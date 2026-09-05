import { useEffect, useRef, useState } from "react";
import { FileUpload } from "@/components/ui/file-upload";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter,
} from "@/components/ui/dialog";
import { Switch } from "@/components/ui/switch";
import { useToast } from "@/hooks/use-toast";
import {
  generateReport,
  Terminal,
} from "@/lib/excel";
import {
  parseExfoXlsx, ExfoXlsxParse, parseAddress, formatDistance, resolveCityFromCLLI,
} from "@/lib/exfo";
import {
  GeocodeHit, loadGeocodeCache, renderEmbeddedMap, autoResolveCity,
  drawConnections, clearConnections, flashPonOnMap, buildOpenInNewTabUrl,
  buildShareData, buildMapViewerUrl,
  startUserLocation, stopUserLocation, UserLocationState,
} from "@/lib/exfo-maps";
import QRCode from "qrcode";

// QR version 40 at error-correction level M holds 2,331 alphanumeric / 1,852 byte chars.
// URLs are encoded in byte mode, so cap below that with a small safety margin.
const QR_URL_MAX_LEN = 1800;

async function tryBuildQrDataUrl(url: string): Promise<string | null> {
  if (!url || url.length > QR_URL_MAX_LEN) return null;
  try {
    return await QRCode.toDataURL(url, { width: 240, margin: 1, errorCorrectionLevel: "M" });
  } catch (err) {
    console.error("[QR] toDataURL failed:", err);
    return null;
  }
}
import { AppToggle } from "@/components/app-toggle";
import { Download, FileJson, Copy, Save, FileSpreadsheet, Settings, MapPin, ExternalLink, Search, Share2, Locate } from "lucide-react";

import { preparePonPower, staggeredFilename, PON_HEADERS, type PonPowerResult } from "@/lib/ponpower";

const API_KEY_STORAGE = "f2job.gmapsApiKey";
const API_KEY_SEEDED = "f2job.gmapsApiKey.seeded";
const DEFAULT_API_KEY = "AIzaSyDWvptOInAO5y7O5pHtVj0GhFQ9aVeYIMc";

interface HomeProps {
  publicMode?: boolean;
}

export default function Home({ publicMode = false }: HomeProps = {}) {
  const { toast } = useToast();
  const [file, setFile] = useState<File | null>(null);
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [orcaText, setOrcaText] = useState("");
  const [conversionError, setConversionError] = useState("");
  const [cableId, setCableId] = useState<string>("");
  const [wireCenterClli, setWireCenterClli] = useState<string>("");
  const [cfas, setCfas] = useState<string>("");
  const [strands, setStrands] = useState<number[]>([]);
  const [averageLoss, setAverageLoss] = useState<string>("-15.5");
  const [jsonOutput, setJsonOutput] = useState<string>("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [parsedWorkbook, setParsedWorkbook] = useState<PonPowerResult | null>(null);
  const [staggeredTerminals, setStaggeredTerminals] = useState<Terminal[]>([]);

  // ----- Map state (mirrors F2 Exfo) -----
  const [parsedExfo, setParsedExfo] = useState<ExfoXlsxParse | null>(null);
  const [apiKey, setApiKey] = useState("");
  const [showSettings, setShowSettings] = useState(false);
  const [city, setCity] = useState("Appleton, WI");
  const [mapStatus, setMapStatus] = useState("");
  const [mapStatusKind, setMapStatusKind] = useState<"err" | "ok" | "">("");
  const [mapLoaded, setMapLoaded] = useState(false);
  const [pfpLocation, setPfpLocation] = useState<GeocodeHit | null>(null);
  const [distances, setDistances] = useState<Map<number, number>>(new Map());
  const [showConnections, setShowConnections] = useState(false);
  const [ponSearch, setPonSearch] = useState("");
  const [shareUrl, setShareUrl] = useState<string | null>(null);
  const [shareQrDataUrl, setShareQrDataUrl] = useState<string | null>(null);
  const [shareDialogOpen, setShareDialogOpen] = useState(false);
  const [sharing, setSharing] = useState(false);
  const [showMyLocation, setShowMyLocation] = useState(true);
  const mapCanvasRef = useRef<HTMLDivElement>(null);
  const mapStateRef = useRef<{
    gmap: any | null; gmarkers: any[]; connections: any[]; showConnections: boolean;
  }>({ gmap: null, gmarkers: [], connections: [], showConnections: false });
  const userLocStateRef = useRef<UserLocationState>({});
  const geocodeCacheRef = useRef<Map<string, GeocodeHit>>(new Map());

  // Load API key + cache once
  useEffect(() => {
    try {
      let v = localStorage.getItem(API_KEY_STORAGE);
      if (v === null && !localStorage.getItem(API_KEY_SEEDED)) {
        v = DEFAULT_API_KEY;
        localStorage.setItem(API_KEY_STORAGE, v);
        localStorage.setItem(API_KEY_SEEDED, "1");
      }
      if (v) setApiKey(v);
    } catch (_) {}
    geocodeCacheRef.current = loadGeocodeCache();
  }, []);

  // Persist API key
  useEffect(() => {
    try { localStorage.setItem(API_KEY_STORAGE, apiKey); } catch (_) {}
  }, [apiKey]);

  // Auto-set city when CLLI changes
  useEffect(() => {
    const hit = resolveCityFromCLLI(wireCenterClli);
    if (hit) setCity(hit.city);
  }, [wireCenterClli]);

  useEffect(() => {
    const handler = (e: any) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };
    window.addEventListener('beforeinstallprompt', handler);
    return () => window.removeEventListener('beforeinstallprompt', handler);
  }, []);

  const handleInstall = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult: any) => {
        if (choiceResult.outcome === 'accepted') {
          setDeferredPrompt(null);
        }
      });
    }
  };

  // Parse file when uploaded
  useEffect(() => {
    if (file) {
      // Reset map state for new file
      setParsedExfo(null);
      setPfpLocation(null);
      setDistances(new Map());
      setMapLoaded(false);
      setMapStatus("");
      setMapStatusKind("");
      if (mapStateRef.current.gmap) {
        for (const m of mapStateRef.current.gmarkers) m.setMap(null);
        mapStateRef.current.gmarkers = [];
        clearConnections(mapStateRef.current);
        mapStateRef.current.gmap = null;
      }
      if (mapCanvasRef.current) mapCanvasRef.current.style.display = "none";

      // Also parse for PFP/terminal map data
      file.arrayBuffer()
        .then(ab => { try { setParsedExfo(parseExfoXlsx(ab)); } catch (_) {} })
        .catch(() => {});
    }
  }, [file]);

  useEffect(() => {
    let cancelled = false;
    setParsedWorkbook(null);
    setStaggeredTerminals([]);
    setConversionError("");
    setStrands([]); setJsonOutput(""); setCableId(""); setCfas("");
    if (!file || !dataFile) { setIsProcessing(false); return; }
    setIsProcessing(true);
    preparePonPower(file, dataFile, orcaText).then(result => {
      if (cancelled) return;
      setParsedWorkbook(result);
      setStaggeredTerminals(result.terminals);
      setCableId(result.cableId);
      setCfas(result.project);
      setStrands(result.terminals.map(t => t.powerTestStrand));
    }).catch(err => {
      if (!cancelled) setConversionError(err instanceof Error ? err.message : "Could not read the uploaded files.");
    }).finally(() => { if (!cancelled) setIsProcessing(false); });
    return () => { cancelled = true; };
  }, [file, dataFile, orcaText]);

  // Generate JSON when dependencies change
  useEffect(() => {
    if (strands.length > 0) {
      const loss = parseFloat(averageLoss);
      if (!isNaN(loss)) {
        const ordered = [...strands].sort((a, b) => a - b);
        const report = generateReport(cableId, ordered, loss, wireCenterClli, cfas);
        setJsonOutput(JSON.stringify(report, null, 2));
      }
    }
  }, [cableId, strands, averageLoss, wireCenterClli, cfas]);

  // Auto-load the map after a fresh upload, mirroring F2 Exfo's behavior.
  // The check on mapLoaded prevents re-entry after the first successful render.
  useEffect(() => {
    if (!parsedExfo || parsedExfo.terminals.length === 0) return;
    if (!apiKey.trim() || mapLoaded) return;
    if (!mapCanvasRef.current) return;
    handleLoadMap();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [parsedExfo, apiKey]);

  const handleSave = () => {
    try {
      const json = JSON.parse(jsonOutput);
      const pretty = JSON.stringify(json, null, 2);
      const base = cfas?.trim() || cableId?.trim() || "report";
      const filename = `${base}.json`;
      const blob = new Blob([pretty], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      toast({
        title: "Saved",
        description: `Saved as ${filename} to your Downloads folder.`,
      });
    } catch (e: any) {
      toast({
        variant: "destructive",
        title: "Could Not Save",
        description: e?.message ?? "The JSON content is invalid and cannot be saved.",
      });
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(jsonOutput);
    toast({
      title: "Copied to Clipboard",
      description: "You can paste it anywhere now.",
    });
  };

  const handleLoadMap = async () => {
    if (!apiKey.trim()) {
      setMapStatus("Add a Google Maps API key in Settings.");
      setMapStatusKind("err");
      return;
    }
    if (!parsedExfo || !mapCanvasRef.current) return;
    if (!parsedExfo.terminals.length) return;
    // The current share link points at a now-stale snapshot; force a rebuild on next export/share.
    setShareUrl(null);
    setShareQrDataUrl(null);
    try {
      // Try auto-resolve if user hasn't manually set city
      const hit = resolveCityFromCLLI(wireCenterClli);
      let useCity = city.trim();
      if (!hit) {
        useCity = await autoResolveCity(
          apiKey.trim(),
          [wireCenterClli, cableId].filter(Boolean),
          parsedExfo.pfpName,
          parsedExfo.terminals,
          geocodeCacheRef.current,
          city.trim(),
          (m, k) => { setMapStatus(m); setMapStatusKind(k || ""); }
        );
        if (useCity && useCity !== city) setCity(useCity);
      }
      const result = await renderEmbeddedMap(
        mapCanvasRef.current,
        mapStateRef.current,
        {
          apiKey: apiKey.trim(),
          city: useCity,
          pfpName: parsedExfo.pfpName,
          terminals: parsedExfo.terminals,
          cache: geocodeCacheRef.current,
          portByRow: portByExfoRow,
          onStatus: (m, k) => { setMapStatus(m); setMapStatusKind(k || ""); },
        }
      );
      setPfpLocation(result.pfpLocation);
      setDistances(result.distances);
      setMapLoaded(true);
      setMapStatus(`Mapped ${result.resolved} terminals in ${result.elapsedSec.toFixed(1)}s${result.failed ? ` — ${result.failed} failed` : ""}.`);
      setMapStatusKind(result.failed ? "err" : "ok");
    } catch (e: any) {
      setMapStatus(e?.message || String(e));
      setMapStatusKind("err");
    }
  };

  const handleShareLink = async () => {
    if (!apiKey.trim() || !parsedExfo) {
      setMapStatus("API key required — set it in Settings.");
      setMapStatusKind("err");
      return;
    }
    if (!mapLoaded) {
      setMapStatus("Load the map before sharing.");
      setMapStatusKind("err");
      return;
    }
    setSharing(true);
    try {
      const data = buildShareData(
        parsedExfo.project || cableId || "Terminal map",
        parsedExfo.pfpName,
        pfpLocation,
        city.trim(),
        parsedExfo.terminals,
        geocodeCacheRef.current,
        showConnections,
        portByExfoRow,
      );
      const url = await buildMapViewerUrl(apiKey.trim(), data);
      setShareUrl(url);
      setShareQrDataUrl(await tryBuildQrDataUrl(url));
      setShareDialogOpen(true);
      try { await navigator.clipboard.writeText(url); } catch (_) {}
      toast({
        title: "Share link copied",
        description: "Direct link to the map — paste anywhere.",
      });
    } catch (e: any) {
      setMapStatus(e?.message || "Could not build share link.");
      setMapStatusKind("err");
    } finally {
      setSharing(false);
    }
  };

  const handleOpenMapInNewTab = () => {
    if (!apiKey.trim() || !parsedExfo) {
      setMapStatus("API key required — set it in Settings.");
      setMapStatusKind("err");
      return;
    }
    const url = buildOpenInNewTabUrl(
      apiKey.trim(),
      parsedExfo.project || cableId || "Terminal map",
      parsedExfo.pfpName,
      pfpLocation,
      city.trim(),
      parsedExfo.terminals,
      geocodeCacheRef.current,
      showConnections,
      portByExfoRow,
    );
    const w = window.open(url, "_blank");
    if (!w) {
      setMapStatus("Popup blocked — allow popups for this page.");
      setMapStatusKind("err");
    }
  };

  // Toggle connections on map
  useEffect(() => {
    mapStateRef.current.showConnections = showConnections;
    if (!mapStateRef.current.gmap || !parsedExfo) return;
    if (showConnections) {
      drawConnections(mapStateRef.current, city, parsedExfo.terminals, geocodeCacheRef.current);
    } else {
      clearConnections(mapStateRef.current);
    }
  }, [showConnections, city, parsedExfo]);

  const handleFlashPon = () => {
    const raw = ponSearch.trim();
    if (!raw) return;
    const n = parseInt(raw, 10);
    if (Number.isNaN(n)) {
      setMapStatus("Enter a numeric Pon count.");
      setMapStatusKind("err");
      return;
    }
    if (!parsedExfo || !mapStateRef.current.gmap || !mapStateRef.current.gmarkers.length) {
      setMapStatus("Load the map first.");
      setMapStatusKind("err");
      return;
    }
    const result = flashPonOnMap(mapStateRef.current, n, parsedExfo.terminals);
    if (!result.matches) {
      setMapStatus(`No terminal contains Pon count ${n}.`);
      setMapStatusKind("err");
    } else {
      setMapStatus(`Flashing ${result.matches} terminal${result.matches > 1 ? "s" : ""} with Pon count ${n}.`);
      setMapStatusKind("ok");
    }
  };

  const handleToggleMyLocation = () => {
    if (showMyLocation) {
      stopUserLocation(userLocStateRef.current);
      setShowMyLocation(false);
      return;
    }
    if (!mapStateRef.current.gmap) {
      setMapStatus("Load the map first.");
      setMapStatusKind("err");
      return;
    }
    startUserLocation(mapStateRef.current.gmap, userLocStateRef.current, (msg) => {
      setMapStatus(msg);
      setMapStatusKind("err");
      setShowMyLocation(false);
    });
    setShowMyLocation(true);
  };

  // Stop GPS watch when the component unmounts or the file changes
  // (the old gmap reference is dropped when a new file is uploaded).
  useEffect(() => {
    return () => { stopUserLocation(userLocStateRef.current); };
  }, []);
  useEffect(() => {
    if (file) {
      // Clear the watch and old marker; the auto-start effect below
      // re-attaches a fresh marker once the new map renders.
      stopUserLocation(userLocStateRef.current);
    }
  }, [file]);

  // Auto-start GPS tracking once the map is up, since "Show my location"
  // defaults to On. Waits for the map's first idle so the marker isn't
  // added mid-render; pans to the user so the blue dot is actually visible.
  useEffect(() => {
    if (!mapLoaded || !showMyLocation) return;
    const gmap = mapStateRef.current.gmap;
    if (!gmap || userLocStateRef.current.watchId != null) return;
    const begin = () => {
      if (!showMyLocation || userLocStateRef.current.watchId != null) return;
      startUserLocation(gmap, userLocStateRef.current, (msg) => {
        setMapStatus(msg);
        setMapStatusKind("err");
        setShowMyLocation(false);
      });
    };
    const w: any = window;
    if (w.google?.maps?.event) {
      w.google.maps.event.addListenerOnce(gmap, "idle", begin);
    } else {
      begin();
    }
  }, [mapLoaded, showMyLocation]);

  const maxDistance = (() => {
    let m: number | null = null;
    Array.from(distances.values()).forEach(v => { if (m == null || v > m) m = v; });
    return m;
  })();

  // Map distance lookup from ExfoTerminal.row (1-based) to Terminal.rowIndex (0-based).
  // Both parsers point at the same physical xlsx rows: ExfoTerminal.row === Terminal.rowIndex + 1.
  const footageByRowIndex = (() => {
    const m = new Map<number, number>();
    for (const t of staggeredTerminals) {
      const original = parsedExfo?.terminals.find(item => item.waldo === t.waldoId && item.terminal === t.terminalName);
      const distance = original ? distances.get(original.row) : undefined;
      if (distance != null) m.set(t.rowIndex, distance);
    }
    return m;
  })();

  // Test-port lookup keyed by ExfoTerminal.row (1-based) so map markers and
  // share data can show the staggered Test Port number alongside the PON count.
  const portByExfoRow = (() => {
    const m = new Map<number, number>();
    for (const t of staggeredTerminals) {
      const original = parsedExfo?.terminals.find(item => item.waldo === t.waldoId && item.terminal === t.terminalName);
      if (original && t.staggeredPort != null) m.set(original.row, t.staggeredPort);
    }
    return m;
  })();

  const ensureShareUrl = async (): Promise<string | null> => {
    if (shareUrl) return shareUrl;
    if (!apiKey.trim() || !parsedExfo || !mapLoaded) return null;
    try {
      const data = buildShareData(
        parsedExfo.project || cableId || "Terminal map",
        parsedExfo.pfpName,
        pfpLocation,
        city.trim(),
        parsedExfo.terminals,
        geocodeCacheRef.current,
        showConnections,
        portByExfoRow,
      );
      const url = await buildMapViewerUrl(apiKey.trim(), data);
      setShareUrl(url);
      const qr = await tryBuildQrDataUrl(url);
      if (qr) setShareQrDataUrl(qr);
      return url;
    } catch (err) {
      console.error("ensureShareUrl failed:", err);
      return null;
    }
  };

  const handleDownloadConverted = async () => {
    if (!parsedWorkbook || staggeredTerminals.length === 0) return;
    try {
      const sheet = parsedWorkbook.workbook.getWorksheet(parsedWorkbook.sheetName)!;
      for (const t of staggeredTerminals) {
        const footage = footageByRowIndex.get(t.rowIndex);
        if (footage != null) sheet.getCell(t.rowIndex + 1, 8).value = Math.round(footage);
      }
      const qrShareUrl = await ensureShareUrl();
      if (qrShareUrl && sheet.getImages().length === 0) {
        const qr = await tryBuildQrDataUrl(qrShareUrl);
        if (qr) {
          const imageId = parsedWorkbook.workbook.addImage({ base64: qr, extension: 'png' });
          if (!sheet.getCell('L1').isMerged) sheet.mergeCells('L1:M1');
          sheet.getCell('L1').value = 'Scan QR for Map';
          sheet.getRow(2).height = 64;
          sheet.addImage(imageId, { tl: { col: 11.63, row: 1.05 }, ext: { width: 84, height: 84 } });
        }
      }
      const bytes = await parsedWorkbook.workbook.xlsx.writeBuffer();
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");

      a.href = url;
      a.download = staggeredFilename(file?.name ?? "PONSHEET.xlsx");
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      toast({
        title: "Spreadsheet Downloaded",
        description: "Staggered PON sheet with Task and Orca tabs downloaded.",
      });
    } catch (e: any) {
      console.error(e);
      toast({
        variant: "destructive",
        title: "Error Generating Spreadsheet",
        description: e?.message ? String(e.message).slice(0, 220) : "Could not build the converted file.",
      });
    }
  };

  return (
    <div className="min-h-screen bg-background p-6 md:p-12 font-sans text-foreground">
      <div className="max-w-6xl mx-auto space-y-8">

        {/* Top center: app toggle */}
        <div className="flex justify-center">
          <AppToggle
            active="ponpower"
            publicMode={publicMode}
            keys={publicMode ? ["ponpower", "exfo"] : undefined}
          />
        </div>

        {/* Header */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-border pb-6">
          <div>
            <h1 className="text-3xl md:text-4xl font-bold tracking-tight bg-gradient-to-r from-primary to-blue-400 bg-clip-text text-transparent">
              FiberOptic Report Gen
            </h1>
            <p className="text-muted-foreground mt-2">
              Offline tool for generating JSON test reports from Excel field data.
            </p>
          </div>
          <div className="flex items-center gap-2 px-3 py-1 rounded-full bg-secondary/50 text-xs font-mono text-muted-foreground border border-white/5">
            <div className="w-2 h-2 rounded-full bg-green-500 animate-pulse" />
            SYSTEM ONLINE
          </div>
          <Button variant="outline" size="sm" onClick={() => setShowSettings(true)} className="gap-2">
            <Settings className="w-4 h-4" /> Settings
          </Button>
          {deferredPrompt && (
            <Button onClick={handleInstall} size="sm" className="gap-2 bg-primary text-primary-foreground hover:bg-primary/90">
              <Download className="w-4 h-4" />
              Install App
            </Button>
          )}
        </div>

        <div className={publicMode ? "max-w-2xl mx-auto" : "grid grid-cols-1 lg:grid-cols-12 gap-8"}>

          {/* Left Column: Inputs */}
          <div className={publicMode ? "" : "lg:col-span-5 space-y-6"}>

            {/* Step 1: Upload */}
            <Card className="border-border/50 bg-card/50 backdrop-blur-sm shadow-lg">
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">1</div>
                  Upload Data
                </CardTitle>
                <CardDescription>
                  Upload a Ponsheet and its task data export (.xlsx).
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <FileUpload
                  label="Upload Ponsheet"
                  helper="Choose your Ponsheet .xlsx workbook"
                  accept=".xlsx"
                  onFileSelect={setFile}
                  onClear={() => setFile(null)}
                />
                <FileUpload
                  label="Upload data file"
                  helper="Choose the task export, like data.xlsx"
                  accept=".xlsx"
                  onFileSelect={setDataFile}
                  onClear={() => setDataFile(null)}
                />
                <div className="space-y-2">
                  <Label htmlFor="orca-paste">Paste Orca status (optional)</Label>
                  <Textarea id="orca-paste" value={orcaText} onChange={e => setOrcaText(e.target.value)}
                    placeholder={"Task\tStatus\n1.478\tC\n1.479\tO"}
                    className="font-mono text-xs min-h-28" />
                  <p className="text-xs text-muted-foreground">Copy and paste Orca rows with Task first and Status second. Headers are optional.</p>
                </div>
                {conversionError && <p role="alert" className="text-sm text-destructive">{conversionError}</p>}
                {parsedWorkbook && <p className="text-sm text-muted-foreground">
                  Matched {parsedWorkbook.matchedTasks} of {parsedWorkbook.terminals.length} terminals to tasks.
                  {parsedWorkbook.matchedTasks < parsedWorkbook.terminals.length && " Unmatched tasks remain blank. Check that both files are for the same project."}
                </p>}
              </CardContent>
            </Card>

            {!publicMode && (
            <>
            {/* Step 2: Configure */}
            <Card className="border-border/50 bg-card/50 backdrop-blur-sm shadow-lg">
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">2</div>
                  Configuration
                </CardTitle>
                <CardDescription>
                  Set parameters for the report generation
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2 col-span-2">
                    <Label htmlFor="cableId">Cable ID</Label>
                    <Input
                      id="cableId"
                      value={cableId}
                      onChange={(e) => setCableId(e.target.value)}
                      placeholder="Extracted from Excel..."
                      className="font-mono bg-secondary/20 border-border focus:border-primary"
                    />
                  </div>

                  <div className="space-y-2 col-span-2">
                    <Label htmlFor="wireCenterClli">Wire Center CLLI</Label>
                    <Input
                      id="wireCenterClli"
                      value={wireCenterClli}
                      onChange={(e) => setWireCenterClli(e.target.value)}
                      placeholder="e.g. LKGNWI01"
                      className="font-mono bg-secondary/20 border-border focus:border-primary"
                    />
                    <p className="text-xs text-muted-foreground">
                      Automatically updates aLoc/zLoc with PFP suffix.
                    </p>
                  </div>

                  <div className="space-y-2 col-span-2">
                    <Label htmlFor="cfas">Project #</Label>
                    <Input
                      id="cfas"
                      value={cfas}
                      onChange={(e) => setCfas(e.target.value)}
                      placeholder="Extracted from Excel..."
                      className="font-mono bg-secondary/20 border-border focus:border-primary"
                    />
                  </div>

                  <div className="space-y-2 col-span-2">
                    <Label htmlFor="avgLoss">Target Average Loss (dB)</Label>
                    <div className="relative">
                      <Input
                        id="avgLoss"
                        type="number"
                        step="0.1"
                        value={averageLoss}
                        onChange={(e) => setAverageLoss(e.target.value)}
                        className="font-mono bg-secondary/20 border-border focus:border-primary pr-12"
                      />
                      <div className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground text-sm font-mono">
                        dB
                      </div>
                    </div>
                    <p className="text-xs text-muted-foreground">
                      Values will be randomized ±0.5 dB from this average.
                    </p>
                  </div>
                </div>

                <div className="pt-4 border-t border-border">
                  <div className="flex justify-between items-center text-sm">
                    <span className="text-muted-foreground">Detected Strands:</span>
                    <span className="font-mono font-bold text-primary">{strands.length}</span>
                  </div>
                </div>
              </CardContent>
            </Card>
            </>
            )}
          </div>

          {!publicMode && (
          <div className="lg:col-span-7 h-full">
            <Card className="h-full flex flex-col border-border/50 bg-card/50 backdrop-blur-sm shadow-lg overflow-hidden">
              <CardHeader className="pb-3 border-b border-border/50 bg-secondary/10">
                <div className="flex justify-between items-center">
                  <CardTitle className="flex items-center gap-2">
                    <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">3</div>
                    JSON Preview
                  </CardTitle>
                  <div className="flex gap-2">
                    <Button variant="outline" size="sm" onClick={handleCopy} className="h-8 text-xs gap-1.5">
                      <Copy className="w-3.5 h-3.5" /> Copy
                    </Button>
                    <Button onClick={handleSave} size="sm" className="h-8 text-xs gap-1.5" disabled={!jsonOutput}>
                      <Save className="w-3.5 h-3.5" /> Save
                    </Button>
                  </div>
                </div>
              </CardHeader>

              <div className="flex-1 relative group">
                <Textarea
                  value={jsonOutput}
                  onChange={(e) => setJsonOutput(e.target.value)}
                  className="w-full h-[600px] lg:h-full resize-none rounded-none border-0 p-4 font-mono text-sm bg-background/50 focus-visible:ring-0 text-blue-300 leading-relaxed"
                  spellCheck={false}
                />
                {!jsonOutput && (
                  <div className="absolute inset-0 flex flex-col items-center justify-center text-muted-foreground pointer-events-none opacity-50">
                    <FileJson className="w-12 h-12 mb-2" />
                    <p>JSON output will appear here</p>
                  </div>
                )}
              </div>
            </Card>
          </div>
          )}

        </div>

        {/* Step 4: Spreadsheet Preview */}
        <Card className="border-border/50 bg-card/50 backdrop-blur-sm shadow-lg overflow-hidden">
          <CardHeader className="pb-3 border-b border-border/50 bg-secondary/10">
            <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-3">
              <div>
                <CardTitle className="flex items-center gap-2">
                  <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">4</div>
                  Spreadsheet Preview
                </CardTitle>
                <CardDescription className="mt-1">
                  Compact PON test sheet with staggered ports and strands, task lookups, and status lookups.
                </CardDescription>
              </div>
              <Button
                onClick={handleDownloadConverted}
                size="sm"
                className="h-8 text-xs gap-1.5"
                disabled={staggeredTerminals.length === 0}
              >
                <FileSpreadsheet className="w-3.5 h-3.5" /> Download XLSX
              </Button>
            </div>
          </CardHeader>
          <CardContent className="p-0">
            {staggeredTerminals.length > 0 ? (
              <div className="max-h-[500px] overflow-auto">
                <Table className="font-mono text-xs">
                  <TableHeader className="sticky top-0 bg-secondary/40 backdrop-blur-sm z-10">
                    <TableRow>
                      {PON_HEADERS.map(header => <TableHead key={header} className="whitespace-nowrap">{header}</TableHead>)}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {staggeredTerminals.map((t) => (
                      <TableRow key={t.rowIndex}>
                        {PON_HEADERS.map((header, i) => (
                          <TableCell key={header} className="whitespace-nowrap">
                            {i === 7 && footageByRowIndex.has(t.rowIndex) ? Math.round(footageByRowIndex.get(t.rowIndex)!).toLocaleString() : parsedWorkbook?.workbook.getWorksheet(parsedWorkbook.sheetName)?.getCell(t.rowIndex + 1, i + 1).text}
                          </TableCell>
                        ))}
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center text-muted-foreground py-16 opacity-50">
                <FileSpreadsheet className="w-12 h-12 mb-2" />
                <p>Upload both files to preview the staggered sheet</p>
              </div>
            )}
          </CardContent>
        </Card>

        {/* Step 5: Terminal map */}
        {parsedExfo && parsedExfo.terminals.length > 0 && (
          <Card className="border-border/50 bg-card/50 backdrop-blur-sm shadow-lg">
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">5</div>
                Terminal map
              </CardTitle>
              <CardDescription>
                {parsedExfo.pfpName ? <>PFP: <code>{parsedExfo.pfpName}</code> — {parseAddress(parsedExfo.pfpName)}</> : "PFP not detected in sheet header."}
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="grid grid-cols-1 md:grid-cols-[1fr_auto] gap-3 items-end">
                <div className="space-y-1.5">
                  <Label>City / State (added to each address for geocoding)</Label>
                  <Input value={city} onChange={e => setCity(e.target.value)} />
                </div>
                <Button onClick={handleLoadMap} size="sm" className="gap-1.5">
                  <MapPin className="w-3.5 h-3.5" /> {mapLoaded ? "Refresh map" : "Load map"}
                </Button>
              </div>
              <div className="flex items-center gap-3 flex-wrap">
                <Button onClick={handleOpenMapInNewTab} variant="outline" size="sm" className="gap-1.5" disabled={!mapLoaded}>
                  <ExternalLink className="w-3.5 h-3.5" /> Open map in new tab
                </Button>
                <Button onClick={handleShareLink} variant="outline" size="sm" className="gap-1.5" disabled={!mapLoaded || sharing}>
                  <Share2 className="w-3.5 h-3.5" /> {sharing ? "Building link…" : "Share link"}
                </Button>
                <Button onClick={handleToggleMyLocation} variant="outline" size="sm" className="gap-1.5" disabled={!mapLoaded}>
                  <Locate className="w-3.5 h-3.5" /> {showMyLocation ? "Hide my location" : "Show my location"}
                </Button>
                <label className="flex items-center gap-2 text-xs cursor-pointer">
                  <Switch checked={showConnections} onCheckedChange={setShowConnections} />
                  Connect pins by strand order
                </label>
                {maxDistance != null && (
                  <span className="text-xs text-muted-foreground">
                    Furthest from PFP: <b>{formatDistance(maxDistance)}</b>
                  </span>
                )}
              </div>
              <div className="flex items-end gap-2">
                <div className="space-y-1.5 flex-1">
                  <Label>Find Pon count</Label>
                  <Input
                    type="number"
                    placeholder="e.g. 125"
                    min="0"
                    value={ponSearch}
                    onChange={e => setPonSearch(e.target.value)}
                    onKeyDown={e => { if (e.key === "Enter") handleFlashPon(); }}
                  />
                </div>
                <Button onClick={handleFlashPon} variant="outline" size="sm" className="gap-1.5">
                  <Search className="w-3.5 h-3.5" /> Flash
                </Button>
              </div>
              {mapStatus && (
                <p className={`text-xs ${mapStatusKind === "err" ? "text-red-400" : mapStatusKind === "ok" ? "text-green-400" : "text-muted-foreground"}`}>
                  {mapStatus}
                </p>
              )}
              <div ref={mapCanvasRef} style={{ width: "100%", height: 460, borderRadius: 8, background: "hsl(var(--secondary) / 0.4)", display: "none" }} />
              <details>
                <summary className="text-sm cursor-pointer text-muted-foreground hover:text-foreground">
                  Addresses ({parsedExfo.terminals.length})
                </summary>
                <div className="max-h-[360px] overflow-auto border border-border rounded mt-2">
                  <Table className="text-xs">
                    <TableHeader>
                      <TableRow>
                        <TableHead>#</TableHead>
                        <TableHead>Terminal</TableHead>
                        <TableHead>Address</TableHead>
                        <TableHead>Waldo</TableHead>
                        <TableHead>Distance from PFP</TableHead>
                        <TableHead>Map</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {parsedExfo.terminals.map((t, i) => {
                        const addr = parseAddress(t.terminal);
                        const q = encodeURIComponent(city ? `${addr}, ${city}` : addr);
                        const d = distances.get(t.row);
                        return (
                          <TableRow key={t.row}>
                            <TableCell>{i + 1}</TableCell>
                            <TableCell className="font-mono whitespace-nowrap">{t.terminal}</TableCell>
                            <TableCell>{addr}</TableCell>
                            <TableCell>{t.waldo}</TableCell>
                            <TableCell>{d != null ? formatDistance(d) : <span className="text-muted-foreground">—</span>}</TableCell>
                            <TableCell><a href={`https://www.google.com/maps/search/?api=1&query=${q}`} target="_blank" rel="noopener noreferrer" className="text-primary hover:underline">Open</a></TableCell>
                          </TableRow>
                        );
                      })}
                    </TableBody>
                  </Table>
                </div>
              </details>
            </CardContent>
          </Card>
        )}
      </div>

      <Dialog open={showSettings} onOpenChange={setShowSettings}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Settings</DialogTitle>
          </DialogHeader>
          <div className="space-y-3">
            <div className="space-y-1.5">
              <Label>Google Maps API key</Label>
              <Input
                type="password"
                value={apiKey}
                onChange={e => setApiKey(e.target.value)}
                placeholder="AIza..."
                autoComplete="off"
                spellCheck={false}
              />
              <p className="text-xs text-muted-foreground">
                Stored locally in your browser. Restrict the key by HTTP referrer in Google Cloud Console.{" "}
                <a className="text-primary hover:underline" target="_blank" rel="noopener noreferrer" href="https://console.cloud.google.com/google/maps-apis/credentials">
                  Create a key →
                </a>
              </p>
            </div>
          </div>
          <DialogFooter>
            <Button onClick={() => setShowSettings(false)}>Close</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={shareDialogOpen} onOpenChange={setShareDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Share Terminal map</DialogTitle>
          </DialogHeader>
          <div className="space-y-3">
            <p className="text-xs text-muted-foreground">
              Anyone with this link can open the map. The URL embeds your Google Maps API key —
              restrict it by HTTP referrer in Google Cloud Console for safety.
            </p>
            <div className="flex gap-2">
              <Input value={shareUrl ?? ""} readOnly onFocus={e => e.currentTarget.select()} className="font-mono text-xs" />
              <Button
                size="sm"
                variant="outline"
                onClick={() => {
                  if (shareUrl) {
                    navigator.clipboard.writeText(shareUrl);
                    toast({ title: "Copied", description: "Link copied to clipboard." });
                  }
                }}
              >
                <Copy className="w-3.5 h-3.5" />
              </Button>
            </div>
            {shareQrDataUrl && (
              <div className="flex flex-col items-center gap-2 pt-2">
                <img
                  src={shareQrDataUrl}
                  alt="QR code linking to the shared map"
                  className="w-44 h-44 rounded-md bg-white p-2"
                />
                <p className="text-xs text-muted-foreground text-center">
                  Scan to open the map. The same QR is embedded in the exported XLSX.
                </p>
                <div className="flex gap-2">
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => {
                      if (!shareQrDataUrl) return;
                      const a = document.createElement("a");
                      a.href = shareQrDataUrl;
                      const base = (parsedExfo?.project || cableId || "terminal-map").replace(/[^a-z0-9_-]+/gi, "_");
                      a.download = `${base}_qr.png`;
                      document.body.appendChild(a);
                      a.click();
                      document.body.removeChild(a);
                    }}
                  >
                    <Download className="w-3.5 h-3.5 mr-1.5" /> Download QR
                  </Button>
                </div>
              </div>
            )}
          </div>
          <DialogFooter>
            <Button onClick={() => setShareDialogOpen(false)}>Close</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
