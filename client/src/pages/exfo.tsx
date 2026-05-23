import { useEffect, useMemo, useRef, useState } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Button } from '@/components/ui/button';
import { Label } from '@/components/ui/label';
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from '@/components/ui/select';
import {
  Table, TableBody, TableCell, TableHead, TableHeader, TableRow,
} from '@/components/ui/table';
import {
  Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter,
} from '@/components/ui/dialog';
import { Switch } from '@/components/ui/switch';
import { useToast } from '@/hooks/use-toast';
import { AppToggle } from '@/components/app-toggle';
import { Settings, Save, AlertTriangle, MapPin, ExternalLink, Search } from 'lucide-react';
import {
  parseExfoXlsx, ExfoTerminal, ExfoXlsxParse, ExfoJobConfig, resolveJobConfig,
  composeTestConfigurations, buildCandidates, findWaldoDuplicateGroups,
  findFiberDuplicateGroups, candidatesToRows, rowsToCsv, formatDistance,
  parseAddress, IOLM_SHORT_LINK_FT, resolveCityFromCLLI,
  clusterFiberGroupsByParticipants, pickCurrentClusterWinner,
} from '@/lib/exfo';
import {
  GeocodeHit, loadGeocodeCache, renderEmbeddedMap, autoResolveCity,
  drawConnections, clearConnections, flashPonOnMap, buildOpenInNewTabHtml,
} from '@/lib/exfo-maps';

const DEFAULT_CFG: ExfoJobConfig = {
  name: '', techID: '', domain: '@att.com',
  aloc: '', zloc: '', wcc: '',
  testType: 'iOLM',
  iolmPreset: 'over', iolmCustom: '', iolmCustomExt: '.iolmcfg',
  opmPreset: 't24', opmCustom: '', opmCustomExt: '.opmcfg',
};

const REQUIRED: Array<[keyof ExfoJobConfig, string]> = [
  ['name', 'CFAS (Job name)'],
  ['techID', 'Tech UID'],
  ['domain', 'Domain'],
  ['aloc', 'CLLI (ALoc)'],
  ['zloc', 'CO CLLI'],
  ['wcc', 'PFP CLLI'],
];

const API_KEY_STORAGE = 'f2job.gmapsApiKey';
const API_KEY_SEEDED = 'f2job.gmapsApiKey.seeded';
const TECH_UID_STORAGE = 'f2job.techUID';
const DEFAULT_API_KEY = 'AIzaSyDWvptOInAO5y7O5pHtVj0GhFQ9aVeYIMc';

type StateBadge = { text: string; tone: 'idle' | 'ready' | 'fail' };

interface ExfoProps {
  publicMode?: boolean;
}

export default function Exfo({ publicMode = false }: ExfoProps = {}) {
  const { toast } = useToast();
  const [parsed, setParsed] = useState<ExfoXlsxParse | null>(null);
  const [fileName, setFileName] = useState('No file chosen.');
  const [cfg, setCfg] = useState<ExfoJobConfig>(DEFAULT_CFG);
  const [excludedTerminals, setExcludedTerminals] = useState<Set<number>>(new Set());
  const [excludedCands, setExcludedCands] = useState<Set<string>>(new Set());

  const [apiKey, setApiKey] = useState('');
  const [showSettings, setShowSettings] = useState(false);
  const [city, setCity] = useState('Appleton, WI');
  const [mapStatus, setMapStatus] = useState('');
  const [mapStatusKind, setMapStatusKind] = useState<'err' | 'ok' | ''>('');
  const [mapLoaded, setMapLoaded] = useState(false);
  const [pfpLocation, setPfpLocation] = useState<GeocodeHit | null>(null);
  const [distances, setDistances] = useState<Map<number, number>>(new Map());
  const [showConnections, setShowConnections] = useState(false);
  const [ponSearch, setPonSearch] = useState('');
  const [iolmAutoApplied, setIolmAutoApplied] = useState(false);
  const [stateBadge, setStateBadge] = useState<StateBadge>({ text: 'Idle', tone: 'idle' });
  const [alertOpen, setAlertOpen] = useState(false);
  const [alertContent, setAlertContent] = useState<{ title: string; body: React.ReactNode }>({ title: '', body: null });

  const mapCanvasRef = useRef<HTMLDivElement>(null);
  const mapStateRef = useRef<{
    gmap: any | null; gmarkers: any[]; connections: any[]; showConnections: boolean;
  }>({ gmap: null, gmarkers: [], connections: [], showConnections: false });
  const geocodeCacheRef = useRef<Map<string, GeocodeHit>>(new Map());

  // Load persisted state on mount
  useEffect(() => {
    try {
      let v: string | null = null;
      v = localStorage.getItem(API_KEY_STORAGE);
      if (v === null && !localStorage.getItem(API_KEY_SEEDED)) {
        v = DEFAULT_API_KEY;
        localStorage.setItem(API_KEY_STORAGE, v);
        localStorage.setItem(API_KEY_SEEDED, '1');
      }
      if (v) setApiKey(v);
    } catch (_) {}
    try {
      const t = localStorage.getItem(TECH_UID_STORAGE);
      if (t) setCfg(c => ({ ...c, techID: t }));
    } catch (_) {}
    geocodeCacheRef.current = loadGeocodeCache();
  }, []);

  useEffect(() => {
    try { localStorage.setItem(API_KEY_STORAGE, apiKey); } catch (_) {}
  }, [apiKey]);

  useEffect(() => {
    try { localStorage.setItem(TECH_UID_STORAGE, cfg.techID.trim()); } catch (_) {}
  }, [cfg.techID]);

  // Auto-fill CFAS from sheet
  useEffect(() => {
    if (parsed?.project && !cfg.name) {
      setCfg(c => ({ ...c, name: parsed.project! }));
    }
  }, [parsed]);

  // Auto-set wcc from PFP name + ALoc when blank
  useEffect(() => {
    if (parsed?.pfpName && !cfg.wcc && cfg.aloc) {
      setCfg(c => ({ ...c, wcc: cfg.aloc + 'PFP' }));
    }
  }, [parsed?.pfpName, cfg.aloc]);

  // Auto-resolve city from CLLI
  useEffect(() => {
    const hit = resolveCityFromCLLI(cfg.aloc) || resolveCityFromCLLI(cfg.zloc) || resolveCityFromCLLI(cfg.wcc);
    if (hit) setCity(hit.city);
  }, [cfg.aloc, cfg.zloc, cfg.wcc]);

  const handleFile = async (file: File) => {
    setFileName(file.name);
    try {
      const ab = await file.arrayBuffer();
      const p = parseExfoXlsx(ab);
      setParsed(p);
      setExcludedTerminals(new Set());
      setExcludedCands(new Set());
      setPfpLocation(null);
      setDistances(new Map());
      setMapLoaded(false);
      setIolmAutoApplied(false);
      toast({ title: 'Loaded', description: `${p.terminals.length} terminals from ${p.sheetName}.` });

      // Auto-resolve city + load map if API key present
      if (apiKey.trim()) {
        const cllis = [cfg.aloc, cfg.zloc, cfg.wcc].filter(Boolean);
        const newCity = await autoResolveCity(
          apiKey.trim(), cllis, p.pfpName, p.terminals,
          geocodeCacheRef.current, city.trim(),
          (m, k) => { setMapStatus(m); setMapStatusKind(k || ''); }
        );
        if (newCity && newCity !== city) setCity(newCity);
        // Trigger map render after city resolution
        if (mapCanvasRef.current) {
          await loadMapWith(p, newCity || city);
        }
      }
    } catch (e: any) {
      toast({ variant: 'destructive', title: 'Parse error', description: e?.message || String(e) });
      setStateBadge({ text: 'Error', tone: 'fail' });
    }
  };

  const missing = useMemo(
    () => REQUIRED.filter(([k]) => !String(cfg[k] || '').trim()).map(([, label]) => label),
    [cfg]
  );

  const waldoGroups = useMemo(
    () => parsed ? findWaldoDuplicateGroups(parsed.terminals) : [],
    [parsed]
  );
  const effectiveExcludedTerminals = useMemo(() => {
    const next = new Set(excludedTerminals);
    if (parsed) {
      const validRows = new Set(parsed.terminals.map(t => t.row));
      Array.from(next).forEach(r => { if (!validRows.has(r)) next.delete(r); });
      for (const g of waldoGroups) {
        const rows = g.items.map(t => t.row);
        const keptCount = rows.filter(r => !next.has(r)).length;
        if (keptCount !== 1) {
          next.delete(rows[0]);
          for (const r of rows.slice(1)) next.add(r);
        }
      }
    }
    return next;
  }, [parsed, waldoGroups, excludedTerminals]);

  const candidates = useMemo(
    () => parsed ? buildCandidates(parsed.terminals, cfg.testType, effectiveExcludedTerminals) : [],
    [parsed, cfg.testType, effectiveExcludedTerminals]
  );

  const fiberGroups = useMemo(() => findFiberDuplicateGroups(candidates), [candidates]);
  const fiberClusters = useMemo(() => clusterFiberGroupsByParticipants(fiberGroups), [fiberGroups]);

  const effectiveExcludedCands = useMemo(() => {
    const next = new Set(excludedCands);
    const validKeys = new Set(candidates.map(c => c.key));
    Array.from(next).forEach(k => { if (!validKeys.has(k)) next.delete(k); });
    for (const g of fiberGroups) {
      const keys = g.items.map(i => i.key);
      const keptCount = keys.filter(k => !next.has(k)).length;
      if (keptCount !== 1) {
        next.delete(keys[0]);
        for (const k of keys.slice(1)) next.add(k);
      }
    }
    return next;
  }, [candidates, fiberGroups, excludedCands]);

  const keptCandidates = useMemo(
    () => candidates.filter(c => !effectiveExcludedCands.has(c.key)),
    [candidates, effectiveExcludedCands]
  );

  const resolvedCfg = useMemo(() => resolveJobConfig(cfg), [cfg]);

  const outputRows = useMemo(
    () => keptCandidates.length ? candidatesToRows(resolvedCfg, keptCandidates) : [],
    [resolvedCfg, keptCandidates]
  );

  const keptTerminals = useMemo(
    () => parsed ? parsed.terminals.filter(t => !effectiveExcludedTerminals.has(t.row)) : [],
    [parsed, effectiveExcludedTerminals]
  );

  const maxDistance = useMemo(() => {
    let m: number | null = null;
    Array.from(distances.values()).forEach(v => { if (m == null || v > m) m = v; });
    return m;
  }, [distances]);

  // Update state badge
  useEffect(() => {
    if (!parsed) { setStateBadge({ text: 'Idle', tone: 'idle' }); return; }
    const conflicts = waldoGroups.length + fiberGroups.length;
    if (conflicts > 0) {
      setStateBadge({ text: 'Conflicts', tone: 'fail' });
    } else if (missing.length === 0 && outputRows.length > 0) {
      setStateBadge({ text: 'Ready to Save', tone: 'ready' });
    } else {
      setStateBadge({ text: 'Idle', tone: 'idle' });
    }
  }, [parsed, waldoGroups.length, fiberGroups.length, missing.length, outputRows.length]);

  // Auto-recommend iOLM preset based on max distance — only auto-apply ONCE per xlsx load
  useEffect(() => {
    if (maxDistance == null || iolmAutoApplied) return;
    const shortLink = maxDistance <= IOLM_SHORT_LINK_FT;
    const target = shortLink ? 'under' : 'over';
    if (cfg.iolmPreset !== target) {
      setCfg(c => ({ ...c, iolmPreset: target }));
    }
    setIolmAutoApplied(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [maxDistance]);

  const showAlert = (title: string, body: React.ReactNode) => {
    setAlertContent({ title, body });
    setAlertOpen(true);
  };

  const handleSaveCsv = () => {
    if (missing.length) {
      showAlert(
        'Fill in Job Settings first',
        <>
          <p>Before saving the CSV, please fill in:</p>
          <ul className="list-disc pl-5 mt-2">
            {missing.map(m => <li key={m}>{m}</li>)}
          </ul>
        </>
      );
      return;
    }
    if (!outputRows.length) {
      showAlert('No data to save', <p>Load a PON test sheet first, then try Save CSV again.</p>);
      return;
    }
    const csv = rowsToCsv(outputRows);
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${cfg.name.trim() || 'exportsheet'}_job.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    toast({ title: 'CSV saved', description: a.download });
  };

  const handleClear = () => {
    setParsed(null);
    setFileName('No file chosen.');
    setExcludedTerminals(new Set());
    setExcludedCands(new Set());
    setPfpLocation(null);
    setDistances(new Map());
    setMapLoaded(false);
    setIolmAutoApplied(false);
    setMapStatus('');
    setMapStatusKind('');
    setShowConnections(false);
    setPonSearch('');
    // Reset job settings (keep techID; that's session-persisted)
    setCfg(c => ({
      ...DEFAULT_CFG,
      techID: c.techID,
    }));
    setCity('Appleton, WI');
    if (mapStateRef.current.gmap) {
      for (const m of mapStateRef.current.gmarkers) m.setMap(null);
      mapStateRef.current.gmarkers = [];
      clearConnections(mapStateRef.current);
      mapStateRef.current.gmap = null;
      mapStateRef.current.showConnections = false;
    }
    if (mapCanvasRef.current) mapCanvasRef.current.style.display = 'none';
  };

  const loadMapWith = async (p: ExfoXlsxParse, useCity: string) => {
    if (!apiKey.trim() || !mapCanvasRef.current) return;
    const terms = p.terminals.filter(t => !effectiveExcludedTerminals.has(t.row));
    if (!terms.length) return;
    try {
      const result = await renderEmbeddedMap(
        mapCanvasRef.current,
        mapStateRef.current,
        {
          apiKey: apiKey.trim(),
          city: useCity.trim(),
          pfpName: p.pfpName,
          terminals: terms,
          cache: geocodeCacheRef.current,
          onStatus: (m, k) => { setMapStatus(m); setMapStatusKind(k || ''); },
        }
      );
      setPfpLocation(result.pfpLocation);
      setDistances(result.distances);
      setMapLoaded(true);
      setMapStatus(`Mapped ${result.resolved} terminals in ${result.elapsedSec.toFixed(1)}s${result.failed ? ` — ${result.failed} failed` : ''}.`);
      setMapStatusKind(result.failed ? 'err' : 'ok');
    } catch (e: any) {
      setMapStatus(e?.message || String(e));
      setMapStatusKind('err');
    }
  };

  const handleLoadMap = async () => {
    if (!apiKey.trim()) {
      setMapStatus('Add a Google Maps API key in Settings.');
      setMapStatusKind('err');
      return;
    }
    if (!parsed) return;
    await loadMapWith(parsed, city);
  };

  const handleOpenInNewTab = () => {
    if (!apiKey.trim()) {
      setMapStatus('API key required — set it in Settings.');
      setMapStatusKind('err');
      return;
    }
    if (!parsed) return;
    const html = buildOpenInNewTabHtml(
      apiKey.trim(),
      parsed.project || cfg.name || 'Terminal map',
      parsed.pfpName,
      pfpLocation,
      city.trim(),
      keptTerminals,
      geocodeCacheRef.current,
      distances,
      showConnections
    );
    const w = window.open('', '_blank');
    if (!w) {
      setMapStatus('Popup blocked — allow popups for this page.');
      setMapStatusKind('err');
      return;
    }
    w.document.open();
    w.document.write(html);
    w.document.close();
  };

  // Toggle connections on the embedded map
  useEffect(() => {
    mapStateRef.current.showConnections = showConnections;
    if (!mapStateRef.current.gmap) return;
    if (showConnections) {
      drawConnections(mapStateRef.current, city, keptTerminals, geocodeCacheRef.current);
    } else {
      clearConnections(mapStateRef.current);
    }
  }, [showConnections, city, keptTerminals]);

  const handleFlashPon = () => {
    const raw = ponSearch.trim();
    if (!raw) return;
    const n = parseInt(raw, 10);
    if (Number.isNaN(n)) {
      setMapStatus('Enter a numeric Pon count.');
      setMapStatusKind('err');
      return;
    }
    if (!mapStateRef.current.gmap || !mapStateRef.current.gmarkers.length) {
      setMapStatus('Load the map first.');
      setMapStatusKind('err');
      return;
    }
    const result = flashPonOnMap(mapStateRef.current, n, keptTerminals);
    if (!result.matches) {
      setMapStatus(`No terminal contains Pon count ${n}.`);
      setMapStatusKind('err');
    } else {
      setMapStatus(`Flashing ${result.matches} terminal${result.matches > 1 ? 's' : ''} with Pon count ${n}.`);
      setMapStatusKind('ok');
    }
  };

  const setOnly = (
    items: Array<{ row?: number; key?: string }>,
    idx: number,
    type: 'waldo' | 'fiber'
  ) => {
    if (type === 'waldo') {
      const next = new Set(excludedTerminals);
      items.forEach((it: any, i) => {
        if (i === idx) next.delete(it.row);
        else next.add(it.row);
      });
      setExcludedTerminals(next);
    } else {
      const next = new Set(excludedCands);
      items.forEach((it: any, i) => {
        if (i === idx) next.delete(it.key);
        else next.add(it.key);
      });
      setExcludedCands(next);
    }
  };

  // For cluster radio: pick a participant terminal -> for each group in cluster, exclude all rows of other terminals
  const handleClusterPick = (cluster: { participantRows: number[]; groups: any[] }, keepRow: number) => {
    const next = new Set(excludedCands);
    for (const g of cluster.groups) {
      let kept: any = null;
      for (const c of g.items) {
        const cRow = Number(c.key.split('|')[0]);
        if (kept === null && cRow === keepRow) {
          next.delete(c.key);
          kept = c;
        } else {
          next.add(c.key);
        }
      }
      if (!kept && g.items.length) {
        next.delete(g.items[0].key);
        for (const c of g.items.slice(1)) next.add(c.key);
      }
    }
    setExcludedCands(next);
  };

  return (
    <div className="min-h-screen bg-background p-6 md:p-12 font-sans text-foreground">
      <div className="max-w-6xl mx-auto space-y-8">
        <div className="flex justify-center">
          <AppToggle
            active="exfo"
            publicMode={publicMode}
            keys={publicMode ? ["ponpower", "exfo"] : undefined}
          />
        </div>

        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-border pb-6">
          <div>
            <h1 className="text-3xl md:text-4xl font-bold tracking-tight bg-gradient-to-r from-primary to-blue-400 bg-clip-text text-transparent">
              F2 Job Creator
            </h1>
            <p className="text-muted-foreground mt-2">
              Convert PON test sheet → Exfo CSV.
            </p>
          </div>
          <div className="flex items-center gap-3">
            <span className={
              `inline-flex items-center px-3 py-1 rounded-full text-xs font-bold border ` +
              (stateBadge.tone === 'ready'
                ? 'border-green-500/40 text-green-400 bg-green-500/10'
                : stateBadge.tone === 'fail'
                ? 'border-red-500/40 text-red-400 bg-red-500/10'
                : 'border-yellow-500/40 text-yellow-400 bg-yellow-500/10')
            }>{stateBadge.text}</span>
            <Button variant="outline" size="sm" onClick={() => setShowSettings(true)} className="gap-2">
              <Settings className="w-4 h-4" /> Settings
            </Button>
          </div>
        </div>

        {/* Step 1: Files */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">1</div>
              Files
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-3">
            <div className="flex items-center gap-3 flex-wrap">
              <label className="cursor-pointer">
                <input
                  type="file"
                  accept=".xlsx,.xlsm,.xlsb,.xls"
                  className="hidden"
                  onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
                />
                <span className="inline-flex items-center justify-center rounded-md text-sm font-medium ring-offset-background transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:pointer-events-none disabled:opacity-50 bg-primary text-primary-foreground hover:bg-primary/90 h-9 px-4 py-2">
                  Choose File
                </span>
              </label>
              <span className="text-sm text-muted-foreground truncate flex-1 min-w-0">{fileName}</span>
            </div>
            <div className="flex gap-2 flex-wrap">
              <Button onClick={handleSaveCsv} size="sm" className="gap-1.5">
                <Save className="w-3.5 h-3.5" /> Save CSV
              </Button>
              <Button onClick={handleClear} variant="outline" size="sm">Clear</Button>
            </div>
            {missing.length > 0 && parsed && (
              <p className="text-xs text-yellow-500">
                Heads up: <b>{missing.join(', ')}</b> in Job Settings must be filled before Save CSV will write the file.
              </p>
            )}
          </CardContent>
        </Card>

        {/* Step 2: Job settings */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">2</div>
              Job settings
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="space-y-1.5">
                <Label>CFAS (Job name)</Label>
                <Input value={cfg.name} onChange={e => setCfg(c => ({ ...c, name: e.target.value }))} placeholder="auto-filled from xlsx" />
              </div>
              <div className="space-y-1.5">
                <Label>Tech UID & Domain</Label>
                <div className="flex gap-2">
                  <Input value={cfg.techID} onChange={e => setCfg(c => ({ ...c, techID: e.target.value }))} placeholder="Tech UID" />
                  <Input value={cfg.domain} onChange={e => setCfg(c => ({ ...c, domain: e.target.value }))} className="w-32" />
                </div>
              </div>
              <div className="space-y-1.5">
                <Label>CLLI (ALoc)</Label>
                <Input value={cfg.aloc} onChange={e => setCfg(c => ({ ...c, aloc: e.target.value }))} placeholder="e.g. MILWWI10" />
              </div>
              <div className="space-y-1.5">
                <Label>CO CLLI (ZLoc)</Label>
                <Input value={cfg.zloc} onChange={e => setCfg(c => ({ ...c, zloc: e.target.value }))} placeholder="e.g. MILWWI10" />
              </div>
              <div className="space-y-1.5">
                <Label>PFP CLLI</Label>
                <Input value={cfg.wcc} onChange={e => setCfg(c => ({ ...c, wcc: e.target.value }))} placeholder="e.g. MILWWI10PFP" />
              </div>
              <div className="space-y-1.5">
                <Label>Test Type</Label>
                <Select value={cfg.testType} onValueChange={v => setCfg(c => ({ ...c, testType: v as any }))}>
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="iOLM">iOLM</SelectItem>
                    <SelectItem value="OTDR">OTDR</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-1.5">
                <div className="flex items-baseline justify-between gap-2">
                  <Label>iOLM test config</Label>
                  {maxDistance != null && (() => {
                    const shortLink = maxDistance <= IOLM_SHORT_LINK_FT;
                    const recPreset = shortLink ? 'under' : 'over';
                    return (
                      <span className={`text-[11px] ${shortLink ? 'text-green-400' : 'text-yellow-500'}`}>
                        Recommended: <b>{shortLink ? 'PON SHORT LINK' : 'PON'}</b>
                        {' · furthest '}{Math.round(maxDistance).toLocaleString()} ft
                        {cfg.iolmPreset !== recPreset && (
                          <>
                            {' · '}
                            <button
                              type="button"
                              onClick={() => setCfg(c => ({ ...c, iolmPreset: recPreset }))}
                              className="text-primary hover:underline"
                            >Apply</button>
                          </>
                        )}
                      </span>
                    );
                  })()}
                </div>
                <Select value={cfg.iolmPreset} onValueChange={v => setCfg(c => ({ ...c, iolmPreset: v as any }))}>
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="over">Span over 8200' → ATT F2 PON.iolmcfg</SelectItem>
                    <SelectItem value="under">Span under 8200' → ATT F2 PON SHORT LINK.iolmcfg</SelectItem>
                    <SelectItem value="none">No configuration</SelectItem>
                    <SelectItem value="custom">Custom…</SelectItem>
                  </SelectContent>
                </Select>
                {cfg.iolmPreset === 'custom' && (
                  <div className="flex gap-2 mt-2">
                    <Input value={cfg.iolmCustom} onChange={e => setCfg(c => ({ ...c, iolmCustom: e.target.value }))} placeholder="config name" />
                    <Input value={cfg.iolmCustomExt} onChange={e => setCfg(c => ({ ...c, iolmCustomExt: e.target.value }))} className="w-28" />
                  </div>
                )}
              </div>
              <div className="space-y-1.5">
                <Label>OPM test config</Label>
                <Select value={cfg.opmPreset} onValueChange={v => setCfg(c => ({ ...c, opmPreset: v as any }))}>
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="t24">ATT F2 Terminal -24dBm</SelectItem>
                    <SelectItem value="p21">ATT F2 PFP -21dBm</SelectItem>
                    <SelectItem value="none">No configuration</SelectItem>
                    <SelectItem value="custom">Custom…</SelectItem>
                  </SelectContent>
                </Select>
                {cfg.opmPreset === 'custom' && (
                  <div className="flex gap-2 mt-2">
                    <Input value={cfg.opmCustom} onChange={e => setCfg(c => ({ ...c, opmCustom: e.target.value }))} placeholder="config name" />
                    <Input value={cfg.opmCustomExt} onChange={e => setCfg(c => ({ ...c, opmCustomExt: e.target.value }))} className="w-28" />
                  </div>
                )}
              </div>
            </div>
            <div className="space-y-1">
              <p className="text-xs text-muted-foreground">Combined testConfigurations (preview)</p>
              <code className="block text-xs bg-secondary/40 border border-border rounded p-2 font-mono break-all">
                {composeTestConfigurations(cfg) || '(blank)'}
              </code>
            </div>
          </CardContent>
        </Card>

        {/* Step 3: Summaries */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">3</div>
              File summaries
            </CardTitle>
          </CardHeader>
          <CardContent className="grid grid-cols-1 md:grid-cols-2 gap-6 text-sm">
            <div>
              <h3 className="text-xs uppercase text-muted-foreground mb-2">Input .xlsx</h3>
              {parsed ? (
                <dl className="grid grid-cols-[160px_1fr] gap-y-1 gap-x-3">
                  <dt className="text-muted-foreground">CFAS</dt><dd>{parsed.project ?? '(none)'}</dd>
                  <dt className="text-muted-foreground">PFP</dt>
                  <dd>{parsed.pfpName ? `${parsed.pfpName} — ${parseAddress(parsed.pfpName)}` : '(not found)'}</dd>
                  {maxDistance != null && (<>
                    <dt className="text-muted-foreground">Furthest from PFP</dt>
                    <dd>{formatDistance(maxDistance)}</dd>
                  </>)}
                  <dt className="text-muted-foreground">Splitter count</dt><dd>{parsed.meta.splitterCount ?? ''}</dd>
                  <dt className="text-muted-foreground">Terminals</dt><dd>{parsed.meta.terminals ?? ''}</dd>
                  <dt className="text-muted-foreground">Total strands</dt><dd>{parsed.meta.totalStrands ?? ''}</dd>
                  <dt className="text-muted-foreground">Power tests</dt><dd>{parsed.meta.powerTests ?? ''}</dd>
                  <dt className="text-muted-foreground">OTDR tests</dt><dd>{parsed.meta.otdrTests ?? ''}</dd>
                </dl>
              ) : <p className="text-muted-foreground">(no file loaded)</p>}
            </div>
            <div>
              <h3 className="text-xs uppercase text-muted-foreground mb-2">Output .csv</h3>
              {outputRows.length ? (
                <dl className="grid grid-cols-[160px_1fr] gap-y-1 gap-x-3">
                  <dt className="text-muted-foreground">Filename</dt><dd>{`${cfg.name.trim() || 'exportsheet'}_job.csv`}</dd>
                  <dt className="text-muted-foreground">Data rows</dt><dd>{outputRows.length - 1}</dd>
                  <dt className="text-muted-foreground">OPM rows</dt><dd>{outputRows.slice(1).filter(r => r[11] === 'OPM').length}</dd>
                  <dt className="text-muted-foreground">{cfg.testType} rows</dt><dd>{outputRows.slice(1).filter(r => r[12] === cfg.testType).length}</dd>
                </dl>
              ) : <p className="text-muted-foreground">(not generated)</p>}
            </div>
          </CardContent>
        </Card>

        {/* Step 4: Conflicts */}
        {(waldoGroups.length > 0 || fiberClusters.length > 0) && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <div className="w-6 h-6 rounded bg-yellow-500/20 flex items-center justify-center text-yellow-500 text-xs font-bold">
                  <AlertTriangle className="w-3.5 h-3.5" />
                </div>
                Conflicts
              </CardTitle>
              <CardDescription>
                The PON sheet has {waldoGroups.length > 0 && <b>{waldoGroups.length} Waldo ID</b>}
                {waldoGroups.length > 0 && fiberClusters.length > 0 && ' and '}
                {fiberClusters.length > 0 && <b>{fiberClusters.length} fiber ID</b>} cluster{(waldoGroups.length + fiberClusters.length) === 1 ? '' : 's'}.
                Pick which to keep — others are excluded from the CSV.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              {waldoGroups.length > 0 && (
                <section>
                  <h3 className="text-xs uppercase text-muted-foreground mb-2">Waldo ID (terminal-level)</h3>
                  <div className="space-y-3">
                    {waldoGroups.map(g => (
                      <div key={g.waldo} className="border border-border rounded p-3 bg-secondary/30">
                        <p className="text-sm font-semibold mb-2">Waldo <code>{g.waldo}</code> — {g.items.length} terminals</p>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                          {g.items.map((t, i) => {
                            const checked = !effectiveExcludedTerminals.has(t.row);
                            return (
                              <label key={t.row} className={`flex items-start gap-2 p-2 rounded border text-xs cursor-pointer ${checked ? 'border-primary bg-primary/5' : 'border-border'}`}>
                                <input
                                  type="radio"
                                  name={`waldo-${g.waldo}`}
                                  checked={checked}
                                  onChange={() => setOnly(g.items, i, 'waldo')}
                                  className="mt-0.5"
                                />
                                <span className="flex-1">
                                  <div><code className="font-bold">{t.cable} · {t.terminal}</code></div>
                                  <div className="text-muted-foreground">
                                    {t.powerStrand != null && `power ${t.powerStrand} · `}
                                    {t.otdrRaw && `OTDR ${t.otdrRaw} · `}
                                    xlsx row {t.row}
                                  </div>
                                </span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                  </div>
                </section>
              )}
              {fiberClusters.length > 0 && (
                <section>
                  <h3 className="text-xs uppercase text-muted-foreground mb-2">Fiber ID (grouped by terminal)</h3>
                  <div className="space-y-3">
                    {fiberClusters.map(cluster => {
                      const fiberIds = cluster.groups.map(g => g.strand).sort((a, b) => a - b);
                      const fiberLabel = fiberIds.length === 1 ? `Fiber ID ${fiberIds[0]}` : `Fiber IDs ${fiberIds.join(', ')}`;
                      const winner = pickCurrentClusterWinner(cluster, effectiveExcludedCands);
                      const participants = cluster.participantRows
                        .map(r => parsed!.terminals.find(t => t.row === r))
                        .filter(Boolean) as ExfoTerminal[];
                      return (
                        <div key={cluster.participantRows.join('_')} className="border border-border rounded p-3 bg-secondary/30">
                          <p className="text-sm font-semibold mb-2">{fiberLabel} — contested by {participants.length} terminals</p>
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                            {participants.map(t => {
                              const checked = winner === t.row;
                              const pointLines: string[] = [];
                              for (const g of cluster.groups) {
                                for (const c of g.items) {
                                  if (Number(c.key.split('|')[0]) === t.row) {
                                    pointLines.push(`${c.cable}_${c.strand} - ${c.terminal}_${c.fiberIndex} · ${c.role}`);
                                  }
                                }
                              }
                              return (
                                <label key={t.row} className={`flex items-start gap-2 p-2 rounded border text-xs cursor-pointer ${checked ? 'border-primary bg-primary/5' : 'border-border'}`}>
                                  <input
                                    type="radio"
                                    name={`cluster-${cluster.participantRows.join('_')}`}
                                    checked={checked}
                                    onChange={() => handleClusterPick(cluster, t.row)}
                                    className="mt-0.5"
                                  />
                                  <span className="flex-1 min-w-0">
                                    <div><b>Keep all from:</b> <code>{t.terminal}</code></div>
                                    <div className="text-muted-foreground text-[11px]">cable {t.cable} · waldo {t.waldo || '—'} · row {t.row}</div>
                                    <div className="mt-1 space-y-0.5">
                                      {pointLines.map((s, i) => (
                                        <div key={i}><code className="text-[10px] text-muted-foreground">{s}</code></div>
                                      ))}
                                    </div>
                                  </span>
                                </label>
                              );
                            })}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </section>
              )}
            </CardContent>
          </Card>
        )}

        {/* Step 5: Map */}
        {parsed && keptTerminals.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">5</div>
                Terminal map
              </CardTitle>
              <CardDescription>
                Tip: enter the <b>CLLI</b> in Job Settings — it pins city/state from the offline lookup before any geocoding runs.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="grid grid-cols-1 md:grid-cols-[1fr_auto] gap-3 items-end">
                <div className="space-y-1.5">
                  <Label>City / State (added to each address for geocoding)</Label>
                  <Input value={city} onChange={e => setCity(e.target.value)} />
                </div>
                <Button onClick={handleLoadMap} size="sm" className="gap-1.5">
                  <MapPin className="w-3.5 h-3.5" /> {mapLoaded ? 'Refresh map' : 'Load map'}
                </Button>
              </div>
              <div className="flex items-center gap-3 flex-wrap">
                <Button onClick={handleOpenInNewTab} variant="outline" size="sm" className="gap-1.5" disabled={!mapLoaded}>
                  <ExternalLink className="w-3.5 h-3.5" /> Open map in new tab
                </Button>
                <label className="flex items-center gap-2 text-xs cursor-pointer">
                  <Switch checked={showConnections} onCheckedChange={setShowConnections} />
                  Connect pins by strand order
                </label>
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
                    onKeyDown={e => { if (e.key === 'Enter') handleFlashPon(); }}
                  />
                </div>
                <Button onClick={handleFlashPon} variant="outline" size="sm" className="gap-1.5">
                  <Search className="w-3.5 h-3.5" /> Flash
                </Button>
              </div>
              {mapStatus && (
                <p className={`text-xs ${mapStatusKind === 'err' ? 'text-red-400' : mapStatusKind === 'ok' ? 'text-green-400' : 'text-muted-foreground'}`}>
                  {mapStatus}
                </p>
              )}
              <div ref={mapCanvasRef} style={{ width: '100%', height: 460, borderRadius: 8, background: 'hsl(var(--secondary) / 0.4)', display: 'none' }} />
              <details open>
                <summary className="text-sm cursor-pointer text-muted-foreground hover:text-foreground">
                  Addresses ({keptTerminals.length})
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
                      {keptTerminals.map((t, i) => {
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

        {/* Step 6: Output preview */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">6</div>
              Output preview
            </CardTitle>
            <CardDescription>
              Locked until every Job Settings field is filled.
            </CardDescription>
          </CardHeader>
          <CardContent>
            {missing.length > 0 ? (
              <p className="text-xs text-yellow-500">
                Fill in <b>{missing.join(', ')}</b> in Job Settings to unlock the preview.
              </p>
            ) : !outputRows.length ? (
              <p className="text-xs text-muted-foreground">Load a PON test sheet to populate the preview.</p>
            ) : (
              <>
                <div className="max-h-[400px] overflow-auto border border-border rounded">
                  <Table className="text-xs font-mono">
                    <TableHeader className="sticky top-0 bg-secondary/40 backdrop-blur-sm">
                      <TableRow>
                        {outputRows[0].map((h: any, i: number) => (
                          <TableHead key={i}>{h}</TableHead>
                        ))}
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {outputRows.slice(1).map((r: any[], i: number) => (
                        <TableRow key={i}>
                          {r.map((v, j) => (
                            <TableCell key={j}>{v == null ? '' : String(v)}</TableCell>
                          ))}
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
                <p className="text-xs text-muted-foreground mt-2">Showing all {outputRows.length - 1} rows.</p>
              </>
            )}
          </CardContent>
        </Card>
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
                Stored locally in your browser. Restrict the key by HTTP referrer in Google Cloud Console.{' '}
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

      <Dialog open={alertOpen} onOpenChange={setAlertOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>{alertContent.title}</DialogTitle>
          </DialogHeader>
          <div className="text-sm">{alertContent.body}</div>
          <DialogFooter>
            <Button onClick={() => setAlertOpen(false)}>OK</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
