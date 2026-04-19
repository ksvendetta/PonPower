import { useEffect, useMemo, useState } from "react";
import { FileUpload } from "@/components/ui/file-upload";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import { AppToggle } from "@/components/app-toggle";
import {
  cleanReport,
  findMissingStrands,
  formatReport,
  IolmReport,
  PFP_SIZES,
  PfpSize,
} from "@/lib/iolm";
import { saveOrOverwrite } from "@/lib/save-file";
import { Download, FileJson, Copy, Save, AlertTriangle } from "lucide-react";

export default function Iolm() {
  const { toast } = useToast();
  const [file, setFile] = useState<File | null>(null);
  const [fileHandle, setFileHandle] = useState<FileSystemFileHandle | null>(null);
  const [rawReport, setRawReport] = useState<IolmReport | null>(null);
  const [cleanedReport, setCleanedReport] = useState<IolmReport | null>(null);
  const [cableId, setCableId] = useState<string>("");
  const [wireCenterClli, setWireCenterClli] = useState<string>("");
  const [cfas, setCfas] = useState<string>("");
  const [pfpSize, setPfpSize] = useState<PfpSize | null>(null);
  const [sortStrands, setSortStrands] = useState<boolean>(true);
  const [jsonOutput, setJsonOutput] = useState<string>("");
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);

  useEffect(() => {
    const handler = (e: any) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };
    window.addEventListener("beforeinstallprompt", handler);
    return () => window.removeEventListener("beforeinstallprompt", handler);
  }, []);

  const handleInstall = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult: any) => {
        if (choiceResult.outcome === "accepted") {
          setDeferredPrompt(null);
        }
      });
    }
  };

  // Parse JSON on upload
  useEffect(() => {
    if (!file) return;
    (async () => {
      try {
        const text = await file.text();
        const parsed = JSON.parse(text) as IolmReport;
        if (!parsed?.Label || !Array.isArray(parsed.Label.Tested)) {
          throw new Error("Missing Label.Tested[] in JSON");
        }
        setRawReport(parsed);
        const cleaned = cleanReport(parsed);
        setCleanedReport(cleaned);
        setCableId(cleaned.Label.cableId ?? "");
        setWireCenterClli(cleaned.Label.wireCenterClli ?? "");
        setCfas(cleaned.Label.cfas ?? "");
        const failCount = (parsed.Label.Tested ?? []).filter(
          (t) => (t.passFail ?? "").toLowerCase() === "fail"
        ).length;
        toast({
          title: "JSON Parsed Successfully",
          description: `Loaded ${cleaned.Label.Tested.length} strand(s). ${failCount} fail→pass applied, events stripped.`,
        });
      } catch (err: any) {
        console.error(err);
        toast({
          variant: "destructive",
          title: "Invalid JSON",
          description: err?.message ?? "Could not parse the uploaded file.",
        });
      }
    })();
  }, [file]);

  // Keep cleanedReport in sync with editable header fields
  useEffect(() => {
    if (!cleanedReport) return;
    const next: IolmReport = {
      ...cleanedReport,
      Label: {
        ...cleanedReport.Label,
        cableId,
        wireCenterClli,
        cfas,
      },
    };
    setCleanedReport(next);
    setJsonOutput(formatReport(next));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cableId, wireCenterClli, cfas]);

  useEffect(() => {
    if (!cleanedReport) return;
    let report: IolmReport = cleanedReport;
    if (sortStrands && Array.isArray(report.Label?.Tested)) {
      report = {
        ...report,
        Label: {
          ...report.Label,
          Tested: [...report.Label.Tested].sort(
            (a, b) => (a.strand ?? 0) - (b.strand ?? 0)
          ),
        },
      };
    }
    setJsonOutput(formatReport(report));
  }, [cleanedReport, sortStrands]);

  const missingStrands = useMemo(
    () => findMissingStrands(cleanedReport, pfpSize),
    [cleanedReport, pfpSize]
  );

  const testedCount = cleanedReport?.Label?.Tested?.length ?? 0;

  const handleOverwrite = async () => {
    if (!cleanedReport || !file) return;
    try {
      const json = JSON.parse(jsonOutput);
      const pretty = formatReport(json as IolmReport);
      const result = await saveOrOverwrite(pretty, file.name, fileHandle);
      toast({
        title:
          result === "overwritten"
            ? "File Overwritten"
            : result === "saved-as"
            ? "Saved"
            : "Downloaded",
        description:
          result === "overwritten"
            ? `Wrote ${file.name} in place.`
            : `Saved as ${file.name}.`,
      });
    } catch (err: any) {
      toast({
        variant: "destructive",
        title: "Could Not Overwrite",
        description: err?.message ?? "Invalid JSON or permission denied.",
      });
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(jsonOutput);
    toast({ title: "Copied to Clipboard", description: "You can paste it anywhere now." });
  };

  return (
    <div className="min-h-screen bg-background p-6 md:p-12 font-sans text-foreground">
      <div className="max-w-6xl mx-auto space-y-8">

        {/* Top center: app toggle */}
        <div className="flex justify-center">
          <AppToggle active="iolm" />
        </div>

        {/* Header */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-border pb-6">
          <div>
            <h1 className="text-3xl md:text-4xl font-bold tracking-tight bg-gradient-to-r from-primary to-blue-400 bg-clip-text text-transparent">
              IOLM Cleanup
            </h1>
            <p className="text-muted-foreground mt-2">
              Offline tool for cleaning OTDR/IOLM JSON reports.
            </p>
          </div>
          <div className="flex items-center gap-2 px-3 py-1 rounded-full bg-secondary/50 text-xs font-mono text-muted-foreground border border-white/5">
            <div className="w-2 h-2 rounded-full bg-green-500 animate-pulse" />
            SYSTEM ONLINE
          </div>
          {deferredPrompt && (
            <Button onClick={handleInstall} size="sm" className="gap-2 bg-primary text-primary-foreground hover:bg-primary/90">
              <Download className="w-4 h-4" />
              Install App
            </Button>
          )}
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">

          {/* Left Column: Inputs */}
          <div className="lg:col-span-5 space-y-6">

            {/* Step 1: Upload */}
            <Card className="border-border/50 bg-card/50 backdrop-blur-sm shadow-lg">
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">1</div>
                  Upload JSON
                </CardTitle>
                <CardDescription>
                  Import your IOLM/OTDR report (.json)
                </CardDescription>
              </CardHeader>
              <CardContent>
                <FileUpload
                  onFileSelect={(f, h) => {
                    setFile(f);
                    setFileHandle(h ?? null);
                  }}
                  accept=".json,application/json"
                  label="Upload JSON"
                  helper="Drag and drop your .json file here, or click to browse"
                  pickerTypes={[
                    {
                      description: "JSON",
                      accept: { "application/json": [".json"] },
                    },
                  ]}
                />
              </CardContent>
            </Card>

            {/* Step 2: Configure */}
            <Card className="border-border/50 bg-card/50 backdrop-blur-sm shadow-lg">
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <div className="w-6 h-6 rounded bg-primary/20 flex items-center justify-center text-primary text-xs font-bold">2</div>
                  Configuration
                </CardTitle>
                <CardDescription>
                  Values extracted from the uploaded JSON
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
                      placeholder="Extracted from JSON..."
                      className="font-mono bg-secondary/20 border-border focus:border-primary"
                    />
                  </div>

                  <div className="space-y-2 col-span-2">
                    <Label htmlFor="wireCenterClli">Wire Center CLLI</Label>
                    <Input
                      id="wireCenterClli"
                      value={wireCenterClli}
                      onChange={(e) => setWireCenterClli(e.target.value)}
                      placeholder="Extracted from JSON..."
                      className="font-mono bg-secondary/20 border-border focus:border-primary"
                    />
                  </div>

                  <div className="space-y-2 col-span-2">
                    <Label htmlFor="cfas">Project #</Label>
                    <Input
                      id="cfas"
                      value={cfas}
                      onChange={(e) => setCfas(e.target.value)}
                      placeholder="Extracted from JSON..."
                      className="font-mono bg-secondary/20 border-border focus:border-primary"
                    />
                  </div>

                  <div className="space-y-2 col-span-2">
                    <Label htmlFor="pfpSize">PFP Size</Label>
                    <Select
                      value={pfpSize ? String(pfpSize) : undefined}
                      onValueChange={(v) => setPfpSize(Number(v) as PfpSize)}
                    >
                      <SelectTrigger
                        id="pfpSize"
                        className="font-mono bg-secondary/20 border-border focus:border-primary"
                      >
                        <SelectValue placeholder="Select PFP size" />
                      </SelectTrigger>
                      <SelectContent>
                        {PFP_SIZES.map((n) => (
                          <SelectItem key={n} value={String(n)}>
                            {n}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    <p className="text-xs text-muted-foreground">
                      Used to compute which fiber strands are missing from the test results.
                    </p>
                  </div>
                </div>

                <div className="pt-4 border-t border-border space-y-3">
                  <div className="flex items-center gap-2">
                    <Checkbox
                      id="sortStrands"
                      checked={sortStrands}
                      onCheckedChange={(v) => setSortStrands(v === true)}
                    />
                    <Label htmlFor="sortStrands" className="text-sm cursor-pointer">
                      Organize strands in order (ascending)
                    </Label>
                  </div>
                  <div className="flex justify-between items-center text-sm">
                    <span className="text-muted-foreground">Detected Strands:</span>
                    <span className="font-mono font-bold text-primary">{testedCount}</span>
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Missing Strands Window */}
            <Card className="border-border/50 bg-card/50 backdrop-blur-sm shadow-lg">
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 text-amber-400" />
                  Missing Strands
                </CardTitle>
                <CardDescription>
                  {pfpSize
                    ? `Strands in 1..${pfpSize} not present in the uploaded test results`
                    : "Select a PFP size to compute missing strands"}
                </CardDescription>
              </CardHeader>
              <CardContent>
                {!pfpSize || !cleanedReport ? (
                  <p className="text-sm text-muted-foreground italic">
                    Upload a JSON file and pick a PFP size to see missing strands.
                  </p>
                ) : missingStrands.length === 0 ? (
                  <p className="text-sm text-green-500 font-semibold">
                    All {pfpSize} strands are present.
                  </p>
                ) : (
                  <>
                    <div className="flex justify-between items-center text-sm mb-2">
                      <span className="text-muted-foreground">Missing:</span>
                      <span className="font-mono font-bold text-amber-400">
                        {missingStrands.length}
                      </span>
                    </div>
                    <div className="max-h-56 overflow-auto rounded-md border border-border bg-background/50 p-2 font-mono text-xs text-amber-200 leading-relaxed">
                      {missingStrands.join(", ")}
                    </div>
                  </>
                )}
              </CardContent>
            </Card>
          </div>

          {/* Right Column: Preview */}
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
                    <Button
                      onClick={handleOverwrite}
                      size="sm"
                      className="h-8 text-xs gap-1.5"
                      disabled={!jsonOutput || !file}
                    >
                      <Save className="w-3.5 h-3.5" /> Overwrite
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

        </div>
      </div>
    </div>
  );
}
