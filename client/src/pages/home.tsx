import { useState, useEffect } from "react";
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
import { useToast } from "@/hooks/use-toast";
import {
  parseExcelFile,
  generateReport,
  parseTerminals,
  computeStaggeredColumns,
  generateConvertedXlsx,
  Terminal,
  ParsedWorkbook,
} from "@/lib/excel";
import { AppToggle } from "@/components/app-toggle";
import { Download, FileJson, Copy, Save, FileSpreadsheet } from "lucide-react";

export default function Home() {
  const { toast } = useToast();
  const [file, setFile] = useState<File | null>(null);
  const [fileHandle, setFileHandle] = useState<FileSystemFileHandle | null>(null);
  const [cableId, setCableId] = useState<string>("");
  const [wireCenterClli, setWireCenterClli] = useState<string>("");
  const [cfas, setCfas] = useState<string>("");
  const [strands, setStrands] = useState<number[]>([]);
  const [averageLoss, setAverageLoss] = useState<string>("-15.5");
  const [jsonOutput, setJsonOutput] = useState<string>("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [parsedWorkbook, setParsedWorkbook] = useState<ParsedWorkbook | null>(null);
  const [staggeredTerminals, setStaggeredTerminals] = useState<Terminal[]>([]);

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
      setIsProcessing(true);
      Promise.all([parseExcelFile(file), parseTerminals(file).catch(() => null)])
        .then(([data, parsed]) => {
          if (data.cableId) setCableId(data.cableId);
          if (data.cfas) setCfas(data.cfas);
          if (data.strands.length > 0) {
            setStrands(data.strands);
            toast({
              title: "Excel Parsed Successfully",
              description: `Found ${data.strands.length} strands. Cable ID: ${data.cableId || 'Not found'}`,
            });
          } else {
            toast({
              variant: "destructive",
              title: "No Data Found",
              description: "Could not find strand data in the uploaded file.",
            });
          }
          if (parsed && parsed.terminals.length > 0) {
            const withStaggered = computeStaggeredColumns(parsed.terminals);
            setParsedWorkbook(parsed);
            setStaggeredTerminals(withStaggered);
          } else {
            setParsedWorkbook(null);
            setStaggeredTerminals([]);
          }
        })
        .catch((err) => {
          console.error(err);
          toast({
            variant: "destructive",
            title: "Error Parsing File",
            description: "Please make sure it's a valid Excel file.",
          });
        })
        .finally(() => setIsProcessing(false));
    }
  }, [file]);

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

  const handleDownloadConverted = () => {
    if (!parsedWorkbook || staggeredTerminals.length === 0) return;
    try {
      const bytes = generateConvertedXlsx(parsedWorkbook, staggeredTerminals);
      const blob = new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const baseName = file?.name.replace(/\.xlsx$/i, "") || "PONSHEET";
      a.href = url;
      a.download = `${baseName}_staggered.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      toast({
        title: "Spreadsheet Downloaded",
        description: "Converted file with staggered columns saved.",
      });
    } catch (e) {
      console.error(e);
      toast({
        variant: "destructive",
        title: "Error Generating Spreadsheet",
        description: "Could not build the converted file.",
      });
    }
  };

  return (
    <div className="min-h-screen bg-background p-6 md:p-12 font-sans text-foreground">
      <div className="max-w-6xl mx-auto space-y-8">

        {/* Top center: app toggle */}
        <div className="flex justify-center">
          <AppToggle active="ponpower" />
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
                  Upload Data
                </CardTitle>
                <CardDescription>
                  Import your field data (.xlsx)
                </CardDescription>
              </CardHeader>
              <CardContent>
                <FileUpload
                  onFileSelect={(f, h) => {
                    setFile(f);
                    setFileHandle(h ?? null);
                  }}
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
                  Original terminals with generated Staggered Port and Staggered Strand columns.
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
                      <TableHead>Waldo ID</TableHead>
                      <TableHead>Terminal</TableHead>
                      <TableHead>Cable ID</TableHead>
                      <TableHead className="text-right">Power Test</TableHead>
                      <TableHead>OTDR Test</TableHead>
                      <TableHead className="text-right">Total Strands</TableHead>
                      <TableHead className="text-right text-primary">Staggered Port</TableHead>
                      <TableHead className="text-right text-primary">Staggered Strand</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {staggeredTerminals.map((t) => (
                      <TableRow key={t.rowIndex}>
                        <TableCell>{t.waldoId}</TableCell>
                        <TableCell className="whitespace-nowrap">{t.terminalName}</TableCell>
                        <TableCell>{t.cableId}</TableCell>
                        <TableCell className="text-right">{t.powerTestStrand}</TableCell>
                        <TableCell>{t.otdrTestStrand}</TableCell>
                        <TableCell className="text-right">{t.totalStrands}</TableCell>
                        <TableCell className="text-right text-primary font-bold">{t.staggeredPort}</TableCell>
                        <TableCell className="text-right text-primary font-bold">{t.staggeredStrand}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center text-muted-foreground py-16 opacity-50">
                <FileSpreadsheet className="w-12 h-12 mb-2" />
                <p>Upload a PON sheet to preview staggered columns</p>
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
