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
import { useToast } from '@/hooks/use-toast';
import { AppToggle } from '@/components/app-toggle';
import { Settings, Save, FileSpreadsheet, AlertTriangle, MapPin } from 'lucide-react';
import {
  parseExfoXlsx, ExfoTerminal, ExfoXlsxParse, ExfoJobConfig, resolveJobConfig,
  composeTestConfigurations, buildCandidates, findWaldoDuplicateGroups,
  findFiberDuplicateGroups, candidatesToRows, rowsToCsv, formatDistance,
  parseAddress, IOLM_SHORT_LINK_FT, resolveCityFromCLLI,
} from '@/lib/exfo';
import {
  GeocodeHit, loadGeocodeCache, renderEmbeddedMap,
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

export default function Exfo() {
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
  const mapCanvasRef = useRef<HTMLDivElement>(null);
  const mapStateRef = useRef<{
    gmap: any | null; gmarkers: any[]; connections: any[]; showConnections: boolean;
  }>({ gmap: null, gmarkers: [], connections: [], showConnections: false });
  const geocodeCacheRef = useRef<Map<string, GeocodeHit>>(new Map());

  useEffect(() => {
    const k = localStorage.getItem('f2job.mapApiKey');
    if (k) setApiKey(k);
    geocodeCacheRef.current = loadGeocodeCache();
  }, []);

  useEffect(() => {
    localStorage.setItem('f2job.mapApiKey', apiKey);
  }, [apiKey]);

  // Auto-fill CFAS from sheet
  useEffect(() => {
    if (parsed?.project && !cfg.name) {
      setCfg(c => ({ ...c, name: parsed.project! }));
    }
  }, [parsed]);

  // Auto-resolve city from CLLI
  useEffect(() => {
    const hit = resolveCityFromCLLI(cfg.aloc) || resolveCityFromCLLI(cfg.zloc) || resolveCityFromCLLI(cfg.wcc);
    if (hit) setCity(hit.city);
  }, [cfg.aloc, cfg.zloc, cfg.wcc]);

  // Auto-set wcc from PFP name if blank
  useEffect(() => {
    if (parsed?.pfpName && !cfg.wcc && cfg.aloc) {
      setCfg(c => ({ ...c, wcc: cfg.aloc + 'PFP' }));
    }
  }, [parsed?.pfpName, cfg.aloc]);

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
      toast({ title: 'Loaded', description: `${p.terminals.length} terminals from ${p.sheetName}.` });
    } catch (e: any) {
      toast({ variant: 'destructive', title: 'Parse error', description: e?.message || String(e) });
    }
  };

  const missing = useMemo(
    () => REQUIRED.filter(([k]) => !String(cfg[k] || '').trim()).map(([, label]) => label),
    [cfg]
  );

  // Auto-prune Waldo dup exclusions
  const waldoGroups = useMemo(
    () => parsed ? findWaldoDuplicateGroups(parsed.terminals) : [],
    [parsed]
  );
  const effectiveExcludedTerminals = useMemo(() => {
    const next = new Set(excludedTerminals);
    if (parsed) {
      const validRows = new Set(parsed.terminals.map(t => t.row));
      for (const r of Array.from(next)) if (!validRows.has(r)) next.delete(r);
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

  const effectiveExcludedCands = useMemo(() => {
    const next = new Set(excludedCands);
    const validKeys = new Set(candidates.map(c => c.key));
    for (const k of Array.from(next)) if (!validKeys.has(k)) next.delete(k);
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

  // Auto-recommend iOLM preset based on max distance
  useEffect(() => {
    if (maxDistance == null) return;
    const shortLink = maxDistance <= IOLM_SHORT_LINK_FT;
    const target = shortLink ? 'under' : 'over';
    if (cfg.iolmPreset !== target) {
      setCfg(c => ({ ...c, iolmPreset: target }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [maxDistance]);

  const handleSaveCsv = () => {
    if (missing.length || !outputRows.length) return;
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
    if (mapStateRef.current.gmap) {
      for (const m of mapStateRef.current.gmarkers) m.setMap(null);
      mapStateRef.current.gmarkers = [];
      for (const l of mapStateRef.current.connections) l.setMap(null);
      mapStateRef.current.connections = [];
      mapStateRef.current.gmap = null;
    }
    if (mapCanvasRef.current) mapCanvasRef.current.style.display = 'none';
  };

  const handleLoadMap = async () => {
    if (!apiKey.trim()) {
      setMapStatus('Add a Google Maps API key in Settings.');
      setMapStatusKind('err');
      return;
    }
    if (!mapCanvasRef.current || !parsed) return;
    try {
      const result = await renderEmbeddedMap(
        mapCanvasRef.current,
        mapStateRef.current,
        {
          apiKey: apiKey.trim(),
          city: city.trim(),
          pfpName: parsed.pfpName,
          terminals: keptTerminals,
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

  const setOnly = (group: { items: Array<{ row?: number; key?: string }> }, idx: number, type: 'waldo' | 'fiber') => {
    if (type === 'waldo') {
      const next = new Set(excludedTerminals);
      group.items.forEach((it: any, i) => {
        if (i === idx) next.delete(it.row);
        else next.add(it.row);
      });
      setExcludedTerminals(next);
    } else {
      const next = new Set(excludedCands);
      group.items.forEach((it: any, i) => {
        if (i === idx) next.delete(it.key);
        else next.add(it.key);
      });
      setExcludedCands(next);
    }
  };

  return (
    <div className="min-h-screen bg-background p-6 md:p-12 font-sans text-foreground">
      <div className="max-w-6xl mx-auto space-y-8">
        <div className="flex justify-center">
          <AppToggle active="exfo" />
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
          <Button variant="outline" size="sm" onClick={() => setShowSettings(true)} className="gap-2">
            <Settings className="w-4 h-4" /> Settings
          </Button>
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
              <Button onClick={handleSaveCsv} disabled={!parsed || missing.length > 0 || !outputRows.length} size="sm" className="gap-1.5">
                <Save className="w-3.5 h-3.5" /> Save CSV
              </Button>
              <Button onClick={handleClear} variant="outline" size="sm">Clear</Button>
            </div>
            {missing.length > 0 && parsed && (
              <p className="text-xs text-yellow-500">
                Save CSV is disabled — fill in <b>{missing.join(', ')}</b> in Job Settings.
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
                  {maxDistance != null && (
                    <span className="text-[11px] text-muted-foreground">
                      Recommended: {maxDistance <= IOLM_SHORT_LINK_FT ? 'SHORT LINK' : 'PON'} (furthest {Math.round(maxDistance).toLocaleString()} ft)
                    </span>
                  )}
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
                <dl className="grid grid-cols-[140px_1fr] gap-y-1 gap-x-3">
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
                <dl className="grid grid-cols-[140px_1fr] gap-y-1 gap-x-3">
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
        {(waldoGroups.length > 0 || fiberGroups.length > 0) && (
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
                {waldoGroups.length > 0 && fiberGroups.length > 0 && ' and '}
                {fiberGroups.length > 0 && <b>{fiberGroups.length} fiber ID</b>} duplicate{waldoGroups.length + fiberGroups.length === 1 ? '' : 's'}.
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
                        <p className="text-sm font-semibold mb-2">Waldo <code>{g.waldo}</code></p>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                          {g.items.map((t, i) => {
                            const checked = !effectiveExcludedTerminals.has(t.row);
                            return (
                              <label key={t.row} className={`flex items-start gap-2 p-2 rounded border text-xs cursor-pointer ${checked ? 'border-primary bg-primary/5' : 'border-border'}`}>
                                <input
                                  type="radio"
                                  name={`waldo-${g.waldo}`}
                                  checked={checked}
                                  onChange={() => setOnly(g, i, 'waldo')}
                                  className="mt-0.5"
                                />
                                <span className="flex-1">
                                  <div><b>{t.terminal}</b></div>
                                  <div className="text-muted-foreground">Cable {t.cable}, Power {t.powerStrand}, OTDR {t.otdrRaw}</div>
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
              {fiberGroups.length > 0 && (
                <section>
                  <h3 className="text-xs uppercase text-muted-foreground mb-2">Fiber ID (row-level)</h3>
                  <div className="space-y-3">
                    {fiberGroups.map(g => (
                      <div key={g.strand} className="border border-border rounded p-3 bg-secondary/30">
                        <p className="text-sm font-semibold mb-2">Strand <code>{g.strand}</code></p>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                          {g.items.map((c, i) => {
                            const checked = !effectiveExcludedCands.has(c.key);
                            return (
                              <label key={c.key} className={`flex items-start gap-2 p-2 rounded border text-xs cursor-pointer ${checked ? 'border-primary bg-primary/5' : 'border-border'}`}>
                                <input
                                  type="radio"
                                  name={`fiber-${g.strand}`}
                                  checked={checked}
                                  onChange={() => setOnly(g, i, 'fiber')}
                                  className="mt-0.5"
                                />
                                <span className="flex-1">
                                  <div><b>{c.terminal}</b> · fiber {c.fiberIndex} · <span className="text-primary">{c.role}</span></div>
                                  <div className="text-muted-foreground">Cable {c.cable}</div>
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
                City/state added to each address for geocoding.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="grid grid-cols-1 md:grid-cols-[1fr_auto] gap-3 items-end">
                <div className="space-y-1.5">
                  <Label>City / State</Label>
                  <Input value={city} onChange={e => setCity(e.target.value)} />
                </div>
                <Button onClick={handleLoadMap} size="sm" className="gap-1.5">
                  <MapPin className="w-3.5 h-3.5" /> {mapLoaded ? 'Refresh map' : 'Load map'}
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
    </div>
  );
}
