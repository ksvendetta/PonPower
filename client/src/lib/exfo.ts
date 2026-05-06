import * as XLSX from 'xlsx';

export interface ExfoTerminal {
  row: number;
  waldo: string;
  terminal: string | null;
  cable: string | null;
  powerStrand: number | null;
  otdrRaw: string;
  otdrStrands: number[];
  total: number | null;
}

export interface ExfoXlsxParse {
  project: string | null;
  terminals: ExfoTerminal[];
  meta: {
    splitterCount?: any;
    terminals?: any;
    totalStrands?: any;
    powerTests?: any;
    otdrTests?: any;
  };
  sheetName: string;
  pfpName: string | null;
}

export interface ExfoJobConfig {
  name: string;
  techID: string;
  domain: string;
  aloc: string;
  zloc: string;
  wcc: string;
  testType: 'iOLM' | 'OTDR';
  iolmPreset: 'over' | 'under' | 'none' | 'custom';
  iolmCustom: string;
  iolmCustomExt: string;
  opmPreset: 't24' | 'p21' | 'none' | 'custom';
  opmCustom: string;
  opmCustomExt: string;
}

export interface ExfoCandidate {
  key: string;
  terminalRow: number;
  terminal: string | null;
  cable: string | null;
  strand: number;
  fiberIndex: number;
  tt01: string;
  tt02: string;
  role: string;
}

export const CSV_HEADER = [
  'name', 'assignees', 'company', 'customer', 'dueDate', 'testPointName',
  'identifier_Cable ID', 'identifier_Fiber ID', 'identifier_ALoc',
  'identifier_ZLoc', 'identifier_WireCenterClli', 'testType_01', 'testType_02',
  'testConfigurations',
];

export function csvEscape(v: any): string {
  if (v === null || v === undefined) return '';
  const s = String(v);
  if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

export function rowsToCsv(rows: any[][]): string {
  return rows.map(r => r.map(csvEscape).join(',')).join('\r\n') + '\r\n';
}

export function parseStrandList(s: any): number[] {
  if (s === null || s === undefined || s === '') return [];
  const out: number[] = [];
  String(s).split(/[,/]/).forEach(part => {
    part = part.trim();
    if (!part) return;
    if (part.includes('-')) {
      const [a, b] = part.split('-').map(n => parseInt(String(n).trim(), 10));
      if (!Number.isNaN(a) && !Number.isNaN(b)) {
        const lo = Math.min(a, b), hi = Math.max(a, b);
        for (let i = lo; i <= hi; i++) out.push(i);
      }
    } else {
      const n = parseInt(part, 10);
      if (!Number.isNaN(n)) out.push(n);
    }
  });
  return out.sort((a, b) => a - b);
}

export function parseExfoXlsx(arrayBuffer: ArrayBuffer): ExfoXlsxParse {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = wb.SheetNames.find(n => /PON|TEST|SHEET/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true }) as any[][];
  const get = (r: number, c: number) => {
    const row = aoa[r - 1];
    return row ? (row[c - 1] ?? null) : null;
  };

  let project: any = get(2, 24);
  if (project === null || project === '') {
    const row2 = aoa[1] || [];
    for (let j = 0; j < row2.length; j++) {
      if (row2[j] !== null && row2[j] !== '') {
        project = String(row2[j]);
        break;
      }
    }
  }
  if (project !== null && project !== undefined) project = String(project);

  let detailHeaderRow = -1;
  for (let r = 1; r <= aoa.length; r++) {
    for (let c = 1; c <= (aoa[r - 1]?.length || 0); c++) {
      const v = get(r, c);
      if (v && String(v).toUpperCase().includes('WALDO ID')) {
        detailHeaderRow = r;
        break;
      }
    }
    if (detailHeaderRow !== -1) break;
  }
  if (detailHeaderRow === -1) throw new Error('Could not locate "WALDO ID" header row in the xlsx.');

  const colMap: { [k: string]: number } = {};
  const headerRow = aoa[detailHeaderRow - 1] || [];
  for (let j = 0; j < headerRow.length; j++) {
    const v = headerRow[j];
    if (!v) continue;
    const up = String(v).toUpperCase().replace(/\s+/g, ' ').trim();
    if (up.includes('WALDO ID')) colMap.waldo = j + 1;
    else if (up === 'TERMINAL') colMap.terminal = j + 1;
    else if (up === 'CABLE ID') colMap.cable = j + 1;
    else if (up.includes('POWER TEST STRAND')) colMap.power = j + 1;
    else if (up.includes('OTDR TEST')) colMap.otdr = j + 1;
    else if (up.includes('TOTAL') && up.includes('STRAND')) colMap.total = j + 1;
  }
  for (const k of ['waldo', 'terminal', 'cable', 'power', 'otdr']) {
    if (!colMap[k]) throw new Error('Could not find column "' + k + '" in the WALDO TERMINAL DETAIL header.');
  }

  const terminals: ExfoTerminal[] = [];
  for (let r = detailHeaderRow + 1; r <= aoa.length; r += 3) {
    const waldo = get(r, colMap.waldo);
    const terminal = get(r, colMap.terminal);
    const cable = get(r, colMap.cable);
    const power = get(r, colMap.power);
    const otdr = get(r, colMap.otdr);
    const total = colMap.total ? get(r, colMap.total) : null;
    if (!terminal && !cable && !power) continue;
    terminals.push({
      row: r,
      waldo: waldo == null || waldo === '' ? '' : String(waldo),
      terminal,
      cable,
      powerStrand: power == null || power === '' ? null : parseInt(String(power), 10),
      otdrRaw: otdr == null ? '' : String(otdr),
      otdrStrands: parseStrandList(otdr),
      total: total == null || total === '' ? null : parseInt(String(total), 10),
    });
  }

  let pfpName: string | null = null;
  for (let r = 1; r <= detailHeaderRow; r++) {
    const row = aoa[r - 1] || [];
    for (const v of row) {
      if (!v) continue;
      const s = String(v).trim();
      if (/^[A-Z]+\s+.+\s+PFP$/.test(s)) {
        pfpName = s;
        break;
      }
    }
    if (pfpName) break;
  }

  const meta: any = {};
  const metaRC: { [k: string]: [number, number] } = {};
  for (let r = 1; r <= aoa.length; r++) {
    for (let c = 1; c <= (aoa[r - 1]?.length || 0); c++) {
      const v = get(r, c);
      if (!v) continue;
      const up = String(v).toUpperCase().trim();
      if (up === 'SPLITTER COUNT') metaRC.splitterCount = [r, c];
      else if (up === 'TERMINALS') metaRC.terminals = [r, c];
      else if (up === 'TOTAL STRANDS') metaRC.totalStrands = [r, c];
      else if (up === 'POWER TESTS') metaRC.powerTests = [r, c];
      else if (up === 'OTDR TESTS') metaRC.otdrTests = [r, c];
    }
  }
  const readBelow = (rc: [number, number] | undefined) => {
    if (!rc) return null;
    for (let d = 1; d <= 4; d++) {
      const v = get(rc[0] + d, rc[1]);
      if (v !== null && v !== undefined && v !== '') return v;
    }
    return null;
  };
  meta.splitterCount = readBelow(metaRC.splitterCount);
  meta.terminals = readBelow(metaRC.terminals);
  meta.totalStrands = readBelow(metaRC.totalStrands);
  meta.powerTests = readBelow(metaRC.powerTests);
  meta.otdrTests = readBelow(metaRC.otdrTests);

  return { project, terminals, meta, sheetName, pfpName };
}

export function resolveIolm(c: ExfoJobConfig): string {
  switch (c.iolmPreset) {
    case 'over': return 'ATT F2 PON.iolmcfg';
    case 'under': return 'ATT F2 PON SHORT LINK.iolmcfg';
    case 'none': return '';
    case 'custom': return (c.iolmCustom || '') + (c.iolmCustomExt || '');
  }
}
export function resolveOpm(c: ExfoJobConfig): string {
  switch (c.opmPreset) {
    case 't24': return 'ATT F2 Terminal -24dBm.opmcfg';
    case 'p21': return 'ATT F2 PFP -21dBm.opmcfg';
    case 'none': return '';
    case 'custom': return (c.opmCustom || '') + (c.opmCustomExt || '');
  }
}
export function composeTestConfigurations(c: ExfoJobConfig): string {
  const iolm = resolveIolm(c);
  const opm = resolveOpm(c);
  let out = '';
  if (c.testType !== 'OTDR') out = iolm;
  if (opm) out = out ? out + '|' + opm : opm;
  return out;
}

export interface ResolvedJobCfg {
  name: string;
  assignees: string;
  company: string;
  customer: string;
  dueDate: string;
  aloc: string;
  zloc: string;
  wcc: string;
  testType: 'iOLM' | 'OTDR';
  testConfigs: string;
}

export function resolveJobConfig(c: ExfoJobConfig): ResolvedJobCfg {
  const techID = c.techID.trim().toUpperCase();
  const domain = c.domain.trim();
  return {
    name: c.name.trim(),
    assignees: techID && domain ? techID + domain : '',
    company: 'AT&T',
    customer: '',
    dueDate: '',
    aloc: c.aloc.trim(),
    zloc: c.zloc.trim(),
    wcc: c.wcc.trim(),
    testType: c.testType,
    testConfigs: composeTestConfigurations(c),
  };
}

export function buildCandidates(
  terminals: ExfoTerminal[],
  testType: 'iOLM' | 'OTDR',
  excludedTerminals: Set<number>
): ExfoCandidate[] {
  const strandTT = testType === 'OTDR' ? 'OTDR' : 'iOLM';
  const cands: ExfoCandidate[] = [];
  for (const t of terminals) {
    if (excludedTerminals.has(t.row)) continue;
    if (t.powerStrand !== null && !Number.isNaN(t.powerStrand) && t.powerStrand > 0) {
      cands.push({
        key: `${t.row}|1`,
        terminalRow: t.row,
        terminal: t.terminal,
        cable: t.cable,
        strand: t.powerStrand,
        fiberIndex: 1,
        tt01: 'OPM',
        tt02: '',
        role: 'OPM',
      });
    }
    t.otdrStrands.forEach((s, k) => {
      const idx = k + 2;
      cands.push({
        key: `${t.row}|${idx}`,
        terminalRow: t.row,
        terminal: t.terminal,
        cable: t.cable,
        strand: s,
        fiberIndex: idx,
        tt01: '',
        tt02: strandTT,
        role: strandTT,
      });
    });
  }
  return cands;
}

export interface WaldoGroup {
  waldo: string;
  items: ExfoTerminal[];
}
export interface FiberGroup {
  strand: number;
  items: ExfoCandidate[];
}

export function findWaldoDuplicateGroups(terminals: ExfoTerminal[]): WaldoGroup[] {
  const byWaldo = new Map<string, ExfoTerminal[]>();
  for (const t of terminals) {
    if (!t.waldo) continue;
    if (!byWaldo.has(t.waldo)) byWaldo.set(t.waldo, []);
    byWaldo.get(t.waldo)!.push(t);
  }
  const groups: WaldoGroup[] = [];
  Array.from(byWaldo.entries()).forEach(([waldo, list]) => {
    if (list.length > 1) groups.push({ waldo, items: list });
  });
  groups.sort((a, b) => String(a.waldo).localeCompare(String(b.waldo)));
  return groups;
}

export interface FiberCluster {
  participantRows: number[];
  groups: FiberGroup[];
}

export function clusterFiberGroupsByParticipants(fiberGroups: FiberGroup[]): FiberCluster[] {
  const map = new Map<string, FiberCluster>();
  for (const g of fiberGroups) {
    const rows = Array.from(new Set(g.items.map(i => Number(i.key.split('|')[0]))))
      .sort((a, b) => a - b);
    const sig = rows.join(',');
    if (!map.has(sig)) map.set(sig, { participantRows: rows, groups: [] });
    map.get(sig)!.groups.push(g);
  }
  const clusters = Array.from(map.values());
  clusters.sort((a, b) => a.groups[0].strand - b.groups[0].strand);
  return clusters;
}

export function pickCurrentClusterWinner(cluster: FiberCluster, excluded: Set<string>): number {
  const counts = new Map<number, number>();
  for (const g of cluster.groups) {
    for (const c of g.items) {
      if (!excluded.has(c.key)) {
        const r = Number(c.key.split('|')[0]);
        counts.set(r, (counts.get(r) || 0) + 1);
      }
    }
  }
  let best = cluster.participantRows[0], bestCount = -1;
  for (const r of cluster.participantRows) {
    const c = counts.get(r) || 0;
    if (c > bestCount) { bestCount = c; best = r; }
  }
  return best;
}

export function findFiberDuplicateGroups(cands: ExfoCandidate[]): FiberGroup[] {
  const byStrand = new Map<number, ExfoCandidate[]>();
  for (const c of cands) {
    if (!byStrand.has(c.strand)) byStrand.set(c.strand, []);
    byStrand.get(c.strand)!.push(c);
  }
  const groups: FiberGroup[] = [];
  Array.from(byStrand.entries()).forEach(([strand, list]) => {
    if (list.length > 1) groups.push({ strand, items: list });
  });
  groups.sort((a, b) => a.strand - b.strand);
  return groups;
}

function makeRow(
  cfg: ResolvedJobCfg,
  terminal: string | null,
  cable: string | null,
  strand: number,
  tt01: string,
  tt02: string,
  fiberIndex: number,
  isFirst: boolean
): any[] {
  const testPoint = `${cable}_${strand} - ${terminal}_${fiberIndex}`;
  return [
    isFirst ? cfg.name : '',
    isFirst ? cfg.assignees : '',
    isFirst ? cfg.company : '',
    isFirst ? cfg.customer : '',
    isFirst ? cfg.dueDate : '',
    testPoint,
    cable ?? '',
    strand,
    cfg.aloc,
    cfg.zloc,
    cfg.wcc,
    tt01,
    tt02,
    isFirst ? cfg.testConfigs : '',
  ];
}

export function candidatesToRows(cfg: ResolvedJobCfg, cands: ExfoCandidate[]): any[][] {
  const rows: any[][] = [CSV_HEADER.slice()];
  let isFirst = true;
  for (const c of cands) {
    rows.push(makeRow(cfg, c.terminal, c.cable, c.strand, c.tt01, c.tt02, c.fiberIndex, isFirst));
    isFirst = false;
  }
  return rows;
}

export function parseAddress(name: string | null | undefined): string {
  if (!name) return '';
  const m = String(name).trim().match(/^[A-Z]+\s+(.+?)\s+[A-Z][A-Z0-9-]*$/);
  return m ? m[1] : String(name).trim();
}

export function terminalStrandNumbers(t: ExfoTerminal): number[] {
  const out: number[] = [];
  if (t.powerStrand != null && !Number.isNaN(t.powerStrand) && t.powerStrand > 0) {
    out.push(t.powerStrand);
  }
  if (Array.isArray(t.otdrStrands)) out.push(...t.otdrStrands);
  return out.sort((a, b) => a - b);
}

export function terminalStrandRange(t: ExfoTerminal): string {
  const nums = terminalStrandNumbers(t);
  if (!nums.length) return '';
  const segs: string[] = [];
  let lo = nums[0], hi = nums[0];
  for (let i = 1; i < nums.length; i++) {
    if (nums[i] === hi + 1) {
      hi = nums[i];
      continue;
    }
    segs.push(lo === hi ? `${lo}` : `${lo}-${hi}`);
    lo = hi = nums[i];
  }
  segs.push(lo === hi ? `${lo}` : `${lo}-${hi}`);
  return segs.join(',');
}

export function manhattanFeet(a: { lat: number; lng: number }, b: { lat: number; lng: number }): number {
  const R = 20902231;
  const toRad = (d: number) => d * Math.PI / 180;
  const nsFt = Math.abs(b.lat - a.lat) * (Math.PI / 180) * R;
  const midLat = toRad((a.lat + b.lat) / 2);
  const ewFt = Math.abs(b.lng - a.lng) * (Math.PI / 180) * R * Math.cos(midLat);
  return nsFt + ewFt;
}

export function formatDistance(feet: number | null | undefined): string {
  if (feet == null || Number.isNaN(feet)) return '';
  const m = feet * 0.3048;
  return `${Math.round(feet).toLocaleString()} ft (${Math.round(m).toLocaleString()} m)`;
}

export const IOLM_SHORT_LINK_FT = 8200 * 0.8;

// CLLI lookup: first 6 chars (4-char city + 2-char state) → "City, ST"
export const CLLI_BUILTIN: { [k: string]: string } = {
  // Wisconsin
  MILWWI: 'Milwaukee, WI', MSKGWI: 'Muskego, WI', CDBGWI: 'Cedarburg, WI',
  HBTSWI: 'Hubertus, WI', JCSNWI: 'Jackson, WI', MNFLWI: 'Menomonee Falls, WI',
  NWBGWI: 'Newburg, WI', PTWAWI: 'Port Washington, WI', WBNDWI: 'West Bend, WI',
  BGBNWI: 'Big Bend, WI', BRFDWI: 'Brookfield, WI', HRLDWI: 'Hartford, WI',
  HRTLWI: 'Hartland, WI', OCNMWI: 'Oconomowoc, WI', PEWKWI: 'Pewaukee, WI',
  SUSXWI: 'Sussex, WI', WTWNWI: 'Watertown, WI', WKSHWI: 'Waukesha, WI',
  CLDNWI: 'Caledonia, WI', KENOWI: 'Kenosha, WI', PLPRWI: 'Pleasant Prairie, WI',
  PRRSWI: 'Kenosha, WI', RACNWI: 'Racine, WI', SMRSWI: 'Somers, WI',
  UNGVWI: 'Union Grove, WI', BELTWI: 'Beloit, WI', BURLWI: 'Burlington, WI',
  DLVNWI: 'Delavan, WI', EVVLWI: 'Evansville, WI', FTATWI: 'Fort Atkinson, WI',
  GNCYWI: 'Genoa City, WI', JNVLWI: 'Janesville, WI',
  APPLWI: 'Appleton, WI', MADIWI: 'Madison, WI', GRBYWI: 'Green Bay, WI',
  WSAUWI: 'Wausau, WI', OSHKWI: 'Oshkosh, WI', NEENWI: 'Neenah, WI',
  EAUCWI: 'Eau Claire, WI', LACRWI: 'La Crosse, WI', FKLNWI: 'Franklin, WI',
  SUPRWI: 'Superior, WI', MNTWWI: 'Manitowoc, WI', STVPWI: 'Stevens Point, WI',
  SHBYWI: 'Sheboygan, WI', MRTNWI: 'Marinette, WI', RHLDWI: 'Rhinelander, WI',
  FDLCWI: 'Fond du Lac, WI', WRPDWI: 'Wisconsin Rapids, WI',
  // Illinois
  CHCGIL: 'Chicago, IL', AURRIL: 'Aurora, IL', NPVLIL: 'Naperville, IL',
  JOLTIL: 'Joliet, IL', RCKFIL: 'Rockford, IL', ELGNIL: 'Elgin, IL',
  PEORIL: 'Peoria, IL', SPFDIL: 'Springfield, IL', CHMGIL: 'Champaign, IL',
  EVSTIL: 'Evanston, IL', OKPKIL: 'Oak Park, IL', ORPKIL: 'Orland Park, IL',
  // Indiana
  INPLIN: 'Indianapolis, IN', FTWNIN: 'Fort Wayne, IN', EVVLIN: 'Evansville, IN',
  SBNDIN: 'South Bend, IN', GARYIN: 'Gary, IN', HMNDIN: 'Hammond, IN',
  BLGTIN: 'Bloomington, IN', MNCIIN: 'Muncie, IN', LFYTIN: 'Lafayette, IN',
  // Michigan
  DTRTMI: 'Detroit, MI', GRRPMI: 'Grand Rapids, MI', LNNGMI: 'Lansing, MI',
  ANARMI: 'Ann Arbor, MI', FLNTMI: 'Flint, MI', WARRMI: 'Warren, MI',
  LIVNMI: 'Livonia, MI', KLMZMI: 'Kalamazoo, MI', SAGWMI: 'Saginaw, MI',
  DRBNMI: 'Dearborn, MI', TROYMI: 'Troy, MI', TVRSMI: 'Traverse City, MI',
  // Ohio
  CLEVOH: 'Cleveland, OH', CLMBOH: 'Columbus, OH', CNCIOH: 'Cincinnati, OH',
  TOLDOH: 'Toledo, OH', AKRNOH: 'Akron, OH', DYTNOH: 'Dayton, OH',
  YNTWOH: 'Youngstown, OH', CNTNOH: 'Canton, OH', LORAOH: 'Lorain, OH',
  PARMOH: 'Parma, OH', HLTSOH: 'Hamilton, OH', SPNGOH: 'Springfield, OH',
  // Minnesota
  MNAPMN: 'Minneapolis, MN', STPLMN: 'Saint Paul, MN', RCHSMN: 'Rochester, MN',
  DULHMN: 'Duluth, MN', BLTNMN: 'Bloomington, MN', BRKPMN: 'Brooklyn Park, MN',
  PLMHMN: 'Plymouth, MN', STCLMN: 'Saint Cloud, MN',
  // Iowa
  DSMNIA: 'Des Moines, IA', CDRPIA: 'Cedar Rapids, IA', DVPTIA: 'Davenport, IA',
  SOCYIA: 'Sioux City, IA', WTLOIA: 'Waterloo, IA', IACYIA: 'Iowa City, IA',
  AMESIA: 'Ames, IA', CNBLIA: 'Council Bluffs, IA',
  // Missouri / Kansas / Nebraska / Dakotas / KY / TN
  STLSMO: 'St. Louis, MO', KSCYMO: 'Kansas City, MO', SPFDMO: 'Springfield, MO',
  COLBMO: 'Columbia, MO', INDPMO: 'Independence, MO',
  WCHTKS: 'Wichita, KS', OVPKKS: 'Overland Park, KS', KSCYKS: 'Kansas City, KS',
  TOPKKS: 'Topeka, KS', OLATKS: 'Olathe, KS', LWRCKS: 'Lawrence, KS',
  OMHANE: 'Omaha, NE', LNCNNE: 'Lincoln, NE', BLVENE: 'Bellevue, NE', GRISNE: 'Grand Island, NE',
  FARGND: 'Fargo, ND', BISMND: 'Bismarck, ND', GRFKND: 'Grand Forks, ND',
  SXFSSD: 'Sioux Falls, SD', RPCYSD: 'Rapid City, SD',
  LSVLKY: 'Louisville, KY', LEXNKY: 'Lexington, KY', BWGRKY: 'Bowling Green, KY',
  NSVLTN: 'Nashville, TN', MMPHTN: 'Memphis, TN', KNXVTN: 'Knoxville, TN',
  CHTTTN: 'Chattanooga, TN', CLRKTN: 'Clarksville, TN',
  // NY / NJ / PA / MA
  NYCMNY: 'New York, NY', NYCKNY: 'New York, NY', NYCXNY: 'New York, NY',
  BFLONY: 'Buffalo, NY', RCSTNY: 'Rochester, NY', SYRCNY: 'Syracuse, NY',
  ALBYNY: 'Albany, NY', YNKRNY: 'Yonkers, NY', HMPSNY: 'Hempstead, NY', WCHSNY: 'White Plains, NY',
  NWRKNJ: 'Newark, NJ', JRCYNJ: 'Jersey City, NJ', PTRSNJ: 'Paterson, NJ',
  ELZBNJ: 'Elizabeth, NJ', TRNTNJ: 'Trenton, NJ', CAMDNJ: 'Camden, NJ',
  PHLAPA: 'Philadelphia, PA', PTBGPA: 'Pittsburgh, PA', ALTNPA: 'Allentown, PA',
  ERIEPA: 'Erie, PA', READPA: 'Reading, PA', SCRNPA: 'Scranton, PA',
  HRBGPA: 'Harrisburg, PA', LANCPA: 'Lancaster, PA',
  BSTNMA: 'Boston, MA', WRCSMA: 'Worcester, MA',
  CMBRMA: 'Cambridge, MA', LWLLMA: 'Lowell, MA', BRCTMA: 'Brockton, MA',
};

export function resolveCityFromCLLI(clli: string): { code: string; city: string } | null {
  if (!clli) return null;
  const code = clli.trim().toUpperCase().slice(0, 6);
  if (code.length !== 6) return null;
  let overrides: { [k: string]: string } = {};
  try {
    const raw = localStorage.getItem('f2job.clliOverrides');
    if (raw) overrides = JSON.parse(raw);
  } catch (_) {}
  const city = overrides[code] || CLLI_BUILTIN[code];
  return city ? { code, city } : null;
}
