import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { parseTerminals, type Terminal } from './excel';
import { buildCandidates, findWaldoDuplicateGroups, findFiberDuplicateGroups, clusterFiberGroupsByParticipants, parseStrandList, type ExfoTerminal } from './exfo';
import { cleanCellText, normalizeWorkbookCells } from './xlsx-export';

export const LEGACY_PON_HEADERS = ['#', 'Terminal', 'Waldo ID', 'PON Count', 'Total Strands', 'Test Port', 'Test Strand', 'Estm FT', 'Real FT', 'Failed Strand', 'Lost', '@FT', 'Task', 'Status'];
export const PON_HEADERS = ['#', 'Terminal', 'Task', 'Waldo ID', 'PON Count', 'Total Strands', 'Test Port', 'Test Strand', 'Estm FT', 'Real FT', 'Failed Strand', 'Lost', '@FT', 'Status'];

function reorderDataColumns(sheet: ExcelJS.Worksheet, columns: number[]) {
  sheet.eachRow((row, r) => {
    if (r < 5) return;
    const cells = columns.map(c => ({ value: row.getCell(c).value, style: { ...row.getCell(c).style } }));
    cells.forEach((cell, i) => { row.getCell(i + 1).value = cell.value; row.getCell(i + 1).style = cell.style; });
  });
}

const TASK_HEADERS = ['Print.Task', 'FRC', 'WAC', 'Terminal Type', 'Terminal Desc', 'Terminal Count', 'Terminal Address'];

export interface PonPowerResult {
  workbook: ExcelJS.Workbook;
  sheetName: string;
  terminals: Terminal[];
  matchedTasks: number;
  project: string;
  cableId: string;
  duplicateGroups: PonDuplicateGroup[];
}

export interface PonDuplicateGroup { key: string; kind: 'waldo' | 'fiber'; cableId: string; counts: number[]; items: Terminal[] }

function asExfoTerminals(terminals: Terminal[]): ExfoTerminal[] {
  return terminals.map(t => ({ row: t.rowIndex, terminal: t.terminalName, waldo: t.waldoId, cable: t.cableId,
    powerStrand: t.powerTestStrand, total: t.totalStrands, otdrRaw: t.otdrTestStrand,
    otdrStrands: parseStrandList(t.otdrTestStrand) }));
}

/** Use F2 Exfo's Waldo-first, then exact-participant Fiber ID conflict workflow. */
export function findPonDuplicateGroups(terminals: Terminal[], choices = new Map<string, number>()): PonDuplicateGroup[] {
  const exfo = asExfoTerminals(terminals);
  const excluded = new Set<number>();
  const waldoGroups = findWaldoDuplicateGroups(exfo).map(group => {
    const key = `waldo:${group.waldo}`;
    const winner = group.items.find(t => t.row === choices.get(key)) ?? group.items[0];
    group.items.forEach(t => { if (t.row !== winner.row) excluded.add(t.row); });
    return { key, kind: 'waldo' as const, cableId: group.waldo, counts: [], items: terminals.filter(t => group.items.some(e => e.row === t.rowIndex)) };
  });
  const clusters = clusterFiberGroupsByParticipants(findFiberDuplicateGroups(buildCandidates(exfo, 'iOLM', excluded)));
  return [...waldoGroups, ...clusters.map(cluster => ({ key: `fiber:${cluster.participantRows.join(',')}`, kind: 'fiber' as const,
    cableId: '', counts: cluster.groups.map(g => g.strand), items: terminals.filter(t => cluster.participantRows.includes(t.rowIndex)) }))];
}

/** Advance through ports 1–4, skipping ports unavailable on this terminal. */
export function staggerPonTerminals(terminals: Terminal[]): Terminal[] {
  let nextPort = 1;
  return terminals.map(t => {
    const originalPorts = t.totalStrands === 2 ? [2, 3] : Array.from({ length: Math.min(t.totalStrands, 4) }, (_, i) => i + 1);
    const ports = originalPorts.filter(p => !t.retainedStrands || t.retainedStrands.includes(t.powerTestStrand + p - originalPorts[0]));
    const port = ports.find(p => p >= nextPort) ?? ports[0];
    nextPort = port % 4 + 1;
    return { ...t, staggeredPort: port, staggeredStrand: t.powerTestStrand + port - originalPorts[0] };
  });
}

export async function preparePonPower(ponFile: File, dataFile: File | null = null, orcaText = '', duplicateChoices: Map<string, number> = new Map()): Promise<PonPowerResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await ponFile.arrayBuffer());
  const source = workbook.worksheets.find(s => s.getCell('B5').text === 'Terminal') ?? workbook.worksheets[0];
  if (!source) throw new Error('The Ponsheet workbook has no worksheets.');
  normalizeWorkbookCells(workbook);
  if (PON_HEADERS.every((h, i) => source.getCell(5, i + 1).text === h)) {
    reorderDataColumns(source, [1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 3, 14]);
  }
  const compact = LEGACY_PON_HEADERS.slice(0, 7).every((h, i) => source.getCell(5, i + 1).text === h);
  let terminals: Terminal[];
  let project: string;
  let cableId: string;
  let sheet: ExcelJS.Worksheet;
  if (compact) {
    sheet = source;
    project = sheet.getCell('D2').text.replace(/^PROJECT:\s*/i, '');
    cableId = sheet.getCell('H2').text.replace(/^CABLE ID:\s*/i, '');
    terminals = [];
    sheet.eachRow((row, r) => {
      if (r <= 5 || !row.getCell(2).text) return;
      const power = Number(row.getCell(4).text.split('-')[0]);
      const total = Number(row.getCell(5).value);
      if (!Number.isInteger(power) || power <= 0 || !Number.isInteger(total) || total <= 0) throw new Error(`Invalid PON count or total strands at row ${r}.`);
      terminals.push({ rowIndex: r - 1, terminalName: row.getCell(2).text, waldoId: row.getCell(3).text, cableId, powerTestStrand: power, totalStrands: total, otdrTestStrand: Array.from({ length: total - 1 }, (_, i) => power + i + 1).join(','), testpQty: '', testpaQty: '' });
    });
  } else {
    const parsed = await parseTerminals(ponFile);
    terminals = parsed.terminals;
    project = source.getCell('X2').text;
    cableId = Array.from(new Set(terminals.map(t => t.cableId))).join(', ');
    let pfp = '';
    source.eachRow(row => row.eachCell(cell => {
      if (!pfp && cell.value != null && /\sPFP$/i.test(cell.text)) pfp = cell.text;
    }));
    workbook.removeWorksheet(source.id);
    sheet = workbook.addWorksheet('PON TEST SHEET');
    sheet.mergeCells('A2:C2'); sheet.mergeCells('D2:G2'); sheet.mergeCells('H2:K2');
    sheet.mergeCells('A3:F3'); sheet.mergeCells('G3:M3');
    sheet.getCell('A2').value = `PFP:  ${pfp}`;
    sheet.getCell('D2').value = `PROJECT:  ${project}`;
    sheet.getCell('H2').value = `CABLE ID:  ${cableId}`;
    sheet.getCell('A3').value = `TOTAL STRANDS:  ${terminals.reduce((n, t) => n + t.totalStrands, 0)}`;
    sheet.getCell('G3').value = `TERMINALS:  ${terminals.length}`;
    sheet.getRow(5).values = LEGACY_PON_HEADERS;
    const widths = [5, 30, 12, 12, 14, 11, 12, 10, 10, 13, 10, 10, 14, 10];
    widths.forEach((width, i) => { sheet.getColumn(i + 1).width = width; });
    terminals = terminals.map((t, i) => {
      const r = i + 6;
      sheet.getRow(r).values = [i + 1, t.terminalName, t.waldoId, `${t.powerTestStrand}-${t.powerTestStrand + t.totalStrands - 1}`, t.totalStrands];
      return { ...t, rowIndex: r - 1 };
    });
    sheet.eachRow(row => {
      row.height = 30;
      row.eachCell({ includeEmpty: true }, cell => {
        cell.font = { name: 'Calibri', size: 11, bold: row.number <= 5 };
        cell.alignment = { vertical: 'middle', wrapText: true };
      });
    });
    sheet.getRow(5).eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } }; });
    sheet.views = [{ state: 'frozen', ySplit: 5 }];
    sheet.autoFilter = `A5:N${terminals.length + 5}`;
    sheet.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '1:5', printArea: `A1:N${terminals.length + 5}` };
  }
  if (!terminals.length) throw new Error('No terminals found in the Ponsheet.');
  terminals = terminals.map(t => ({ ...t, terminalName: cleanCellText(t.terminalName), waldoId: cleanCellText(t.waldoId), cableId: cleanCellText(t.cableId) }));
  const duplicateGroups = findPonDuplicateGroups(terminals, duplicateChoices);
  const excludedRows = new Set<number>();
  for (const group of duplicateGroups.filter(g => g.kind === 'waldo')) {
    const winner = group.items.find(t => t.rowIndex === duplicateChoices.get(group.key)) ?? group.items[0];
    group.items.forEach(t => { if (t !== winner) excludedRows.add(t.rowIndex); });
  }
  const candidates = buildCandidates(asExfoTerminals(terminals), 'iOLM', excludedRows);
  const excludedCandidates = new Set<string>();
  for (const cluster of clusterFiberGroupsByParticipants(findFiberDuplicateGroups(candidates))) {
    const winner = duplicateChoices.get(`fiber:${cluster.participantRows.join(',')}`);
    for (const group of cluster.groups) {
      const kept = group.items.find(c => c.terminalRow === winner) ?? group.items[0];
      group.items.forEach(c => { if (c.key !== kept.key) excludedCandidates.add(c.key); });
    }
  }
  terminals = terminals.map(t => ({ ...t, retainedStrands: candidates.filter(c => c.terminalRow === t.rowIndex && !excludedCandidates.has(c.key)).map(c => c.strand) }));
  terminals.forEach(t => { if (!t.retainedStrands?.length) excludedRows.add(t.rowIndex); });
  if (excludedRows.size) {
    const oldLast = sheet.rowCount;
    const kept = terminals.filter(t => !excludedRows.has(t.rowIndex));
    const values = kept.map(t => sheet.getRow(t.rowIndex + 1).values);
    for (let r = 6; r <= oldLast; r++) sheet.getRow(r).values = [];
    terminals = kept.map((t, i) => {
      const r = i + 6;
      sheet.getRow(r).values = values[i];
      sheet.getCell(r, 1).value = i + 1;
      return { ...t, rowIndex: r - 1 };
    });
    sheet.spliceRows(terminals.length + 6, oldLast - terminals.length - 5);
    sheet.getCell('A3').value = `TOTAL STRANDS: ${terminals.reduce((n, t) => n + t.totalStrands, 0)}`;
    sheet.getCell('G3').value = `TERMINALS: ${terminals.length}`;
    sheet.pageSetup.printArea = `A1:N${terminals.length + 5}`;
  }

  if (dataFile) {
    const data = XLSX.read(await dataFile.arrayBuffer(), { type: 'array' });
    const rows = data.SheetNames.map(name => XLSX.utils.sheet_to_json<string[]>(data.Sheets[name], { header: 1, defval: '' }))
      .find(rows => TASK_HEADERS.every(h => rows[0]?.some(value => String(value).trim() === h)));
    if (!rows) throw new Error('Data file must contain Print.Task, FRC, WAC, Terminal Type, Terminal Desc, Terminal Count, and Terminal Address headers.');
    const columns = TASK_HEADERS.map(h => rows[0].findIndex(value => String(value).trim() === h));
    const dataRows: ExcelJS.CellValue[][] = [];
    for (const row of rows.slice(1)) {
      const values = columns.map(c => row[c] ?? '');
      dataRows.push(values);
    }
    const oldTask = workbook.getWorksheet('Task');
    if (oldTask) workbook.removeWorksheet(oldTask.id);
    const taskSheet = workbook.addWorksheet('Task');
    taskSheet.addRow(TASK_HEADERS);
    dataRows.forEach(row => taskSheet.addRow(row));
    [14, 10, 10, 16, 24, 65, 38].forEach((width, i) => { taskSheet.getColumn(i + 1).width = width; });
    taskSheet.views = [{ state: 'frozen', ySplit: 1 }];
    taskSheet.autoFilter = `A1:G${dataRows.length + 1}`;
  }
  // Keep an existing Task sheet when no data is uploaded, or create a blank one
  // so exported lookups can be used after task data is added in Excel.
  const taskSheet = workbook.getWorksheet('Task') ?? workbook.addWorksheet('Task');
  if (!taskSheet.getCell('A1').value) taskSheet.addRow(TASK_HEADERS);
  const taskRows: { count: string; task: string }[] = [];
  taskSheet.eachRow((row, r) => {
    if (r > 1) taskRows.push({ count: row.getCell(6).text.toUpperCase(), task: row.getCell(1).text });
  });
  const orcaHeaders = ['Task', 'Status', 'Open Flag', 'Col D', 'Terminal Desc', 'WAC', 'FRC', 'Col H'];
  const existingOrca = workbook.getWorksheet('Orca');
  let orcaRows: ExcelJS.CellValue[][] = [];
  if (orcaText.trim()) {
    const pasted = XLSX.read(orcaText.trim(), { type: 'string', raw: true, FS: orcaText.includes('\t') ? '\t' : ',' });
    const rows = XLSX.utils.sheet_to_json<string[]>(pasted.Sheets[pasted.SheetNames[0]], { header: 1, defval: '' });
    const hasHeader = rows[0]?.[0]?.trim().toLowerCase() === 'task';
    const values = hasHeader ? rows.slice(1) : rows;
    const invalid = values.find(row => row.some(v => String(v).trim()) && (!Number.isFinite(Number(row[0])) || !String(row[0]).trim()));
    if (invalid || !values.length) throw new Error('Paste Orca rows with Task in the first column and Status in the second column.');
    orcaRows = values.filter(row => String(row[0]).trim()).map(row => [Number(row[0]), ...row.slice(1, 8)]);
  } else if (existingOrca) {
    // Migrate old A/B workbooks and preserve already shifted B/C workbooks.
    const taskColumn = existingOrca.getCell('B1').text.trim().toLowerCase() === 'task' ? 2 : 1;
    existingOrca.eachRow((row, r) => {
      if (r <= 1 || !row.getCell(taskColumn).text.trim()) return;
      const values = orcaHeaders.map((_, i) => row.getCell(taskColumn + i).value);
      const task = row.getCell(taskColumn).text.trim();
      if (Number.isFinite(Number(task))) values[0] = Number(task);
      orcaRows.push(values);
    });
  }
  if (existingOrca) workbook.removeWorksheet(existingOrca.id);
  const orca = workbook.addWorksheet('Orca');
  orca.addRow(['', ...orcaHeaders]);
  orcaRows.forEach(row => orca.addRow(['', ...row]));
  orca.getCell('A2').value = 'Paste Here';
  // Cover future pasted rows as well as the currently populated Orca data.
  orca.addConditionalFormatting({
    ref: 'A2:I1048576',
    rules: [{
      type: 'expression', priority: 1, formulae: ['MOD(ROW(),2)=1'],
      style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' }, bgColor: { argb: 'FFF2F2F2' } } },
    }],
  });
  const statuses = new Map<string, string>();
  orca.eachRow((row, r) => {
    const task = row.getCell(2).text.trim();
    if (r > 1 && task && !statuses.has(String(Number(task)))) statuses.set(String(Number(task)), row.getCell(3).text);
  });
  terminals = staggerPonTerminals(terminals);
  let matchedTasks = 0;
  for (const t of terminals) {
    const r = t.rowIndex + 1;
    sheet.getCell(r, 6).value = t.staggeredPort!;
    sheet.getCell(r, 7).value = t.staggeredStrand!;
    const task = taskRows.find(row => row.count.includes(`${t.cableId.trim()},${t.powerTestStrand}-`.toUpperCase()))?.task;
    if (task != null) matchedTasks++;
    const escapedCable = t.cableId.trim().replace(/[~*?]/g, '~$&').replace(/"/g, '""');
    sheet.getCell(r, 13).value = { formula: `IFERROR(INDEX('Task'!$A:$A,MATCH("*${escapedCable},"&TRIM(LEFT(E${r},FIND("-",E${r})-1))&"-*",'Task'!$F:$F,0)),"")`, result: task ?? '' };
    sheet.getCell(r, 14).value = { formula: `IF(C${r}="","",IFERROR(INDEX('Orca'!$C:$C,MATCH(--C${r},'Orca'!$B:$B,0)),""))`, result: task == null ? '' : statuses.get(String(Number(task))) ?? '' };
    // Status conditional fills override this light-gray alternating background.
    for (let c = 1; c <= 14; c++) {
      const cell = sheet.getCell(r, c);
      // Imported cells can share a style object; detach before assigning row fills.
      cell.style = { ...cell.style };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: r % 2 === 0 ? 'FFFFFFFF' : 'FFF2F2F2' } };
    }
  }
  const lastRow = Math.max(...terminals.map(t => t.rowIndex + 1));
  sheet.autoFilter = `A5:N${lastRow}`;
  // Replace status rules from previously converted inputs; retain other rules.
  sheet.removeConditionalFormatting((formatting: ExcelJS.ConditionalFormattingOptions) => {
    formatting.rules = formatting.rules.filter(rule =>
      !('formulae' in rule && rule.formulae?.some(formula => /\$N\$?\d+\s*=\s*"[OC]"/i.test(String(formula)))));
    return formatting.rules.length > 0;
  });
  sheet.addConditionalFormatting({
    ref: `A6:N${lastRow}`,
    // Excel's differential fills need the background color as well as the
    // foreground color; foreground-only rules can load without visible fill.
    rules: [
      { type: 'expression', priority: 1, formulae: ['$N6="O"'], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' }, bgColor: { argb: 'FFFCE4D6' } } } },
      { type: 'expression', priority: 2, formulae: ['$N6="C"'], style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' }, bgColor: { argb: 'FFC6E0B4' } } } },
    ],
  });
  reorderDataColumns(sheet, [1, 2, 13, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14]);
  workbook.calcProperties.fullCalcOnLoad = true;
  normalizeWorkbookCells(workbook);
  return { workbook, sheetName: sheet.name, terminals, matchedTasks, project, cableId, duplicateGroups };
}

export function staggeredFilename(name: string): string {
  return `${name.replace(/\.xlsx$/i, '').replace(/_staggered$/i, '')}_staggered.xlsx`;
}

/** Apply only reviewed Google distances; clear stale imported or previously exported outliers. */
export function applyPonDistances(result: PonPowerResult, distances: Map<number, number>, suppressedRows: Set<number>) {
  const sheet = result.workbook.getWorksheet(result.sheetName)!;
  for (const terminal of result.terminals) {
    const row = terminal.rowIndex + 1;
    if (suppressedRows.has(row)) sheet.getCell(row, 9).value = null;
    else if (distances.has(row)) sheet.getCell(row, 9).value = Math.round(distances.get(row)!);
  }
}
