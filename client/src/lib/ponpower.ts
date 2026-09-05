import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { parseTerminals, type Terminal } from './excel';

export const PON_HEADERS = ['#', 'Terminal', 'Waldo ID', 'PON Count', 'Total Strands', 'Test Port', 'Test Strand', 'Estm FT', 'Real FT', 'Failed Strand', 'Lost', '@FT', 'Task', 'Status'];
const TASK_HEADERS = ['Print.Task', 'FRC', 'WAC', 'Terminal Type', 'Terminal Desc', 'Terminal Count', 'Terminal Address'];
const addressKey = (value: string) => value.replace(/\s+CFST\s*$/i, '').trim().replace(/\s+/g, ' ').toUpperCase();

export interface PonPowerResult {
  workbook: ExcelJS.Workbook;
  sheetName: string;
  terminals: Terminal[];
  matchedTasks: number;
  project: string;
  cableId: string;
}

/** Advance through ports 1–4, skipping ports unavailable on this terminal. */
export function staggerPonTerminals(terminals: Terminal[]): Terminal[] {
  let nextPort = 1;
  return terminals.map(t => {
    const ports = t.totalStrands === 2 ? [2, 3] : Array.from({ length: Math.min(t.totalStrands, 4) }, (_, i) => i + 1);
    const port = ports.find(p => p >= nextPort) ?? ports[0];
    nextPort = port % 4 + 1;
    return { ...t, staggeredPort: port, staggeredStrand: t.powerTestStrand + port - ports[0] };
  });
}

export async function preparePonPower(ponFile: File, dataFile: File, orcaText = ''): Promise<PonPowerResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await ponFile.arrayBuffer());
  const source = workbook.worksheets.find(s => s.getCell('B5').text === 'Terminal') ?? workbook.worksheets[0];
  if (!source) throw new Error('The Ponsheet workbook has no worksheets.');
  const compact = PON_HEADERS.slice(0, 7).every((h, i) => source.getCell(5, i + 1).text === h);
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
      terminals.push({ rowIndex: r - 1, terminalName: row.getCell(2).text, waldoId: row.getCell(3).text, cableId, powerTestStrand: power, totalStrands: total, otdrTestStrand: '', testpQty: '', testpaQty: '' });
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
    sheet.getRow(5).values = PON_HEADERS;
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

  const data = XLSX.read(await dataFile.arrayBuffer(), { type: 'array' });
  const rows = data.SheetNames.map(name => XLSX.utils.sheet_to_json<string[]>(data.Sheets[name], { header: 1, defval: '' }))
    .find(rows => TASK_HEADERS.every(h => rows[0]?.some(value => String(value).trim() === h)));
  if (!rows) throw new Error('Data file must contain Print.Task, FRC, WAC, Terminal Type, Terminal Desc, Terminal Count, and Terminal Address headers.');
  const columns = TASK_HEADERS.map(h => rows[0].findIndex(value => String(value).trim() === h));
  const dataRows: ExcelJS.CellValue[][] = [];
  const tasks = new Map<string, ExcelJS.CellValue>();
  for (const row of rows.slice(1)) {
    const values = columns.map(c => row[c] ?? '');
    dataRows.push(values);
    const key = addressKey(String(values[6]));
    if (!key) continue;
    if (tasks.has(key) && tasks.get(key) !== values[0]) throw new Error(`Multiple tasks found for ${values[6]}.`);
    tasks.set(key, values[0]);
  }
  const oldTask = workbook.getWorksheet('Task');
  if (oldTask) workbook.removeWorksheet(oldTask.id);
  const taskSheet = workbook.addWorksheet('Task');
  taskSheet.addRow(TASK_HEADERS);
  dataRows.forEach(row => taskSheet.addRow(row));
  [14, 10, 10, 16, 24, 65, 38].forEach((width, i) => { taskSheet.getColumn(i + 1).width = width; });
  taskSheet.views = [{ state: 'frozen', ySplit: 1 }];
  taskSheet.autoFilter = `A1:G${dataRows.length + 1}`;
  if (orcaText.trim()) {
    const pasted = XLSX.read(orcaText.trim(), { type: 'string', raw: true, FS: orcaText.includes('\t') ? '\t' : ',' });
    const rows = XLSX.utils.sheet_to_json<string[]>(pasted.Sheets[pasted.SheetNames[0]], { header: 1, defval: '' });
    const hasHeader = rows[0]?.[0]?.trim().toLowerCase() === 'task';
    const values = hasHeader ? rows.slice(1) : rows;
    const invalid = values.find(row => row.some(v => String(v).trim()) && (!Number.isFinite(Number(row[0])) || !String(row[0]).trim()));
    if (invalid || !values.length) throw new Error('Paste Orca rows with Task in the first column and Status in the second column.');
    const existing = workbook.getWorksheet('Orca');
    if (existing) workbook.removeWorksheet(existing.id);
    const pastedOrca = workbook.addWorksheet('Orca');
    pastedOrca.addRow(hasHeader ? rows[0] : ['Task', 'Status', 'Open Flag', 'Col D', 'Terminal Desc', 'WAC', 'FRC', 'Col H']);
    values.filter(row => String(row[0]).trim()).forEach(row => pastedOrca.addRow([Number(row[0]), ...row.slice(1)]));
  }
  const orca = workbook.getWorksheet('Orca') ?? workbook.addWorksheet('Orca');
  if (!orca.getCell('A1').value) orca.addRow(['Task', 'Status', 'Open Flag', 'Col D', 'Terminal Desc', 'WAC', 'FRC', 'Col H']);
  const statuses = new Map<string, string>();
  orca.eachRow((row, r) => { if (r > 1) statuses.set(String(Number(row.getCell(1).text)), row.getCell(2).text); });
  terminals = staggerPonTerminals(terminals);
  let matchedTasks = 0;
  for (const t of terminals) {
    const r = t.rowIndex + 1;
    sheet.getCell(r, 6).value = t.staggeredPort!;
    sheet.getCell(r, 7).value = t.staggeredStrand!;
    const task = tasks.get(addressKey(t.terminalName));
    if (task != null) matchedTasks++;
    sheet.getCell(r, 13).value = { formula: `IFERROR(INDEX('Task'!$A:$A,MATCH(TRIM(SUBSTITUTE(B${r}," CFST","")),'Task'!$G:$G,0)),"")`, result: task == null ? '' : String(task) };
    sheet.getCell(r, 14).value = { formula: `IF(M${r}="","",IFERROR(INDEX('Orca'!$B:$B,MATCH(--M${r},'Orca'!$A:$A,0)),""))`, result: task == null ? '' : statuses.get(String(Number(task))) ?? '' };
    // Clear imported static status colors so blank statuses remain unfilled.
    for (let c = 1; c <= 14; c++) sheet.getCell(r, c).fill = { type: 'pattern', pattern: 'none' };
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
  workbook.calcProperties.fullCalcOnLoad = true;
  return { workbook, sheetName: sheet.name, terminals, matchedTasks, project, cableId };
}

export function staggeredFilename(name: string): string {
  return `${name.replace(/\.xlsx$/i, '').replace(/_staggered$/i, '')}_staggered.xlsx`;
}
