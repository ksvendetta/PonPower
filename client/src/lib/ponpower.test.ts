import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import { unzipSync, strFromU8 } from 'fflate';
import { PON_HEADERS, preparePonPower, staggeredFilename, findPonDuplicateGroups, applyPonDistances } from './ponpower';
import { writeExcelWorkbook } from './xlsx-export';
import { parseExfoXlsx } from './exfo';

async function file(path: string) {
  return new File([await readFile(path)], path.split('/').at(-1)!);
}

test('preview/export work without data, preserving embedded tasks when available', async () => {
  for (const path of ['examples/PON_TEST_SHEET__17_.xlsx', 'examples/PON_TEST_SHEET__17__stag_Text.xlsx', 'PONSHEET.xlsx']) {
    const result = await preparePonPower(await file(path));
    assert.ok(result.terminals.length);
    const saved = new ExcelJS.Workbook();
    await saved.xlsx.load(await result.workbook.xlsx.writeBuffer());
    const sheet = saved.getWorksheet(result.sheetName)!;
    assert.ok(sheet.getCell('F6').value);
    assert.ok(sheet.getCell('M6').formula.includes("'Task'!$F:$F"));
    assert.ok(saved.getWorksheet('Task'));
    if (path.includes('stag_Text')) assert.equal(result.matchedTasks, 188);
    else {
      assert.equal(result.matchedTasks, 0);
      assert.equal(sheet.getCell('M6').text, '');
      assert.equal(sheet.getCell('N6').text, '');
    }
  }
});

test('task lookup matches cable and starting count only, including split counts and first match', async () => {
  const pon = new ExcelJS.Workbook();
  const sheet = pon.addWorksheet('PON TEST SHEET');
  sheet.getCell('D2').value = 'PROJECT: example';
  sheet.getCell('H2').value = 'CABLE ID: PON4250WRR';
  sheet.getRow(5).values = PON_HEADERS;
  sheet.getRow(6).values = [1, 'unrelated address CFST', '123', '45-46', 2];
  const data = new ExcelJS.Workbook();
  const tasks = data.addWorksheet('Export');
  tasks.addRow(['Print.Task', 'FRC', 'WAC', 'Terminal Type', 'Terminal Desc', 'Terminal Count', 'Terminal Address']);
  tasks.addRow(['wrong cable', '', '', '', '', '["OTHER,45-45"]', 'unrelated address']);
  tasks.addRow(['wrong count', '', '', '', '', '["PON4250WRR,145-145"]', 'unrelated address']);
  tasks.addRow(['1.493', '', '', '', '', '["X,1-1","PON4250WRR,45-45","PON4250WRR,46-46","X,4-4"]', 'different address']);
  tasks.addRow(['duplicate', '', '', '', '', '["PON4250WRR,45-46"]', 'different address']);
  const result = await preparePonPower(
    new File([await pon.xlsx.writeBuffer()], 'pon.xlsx'),
    new File([await data.xlsx.writeBuffer()], 'data.xlsx'), 'Task\tStatus\n1.493\tC');
  const saved = new ExcelJS.Workbook();
  await saved.xlsx.load(await result.workbook.xlsx.writeBuffer());
  assert.equal(result.matchedTasks, 1);
  assert.equal(saved.getWorksheet(result.sheetName)!.getCell('M6').result, '1.493');
  assert.equal(saved.getWorksheet(result.sheetName)!.getCell('N6').result, 'C');
  assert.equal(saved.getWorksheet(result.sheetName)!.getCell('M6').formula,
    `IFERROR(INDEX('Task'!$A:$A,MATCH("*PON4250WRR,"&TRIM(LEFT(D6,FIND("-",D6)-1))&"-*",'Task'!$F:$F,0)),"")`);
});

test('all 188 staggered ports and strands match the reference; task rows survive export', async () => {
  const pon = await file('examples/PON_TEST_SHEET__17__stag_Text.xlsx');
  const original = new ExcelJS.Workbook();
  await original.xlsx.load(await pon.arrayBuffer());
  const result = await preparePonPower(pon, await file('examples/data.xlsx'));
  assert.equal(result.terminals.length, 188);
  assert.equal(result.matchedTasks, 188);
  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load(await result.workbook.xlsx.writeBuffer());
  const actual = reloaded.getWorksheet('PON TEST SHEET')!;
  const expected = original.getWorksheet('PON TEST SHEET')!;
  for (let r = 6; r <= 193; r++) {
    for (let c = 1; c <= 12; c++) assert.deepEqual(actual.getCell(r, c).value, expected.getCell(r, c).value, `cell ${r},${c}`);
    assert.ok(actual.getCell(r, 13).result, `task at row ${r}`);
  }
  assert.equal(reloaded.getWorksheet('Task')!.getCell('A2').value, '1.499');
  assert.equal(reloaded.getWorksheet('Task')!.getCell('G2').value, 'F 4303 W WOODWARD DR');
  assert.equal(reloaded.getWorksheet('Orca')!.rowCount, 189);
  assert.deepEqual(actual.model.merges, expected.model.merges);
  assert.equal(reloaded.model.media.length, original.model.media.length);
});

test('legacy Ponsheet converts to compact layout and unmatched project tasks stay blank', async () => {
  const result = await preparePonPower(await file('PONSHEET.xlsx'), await file('examples/data.xlsx'));
  assert.equal(result.terminals.length, 60);
  assert.equal(result.matchedTasks, 0);
  const s = result.workbook.getWorksheet('PON TEST SHEET')!;
  assert.equal(s.getCell('B6').text, 'S 920 E POTTER AVE CFST');
  assert.equal(s.getCell('F6').value, 1);
  assert.equal(s.getCell('G6').value, 5);
  assert.equal(s.getCell('M6').result, '');
  assert.equal(s.getCell('H6').value, null);
});

test('pasted Orca rows populate status lookups and reject malformed input', async () => {
  const pon = await file('examples/PON_TEST_SHEET__17__stag_Text.xlsx');
  const data = await file('examples/data.xlsx');
  const result = await preparePonPower(pon, data, 'Task\tStatus\n1.499\tO');
  assert.equal(result.workbook.getWorksheet('PON TEST SHEET')!.getCell('N17').result, 'O');
  assert.equal(result.workbook.getWorksheet('Orca')!.rowCount, 2);
  const withoutHeader = await preparePonPower(pon, data, '1.499\tC');
  assert.equal(withoutHeader.workbook.getWorksheet('PON TEST SHEET')!.getCell('N17').result, 'C');
  await assert.rejects(preparePonPower(pon, data, 'wrong\tO'), /Paste Orca rows/);
  await assert.rejects(preparePonPower(pon, await file('PONSHEET.xlsx')), /Data file must contain/);
});

test('download name gets one staggered suffix', () => {
  assert.equal(staggeredFilename('PON_TEST_SHEET__17_.xlsx'), 'PON_TEST_SHEET__17__staggered.xlsx');
  assert.equal(staggeredFilename('PON_TEST_SHEET__17__staggered.xlsx'), 'PON_TEST_SHEET__17__staggered.xlsx');
});

test('reconstructed source and data reproduce every reference terminal', async () => {
  const source = await file('examples/PON_TEST_SHEET__17_.xlsx');
  const sourceBook = new ExcelJS.Workbook();
  await sourceBook.xlsx.load(await source.arrayBuffer());
  assert.equal(sourceBook.worksheets.length, 1);
  assert.equal(sourceBook.worksheets[0].getCell('G7').value, 31);
  assert.equal(sourceBook.worksheets[0].getCell('M6').value, null);
  const result = await preparePonPower(source, await file('examples/data.xlsx'));
  const reference = new ExcelJS.Workbook();
  await reference.xlsx.load(await (await file('examples/PON_TEST_SHEET__17__stag_Text.xlsx')).arrayBuffer());
  assert.equal(result.matchedTasks, 188);
  for (let r = 6; r <= 193; r++) {
    for (let c = 1; c <= 12; c++) {
      assert.deepEqual(result.workbook.getWorksheet(result.sheetName)!.getCell(r, c).value,
        reference.worksheets[0].getCell(r, c).value, `row ${r}, column ${c}`);
    }
  }
});

test('export enables every column filter and dynamic O/C row fills for both Ponsheet layouts', async () => {
  for (const path of ['examples/PON_TEST_SHEET__17__stag_Text.xlsx', 'PONSHEET.xlsx']) {
    const result = await preparePonPower(await file(path), await file('examples/data.xlsx'), 'Task\tStatus\n1.499\tO\n1.610\tC\n1.517\t');
    const saved = new ExcelJS.Workbook();
    const bytes = await result.workbook.xlsx.writeBuffer();
    // Check the actual differential-style XML, not just ExcelJS round trips.
    const styles = strFromU8(unzipSync(new Uint8Array(bytes))['xl/styles.xml']);
    const dxfs = styles.match(/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/)?.[0] ?? '';
    assert.match(dxfs, /<bgColor rgb="FFFCE4D6"\s*\/>/);
    assert.match(dxfs, /<bgColor rgb="FFC6E0B4"\s*\/>/);
    await saved.xlsx.load(bytes);
    const sheet = saved.getWorksheet(result.sheetName)!;
    const lastRow = result.terminals.length + 5;
    assert.equal(sheet.autoFilter, `A5:N${lastRow}`);
    const rules = sheet.model.conditionalFormattings!;
    assert.equal(rules.length, 1);
    assert.equal(rules[0].ref, `A6:N${lastRow}`);
    assert.deepEqual(rules[0].rules.map(rule => 'formulae' in rule ? rule.formulae : []), [['$N6="O"'], ['$N6="C"']]);
    assert.deepEqual(rules[0].rules.map(rule => rule.style?.fill), [
      { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' }, bgColor: { argb: 'FFFCE4D6' } },
      { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' }, bgColor: { argb: 'FFC6E0B4' } },
    ]);
    assert.equal(sheet.getCell('N7').text, '');
    assert.deepEqual(sheet.getCell('A7').fill, { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } });
  }
});

test('all cells export centered with consistent fonts, cleaned strings, gray stripes, and valid QR anchors', async () => {
  const result = await preparePonPower(await file('examples/PON_TEST_SHEET__17_.xlsx'));
  const sheet = result.workbook.getWorksheet(result.sheetName)!;
  const originalImageCount = sheet.getImages().length;
  sheet.getCell('B6').value = { richText: [{ text: '  Terminal\u00a0\u200b ', font: { name: 'Arial', size: 22 } }] };
  sheet.getCell('B6').alignment = { horizontal: 'left', vertical: 'top', indent: 3, textRotation: 45 };
  // Same anchor path used when the web app adds a QR on download.
  const imageId = result.workbook.addImage({ base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a0l8AAAAASUVORK5CYII=', extension: 'png' });
  sheet.addImage(imageId, { tl: { col: 11.63, row: 1.05 }, ext: { width: 84, height: 84 } });
  const bytes = await writeExcelWorkbook(result.workbook);
  const zip = unzipSync(bytes);
  const drawing = strFromU8(zip['xl/drawings/drawing1.xml']);
  assert.match(drawing, /<xdr:oneCellAnchor>/);
  assert.doesNotMatch(drawing, /<xdr:oneCellAnchor[^>]*editAs/);
  assert.match(drawing, /cx="800100" cy="800100"/);
  assert.ok(zip['xl/media/image1.png']);
  const saved = new ExcelJS.Workbook();
  await saved.xlsx.load(bytes);
  assert.equal(saved.getWorksheet(result.sheetName)!.getImages().length, originalImageCount + 1);
  assert.equal(saved.getWorksheet(result.sheetName)!.getCell('B6').text, 'Terminal');
  for (const ws of saved.worksheets) ws.eachRow(row => row.eachCell({ includeEmpty: true }, cell => {
    assert.equal(cell.alignment.horizontal, 'center');
    assert.equal(cell.alignment.vertical, 'middle');
    assert.equal(cell.alignment.indent, undefined);
    assert.equal(cell.font.name, 'Calibri');
    assert.equal(cell.font.size, 11);
  }));
});

test('overlapping PON groups are cable-scoped and chosen terminals alone survive export', async () => {
  const pon = new ExcelJS.Workbook();
  const sheet = pon.addWorksheet('PON TEST SHEET');
  sheet.getRow(5).values = PON_HEADERS;
  sheet.getCell('H2').value = 'CABLE ID: PON4250WRR';
  sheet.getRow(6).values = [1, 'first', '123', '45-46', 2];
  sheet.getRow(7).values = [2, 'second', '456', '45-48', 4];
  sheet.getRow(8).values = [3, 'third', '789', '49-50', 2];
  const input = new File([await pon.xlsx.writeBuffer()], 'duplicates.xlsx');
  const initial = await preparePonPower(input);
  assert.equal(initial.duplicateGroups.length, 1);
  assert.deepEqual(initial.duplicateGroups[0].counts, [45, 46]);
  const chosen = await preparePonPower(input, null, '', new Map([[initial.duplicateGroups[0].key, 6]]));
  assert.deepEqual(chosen.terminals.map(t => t.terminalName), ['second', 'third']);
  assert.equal(chosen.workbook.getWorksheet(chosen.sheetName)!.getCell('B6').text, 'second');
  assert.equal(chosen.workbook.getWorksheet(chosen.sheetName)!.getCell('B8').text, '');
  assert.equal(chosen.terminals[1].staggeredPort, 2);
  const parsed = parseExfoXlsx((await writeExcelWorkbook(chosen.workbook)).buffer as ArrayBuffer);
  assert.equal(parsed.terminals.length, 2);
  assert.equal(parsed.terminals[0].terminal, 'second');
  assert.equal(findPonDuplicateGroups([{ ...initial.terminals[0], cableId: 'other' }, initial.terminals[1]]).length, 0);
});

test('rejected distances clear imported and previously exported footage; confirming restores it', async () => {
  const result = await preparePonPower(await file('examples/PON_TEST_SHEET__17_.xlsx'));
  const sheet = result.workbook.getWorksheet(result.sheetName)!;
  sheet.getCell('H6').value = 500000;
  applyPonDistances(result, new Map([[7, 1000]]), new Set([6]));
  const saved = new ExcelJS.Workbook();
  await saved.xlsx.load(await writeExcelWorkbook(result.workbook));
  assert.equal(saved.getWorksheet(result.sheetName)!.getCell('H6').value, null);
  assert.equal(saved.getWorksheet(result.sheetName)!.getCell('H7').value, 1000);
  applyPonDistances(result, new Map([[6, 500000]]), new Set());
  assert.equal(sheet.getCell('H6').value, 500000);
});


