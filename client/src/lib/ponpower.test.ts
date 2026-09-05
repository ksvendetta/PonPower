import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import { unzipSync, strFromU8 } from 'fflate';
import { preparePonPower, staggeredFilename } from './ponpower';

async function file(path: string) {
  return new File([await readFile(path)], path.split('/').at(-1)!);
}

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
    assert.deepEqual(sheet.getCell('A7').fill, { type: 'pattern', pattern: 'none' });
  }
});


