import type ExcelJS from 'exceljs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

export const cleanCellText = (text: string) => text.replace(/\u00a0/g, ' ').replace(/[\u200b-\u200d\ufeff]/g, '').replace(/\s*[\r\n]+\s*/g, ' ').trim();

/** Remove imported alignment/font overrides, including those inside rich text. */
export function normalizeWorkbookCells(workbook: ExcelJS.Workbook) {
  for (const sheet of workbook.worksheets) {
    const isPon = sheet.findRow(5)?.findCell(2)?.text === 'Terminal';
    const widths = new Map<number, number>();
    sheet.properties.defaultRowHeight = 15;
    sheet.eachRow({ includeEmpty: true }, row => {
      if (!isPon || row.number >= 5) row.height = 15;
      for (let c = 1; c <= Math.max(sheet.columnCount, 14); c++) {
        const cell = row.getCell(c);
        if (cell.isMerged && cell.master !== cell) continue;
        if (typeof cell.value === 'string') cell.value = cleanCellText(cell.value);
        else if (cell.value && typeof cell.value === 'object' && 'richText' in cell.value) {
          cell.value = cleanCellText(cell.value.richText.map(part => part.text).join(''));
        }
        cell.font = { name: 'Calibri', size: 11, bold: sheet.name === 'PON TEST SHEET' ? row.number <= 5 : row.number === 1 };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false, shrinkToFit: true };
        if (!cell.isMerged && (!isPon || row.number >= 5)) {
          // Excel has no persisted AutoFit command. Size to content with padding;
          // shrink unusually long values instead of creating oversized columns.
          const width = Math.max(9, cell.text.length * 1.1 + (row.number === (isPon ? 5 : 1) ? 3 : 2));
          widths.set(c, Math.max(widths.get(c) ?? 9, Math.min(isPon ? 45 : 55, width)));
        }
      }
    });
    widths.forEach((width, column) => { sheet.getColumn(column).width = width; });
    if (sheet.name === 'Orca') {
      sheet.getCell('A2').font = { name: 'Calibri', size: 11, bold: true };
    }
  }
}

/** ExcelJS emits editAs on oneCellAnchor, although OOXML permits it only on twoCellAnchor. */
export async function writeExcelWorkbook(workbook: ExcelJS.Workbook): Promise<Uint8Array> {
  normalizeWorkbookCells(workbook);
  const files = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()));
  // ExcelJS's default font writes color/family/scheme before size/name. OOXML's
  // font sequence requires size before color, then name/family/charset/scheme.
  const fontOrder = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
  const styles = strFromU8(files['xl/styles.xml']).replace(/<font>([\s\S]*?)<\/font>/g, (_font, body: string) => {
    const children = body.match(/<\w+\b[^>]*\/>/g) ?? [];
    children.sort((a, b) => fontOrder.indexOf(a.match(/^<(\w+)/)![1]) - fontOrder.indexOf(b.match(/^<(\w+)/)![1]));
    return `<font>${children.join('')}</font>`;
  });
  files['xl/styles.xml'] = strToU8(styles);
  for (const name of Object.keys(files)) {
    if (!/^xl\/drawings\/drawing\d+\.xml$/.test(name)) continue;
    const xml = strFromU8(files[name]);
    files[name] = strToU8(xml.replace(/(<xdr:oneCellAnchor\b[^>]*?)\s+editAs="[^"]*"/g, '$1'));
  }
  return zipSync(files);
}
