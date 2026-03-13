/**
 * pdfRenderer.ts
 *
 * Pure client-side PDF generation using the browser's native print engine.
 * Zero server dependencies — works on localhost, Vercel, anywhere.
 *
 * Strategy:
 *   1. Read the ExcelJS workbook we already built
 *   2. Render each sheet as pixel-accurate HTML tables
 *   3. Open a hidden iframe, inject the HTML + print CSS
 *   4. Call iframe.contentWindow.print() — browser shows Save as PDF dialog
 *
 * This is exactly what Excel does when you File → Print → Save as PDF.
 * No LibreOffice, no Puppeteer, no server round-trip.
 */

import ExcelJS from 'exceljs';

// ─── COLOUR UTILS ─────────────────────────────────────────────────────────────

const argbToHex = (argb?: string): string => {
  if (!argb || argb.length < 6) return '';
  // ExcelJS stores as AARRGGBB — drop the alpha byte
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  return `#${hex}`;
};

const getBgColor = (cell: ExcelJS.Cell): string => {
  const fill = cell.fill as ExcelJS.FillPattern | undefined;
  if (fill?.type === 'pattern' && fill.fgColor?.argb) {
    const hex = argbToHex(fill.fgColor.argb);
    if (hex && hex !== '#000000' && hex !== '#FFFFFF' && hex !== '#ffffff') return hex;
  }
  return '';
};

const getFontColor = (cell: ExcelJS.Cell): string => {
  const font = cell.font as ExcelJS.Font | undefined;
  if (font?.color?.argb) {
    const hex = argbToHex(font.color.argb);
    if (hex) return hex;
  }
  return '#000000';
};

// ─── BORDER UTILS ─────────────────────────────────────────────────────────────

const borderSide = (side?: ExcelJS.BorderStyle): string => {
  if (!side) return 'none';
  if (side === 'thin') return '0.5pt solid #000';
  if (side === 'medium') return '1pt solid #000';
  if (side === 'thick') return '1.5pt solid #000';
  return '0.5pt solid #000';
};

const cellBorderStyle = (cell: ExcelJS.Cell): string => {
  const b = cell.border as ExcelJS.Borders | undefined;
  if (!b) return '';
  const parts: string[] = [];
  if (b.top)    parts.push(`border-top: ${borderSide(b.top.style)}`);
  if (b.bottom) parts.push(`border-bottom: ${borderSide(b.bottom.style)}`);
  if (b.left)   parts.push(`border-left: ${borderSide(b.left.style)}`);
  if (b.right)  parts.push(`border-right: ${borderSide(b.right.style)}`);
  return parts.join('; ');
};

// ─── VALUE FORMATTING ─────────────────────────────────────────────────────────

const formatCellValue = (cell: ExcelJS.Cell): string => {
  const val = cell.value;
  if (val === null || val === undefined || val === '') return '';

  // Rich text
  if (typeof val === 'object' && 'richText' in val) {
    return (val as ExcelJS.CellRichTextValue).richText
      .map(r => escHtml(String(r.text ?? '')))
      .join('');
  }

  // Date
  if (val instanceof Date) {
    const dd = String(val.getDate()).padStart(2, '0');
    const mmm = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][val.getMonth()];
    const yy = String(val.getFullYear()).slice(-2);
    return `${dd}-${mmm}-${yy}`;
  }

  // Formula — use result if available
  if (typeof val === 'object' && 'formula' in val) {
    const r = (val as ExcelJS.CellFormulaValue).result;
    if (r instanceof Date) return formatCellValue({ ...cell, value: r } as ExcelJS.Cell);
    return r !== undefined && r !== null ? escHtml(String(r)) : '';
  }

  // Number with format
  if (typeof val === 'number') {
    const fmt = cell.numFmt;
    if (fmt?.includes('$')) return `$${val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    return escHtml(String(val));
  }

  // Plain string — preserve line breaks as <br>
  return escHtml(String(val)).replace(/\n/g, '<br>');
};

const escHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// ─── ALIGNMENT UTILS ─────────────────────────────────────────────────────────

const textAlign = (cell: ExcelJS.Cell): string => {
  const h = (cell.alignment as ExcelJS.Alignment | undefined)?.horizontal;
  if (h === 'center') return 'center';
  if (h === 'right')  return 'right';
  return 'left';
};

const vertAlign = (cell: ExcelJS.Cell): string => {
  const v = (cell.alignment as ExcelJS.Alignment | undefined)?.vertical;
  if (v === 'bottom') return 'bottom';
  if (v === 'middle') return 'middle';
  return 'top';
};

// ─── SHEET → HTML ─────────────────────────────────────────────────────────────

const sheetToHtml = (sheet: ExcelJS.Worksheet): string => {
  const maxRow = sheet.rowCount;
  const maxCol = sheet.columnCount;

  // Build column width map (pt — 1 Excel char ≈ 7px ≈ 5.25pt)
  const colWidths: number[] = [];
  for (let c = 1; c <= maxCol; c++) {
    const col = sheet.getColumn(c);
    colWidths[c] = Math.round((col.width ?? 8) * 5.25);
  }

  // Build merge map: "row,col" → { rowspan, colspan }
  type MergeInfo = { rowspan: number; colspan: number };
  const mergeMap = new Map<string, MergeInfo>();
  const skipSet  = new Set<string>();

  // ensure merges are loaded
  type MergeRange = { top: number; left: number; bottom: number; right: number };
  const merges = (sheet as unknown as { _merges?: Record<string, MergeRange> })._merges;
  if (merges) {
    Object.values(merges).forEach((merge: MergeRange) => {
      const key = `${merge.top},${merge.left}`;
      mergeMap.set(key, {
        rowspan: merge.bottom - merge.top + 1,
        colspan: merge.right  - merge.left + 1,
      });
      for (let r = merge.top; r <= merge.bottom; r++) {
        for (let c = merge.left; c <= merge.right; c++) {
          if (r !== merge.top || c !== merge.left) {
            skipSet.add(`${r},${c}`);
          }
        }
      }
    });
  }

  // Build HTML
  let html = '<table>';

  // colgroup for widths
  html += '<colgroup>';
  for (let c = 1; c <= maxCol; c++) {
    html += `<col style="width:${colWidths[c]}pt">`;
  }
  html += '</colgroup>';

  for (let r = 1; r <= maxRow; r++) {
    const row = sheet.getRow(r);
    const rowH = Math.round((row.height ?? 15) * 0.75); // Excel pt → CSS pt

    html += `<tr style="height:${rowH}pt">`;

    for (let c = 1; c <= maxCol; c++) {
      const key = `${r},${c}`;
      if (skipSet.has(key)) continue;

      const cell = sheet.getCell(r, c);
      const merge = mergeMap.get(key);

      const rs = merge?.rowspan ?? 1;
      const cs = merge?.colspan ?? 1;

      const bg    = getBgColor(cell);
      const color = getFontColor(cell);
      const font  = cell.font as ExcelJS.Font | undefined;
      const bold  = font?.bold ? 'font-weight:bold;' : '';
      const size  = font?.size ? `font-size:${font.size}pt;` : 'font-size:8pt;';
      const wrap  = (cell.alignment as ExcelJS.Alignment | undefined)?.wrapText
        ? 'white-space:pre-wrap; word-break:break-word;'
        : 'white-space:nowrap; overflow:hidden;';

      const styles = [
        cellBorderStyle(cell),
        bg    ? `background:${bg}` : '',
        `color:${color}`,
        bold,
        size,
        `font-family:Arial,sans-serif`,
        `text-align:${textAlign(cell)}`,
        `vertical-align:${vertAlign(cell)}`,
        wrap,
        'padding:1pt 2pt',
      ].filter(Boolean).join('; ');

      const rsAttr = rs > 1 ? ` rowspan="${rs}"` : '';
      const csAttr = cs > 1 ? ` colspan="${cs}"` : '';

      html += `<td${rsAttr}${csAttr} style="${styles}">${formatCellValue(cell)}</td>`;
    }

    html += '</tr>';
  }

  html += '</table>';
  return html;
};

// ─── FULL HTML PAGE ───────────────────────────────────────────────────────────

const buildHtmlPage = (sheets: { name: string; html: string }[]): string => {
  const bodies = sheets
    .map((s, i) => `
      ${i > 0 ? '<div class="page-break"></div>' : ''}
      <div class="sheet-title">${escHtml(s.name)}</div>
      <div class="sheet-wrap">${s.html}</div>
    `)
    .join('');

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 8pt; background: white; }

  .sheet-title {
    font-size: 7pt;
    color: #666;
    padding: 4pt 0 2pt 0;
    border-bottom: 0.5pt solid #ccc;
    margin-bottom: 4pt;
  }

  .sheet-wrap {
    overflow: hidden;
  }

  table {
    border-collapse: collapse;
    table-layout: fixed;
    width: 100%;
    /* fitToWidth=1: scale table to fit A4 printable width */
    max-width: 100%;
  }

  td {
    border: none; /* borders set per-cell via inline styles */
    overflow: hidden;
  }

  .page-break {
    page-break-before: always;
    padding-top: 8pt;
  }

  @page {
    size: A4 portrait;
    /* Matches original: L=0.1in R=0.1in T=0.25in B=0.25in */
    margin: 6.35mm 2.54mm 6.35mm 2.54mm;
  }

  @media print {
    .sheet-title { display: none; }
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
</style>
</head>
<body>
${bodies}
</body>
</html>`;
};

// ─── PUBLIC API ───────────────────────────────────────────────────────────────

/**
 * Convert an ExcelJS workbook to PDF by rendering it as HTML and
 * triggering the browser's native print-to-PDF dialog.
 *
 * @param workbook  The already-built ExcelJS workbook
 * @param sheetNames  Which sheet names to include (in order). Defaults to all.
 * @param pdfFilename  Suggested filename shown in the browser Save dialog
 */
export const printWorkbookAsPdf = (
  workbook: ExcelJS.Workbook,
  sheetNames?: string[],
  pdfFilename?: string,
): void => {
  const targetSheets = sheetNames
    ? sheetNames.map(n => workbook.getWorksheet(n)).filter(Boolean) as ExcelJS.Worksheet[]
    : workbook.worksheets;

  const sheets = targetSheets.map(ws => ({
    name: ws.name,
    html: sheetToHtml(ws),
  }));

  const html = buildHtmlPage(sheets);

  // Inject into a hidden iframe and print
  const iframe = document.createElement('iframe');
  iframe.style.cssText = 'position:fixed;top:-9999px;left:-9999px;width:0;height:0;border:none';
  document.body.appendChild(iframe);

  const doc = iframe.contentDocument!;
  doc.open();
  doc.write(html);
  doc.close();

  // Give images/fonts a moment, then print
  iframe.onload = () => {
    // Set the suggested filename via document.title — browser uses this for Save As
    if (pdfFilename) iframe.contentDocument!.title = pdfFilename.replace(/\.xlsx$/i, '');

    setTimeout(() => {
      iframe.contentWindow!.focus();
      iframe.contentWindow!.print();
      // Remove iframe after print dialog closes
      setTimeout(() => document.body.removeChild(iframe), 1000);
    }, 300);
  };
};