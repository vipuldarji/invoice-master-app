import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { MasterData } from '../excelGenerator';

// ─── CONSTANTS ───────────────────────────────────────────────────────────────
const BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' }, left: { style: 'thin' },
  bottom: { style: 'thin' }, right: { style: 'thin' },
};
const FONT_BOLD   = { bold: true,  size: 9,  name: 'Arial' };
const FONT_NORMAL = { bold: false, size: 9,  name: 'Arial' };
const FONT_TITLE  = { bold: true,  size: 14, name: 'Arial', underline: true };
const FONT_SMALL  = { bold: true,  size: 7,  name: 'Arial' };

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const toDate = (s: string): Date | string => {
  if (!s) return s;
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
};

const formatExpiry = (s: string): string => {
  if (!s) return s;
  if (/^\d{2}\/\d{4}$/.test(s)) return s;
  if (s.includes('-')) {
    const [year, month] = s.split('-');
    return `${month}/${year}`;
  }
  return s;
};

const formatWeight = (w: string | number): string => {
  const n = Number(w);
  return isNaN(n) ? String(w) : Number.isInteger(n) ? String(n) : String(n);
};

// ─── MAIN EXPORT ─────────────────────────────────────────────────────────────
export const addPackingListSheet = (workbook: ExcelJS.Workbook, data: MasterData) => {
  const sheet = workbook.addWorksheet('PACKING', {
    views: [{ showGridLines: false }],
    pageSetup: {
      paperSize: 9,
      orientation: 'portrait',
      fitToPage: true, fitToWidth: 1, fitToHeight: 1,
      margins: { left: 0.1, right: 0.1, top: 0.25, bottom: 0.25, header: 0, footer: 0 },
    },
  });

  // ── Column widths (match original PACKING sheet exactly) ──────────────────
  // A(1)=margin, B(2)=marks, C(3)=description, D(4)=HSN, E(5)=pack,
  // F(6)=batch, G(7)=expiry, H(8)=UQC, I(9)=qty, J(10)=gross wt, K(11)=net wt
  sheet.columns = [
    { width: 2.7   }, // A — left margin
    { width: 13.28 }, // B — Marks & Nos / Sr.No
    { width: 84.28 }, // C — Description of Goods (exact from original)
    { width: 15.28 }, // D — HSN CODE
    { width: 18.56 }, // E — Pack
    { width: 23.99 }, // F — Batch No.
    { width: 21.13 }, // G — Expiry Date
    { width: 15.28 }, // H — Standard UQC
    { width: 14.70 }, // I — Quantity (NOS)
    { width: 18.28 }, // J — Gross Weight in KGS
    { width: 13.28 }, // K — Nett Weight in KGS
  ];

  sheet.pageSetup.printArea = 'A1:K69';

  // ── Cell helpers ──────────────────────────────────────────────────────────
  const sc = (
    coord: string,
    value: ExcelJS.CellValue,
    bold = false,
    numFmt?: string,
    hAlign: ExcelJS.Alignment['horizontal'] = 'left',
    vAlign: ExcelJS.Alignment['vertical']   = 'middle',
  ) => {
    const cell = sheet.getCell(coord);
    cell.value = value;
    cell.border = BORDER;
    cell.font = bold ? FONT_BOLD : FONT_NORMAL;
    cell.alignment = { horizontal: hAlign, vertical: vAlign, wrapText: true };
    if (numFmt) cell.numFmt = numFmt;
    return cell;
  };

  const msc = (
    range: string,
    firstCoord: string,
    value: ExcelJS.CellValue,
    bold = false,
    numFmt?: string,
    hAlign: ExcelJS.Alignment['horizontal'] = 'left',
    vAlign: ExcelJS.Alignment['vertical']   = 'middle',
  ) => {
    sheet.mergeCells(range);
    return sc(firstCoord, value, bold, numFmt, hAlign, vAlign);
  };

  const borderRow = (r: number, fromCol = 2, toCol = 11) => {
    for (let c = fromCol; c <= toCol; c++) {
      sheet.getCell(r, c).border = BORDER;
      sheet.getCell(r, c).font = FONT_NORMAL;
    }
  };

  // Drug lic strings
  const lic1str = data.drugLicDate1
    ? `${data.drugLicNo1} Dated: ${data.drugLicDate1}`
    : data.drugLicNo1;
  const lic2str = data.drugLicDate2
    ? `${data.drugLicNo2} Dated: ${data.drugLicDate2}`
    : data.drugLicNo2;

  // ════════════════════════════════════════════════════════════════════════════
  // R1 — blank spacer
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(1).height = 17.25;

  // ════════════════════════════════════════════════════════════════════════════
  // R2 — PACKING LIST title
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(2).height = 31.5;
  const titleCell = msc('B2:K2', 'B2', 'PACKING LIST', true, undefined, 'center', 'middle');
  titleCell.font = FONT_TITLE;

  // ════════════════════════════════════════════════════════════════════════════
  // R3 — EXPORTER label | PACKING LIST No. | invoiceNo | PACKING LIST DATE | date
  // Original: B='EXPORTER :' | F='PACKING LIST No.' | G=invoiceNo
  //           H='PACKING LIST DATE' | J=invoiceDate (date value)
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(3).height = 24;
  sc('B3', 'EXPORTER :', true);
  sc('C3', '');
  sc('D3', '');
  sc('E3', '');
  sc('F3', 'PACKING LIST No.       ', true);
  sc('G3', data.invoiceNo,  true);
  msc('H3:I3', 'H3', 'PACKING LIST DATE          ', true);
  sc('J3', toDate(data.invoiceDate) as ExcelJS.CellValue, true, 'dd-mmm-yy');
  sc('K3', '');

  // ════════════════════════════════════════════════════════════════════════════
  // R4 — exporterName | IEC No. | iecNo
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(4).height = 22.5;
  sc('B4', data.exporterName, true);
  sc('C4', ''); sc('D4', ''); sc('E4', '');
  sc('F4', 'IEC No.:', true);
  sc('G4', '');
  sc('H4', data.iecNo, true);
  sc('I4', ''); sc('J4', ''); sc('K4', '');

  // ════════════════════════════════════════════════════════════════════════════
  // R5 — addrLine1 | Company GSTN | gstNo
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(5).height = 18.75;
  sc('B5', data.exporterAddressLine1);
  sc('C5', ''); sc('D5', ''); sc('E5', '');
  sc('F5', "Company's GSTN No. :  ", true);
  sc('G5', '');
  sc('H5', data.companyGstNo, true);
  sc('I5', ''); sc('J5', ''); sc('K5', '');

  // ════════════════════════════════════════════════════════════════════════════
  // R6 — addrLine2 | IGST STATUS | gstStatus
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(6).height = 22.5;
  sc('B6', data.exporterAddressLine2);
  sc('C6', ''); sc('D6', ''); sc('E6', '');
  sc('F6', 'IGST PAYMENT STATUS : ', true);
  sc('G6', '');
  sc('H6', data.gstStatus, true);
  sc('I6', ''); sc('J6', ''); sc('K6', '');

  // ════════════════════════════════════════════════════════════════════════════
  // R7 — addrLine3 | Drug Lic No. | lic1
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(7).height = 18.75;
  sc('B7', data.exporterAddressLine3);
  sc('C7', ''); sc('D7', ''); sc('E7', '');
  sc('F7', 'Drug Lic No.:', true);
  sc('G7', '');
  msc('H7:K7', 'H7', lic1str, true);

  // ════════════════════════════════════════════════════════════════════════════
  // R8 — phone | lic2 (no label)
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(8).height = 18.75;
  sc('B8', `PHONE NO-${data.exporterPhone}`);
  sc('C8', ''); sc('D8', ''); sc('E8', '');
  sc('F8', ''); sc('G8', '');
  msc('H8:K8', 'H8', lic2str, true);

  // ════════════════════════════════════════════════════════════════════════════
  // R9 — email | Buyer's Order Ref
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(9).height = 18;
  sc('B9', `Email ID: ${data.exporterEmail}`);
  sc('C9', ''); sc('D9', ''); sc('E9', '');
  sc('F9', "Buyer's Order Ref.No. :", true);
  sc('G9', '');
  msc('H9:K9', 'H9', data.buyerOrderRef || '');

  // ════════════════════════════════════════════════════════════════════════════
  // R10 — blank | Exporter Ref
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(10).height = 18;
  sc('B10', ''); sc('C10', ''); sc('D10', ''); sc('E10', '');
  sc('F10', 'Exporter Ref. and Date :', true);
  sc('G10', '');
  msc('H10:K10', 'H10', data.exporterRef || '');

  // ════════════════════════════════════════════════════════════════════════════
  // R11 — CONSIGNEE : | BUYER (IF OTHER THAN CONSIGNEE) :
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(11).height = 18.75;
  sc('B11', 'CONSIGNEE  :', true);
  sc('C11', ''); sc('D11', ''); sc('E11', '');
  sc('F11', ''); sc('G11', ''); sc('H11', '');
  msc('I11:K11', 'I11', 'BUYER (IF OTHER THAN CONSIGNEE)  :', true);

  // ════════════════════════════════════════════════════════════════════════════
  // R12 — consigneeName | buyerName
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(12).height = 22.5;
  sc('B12', data.consigneeName, true);
  sc('C12', ''); sc('D12', ''); sc('E12', '');
  sc('F12', ''); sc('G12', ''); sc('H12', '');
  msc('I12:K12', 'I12', data.buyerName || '');

  // ════════════════════════════════════════════════════════════════════════════
  // R13 — consigneeAddress | blank buyer address
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(13).height = 18.75;
  sc('B13', data.consigneeAddress);
  sc('C13', ''); sc('D13', ''); sc('E13', '');
  sc('F13', ''); sc('G13', ''); sc('H13', '');
  msc('I13:K13', 'I13', '');

  // ════════════════════════════════════════════════════════════════════════════
  // R14–R17 — blank filler rows (original has 4 blank rows before logistics)
  // ════════════════════════════════════════════════════════════════════════════
  for (let r = 14; r <= 17; r++) {
    sheet.getRow(r).height = r <= 15 ? 17.25 : 21.6;
    borderRow(r);
  }

  // ════════════════════════════════════════════════════════════════════════════
  // R18–R23 — Logistics block (identical layout to INVOICE sheet)
  // R18: labels  PRE-CARRIAGE BY | PLACE OF RECEIPT | COUNTRY OF ORIGIN | INDIA
  // R19: values  By AIR | Ahmedabad Airport | COUNTRY OF FINAL DEST | KENYA
  // R20: labels  VESSEL/FLIGHT | PORT OF LOADING | TERMS OF DELIVERY | value
  // R21: values  blank | Ahmedabad Airport | blank
  // R22: labels  PORT OF DISCHARGE | FINAL DESTINATION | PAYMENT TERMS | value
  // R23: values  NAIROBI,KENYA | KENYA | blank
  // ════════════════════════════════════════════════════════════════════════════
  const logH = 21.6;

  sheet.getRow(18).height = logH;
  msc('B18:C18', 'B18', '  PRE-CARRIAGE BY',            true, undefined, 'center');
  msc('D18:E18', 'D18', 'PLACE OF RECEIPT',             true, undefined, 'center');
  sc ('F18', 'COUNTRY OF ORIGIN OF GOODS ', true, undefined, 'center');
  sc ('G18', '', false, undefined, 'center');
  msc('H18:I18', 'H18', 'INDIA ', true, undefined, 'center');
  sc ('J18', '', false, undefined, 'center');
  sc ('K18', '', false, undefined, 'center');

  sheet.getRow(19).height = 25.15;
  msc('B19:C19', 'B19', data.preCarriage,               false, undefined, 'center');
  msc('D19:E19', 'D19', data.placeOfReceipt,            false, undefined, 'center');
  sc ('F19', 'COUNTRY OF FINAL DESTINATION ', true, undefined, 'center');
  sc ('G19', '', false, undefined, 'center');
  msc('H19:K19', 'H19', data.finalDestination,          false, undefined, 'center');

  sheet.getRow(20).height = 25.15;
  msc('B20:C20', 'B20', '  VESSEL/FLIGHT No.',          true, undefined, 'center');
  msc('D20:E20', 'D20', 'PORT OF LOADING',              true, undefined, 'center');
  sc ('F20', 'TERMS OF DELIVERY  ', true, undefined, 'center');
  sc ('G20', '', false, undefined, 'center');
  msc('H20:K20', 'H20', data.termsOfDelivery,           false, undefined, 'center');

  sheet.getRow(21).height = 25.15;
  msc('B21:C21', 'B21', data.vesselFlight || '',        false, undefined, 'center');
  msc('D21:E21', 'D21', data.portOfLoading,             false, undefined, 'center');
  sc ('F21', ''); sc('G21', '');
  msc('H21:K21', 'H21', '',                             false, undefined, 'center');

  sheet.getRow(22).height = 25.15;
  msc('B22:C22', 'B22', '  PORT OF DISCHARGE',          true, undefined, 'center');
  msc('D22:E22', 'D22', 'FINAL DESTINATION',            true, undefined, 'center');
  sc ('F22', 'PAYMENT TERMS ', true, undefined, 'center');
  sc ('G22', '');
  msc('H22:K23', 'H22', data.paymentTerms,              false, undefined, 'center');

  sheet.getRow(23).height = 25.15;
  msc('B23:C23', 'B23', data.portOfDischarge,           false, undefined, 'center');
  msc('D23:E23', 'D23', data.finalDestination,          false, undefined, 'center');
  sc ('F23', ''); sc('G23', '');
  // H23:K23 already merged above via H22:K23

  // ════════════════════════════════════════════════════════════════════════════
  // R24 — TABLE COLUMN HEADERS
  // PACKING differs from INVOICE: J=Gross Weight in KGS, K=Nett Weight in KGS
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(24).height = 73.15;
  const headers: [number, string][] = [
    [2,  'Marks & Nos.'],
    [3,  'Description of Goods'],
    [4,  'HSN CODE'],
    [5,  'Pack'],
    [6,  'Batch No.'],
    [7,  'Expiry Date'],
    [8,  'Standard UQC'],
    [9,  'Quantity (NOS)'],
    [10, 'Gross Weight in KGS'],
    [11, 'Nett Weight in KGS'],
  ];
  headers.forEach(([col, label]) => {
    const cell = sheet.getCell(24, col);
    cell.value = label;
    cell.font = FONT_BOLD;
    cell.border = BORDER;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // ════════════════════════════════════════════════════════════════════════════
  // R25 — blank spacer (matches original)
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(25).height = 17.25;
  borderRow(25);

  // ════════════════════════════════════════════════════════════════════════════
  // ITEM ROWS — start at R26, 2 rows per item
  //
  // Item row:  B=srNo | C=name\ndesc\n | D=HSN(int) | E=pack | F=batch
  //            G=expiry(MM/YYYY) | H=UQC | I=qty | J=grossWeight | K=netWeight
  //
  // NOTE: Only item 1 has actual grossWeight/netWeight values (59/50).
  //       Items 2-13 have None — matches original (only row 1 shows totals in J/K).
  //
  // State row: B=border | C:K merged = "STATE CODE :  xx, GSTIN No.: ... DISTRICT CODE :  ..."
  // ════════════════════════════════════════════════════════════════════════════
  let row = 26;
  let totalQty = 0;

  data.items.forEach((item, idx) => {
    const isFirst = idx === 0;
    // Per-item row heights from original PACKING sheet
    const PACK_ITEM_HEIGHTS = [43.15, 46.9, 47.25, 41.45, 47.25, 47.25, 57.0, 47.25, 47.25, 63.0, 47.25, 47.25, 47.25];
    sheet.getRow(row).height = PACK_ITEM_HEIGHTS[idx] ?? 45;

    // B: Sr.No
    const bCell = sheet.getCell(row, 2);
    bCell.value = idx + 1;
    bCell.border = BORDER;
    bCell.font = FONT_NORMAL;
    bCell.alignment = { horizontal: 'center', vertical: 'top' };

    // C: productName + "\n" + description (same format as INVOICE)
    const cCell = sheet.getCell(row, 3);
    cCell.value = item.description
      ? `${item.productName}\n${item.description}\n`
      : item.productName;
    cCell.border = BORDER;
    cCell.font = FONT_NORMAL;
    cCell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };

    // D: HSN as integer
    const dCell = sheet.getCell(row, 4);
    dCell.value = Number(item.hsnSac) || item.hsnSac;
    dCell.border = BORDER;
    dCell.font = FONT_NORMAL;
    dCell.alignment = { horizontal: 'center', vertical: 'top' };

    // E: Pack
    const eCell = sheet.getCell(row, 5);
    eCell.value = item.packSize;
    eCell.border = BORDER;
    eCell.font = FONT_NORMAL;
    eCell.alignment = { horizontal: 'center', vertical: 'top' };

    // F: Batch No.
    const fCell = sheet.getCell(row, 6);
    fCell.value = item.batchNo;
    fCell.border = BORDER;
    fCell.font = FONT_NORMAL;
    fCell.alignment = { horizontal: 'center', vertical: 'top' };

    // G: Expiry Date MM/YYYY
    const gCell = sheet.getCell(row, 7);
    gCell.value = formatExpiry(item.expDate);
    gCell.border = BORDER;
    gCell.font = FONT_NORMAL;
    gCell.alignment = { horizontal: 'center', vertical: 'top' };

    // H: Standard UQC " {netWeight} {uom}"
    const hCell = sheet.getCell(row, 8);
    hCell.value = ` ${item.netWeight} ${item.uom}`;
    hCell.border = BORDER;
    hCell.font = FONT_NORMAL;
    hCell.alignment = { horizontal: 'center', vertical: 'top' };

    // I: Quantity (NOS)
    const q = Number(item.quantity) || 0;
    const iCell = sheet.getCell(row, 9);
    iCell.value = q;
    iCell.border = BORDER;
    iCell.font = FONT_NORMAL;
    iCell.alignment = { horizontal: 'center', vertical: 'top' };
    totalQty += q;

    // J: Gross Weight in KGS — only first item shows the total gross weight
    //    (matches original: R26J=59, R28J=None, R30J=None ... R50J=None)
    const jCell = sheet.getCell(row, 10);
    jCell.value = isFirst ? (Number(data.totalGrossWeight) || null) : null;
    jCell.border = BORDER;
    jCell.font = FONT_NORMAL;
    jCell.alignment = { horizontal: 'center', vertical: 'top' };

    // K: Nett Weight in KGS — only first item shows the total net weight
    const kCell = sheet.getCell(row, 11);
    kCell.value = isFirst ? (Number(data.totalNetWeight) || null) : null;
    kCell.border = BORDER;
    kCell.font = FONT_NORMAL;
    kCell.alignment = { horizontal: 'center', vertical: 'top' };

    row++;

    // State code sub-row
    sheet.getRow(row).height = 30;
    sheet.getCell(row, 2).border = BORDER;
    sheet.getCell(row, 2).font = FONT_NORMAL;

    sheet.mergeCells(`C${row}:K${row}`);
    const scCell = sheet.getCell(row, 3);
    scCell.value = `STATE CODE :  ${item.stateCode}, GSTIN No.: ${item.supplierGstin}                                            DISTRICT CODE :  ${item.distCode}`;
    scCell.border = BORDER;
    scCell.font = { size: 8, name: 'Arial' };
    scCell.alignment = { vertical: 'middle', horizontal: 'left' };

    row++;
  });

  // ════════════════════════════════════════════════════════════════════════════
  // Pad blank rows so totals land at R63 (same target as INVOICE)
  // ════════════════════════════════════════════════════════════════════════════
  const totalsTargetRow = 63;
  while (row < totalsTargetRow) {
    sheet.getRow(row).height = 12;
    borderRow(row);
    row++;
  }

  // ════════════════════════════════════════════════════════════════════════════
  // R63 — TOTALS ROW
  // Original: B:C=' No. of Corrugated Boxes :   06' | D:E=' Gross Weight :  59 KGS'
  //           F:G=' Nett Weight :  50 KGS' | H=blank | I=totalQty | J=grossTotal | K=netTotal
  // ════════════════════════════════════════════════════════════════════════════
  sheet.getRow(row).height = 25.5;

  msc('B63:C63', 'B63',
    ` No. of  Corrugated Boxes :   ${data.totalCorrugatedBoxes}`,
    true, undefined, 'left', 'middle');

  msc('D63:E63', 'D63',
    ` Gross Weight :  ${formatWeight(data.totalGrossWeight)} KGS`,
    true, undefined, 'left', 'middle');

  msc('F63:G63', 'F63',
    ` Nett Weight :  ${formatWeight(data.totalNetWeight)} KGS`,
    true, undefined, 'left', 'middle');

  // H: blank
  const h63 = sheet.getCell(63, 8);
  h63.border = BORDER; h63.font = FONT_NORMAL;

  // I: total qty
  const i63 = sheet.getCell(63, 9);
  i63.value = totalQty;
  i63.border = BORDER;
  i63.font = FONT_BOLD;
  i63.alignment = { horizontal: 'center', vertical: 'middle' };

  // J: gross weight total
  const j63 = sheet.getCell(63, 10);
  j63.value = Number(data.totalGrossWeight) || 0;
  j63.border = BORDER;
  j63.font = FONT_BOLD;
  j63.alignment = { horizontal: 'center', vertical: 'middle' };

  // K: net weight total
  const k63 = sheet.getCell(63, 11);
  k63.value = Number(data.totalNetWeight) || 0;
  k63.border = BORDER;
  k63.font = FONT_BOLD;
  k63.alignment = { horizontal: 'center', vertical: 'middle' };

  row = 64;

  // ════════════════════════════════════════════════════════════════════════════
  // R64–R68 — Box dimensions + signatory
  // Original:
  //   R64: B=Dim Box#01 | D=Dim Box#06 | I='FOR COSMOS HEALTHCARE,'
  //   R65: B=Dim Box#02  (D:H merged blank)
  //   R66: B=Dim Box#03
  //   R67: B=Dim Box#04
  //   R68: B=Dim Box#05
  //   R69: I='AUTHORISED SIGNATORY'
  // ════════════════════════════════════════════════════════════════════════════
  const dims = data.boxDimensions || [];

  // Build dim label helper
  const dimLabel = (idx: number): string => {
    const d = dims[idx];
    if (!d) return '';
    return ` Dimension for Box #  ${d.boxNo} :   ${d.dimensions}`;
  };

  // R64
  sheet.getRow(64).height = 20.25;
  sc('B64', dimLabel(0), false);
  sc('C64', '');
  sc('D64', dimLabel(5) || '', false); // Box #06 goes in D64
  sc('E64', '');
  sc('F64', ''); sc('G64', ''); sc('H64', '');
  msc('I64:K64', 'I64', `FOR ${data.exporterName.toUpperCase()},`, true, undefined, 'center', 'middle');

  // R65–R68: Box dims 02–05
  [65, 66, 67, 68].forEach((r, i) => {
    sheet.getRow(r).height = 18.75;
    sc(`B${r}`, dimLabel(i + 1), false);
    sc(`C${r}`, '');
    // D:H merged blank
    msc(`D${r}:H${r}`, `D${r}`, '', false);
    // I:K blank (part of signatory block — keep bordered)
    sheet.getCell(r, 9).border = BORDER;
    sheet.getCell(r, 9).font = FONT_NORMAL;
    sheet.getCell(r, 10).border = BORDER;
    sheet.getCell(r, 10).font = FONT_NORMAL;
    sheet.getCell(r, 11).border = BORDER;
    sheet.getCell(r, 11).font = FONT_NORMAL;
  });

  // R69 — AUTHORISED SIGNATORY
  sheet.getRow(69).height = 12;
  borderRow(69, 2, 8);
  msc('I69:K69', 'I69', 'AUTHORISED SIGNATORY', true, undefined, 'center', 'middle');
};

// ─── STANDALONE XLSX EXPORT ───────────────────────────────────────────────────
export const generatePackingList = async (data: MasterData) => {
  const workbook = new ExcelJS.Workbook();
  addPackingListSheet(workbook, data);
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(
    new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    `PackingList_${data.invoiceNo}.xlsx`,
  );
};