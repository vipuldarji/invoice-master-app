import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { MasterData } from '../excelGenerator';

// --- CONSTANTS ---
const BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
};
const FONT_BOLD   = { bold: true, size: 9, name: 'Arial' };
const FONT_NORMAL = { size: 9, name: 'Arial' };
const FONT_TITLE  = { bold: true, size: 14, underline: true, name: 'Arial' };
const FONT_SMALL  = { size: 7, name: 'Arial', bold: true };

// --- AMOUNT IN WORDS ---
const numberToWords = (num: number): string => {
  if (num === 0) return 'ZERO';
  const ones  = ['','ONE','TWO','THREE','FOUR','FIVE','SIX','SEVEN','EIGHT','NINE'];
  const teens = ['TEN','ELEVEN','TWELVE','THIRTEEN','FOURTEEN','FIFTEEN','SIXTEEN','SEVENTEEN','EIGHTEEN','NINETEEN'];
  const tens  = ['','','TWENTY','THIRTY','FORTY','FIFTY','SIXTY','SEVENTY','EIGHTY','NINETY'];
  const convert = (n: number): string => {
    if (n < 10)      return ones[n];
    if (n < 20)      return teens[n - 10];
    if (n < 100)     return tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
    if (n < 1000)    return ones[Math.floor(n / 100)] + ' HUNDRED' + (n % 100 ? ' ' + convert(n % 100) : '');
    if (n < 1000000) return convert(Math.floor(n / 1000)) + ' THOUSAND' + (n % 1000 ? ' ' + convert(n % 1000) : '');
    return convert(Math.floor(n / 1000000)) + ' MILLION' + (n % 1000000 ? ' ' + convert(n % 1000000) : '');
  };
  const dollars = Math.floor(num);
  const cents   = Math.round((num - dollars) * 100);
  let result = 'US DOLLARS ' + convert(dollars);
  if (cents > 0) result += ' AND ' + convert(cents) + ' CENTS';
  return result + ' ONLY';
};

// Convert YYYY-MM-DD string to JS Date for Excel date serial
const toDate = (s: string): Date | string => {
  if (!s) return s;
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
};

// Format expiry YYYY-MM-DD → MM/YYYY
const formatExpiry = (s: string): string => {
  if (!s) return s;
  if (/^\d{2}\/\d{4}$/.test(s)) return s;
  if (s.includes('-')) {
    const parts = s.split('-');
    if (parts.length >= 2) return `${parts[1]}/${parts[0]}`;
  }
  return s;
};

// FIX-2: Strip trailing decimals from whole-number weights
// e.g. totalGrossWeight=59 → "59", not "59.000"
const formatWeight = (w: string | number): string => {
  const n = Number(w);
  if (isNaN(n)) return String(w);
  return Number.isInteger(n) ? String(n) : String(n);
};

export const addCommercialInvoiceSheet = (workbook: ExcelJS.Workbook, data: MasterData) => {
  const sheet = workbook.addWorksheet('INVOICE', {
    views: [{ showGridLines: false }],
    pageSetup: {
      paperSize: 9, orientation: 'portrait', fitToPage: true,
      margins: { left: 0.2, right: 0.2, top: 0.2, bottom: 0.2, header: 0, footer: 0 }
    }
  });

  // 11 columns: A–K  (col index 1–11)
  // Original Invoice.xls uses B(2) through K(11) for content; A(1) is narrow margin
  sheet.columns = [
    { width: 5  }, // A col 1 — narrow left margin
    { width: 6  }, // B col 2 — Sr.No
    { width: 38 }, // C col 3 — Description of Goods (widest)
    { width: 10 }, // D col 4 — HSN CODE
    { width: 8  }, // E col 5 — Pack
    { width: 10 }, // F col 6 — Batch No.
    { width: 10 }, // G col 7 — Expiry Date
    { width: 10 }, // H col 8 — Standard UQC
    { width: 10 }, // I col 9 — Quantity
    { width: 12 }, // J col 10 — Rate Per Unit
    { width: 12 }, // K col 11 — Amount
  ];

  // Helper: set single cell
  const sc = (
    coord: string,
    value: ExcelJS.CellValue,
    bold = false,
    numFmt?: string,
    hAlign: ExcelJS.Alignment['horizontal'] = 'left',
    vAlign: ExcelJS.Alignment['vertical'] = 'middle'
  ) => {
    const cell = sheet.getCell(coord);
    cell.value = value;
    cell.border = BORDER;
    cell.font = bold ? FONT_BOLD : FONT_NORMAL;
    cell.alignment = { horizontal: hAlign, vertical: vAlign, wrapText: true };
    if (numFmt) cell.numFmt = numFmt;
    return cell;
  };

  // Helper: merge then set first cell
  const msc = (
    range: string,
    firstCoord: string,
    value: ExcelJS.CellValue,
    bold = false,
    numFmt?: string,
    hAlign: ExcelJS.Alignment['horizontal'] = 'left',
    vAlign: ExcelJS.Alignment['vertical'] = 'middle'
  ) => {
    sheet.mergeCells(range);
    return sc(firstCoord, value, bold, numFmt, hAlign, vAlign);
  };

  // Drug lic strings
  const lic1str = data.drugLicDate1 ? `${data.drugLicNo1} Dated: ${data.drugLicDate1}` : data.drugLicNo1;
  const lic2str = data.drugLicDate2 ? `${data.drugLicNo2} Dated: ${data.drugLicDate2}` : data.drugLicNo2;

  // ═══════════════════════════════════════════════════════
  // ROW 1: empty — matches original (row 1 blank in Invoice.xls)
  // ═══════════════════════════════════════════════════════
  sheet.getRow(1).height = 8;

  // ═══════════════════════════════════════════════════════
  // ROW 2: INVOICE title — B2:K2 merged, centered
  // Original: R2C2 = 'INVOICE'
  // ═══════════════════════════════════════════════════════
  sheet.getRow(2).height = 25;
  const titleCell = msc('B2:K2', 'B2', 'INVOICE', true, undefined, 'center', 'middle');
  titleCell.font = FONT_TITLE;

  // ═══════════════════════════════════════════════════════
  // ROW 3: EXPORTER label | INVOICE No. | Invoice Date
  // Original: R3C2='EXPORTER :' | R3C6='INVOICE No.' | R3C7=invoiceNo | R3C8=' INVOICE DATE ' | R3C10=invoiceDate
  // ═══════════════════════════════════════════════════════
  sheet.getRow(3).height = 18;
  sc('B3', 'EXPORTER :', true);
  sc('C3', '', false);
  sc('D3', '', false);
  sc('E3', '', false);
  sc('F3', 'INVOICE No.      ', true);
  sc('G3', data.invoiceNo, true);
  sc('H3', ' INVOICE DATE ', true);
  sc('I3', '', false);
  sc('J3', toDate(data.invoiceDate) as ExcelJS.CellValue, true, 'dd-mmm-yy');
  sc('K3', '', false);

  // ═══════════════════════════════════════════════════════
  // ROW 4: exporterName | IEC No.
  // Original: R4C2=exporterName | R4C6='IEC No.' | R4C8=iecNo
  // ═══════════════════════════════════════════════════════
  sheet.getRow(4).height = 16;
  sc('B4', data.exporterName, true);
  sc('C4', '', false);
  sc('D4', '', false);
  sc('E4', '', false);
  sc('F4', 'IEC No.', true);
  sc('G4', '', false);
  sc('H4', data.iecNo, true);
  sc('I4', '', false);
  sc('J4', '', false);
  sc('K4', '', false);

  // ═══════════════════════════════════════════════════════
  // ROW 5: addr1 | Company GSTN
  // Original: R5C2=addr1 | R5C6="Company's GSTN No." | R5C8=gstNo
  // ═══════════════════════════════════════════════════════
  sheet.getRow(5).height = 16;
  sc('B5', data.exporterAddressLine1, false);
  sc('C5', '', false);
  sc('D5', '', false);
  sc('E5', '', false);
  sc('F5', "Company's GSTN No. ", true);
  sc('G5', '', false);
  sc('H5', data.companyGstNo, true);
  sc('I5', '', false);
  sc('J5', '', false);
  sc('K5', '', false);

  // ═══════════════════════════════════════════════════════
  // ROW 6: addr2 | IGST PAYMENT STATUS
  // Original: R6C2=addr2 | R6C6='IGST PAYMENT STATUS :' | R6C8=gstStatus
  // ═══════════════════════════════════════════════════════
  sheet.getRow(6).height = 16;
  sc('B6', data.exporterAddressLine2, false);
  sc('C6', '', false);
  sc('D6', '', false);
  sc('E6', '', false);
  sc('F6', 'IGST PAYMENT STATUS : ', true);
  sc('G6', '', false);
  sc('H6', data.gstStatus, true);
  sc('I6', '', false);
  sc('J6', '', false);
  sc('K6', '', false);

  // ═══════════════════════════════════════════════════════
  // ROW 7: addr3 | Drug Lic No. label | lic1
  // Original: R7C2=addr3 | R7C6='Drug Lic No.' | R7C8=lic1
  // ═══════════════════════════════════════════════════════
  sheet.getRow(7).height = 16;
  sc('B7', data.exporterAddressLine3, false);
  sc('C7', '', false);
  sc('D7', '', false);
  sc('E7', '', false);
  sc('F7', 'Drug Lic No.', true);
  sc('G7', '', false);
  msc('H7:K7', 'H7', lic1str, true);

  // ═══════════════════════════════════════════════════════
  // ROW 8: phone | lic2 (no label)
  // Original: R8C2='PHONE NO-+91...' | R8C8=lic2
  // ═══════════════════════════════════════════════════════
  sheet.getRow(8).height = 16;
  sc('B8', `PHONE NO-${data.exporterPhone}`, false);
  sc('C8', '', false);
  sc('D8', '', false);
  sc('E8', '', false);
  sc('F8', '', false);
  sc('G8', '', false);
  msc('H8:K8', 'H8', lic2str, true);

  // ═══════════════════════════════════════════════════════
  // ROW 9: email | Buyer's Order Ref label
  // Original: R9C2='Email ID:...' | R9C6="Buyer's Order Ref.No. :"
  // ═══════════════════════════════════════════════════════
  sheet.getRow(9).height = 16;
  sc('B9', `Email ID: ${data.exporterEmail}`, false);
  sc('C9', '', false);
  sc('D9', '', false);
  sc('E9', '', false);
  sc('F9', "Buyer's Order Ref.No. :", true);
  sc('G9', '', false);
  msc('H9:K9', 'H9', data.buyerOrderRef || '', false);

  // ═══════════════════════════════════════════════════════
  // ROW 10: blank left | Exporter Ref label
  // Original: R10C6='Exporter Ref. and Date :'
  // ═══════════════════════════════════════════════════════
  sheet.getRow(10).height = 16;
  sc('B10', '', false);
  sc('C10', '', false);
  sc('D10', '', false);
  sc('E10', '', false);
  sc('F10', 'Exporter Ref. and Date :', true);
  sc('G10', '', false);
  msc('H10:K10', 'H10', data.exporterRef || '', false);

  // ═══════════════════════════════════════════════════════
  // ROW 11: CONSIGNEE label | BUYER label
  // Original: R11C2='CONSIGNEE  :' | R11C6='BUYER (IF OTHER THAN CONSIGNEE)  :'
  // ═══════════════════════════════════════════════════════
  sheet.getRow(11).height = 16;
  sc('B11', 'CONSIGNEE  :', true);
  sc('C11', '', false);
  sc('D11', '', false);
  sc('E11', '', false);
  msc('F11:K11', 'F11', 'BUYER (IF OTHER THAN CONSIGNEE)  :', true);

  // ═══════════════════════════════════════════════════════
  // ROW 12: consigneeName | buyer name
  // Original: R12C2='TO THE ORDER OF BUYER'
  // ═══════════════════════════════════════════════════════
  sheet.getRow(12).height = 16;
  sc('B12', data.consigneeName, true);
  sc('C12', '', false);
  sc('D12', '', false);
  sc('E12', '', false);
  msc('F12:K12', 'F12', data.buyerName || '', false);

  // ═══════════════════════════════════════════════════════
  // ROW 13: consigneeAddress | blank right
  // Original: R13C2='NAIROBI, KENYA'
  // ═══════════════════════════════════════════════════════
  sheet.getRow(13).height = 16;
  sc('B13', data.consigneeAddress, false);
  sc('C13', '', false);
  sc('D13', '', false);
  sc('E13', '', false);
  msc('F13:K13', 'F13', '', false);

  // ═══════════════════════════════════════════════════════
  // ROWS 14–17: blank filler rows
  // ═══════════════════════════════════════════════════════
  for (let r = 14; r <= 17; r++) {
    sheet.getRow(r).height = 12;
    for (let col = 2; col <= 11; col++) {
      const cell = sheet.getCell(r, col);
      cell.border = BORDER;
      cell.font = FONT_NORMAL;
    }
  }

  // ═══════════════════════════════════════════════════════
  // ROW 18: Logistics HEADER row — labels only
  // Original: R18C2='  PRE-CARRIAGE BY' | R18C4='PLACE OF RECEIPT'
  //           R18C6='COUNTRY OF ORIGIN OF GOODS' | R18C8='INDIA'
  // ═══════════════════════════════════════════════════════
  sheet.getRow(18).height = 20;
  msc('B18:C18', 'B18', '  PRE-CARRIAGE BY',            true, undefined, 'center');
  msc('D18:E18', 'D18', 'PLACE OF RECEIPT',             true, undefined, 'center');
  msc('F18:G18', 'F18', 'COUNTRY OF ORIGIN OF GOODS ', true, undefined, 'center');
  msc('H18:K18', 'H18', 'INDIA ',                       true, undefined, 'center');

  // ROW 19: logistics VALUES row 1
  // Original: R19C2='By AIR' | R19C4='Ahmedabad Airport'
  //           R19C6='COUNTRY OF FINAL DESTINATION' | R19C8='KENYA'
  sheet.getRow(19).height = 20;
  msc('B19:C19', 'B19', data.preCarriage,               false, undefined, 'center');
  msc('D19:E19', 'D19', data.placeOfReceipt,            false, undefined, 'center');
  msc('F19:G19', 'F19', 'COUNTRY OF FINAL DESTINATION', true,  undefined, 'center');
  msc('H19:K19', 'H19', data.finalDestination,          false, undefined, 'center');

  // ROW 20: labels row 2
  // Original: R20C2='  VESSEL/FLIGHT NO.' | R20C4='PORT OF LOADING'
  //           R20C6='TERMS OF DELIVERY' | R20C8=termsValue
  sheet.getRow(20).height = 20;
  msc('B20:C20', 'B20', '  VESSEL/FLIGHT NO.',  true,  undefined, 'center');
  msc('D20:E20', 'D20', 'PORT OF LOADING ',     true,  undefined, 'center');
  msc('F20:G20', 'F20', 'TERMS OF DELIVERY  ',  true,  undefined, 'center');
  msc('H20:K20', 'H20', data.termsOfDelivery,   false, undefined, 'center');

  // ROW 21: values row 2 (vessel/flight, portOfLoading)
  // Original: R21C4='Ahmedabad Airport' (portOfLoading value)
  sheet.getRow(21).height = 20;
  msc('B21:C21', 'B21', data.vesselFlight || '', false, undefined, 'center');
  msc('D21:E21', 'D21', data.portOfLoading,      false, undefined, 'center');
  msc('F21:K21', 'F21', '',                      false, undefined, 'center');

  // ROW 22: labels row 3
  // Original: R22C2='  PORT OF DISCHARGE' | R22C4='FINAL DESTINATION'
  //           R22C6='PAYMENT TERMS' | R22C8=paymentTerms
  sheet.getRow(22).height = 20;
  msc('B22:C22', 'B22', '  PORT OF DISCHARGE', true,  undefined, 'center');
  msc('D22:E22', 'D22', 'FINAL DESTINATION',   true,  undefined, 'center');
  msc('F22:G22', 'F22', 'PAYMENT TERMS ',      true,  undefined, 'center');
  msc('H22:K22', 'H22', data.paymentTerms,     false, undefined, 'center');

  // ROW 23: values row 3
  // Original: R23C2='NAIROBI,KENYA' | R23C4='KENYA'
  sheet.getRow(23).height = 20;
  msc('B23:C23', 'B23', data.portOfDischarge,  false, undefined, 'center');
  msc('D23:E23', 'D23', data.finalDestination, false, undefined, 'center');
  msc('F23:K23', 'F23', '',                    false, undefined, 'center');

  // ═══════════════════════════════════════════════════════
  // ROW 24: TABLE HEADERS
  // Original: R24C2='Marks & Nos.' C3='Description of Goods' C4='HSN CODE' C5='Pack'
  //           C6='Batch No.' C7='Expiry Date' C8='Standard UQC' C9='Quantity (NOS)'
  //           C10='Rate Per Unit / USD' C11='Amount / USD'
  // ═══════════════════════════════════════════════════════
  sheet.getRow(24).height = 35;
  const tableHeads = [
    [2,  'Marks & Nos.'],
    [3,  'Description of Goods'],
    [4,  'HSN CODE'],
    [5,  'Pack'],
    [6,  'Batch No.'],
    [7,  'Expiry Date'],
    [8,  'Standard UQC'],
    [9,  'Quantity (NOS)'],
    [10, 'Rate Per Unit  / USD'],
    [11, 'Amount / USD'],
  ] as [number, string][];
  tableHeads.forEach(([col, label]) => {
    const cell = sheet.getCell(24, col);
    cell.value = label;
    cell.font = FONT_BOLD;
    cell.border = BORDER;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // ═══════════════════════════════════════════════════════
  // ROW 25: blank spacer before items (original has blank R25, items start R26)
  // ═══════════════════════════════════════════════════════
  sheet.getRow(25).height = 8;
  for (let col = 2; col <= 11; col++) {
    sheet.getCell(25, col).border = BORDER;
  }

  let row = 26;
  let totalQty = 0;
  let totalAmount = 0;

  // ═══════════════════════════════════════════════════════
  // PRODUCT ROWS — start at R26, 2 rows per item (item row + state code row)
  //
  // FIX-1: Col C = productName + "\n" + description + "\n"
  //   Original always has two-line cell: name on line 1, description on line 2
  //   e.g. "TOBRINE EYE/EAR DROPS 5ML\nTOBRAMYCIN OPHTHALMIC SOLUTION USP 0.3% W/V\n"
  //
  // FIX-3: Col D HSN written as Number (original stores as integer e.g. 30049099)
  //
  // Original item row:  B=srNo, C=name\ndesc\n, D=HSN(int), E=pack,
  //                     F=batch, G=expiry(MM/YYYY), H=UQC, I=qty, J=rate, K=amount
  // Original state row: B=border only, C:K merged =
  //   "STATE CODE :  24, GSTIN No.: 24AAAFI5671D1ZP                                            DISTRICT CODE :  SURENDRANAGAR"
  // ═══════════════════════════════════════════════════════
  data.items.forEach((item, idx) => {
    sheet.getRow(row).height = 45;

    // Col B: Sr.No
    const bCell = sheet.getCell(row, 2);
    bCell.value = idx + 1;
    bCell.border = BORDER;
    bCell.font = FONT_NORMAL;
    bCell.alignment = { horizontal: 'center', vertical: 'top' };

    // FIX-1: Col C — productName + "\n" + description (matches original two-line cell format)
    const cCell = sheet.getCell(row, 3);
    cCell.value = item.description
      ? `${item.productName}\n${item.description}\n`
      : item.productName;
    cCell.border = BORDER;
    cCell.font = FONT_NORMAL;
    cCell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };

    // FIX-3: Col D — HSN as Number (original: integer 30049099, not string '30049099')
    const dCell = sheet.getCell(row, 4);
    dCell.value = Number(item.hsnSac) || item.hsnSac;
    dCell.border = BORDER;
    dCell.font = FONT_NORMAL;
    dCell.alignment = { horizontal: 'center', vertical: 'top' };

    // Col E: Pack
    const eCell = sheet.getCell(row, 5);
    eCell.value = item.packSize;
    eCell.border = BORDER;
    eCell.font = FONT_NORMAL;
    eCell.alignment = { horizontal: 'center', vertical: 'top' };

    // Col F: Batch No.
    const fCell = sheet.getCell(row, 6);
    fCell.value = item.batchNo;
    fCell.border = BORDER;
    fCell.font = FONT_NORMAL;
    fCell.alignment = { horizontal: 'center', vertical: 'top' };

    // Col G: Expiry Date MM/YYYY
    const gCell = sheet.getCell(row, 7);
    gCell.value = formatExpiry(item.expDate);
    gCell.border = BORDER;
    gCell.font = FONT_NORMAL;
    gCell.alignment = { horizontal: 'center', vertical: 'top' };

    // Col H: Standard UQC — leading space matches original e.g. " 1 KGS"
    const hCell = sheet.getCell(row, 8);
    hCell.value = ` ${item.netWeight} ${item.uom}`;
    hCell.border = BORDER;
    hCell.font = FONT_NORMAL;
    hCell.alignment = { horizontal: 'center', vertical: 'top' };

    // Col I: Quantity
    const q = Number(item.quantity) || 0;
    const iCell = sheet.getCell(row, 9);
    iCell.value = q;
    iCell.border = BORDER;
    iCell.font = FONT_NORMAL;
    iCell.alignment = { horizontal: 'center', vertical: 'top' };
    totalQty += q;

    // Col J: Rate Per Unit
    const p = Number(item.price) || 0;
    const jCell = sheet.getCell(row, 10);
    jCell.value = p;
    jCell.border = BORDER;
    jCell.font = FONT_NORMAL;
    jCell.numFmt = '"$"#,##0.00';
    jCell.alignment = { horizontal: 'right', vertical: 'top' };

    // Col K: Amount
    const amt = q * p;
    const kCell = sheet.getCell(row, 11);
    kCell.value = amt;
    kCell.border = BORDER;
    kCell.font = FONT_NORMAL;
    kCell.numFmt = '"$"#,##0.00';
    kCell.alignment = { horizontal: 'right', vertical: 'top' };
    totalAmount += amt;

    row++;

    // STATE CODE sub-row — col B border only, C:K merged
    sheet.getRow(row).height = 15;
    const bBorder = sheet.getCell(row, 2);
    bBorder.border = BORDER;
    bBorder.font = FONT_NORMAL;

    sheet.mergeCells(`C${row}:K${row}`);
    const scCell = sheet.getCell(row, 3);
    scCell.value = `STATE CODE :  ${item.stateCode}, GSTIN No.: ${item.supplierGstin}                                            DISTRICT CODE :  ${item.distCode}`;
    scCell.border = BORDER;
    scCell.font = { size: 8, name: 'Arial' };
    scCell.alignment = { vertical: 'middle', horizontal: 'left' };

    row++;
  });

  // ═══════════════════════════════════════════════════════
  // PAD blank rows so totals always land at R63 (matches original)
  // 13 items × 2 rows each = R26–R51 (row=52 after loop)
  // Target totals row = 63  →  pad R52–R62 = 11 rows
  // For any item count: pad until row reaches 63
  // ═══════════════════════════════════════════════════════
  const totalsTargetRow = 63;
  while (row < totalsTargetRow) {
    sheet.getRow(row).height = 12;
    for (let col = 2; col <= 11; col++) {
      const cell = sheet.getCell(row, col);
      cell.border = BORDER;
      cell.font = FONT_NORMAL;
    }
    row++;
  }

  // ═══════════════════════════════════════════════════════
  // TOTALS ROW (R63)
  // Original: R63C2=' No. of Corrugated Boxes :   06'
  //           R63C9=totalQty | R63C10='CIF VALUE' | R63C11=cifTotal
  // ═══════════════════════════════════════════════════════
  sheet.getRow(row).height = 20;

  sheet.mergeCells(`B${row}:H${row}`);
  const boxInfoCell = sheet.getCell(row, 2);
  boxInfoCell.value = ` No. of  Corrugated Boxes :   ${data.totalCorrugatedBoxes}`;
  boxInfoCell.border = BORDER;
  boxInfoCell.font = FONT_BOLD;
  boxInfoCell.alignment = { vertical: 'middle', horizontal: 'left' };

  const qtyCell = sheet.getCell(row, 9);
  qtyCell.value = totalQty;
  qtyCell.border = BORDER;
  qtyCell.font = FONT_BOLD;
  qtyCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const cifLabel = sheet.getCell(row, 10);
  cifLabel.value = 'CIF VALUE';
  cifLabel.border = BORDER;
  cifLabel.font = FONT_BOLD;
  cifLabel.alignment = { horizontal: 'left', vertical: 'middle' };

  const cifAmt = sheet.getCell(row, 11);
  cifAmt.value = totalAmount;
  cifAmt.border = BORDER;
  cifAmt.font = FONT_BOLD;
  cifAmt.numFmt = '"$"#,##0.00';
  cifAmt.alignment = { horizontal: 'right', vertical: 'middle' };
  row++;

  // ═══════════════════════════════════════════════════════
  // SUMMARY ROWS (R64, R65, R66)
  // Original: R64C2=' Gross Weight :  59 KGS'  | R64C10='FREIGHT VALUE' | R64C11=393
  //           R65C2=' Nett Weight :  50 KGS'   | R65C10='INSURANCE'     | R65C11=7
  //           R66C2='AMOUNT CHARGEABLE  :'      | R66C10='FOB VALUE'     | R66C11=6337
  //
  // FIX-2: Use formatWeight() so whole-number weights show as "59" not "59.000"
  // ═══════════════════════════════════════════════════════
  const fr  = Number(data.freightValue)   || 0;
  const ins = Number(data.insuranceValue) || 0;
  const fob = totalAmount - fr - ins;

  const summaryRow = (leftLabel: string, rightLabel: string, rightValue: number) => {
    sheet.getRow(row).height = 18;
    sheet.mergeCells(`B${row}:I${row}`);
    const lc = sheet.getCell(row, 2);
    lc.value = leftLabel;
    lc.border = BORDER;
    lc.font = FONT_BOLD;
    lc.alignment = { vertical: 'middle', horizontal: 'left' };

    const rl = sheet.getCell(row, 10);
    rl.value = rightLabel;
    rl.border = BORDER;
    rl.font = FONT_BOLD;
    rl.alignment = { horizontal: 'left', vertical: 'middle' };

    const rv = sheet.getCell(row, 11);
    rv.value = rightValue;
    rv.border = BORDER;
    rv.font = FONT_BOLD;
    rv.numFmt = '"$"#,##0.00';
    rv.alignment = { horizontal: 'right', vertical: 'middle' };
    row++;
  };

  // FIX-2: formatWeight() strips ".000" → "59" not "59.000"
  summaryRow(` Gross Weight :  ${formatWeight(data.totalGrossWeight)} KGS`, 'FREIGHT VALUE', fr);
  summaryRow(` Nett Weight :  ${formatWeight(data.totalNetWeight)} KGS`,    'INSURANCE',     ins);
  summaryRow('AMOUNT CHARGEABLE  :',                                        'FOB VALUE',     fob);

  // ═══════════════════════════════════════════════════════
  // AMOUNT IN WORDS — B:K merged, 2 rows
  // ═══════════════════════════════════════════════════════
  sheet.getRow(row).height = 20;
  sheet.mergeCells(`B${row}:K${row + 1}`);
  const wordsCell = sheet.getCell(row, 2);
  wordsCell.value = numberToWords(totalAmount);
  wordsCell.border = BORDER;
  wordsCell.font = FONT_BOLD;
  wordsCell.alignment = { vertical: 'top', wrapText: true };
  row += 2;

  // ═══════════════════════════════════════════════════════
  // COMPLIANCE & SIGNATURE — B:G merged (legal) | H:K merged (signature)
  // ═══════════════════════════════════════════════════════
  const isPaid = (data.gstStatus || '').toUpperCase().includes('PAID');
  const igstLine = isPaid
    ? '* SUPPLY MEANT FOR EXPORT UNDER WITH PAYMENT OF INTEGRATED TAX (IGST)'
    : '* SUPPLY MEANT FOR EXPORT UNDER LETTER OF UNDERTAKING WITHOUT PAYMENT OF IGST';

  const legal = [
    '(WE UNDERTAKE TO ABIDE BY PROVISIONS OF FOREIGN EXCHANGE MANAGEMENT ACT,1999, AS  ',
    'AMENDED FROM TIME TO TIME, INCLUDING REALIZATION/REPATRIATION OF FOREIGN EXCHANGE TO OR FROM INDIA)',
    igstLine,
    '* * No DBK or RoDTEP for All Items ( ITS FREE SHIPPING BILL)& GNX 100',
    'Declaration:',
    'We declare that this invoice shows actual price of the goods described and that all particulars are true and correct.',
  ].join('\n');

  sheet.mergeCells(`B${row}:G${row + 5}`);
  const legalCell = sheet.getCell(row, 2);
  legalCell.value = legal;
  legalCell.border = BORDER;
  legalCell.font = FONT_SMALL;
  legalCell.alignment = { wrapText: true, vertical: 'top' };

  sheet.mergeCells(`H${row}:K${row + 5}`);
  const sigCell = sheet.getCell(row, 8);
  sigCell.value = `FOR ${data.exporterName.toUpperCase()},\n\n\n\nAUTHORISED SIGNATORY`;
  sigCell.border = BORDER;
  sigCell.font = FONT_BOLD;
  sigCell.alignment = { horizontal: 'center', vertical: 'bottom', wrapText: true };
};

export const generateCommercialInvoice = async (data: MasterData) => {
  const workbook = new ExcelJS.Workbook();
  addCommercialInvoiceSheet(workbook, data);
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), `Invoice_Commercial_${data.invoiceNo}.xlsx`);
};