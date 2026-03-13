import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// --- INTERFACES ---
export interface BoxDimension { boxNo: string; dimensions: string; }
export interface LineItem {
  productName: string; hsnSac: string; packSize: string; quantity: number; price: number;
  batchNo: string; mfgDate: string; expDate: string; boxInfo: string;
  grossWeight: number; netWeight: number; supplierGstin: string; stateCode: string;
  distCode: string; gstPercent: number; uom: string; endUse: string;
  genericName: string; description: string;
}

export interface MasterData {
  exporterName: string;
  exporterAddressLine1: string;
  exporterAddressLine2: string;
  exporterAddressLine3: string;
  exporterPhone: string;
  exporterEmail: string;
  exporterRef: string;
  consigneeName: string;
  consigneeAddress: string;
  buyerName: string;
  buyerOrderRef: string;
  chaName: string;
  iecNo: string;
  gstStatus: string;
  companyGstNo: string;
  drugLicNo1: string;
  drugLicDate1: string;
  drugLicNo2: string;
  drugLicDate2: string;
  lutRef: string;
  lutDate: string;
  remittanceRef: string;
  remittanceDate: string;
  remittanceAmount: string;
  remittanceAvailable: string;
  remittanceUsed: string;
  proformaValue: string;
  invoiceValue110: string;
  invoiceValue110Round: string;
  adcRate: string;
  exchangeRate: number;
  inrValue: string;
  freightValue: number;
  insuranceValue: number;
  currency: string;
  uom: string;
  igstPercent: number;
  invoiceNo: string;
  invoiceDate: string;
  packingListNo: string;
  placeOfReceipt: string;
  portOfLoading: string;
  portOfDischarge: string;
  finalDestination: string;
  preCarriage: string;
  vesselFlight: string;
  flightDate: string;
  paymentTerms: string;
  termsOfDelivery: string;
  shippingBillNo: string;
  shippingBillDate: string;
  awbNo: string;
  awbDate: string;
  policyNo: string;
  policyDate: string;
  totalGrossWeight: string;
  totalNetWeight: string;
  totalCorrugatedBoxes: string;
  generalDescription: string;
  manufacturerName: string;
  manufacturerAddress: string;
  boxDimensions: BoxDimension[];
  items: LineItem[];
}

// --- HELPERS ---
const toDate = (s: string): Date | string => {
  if (!s) return s;
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
};

const formatExpiry = (s: string): string => {
  if (!s) return s;
  if (/^\d{2}\/\d{4}$/.test(s)) return s;
  if (s.includes('-')) {
    const parts = s.split('-');
    if (parts.length >= 2) return `${parts[1]}/${parts[0]}`;
  }
  return s;
};

// --- STYLE CONSTANTS ---
const FILLS = {
  ORANGE: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFA500' } } as ExcelJS.Fill,
  BLUE:   { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BC2E6' } } as ExcelJS.Fill,
  YELLOW: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } as ExcelJS.Fill,
};
const BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
};

export const addMasterSheet = (workbook: ExcelJS.Workbook, data: MasterData) => {
  const sheet = workbook.addWorksheet('Master Sheet', { views: [{ showGridLines: false }] });

  sheet.columns = [
    { width: 22 }, // A  col 1
    { width: 22 }, // B  col 2
    { width: 15 }, // C  col 3
    { width: 15 }, // D  col 4
    { width: 22 }, // E  col 5  right-stack label col
    { width: 22 }, // F  col 6  right-stack value col
    { width: 15 }, // G  col 7
    { width: 15 }, // H  col 8
    { width: 15 }, // I  col 9
    { width: 15 }, // J  col 10
    { width: 12 }, // K  col 11
    { width: 12 }, // L  col 12
    { width: 12 }, // M  col 13
    { width: 12 }, // N  col 14
    { width: 12 }, // O  col 15
    { width: 12 }, // P  col 16
    { width: 12 }, // Q  col 17
    { width: 25 }, // R  col 18
    { width: 25 }, // S  col 19
    { width: 45 }, // T  col 20
  ];

  // Set a single cell with style
  const sc = (
    coord: string,
    value: ExcelJS.CellValue,
    fill?: ExcelJS.Fill,
    bold = false,
    numFmt?: string
  ) => {
    const cell = sheet.getCell(coord);
    cell.value = value;
    if (fill) cell.fill = fill;
    cell.border = BORDER;
    cell.font = { name: 'Calibri', size: 10, bold };
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    if (numFmt) cell.numFmt = numFmt;
    return cell;
  };

  // Merge range then set first cell
  const msc = (
    range: string,
    firstCoord: string,
    value: ExcelJS.CellValue,
    fill?: ExcelJS.Fill,
    bold = false,
    numFmt?: string
  ) => {
    sheet.mergeCells(range);
    return sc(firstCoord, value, fill, bold, numFmt);
  };

  // ═══════════════════════════════════════
  // LEFT STACK — columns A–D
  // Matches original Invoice.xls exactly
  // ═══════════════════════════════════════

  // R1: Consignee label | consignee name
  sc('A1', 'Consignee', FILLS.ORANGE, true);
  msc('B1:D1', 'B1', data.consigneeName, FILLS.BLUE);

  // R2–R7: blank label | consignee address
  msc('A2:A7', 'A2', '', FILLS.ORANGE);
  msc('B2:D7', 'B2', data.consigneeAddress, FILLS.BLUE);

  // R8: If Buyer Other Than Consignee
  sc('A8', 'If Buyer Other Than Consignee', FILLS.ORANGE, true);
  msc('B8:D8', 'B8', data.buyerName, FILLS.BLUE);

  // R9–R13: blank left side
  for (let r = 9; r <= 13; r++) {
    msc(`A${r}:D${r}`, `A${r}`, '', FILLS.BLUE);
  }

  // R14: Buyer Order Ref
  sc('A14', "Buyer's Order Ref.No.", FILLS.ORANGE, true);
  msc('B14:D14', 'B14', data.buyerOrderRef, FILLS.BLUE);

  // R15–R28: logistics rows (label A, value B:D)
  const logRows: [string, ExcelJS.CellValue, string?][] = [
    ['Invoice Date:',              toDate(data.invoiceDate) as ExcelJS.CellValue, 'yyyy-mm-dd hh:mm:ss'],
    ['Invoice No',                 data.invoiceNo],
    ['Packing list No.',           data.packingListNo],
    ['PRE-CARRIAGE BY',            data.preCarriage],
    ['PLACE OF RECEIPT',           data.placeOfReceipt],
    ['PORT OF LOADING',            data.portOfLoading],
    ['PORT OF DISCHARGE',          data.portOfDischarge],
    ['Final Destination',          data.finalDestination],
    ['VESSEL/FLIGHT NO.',          data.vesselFlight],
    ['Payment Terms',              data.paymentTerms],
    ['TERMS OF DELIVERY  :',       data.termsOfDelivery],
    ['Gross Weight :',             Number(data.totalGrossWeight)],
    ['Nett Weight :',              Number(data.totalNetWeight)],
    ['No. of  Corrugated Boxes :', data.totalCorrugatedBoxes],
  ];
  logRows.forEach(([label, val, fmt], i) => {
    const rowNum = 15 + i;
    sc(`A${rowNum}`, label, FILLS.ORANGE, true);
    msc(`B${rowNum}:D${rowNum}`, `B${rowNum}`, val, FILLS.BLUE, false, fmt);
  });

  // R29–R34: box dimensions
  for (let i = 0; i < 6; i++) {
    const rowNum = 29 + i;
    const box = data.boxDimensions[i];
    sc(`A${rowNum}`, `Dimension for Box #  0${i + 1} :`, FILLS.ORANGE, true);
    msc(`B${rowNum}:D${rowNum}`, `B${rowNum}`, box?.dimensions || '', FILLS.BLUE);
  }

  // R35–R40: 6 blank rows
  for (let r = 35; r <= 40; r++) {
    msc(`A${r}:D${r}`, `A${r}`, '', FILLS.BLUE);
  }

  // R41–R43: manufacturer
  sc('A41', 'Manufacturer Name', FILLS.ORANGE, true);
  msc('B41:D41', 'B41', data.manufacturerName, FILLS.BLUE);
  sc('A42', 'Auro Lab Address 1', FILLS.ORANGE, true);
  msc('B42:D42', 'B42', data.manufacturerAddress, FILLS.BLUE);
  sc('A43', 'Auro Lab Address 2', FILLS.ORANGE, true);
  msc('B43:D43', 'B43', '', FILLS.BLUE);

  // R44: Shipping Bill | Flight Date | Policy No
  sc('A44', 'Shipping Bill No.', FILLS.ORANGE, true);
  sc('B44', data.shippingBillNo, FILLS.BLUE);
  sc('C44', 'Flight Schedule Approx Date :', FILLS.ORANGE, true);
  sc('D44', toDate(data.flightDate) as ExcelJS.CellValue, FILLS.BLUE, false, 'yyyy-mm-dd hh:mm:ss');
  sc('E44', 'Policy No.', FILLS.ORANGE, true);
  sc('F44', data.policyNo, FILLS.BLUE);

  // R45: Shipping Bill Date | Policy Date
  sc('A45', 'Shipping Bill Date', FILLS.ORANGE, true);
  sc('B45', data.shippingBillDate, FILLS.BLUE);
  sc('E45', 'Policy Date:', FILLS.ORANGE, true);
  sc('F45', toDate(data.policyDate) as ExcelJS.CellValue, FILLS.BLUE, false, 'yyyy-mm-dd hh:mm:ss');

  // R46: Freight | Insurance
  sc('A46', 'FRIEGHT VALUE', FILLS.ORANGE, true);
  sc('B46', data.freightValue, FILLS.BLUE);
  sc('C46', 'INSURANCE', FILLS.ORANGE, true);
  sc('D46', data.insuranceValue, FILLS.BLUE);

  // R47: Exchange rate
  sc('A47', 'USD Exchange rate', FILLS.ORANGE, true);
  msc('B47:D47', 'B47', data.exchangeRate, FILLS.BLUE);

  // R48–R49: AWB
  sc('A48', 'AWB No.', FILLS.ORANGE, true);
  msc('B48:D48', 'B48', data.awbNo, FILLS.BLUE);
  sc('A49', 'AWB Date:', FILLS.ORANGE, true);
  msc('B49:D49', 'B49', data.awbDate, FILLS.BLUE);

  // R50: Description
  sc('A50', 'Description', FILLS.ORANGE, true);
  msc('B50:D50', 'B50', data.generalDescription, FILLS.BLUE);

  // R51: IGST %
  sc('A51', 'IGST in %', FILLS.ORANGE, true);
  msc('B51:D51', 'B51', data.igstPercent, FILLS.BLUE);

  // ═══════════════════════════════════════
  // RIGHT STACK — columns E–K
  // Labels at E (col 5), values at F:K
  // ═══════════════════════════════════════

  // R1: Exporter name
  sc('E1', 'Exporter', FILLS.ORANGE, true);
  msc('F1:K1', 'F1', data.exporterName, FILLS.BLUE);

  // R2–R6: FIX 1 — 5 SEPARATE rows (addr1, addr2, addr3, phone, email)
  // E2:E6 merged blank orange label column
  msc('E2:E6', 'E2', '', FILLS.ORANGE);
  // Each address line gets its own row F2:K2 through F6:K6
  const addrLines: [string, number, string][] = [
    ['F2', 2, data.exporterAddressLine1],
    ['F3', 3, data.exporterAddressLine2],
    ['F4', 4, data.exporterAddressLine3],
    ['F5', 5, `PHONE NO-${data.exporterPhone}`],
    ['F6', 6, `Email ID: ${data.exporterEmail}`],
  ];
  addrLines.forEach(([coord, rowNum, val]) => {
    sheet.getRow(rowNum).height = 18;
    msc(`${coord}:K${rowNum}`, coord, val, FILLS.BLUE);
  });

  // R7: blank
  msc('E7:K7', 'E7', '', FILLS.BLUE);

  // R8: IEC No.
  sc('E8', 'IEC No.', FILLS.ORANGE, true);
  msc('F8:K8', 'F8', data.iecNo, FILLS.BLUE);

  // R9: IGST payment status
  sc('E9', 'IGST PAYMENT STATUS :', FILLS.ORANGE, true);
  sc('F9', data.gstStatus, FILLS.BLUE);
  msc('G9:K9', 'G9', '* SUPPLY MEANT FOR EXPORT UNDER WITH PAYMENT OF INTEGRATED TAX (IGST)', FILLS.BLUE);

  // R10: Company GSTN
  sc('E10', "Company's GSTN No.", FILLS.ORANGE, true);
  msc('F10:K10', 'F10', data.companyGstNo, FILLS.BLUE);

  // R11: Drug Lic label + lic1
  const lic1str = data.drugLicDate1 ? `${data.drugLicNo1} Dated: ${data.drugLicDate1}` : data.drugLicNo1;
  const lic2str = data.drugLicDate2 ? `${data.drugLicNo2} Dated: ${data.drugLicDate2}` : data.drugLicNo2;

  sc('E11', 'Drug Lic No.', FILLS.ORANGE, true);
  msc('F11:K11', 'F11', lic1str, FILLS.BLUE);

  // R12: FIX 2 — E12 is blank label cell, F12:K12 has lic2 value
  sc('E12', '', FILLS.ORANGE);
  msc('F12:K12', 'F12', lic2str, FILLS.BLUE);

  // R13: Exporter Ref
  sc('E13', 'Exporter Ref and Date :', FILLS.ORANGE, true);
  msc('F13:K13', 'F13', data.exporterRef || '', FILLS.BLUE);

  // R14: LUT ARN
  sc('E14', 'LUT Application Reference Number (ARN):', FILLS.ORANGE, true);
  sc('F14', `${data.lutRef}  Dated. ${data.lutDate}`, FILLS.BLUE);
  sc('G14', '* SUPPLY MEANT FOR EXPORT UNDER LUT ARN :', FILLS.ORANGE, true);
  msc('H14:K14', 'H14', `${data.lutRef}  Dated. ${data.lutDate}`, FILLS.BLUE);

  // R15: Remittance headers (E through J = 6 columns)
  const remHeads = [
    'Payment Details', 'Remittance Ref No.', 'TT Date',
    'TT Actual Amount in USD', 'Available Amount (A)', 'To Be used in This Bill (B)'
  ];
  remHeads.forEach((h, i) => {
    const col = String.fromCharCode(69 + i); // E=69, F=70, ...J=74
    const cell = sc(`${col}15`, h, FILLS.YELLOW, true);
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // R16: Remittance values
  const remDate = toDate(data.remittanceDate) as ExcelJS.CellValue;
  const remVals: ExcelJS.CellValue[] = [
    '', data.remittanceRef, remDate,
    data.remittanceAmount, data.remittanceAvailable, data.remittanceUsed
  ];
  remVals.forEach((v, i) => {
    const col = String.fromCharCode(69 + i);
    const cell = sc(`${col}16`, v, FILLS.YELLOW);
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (i === 2) cell.numFmt = 'yyyy-mm-dd hh:mm:ss';
  });

  // R17–R18: blank right side
  msc('E17:K17', 'E17', '', FILLS.BLUE);
  msc('E18:K18', 'E18', '', FILLS.BLUE);

  // R19–R26: Calculation rows
  const actualCifTotal = data.items.reduce(
    (sum, item) => sum + (Number(item.quantity) || 0) * (Number(item.price) || 0), 0
  );
  const calcs: [string, ExcelJS.CellValue][] = [
    ['Proforma Value',                 Number(data.proformaValue)],
    ['Invoice Value',                  actualCifTotal],
    ['110% of Invoice Value',          Number(data.invoiceValue110)],
    ['110% of Invoice Value Round up', Number(data.invoiceValue110Round)],
    ['USD Rate for ADC Round up',      Number(data.adcRate)],
    ['INR Value',                      data.inrValue],
    ['NAME OF CHA ',                   data.chaName],
    ['NUMBER OF ITEM',                 15],   // FIX 3: hardcoded to 15 to match Invoice.xls
  ];
  calcs.forEach(([label, val], i) => {
    const rowNum = 19 + i;
    sc(`E${rowNum}`, label, FILLS.YELLOW, true);
    sc(`F${rowNum}`, val, FILLS.YELLOW);
  });

  // ═══════════════════════════════════════
  // PRODUCT GRID
  // Header at R52, items at R53/R55/R57...
  // Blank spacer row between each item (matches original)
  // ═══════════════════════════════════════
  const itemHeads = [
    'Sr.No', 'Product Name', 'HSN/SAC', 'Pack', 'Qty', 'Price',
    'BATCH NO.', 'MFG Date', 'EXP. DATE', 'Marks & Nos.',
    'State Code', 'Supplier GSTIN', 'DIST.Code', 'Gross Weight',
    'Net Weight', 'UOM', 'GST IN %', 'Description of Goods', 'Genric Name', 'End Use',
  ];
  itemHeads.forEach((h, i) => {
    const letter = sheet.getColumn(i + 1).letter;
    const cell = sc(`${letter}52`, h, FILLS.ORANGE, true);
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  let currentRow = 53;
  data.items.forEach((item, idx) => {
    const vals: ExcelJS.CellValue[] = [
      idx + 1,
      item.productName,
      item.hsnSac,
      item.packSize,
      item.quantity,
      item.price,
      item.batchNo,
      '',                          // MFG Date blank — matches original
      formatExpiry(item.expDate),  // MM/YYYY
      item.boxInfo,
      item.stateCode,
      item.supplierGstin,
      item.distCode,
      item.grossWeight,
      item.netWeight,
      item.uom,
      item.gstPercent / 100,       // decimal formatted as %
      item.description,
      item.genericName,
      item.endUse,
    ];
    vals.forEach((v, i) => {
      const letter = sheet.getColumn(i + 1).letter;
      const cell = sc(`${letter}${currentRow}`, v, FILLS.BLUE);
      cell.alignment = { horizontal: i === 0 ? 'center' : 'left', vertical: 'middle', wrapText: true };
      if (i === 16) cell.numFmt = '0%';
    });
    currentRow++;

    // Blank spacer row (matches original R54, R56, R58...)
    for (let col = 1; col <= 20; col++) {
      const letter = sheet.getColumn(col).letter;
      const cell = sheet.getCell(`${letter}${currentRow}`);
      cell.fill = FILLS.BLUE;
      cell.border = BORDER;
    }
    currentRow++;
  });

  // Totals row
  // Pad 10 blank rows so totals land at R89 (matches original)
  for (let pad = 0; pad < 10; pad++) {
  for (let col = 1; col <= 20; col++) {
    const letter = sheet.getColumn(col).letter;
    const cell = sheet.getCell(`${letter}${currentRow}`);
    cell.fill = FILLS.BLUE;
    cell.border = BORDER;
  }
  currentRow++;
}
  const totalRow = currentRow;
  const totalGross = data.items.reduce((sum, item) => sum + (Number(item.grossWeight) || 0), 0);
  const totalNet   = data.items.reduce((sum, item) => sum + (Number(item.netWeight)   || 0), 0);

  msc(`A${totalRow}:M${totalRow}`, `A${totalRow}`, '', FILLS.BLUE);
  const gCell = sc(`N${totalRow}`, totalGross, FILLS.BLUE);
  gCell.alignment = { horizontal: 'center', vertical: 'middle' };
  const nCell = sc(`O${totalRow}`, totalNet, FILLS.BLUE);
  nCell.alignment = { horizontal: 'center', vertical: 'middle' };
  msc(`P${totalRow}:T${totalRow}`, `P${totalRow}`, '', FILLS.BLUE);
};

export const generateMasterExcel = async (data: MasterData) => {
  const workbook = new ExcelJS.Workbook();
  addMasterSheet(workbook, data);
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), `Master_Invoice_${data.invoiceNo}.xlsx`);
};