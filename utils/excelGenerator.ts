import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// --- 1. INTERFACES (Matches app/page.tsx 100%) ---

export interface BoxDimension {
  boxNo: string;
  dimensions: string;
}

export interface LineItem {
  productName: string;
  hsnSac: string;
  packSize: string;
  quantity: number;
  price: number;
  batchNo: string;
  mfgDate: string;
  expDate: string;
  boxInfo: string;
  grossWeight: number;
  netWeight: number;
  supplierGstin: string;
  stateCode: string;
  distCode: string;
  gstPercent: number;
  uom: string;
  endUse: string;
  genericName: string;
  description: string;
}

export interface MasterData {
  // Parties
  exporterName: string;
  exporterAddress: string;
  exporterPhone: string;
  exporterEmail: string;
  exporterRef: string;
  consigneeName: string;
  consigneeAddress: string;
  buyerName: string;
  buyerOrderRef: string;
  chaName: string;

  // Regulatory
  iecNo: string;
  gstStatus: string;
  companyGstNo: string;
  drugLicNo: string;
  lutRef: string;
  lutDate: string;
  
  // Remittance
  remittanceRef: string;
  remittanceDate: string;
  remittanceAmount: string;
  remittanceAvailable: string;
  remittanceUsed: string;

  // Financials
  proformaValue: string;
  invoiceValue110: string;
  invoiceValue110Round: string; 
  adcRate: string;
  inrValue: string;
  freightValue: number;
  insuranceValue: number;
  exchangeRate: number;
  currency: string;
  uom: string;

  // Logistics
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
  
  // Shipping Docs
  shippingBillNo: string;
  shippingBillDate: string;
  awbNo: string;
  awbDate: string;
  policyNo: string;
  policyDate: string;

  // Packing Summary
  totalGrossWeight: string;
  totalNetWeight: string;
  totalCorrugatedBoxes: string;
  generalDescription: string;
  globalIgst: string;

  // Manufacturer
  manufacturerName: string;
  manufacturerAddress: string;

  // Arrays
  boxDimensions: BoxDimension[];
  items: LineItem[];
}

// --- 2. STYLING CONSTANTS ---

const BORDER_ALL: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

const FILLS = {
  BLUE: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99CCFF' } } as ExcelJS.Fill,
  ORANGE: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } } as ExcelJS.Fill,
  YELLOW: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } } as ExcelJS.Fill,
};

const FONTS = {
  HEADER: { bold: true, size: 11, name: 'Calibri' },
  BOLD: { bold: true, size: 10, name: 'Calibri' }
};

// --- 3. GENERATOR LOGIC ---

export const generateMasterExcel = async (data: MasterData) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Master Sheet', {
    views: [{ showGridLines: false }]
  });

  // Setup ample columns to fit all data
  sheet.columns = [
    { key: 'A', width: 25 }, { key: 'B', width: 25 }, { key: 'C', width: 10 }, 
    { key: 'D', width: 20 }, { key: 'E', width: 15 }, { key: 'F', width: 15 }, { key: 'G', width: 15 },
    { key: 'H', width: 15 }, { key: 'I', width: 15 }, { key: 'J', width: 15 }, { key: 'K', width: 15 },
    { key: 'L', width: 15 }, { key: 'M', width: 15 }, { key: 'N', width: 15 }, { key: 'O', width: 15 }
  ];

  let row = 1;

  // HEADER BLOCK
  sheet.mergeCells(`A${row}:C${row}`);
  sheet.getCell(`A${row}`).value = "CONSIGNEE";
  sheet.getCell(`A${row}`).fill = FILLS.ORANGE;
  sheet.getCell(`A${row}`).font = FONTS.HEADER;
  sheet.getCell(`A${row}`).border = BORDER_ALL;

  sheet.mergeCells(`D${row}:G${row}`);
  sheet.getCell(`D${row}`).value = "EXPORTER";
  sheet.getCell(`D${row}`).fill = FILLS.ORANGE;
  sheet.getCell(`D${row}`).font = FONTS.HEADER;
  sheet.getCell(`D${row}`).border = BORDER_ALL;
  
  row++;

  // ADDRESS BLOCK
  sheet.mergeCells(`A${row}:C${row + 3}`);
  sheet.getCell(`A${row}`).value = `${data.consigneeName}\n${data.consigneeAddress}\n\n${data.buyerName ? 'BUYER: ' + data.buyerName : ''}`;
  sheet.getCell(`A${row}`).alignment = { wrapText: true, vertical: 'top' };
  sheet.getCell(`A${row}`).border = BORDER_ALL;

  sheet.mergeCells(`D${row}:G${row + 3}`);
  sheet.getCell(`D${row}`).value = `${data.exporterName}\n${data.exporterAddress}\nPhone: ${data.exporterPhone}\nEmail: ${data.exporterEmail}`;
  sheet.getCell(`D${row}`).alignment = { wrapText: true, vertical: 'top' };
  sheet.getCell(`D${row}`).border = BORDER_ALL;

  row += 4;

  // HELPER FUNCTIONS
  const drawRow = (r: number, label: string, val: string, color: ExcelJS.Fill) => {
    sheet.getCell(`A${r}`).value = label;
    sheet.getCell(`A${r}`).fill = color;
    sheet.getCell(`A${r}`).font = FONTS.BOLD;
    sheet.getCell(`A${r}`).border = BORDER_ALL;
    sheet.mergeCells(`B${r}:C${r}`);
    sheet.getCell(`B${r}`).value = val;
    sheet.getCell(`B${r}`).border = BORDER_ALL;
  };

  const drawRegRow = (r: number, label: string, val: string) => {
    sheet.getCell(`D${r}`).value = label;
    sheet.getCell(`D${r}`).fill = FILLS.ORANGE;
    sheet.getCell(`D${r}`).font = FONTS.BOLD;
    sheet.getCell(`D${r}`).border = BORDER_ALL;
    sheet.mergeCells(`E${r}:G${r}`);
    sheet.getCell(`E${r}`).value = val;
    sheet.getCell(`E${r}`).border = BORDER_ALL;
  };

  const startLogisticsRow = row;

  // --- LOGISTICS (Left Side) ---
  drawRow(row, "Invoice No", data.invoiceNo, FILLS.BLUE); row++;
  drawRow(row, "Date", data.invoiceDate, FILLS.BLUE); row++;
  drawRow(row, "Packing List No", data.packingListNo, FILLS.ORANGE); row++;
  drawRow(row, "Port of Loading", data.portOfLoading, FILLS.ORANGE); row++;
  drawRow(row, "Port of Discharge", data.portOfDischarge, FILLS.ORANGE); row++;
  drawRow(row, "Final Destination", data.finalDestination, FILLS.ORANGE); row++;
  drawRow(row, "Payment Terms", data.paymentTerms, FILLS.ORANGE); row++;
  
  // --- FINANCIALS (Yellow Block) ---
  drawRow(row, "Proforma Value", data.proformaValue, FILLS.YELLOW); row++;
  drawRow(row, "110% Value", data.invoiceValue110, FILLS.YELLOW); row++;
  drawRow(row, "110% Round Up", data.invoiceValue110Round, FILLS.YELLOW); row++; 
  drawRow(row, "ADC Rate", data.adcRate, FILLS.YELLOW); row++;
  drawRow(row, "INR Value", data.inrValue, FILLS.YELLOW); 

  // --- REGULATORY (Right Side) ---
  let regRow = startLogisticsRow;
  drawRegRow(regRow, "IEC No.", data.iecNo); regRow++;
  drawRegRow(regRow, "GST Status", data.gstStatus); regRow++;
  drawRegRow(regRow, "Exporter Ref", data.exporterRef); regRow++;
  drawRegRow(regRow, "LUT Ref", data.lutRef); regRow++;
  drawRegRow(regRow, "LUT Date", data.lutDate); regRow++;
  
  // Remittance
  drawRegRow(regRow, "Remittance Ref", data.remittanceRef); regRow++;
  drawRegRow(regRow, "TT Date", data.remittanceDate); regRow++;
  drawRegRow(regRow, "TT Amount", data.remittanceAmount); regRow++;
  drawRegRow(regRow, "Available Amt", data.remittanceAvailable); regRow++;
  drawRegRow(regRow, "Amount Used", data.remittanceUsed); 

  row = Math.max(row, regRow) + 2;

  // --- BOX DIMENSIONS ---
  sheet.getCell(`A${row}`).value = "PACKING DIMENSIONS";
  sheet.getCell(`A${row}`).fill = FILLS.ORANGE;
  sheet.getCell(`A${row}`).font = FONTS.HEADER;
  sheet.getCell(`A${row}`).border = BORDER_ALL;
  row++;

  data.boxDimensions.forEach((box) => {
    sheet.getCell(`A${row}`).value = box.boxNo;
    sheet.getCell(`A${row}`).fill = FILLS.ORANGE;
    sheet.getCell(`A${row}`).border = BORDER_ALL;
    sheet.mergeCells(`B${row}:C${row}`);
    sheet.getCell(`B${row}`).value = box.dimensions;
    sheet.getCell(`B${row}`).fill = FILLS.BLUE;
    sheet.getCell(`B${row}`).border = BORDER_ALL;
    row++;
  });
  row++;

  // --- ITEMS HEADER (100% MATCH) ---
  const headers = [
    'Sr.', 'Product Name', 'HSN', 'Pack', 'Qty', 'Price', 'Batch', 'Mfg Date', 'Exp Date', 
    'Box Info', 'State', 'Supp GST', 'Dist', 'Gr Wt', 'Net Wt', 'UOM', 'GST %', 'End Use'
  ];
  
  headers.forEach((h, i) => {
    const cell = sheet.getCell(row, i + 1);
    cell.value = h;
    cell.fill = FILLS.ORANGE;
    cell.font = FONTS.HEADER;
    cell.border = BORDER_ALL;
    cell.alignment = { horizontal: 'center' };
  });
  row++;

  // --- ITEMS DATA LOOP ---
  data.items.forEach((item, index) => {
    const vals = [
      index + 1,
      item.productName,
      item.hsnSac,
      item.packSize,
      item.quantity,
      item.price,
      item.batchNo,
      item.mfgDate, 
      item.expDate,
      item.boxInfo,
      item.stateCode, 
      item.supplierGstin, 
      item.distCode, 
      item.grossWeight,
      item.netWeight, 
      item.uom, 
      item.gstPercent, 
      item.endUse 
    ];

    vals.forEach((v, i) => {
       const cell = sheet.getCell(row, i + 1);
       cell.value = v;
       cell.border = BORDER_ALL;
       cell.alignment = { horizontal: 'center' };
    });
    row++;
  });
  
  // --- FOOTER ---
  row++;
  sheet.mergeCells(`A${row}:E${row}`);
  sheet.getCell(`A${row}`).value = `Description: ${data.generalDescription} | IGST: ${data.globalIgst}`;
  sheet.getCell(`A${row}`).fill = FILLS.BLUE;
  sheet.getCell(`A${row}`).border = BORDER_ALL;

  // --- SAVE ---
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, `Master_Invoice_${data.invoiceNo}.xlsx`);
};