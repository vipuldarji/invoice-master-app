import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { MasterData } from '../excelGenerator';

// --- STYLING CONSTANTS ---
const BORDER_ALL: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' }
};

const FONT_HEADER = { bold: true, size: 10, name: 'Arial' };
const FONT_TITLE = { bold: true, size: 14, underline: true, name: 'Arial' };
const FONT_BOLD = { bold: true, size: 9, name: 'Arial' };

// --- 1. SHEET GENERATION LOGIC (Reusable) ---
export const addCommercialInvoiceSheet = (workbook: ExcelJS.Workbook, data: MasterData) => {
  const sheet = workbook.addWorksheet('INVOICE', { 
    views: [{ showGridLines: false }]
  });

  // 1. Column Setup
  sheet.columns = [
    { key: 'A', width: 6 },  // Marks & Nos
    { key: 'B', width: 45 }, // Description
    { key: 'C', width: 12 }, // HSN
    { key: 'D', width: 8 },  // Pack
    { key: 'E', width: 12 }, // Batch
    { key: 'F', width: 12 }, // Expiry
    { key: 'G', width: 10 }, // UQC
    { key: 'H', width: 10 }, // Qty
    { key: 'I', width: 15 }, // Rate
    { key: 'J', width: 15 }, // Amount
  ];

  let row = 1;

  // --- HEADER TITLE ---
  sheet.mergeCells(`A${row}:J${row}`);
  const titleCell = sheet.getCell(`A${row}`);
  titleCell.value = "INVOICE";
  titleCell.alignment = { horizontal: 'center' };
  titleCell.font = FONT_TITLE;
  titleCell.border = BORDER_ALL;
  row++;

  // --- TOP BLOCK ---
  // Exporter (Left)
  sheet.mergeCells(`A${row}:D${row + 6}`);
  const exporterCell = sheet.getCell(`A${row}`);
  exporterCell.value = `EXPORTER :\n${data.exporterName}\n${data.exporterAddress}\nPhone: ${data.exporterPhone}\nEmail: ${data.exporterEmail}`;
  exporterCell.alignment = { vertical: 'top', wrapText: true };
  exporterCell.font = FONT_BOLD;
  exporterCell.border = BORDER_ALL;

  // References (Right)
  const drawRightField = (r: number, label: string, value: string) => {
    sheet.mergeCells(`E${r}:F${r}`);
    sheet.getCell(`E${r}`).value = label;
    sheet.getCell(`E${r}`).font = FONT_BOLD;
    sheet.getCell(`E${r}`).border = BORDER_ALL;
    
    sheet.mergeCells(`G${r}:J${r}`);
    sheet.getCell(`G${r}`).value = value;
    sheet.getCell(`G${r}`).font = FONT_BOLD;
    sheet.getCell(`G${r}`).border = BORDER_ALL;
  };

  drawRightField(row, "INVOICE No.", data.invoiceNo);
  drawRightField(row, "INVOICE DATE", data.invoiceDate); row++;
  drawRightField(row, "IEC No.", data.iecNo); row++;
  drawRightField(row, "Company GSTN", data.companyGstNo); row++;
  
  // IGST Status
  sheet.mergeCells(`E${row}:F${row}`);
  sheet.getCell(`E${row}`).value = "IGST PAYMENT STATUS :";
  sheet.getCell(`E${row}`).font = FONT_BOLD;
  sheet.getCell(`E${row}`).border = BORDER_ALL;
  sheet.mergeCells(`G${row}:J${row}`);
  sheet.getCell(`G${row}`).value = data.gstStatus;
  sheet.getCell(`G${row}`).font = FONT_BOLD;
  sheet.getCell(`G${row}`).border = BORDER_ALL;
  row++;

  drawRightField(row, "Drug Lic No.", data.drugLicNo); row++;
  drawRightField(row, "Buyer's Order Ref.", data.buyerOrderRef); row++;
  drawRightField(row, "Exporter Ref.", data.exporterRef); row++;

  // --- CONSIGNEE & BUYER ---
  sheet.mergeCells(`A${row}:D${row + 4}`);
  const consigneeCell = sheet.getCell(`A${row}`);
  consigneeCell.value = `CONSIGNEE :\n${data.consigneeName}\n${data.consigneeAddress}`;
  consigneeCell.alignment = { vertical: 'top', wrapText: true };
  consigneeCell.font = FONT_BOLD;
  consigneeCell.border = BORDER_ALL;

  sheet.mergeCells(`E${row}:J${row + 4}`);
  const buyerCell = sheet.getCell(`E${row}`);
  buyerCell.value = `BUYER (IF OTHER THAN CONSIGNEE) :\n${data.buyerName || 'Same as Consignee'}`;
  buyerCell.alignment = { vertical: 'top', wrapText: true };
  buyerCell.font = FONT_BOLD;
  buyerCell.border = BORDER_ALL;
  row += 5;

  // --- LOGISTICS ---
  const drawLogistics = (c1: string, c2: string, label: string, val: string) => {
    sheet.mergeCells(`${c1}${row}:${c2}${row}`);
    const cell = sheet.getCell(`${c1}${row}`);
    cell.value = label + "\n" + (val || "");
    cell.alignment = { wrapText: true, horizontal: 'center' };
    cell.font = FONT_BOLD; 
    cell.border = BORDER_ALL;
  };

  drawLogistics("A", "B", "PRE-CARRIAGE BY", data.preCarriage);
  drawLogistics("C", "D", "PLACE OF RECEIPT", data.placeOfReceipt);
  
  sheet.mergeCells(`E${row}:G${row}`);
  sheet.getCell(`E${row}`).value = "COUNTRY OF ORIGIN";
  sheet.getCell(`E${row}`).border = BORDER_ALL;
  sheet.mergeCells(`H${row}:J${row}`);
  sheet.getCell(`H${row}`).value = "INDIA";
  sheet.getCell(`H${row}`).border = BORDER_ALL;
  sheet.getCell(`H${row}`).font = FONT_BOLD;
  row++;

  drawLogistics("A", "B", "VESSEL/FLIGHT NO.", data.vesselFlight);
  drawLogistics("C", "D", "PORT OF LOADING", data.portOfLoading);

  sheet.mergeCells(`E${row}:G${row}`);
  sheet.getCell(`E${row}`).value = "COUNTRY OF FINAL DEST";
  sheet.getCell(`E${row}`).border = BORDER_ALL;
  sheet.mergeCells(`H${row}:J${row}`);
  sheet.getCell(`H${row}`).value = data.finalDestination;
  sheet.getCell(`H${row}`).border = BORDER_ALL;
  sheet.getCell(`H${row}`).font = FONT_BOLD;
  row++;

  drawLogistics("A", "B", "PORT OF DISCHARGE", data.portOfDischarge);
  drawLogistics("C", "D", "FINAL DESTINATION", data.finalDestination);

  sheet.mergeCells(`E${row}:E${row + 1}`);
  sheet.getCell(`E${row}`).value = "TERMS OF DELIVERY";
  sheet.getCell(`E${row}`).alignment = { wrapText: true, vertical: 'top' };
  sheet.getCell(`E${row}`).border = BORDER_ALL;

  sheet.mergeCells(`F${row}:J${row + 1}`);
  sheet.getCell(`F${row}`).value = data.termsOfDelivery;
  sheet.getCell(`F${row}`).font = FONT_BOLD;
  sheet.getCell(`F${row}`).alignment = { wrapText: true, vertical: 'top' };
  sheet.getCell(`F${row}`).border = BORDER_ALL;
  row++;

  sheet.mergeCells(`A${row}:B${row}`);
  sheet.getCell(`A${row}`).value = "";
  sheet.getCell(`A${row}`).border = BORDER_ALL;
  sheet.mergeCells(`C${row}:D${row}`);
  sheet.getCell(`C${row}`).value = "";
  sheet.getCell(`C${row}`).border = BORDER_ALL;
  row++; 
  
  sheet.mergeCells(`E${row}:E${row}`);
  sheet.getCell(`E${row}`).value = "PAYMENT TERMS";
  sheet.getCell(`E${row}`).border = BORDER_ALL;
  sheet.mergeCells(`F${row}:J${row}`);
  sheet.getCell(`F${row}`).value = data.paymentTerms;
  sheet.getCell(`F${row}`).font = FONT_BOLD;
  sheet.getCell(`F${row}`).border = BORDER_ALL;
  row++;

  // --- ITEM HEADERS ---
  const headers = [
    "Marks & Nos", "Description of Goods", "HSN CODE", "Pack", "Batch No.", 
    "Expiry Date", "Standard UQC", "Quantity (NOS)", "Rate Per Unit / USD", "Amount / USD"
  ];
  headers.forEach((h, i) => {
    const cell = sheet.getCell(row, i + 1);
    cell.value = h;
    cell.font = FONT_HEADER;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = BORDER_ALL;
  });
  row++;

  // --- ITEM DATA ---
  let totalQty = 0;
  let totalAmount = 0;

  data.items.forEach((item, index) => {
    const cellA = sheet.getCell(row, 1);
    cellA.value = index + 1;
    cellA.alignment = { vertical: 'top', horizontal: 'center' };
    cellA.border = BORDER_ALL;

    const cellB = sheet.getCell(row, 2);
    const descText = `${item.productName.toUpperCase()}\n${item.description || ''}\n\nSTATE CODE: ${item.stateCode}, GSTIN: ${item.supplierGstin}\nDISTRICT CODE: ${item.distCode}`;
    cellB.value = descText;
    cellB.alignment = { wrapText: true, vertical: 'top' };
    cellB.font = FONT_BOLD;
    cellB.border = BORDER_ALL;

    sheet.getCell(row, 3).value = item.hsnSac;
    sheet.getCell(row, 4).value = item.packSize;
    sheet.getCell(row, 5).value = item.batchNo;
    sheet.getCell(row, 6).value = item.expDate;
    sheet.getCell(row, 7).value = `${item.netWeight} ${item.uom}`;

    const qty = Number(item.quantity) || 0;
    sheet.getCell(row, 8).value = qty;
    totalQty += qty;

    const rate = Number(item.price) || 0;
    sheet.getCell(row, 9).value = rate;
    sheet.getCell(row, 9).numFmt = '"$"#,##0.00';

    const amount = qty * rate;
    sheet.getCell(row, 10).value = amount;
    sheet.getCell(row, 10).numFmt = '"$"#,##0.00';
    totalAmount += amount;

    for(let c=3; c<=10; c++) {
        sheet.getCell(row, c).alignment = { vertical: 'top', horizontal: 'center' };
        sheet.getCell(row, c).border = BORDER_ALL;
    }
    row++;
  });

  // --- FOOTER & TOTALS ---
  const totalRow = row;
  sheet.mergeCells(`A${row}:G${row}`);
  sheet.getCell(`A${row}`).border = BORDER_ALL;
  
  sheet.getCell(`H${row}`).value = totalQty;
  sheet.getCell(`H${row}`).font = FONT_BOLD;
  sheet.getCell(`H${row}`).border = BORDER_ALL;

  sheet.getCell(`I${row}`).value = "CIF VALUE $";
  sheet.getCell(`I${row}`).font = FONT_BOLD;
  sheet.getCell(`I${row}`).border = BORDER_ALL;

  sheet.getCell(`J${row}`).value = totalAmount;
  sheet.getCell(`J${row}`).numFmt = '"$"#,##0.00';
  sheet.getCell(`J${row}`).font = FONT_BOLD;
  sheet.getCell(`J${row}`).border = BORDER_ALL;
  row++;

  // Freight
  const freight = Number(data.freightValue) || 0;
  sheet.mergeCells(`I${row}:I${row}`);
  sheet.getCell(`I${row}`).value = "FREIGHT VALUE $";
  sheet.getCell(`I${row}`).border = BORDER_ALL;
  sheet.getCell(`J${row}`).value = freight;
  sheet.getCell(`J${row}`).numFmt = '"$"#,##0.00';
  sheet.getCell(`J${row}`).border = BORDER_ALL;

  // Box Info
  sheet.mergeCells(`A${row}:E${row}`);
  sheet.getCell(`A${row}`).value = `No. of Corrugated Boxes :  ${data.totalCorrugatedBoxes}`;
  sheet.getCell(`A${row}`).font = FONT_BOLD;
  sheet.getCell(`A${row}`).border = { left: { style: 'thin' } }; 
  
  row++;

  // Insurance
  const insurance = Number(data.insuranceValue) || 0;
  sheet.mergeCells(`I${row}:I${row}`);
  sheet.getCell(`I${row}`).value = "INSURANCE $";
  sheet.getCell(`I${row}`).border = BORDER_ALL;
  sheet.getCell(`J${row}`).value = insurance;
  sheet.getCell(`J${row}`).numFmt = '"$"#,##0.00';
  sheet.getCell(`J${row}`).border = BORDER_ALL;

  // Gross Weight
  sheet.mergeCells(`A${row}:E${row}`);
  sheet.getCell(`A${row}`).value = `Gross Weight :  ${data.totalGrossWeight} ${data.uom}`;
  sheet.getCell(`A${row}`).font = FONT_BOLD;
  sheet.getCell(`A${row}`).border = { left: { style: 'thin' } };

  row++;

  // FOB
  const fob = totalAmount - freight - insurance;
  sheet.mergeCells(`I${row}:I${row}`);
  sheet.getCell(`I${row}`).value = "FOB VALUE $";
  sheet.getCell(`I${row}`).font = FONT_BOLD;
  sheet.getCell(`I${row}`).border = BORDER_ALL;
  sheet.getCell(`J${row}`).value = fob;
  sheet.getCell(`J${row}`).numFmt = '"$"#,##0.00';
  sheet.getCell(`J${row}`).font = FONT_BOLD;
  sheet.getCell(`J${row}`).border = BORDER_ALL;

  // Net Weight
  sheet.mergeCells(`A${row}:E${row}`);
  sheet.getCell(`A${row}`).value = `Nett Weight :  ${data.totalNetWeight} ${data.uom}`;
  sheet.getCell(`A${row}`).font = FONT_BOLD;
  sheet.getCell(`A${row}`).border = { left: { style: 'thin' }, bottom: { style: 'thin' } };

  row++;

  // Words & Declaration
  sheet.mergeCells(`A${row}:J${row+1}`);
  sheet.getCell(`A${row}`).value = "AMOUNT CHARGEABLE (IN WORDS):\n" + "US DOLLARS ...ONLY"; 
  sheet.getCell(`A${row}`).alignment = { vertical: 'top', wrapText: true };
  sheet.getCell(`A${row}`).font = FONT_BOLD;
  sheet.getCell(`A${row}`).border = BORDER_ALL;
  row += 2;

  // Dynamic Declaration
  const isPaid = (data.gstStatus || "").toUpperCase().includes("PAID");
  const declText = isPaid 
    ? "* SUPPLY MEANT FOR EXPORT UNDER WITH PAYMENT OF INTEGRATED TAX (IGST)"
    : "* SUPPLY MEANT FOR EXPORT UNDER LETTER OF UNDERTAKING WITHOUT PAYMENT OF IGST *";

  sheet.mergeCells(`A${row}:F${row+3}`);
  sheet.getCell(`A${row}`).value = `DECLARATION:\nWe declare that this invoice shows actual price of the goods described and that all particulars are true and correct.\n${declText}`;
  sheet.getCell(`A${row}`).alignment = { wrapText: true, vertical: 'top' };
  sheet.getCell(`A${row}`).font = { size: 8, name: 'Arial' };
  sheet.getCell(`A${row}`).border = BORDER_ALL;

  sheet.mergeCells(`G${row}:J${row+3}`);
  sheet.getCell(`G${row}`).value = `For ${data.exporterName},\n\n\n\nAUTHORISED SIGNATORY`;
  sheet.getCell(`G${row}`).alignment = { horizontal: 'center', vertical: 'bottom', wrapText: true };
  sheet.getCell(`G${row}`).font = FONT_BOLD;
  sheet.getCell(`G${row}`).border = BORDER_ALL;
};

// --- 2. EXPORT HANDLER (Legacy Support) ---
export const generateCommercialInvoice = async (data: MasterData) => {
  const workbook = new ExcelJS.Workbook();
  addCommercialInvoiceSheet(workbook, data);
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, `Invoice_Commercial_${data.invoiceNo}.xlsx`);
};