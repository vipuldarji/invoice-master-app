/**
 * downloadUtils.ts
 *
 * Central module for all invoice document generation and download.
 * page.tsx imports only from here — no ExcelJS/file-saver/pdf imports in the component.
 *
 * PDF: pure client-side via browser print engine (pdfRenderer.ts).
 *      Zero server dependencies. Works on localhost + Vercel + anywhere.
 *      User gets the native "Save as PDF" print dialog.
 */

import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { MasterData, addMasterSheet } from './excelGenerator';
import { addCommercialInvoiceSheet } from './generators/commercialInvoice';
import { addPackingListSheet } from './generators/packingList';
import { printWorkbookAsPdf } from  './pdfRenderer';

// ─── TYPES ────────────────────────────────────────────────────────────────────

export type DownloadFormat = 'xlsx' | 'pdf' | 'both';

type SheetKey = 'master' | 'invoice' | 'packing';

// Sheet name as it appears in the workbook (must match addXxxSheet calls)
const SHEET_NAMES: Record<SheetKey, string> = {
  master:  'Master Sheet',
  invoice: 'INVOICE',
  packing: 'PACKING',
};

// ─── WORKBOOK BUILDER ─────────────────────────────────────────────────────────

const buildWorkbook = async (
  data: MasterData,
  sheets: SheetKey[],
): Promise<ExcelJS.Workbook> => {
  const wb = new ExcelJS.Workbook();
  if (sheets.includes('master'))  addMasterSheet(wb, data);
  if (sheets.includes('invoice')) addCommercialInvoiceSheet(wb, data);
  if (sheets.includes('packing')) addPackingListSheet(wb, data);
  return wb;
};

// ─── XLSX DOWNLOAD ────────────────────────────────────────────────────────────

const downloadAsXlsx = async (wb: ExcelJS.Workbook, filename: string): Promise<void> => {
  const buffer = await wb.xlsx.writeBuffer();
  saveAs(
    new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    filename,
  );
};

// ─── PDF DOWNLOAD ─────────────────────────────────────────────────────────────

const downloadAsPdf = (
  wb: ExcelJS.Workbook,
  sheets: SheetKey[],
  filename: string,
): void => {
  const sheetNames = sheets.map(k => SHEET_NAMES[k]);
  printWorkbookAsPdf(wb, sheetNames, filename);
};

// ─── HIGH-LEVEL HANDLERS ──────────────────────────────────────────────────────

/**
 * Complete Set — all sheets (Master + Invoice + Packing)
 */
export const downloadCompleteSet = async (
  data: MasterData,
  format: DownloadFormat = 'both',
): Promise<void> => {
  const sheets: SheetKey[] = ['master', 'invoice', 'packing'];
  const filename = `Complete_Set_${data.invoiceNo || 'DRAFT'}.xlsx`;
  const wb = await buildWorkbook(data, sheets);

  if (format === 'xlsx' || format === 'both') await downloadAsXlsx(wb, filename);
  if (format === 'pdf'  || format === 'both') downloadAsPdf(wb, sheets, filename);
};

/**
 * Commercial Invoice only
 */
export const downloadCommercialInvoice = async (
  data: MasterData,
  format: DownloadFormat = 'both',
): Promise<void> => {
  const sheets: SheetKey[] = ['invoice'];
  const filename = `Invoice_${data.invoiceNo || 'DRAFT'}.xlsx`;
  const wb = await buildWorkbook(data, sheets);

  if (format === 'xlsx' || format === 'both') await downloadAsXlsx(wb, filename);
  if (format === 'pdf'  || format === 'both') downloadAsPdf(wb, sheets, filename);
};

/**
 * Packing List only
 */
export const downloadPackingList = async (
  data: MasterData,
  format: DownloadFormat = 'both',
): Promise<void> => {
  const sheets: SheetKey[] = ['packing'];
  const filename = `PackingList_${data.invoiceNo || 'DRAFT'}.xlsx`;
  const wb = await buildWorkbook(data, sheets);

  if (format === 'xlsx' || format === 'both') await downloadAsXlsx(wb, filename);
  if (format === 'pdf'  || format === 'both') downloadAsPdf(wb, sheets, filename);
};

/**
 * Master Sheet only (xlsx only — it's a data sheet, not a printable doc)
 */
export const downloadMasterSheet = async (
  data: MasterData,
  format: DownloadFormat = 'xlsx',
): Promise<void> => {
  const sheets: SheetKey[] = ['master'];
  const filename = `Master_${data.invoiceNo || 'DRAFT'}.xlsx`;
  const wb = await buildWorkbook(data, sheets);

  if (format === 'xlsx' || format === 'both') await downloadAsXlsx(wb, filename);
  if (format === 'pdf'  || format === 'both') downloadAsPdf(wb, sheets, filename);
};