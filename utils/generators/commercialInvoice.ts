import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { MasterData } from '../excelGenerator';

// ─── FONTS ────────────────────────────────────────────────────────────────────
const F = (bold: boolean, size: number, underline = false): Partial<ExcelJS.Font> => ({
  bold, size, name: 'Arial', ...(underline ? { underline: true } : {}),
});
const F9    = F(false, 9);
const F9B   = F(true,  9);
const F10   = F(false, 10);
const F12   = F(false, 12);
const F12B  = F(true,  12);
const F13   = F(false, 13);
const F13B  = F(true,  13);
const F14   = F(false, 14);
const F14B  = F(true,  14);
const F14BU = F(true,  14, true);
const F15B  = F(true,  15);
const F16B  = F(true,  16);

// ─── BORDERS ─────────────────────────────────────────────────────────────────
const T: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FF000000' } };
const mk = (t=false,b=false,l=false,r=false): Partial<ExcelJS.Borders> => ({
  ...(t ? {top:T}    : {}),
  ...(b ? {bottom:T} : {}),
  ...(l ? {left:T}   : {}),
  ...(r ? {right:T}  : {}),
});
const TBLR = mk(true,true,true,true);
const TLR  = mk(true,false,true,true);
const TBR  = mk(true,true,false,true);
const TBLR_ = TBLR;
const TL   = mk(true,false,true,false);
const TR_  = mk(true,false,false,true);
const T_   = mk(true,false,false,false);
const BL   = mk(false,true,true,false);
const BLR  = mk(false,true,true,true);
const BR_  = mk(false,true,false,true);
const B_   = mk(false,true,false,false);
const LR   = mk(false,false,true,true);
const L_   = mk(false,false,true,false);
const R_   = mk(false,false,false,true);
const LR_only = LR;
const NONE = mk();

// ─── ALIGNMENT ────────────────────────────────────────────────────────────────
type H = ExcelJS.Alignment['horizontal'];
type V = ExcelJS.Alignment['vertical'];
const AL = (h: H, v: V = 'middle', wrap = true): Partial<ExcelJS.Alignment> =>
  ({ horizontal: h, vertical: v, wrapText: wrap });

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const toDate = (s: string): Date | string => {
  if (!s) return s;
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : d;
};

const formatExpiry = (s: string): string => {
  if (!s) return s;
  if (/^\d{2}\/\d{4}$/.test(s)) return s;
  if (s.includes('-')) { const p = s.split('-'); return `${p[1]}/${p[0]}`; }
  return s;
};

const formatWeight = (w: string | number): string => {
  const n = Number(w);
  return isNaN(n) ? String(w) : String(n);
};

// ═════════════════════════════════════════════════════════════════════════════
// MAIN FUNCTION
// ═════════════════════════════════════════════════════════════════════════════
export const addCommercialInvoiceSheet = (workbook: ExcelJS.Workbook, data: MasterData) => {
  const sheet = workbook.addWorksheet('INVOICE', {
    views: [{ showGridLines: false }],
  });
  // Page setup must be set this way — fitToPage + fitToWidth/fitToHeight cannot coexist with scale
  sheet.pageSetup.paperSize    = 9;
  sheet.pageSetup.orientation  = 'portrait';
  sheet.pageSetup.fitToPage    = true;
  sheet.pageSetup.fitToWidth   = 1;
  sheet.pageSetup.fitToHeight  = 1;
  sheet.pageSetup.margins = { left: 0.1, right: 0.1, top: 0.25, bottom: 0.25, header: 0, footer: 0 };

  // ── Column widths ────────────────────────────────────────────────────────
  sheet.columns = [
    { width: 2.7   }, // A
    { width: 13.28 }, // B
    { width: 75.42 }, // C
    { width: 17.42 }, // D
    { width: 17.28 }, // E
    { width: 25.41 }, // F
    { width: 20.85 }, // G
    { width: 15.70 }, // H
    { width: 14.28 }, // I
    { width: 23.70 }, // J
    { width: 22.70 }, // K
  ];

  // ── Cell helpers ─────────────────────────────────────────────────────────
  const cell = (coord: string) => sheet.getCell(coord);
  const rc   = (r: number, c: number) => sheet.getCell(r, c);

  const set = (
    coord: string,
    val: ExcelJS.CellValue,
    font: Partial<ExcelJS.Font>,
    align: Partial<ExcelJS.Alignment>,
    border: Partial<ExcelJS.Borders>,
  ) => {
    const c = sheet.getCell(coord);
    c.value     = val;
    c.font      = font as ExcelJS.Font;
    c.alignment = align as ExcelJS.Alignment;
    c.border    = border as ExcelJS.Borders;
    return c;
  };

  const merge = (range: string, val: ExcelJS.CellValue,
    font: Partial<ExcelJS.Font>, align: Partial<ExcelJS.Alignment>,
    border: Partial<ExcelJS.Borders>) => {
    sheet.mergeCells(range);
    return set(range.split(':')[0], val, font, align, border);
  };

  // Apply just border to a cell (no value change)
  const brd = (coord: string, border: Partial<ExcelJS.Borders>) => {
    sheet.getCell(coord).border = border as ExcelJS.Borders;
  };
  const brdc = (r: number, c: number, border: Partial<ExcelJS.Borders>) => {
    sheet.getCell(r, c).border = border as ExcelJS.Borders;
  };

  // Drug lic strings
  const lic1 = data.drugLicDate1 ? `${data.drugLicNo1} Dated: ${data.drugLicDate1}` : data.drugLicNo1;
  const lic2 = data.drugLicDate2 ? `${data.drugLicNo2} Dated: ${data.drugLicDate2}` : data.drugLicNo2;

  // ══════════════════════════════════════════════════════════════════════════
  // R1 — spacer (orig font sz10 on all cells)
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(1).height = 17.25;
  for (let c = 1; c <= 11; c++) sheet.getCell(1, c).font = F10 as ExcelJS.Font;

  // ══════════════════════════════════════════════════════════════════════════
  // R2 — INVOICE title (B2:K2 merged, 26pt bold underline, center)
  // NOTE: Original B2 is sz=26B — use F26BU
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(2).height = 31.5;
  merge('B2:K2', 'INVOICE', F(true, 26, true), AL('center'), TBLR);

  // ══════════════════════════════════════════════════════════════════════════
  // R3 — Header row
  // Original borders: B=tl, C=t, D=t, E=tr, F=tblr(full), G=tbr, H=tr, I=tr, J=tblr, K=tbr
  // Original fonts:   B=14B, C/D/E=12, F=14B, G=15B, H=14B, I=10, J=15B, K=10
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(3).height = 24;
  set('B3', 'EXPORTER :',                                          F14B, AL('left'),  TL);
  set('C3', '',                                                    F12,  AL('left'),  T_);
  set('D3', '',                                                    F12,  AL('left'),  T_);
  set('E3', '',                                                    F12,  AL('left'),  TR_);
  set('F3', 'INVOICE No.      ',                                   F14B, AL('left'),  TBLR);
  set('G3', data.invoiceNo,                                        F15B, AL('left'),  TBR);
  // H3:I3 merged — original H=tr border on merged cell
  merge('H3:I3', ' INVOICE DATE ', F14B, AL('left'), TR_);
  // J3:K3 merged — original J=tblr
  merge('J3:K3', toDate(data.invoiceDate) as ExcelJS.CellValue, F15B, AL('left'), TBLR);
  sheet.getCell('J3').numFmt = 'dd-mmm-yy';

  // ══════════════════════════════════════════════════════════════════════════
  // R4 — exporter name | IEC No.
  // Original borders: B=l, C/D/E=none, F=tbl, G=tbr, H/I/J/K=tb
  // Fonts: B=14B, F=14(not bold!), G=13B, H=14, I/J=12B, K=12
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(4).height = 22.5;
  set('B4', data.exporterName, F14B, AL('left'), L_);
  // C4/D4/E4 — no border, sz12 (must set explicitly or ExcelJS defaults to 11)
  for (const c of [3,4,5]) { sheet.getCell(4,c).font = F12 as ExcelJS.Font; }
  set('F4', 'IEC No.', F14, AL('left'), mk(true,true,true,false));  // tbl
  set('G4', '',        F13B, AL('left'), TBR);
  set('H4', data.iecNo, F14, AL('left'), mk(true,true,false,false));
  set('I4', '',         F12B, AL('left'), mk(true,true,false,false));
  set('J4', '',         F12B, AL('left'), mk(true,true,false,false));
  set('K4', '',         F12,  AL('center'), TBR);

  // ══════════════════════════════════════════════════════════════════════════
  // R5 — addr line 1 | GSTN
  // Original: B=l, F=tbl, G=tbr, H/I/J/K=tb
  // Fonts: B=14B, F=14, G=14, H=14, I/J/K=12
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(5).height = 18.75;
  set('B5', data.exporterAddressLine1, F14B, AL('left'), L_);
  for (const c of [3,4,5]) { sheet.getCell(5,c).font = F12 as ExcelJS.Font; }
  set('F5', "Company's GSTN No. ", F14, AL('left'),  mk(true,true,true,false));
  set('G5', '',                     F14, AL('left'),  TBR);
  set('H5', data.companyGstNo,      F14, AL('left'),  mk(true,true,false,false));
  set('I5', '',  F12, AL('left'),   mk(true,true,false,false));
  set('J5', '',  F12, AL('left'),   mk(true,true,false,false));
  set('K5', '',  F12, AL('left'),   TBR);

  // ══════════════════════════════════════════════════════════════════════════
  // R6 — addr line 2 | IGST STATUS
  // Original: B=l, F=tbl, G=tbr, H/I/J/K=tb
  // Fonts: B=14B, F=14B, G=14, H=14B, I/J/K=12
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(6).height = 22.5;
  set('B6', data.exporterAddressLine2, F14B, AL('left'), L_);
  // C3=14, C4/C5=12 (orig has 14 on C3 specifically)
  sheet.getCell(6,3).font = F14 as ExcelJS.Font;
  sheet.getCell(6,4).font = F12 as ExcelJS.Font;
  sheet.getCell(6,5).font = F12 as ExcelJS.Font;
  set('F6', 'IGST PAYMENT STATUS : ', F14B, AL('left'), mk(true,true,true,false));
  set('G6', '',                        F14,  AL('left'), TBR);
  set('H6', data.gstStatus,            F14B, AL('left'), mk(true,true,false,false));
  set('I6', '',  F12, AL('left'), mk(true,true,false,false));
  set('J6', '',  F12, AL('left'), mk(true,true,false,false));
  set('K6', '',  F12, AL('left'), TBR);

  // ══════════════════════════════════════════════════════════════════════════
  // R7 — addr line 3 | Drug Lic No. 1
  // Original: B=l, F=l, G=r, H-J=none, K=r
  // Fonts: B=14B, F/G/H=14
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(7).height = 18.75;
  set('B7', data.exporterAddressLine3, F14B, AL('left'), L_);
  for (const c of [3,4,5]) { sheet.getCell(7,c).font = F12 as ExcelJS.Font; }
  set('F7', 'Drug Lic No.', F14, AL('left'), L_);
  set('G7', '',             F14, AL('left'), R_);
  // H7: value, no border
  const h7 = sheet.getCell('H7');
  h7.value = lic1; h7.font = F14 as ExcelJS.Font; h7.border = NONE as ExcelJS.Borders;
  // I7/J7/K7 — sz12, no border
  for (const c of [9,10,11]) { sheet.getCell(7,c).font = F12 as ExcelJS.Font; }
  brd('K7', R_);

  // ══════════════════════════════════════════════════════════════════════════
  // R8 — phone | Drug Lic No. 2 (in H8, no border)
  // Original: B=l, F=l, G=r, H=none, K=r
  // Fonts: B=14B, F/G/H=14
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(8).height = 19.5;
  set('B8', `PHONE NO-${data.exporterPhone}`, F14B, AL('left'), L_);
  for (const c of [3,4,5]) { sheet.getCell(8,c).font = F12 as ExcelJS.Font; }
  set('F8', '', F14, AL('left'), L_);
  set('G8', '', F14, AL('left'), R_);
  const h8 = sheet.getCell('H8');
  h8.value = lic2; h8.font = F14 as ExcelJS.Font; h8.border = NONE as ExcelJS.Borders;
  // I8/J8/K8 — sz12, no border
  for (const c of [9,10,11]) { sheet.getCell(8,c).font = F12 as ExcelJS.Font; }
  brd('K8', R_);

  // ══════════════════════════════════════════════════════════════════════════
  // R9 — email | Buyer's Order Ref
  // Original: B=l, F=tbl, G=tbr, H/I/J/K=tb
  // Fonts: B=14B, F/G=12, H/I/J=12B, K=12
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(9).height = 18;
  set('B9', `Email ID: ${data.exporterEmail}`, F14B, AL('left'), L_);
  for (const c of [3,4,5]) { sheet.getCell(9,c).font = F12 as ExcelJS.Font; }
  set('F9', "Buyer's Order Ref.No. :", F12, AL('left'), mk(true,true,true,false));
  set('G9', '', F12, AL('left'), TBR);
  set('H9', data.buyerOrderRef || '', F12B, AL('left'), mk(true,true,false,false));
  set('I9', '', F12B, AL('left'), mk(true,true,false,false));
  set('J9', '', F12B, AL('left'), mk(true,true,false,false));
  set('K9', '', F12,  AL('left'), TBR);

  // ══════════════════════════════════════════════════════════════════════════
  // R10 — blank | Exporter Ref. and Date
  // Original: B=bl, C/D/E=b, F=bl, G=br, H/I/J/K=b
  // Fonts: B/C/D/E=12, F/G=12, H/I/J=12B, K=12
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(10).height = 18;
  set('B10', '', F12, AL('left'), BL);
  set('C10', '', F12, AL('left'), B_);
  set('D10', '', F12, AL('left'), B_);
  set('E10', '', F12, AL('left'), B_);
  set('F10', 'Exporter Ref. and Date :', F12, AL('left'), BL);
  set('G10', '', F12, AL('left'), BR_);
  set('H10', data.exporterRef || '', F12B, AL('left'), B_);
  set('I10', '', F12B, AL('left'), B_);
  set('J10', '', F12B, AL('left'), B_);
  set('K10', '', F12,  AL('left'), BR_);

  // ══════════════════════════════════════════════════════════════════════════
  // R11 — CONSIGNEE / BUYER labels
  // Original: B=tl, C/D/E=t, F=l, I=r, K=r
  // Fonts: B=14B, F=14B, I=12B
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(11).height = 18.75;
  set('B11', 'CONSIGNEE  :', F14B, AL('left'), TL);
  // C/D/E — top border, sz12
  for (const col of ['C11','D11','E11']) { sheet.getCell(col).border = T_ as ExcelJS.Borders; sheet.getCell(col).font = F12 as ExcelJS.Font; sheet.getCell(col).alignment = { horizontal: 'left' } as ExcelJS.Alignment; }
  set('F11', 'BUYER (IF OTHER THAN CONSIGNEE)  :', F14B, AL('left'), L_);
  // G/H — sz12, left
  for (const col of ['G11','H11']) { sheet.getCell(col).font = F12 as ExcelJS.Font; sheet.getCell(col).alignment = { horizontal: 'left' } as ExcelJS.Alignment; }
  // I11:K11 merge — original has this merge, I11=sz12B/center/r
  sheet.mergeCells('I11:K11');
  const i11 = sheet.getCell('I11');
  i11.font = F12B as ExcelJS.Font; i11.border = R_ as ExcelJS.Borders;
  i11.alignment = { horizontal: 'center' } as ExcelJS.Alignment;

  // ══════════════════════════════════════════════════════════════════════════
  // R12-R17 — consignee address block
  // Original: B=l, F=l, K=r — all other cols NO border
  // Fonts: original has sz14B-16B on these rows (carry-over from row above)
  // ══════════════════════════════════════════════════════════════════════════
  // R12-R17 font data from original:
  // R12: B=15B, C/D/E=14B/left, F=16B, G-K=14B
  // R13-R16: B=14B, C/D/E=14B/left, F=15B, G-K=14B
  // R17: B=14B, C/D/E=14B/left, F=14B, G-K=14B
  const consigneeData: [number, string, number, number][] = [
    [12, data.consigneeName,    15, 16],
    [13, data.consigneeAddress, 14, 15],
    [14, '',  14, 15], [15, '', 14, 15], [16, '', 14, 15], [17, '', 14, 14],
  ];
  consigneeData.forEach(([r, val, bSz, fSz]) => {
    sheet.getRow(r).height = 21.6;
    set(`B${r}`, val, F(true, bSz), AL('left'), L_);
    // C/D/E — 14B, left, no border
    for (const c of [3,4,5]) {
      sheet.getCell(r, c).font = F14B as ExcelJS.Font;
      sheet.getCell(r, c).alignment = { horizontal: 'left' } as ExcelJS.Alignment;
    }
    set(`F${r}`, '', F(true, fSz), AL('left'), L_);
    // G-K — 14B, no border
    for (const c of [7,8,9,10,11]) {
      sheet.getCell(r, c).font = F14B as ExcelJS.Font;
    }
    brd(`K${r}`, R_);
  });

  // ══════════════════════════════════════════════════════════════════════════
  // R18-R23 — LOGISTICS BLOCK
  //
  // R18: B18:C18=tlr/tbr merged, D18:E18=tr/tbr merged, F=tb, G=tbr, H18:I18=tbl/tb merged, J=tb, K=tbr
  // R19: B19:C19=blr/br, D19:E19=br/br, F=tb, G=tbr, H=tb, I/J/K=tb/tbr
  // R20: B20:C20=tlr/tr, D20:E20=tr/tr, F=t, G=tr, H/I/J=t, K=tr
  // R21: B21:C21=blr/br, D21:E21=br/br, F=b, G=br, H/I/J=b, K=br
  // R22: B22:C22=lr/r, D22:E22=r/r, F=none, G=r, H22:K23=lr..
  // R23: B23:C23=blr/br, D23:E23=br/br, F=none, G=r
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(18).height = 27;
  merge('B18:C18', '  PRE-CARRIAGE BY', F14, AL('center'), TLR);
  merge('D18:E18', 'PLACE OF RECEIPT',  F14, AL('center'), TR_);
  set('F18', 'COUNTRY OF ORIGIN OF GOODS ', F12, AL('left'), mk(true,true,false,false));
  set('G18', '', F12, AL('left'), TBR);
  merge('H18:I18', 'INDIA ', F14B, AL('left'), mk(true,true,true,false));
  set('J18', '', F14, AL('left'), mk(true,true,false,false));
  set('K18', '', F14, AL('left'), TBR);

  sheet.getRow(19).height = 23.25;
  merge('B19:C19', data.preCarriage,    F14B, AL('center'), BLR);
  merge('D19:E19', data.placeOfReceipt, F14B, AL('center'), BR_);
  set('F19', 'COUNTRY OF FINAL DESTINATION', F12, AL('left'), mk(true,true,false,false));
  set('G19', '', F12B, AL('left'), TBR);  // G=12B
  set('H19', data.finalDestination, F14B, AL('left'), mk(true,true,false,false));
  set('I19', '', F14B, AL('left'), mk(true,true,false,false));
  set('J19', '', F14B, AL('left'), mk(true,true,false,false));
  set('K19', '', F14B, AL('left'), TBR);

  sheet.getRow(20).height = 23.25;
  merge('B20:C20', '  VESSEL/FLIGHT NO.', F14, AL('center'), TLR);
  merge('D20:E20', 'PORT OF LOADING ',    F14, AL('center'), TR_);
  set('F20', 'TERMS OF DELIVERY  ', F13, AL('left'), T_);
  set('G20', '', F12, AL('left'), TR_);
  set('H20', data.termsOfDelivery, F15B, AL('left'), T_);
  set('I20', '', F13B, AL('left'), T_);   // I=13B matches original
  set('J20', '', F14B, AL('left'), T_);   // J=14B
  set('K20', '', F14B, AL('left'), TR_);  // K=14B

  sheet.getRow(21).height = 26.45;
  merge('B21:C21', data.vesselFlight || '', F13B, AL('center'), BLR);
  merge('D21:E21', data.portOfLoading,      F15B, AL('center'), BR_);
  set('F21', '', F13B, AL('left'),  B_);
  set('G21', '', F12B, AL('left'),  BR_);  // G=12B
  set('H21', '', F14B, AL('center'), B_);
  set('I21', '', F14B, AL('center'), B_);
  set('J21', '', F14B, AL('center'), B_);
  set('K21', '', F14B, AL('center'), BR_);

  sheet.getRow(22).height = 24;
  merge('B22:C22', '  PORT OF DISCHARGE', F14, AL('center'), LR);
  merge('D22:E22', 'FINAL DESTINATION',   F14, AL('center'), R_);
  set('F22', 'PAYMENT TERMS ', F13, AL('left'), NONE);
  set('G22', '', F12, AL('left'), R_);
  merge('H22:K23', data.paymentTerms, F14B, AL('left'), LR);

  sheet.getRow(23).height = 32.45;
  merge('B23:C23', data.portOfDischarge,  F15B, AL('center'), BLR);
  merge('D23:E23', data.finalDestination, F15B, AL('center'), BR_);
  set('F23', '', F12B, AL('left'), NONE);
  set('G23', '', F12B, AL('left'), R_);  // G=12B matches original

  // ══════════════════════════════════════════════════════════════════════════
  // R24 — column headers
  // Original: B/C/D/E=blr, F-K=tblr. Font=13B. Alignment=center/middle.
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(24).height = 63.6;
  const hdrBorders = ['B','C','D','E'].map(c => [c, BLR]) as [string, Partial<ExcelJS.Borders>][];
  const hdrBorders2 = ['F','G','H','I','J','K'].map(c => [c, TBLR]) as [string, Partial<ExcelJS.Borders>][];
  const hdrDefs: [string, string][] = [
    ['B','Marks & Nos.'], ['C','Description of Goods'], ['D','HSN CODE'],
    ['E','Pack'], ['F','Batch No.'], ['G','Expiry Date'],
    ['H','Standard UQC'], ['I','Quantity (NOS)'],
    ['J','Rate Per Unit  / USD'], ['K','Amount / USD'],
  ];
  [...hdrBorders, ...hdrBorders2].forEach(([col, bord]) => {
    sheet.getCell(`${col}24`).border = bord as ExcelJS.Borders;
  });
  hdrDefs.forEach(([col, label]) => {
    set(`${col}24`, label, F13B, AL('center', 'middle'), sheet.getCell(`${col}24`).border as Partial<ExcelJS.Borders>);
  });

  // ══════════════════════════════════════════════════════════════════════════
  // R25 — spacer below headers
  // Original: same lr/r pattern, font 13, specific alignments per col
  // C2=lr/left, C3=lr/left, C4=lr/left, C5=r/left, C6=r/left, C7=r/left,
  // C8=lr/left, C9=r/left, C10=lr/left, C11=r/right
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(25).height = 17.25;
  const r25borders = { 2:LR,3:LR,4:LR,5:R_,6:R_,7:R_,8:LR,9:R_,10:LR,11:R_ };
  const r25halign: Record<number,H> = { 2:'left',3:'left',4:'left',5:'left',6:'left',7:'left',8:'left',9:'left',10:'left',11:'right' };
  for (let c = 2; c <= 11; c++) {
    const cl = sheet.getCell(25, c);
    cl.font   = F13 as ExcelJS.Font;
    cl.border = (r25borders as Record<number, Partial<ExcelJS.Borders>>)[c] as ExcelJS.Borders;
    cl.alignment = { horizontal: r25halign[c] } as ExcelJS.Alignment;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // ITEM ROWS — R26 onward, 2 rows per item
  //
  // Item row: ALL cells = lr/r pattern, font 12B, bold
  //   C2=lr/center, C3=lr/left, C4=lr/center, C5=r/center, C6=r/center,
  //   C7=r/center, C8=lr/center, C9=r/center, C10=lr/RIGHT, C11=r/RIGHT
  //
  // State row: font 12 (not bold) on C3, 13B on C2+C4-C11
  //   C2=lr/center(13B), C3=lr/left(12), C4=lr(13), C5=r(13), C6=r(13),
  //   C7=r(13), C8=lr/center(13), C9=r/left(13), C10=lr/center(13), C11=r/center(13)
  // ══════════════════════════════════════════════════════════════════════════
  const ITEM_ROW_H = [40.9, 37.9, 47.25, 40.9, 63.0, 47.25, 58.9, 63.0, 47.25, 63.0, 47.25, 47.25, 47.25];

  let row = 26;
  let totalQty    = 0;
  let totalAmount = 0;

  data.items.forEach((item, idx) => {
    sheet.getRow(row).height = ITEM_ROW_H[idx] ?? 45;

    const qty    = Number(item.quantity) || 0;
    const price  = Number(item.price)    || 0;
    const amount = Math.round(qty * price * 100) / 100;
    totalQty    += qty;
    totalAmount  = Math.round((totalAmount + amount) * 100) / 100;

    // item row — font 12B bold, lr/r border pattern, alignment matches original
    set(`B${row}`, idx + 1,  F12B, AL('center','top',false), LR);
    const desc = `${item.productName}\n${item.description || ''}\n`;
    set(`C${row}`, desc,     F12B, { horizontal:'left', vertical:'top', wrapText:true } as ExcelJS.Alignment, LR);
    set(`D${row}`, Number(item.hsnSac) || item.hsnSac, F12B, AL('center','top',false), LR);
    set(`E${row}`, item.packSize,              F12B, AL('center','top',false), R_);
    set(`F${row}`, item.batchNo,               F12B, AL('center','top',false), R_);
    set(`G${row}`, formatExpiry(item.expDate), F12B, AL('center','top',false), R_);
    set(`H${row}`, ` ${item.netWeight} ${item.uom}`, F12B, AL('center','top',false), LR);
    set(`I${row}`, qty,   F12B, AL('center','top',false), R_);
    set(`J${row}`, price, F12B, AL('right','top',false), LR);
    sheet.getCell(`J${row}`).numFmt = '"$"#,##0.00';
    set(`K${row}`, amount, F12B, AL('right','top',false), R_);
    sheet.getCell(`K${row}`).numFmt = '"$"#,##0.00';
    row++;

    // state sub-row — NO MERGE, individual cells
    sheet.getRow(row).height = 30;
    const stateText = `STATE CODE :  ${item.stateCode}, GSTIN No.: ${item.supplierGstin}                                            DISTRICT CODE :  ${item.distCode}`;

    // C2: 13B, center, lr
    const b2 = sheet.getCell(row, 2);
    b2.font = F13B as ExcelJS.Font; b2.border = LR as ExcelJS.Borders;
    b2.alignment = { horizontal: 'center' } as ExcelJS.Alignment;

    // C3: 12 not bold, left, lr — state text
    const s3 = sheet.getCell(row, 3);
    s3.value = stateText; s3.font = F12 as ExcelJS.Font;
    s3.alignment = { horizontal: 'left', vertical: 'middle' } as ExcelJS.Alignment;
    s3.border = LR as ExcelJS.Borders;

    // C4=lr(13), C5=r(13), C6=r(13), C7=r(13), C8=lr/center(13), C9=r/left(13), C10=lr/center(13), C11=r/center(13)
    const stBorders: Record<number, Partial<ExcelJS.Borders>> = { 4:LR, 5:R_, 6:R_, 7:R_, 8:LR, 9:R_, 10:LR, 11:R_ };
    const stAlign:   Record<number, H> = { 4:'left', 5:'left', 6:'left', 7:'left', 8:'center', 9:'left', 10:'center', 11:'center' };
    for (let c = 4; c <= 11; c++) {
      const sc = sheet.getCell(row, c);
      sc.font = F13 as ExcelJS.Font;
      sc.border = (stBorders[c] || NONE) as ExcelJS.Borders;
      sc.alignment = { horizontal: stAlign[c] } as ExcelJS.Alignment;
    }
    row++;
  });

  // ══════════════════════════════════════════════════════════════════════════
  // PAD ROWS — fill to R63
  // Original pad rows carry forward alternating 12B/13 font sizes from item/state rows.
  // Pad row 0 (R52) = item row style (12B), pad row 1 (R53) = state row style (13), etc.
  // Heights: [15.75, 16.5, 66, 16.5, 16.5, 16.5, 16.5, 16.5, 16.5, 16.5, 19.5]
  // ══════════════════════════════════════════════════════════════════════════
  const totalsRow  = 63;
  const padHeights = [15.75, 16.5, 66, 16.5, 16.5, 16.5, 16.5, 16.5, 16.5, 16.5, 19.5];
  let padIdx = 0;

  // Exact pad row specs from original (R52-R62):
  // R52 (padIdx=0): 12B — item style
  // R53 (padIdx=1): C2=13B/cen, C3=12/lef, C4-C7=13/lef, C8=13/cen, C9=13/lef, C10-11=13/cen
  // R54-R61 (padIdx=2-9): item-style but sz=13B (not 12B), except C3 which cycles
  //   Even padIdx (2,4,6,8): ALL 13B — C2-11 all bold sz13, alignments like item row
  //   Odd padIdx (3,5,7,9): C2=13B/cen, C3=13/lef (NOT 12), C4-C7=13/lef, C8=13/cen etc
  // R62 (padIdx=10): 13 NOT bold — C2-C8=13/lef, C9=13/cen, C10-C11=14/cen
  const padBorders = { 2:LR, 3:LR, 4:LR, 5:R_, 6:R_, 7:R_, 8:LR, 9:R_, 10:LR, 11:R_ } as Record<number, Partial<ExcelJS.Borders>>;

  // Item-style alignments (even pads after R52)
  const itemAligns: Record<number,H> = { 2:'center',3:'left',4:'center',5:'center',6:'center',7:'center',8:'center',9:'center',10:'right',11:'right' };
  // State-style alignments (odd pads)
  const stateAligns: Record<number,H> = { 2:'center',3:'left',4:'left',5:'left',6:'left',7:'left',8:'center',9:'left',10:'center',11:'center' };

  while (row < totalsRow) {
    sheet.getRow(row).height = padHeights[padIdx] ?? 16.5;

    if (padIdx === 10) {
      // R62 — last pad: 13 NOT bold, specific alignment
      const r62aligns: Record<number,H> = { 2:'left',3:'left',4:'left',5:'left',6:'left',7:'left',8:'left',9:'center',10:'center',11:'center' };
      const r62sizes:  Record<number,number> = { 2:13,3:13,4:13,5:13,6:13,7:13,8:13,9:13,10:14,11:14 };
      for (let c = 2; c <= 11; c++) {
        const cl = sheet.getCell(row, c);
        cl.border    = padBorders[c] as ExcelJS.Borders;
        cl.font      = F(false, r62sizes[c]) as ExcelJS.Font;
        cl.alignment = { horizontal: r62aligns[c] } as ExcelJS.Alignment;
      }
    } else if (padIdx === 0) {
      // R52 — first pad: 12B, item-style
      for (let c = 2; c <= 11; c++) {
        const cl = sheet.getCell(row, c);
        cl.border    = padBorders[c] as ExcelJS.Borders;
        cl.font      = F12B as ExcelJS.Font;
        cl.alignment = { horizontal: itemAligns[c] } as ExcelJS.Alignment;
      }
    } else if (padIdx % 2 === 0) {
      // Even pads (R54,R56,R58,R60): 13B item-style (sz13 not 12)
      // R58/R60 have C6/C7 at sz12B
      const isLateEven = padIdx >= 6;
      for (let c = 2; c <= 11; c++) {
        const cl = sheet.getCell(row, c);
        cl.border    = padBorders[c] as ExcelJS.Borders;
        const sz = (isLateEven && (c === 6 || c === 7)) ? 12 : 13;
        cl.font      = F(true, sz) as ExcelJS.Font;
        cl.alignment = { horizontal: itemAligns[c] } as ExcelJS.Alignment;
      }
    } else {
      // Odd pads: C2=13B, C3 size depends on which odd pad, C4+=13 not bold
      // R53(padIdx=1): C3=sz12. R55+(padIdx>=3): C3=sz13
      // R59(padIdx=7): C6=sz12. All others: C6=sz13
      const c3size = padIdx === 1 ? 12 : 13;
      for (let c = 2; c <= 11; c++) {
        const cl = sheet.getCell(row, c);
        cl.border    = padBorders[c] as ExcelJS.Borders;
        const isBold = c === 2;
        let sz = 13;
        if (c === 3) sz = c3size;
        if (c === 6 && padIdx === 7) sz = 12; // only R59 has C6=12
        cl.font      = F(isBold, sz) as ExcelJS.Font;
        cl.alignment = { horizontal: stateAligns[c] } as ExcelJS.Alignment;
      }
    }
    row++; padIdx++;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // R63 — TOTALS ROW
  // B=tl/16B, C=t/16B, D=t/14, E=tr/14, F/G/H=t/14, I=tblr/14B, J=tlr/14B(right), K=tr/15B(right)
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(63).height = 25.9;
  set('B63', ` No. of  Corrugated Boxes :   ${data.totalCorrugatedBoxes}`, F16B, AL('left'), TL);
  set('C63', '',  F16B, AL('left'),  T_);
  set('D63', ' ', F14,  AL('left'),  T_);
  set('E63', '',  F14,  AL('left'),  TR_);
  set('F63', '',  F14,  AL('left'),  T_);
  set('G63', '',  F14,  AL('left'),  T_);
  set('H63', '',  F14,  AL('left'),  T_);
  set('I63', totalQty,    F14B, AL('center'), TBLR);
  set('J63', 'CIF VALUE', F14B, AL('right'),  TLR);
  set('K63', totalAmount, F15B, AL('right'),  TR_);
  sheet.getCell('K63').numFmt = '"$"#,##0.00';
  // Align D/E/F/G to match original 'general' + 'left' mix
  sheet.getCell('C63').alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  sheet.getCell('D63').alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  sheet.getCell('E63').alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  sheet.getCell('F63').alignment = { horizontal: 'left' } as ExcelJS.Alignment;

  // ══════════════════════════════════════════════════════════════════════════
  // R64 — Gross weight | FREIGHT VALUE
  // B=l/16B, C=none/16B, D/E=none+r/14, F/G/H=none/14, I=lr/14, J=tblr/13B, K=tblr/15B
  // ══════════════════════════════════════════════════════════════════════════
  const fr  = Number(data.freightValue)   || 0;
  const ins = Number(data.insuranceValue) || 0;
  const fob = Math.round((totalAmount - fr - ins) * 100) / 100;

  sheet.getRow(64).height = 32.45;
  set('B64', ` Gross Weight :  ${formatWeight(data.totalGrossWeight)} KGS`, F16B, AL('left'), L_);
  const c64 = sheet.getCell('C64'); c64.font = F16B as ExcelJS.Font; c64.border = NONE as ExcelJS.Borders; c64.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  // D/E/F/G/H — sz14, set font + alignment
  for (const [coord, ha] of [['D64','left'],['E64','left'],['F64','left'],['G64','center'],['H64','center']] as [string,H][]) {
    const cl = sheet.getCell(coord); cl.font = F14 as ExcelJS.Font; cl.alignment = { horizontal: ha } as ExcelJS.Alignment;
  }
  brd('E64', R_);
  const i64 = sheet.getCell('I64'); i64.font = F14 as ExcelJS.Font; i64.border = LR as ExcelJS.Borders; i64.alignment = { horizontal: 'center' } as ExcelJS.Alignment;
  set('J64', 'FREIGHT VALUE', F13B, AL('right'), TBLR);
  set('K64', fr,              F15B, AL('right'), TBLR);
  sheet.getCell('K64').numFmt = '"$"#,##0.00';

  // ══════════════════════════════════════════════════════════════════════════
  // R65 — Net weight | INSURANCE
  // B=bl/16B, C=b/16B, D=b/14, E=br/14, F/G/H=none/14, I=blr/14, J=tblr/13B, K=tblr/15B
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(65).height = 30;
  set('B65', ` Nett Weight :  ${formatWeight(data.totalNetWeight)} KGS`, F16B, AL('left'), BL);
  set('C65', '', F16B, AL('left'),  B_);
  set('D65', '', F14,  AL('left'),  B_);
  set('E65', '', F14,  AL('left'),  BR_);
  // F/G/H — sz14, no border
  for (const [coord, ha] of [['F65','left'],['G65','center'],['H65','center']] as [string,H][]) {
    const cl = sheet.getCell(coord); cl.font = F14 as ExcelJS.Font; cl.alignment = { horizontal: ha } as ExcelJS.Alignment;
  }
  const i65 = sheet.getCell('I65'); i65.font = F14 as ExcelJS.Font; i65.border = BLR as ExcelJS.Borders; i65.alignment = { horizontal: 'center' } as ExcelJS.Alignment;
  set('J65', 'INSURANCE', F13B, AL('right'), TBLR);
  set('K65', ins,         F15B, AL('right'), TBLR);
  sheet.getCell('K65').numFmt = '"$"#,##0.00';

  // ══════════════════════════════════════════════════════════════════════════
  // R66 — AMOUNT CHARGEABLE | FOB VALUE
  // B=tl/14B, C=t/14, D-H=t/13, I=tlr/13, J=tlr/14B, K=tr/15B
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(66).height = 19.5;
  set('B66', 'AMOUNT CHARGEABLE  :', F14B, AL('left'), TL);
  set('C66', '', F14, AL('left'),  T_);
  set('D66', '', F13, AL('left'),  T_);
  set('E66', '', F13, AL('left'),  T_);
  set('F66', '', F13, AL('left'),  T_);
  set('G66', '', F13, AL('left'),  T_);
  set('H66', '', F13, AL('left'),  T_);
  set('I66', '', F13, AL('left'),  TLR);
  sheet.getCell('I66').alignment = AL('left'); // original ha=general (empty)
  set('J66', 'FOB VALUE', F14B, AL('right'), TLR);
  set('K66', fob,         F15B, AL('right'), TR_);
  sheet.getCell('K66').numFmt = '"$"#,##0.00';

  // ══════════════════════════════════════════════════════════════════════════
  // R67 — closing border row (original had broken #VALUE! formula)
  // B=bl/16B, C=b/16B, D-H=b/13, I=blr/13, J/K=blr+br/14
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(67).height = 25.5;
  const r67fonts: Record<number, Partial<ExcelJS.Font>> = { 2:F16B,3:F16B,4:F13,5:F13,6:F13,7:F13,8:F13,9:F13,10:F14,11:F14 };
  const r67borders: Record<number, Partial<ExcelJS.Borders>> = { 2:BL,3:B_,4:B_,5:B_,6:B_,7:B_,8:B_,9:BLR,10:BLR,11:BR_ };
  for (let c = 2; c <= 11; c++) {
    const cl = sheet.getCell(67, c);
    cl.font   = (r67fonts[c] || F13) as ExcelJS.Font;
    cl.border = (r67borders[c] || NONE) as ExcelJS.Borders;
    cl.alignment = AL('left'); // ha=general (matches original)
  }

  // ══════════════════════════════════════════════════════════════════════════
  // R68:R69 — Legal text block
  // B68:H69 merged: border tlr on B68 (merge head), font 13
  // C68-H68 top border only (sz10), I68=tl/14B, J68=t/13, K68=tr/13
  // R69: B=l, H=r, I=l, K=r  (merge bottom row borders)
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(68).height = 31.5;
  sheet.getRow(69).height = 4.5;
  const legalText =
    '(WE UNDERTAKE TO ABIDE BY PROVISIONS OF FOREIGN EXCHANGE MANAGEMENT ACT,1999, AS  \n' +
    'AMENDED FROM TIME TO TIME, INCLUDING REALIZATION/REPATRIATION OF FOREIGN EXCHANGE TO OR FROM INDIA) \n';
  merge('B68:H69', legalText, F13, AL('left', 'top'), TLR);
  // I68 = tl / 14B / general (not left)
  set('I68', `FOR ${data.exporterName.toUpperCase()},`, F14B, AL('left') as ExcelJS.Alignment, TL);
  // J68 = t / sz13
  const j68 = sheet.getCell('J68'); j68.border = T_ as ExcelJS.Borders; j68.font = F13 as ExcelJS.Font;
  // K68 = tr / sz13
  const k68 = sheet.getCell('K68'); k68.border = TR_ as ExcelJS.Borders; k68.font = F13 as ExcelJS.Font;
  // R69 right panel: I=l/sz13/left, J=none/sz13/left, K=r/sz13/left
  const i69 = sheet.getCell('I69'); i69.border = L_ as ExcelJS.Borders; i69.font = F13 as ExcelJS.Font; i69.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const j69 = sheet.getCell('J69'); j69.font = F13 as ExcelJS.Font; j69.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const k69 = sheet.getCell('K69'); k69.border = R_ as ExcelJS.Borders; k69.font = F13 as ExcelJS.Font; k69.alignment = { horizontal: 'left' } as ExcelJS.Alignment;

  // ══════════════════════════════════════════════════════════════════════════
  // R70 — IGST note
  // B=tbl/14B, C-H=tb/13B, I=l/14B, K=r/13
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(70).height = 25.5;
  const isPaid = (data.gstStatus || '').toUpperCase().includes('PAID');
  const igstLine = isPaid
    ? '* SUPPLY MEANT FOR EXPORT UNDER WITH PAYMENT OF INTEGRATED TAX (IGST)                               '
    : '* SUPPLY MEANT FOR EXPORT UNDER LETTER OF UNDERTAKING WITHOUT PAYMENT OF IGST';
  // B=tbl/14B/general, C-G=tb/13B/general, H=tbr/13B/general
  const b70 = sheet.getCell('B70'); b70.value = igstLine; b70.font = F14B as ExcelJS.Font; b70.border = mk(true,true,true,false) as ExcelJS.Borders; b70.alignment = AL('left');
  for (let c = 3; c <= 7; c++) { const cl = sheet.getCell(70,c); cl.font = F13B as ExcelJS.Font; cl.border = mk(true,true,false,false) as ExcelJS.Borders; cl.alignment = AL('left'); }
  const h70 = sheet.getCell('H70'); h70.font = F13B as ExcelJS.Font; h70.border = mk(true,true,false,true) as ExcelJS.Borders; h70.alignment = AL('left');
  // I=l/14B/general, J=none/13/left, K=r/13/left
  const i70 = sheet.getCell('I70'); i70.font = F14B as ExcelJS.Font; i70.border = L_ as ExcelJS.Borders; i70.alignment = AL('left');
  const j70 = sheet.getCell('J70'); j70.font = F13 as ExcelJS.Font; j70.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const k70 = sheet.getCell('K70'); k70.font = F13 as ExcelJS.Font; k70.border = R_ as ExcelJS.Borders; k70.alignment = { horizontal: 'left' } as ExcelJS.Alignment;

  // ══════════════════════════════════════════════════════════════════════════
  // R71 — tiny gap row
  // B=tlr/13B, C-H=t/10, I=l/14B, K=r/13
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(71).height = 7.9;
  // B71:H71 — merged in original, tlr border, 13B, left
  merge('B71:H71', '', F13B, AL('left'), TLR);
  // I=l/14B/general, J=none/13/left, K=r/13/left
  const i71 = sheet.getCell('I71'); i71.font = F14B as ExcelJS.Font; i71.border = L_ as ExcelJS.Borders; i71.alignment = AL('left');
  const j71 = sheet.getCell('J71'); j71.font = F13 as ExcelJS.Font; j71.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const k71 = sheet.getCell('K71'); k71.font = F13 as ExcelJS.Font; k71.border = R_ as ExcelJS.Borders; k71.alignment = { horizontal: 'left' } as ExcelJS.Alignment;

  // ══════════════════════════════════════════════════════════════════════════
  // R72 — DBK/RoDTEP note
  // B=tlr/16B, C-H=t/10, I=l/14B, K=r/13
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(72).height = 25.15;
  // B72:H72 — merged in original, tlr border, 16B, left
  merge('B72:H72',
    `* * ${data.dbkRodtepNote || 'No DBK or RoDTEP for All Items ( ITS FREE SHIPPING BILL)& GNX 100'}`,
    F16B, AL('left'), TLR);
  // I=l/14B/general, J=none/13/left, K=r/13/left
  const i72 = sheet.getCell('I72'); i72.font = F14B as ExcelJS.Font; i72.border = L_ as ExcelJS.Borders; i72.alignment = AL('left');
  const j72 = sheet.getCell('J72'); j72.font = F13 as ExcelJS.Font; j72.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const k72 = sheet.getCell('K72'); k72.font = F13 as ExcelJS.Font; k72.border = R_ as ExcelJS.Borders; k72.alignment = { horizontal: 'left' } as ExcelJS.Alignment;

  // ══════════════════════════════════════════════════════════════════════════
  // R73 — Declaration label
  // B=tl/13B, C-H=t/13, I=l/13, K=r/13
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(73).height = 17.25;
  set('B73', 'Declaration:', F13B, AL('left'), TL);
  for (let c = 3; c <= 8; c++) {
    const cl = sheet.getCell(73, c);
    cl.font = F13 as ExcelJS.Font;
    cl.border = (c === 8 ? TR_ : T_) as ExcelJS.Borders;
    cl.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
    if (c === 4) cl.value = ' ';
  }
  // I=l/13/left, J=none/13/left, K=r/13/left
  const i73 = sheet.getCell('I73'); i73.font = F13 as ExcelJS.Font; i73.border = L_ as ExcelJS.Borders; i73.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const j73 = sheet.getCell('J73'); j73.font = F13 as ExcelJS.Font; j73.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const k73 = sheet.getCell('K73'); k73.font = F13 as ExcelJS.Font; k73.border = R_ as ExcelJS.Borders; k73.alignment = { horizontal: 'left' } as ExcelJS.Alignment;

  // ══════════════════════════════════════════════════════════════════════════
  // R74 — Declaration text + AUTHORISED SIGNATORY
  // B=blr/13, C-H=b/10, I=bl/14(NOT bold), J=b/13, K=br/13
  // ══════════════════════════════════════════════════════════════════════════
  sheet.getRow(74).height = 24.6;
  merge('B74:H74',
    'We declare that this invoice shows actual price of the goods described and that all particulars are true and correct.',
    F13, AL('left', 'middle'), BLR);
  // I=bl/14(NOT bold)/left, J=b/13/left, K=br/13/left
  const i74 = sheet.getCell('I74'); i74.value = 'AUTHORISED SIGNATORY'; i74.font = F14 as ExcelJS.Font; i74.border = BL as ExcelJS.Borders; i74.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const j74 = sheet.getCell('J74'); j74.font = F13 as ExcelJS.Font; j74.border = B_ as ExcelJS.Borders; j74.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
  const k74 = sheet.getCell('K74'); k74.font = F13 as ExcelJS.Font; k74.border = BR_ as ExcelJS.Borders; k74.alignment = { horizontal: 'left' } as ExcelJS.Alignment;
};

// ─── STANDALONE XLSX EXPORT ──────────────────────────────────────────────────
export const generateCommercialInvoice = async (data: MasterData) => {
  const workbook = new ExcelJS.Workbook();
  addCommercialInvoiceSheet(workbook, data);
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(
    new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    `Invoice_${data.invoiceNo}.xlsx`,
  );
};