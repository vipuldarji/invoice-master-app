"use client";

import React, { useState } from 'react';
import { useForm, useFieldArray, SubmitHandler, UseFormRegister, Path } from 'react-hook-form';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import {
  Download, Package, Plus, Trash2, ChevronDown, ChevronRight,
  Plane, FileText, Layers, FileBadge, Box, DollarSign,
  Building2, Ship, ClipboardList, Weight, CheckCircle2,
  Loader2, FileDown, FileSpreadsheet,
} from 'lucide-react';
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem,
  DropdownMenuTrigger, DropdownMenuSeparator, DropdownMenuLabel,
} from "@/components/ui/dropdown-menu";
import { MasterData } from '@/utils/excelGenerator';
import {
  downloadCompleteSet,
  downloadCommercialInvoice,
  downloadPackingList,
  downloadMasterSheet,
  type DownloadFormat,
} from '@/utils/downloadUtils';

// ─── FIELD COMPONENT ──────────────────────────────────────────────────────────

interface FieldProps {
  label: string;
  register: UseFormRegister<MasterData>;
  name: Path<MasterData>;
  type?: string;
  placeholder?: string;
  className?: string;
  required?: boolean;
}

const Field: React.FC<FieldProps> = ({
  label, register, name, type = "text", placeholder = "", className = "", required = false,
}) => (
  <div className={`flex flex-col gap-1.5 ${className}`}>
    <label className="text-[10px] font-semibold tracking-widest text-slate-400 uppercase flex items-center gap-1">
      {label}
      {required && <span className="text-rose-400">*</span>}
    </label>
    <Input
      type={type}
      {...register(name)}
      placeholder={placeholder}
      className="h-9 text-[13px] bg-white/60 border-slate-200 focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 rounded-lg text-slate-800 placeholder:text-slate-300 transition-all"
    />
  </div>
);

// ─── SECTION CARD ─────────────────────────────────────────────────────────────

interface SectionProps {
  icon: React.ReactNode;
  title: string;
  subtitle?: string;
  accent: string;
  children: React.ReactNode;
  badge?: string;
}

const Section: React.FC<SectionProps> = ({ icon, title, subtitle, accent, children, badge }) => (
  <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
    <div className="flex items-center gap-3 px-6 py-4" style={{ borderLeft: `4px solid ${accent}` }}>
      <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0" style={{ background: `${accent}18` }}>
        <span style={{ color: accent }}>{icon}</span>
      </div>
      <div className="flex-1">
        <h3 className="text-[13px] font-bold text-slate-700 tracking-tight">{title}</h3>
        {subtitle && <p className="text-[11px] text-slate-400 mt-0.5">{subtitle}</p>}
      </div>
      {badge && (
        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold" style={{ background: `${accent}18`, color: accent }}>
          {badge}
        </span>
      )}
    </div>
    <div className="px-6 py-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-x-6 gap-y-4">
      {children}
    </div>
  </div>
);

const FullWidth: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div className="col-span-1 sm:col-span-2 lg:col-span-3">{children}</div>
);

const Half: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <div className="col-span-1">{children}</div>
);

// ─── STEP DEFINITIONS ─────────────────────────────────────────────────────────

const steps = [
  { id: 0, label: "Parties",    icon: <Building2 size={14} />,     color: "#6366f1" },
  { id: 1, label: "Regulatory", icon: <FileBadge size={14} />,     color: "#0ea5e9" },
  { id: 2, label: "Financials", icon: <DollarSign size={14} />,    color: "#10b981" },
  { id: 3, label: "Logistics",  icon: <Plane size={14} />,         color: "#f59e0b" },
  { id: 4, label: "Packing",    icon: <Box size={14} />,           color: "#8b5cf6" },
  { id: 5, label: "Line Items", icon: <ClipboardList size={14} />, color: "#ef4444" },
];

// ─── MAIN COMPONENT ───────────────────────────────────────────────────────────

export default function MasterInvoiceApp() {
  const [activeStep, setActiveStep] = useState(0);
  const [expandedRow, setExpandedRow] = useState<number | null>(null);
  const [downloading, setDownloading] = useState<string | null>(null); // key of in-progress download

  const { register, control, handleSubmit, watch } = useForm<MasterData>({
    defaultValues: {
      // Exporter
      exporterName: "",
      exporterAddressLine1: "",
      exporterAddressLine2: "",
      exporterAddressLine3: "",
      exporterPhone: "", exporterEmail: "", exporterRef: "",
      // Consignee
      consigneeName: "", consigneeAddress: "",
      buyerName: "", buyerOrderRef: "", chaName: "",
      // Regulatory
      iecNo: "", gstStatus: "", companyGstNo: "",
      drugLicNo1: "", drugLicDate1: "",
      drugLicNo2: "", drugLicDate2: "",
      lutRef: "", lutDate: "",
      // Remittance
      remittanceRef: "", remittanceDate: "", remittanceAmount: "",
      remittanceAvailable: "", remittanceUsed: "",
      // Financials
      proformaValue: "", invoiceValue110: "", invoiceValue110Round: "",
      adcRate: "", exchangeRate: 0, inrValue: "",
      freightValue: 0, insuranceValue: 0, currency: "USD", uom: "KGS",
      igstPercent: 0.05,
      // Logistics
      invoiceNo: "", invoiceDate: new Date().toISOString().split('T')[0],
      packingListNo: "", placeOfReceipt: "", portOfLoading: "",
      portOfDischarge: "", finalDestination: "",
      preCarriage: "By AIR", vesselFlight: "", flightDate: "",
      paymentTerms: "", termsOfDelivery: "",
      // Shipping docs
      shippingBillNo: "", shippingBillDate: "",
      awbNo: "", awbDate: "", policyNo: "", policyDate: "",
      // Packing
      totalGrossWeight: "", totalNetWeight: "", totalCorrugatedBoxes: "",
      generalDescription: "",
      manufacturerName: "", manufacturerAddress: "",
      // Items
      items: [{
        productName: "", hsnSac: "", packSize: "", quantity: 0, price: 0,
        batchNo: "", mfgDate: "", expDate: "", boxInfo: "",
        grossWeight: 0, netWeight: 0, supplierGstin: "", stateCode: "",
        distCode: "", gstPercent: 0, uom: "", endUse: "",
        genericName: "", description: ""
      }],
      boxDimensions: [{ boxNo: "Box # 01", dimensions: "" }]
    }
  });

  const { fields: itemFields, append: appendItem, remove: removeItem } = useFieldArray({ control, name: "items" });
  const { fields: boxFields, append: appendBox, remove: removeBox } = useFieldArray({ control, name: "boxDimensions" });

  const watchedItems = watch("items");
  const watchedInvoiceNo = watch("invoiceNo");
  const totalValue = watchedItems?.reduce((sum, item) =>
    sum + ((Number(item.quantity) || 0) * (Number(item.price) || 0)), 0) || 0;

  // ─── DOWNLOAD HANDLERS ────────────────────────────────────────────────────
  // Each handler: validate form → build xlsx → optionally convert to PDF via /api/pdf
  // 'format' controls whether we emit xlsx, pdf, or both.

  const withDownload = (key: string, fn: () => Promise<void>) => async () => {
    setDownloading(key);
    try {
      await fn();
    } catch (e) {
      console.error('[download]', e);
      alert(`Download failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setDownloading(null);
    }
  };

  const onDownloadMaster: SubmitHandler<MasterData> = (data) =>
    withDownload('master', () => downloadMasterSheet(data, 'xlsx'))();

  const onDownloadCommercial: SubmitHandler<MasterData> = (data) =>
    withDownload('invoice-xlsx', () => downloadCommercialInvoice(data, 'xlsx'))();

  const onDownloadCommercialPdf: SubmitHandler<MasterData> = (data) =>
    withDownload('invoice-pdf', () => downloadCommercialInvoice(data, 'pdf'))();

  const onDownloadPacking: SubmitHandler<MasterData> = (data) =>
    withDownload('packing-xlsx', () => downloadPackingList(data, 'xlsx'))();

  const onDownloadPackingPdf: SubmitHandler<MasterData> = (data) =>
    withDownload('packing-pdf', () => downloadPackingList(data, 'pdf'))();

  const onDownloadCombined: SubmitHandler<MasterData> = (data) =>
    withDownload('complete-xlsx', () => downloadCompleteSet(data, 'xlsx'))();

  const onDownloadCombinedPdf: SubmitHandler<MasterData> = (data) =>
    withDownload('complete-pdf', () => downloadCompleteSet(data, 'pdf'))();

  const onDownloadCombinedBoth: SubmitHandler<MasterData> = (data) =>
    withDownload('complete-both', () => downloadCompleteSet(data, 'both'))();

  // ─── DEMO / TEST DATA ─────────────────────────────────────────────────────
  // All values below are generic placeholders for development and testing only.
  // No real company, product, or regulatory data is present.
  // Replace with your actual data at runtime — do NOT commit real data here.

  const fillTestData = () => {
    control._reset({
      exporterName: "YOUR COMPANY NAME",
      exporterAddressLine1: "UNIT 1, BUSINESS PARK,",
      exporterAddressLine2: "INDUSTRIAL AREA,",
      exporterAddressLine3: "CITY-000000, STATE, INDIA.",
      exporterPhone: "+91-0000000000",
      exporterEmail: "exports@yourcompany.com",
      exporterRef: "",

      consigneeName: "TO THE ORDER OF BUYER",
      consigneeAddress: "CITY, DESTINATION COUNTRY",
      buyerName: "",
      buyerOrderRef: "",
      chaName: "YOUR CHA NAME",

      iecNo: "XXXXXXXXXX",
      gstStatus: "PAID",
      companyGstNo: "00XXXXXXXXXXXXX",

      drugLicNo1: "00X-XX-XXX-000000",
      drugLicDate1: "01/01/2020",
      drugLicNo2: "00X-XX-XXX-000001",
      drugLicDate2: "01/01/2020",

      lutRef: "XXXXXXXXXXXXXXX",
      lutDate: "01/01/2020",

      remittanceRef:       "TT-DEMO001",
      remittanceDate:      "2025-01-01",
      remittanceAmount:    "0",
      remittanceAvailable: "0",
      remittanceUsed:      "0",

      proformaValue:        "0.00",
      invoiceValue110:      "0.00",
      invoiceValue110Round: "0.00",
      adcRate:              "0.00",
      exchangeRate:         0,
      inrValue:             "0",
      freightValue:         0,
      insuranceValue:       0,
      currency:             "USD",
      uom:                  "KGS",
      igstPercent:          0.05,

      invoiceNo:        "INV-000001",
      invoiceDate:      "2025-01-01",
      packingListNo:    "INV-000001",
      placeOfReceipt:   "Origin Airport",
      portOfLoading:    "Origin Airport",
      portOfDischarge:  "DESTINATION PORT",
      finalDestination: "DESTINATION COUNTRY",
      preCarriage:      "By AIR",
      vesselFlight:     "",
      flightDate:       "2025-01-01",
      paymentTerms:     "100% ADVANCE WITH ORDER",
      termsOfDelivery:  "By AIR CIF DESTINATION",

      shippingBillNo: "", shippingBillDate: "",
      awbNo: "", awbDate: "",
      policyNo: "",
      policyDate: "2025-01-01",

      totalGrossWeight:     "0.000",
      totalNetWeight:       "0.000",
      totalCorrugatedBoxes: "00",
      generalDescription:   "PHARMACEUTICAL PRODUCTS / OPHTHALMIC / MEDICAL DEVICES",
      manufacturerName:     "MANUFACTURER A / MANUFACTURER B / MANUFACTURER C",
      manufacturerAddress:  "CITY A (STATE) / CITY B (STATE) / CITY C (STATE)",

      items: [
        {
          productName: "SAMPLE PRODUCT 01 5ML",
          description: "ACTIVE INGREDIENT A OPHTHALMIC SOLUTION 0.3% W/V",
          genericName: "ACTIVE INGREDIENT A",
          endUse: "Sample therapeutic indication",
          hsnSac: "30049099", packSize: "5ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-001", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 01",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 02 3ML",
          description: "ACTIVE INGREDIENT B EYE DROPS 0.04MG",
          genericName: "ACTIVE INGREDIENT B",
          endUse: "Sample therapeutic indication",
          hsnSac: "30049099", packSize: "3ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-002", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 02",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 12, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 03 5ML",
          description: "ACTIVE INGREDIENT C EYE DROPS 1% W/V",
          genericName: "ACTIVE INGREDIENT C",
          endUse: "Sample therapeutic indication",
          hsnSac: "30049099", packSize: "5ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-003", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 03",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 12, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 04 3ML",
          description: "ACTIVE INGREDIENT D EYE DROPS 0.05% W/V",
          genericName: "ACTIVE INGREDIENT D",
          endUse: "Sample therapeutic indication",
          hsnSac: "30043110", packSize: "3ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-004", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 04",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 05 10ML",
          description: "ACTIVE INGREDIENT E EYE DROPS 0.5% W/V",
          genericName: "ACTIVE INGREDIENT E",
          endUse: "Sample therapeutic indication",
          hsnSac: "30049099", packSize: "10ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-005", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 05",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 06 10ML",
          description: "ACTIVE INGREDIENT F EYE DROPS 5MG",
          genericName: "ACTIVE INGREDIENT F",
          endUse: "Sample therapeutic indication",
          hsnSac: "30049099", packSize: "10ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-006", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 06",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 07 5ML",
          description: "FDC OF INGREDIENT G + INGREDIENT H EYE DROPS, SUSPENSION",
          genericName: "INGREDIENT G + INGREDIENT H",
          endUse: "Sample therapeutic indication",
          hsnSac: "30042019", packSize: "5ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-007", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 07",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 08 10ML",
          description: "ACTIVE INGREDIENT I OPHTHALMIC SUSPENSION 10MG",
          genericName: "ACTIVE INGREDIENT I",
          endUse: "Sample therapeutic indication",
          hsnSac: "30043200", packSize: "10ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-008", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 08",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 09 5ML",
          description: "FDC OF INGREDIENT J + INGREDIENT K OPHTHALMIC SUSPENSION",
          genericName: "INGREDIENT J + INGREDIENT K",
          endUse: "Sample therapeutic indication",
          hsnSac: "30042019", packSize: "5ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-009", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 09",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 10 5ML",
          description: "ACTIVE INGREDIENT L OPHTHALMIC SUSPENSION",
          genericName: "ACTIVE INGREDIENT L",
          endUse: "Sample therapeutic indication",
          hsnSac: "30043200", packSize: "5ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-010", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 10",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 12, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 11 5ML",
          description: "ACTIVE INGREDIENT M EYE DROPS 1MG",
          genericName: "ACTIVE INGREDIENT M",
          endUse: "Sample therapeutic indication",
          hsnSac: "30043200", packSize: "5ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-011", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 11",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 12, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 12 1ML VIAL",
          description: "ACTIVE INGREDIENT N STERILE SOLUTION 0.8MG",
          genericName: "ACTIVE INGREDIENT N",
          endUse: "Sample surgical indication",
          hsnSac: "30051090", packSize: "1ML VIAL",
          quantity: 0, price: 0,
          batchNo: "BATCH-012", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 12",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 12, uom: "KGS",
        },
        {
          productName: "SAMPLE PRODUCT 13 5ML",
          description: "ACTIVE INGREDIENT O OPHTHALMIC SOLUTION 0.01% W/V",
          genericName: "ACTIVE INGREDIENT O",
          endUse: "Sample therapeutic indication",
          hsnSac: "30049099", packSize: "5ML",
          quantity: 0, price: 0,
          batchNo: "BATCH-013", mfgDate: "2025-01-01", expDate: "2027-01-01",
          boxInfo: "BOX # 13",
          grossWeight: 0, netWeight: 0,
          supplierGstin: "00XXXXXXXXXXXXX", stateCode: "00", distCode: "DISTRICT",
          gstPercent: 5, uom: "KGS",
        },
      ],

      boxDimensions: [
        { boxNo: "Box # 01", dimensions: "00 cms X 00 cms X 00 cms" },
        { boxNo: "Box # 02", dimensions: "00 cms X 00 cms X 00 cms" },
        { boxNo: "Box # 03", dimensions: "00 cms X 00 cms X 00 cms" },
        { boxNo: "Box # 04", dimensions: "00 cms X 00 cms X 00 cms" },
        { boxNo: "Box # 05", dimensions: "00 cms X 00 cms X 00 cms" },
        { boxNo: "Box # 06", dimensions: "00 cms X 00 cms X 00 cms" },
      ],
    });
  };

  // ─── DOWNLOAD MENU ────────────────────────────────────────────────────────
  // Spinner helper — shows while a specific download key is in progress
  const Spin = ({ k }: { k: string }) =>
    downloading === k ? <Loader2 size={12} className="ml-auto animate-spin" /> : null;

  const DownloadMenu = () => (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button
          disabled={!!downloading}
          className="bg-indigo-600 hover:bg-indigo-700 text-white shadow-md shadow-indigo-200 rounded-xl h-9 px-4 text-[13px] font-semibold disabled:opacity-60"
        >
          {downloading
            ? <Loader2 size={14} className="mr-2 animate-spin" />
            : <Download size={14} className="mr-2" />}
          <span className="hidden sm:inline">Download</span>
          <ChevronDown size={12} className="ml-1.5" />
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="w-64 rounded-xl shadow-xl border-slate-100 p-1.5">

        {/* ── Complete Set ─────────────────────────────────────────────── */}
        <DropdownMenuLabel className="text-[11px] text-slate-400 font-medium px-2 pt-1 pb-0.5 uppercase tracking-wide">
          Complete Set
        </DropdownMenuLabel>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadCombinedBoth)}
          className="rounded-lg cursor-pointer font-semibold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 p-3 mb-0.5"
        >
          <Layers size={14} className="mr-2 shrink-0" />
          All Sheets — XLSX + PDF
          <Spin k="complete-both" />
        </DropdownMenuItem>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadCombined)}
          className="rounded-lg cursor-pointer p-2.5 mb-0.5 text-[13px]"
        >
          <FileSpreadsheet size={13} className="mr-2 shrink-0 text-emerald-600" />
          All Sheets — XLSX only
          <Spin k="complete-xlsx" />
        </DropdownMenuItem>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadCombinedPdf)}
          className="rounded-lg cursor-pointer p-2.5 text-[13px]"
        >
          <FileDown size={13} className="mr-2 shrink-0 text-rose-500" />
          All Sheets — PDF only
          <Spin k="complete-pdf" />
        </DropdownMenuItem>

        <DropdownMenuSeparator className="my-1.5" />

        {/* ── Commercial Invoice ───────────────────────────────────────── */}
        <DropdownMenuLabel className="text-[11px] text-slate-400 font-medium px-2 pt-1 pb-0.5 uppercase tracking-wide">
          Commercial Invoice
        </DropdownMenuLabel>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadCommercial)}
          className="rounded-lg cursor-pointer p-2.5 mb-0.5 text-[13px]"
        >
          <FileSpreadsheet size={13} className="mr-2 shrink-0 text-emerald-600" />
          Invoice — XLSX
          <Spin k="invoice-xlsx" />
        </DropdownMenuItem>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadCommercialPdf)}
          className="rounded-lg cursor-pointer p-2.5 text-[13px]"
        >
          <FileDown size={13} className="mr-2 shrink-0 text-rose-500" />
          Invoice — PDF
          <Spin k="invoice-pdf" />
        </DropdownMenuItem>

        <DropdownMenuSeparator className="my-1.5" />

        {/* ── Packing List ─────────────────────────────────────────────── */}
        <DropdownMenuLabel className="text-[11px] text-slate-400 font-medium px-2 pt-1 pb-0.5 uppercase tracking-wide">
          Packing List
        </DropdownMenuLabel>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadPacking)}
          className="rounded-lg cursor-pointer p-2.5 mb-0.5 text-[13px]"
        >
          <FileSpreadsheet size={13} className="mr-2 shrink-0 text-emerald-600" />
          Packing List — XLSX
          <Spin k="packing-xlsx" />
        </DropdownMenuItem>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadPackingPdf)}
          className="rounded-lg cursor-pointer p-2.5 text-[13px]"
        >
          <FileDown size={13} className="mr-2 shrink-0 text-rose-500" />
          Packing List — PDF
          <Spin k="packing-pdf" />
        </DropdownMenuItem>

        <DropdownMenuSeparator className="my-1.5" />

        {/* ── Master Sheet ─────────────────────────────────────────────── */}
        <DropdownMenuLabel className="text-[11px] text-slate-400 font-medium px-2 pt-1 pb-0.5 uppercase tracking-wide">
          Master Sheet
        </DropdownMenuLabel>
        <DropdownMenuItem
          onClick={handleSubmit(onDownloadMaster)}
          className="rounded-lg cursor-pointer p-2.5 text-[13px]"
        >
          <FileText size={13} className="mr-2 shrink-0 text-slate-500" />
          Master Data — XLSX
          <Spin k="master" />
        </DropdownMenuItem>

      </DropdownMenuContent>
    </DropdownMenu>
  );

  // ─── RENDER ───────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen font-sans" style={{ background: 'linear-gradient(135deg, #f8faff 0%, #f0f4ff 50%, #fafafa 100%)' }}>

      {/* ── STICKY HEADER ── */}
      <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-xl border-b border-slate-100 shadow-sm">
        <div className="max-w-screen-2xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between gap-4">

          {/* Logo */}
          <div className="flex items-center gap-3 shrink-0">
            <div className="w-9 h-9 bg-indigo-600 rounded-xl flex items-center justify-center shadow-md shadow-indigo-200">
              <Package size={18} className="text-white" />
            </div>
            <div className="hidden sm:block">
              <div className="text-[14px] font-bold text-slate-800 leading-tight tracking-tight">Master Invoice Engine</div>
              <div className="text-[10px] text-indigo-400 font-semibold tracking-widest uppercase">Export Documentation Suite · Ver 2026</div>
            </div>
          </div>

          {/* Step tabs */}
          <nav className="hidden lg:flex items-center gap-1 bg-slate-50 rounded-xl p-1 border border-slate-100">
            {steps.map((step) => (
              <button
                key={step.id}
                onClick={() => setActiveStep(step.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-semibold transition-all ${
                  activeStep === step.id
                    ? 'bg-white text-slate-800 shadow-sm border border-slate-200'
                    : 'text-slate-400 hover:text-slate-600'
                }`}
              >
                <span style={{ color: activeStep === step.id ? step.color : undefined }}>{step.icon}</span>
                {step.label}
              </button>
            ))}
          </nav>

          {/* Right */}
          <div className="flex items-center gap-3">
            <div className="hidden sm:flex flex-col items-end">
              <div className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Invoice Total</div>
              <div className="text-[18px] font-bold text-indigo-600 font-mono leading-tight">
                ${totalValue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
              </div>
            </div>
            <button
              onClick={fillTestData}
              className="hidden md:flex items-center gap-1.5 px-3 py-2 rounded-lg border border-amber-200 bg-amber-50 text-amber-600 text-[11px] font-semibold hover:bg-amber-100 transition-colors"
            >
              ⚡ Test Data
            </button>
            <DownloadMenu />
          </div>
        </div>

        {/* Mobile step strip */}
        <div className="lg:hidden flex overflow-x-auto gap-1 px-4 pb-3 scrollbar-none">
          {steps.map((step) => (
            <button
              key={step.id}
              onClick={() => setActiveStep(step.id)}
              className={`shrink-0 flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-semibold transition-all border ${
                activeStep === step.id ? 'bg-white border-slate-200 text-slate-800 shadow-sm' : 'border-transparent text-slate-400'
              }`}
            >
              {step.icon} {step.label}
            </button>
          ))}
        </div>
      </header>

      {/* ── MAIN CONTENT ── */}
      <main className="max-w-screen-2xl mx-auto px-4 sm:px-6 py-6 space-y-4">

        {/* Progress bar */}
        <div className="flex items-center gap-3">
          <div className="flex-1 h-1.5 bg-slate-100 rounded-full overflow-hidden">
            <div
              className="h-full bg-indigo-500 rounded-full transition-all duration-500 ease-out"
              style={{ width: `${((activeStep + 1) / steps.length) * 100}%` }}
            />
          </div>
          <span className="text-[11px] text-slate-400 font-semibold shrink-0">
            {activeStep + 1} / {steps.length}
          </span>
        </div>

        {/* ── STEP 0: PARTIES ── */}
        {activeStep === 0 && (
          <div className="space-y-4">
            <Section icon={<Building2 size={16} />} title="Exporter Details" subtitle="Your company's export information" accent="#6366f1">
              <FullWidth>
                <Field label="Exporter / Company Name" register={register} name="exporterName" required placeholder="YOUR COMPANY NAME" />
              </FullWidth>
              <FullWidth>
                <Field label="Address Line 1" register={register} name="exporterAddressLine1" placeholder="UNIT 1, BUSINESS PARK," />
              </FullWidth>
              <FullWidth>
                <Field label="Address Line 2" register={register} name="exporterAddressLine2" placeholder="INDUSTRIAL AREA," />
              </FullWidth>
              <FullWidth>
                <Field label="Address Line 3 (City / PIN / Country)" register={register} name="exporterAddressLine3" placeholder="CITY-000000, STATE, INDIA." />
              </FullWidth>
              <Field label="Phone Number" register={register} name="exporterPhone" placeholder="+91-0000000000" />
              <Field label="Email Address" register={register} name="exporterEmail" type="email" placeholder="exports@yourcompany.com" />
              <Field label="Exporter Reference No." register={register} name="exporterRef" placeholder="Reference number (optional)" />
            </Section>

            <Section icon={<Building2 size={16} />} title="Consignee Details" subtitle="Who receives the shipment" accent="#0ea5e9">
              <Half>
                <Field label="Consignee Name" register={register} name="consigneeName" required placeholder="TO THE ORDER OF BUYER" />
              </Half>
              <Half>
                <Field label="Consignee Address" register={register} name="consigneeAddress" placeholder="City, Country" />
              </Half>
              <Field label="Buyer Name (if different from consignee)" register={register} name="buyerName" placeholder="Leave blank if same" />
              <Field label="Buyer Order Reference" register={register} name="buyerOrderRef" placeholder="PO-2025-XXXXX" />
              <Field label="CHA (Clearing Agent) Name" register={register} name="chaName" placeholder="YOUR CHA NAME" />
            </Section>

            <Section icon={<Building2 size={16} />} title="Manufacturer Details" subtitle="Product origin and manufacturing details" accent="#f59e0b">
              <Half>
                <Field label="Manufacturer Name(s)" register={register} name="manufacturerName" placeholder="MANUFACTURER A / MANUFACTURER B..." />
              </Half>
              <Half>
                <Field label="Manufacturer Address(es)" register={register} name="manufacturerAddress" placeholder="CITY A (STATE) / CITY B (STATE)..." />
              </Half>
            </Section>
          </div>
        )}

        {/* ── STEP 1: REGULATORY ── */}
        {activeStep === 1 && (
          <div className="space-y-4">
            <Section icon={<FileBadge size={16} />} title="Regulatory Identifiers" subtitle="Licenses, GST, and compliance codes" accent="#0ea5e9">
              <Field label="IEC Number" register={register} name="iecNo" required placeholder="XXXXXXXXXX" />
              <Field label="Company GST Number" register={register} name="companyGstNo" placeholder="00XXXXXXXXXXXXX" />
              <Field label="GST Payment Status" register={register} name="gstStatus" placeholder="PAID / LUT" />
              <Half>
                <Field label="Drug License No. 1" register={register} name="drugLicNo1" placeholder="00X-XX-XXX-000000" />
              </Half>
              <Half>
                <Field label="Drug Lic 1 — Date Issued" register={register} name="drugLicDate1" placeholder="DD/MM/YYYY" />
              </Half>
              <Half>
                <Field label="Drug License No. 2" register={register} name="drugLicNo2" placeholder="00X-XX-XXX-000001" />
              </Half>
              <Half>
                <Field label="Drug Lic 2 — Date Issued" register={register} name="drugLicDate2" placeholder="DD/MM/YYYY" />
              </Half>
              <Field label="LUT Reference Number" register={register} name="lutRef" placeholder="XXXXXXXXXXXXXXX" />
              <Field label="LUT Date" register={register} name="lutDate" placeholder="DD/MM/YYYY" />
            </Section>

            <Section icon={<ClipboardList size={16} />} title="Goods Description & Tax" subtitle="General shipment description and tax configuration" accent="#10b981">
              <FullWidth>
                <Field label="General Description of Goods" register={register} name="generalDescription" placeholder="PHARMACEUTICAL EYE DROPS & INJECTION / OPHTHALMIC MEDICAL DEVICES..." />
              </FullWidth>
              <Field label="Global IGST Rate (decimal, e.g. 0.05)" register={register} name="igstPercent" type="number" placeholder="0.05" />
              <Field label="Currency" register={register} name="currency" placeholder="USD" />
              <Field label="Default Unit of Measure (UOM)" register={register} name="uom" placeholder="KGS" />
            </Section>
          </div>
        )}

        {/* ── STEP 2: FINANCIALS ── */}
        {activeStep === 2 && (
          <div className="space-y-4">
            <Section icon={<DollarSign size={16} />} title="Remittance / Advance Payment" subtitle="TT and advance payment tracking" accent="#10b981">
              <Field label="TT / Remittance Reference" register={register} name="remittanceRef" placeholder="TT-REF-001" />
              <Field label="TT Date" register={register} name="remittanceDate" type="date" />
              <Field label="TT Amount (USD)" register={register} name="remittanceAmount" placeholder="0.00" />
              <Field label="Amount Available" register={register} name="remittanceAvailable" placeholder="0.00" />
              <Field label="Amount To Use (This Shipment)" register={register} name="remittanceUsed" placeholder="0.00" />
            </Section>

            <Section icon={<DollarSign size={16} />} title="Invoice Valuation" subtitle="Proforma, exchange rates, and computed values" accent="#6366f1">
              <Field label="Proforma Invoice Value" register={register} name="proformaValue" placeholder="0.00" />
              <Field label="110% Invoice Value" register={register} name="invoiceValue110" placeholder="0.00" />
              <Field label="110% Value (Rounded)" register={register} name="invoiceValue110Round" placeholder="0.00" />
              <Field label="ADC Rate" register={register} name="adcRate" placeholder="0.00" />
              <Field label="Exchange Rate (INR per USD)" register={register} name="exchangeRate" type="number" placeholder="0.00" />
              <Field label="Total INR Value" register={register} name="inrValue" placeholder="0" />
              <Field label="Freight Value (USD)" register={register} name="freightValue" type="number" placeholder="0.00" />
              <Field label="Insurance Value (USD)" register={register} name="insuranceValue" type="number" placeholder="0.00" />
            </Section>
          </div>
        )}

        {/* ── STEP 3: LOGISTICS ── */}
        {activeStep === 3 && (
          <div className="space-y-4">
            <Section icon={<FileText size={16} />} title="Invoice & Document Numbers" subtitle="Reference numbers for all export documents" accent="#f59e0b">
              <Field label="Invoice Number" register={register} name="invoiceNo" required placeholder="INV-000001" />
              <Field label="Invoice Date" register={register} name="invoiceDate" type="date" required />
              <Field label="Packing List Number" register={register} name="packingListNo" placeholder="INV-000001" />
              <Field label="Shipping Bill Number" register={register} name="shippingBillNo" placeholder="SB-XXXXXXX" />
              <Field label="Shipping Bill Date" register={register} name="shippingBillDate" type="date" />
              <Field label="AWB Number" register={register} name="awbNo" placeholder="AWB-XXXXXXXXXX" />
              <Field label="AWB Date" register={register} name="awbDate" type="date" />
              <Field label="Insurance Policy Number" register={register} name="policyNo" placeholder="Policy number" />
              <Field label="Policy Date" register={register} name="policyDate" type="date" />
            </Section>

            <Section icon={<Plane size={16} />} title="Routing & Shipment Details" subtitle="Ports, destinations, and transit information" accent="#0ea5e9">
              <Field label="Pre-Carriage Mode" register={register} name="preCarriage" placeholder="By AIR" />
              <Field label="Place of Receipt" register={register} name="placeOfReceipt" placeholder="Origin Airport" />
              <Field label="Port of Loading" register={register} name="portOfLoading" placeholder="Origin Airport" />
              <Field label="Port of Discharge" register={register} name="portOfDischarge" placeholder="DESTINATION PORT" />
              <Field label="Final Destination" register={register} name="finalDestination" placeholder="DESTINATION COUNTRY" />
              <Field label="Vessel / Flight Number" register={register} name="vesselFlight" placeholder="Flight/vessel code" />
              <Field label="Flight / Departure Date" register={register} name="flightDate" type="date" />
              <FullWidth>
                <Field label="Terms of Delivery (Incoterms)" register={register} name="termsOfDelivery" placeholder="By AIR CIF DESTINATION" />
              </FullWidth>
              <FullWidth>
                <Field label="Payment Terms" register={register} name="paymentTerms" placeholder="100% ADVANCE WITH ORDER" />
              </FullWidth>
            </Section>
          </div>
        )}

        {/* ── STEP 4: PACKING ── */}
        {activeStep === 4 && (
          <div className="space-y-4">
            <Section icon={<Weight size={16} />} title="Weight Summary" subtitle="Total shipment weights and carton count" accent="#8b5cf6">
              <Field label="Total Gross Weight (KGS)" register={register} name="totalGrossWeight" placeholder="0.000" />
              <Field label="Total Net Weight (KGS)" register={register} name="totalNetWeight" placeholder="0.000" />
              <Field label="Total Corrugated Boxes" register={register} name="totalCorrugatedBoxes" placeholder="00" />
            </Section>

            {/* Box Dimensions */}
            <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
              <div className="flex items-center justify-between px-6 py-4" style={{ borderLeft: '4px solid #8b5cf6' }}>
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0" style={{ background: '#8b5cf618' }}>
                    <Box size={16} style={{ color: '#8b5cf6' }} />
                  </div>
                  <div>
                    <h3 className="text-[13px] font-bold text-slate-700">Box Dimensions</h3>
                    <p className="text-[11px] text-slate-400">Individual carton measurements (L × W × H)</p>
                  </div>
                  <span className="ml-2 px-2 py-0.5 rounded-full text-[10px] font-bold" style={{ background: '#8b5cf618', color: '#8b5cf6' }}>
                    {boxFields.length} {boxFields.length === 1 ? 'box' : 'boxes'}
                  </span>
                </div>
                <button
                  onClick={() => appendBox({ boxNo: `Box # ${String(boxFields.length + 1).padStart(2, '0')}`, dimensions: "" })}
                  className="flex items-center gap-1.5 px-3 py-2 rounded-xl bg-violet-50 border border-violet-200 text-violet-600 text-[11px] font-semibold hover:bg-violet-100 transition-colors"
                >
                  <Plus size={13} /> Add Box
                </button>
              </div>
              <div className="px-6 py-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                {boxFields.map((field, index) => (
                  <div key={field.id} className="flex items-center gap-2 p-3 bg-slate-50 rounded-xl border border-slate-100 hover:border-violet-200 transition-colors">
                    <Input
                      {...register(`boxDimensions.${index}.boxNo` as const)}
                      className="w-28 h-8 text-[11px] bg-white font-mono border-slate-200 rounded-lg shrink-0"
                    />
                    <Input
                      {...register(`boxDimensions.${index}.dimensions` as const)}
                      placeholder="L x W x H cms"
                      className="h-8 text-[11px] bg-white border-slate-200 rounded-lg flex-1 min-w-0"
                    />
                    <button
                      onClick={() => removeBox(index)}
                      className="w-7 h-7 flex items-center justify-center rounded-lg text-slate-300 hover:text-rose-500 hover:bg-rose-50 transition-colors shrink-0"
                    >
                      <Trash2 size={13} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── STEP 5: LINE ITEMS ── */}
        {activeStep === 5 && (
          <div className="space-y-4">

            {/* Summary strip */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              {[
                { label: "Total Items",    value: itemFields.length,                                                                   color: "#6366f1" },
                { label: "Total Qty",      value: watchedItems?.reduce((s, i) => s + (Number(i.quantity) || 0), 0).toLocaleString(),   color: "#0ea5e9" },
                { label: "Invoice Value",  value: `$${totalValue.toLocaleString('en-US', { minimumFractionDigits: 2 })}`,              color: "#10b981" },
                { label: "Invoice No.",    value: watchedInvoiceNo || "—",                                                             color: "#f59e0b" },
              ].map((stat) => (
                <div key={stat.label} className="bg-white rounded-xl border border-slate-100 px-4 py-3 shadow-sm">
                  <div className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">{stat.label}</div>
                  <div className="text-[18px] font-bold mt-0.5 font-mono truncate" style={{ color: stat.color }}>{stat.value}</div>
                </div>
              ))}
            </div>

            {/* Items table */}
            <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
              {/* Header */}
              <div className="flex items-center justify-between px-6 py-4" style={{ borderLeft: '4px solid #ef4444' }}>
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0" style={{ background: '#ef444418' }}>
                    <ClipboardList size={16} style={{ color: '#ef4444' }} />
                  </div>
                  <div>
                    <h3 className="text-[13px] font-bold text-slate-700">Line Items</h3>
                    <p className="text-[11px] text-slate-400">Click any row to expand all 20 fields</p>
                  </div>
                  <span className="ml-2 px-2 py-0.5 rounded-full text-[10px] font-bold" style={{ background: '#ef444418', color: '#ef4444' }}>
                    {itemFields.length} items
                  </span>
                </div>
                <button
                  onClick={() => {
                    appendItem({ productName: "", hsnSac: "", packSize: "", quantity: 0, price: 0, batchNo: "", mfgDate: "", expDate: "", boxInfo: "", grossWeight: 0, netWeight: 0, supplierGstin: "", stateCode: "", distCode: "", gstPercent: 0, uom: "", endUse: "", genericName: "", description: "" });
                    setExpandedRow(itemFields.length);
                  }}
                  className="flex items-center gap-1.5 px-4 py-2 rounded-xl bg-rose-600 text-white text-[12px] font-semibold hover:bg-rose-700 transition-colors shadow-md shadow-rose-100"
                >
                  <Plus size={14} /> Add Item
                </button>
              </div>

              {/* Column headers */}
              <div className="hidden sm:grid px-6 py-2 bg-slate-50 border-y border-slate-100 text-[10px] font-bold text-slate-400 uppercase tracking-wider gap-3"
                style={{ gridTemplateColumns: '1.5rem 2.5fr 1fr 1fr 1fr 1fr 1.2fr 1.8rem' }}>
                <span>#</span>
                <span>Product Name</span>
                <span>HSN</span>
                <span>Pack</span>
                <span>Qty</span>
                <span>Price</span>
                <span>Total</span>
                <span></span>
              </div>

              {/* Accordion rows */}
              <div className="divide-y divide-slate-50">
                {itemFields.map((field, index) => {
                  const item = watchedItems?.[index];
                  const lineTotal = (Number(item?.quantity) || 0) * (Number(item?.price) || 0);
                  const isOpen = expandedRow === index;

                  return (
                    <div key={field.id} className={`transition-colors duration-150 ${isOpen ? 'bg-indigo-50/30' : 'hover:bg-slate-50/70'}`}>

                      {/* Summary row */}
                      <div
                        className="grid px-6 py-3 cursor-pointer items-center gap-3"
                        style={{ gridTemplateColumns: '1.5rem 2.5fr 1fr 1fr 1fr 1fr 1.2fr 1.8rem' }}
                        onClick={() => setExpandedRow(isOpen ? null : index)}
                      >
                        <span className="text-[11px] font-bold text-slate-300">{index + 1}</span>
                        <div className="flex items-center gap-2 min-w-0">
                          <ChevronRight size={13} className={`text-slate-300 shrink-0 transition-transform duration-200 ${isOpen ? 'rotate-90' : ''}`} />
                          <span className="text-[12px] font-semibold text-slate-700 truncate">
                            {item?.productName || <span className="text-slate-300 font-normal italic">Unnamed product</span>}
                          </span>
                        </div>
                        <span className="text-[11px] font-mono text-slate-400 truncate hidden sm:block">{item?.hsnSac || '—'}</span>
                        <span className="text-[11px] text-slate-400 hidden sm:block">{item?.packSize || '—'}</span>
                        <span className="text-[12px] font-semibold text-slate-600 hidden sm:block">{Number(item?.quantity) || '—'}</span>
                        <span className="text-[11px] font-mono text-slate-400 hidden sm:block">${Number(item?.price || 0).toFixed(2)}</span>
                        <span className="text-[12px] font-bold font-mono" style={{ color: lineTotal > 0 ? '#6366f1' : '#cbd5e1' }}>
                          ${lineTotal.toFixed(2)}
                        </span>
                        <button
                          onClick={(e) => { e.stopPropagation(); removeItem(index); if (expandedRow === index) setExpandedRow(null); }}
                          className="w-7 h-7 flex items-center justify-center rounded-lg text-slate-200 hover:text-rose-500 hover:bg-rose-50 transition-colors"
                        >
                          <Trash2 size={13} />
                        </button>
                      </div>

                      {/* Expanded: ALL 20 FIELDS */}
                      {isOpen && (
                        <div className="mx-4 mb-4 p-5 rounded-xl bg-white border border-indigo-100 shadow-sm grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-x-5 gap-y-4">

                          {/* Group 1: Identity */}
                          <div className="col-span-1 sm:col-span-2 lg:col-span-2">
                            <Field label="Product Name" register={register} name={`items.${index}.productName` as Path<MasterData>} placeholder="PRODUCT NAME" required />
                          </div>
                          <div className="col-span-1 sm:col-span-2 lg:col-span-2">
                            <Field label="Description (Invoice)" register={register} name={`items.${index}.description` as Path<MasterData>} placeholder="ACTIVE INGREDIENT / FORMULATION DETAILS" />
                          </div>
                          <div className="col-span-1 sm:col-span-2 lg:col-span-2">
                            <Field label="Generic / Chemical Name" register={register} name={`items.${index}.genericName` as Path<MasterData>} placeholder="GENERIC NAME" />
                          </div>
                          <div className="col-span-1 sm:col-span-2 lg:col-span-2">
                            <Field label="End Use / Therapeutic Use" register={register} name={`items.${index}.endUse` as Path<MasterData>} placeholder="Therapeutic indication..." />
                          </div>

                          {/* Divider */}
                          <div className="col-span-1 sm:col-span-2 lg:col-span-4 border-t border-slate-100 pt-1">
                            <span className="text-[9px] font-bold text-slate-300 uppercase tracking-widest">Commercial Details</span>
                          </div>

                          {/* Group 2: Commercial */}
                          <Field label="HSN / SAC Code" register={register} name={`items.${index}.hsnSac` as Path<MasterData>} placeholder="30049099" />
                          <Field label="Pack Size" register={register} name={`items.${index}.packSize` as Path<MasterData>} placeholder="5ML" />
                          <Field label="Quantity" register={register} name={`items.${index}.quantity` as Path<MasterData>} type="number" placeholder="0" />
                          <Field label="Unit Price (USD)" register={register} name={`items.${index}.price` as Path<MasterData>} type="number" placeholder="0.00" />

                          {/* Group 3: Batch */}
                          <div className="col-span-1 sm:col-span-2 lg:col-span-4 border-t border-slate-100 pt-1">
                            <span className="text-[9px] font-bold text-slate-300 uppercase tracking-widest">Batch & Expiry</span>
                          </div>
                          <Field label="Batch Number" register={register} name={`items.${index}.batchNo` as Path<MasterData>} placeholder="BATCH-001" />
                          <Field label="Manufacturing Date" register={register} name={`items.${index}.mfgDate` as Path<MasterData>} type="date" />
                          <Field label="Expiry Date" register={register} name={`items.${index}.expDate` as Path<MasterData>} type="date" />
                          <Field label="Marks / Box Info" register={register} name={`items.${index}.boxInfo` as Path<MasterData>} placeholder="BOX # 01" />

                          {/* Group 4: Weight & UOM */}
                          <div className="col-span-1 sm:col-span-2 lg:col-span-4 border-t border-slate-100 pt-1">
                            <span className="text-[9px] font-bold text-slate-300 uppercase tracking-widest">Weight & Units</span>
                          </div>
                          <Field label="Gross Weight (KGS)" register={register} name={`items.${index}.grossWeight` as Path<MasterData>} type="number" placeholder="0.00" />
                          <Field label="Net Weight (KGS)" register={register} name={`items.${index}.netWeight` as Path<MasterData>} type="number" placeholder="0.00" />
                          <Field label="Unit of Measure" register={register} name={`items.${index}.uom` as Path<MasterData>} placeholder="KGS" />
                          <Field label="GST %" register={register} name={`items.${index}.gstPercent` as Path<MasterData>} type="number" placeholder="5" />

                          {/* Group 5: Tax & Supplier */}
                          <div className="col-span-1 sm:col-span-2 lg:col-span-4 border-t border-slate-100 pt-1">
                            <span className="text-[9px] font-bold text-slate-300 uppercase tracking-widest">Tax & Supplier</span>
                          </div>
                          <Field label="Supplier GSTIN" register={register} name={`items.${index}.supplierGstin` as Path<MasterData>} placeholder="00XXXXXXXXXXXXX" />
                          <Field label="State Code" register={register} name={`items.${index}.stateCode` as Path<MasterData>} placeholder="00" />
                          <Field label="District Code" register={register} name={`items.${index}.distCode` as Path<MasterData>} placeholder="DISTRICT NAME" />

                          {/* Computed line total */}
                          <div className="flex flex-col gap-1.5">
                            <label className="text-[10px] font-semibold tracking-widest text-slate-400 uppercase">Line Total</label>
                            <div className="h-9 flex items-center px-3 rounded-lg bg-indigo-50 border border-indigo-100 text-[14px] font-bold text-indigo-700 font-mono">
                              ${lineTotal.toFixed(2)}
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Footer */}
              <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2">
                <div className="flex items-center gap-2 text-[11px] text-slate-400">
                  <CheckCircle2 size={13} className="text-emerald-400 shrink-0" />
                  All 20 fields per item · Click any row to expand and edit
                </div>
                <div className="text-[15px] font-bold text-indigo-700 font-mono">
                  Grand Total: ${totalValue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ── NAVIGATION ── */}
        <div className="flex items-center justify-between pt-2 pb-8">
          <button
            onClick={() => setActiveStep(s => Math.max(0, s - 1))}
            disabled={activeStep === 0}
            className="flex items-center gap-2 px-5 py-2.5 rounded-xl border border-slate-200 bg-white text-slate-500 text-[13px] font-semibold hover:bg-slate-50 disabled:opacity-30 disabled:cursor-not-allowed transition-all shadow-sm"
          >
            ← Previous
          </button>

          {/* Dot indicators */}
          <div className="flex items-center gap-1.5">
            {steps.map((s) => (
              <button
                key={s.id}
                onClick={() => setActiveStep(s.id)}
                className={`rounded-full transition-all duration-300 ${
                  activeStep === s.id ? 'w-6 h-2 bg-indigo-500' : 'w-2 h-2 bg-slate-200 hover:bg-slate-300'
                }`}
              />
            ))}
          </div>

          {activeStep < steps.length - 1 ? (
            <button
              onClick={() => setActiveStep(s => Math.min(steps.length - 1, s + 1))}
              className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-indigo-600 text-white text-[13px] font-semibold hover:bg-indigo-700 transition-all shadow-md shadow-indigo-200"
            >
              Next →
            </button>
          ) : (
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <button
                  disabled={!!downloading}
                  className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-emerald-600 text-white text-[13px] font-semibold hover:bg-emerald-700 transition-all shadow-md shadow-emerald-200 disabled:opacity-60"
                >
                  {downloading
                    ? <Loader2 size={14} className="animate-spin" />
                    : <Download size={14} />}
                  Generate Invoice
                </button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-64 rounded-xl shadow-xl border-slate-100 p-1.5">
                <DropdownMenuLabel className="text-[11px] text-slate-400 font-medium px-2 pt-1 pb-0.5 uppercase tracking-wide">Complete Set</DropdownMenuLabel>
                <DropdownMenuItem onClick={handleSubmit(onDownloadCombinedBoth)} className="rounded-lg cursor-pointer font-semibold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 p-3 mb-0.5">
                  <Layers size={14} className="mr-2 shrink-0" /> All Sheets — XLSX + PDF
                  {downloading === 'complete-both' && <Loader2 size={12} className="ml-auto animate-spin" />}
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleSubmit(onDownloadCombined)} className="rounded-lg cursor-pointer p-2.5 mb-0.5 text-[13px]">
                  <FileSpreadsheet size={13} className="mr-2 shrink-0 text-emerald-600" /> All Sheets — XLSX only
                  {downloading === 'complete-xlsx' && <Loader2 size={12} className="ml-auto animate-spin" />}
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleSubmit(onDownloadCombinedPdf)} className="rounded-lg cursor-pointer p-2.5 text-[13px]">
                  <FileDown size={13} className="mr-2 shrink-0 text-rose-500" /> All Sheets — PDF only
                  {downloading === 'complete-pdf' && <Loader2 size={12} className="ml-auto animate-spin" />}
                </DropdownMenuItem>
                <DropdownMenuSeparator className="my-1.5" />
                <DropdownMenuLabel className="text-[11px] text-slate-400 font-medium px-2 pt-1 pb-0.5 uppercase tracking-wide">Individual</DropdownMenuLabel>
                <DropdownMenuItem onClick={handleSubmit(onDownloadCommercial)} className="rounded-lg cursor-pointer p-2.5 text-[13px]">
                  <FileSpreadsheet size={13} className="mr-2 shrink-0 text-emerald-600" /> Invoice — XLSX
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleSubmit(onDownloadCommercialPdf)} className="rounded-lg cursor-pointer p-2.5 text-[13px]">
                  <FileDown size={13} className="mr-2 shrink-0 text-rose-500" /> Invoice — PDF
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleSubmit(onDownloadPacking)} className="rounded-lg cursor-pointer p-2.5 text-[13px]">
                  <FileSpreadsheet size={13} className="mr-2 shrink-0 text-emerald-600" /> Packing List — XLSX
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleSubmit(onDownloadPackingPdf)} className="rounded-lg cursor-pointer p-2.5 text-[13px]">
                  <FileDown size={13} className="mr-2 shrink-0 text-rose-500" /> Packing List — PDF
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          )}
        </div>
      </main>
    </div>
  );
}