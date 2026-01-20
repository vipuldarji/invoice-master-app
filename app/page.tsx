"use client";

import React from 'react';
import { useForm, useFieldArray, SubmitHandler, UseFormRegister, Path } from 'react-hook-form';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Separator } from '@/components/ui/separator';
import { Download, Package, Plus, Trash2, Box, Plane, FileBadge, DollarSign, ChevronDown, FileText, Layers } from 'lucide-react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";

// Imported Helper Functions for Excel
import { generateMasterExcel, MasterData, addMasterSheet } from '@/utils/excelGenerator';
import { generateCommercialInvoice, addCommercialInvoiceSheet } from '@/utils/generators/commercialInvoice';

// --- HELPER COMPONENT PROPS INTERFACE ---
interface ExcelRowProps {
  label: string;
  register: UseFormRegister<MasterData>;
  name: Path<MasterData>; 
  type?: string;
  placeholder?: string;
  className?: string;
}

// 1. Side-by-Side (Excel Row Style)
const ExcelRow: React.FC<ExcelRowProps> = ({ label, register, name, type = "text", placeholder = "", className = "" }) => (
  <div className={`grid grid-cols-12 gap-2 items-center ${className}`}>
    <span className="col-span-4 text-[10px] font-bold text-slate-500 uppercase text-right">{label}:</span>
    <div className="col-span-8">
      <Input type={type} {...register(name)} placeholder={placeholder} className="h-7 text-xs bg-white" />
    </div>
  </div>
);

// 2. Stacked (Standard Form Style)
const StackedField: React.FC<ExcelRowProps> = ({ label, register, name, placeholder = "", className = "" }) => (
  <div className={`flex flex-col gap-1 ${className}`}>
    <span className="text-[10px] font-bold text-slate-500 uppercase ml-1">{label}</span>
    <Input {...register(name)} placeholder={placeholder} className="h-8 text-xs bg-white" />
  </div>
);

export default function MasterInvoiceApp() {
  const { register, control, handleSubmit, watch } = useForm<MasterData>({
    defaultValues: {
      // --- PARTIES (Env vars for security) ---
      exporterName: process.env.NEXT_PUBLIC_EXPORTER_NAME || "",
      exporterAddress: process.env.NEXT_PUBLIC_EXPORTER_ADDRESS || "",
      exporterPhone: "",
      exporterEmail: "",
      exporterRef: "",
      consigneeName: "",
      consigneeAddress: "",
      buyerName: "",
      buyerOrderRef: "",
      chaName: "",

      // --- REGULATORY ---
      iecNo: "",
      gstStatus: "",
      companyGstNo: "",
      drugLicNo: "",
      lutRef: "",
      lutDate: "", 
      
      // --- REMITTANCE ---
      remittanceRef: "",
      remittanceDate: "",
      remittanceAmount: "",
      remittanceAvailable: "",
      remittanceUsed: "",

      // --- FINANCIALS ---
      proformaValue: "",
      invoiceValue110: "",
      invoiceValue110Round: "",
      adcRate: "", 
      exchangeRate: 0,
      inrValue: "",
      freightValue: 0,
      insuranceValue: 0,
      currency: "USD",
      uom: "KGS",

      // --- LOGISTICS ---
      invoiceNo: "",
      invoiceDate: new Date().toISOString().split('T')[0],
      packingListNo: "",
      placeOfReceipt: "",
      portOfLoading: "",
      portOfDischarge: "",
      finalDestination: "",
      preCarriage: "By AIR",
      vesselFlight: "",
      flightDate: "",
      paymentTerms: "",
      termsOfDelivery: "",
      shippingBillNo: "",
      shippingBillDate: "",
      awbNo: "",
      awbDate: "",
      policyNo: "",
      policyDate: "",

      // --- PACKING ---
      totalGrossWeight: "",
      totalNetWeight: "",
      totalCorrugatedBoxes: "",
      generalDescription: "",
      globalIgst: "",

      // --- MANUFACTURER ---
      manufacturerName: "",
      manufacturerAddress: "",

      // --- ARRAYS ---
      items: [{ 
        productName: "", hsnSac: "", packSize: "", quantity: 0, price: 0, 
        batchNo: "", mfgDate: "", expDate: "", boxInfo: "", 
        grossWeight: 0, netWeight: 0, supplierGstin: "", stateCode: "",
        distCode: "", gstPercent: 0, uom: "", endUse: "",
        genericName: "", description: "" 
      }],
      boxDimensions: [
        { boxNo: "Box # 01", dimensions: "" }
      ]
    }
  });

  const { fields: itemFields, append: appendItem, remove: removeItem } = useFieldArray({
    control,
    name: "items"
  });

  const { fields: boxFields, append: appendBox, remove: removeBox } = useFieldArray({
    control,
    name: "boxDimensions"
  });

  const watchedItems = watch("items");
  const totalValue = watchedItems?.reduce((sum, item) => sum + ((Number(item.quantity) || 0) * (Number(item.price) || 0)), 0) || 0;

  // --- 1. DOWNLOAD MASTER SHEET ONLY ---
  const onDownloadMaster: SubmitHandler<MasterData> = async (data) => {
    try { await generateMasterExcel(data); } 
    catch (error) { console.error(error); alert("Failed."); }
  };

  // --- 2. DOWNLOAD COMMERCIAL INVOICE ONLY ---
  const onDownloadCommercial: SubmitHandler<MasterData> = async (data) => {
    try { await generateCommercialInvoice(data); } 
    catch (error) { console.error(error); alert("Failed."); }
  };

  // --- 3. DOWNLOAD COMPLETE SET (ALL TABS) ---
  const onDownloadCombined: SubmitHandler<MasterData> = async (data) => {
    try {
      // Create Empty Workbook
      const workbook = new ExcelJS.Workbook();
      
      // Add Sheet 1: Master Data
      addMasterSheet(workbook, data);
      
      // Add Sheet 2: Commercial Invoice
      addCommercialInvoiceSheet(workbook, data);
      
      // Save File
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `Complete_Set_${data.invoiceNo || 'DRAFT'}.xlsx`);
      
    } catch (error) {
      console.error(error);
      alert("Failed to generate Combined Excel.");
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 pb-20 font-sans text-xs md:text-sm">
      
      {/* HEADER */}
      <header className="sticky top-0 z-50 bg-white/90 backdrop-blur border-b px-6 py-3 flex justify-between items-center shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-blue-700 text-white p-2 rounded">
            <Package size={18} />
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-tight text-slate-800">Master Invoice Engine</h1>
            <p className="text-[10px] text-slate-500 font-bold uppercase">Dynamic Build • Ver 2026</p>
          </div>
        </div>
        <div className="flex gap-4 items-center">
           <div className="text-right hidden md:block">
              <div className="text-[10px] text-slate-400 uppercase font-bold">Total Invoice Value</div>
              <div className="text-xl font-mono font-bold text-blue-600">${totalValue.toFixed(2)}</div>
           </div>
           
           {/* DOWNLOAD MENU */}
           <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <Button className="bg-blue-700 hover:bg-blue-800 text-white shadow-lg">
                <Download className="w-4 h-4 mr-2" /> Download <ChevronDown className="w-4 h-4 ml-2"/>
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="w-64">
              
              {/* COMBINED BUTTON */}
              <DropdownMenuItem onClick={handleSubmit(onDownloadCombined)} className="cursor-pointer bg-blue-50 text-blue-700 font-bold focus:bg-blue-100">
                <Layers className="w-4 h-4 mr-2" /> Download Complete Set
              </DropdownMenuItem>
              
              <Separator className="my-1"/>
              
              {/* INDIVIDUAL BUTTONS */}
              <DropdownMenuItem onClick={handleSubmit(onDownloadMaster)} className="cursor-pointer">
                <FileText className="w-4 h-4 mr-2" /> Master Data Sheet Only
              </DropdownMenuItem>
              <DropdownMenuItem onClick={handleSubmit(onDownloadCommercial)} className="cursor-pointer">
                <FileBadge className="w-4 h-4 mr-2" /> Commercial Invoice Only
              </DropdownMenuItem>
              
              <DropdownMenuItem disabled>
                <span className="opacity-50">Packing List (Coming Soon)</span>
              </DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>

        </div>
      </header>

      <main className="max-w-[1800px] mx-auto p-4 space-y-4">
        
        {/* ROW 1: PARTIES & REGULATORY */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
          
          {/* 1. EXPORTER & FINANCIALS */}
          <div className="lg:col-span-4 space-y-4">
             <Card>
               <CardHeader className="py-2 px-4 bg-slate-100 border-b"><CardTitle className="text-xs font-bold uppercase text-slate-600">Exporter & Regulatory</CardTitle></CardHeader>
               <CardContent className="p-4 space-y-3">
                  <StackedField label="Exporter Name" register={register} name="exporterName" />
                  <StackedField label="Exporter Address" register={register} name="exporterAddress" />
                  
                  <div className="grid grid-cols-2 gap-2">
                    <StackedField label="Phone" register={register} name="exporterPhone" />
                    <StackedField label="Email" register={register} name="exporterEmail" />
                  </div>
                  
                  <Separator className="my-2" />
                  
                  {/* Regulatory Block */}
                  <div className="space-y-2 bg-slate-50 p-2 rounded border">
                    <ExcelRow label="IEC No" register={register} name="iecNo" />
                    <ExcelRow label="Co. GSTN" register={register} name="companyGstNo" />
                    <ExcelRow label="Drug Lic" register={register} name="drugLicNo" />
                    <div className="grid grid-cols-2 gap-2">
                      <StackedField label="LUT Ref" register={register} name="lutRef" />
                      <StackedField label="LUT Date" register={register} name="lutDate" placeholder="dd-mm-yyyy" />
                    </div>
                    <ExcelRow label="Exp Ref" register={register} name="exporterRef" />
                    <ExcelRow label="Status" register={register} name="gstStatus" />
                  </div>

                  {/* Financials Block */}
                  <div className="bg-yellow-50 p-3 rounded border border-yellow-100 space-y-2 mt-2">
                     <div className="flex items-center gap-2 mb-2">
                        <DollarSign className="w-3 h-3 text-yellow-600"/>
                        <span className="text-[10px] font-bold text-yellow-700 uppercase">Remittance & Value</span>
                     </div>
                     <ExcelRow label="TT Ref" register={register} name="remittanceRef" />
                     <ExcelRow label="TT Date" register={register} name="remittanceDate" type="date" />
                     <ExcelRow label="TT Amount" register={register} name="remittanceAmount" />
                     <ExcelRow label="Available" register={register} name="remittanceAvailable" />
                     <ExcelRow label="To Use" register={register} name="remittanceUsed" />
                     <Separator className="bg-yellow-200 my-2" />
                     <ExcelRow label="Proforma" register={register} name="proformaValue" />
                     <ExcelRow label="110% Value" register={register} name="invoiceValue110" />
                     <ExcelRow label="110% Round" register={register} name="invoiceValue110Round" />
                     <ExcelRow label="ADC Rate" register={register} name="adcRate" />
                     <ExcelRow label="Exch Rate" register={register} name="exchangeRate" type="number" />
                     <ExcelRow label="INR Value" register={register} name="inrValue" />
                  </div>
               </CardContent>
             </Card>
          </div>

          {/* 2. CONSIGNEE & BUYER */}
          <div className="lg:col-span-4 space-y-4">
             <Card className="h-full">
               <CardHeader className="py-2 px-4 bg-slate-100 border-b"><CardTitle className="text-xs font-bold uppercase text-slate-600">Consignee & Manufacturer</CardTitle></CardHeader>
               <CardContent className="p-4 space-y-4">
                  <StackedField label="Consignee Name" register={register} name="consigneeName" />
                  <StackedField label="Consignee Address" register={register} name="consigneeAddress" />
                  
                  <Separator />
                  
                  <StackedField label="Buyer (If different)" register={register} name="buyerName" />
                  <ExcelRow label="Buyer Order Ref" register={register} name="buyerOrderRef" />
                  <ExcelRow label="CHA Name" register={register} name="chaName" />
                  
                  <div className="bg-orange-50 p-3 rounded border border-orange-100 mt-2 space-y-2">
                     <div className="text-[10px] font-bold text-orange-400 uppercase">Manufacturer Details</div>
                     <StackedField label="Name" register={register} name="manufacturerName" />
                     <StackedField label="Address" register={register} name="manufacturerAddress" />
                  </div>
                  
                  <div className="bg-blue-50 p-3 rounded border border-blue-100 mt-2 space-y-2">
                     <div className="text-[10px] font-bold text-blue-400 uppercase">Description & Tax</div>
                     <StackedField label="General Description" register={register} name="generalDescription" />
                     <ExcelRow label="Global IGST" register={register} name="globalIgst" />
                  </div>
               </CardContent>
             </Card>
          </div>

          {/* 3. LOGISTICS & DOCS */}
          <div className="lg:col-span-4 space-y-4">
             <Card className="h-full">
               <CardHeader className="py-2 px-4 bg-slate-100 border-b flex justify-between">
                  <CardTitle className="text-xs font-bold uppercase text-slate-600">Logistics & Shipping Docs</CardTitle>
                  <Plane className="w-4 h-4 text-slate-400" />
               </CardHeader>
               <CardContent className="p-4 space-y-2">
                  <ExcelRow label="Invoice No" register={register} name="invoiceNo" className="bg-blue-50 p-1 rounded" />
                  <ExcelRow label="Invoice Date" register={register} name="invoiceDate" type="date" />
                  <ExcelRow label="Packing List" register={register} name="packingListNo" />
                  <Separator className="my-1"/>
                  <ExcelRow label="Pre-Carriage" register={register} name="preCarriage" />
                  <ExcelRow label="Receipt Place" register={register} name="placeOfReceipt" />
                  <ExcelRow label="Port Loading" register={register} name="portOfLoading" />
                  <ExcelRow label="Port Discharge" register={register} name="portOfDischarge" />
                  <ExcelRow label="Final Dest" register={register} name="finalDestination" />
                  <Separator className="my-1"/>
                  <ExcelRow label="Vessel/Flight" register={register} name="vesselFlight" />
                  <ExcelRow label="Flight Date" register={register} name="flightDate" type="date" />
                  <StackedField label="Terms of Delivery" register={register} name="termsOfDelivery" />
                  <StackedField label="Payment Terms" register={register} name="paymentTerms" />
                  
                  <div className="bg-slate-50 p-2 rounded border mt-2 space-y-2">
                    <span className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Shipping Documents</span>
                    <ExcelRow label="SB No" register={register} name="shippingBillNo" />
                    <ExcelRow label="SB Date" register={register} name="shippingBillDate" type="date" />
                    <ExcelRow label="AWB No" register={register} name="awbNo" />
                    <ExcelRow label="AWB Date" register={register} name="awbDate" type="date" />
                    <ExcelRow label="Policy No" register={register} name="policyNo" />
                    <ExcelRow label="Policy Date" register={register} name="policyDate" type="date" />
                  </div>
                  
                  <div className="grid grid-cols-2 gap-2 mt-2">
                    <StackedField label="Freight Val" register={register} name="freightValue" />
                    <StackedField label="Insurance" register={register} name="insuranceValue" />
                  </div>
               </CardContent>
             </Card>
          </div>
        </div>

        {/* ROW 2: PACKING & TOTALS */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
           {/* BOX DIMENSIONS */}
           <div className="lg:col-span-8">
              <Card>
                 <CardHeader className="py-2 px-4 bg-yellow-50 border-b border-yellow-100 flex justify-between items-center">
                    <CardTitle className="text-xs font-bold uppercase text-yellow-700 flex gap-2 items-center">
                       <Box className="w-4 h-4" /> Packing Dimensions
                    </CardTitle>
                    <Button size="sm" variant="outline" onClick={() => appendBox({ boxNo: `Box # 0${boxFields.length + 1}`, dimensions: "" })} className="h-6 text-xs bg-white">
                       <Plus className="w-3 h-3 mr-1" /> Add Box
                    </Button>
                 </CardHeader>
                 <CardContent className="p-4">
                    <div className="grid grid-cols-3 gap-4">
                       {boxFields.map((field, index) => (
                          <div key={field.id} className="flex gap-2 items-center">
                             <Input {...register(`boxDimensions.${index}.boxNo` as const)} className="w-24 bg-slate-50 font-mono text-xs" />
                             <Input {...register(`boxDimensions.${index}.dimensions` as const)} placeholder="L x W x H cms" className="text-xs" />
                             <Button variant="ghost" size="icon" onClick={() => removeBox(index)} className="h-6 w-6 text-slate-300 hover:text-red-400"><Trash2 className="w-3 h-3"/></Button>
                          </div>
                       ))}
                    </div>
                 </CardContent>
              </Card>
           </div>

           {/* WEIGHT TOTALS */}
           <div className="lg:col-span-4">
              <Card className="h-full">
                 <CardHeader className="py-2 px-4 bg-slate-100 border-b"><CardTitle className="text-xs font-bold uppercase text-slate-600">Weight Summary</CardTitle></CardHeader>
                 <CardContent className="p-4 space-y-2">
                    <ExcelRow label="Total Gross Wt" register={register} name="totalGrossWeight" />
                    <ExcelRow label="Total Net Wt" register={register} name="totalNetWeight" />
                    <ExcelRow label="Total Boxes" register={register} name="totalCorrugatedBoxes" />
                 </CardContent>
              </Card>
           </div>
        </div>

        {/* ROW 3: ITEMS (100% COMPLETE MATCH) */}
        <Card className="shadow-lg border-slate-300 overflow-hidden flex flex-col">
          <CardHeader className="flex flex-row items-center justify-between bg-slate-800 text-white py-2 px-4">
             <div className="flex items-center gap-2">
                <FileBadge className="w-4 h-4"/>
                <CardTitle className="text-sm font-bold uppercase">Line Items ({itemFields.length})</CardTitle>
             </div>
             <Button size="sm" onClick={() => appendItem({ 
                productName: "", hsnSac: "", packSize: "", quantity: 0, price: 0, 
                batchNo: "", mfgDate: "", expDate: "", boxInfo: "", 
                grossWeight: 0, netWeight: 0, supplierGstin: "", stateCode: "",
                distCode: "", gstPercent: 0, uom: "", endUse: "",
                genericName: "", description: ""
              })} className="bg-blue-500 hover:bg-blue-400 text-white h-7 text-xs border-0">
                <Plus className="w-3 h-3 mr-2" /> Add Item
             </Button>
          </CardHeader>
          
          <CardContent className="p-0 overflow-x-auto">
             <table className="w-full text-left text-xs whitespace-nowrap">
               <thead className="bg-slate-100 text-slate-600 font-bold uppercase sticky top-0 z-10 border-b border-slate-200">
                 <tr>
                   <th className="px-2 py-2 w-8">#</th>
                   <th className="px-2 py-2 min-w-[150px]">Product Name</th>
                   <th className="px-2 py-2 w-20">HSN</th>
                   <th className="px-2 py-2 w-16">Pack</th>
                   <th className="px-2 py-2 w-16">Qty</th>
                   <th className="px-2 py-2 w-16">Price</th>
                   <th className="px-2 py-2 w-24">Batch</th>
                   <th className="px-2 py-2 w-24">Mfg Date</th>
                   <th className="px-2 py-2 w-24">Exp Date</th>
                   <th className="px-2 py-2 w-24">Marks/Nos</th>
                   <th className="px-2 py-2 w-16">State</th>
                   <th className="px-2 py-2 w-20">Supp GST</th>
                   <th className="px-2 py-2 w-20">Dist Code</th>
                   <th className="px-2 py-2 w-16">Gr Wt</th>
                   <th className="px-2 py-2 w-16">Net Wt</th>
                   <th className="px-2 py-2 w-16">UOM</th>
                   <th className="px-2 py-2 w-16">GST %</th>
                   <th className="px-2 py-2 w-24">Desc.</th>
                   <th className="px-2 py-2 w-24">Generic</th>
                   <th className="px-2 py-2 w-24">End Use</th>
                   <th className="px-2 py-2 w-8"></th>
                 </tr>
               </thead>
               <tbody className="divide-y divide-slate-100">
                 {itemFields.map((field, index) => (
                   <tr key={field.id} className="group hover:bg-blue-50 transition-colors">
                     <td className="px-2 py-1 text-slate-400 font-mono">{index + 1}</td>
                     
                     <td className="px-1 py-1"><Input {...register(`items.${index}.productName` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" placeholder="Name" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.hsnSac` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.packSize` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input type="number" {...register(`items.${index}.quantity` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white font-bold text-blue-700" /></td>
                     <td className="px-1 py-1"><Input type="number" step="0.01" {...register(`items.${index}.price` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.batchNo` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white bg-yellow-50/50" /></td>
                     <td className="px-1 py-1"><Input type="date" {...register(`items.${index}.mfgDate` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input type="date" {...register(`items.${index}.expDate` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.boxInfo` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white bg-yellow-50/50" /></td>
                     
                     <td className="px-1 py-1"><Input {...register(`items.${index}.stateCode` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.supplierGstin` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.distCode` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input type="number" step="0.01" {...register(`items.${index}.grossWeight` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input type="number" step="0.01" {...register(`items.${index}.netWeight` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.uom` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.gstPercent` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.description` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.genericName` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>
                     <td className="px-1 py-1"><Input {...register(`items.${index}.endUse` as const)} className="h-7 text-xs border-transparent focus:border-blue-300 bg-transparent focus:bg-white" /></td>

                     <td className="px-1 py-1 text-center">
                       <Button variant="ghost" size="icon" onClick={() => removeItem(index)} className="h-6 w-6 text-slate-300 hover:text-red-600">
                         <Trash2 className="w-3 h-3" />
                       </Button>
                     </td>
                   </tr>
                 ))}
               </tbody>
             </table>
          </CardContent>
          <div className="bg-slate-50 border-t p-1 text-[10px] text-center text-slate-400">
             Grid now matches 100% of the Invoice columns. Scroll right to edit all fields.
          </div>
        </Card>
      </main>
    </div>
  );
}