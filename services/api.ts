import { POData, ExportFormat } from '../types';
import { USE_MOCK_API, MOCK_DELAY_MS, SAMPLE_SUPPLIERS, SAMPLE_REQUESTORS } from '../constants';

type ExcelExportMode = 'new' | 'existing';

interface ExcelExportOptions {
  excelMode?: ExcelExportMode;
  mode?: ExcelExportMode;
}

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));
const arrayBufferToBase64 = (buffer: ArrayBuffer) => {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
};

const parseDateHelper = (dateStr: string | undefined): Date | null => {
  if (!dateStr) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const [y, m, d] = dateStr.split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
};

const formatDate = (dateStr: string | undefined): string => {
  try {
    const date = parseDateHelper(dateStr);
    if (!date) return dateStr || "-";
    const day = date.getDate();
    const month = date.toLocaleString('en-US', { month: 'short' });
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
  } catch (e) {
    return dateStr || "-";
  }
};

const getDaysDiff = (startStr: string | undefined, endStr: string | undefined): string => {
  try {
    if (!startStr || !endStr) return "-";
    const d1 = new Date(startStr);
    const d2 = new Date(endStr);
    if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return "-";
    const utc1 = Date.UTC(d1.getFullYear(), d1.getMonth(), d1.getDate());
    const utc2 = Date.UTC(d2.getFullYear(), d2.getMonth(), d2.getDate());
    return Math.floor((utc2 - utc1) / (1000 * 60 * 60 * 24)).toString();
  } catch (e) {
    return "-";
  }
};

const generateMockPOData = (fileName: string): POData => {
  const itemCount = Math.floor(Math.random() * 3) + 1;
  const items = Array.from({ length: itemCount }).map((_, i) => ({
    id: Math.random().toString(36).substr(2, 9),
    description: `1. Item Description ${i + 1}\nAdditional details line 1\nAdditional details line 2`,
    quantity: Math.floor(Math.random() * 20) + 1,
    rate: parseFloat((Math.random() * 100).toFixed(2)),
    amount: 0
  }));
  items.forEach(item => { item.amount = parseFloat((item.quantity * item.rate).toFixed(2)); });
  const subtotal = items.reduce((sum, item) => sum + item.amount, 0);
  const tax = parseFloat((subtotal * 0.1).toFixed(2));
  const total = subtotal + tax;
  const today = new Date();
  return {
    id: Math.random().toString(36).substr(2, 9),
    fileName,
    po_number: `PO-2024-${Math.floor(Math.random() * 1000).toString().padStart(3, '0')}`,
    supplier_name: SAMPLE_SUPPLIERS[Math.floor(Math.random() * SAMPLE_SUPPLIERS.length)],
    date: today.toISOString().split('T')[0],
    items,
    tax,
    total: parseFloat(total.toFixed(2)),
    status: 'completed',
    category: "IT",
    requestor_name: SAMPLE_REQUESTORS[Math.floor(Math.random() * SAMPLE_REQUESTORS.length)],
    pr_number: `PR-${Math.floor(Math.random() * 10000)}`,
    pr_date: today.toISOString().split('T')[0],
    description: `Supply of ${items[0].description}`,
    delivery_date_agreed: '',
    actual_delivery_date: '',
    agreed_vs_actual: '',
    remarks: '',
    no_of_days_remarks: '',
    no_of_days_delay: '',
    negotiation_remarks: ''
  };
};

export const api = {
  uploadAndExtract: async (file: File): Promise<POData> => {
    if (USE_MOCK_API) {
      throw new Error('Mock extraction is disabled. Set USE_MOCK_API=false to use the real Gemini API.');
    }
    try {
      const GEMINI_API_KEY = import.meta.env.VITE_GEMINI_API_KEY;
      if (!GEMINI_API_KEY) {
        console.warn("API_KEY not found. Falling back to mock data.");
        await sleep(1000);
        return generateMockPOData(file.name);
      }
      if (file.type !== 'application/pdf') throw new Error('Only PDF uploads are supported.');

      const buffer = await file.arrayBuffer();
      const base64 = arrayBufferToBase64(buffer);
      const response = await fetch('/api/po/extract', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileName: file.name, mimeType: file.type, base64 })
      });
      if (!response.ok) throw new Error((await response.text()) || 'PO extraction failed.');

      const payload = await response.json();
      if (!payload?.success || !payload?.data) throw new Error(payload?.error || 'PO extraction failed.');

      const result = payload.data;
      return {
        id: Math.random().toString(36).substr(2, 9),
        fileName: file.name,
        po_number: result.po_number || "",
        supplier_name: result.supplier_name || "",
        date: result.date || "",
        items: result.items?.flatMap((item: any, idx: number) => {
          const fullDesc = item.description || "";
          const splitDescriptions = fullDesc.split(/\n(?=\d+\.\s)/);
          if (splitDescriptions.length > 1) {
            return splitDescriptions.map((descPart: string, subIdx: number) => ({
              id: `${idx}-${subIdx}`, description: descPart.trim(),
              quantity: Number(item.quantity) || 0, rate: Number(item.rate) || 0, amount: Number(item.amount) || 0
            }));
          }
          return [{ id: `${idx}`, description: fullDesc, quantity: Number(item.quantity) || 0, rate: Number(item.rate) || 0, amount: Number(item.amount) || 0 }];
        }) || [],
        tax: Number(result.tax) || 0,
        total: Number(result.total) || 0,
        status: 'completed',
        category: result.category || "",
        requestor_name: result.requestor_name || "",
        pr_number: result.pr_number || "",
        pr_date: result.pr_date || "",
        description: result.description || "",
        delivery_date_agreed: result.delivery_date_agreed || "",
        actual_delivery_date: result.actual_delivery_date || "",
        agreed_vs_actual: result.agreed_vs_actual || "",
        remarks: result.remarks || "",
        no_of_days_remarks: result.no_of_days_remarks || "",
        no_of_days_delay: result.no_of_days_delay || "",
        negotiation_remarks: result.negotiation_remarks || ""
      };
    } catch (error: any) {
      console.error("Extraction Error:", error);
      let errorMsg = 'Failed to extract data.';
      if (error?.status === 429 || (error?.message && (error.message.includes('429') || error.message.toLowerCase().includes('quota')))) {
        errorMsg = 'API Quota Exceeded. Please try again later.';
      } else if (error?.message) {
        errorMsg = error.message;
      }
      return {
        id: Math.random().toString(36), fileName: file.name, po_number: 'Error', supplier_name: 'Error',
        date: '', items: [], tax: 0, total: 0, status: 'error', errorMessage: errorMsg,
        category: '', requestor_name: '', pr_number: '', pr_date: '', description: ''
      };
    }
  },

  generateExport: async (
    data: POData[],
    _templateType: 'default' | 'custom',
    format: ExportFormat | 'docx' | 'jpeg' | 'jpg',
    _customTemplate?: File,
    excelOptions?: ExcelExportOptions
  ): Promise<Blob> => {
    await sleep(500);

    const flattenedRows: any[] = [];
    let slNoCounter = 1;
    const isExcel = format === 'xlsx';
    const excelMode = excelOptions?.excelMode ?? excelOptions?.mode ?? 'new';

    data.forEach(po => {
      const currentSlNo = slNoCounter++;
      const daysPrToPo = getDaysDiff(po.pr_date, po.date);

      const processDate = (val: string | undefined): Date | string => {
        if (isExcel) { const d = parseDateHelper(val); return d || (val || "-"); }
        return formatDate(val);
      };

      const fullMetaData = {
        slNo: currentSlNo, category: "-", requestor: po.requestor_name || "",
        prNumber: po.pr_number || "", prDate: processDate(po.pr_date),
        poNumber: po.po_number || "", poDate: processDate(po.date),
        vendorName: po.supplier_name || "", tax: po.tax || 0, grandTotal: po.total || 0,
        status: po.status || "", deliveryDateAgreed: "-", actualDeliveryDate: "-",
        agreedVsActual: "-", remarks: "-", daysPrToPo,
        remarksPrToPo: "-", daysPoToDelivery: "-", remarksPoToDelivery: "-",
        negotiationSavings: "-", remarksNegotiation: "-"
      };

      const emptyMetaData = {
        slNo: "", category: "", requestor: "", prNumber: "", prDate: "",
        poNumber: "", poDate: "", vendorName: "", tax: "", grandTotal: "",
        status: "", deliveryDateAgreed: "", actualDeliveryDate: "", agreedVsActual: "",
        remarks: "", daysPrToPo: "", remarksPrToPo: "", daysPoToDelivery: "",
        remarksPoToDelivery: "", negotiationSavings: "", remarksNegotiation: ""
      };

      if (po.items && po.items.length > 0) {
        po.items.forEach((item, index) => {
          const meta = index === 0 ? fullMetaData : emptyMetaData;
          flattenedRows.push({ ...meta, description: item.description || "", qty: item.quantity || 0, unitRate: item.rate || 0 });
        });
      } else {
        flattenedRows.push({ ...fullMetaData, description: po.description || "Purchase Order", qty: 0, unitRate: 0 });
      }
    });

    // ✅ Original 24-column format
    const headers = [
      "Sl No", "Category", "Requestor Name", "PR Number", "PR Date",
      "Purchase Order No", "PO Date", "Vendor Name", "Description",
      "Qty", "Unit Rate", "Tax", "Grand Total", "Status",
      "Delivery Date agreed as per PO", "Actual Delivery Date",
      "Agreed vs Actual Delivery Date", "Remarks",
      "No of days from PR to PO", "Remarks",
      "No of Days from PO to Delivery", "Remarks",
      "Negotiation Savings", "Remarks"
    ];

    // ✅ 24 values matching exactly
    const mapRowToArray = (r: any) => [
      r.slNo, r.category, r.requestor, r.prNumber, r.prDate,
      r.poNumber, r.poDate, r.vendorName, r.description,
      r.qty, r.unitRate, r.tax, r.grandTotal, r.status,
      r.deliveryDateAgreed, r.actualDeliveryDate, r.agreedVsActual, r.remarks,
      r.daysPrToPo, r.remarksPrToPo,
      r.daysPoToDelivery, r.remarksPoToDelivery,
      r.negotiationSavings, r.remarksNegotiation
    ];

    if (format === 'csv') {
      const rowStrings = flattenedRows.map(r => mapRowToArray(r).map(val => `"${val}"`).join(","));
      return new Blob([headers.join(",") + "\n" + rowStrings.join("\n")], { type: 'text/csv' });
    }
    if (format === 'json') {
      return new Blob([JSON.stringify(flattenedRows, null, 2)], { type: 'application/json' });
    }
    if (format === 'docx') {
      const htmlContent = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><title>PO Export</title></head><body>
        <h1>Purchase Order Tracker</h1>
        <table border="1" style="border-collapse:collapse;width:100%;font-family:Arial;font-size:8pt;">
          <thead><tr style="background-color:#f0f0f0;">${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
          <tbody>${flattenedRows.map(r => `<tr>${mapRowToArray(r).map(val => `<td>${val}</td>`).join('')}</tr>`).join('')}</tbody>
        </table></body></html>`;
      return new Blob([htmlContent], { type: 'application/msword' });
    }
    if (format === 'pdf') {
      try {
        const response = await fetch('/api/pdf/export', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ headers, rows: flattenedRows.map(mapRowToArray) })
        });
        if (!response.ok) throw new Error((await response.text()) || 'PDF export failed.');
        const contentType = response.headers.get('content-type') || '';
        if (!contentType.includes('application/pdf')) {
          throw new Error(`Unexpected content-type: ${contentType}`);
        }
        return await response.blob();
      } catch (error) {
        console.error('Backend PDF generation failed:', error);
        throw error;
      }
    }
    if (format === 'jpeg' || format === 'jpg') {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error("Canvas context not supported");
      const lineHeight = 20;
      const width = 1800;
      const height = (flattenedRows.length + 5) * lineHeight;
      canvas.width = width; canvas.height = height;
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, width, height);
      ctx.fillStyle = '#000000'; ctx.font = '12px monospace';
      let y = 30;
      ctx.fillText("PO TRACKER EXPORT", 20, y); y += 30;
      ctx.fillText("Sl | PO No      | Vendor          | Description                    | Qty | Rate  | Total  | Delivery", 20, y); y += 10;
      ctx.fillText("-".repeat(200), 20, y); y += 20;
      flattenedRows.forEach(row => {
        const sl = row.slNo ? row.slNo.toString() : "";
        const line = `${sl.padEnd(3)}| ${row.poNumber.padEnd(11)}| ${row.vendorName.slice(0,15).padEnd(16)}| ${row.description.slice(0,30).padEnd(31)}| ${row.qty.toString().padEnd(4)}| ${row.unitRate.toString().padEnd(6)}| ${row.grandTotal.toString().padEnd(7)}| ${row.deliveryDateAgreed}`;
        ctx.fillText(line, 20, y); y += lineHeight;
      });
      return new Promise((resolve) => { canvas.toBlob((blob) => { resolve(blob || new Blob([])); }, 'image/jpeg', 0.9); });
    }

    // Excel — send to backend
    try {
      const response = await fetch('/api/excel/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ excelMode, mode: excelMode, headers, rows: flattenedRows.map(mapRowToArray) })
      });
      if (!response.ok) throw new Error((await response.text()) || 'Excel export failed.');
      const contentType = response.headers.get('content-type') || '';
      if (!contentType.includes('application/json') && !contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
        throw new Error(`Unexpected content-type: ${contentType}`);
      }
      return await response.blob();
    } catch (error) {
      console.error('Backend Excel generation failed:', error);
      throw error;
    }
  },

  downloadMasterExcel: async (): Promise<Blob> => {
    const response = await fetch('/api/download-excel');
    if (!response.ok) {
      let message = response.statusText;
      try {
        const payload = await response.json();
        message = payload?.error || message;
      } catch {
        // ignore json parse errors
      }
      throw new Error(message || 'Failed to download master Excel.');
    }
    return response.blob();
  },

  uploadMasterExcel: async (file: File): Promise<{ success: boolean; message?: string }> => {
    const formData = new FormData();
    formData.append('file', file);
    const response = await fetch('/api/upload-excel', { method: 'POST', body: formData });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || !payload?.success) {
      const errorMsg = payload?.error || response.statusText || 'Failed to upload Excel.';
      throw new Error(errorMsg);
    }
    return payload;
  },

  generateExcelDownloadUrl: async (
    data: POData[],
    templateType: 'default' | 'custom',
    customTemplate?: File,
    excelOptions?: ExcelExportOptions
  ): Promise<string> => {
    const resultBlob = await api.generateExport(data, templateType, 'xlsx', customTemplate, excelOptions);
    const resultText = await resultBlob.text();
    let parsed: { success?: boolean; downloadUrl?: string; error?: string };
    try { parsed = JSON.parse(resultText); }
    catch { throw new Error('Excel generation did not return a valid download link response.'); }
    if (!parsed.success || !parsed.downloadUrl) throw new Error(parsed.error || 'Link failed to generate.');
    return parsed.downloadUrl;
  }
};
