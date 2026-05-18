import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import multer from 'multer';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import { GoogleGenAI, Type } from '@google/genai';
import PDFDocument from 'pdfkit';

const PORT = Number(process.env.PORT || 3001);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const STORAGE_DIR = path.resolve(__dirname, '..', 'storage');
const MASTER_FILE_PATH = path.join(STORAGE_DIR, 'po_master.xlsx');
const SHEET_NAME = 'PO Tracker';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/legacy/build/pdf.worker.mjs',
  import.meta.url
).toString();

// ─── Utilities ────────────────────────────────────────────────────────────────

const normalizeHeader = (value) => String(value ?? '').trim().toLowerCase();

const headersMatch = (candidateRow, expectedHeaders) =>
  expectedHeaders.every((header, index) => normalizeHeader(candidateRow[index]) === normalizeHeader(header));

const isRowEmpty = (row) => !Array.isArray(row) || row.every((cell) => String(cell ?? '').trim() === '');


const normalizeCell = (value) => {
  if (value instanceof Date) return value.toISOString();
  return String(value ?? '').trim();
};

const normalizeWhitespace = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();

const normalizeKeyText = (value) => normalizeWhitespace(value).toUpperCase();

const normalizeAmountKey = (value) => {
  const raw = normalizeCell(value).replace(/,/g, '');
  const parsed = Number(raw);
  if (Number.isFinite(parsed)) return parsed.toFixed(2);
  return normalizeKeyText(raw);
};

const normalizePrNumber = (value) => {
  const raw = normalizeCell(value);
  if (!raw) return '';

  const compact = raw.replace(/\s+/g, '').toUpperCase();
  const normalizedSeparators = compact.replace(/[^A-Z0-9]+/g, '/');
  const withPrefix = normalizedSeparators.replace(/^PR(?!\/)/, 'PR/');
  const cleaned = withPrefix.replace(/\/+/g, '/');
  const match = cleaned.match(/^PR\/(\d{4})\/(\d+)$/);

  const normalized = match
    ? `PR/${match[1]}/${match[2].padStart(3, '0')}`
    : '';

  if (raw) {
    console.log(`[normalize] PR Number - Original: "${raw}" | Normalized: "${normalized}"`);
  }

  return normalized;
};

const toSlNo = (value) => {
  const parsed = Number(normalizeCell(value));
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
};

const formatPrDateCell = (value) => {
  if (value === null || value === undefined) return value;
  const raw = normalizeCell(value);
  if (!raw) return raw;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return value;
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const day = String(date.getDate()).padStart(2, '0');
  const month = months[date.getMonth()];
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
};

// ─── PR Number sort utility ───────────────────────────────────────────────────
// Extract trailing numeric digits for sorting: "PR/2025/084" → 84
const extractPrSortValue = (prValue) => {
  const str = normalizeCell(prValue);
  const match = str.match(/(\d+)\s*$/);
  return match ? Number(match[1]) : Number.POSITIVE_INFINITY;
};

// ─── Deduplication ────────────────────────────────────────────────────────────
// Dedup key priority:
// 1) Normalized PR Number (col 3) + Description (col 8)
//    (allows multiple line items under same PR with different descriptions)
// 2) Fallback: Vendor (col 7) + Total (col 12) + Description (col 8)
const buildDuplicateKey = (row) => {
  const normalizedPr = normalizePrNumber(row?.[3]);
  const description = normalizeKeyText(row?.[8]);
  if (normalizedPr) {
    if (!description) return '';
    return `pr:${normalizedPr}||desc:${description}`;
  }

  const vendor = normalizeKeyText(row?.[7]);
  const total = normalizeAmountKey(row?.[12]);
  const fallbackKey = [vendor, total, description].join('||');
  if (!fallbackKey || fallbackKey === '||') return '';
  return `fallback:${fallbackKey}`;
};

const dedupeRows = (rows) => {
  const seen = new Set();
  return rows.filter((row) => {
    const key = buildDuplicateKey(row);
    if (!key || key === '||') return false;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

// ─── Grouping & Sorting ───────────────────────────────────────────────────────

const groupRowsByPo = (rows) => {
  const groups = [];
  let currentGroup = [];
  let lastSlNo = null;

  rows.forEach((row) => {
    const slNo = toSlNo(row[0]);
    if (slNo && slNo !== lastSlNo) {
      if (currentGroup.length > 0) groups.push(currentGroup);
      currentGroup = [row];
      lastSlNo = slNo;
    } else {
      if (currentGroup.length === 0) currentGroup = [row];
      else currentGroup.push(row);
      if (slNo) lastSlNo = slNo;
    }
  });

  if (currentGroup.length > 0) groups.push(currentGroup);
  return groups;
};

// Sort PO groups by PR Number ascending (extracted as-is from PDF, never modified)
// Then reassign Sl No sequentially based on sorted order
const sortPoGroupsByPrNumber = (rows) =>
  groupRowsByPo(rows)
    .sort((groupA, groupB) => {
      const prA = extractPrSortValue(groupA[0]?.[3]);
      const prB = extractPrSortValue(groupB[0]?.[3]);
      return prA - prB;
    })
    .map((group, index) =>
      group.map((row, rowIndex) => {
        const r = Array.isArray(row) ? [...row] : [];
        r[0] = rowIndex === 0 ? index + 1 : '';
        return r;
      })
    )
    .flat();

const normalizeSlNoDisplay = (rows) => {
  let lastSlNo = null;
  return rows.map((row) => {
    const r = Array.isArray(row) ? [...row] : [];
    const slNo = toSlNo(r[0]);
    if (slNo) {
      if (slNo === lastSlNo) r[0] = '';
      else lastSlNo = slNo;
    }
    return r;
  });
};

const getPdfGroupKey = (row) => {
  const slNo = toSlNo(row?.[0]);
  if (slNo) return `sl:${slNo}`;
  const pr = normalizePrNumber(row?.[3]);
  if (pr) return `pr:${pr}`;
  const po = normalizeKeyText(row?.[5]);
  if (po) return `po:${po}`;
  const vendor = normalizeKeyText(row?.[7]);
  const total = normalizeAmountKey(row?.[12]);
  if (vendor || total) return `fallback:${vendor}||${total}`;
  return 'fallback:unknown';
};

// ─── FIX: Detect continuation rows (extra line items of the same PDF) ─────────
// A continuation row has no SL No, no PR number, no PO number, no vendor —
// only description/qty/rate columns are filled. It must stay with the previous group.
const isContinuationRow = (row) => {
  const slNo = toSlNo(row?.[0]);
  const pr = normalizePrNumber(row?.[3]);
  const po = normalizeKeyText(row?.[5]);
  const vendor = normalizeKeyText(row?.[7]);
  return !slNo && !pr && !po && !vendor;
};

// ─── FIX: Group rows by PDF — continuation rows always attach to current group ─
const groupRowsByPdfKey = (rows) => {
  const groups = [];
  let currentGroup = [];
  let lastKey = null;

  rows.forEach((row) => {
    // Continuation row (only desc/qty/rate filled) → glue to current group, never start a new one
    if (isContinuationRow(row)) {
      if (currentGroup.length > 0) {
        currentGroup.push(row);
      } else {
        currentGroup = [row];
      }
      return;
    }

    const key = getPdfGroupKey(row);
    if (lastKey !== null && key !== lastKey) {
      if (currentGroup.length > 0) groups.push(currentGroup);
      currentGroup = [row];
    } else {
      currentGroup.push(row);
    }
    lastKey = key;
  });

  if (currentGroup.length > 0) groups.push(currentGroup);
  return groups;
};

// ─── Workbook builder ─────────────────────────────────────────────────────────

const buildWorkbook = (headers, rows) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(SHEET_NAME);
  worksheet.addRow(headers);

  const normalizedRows = rows.map((row) => {
    const r = Array.isArray(row) ? [...row] : [];
    if (r.length > 4) r[4] = formatPrDateCell(r[4]);
    return r.map((cell) =>
      typeof cell === 'string'
        ? cell.replace(/\r?\n|\r/g, ' ').trim()
        : cell
    );
  });
  const nonEmptyRows = normalizedRows.filter((row) => !isRowEmpty(row));
  const summaryRows = nonEmptyRows.filter(isCombinedGrandTotalRow);
  const dataRows = nonEmptyRows.filter((row) => !isCombinedGrandTotalRow(row));
  const finalRows = [
    ...normalizeSlNoDisplay(sortPoGroupsByPrNumber(dedupeRows(dataRows))),
    ...summaryRows
  ];

  // ─── FIX: Write full row for every line item; 2 empty rows only between PDFs ─
  const groupedRows = groupRowsByPdfKey(finalRows);
  groupedRows.forEach((group, groupIndex) => {
    group.forEach((row) => {
      worksheet.addRow(row); // full row written for every line item
    });
    if (groupIndex < groupedRows.length - 1) {
      worksheet.addRow([]);
      worksheet.addRow([]); // 2 empty rows only after finishing one PDF's rows
    }
  });

  return workbook;
};

// ─── PDF builder ─────────────────────────────────────────────────────────────
const buildPdfBuffer = (headers, rows) =>
  new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 30 });
    const chunks = [];
    doc.on('data', (c) => chunks.push(c));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const colCount = Math.max(headers.length, 1);
    const colWidth = pageWidth / colCount;

    const measureRowHeight = (cells, isHeader = false) => {
      doc.font(isHeader ? 'Helvetica-Bold' : 'Helvetica');
      doc.fontSize(isHeader ? 8 : 7);
      let maxH = isHeader ? 12 : 10;
      const stringCells = cells.map((c) => String(c ?? ''));
      stringCells.forEach((text) => {
        const h = doc.heightOfString(text, { width: colWidth - 4, align: 'left' });
        if (h > maxH) maxH = h;
      });
      return maxH;
    };

    const drawRow = (cells, y, isHeader = false, fixedHeight) => {
      doc.font(isHeader ? 'Helvetica-Bold' : 'Helvetica');
      doc.fontSize(isHeader ? 8 : 7);
      let x = doc.page.margins.left;

      const stringCells = cells.map((c) => String(c ?? ''));
      const maxH = typeof fixedHeight === 'number' ? fixedHeight : measureRowHeight(stringCells, isHeader);

      stringCells.forEach((text) => {
        doc.rect(x, y, colWidth, maxH).strokeColor('#d0d0d0').stroke();
        doc.text(text, x + 2, y + 2, { width: colWidth - 4, align: 'left' });
        x += colWidth;
      });

      return maxH;
    };

    const truncateToHeight = (text, maxHeight) => {
      const raw = String(text ?? '');
      if (!raw) return raw;
      if (doc.heightOfString(raw, { width: colWidth - 4, align: 'left' }) <= maxHeight) return raw;
      const suffix = '...';
      let low = 0;
      let high = raw.length;
      let best = '';
      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        const candidate = raw.slice(0, mid).trimEnd() + suffix;
        const h = doc.heightOfString(candidate, { width: colWidth - 4, align: 'left' });
        if (h <= maxHeight) {
          best = candidate;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }
      return best || suffix;
    };

    let y = doc.page.margins.top;
    const headerHeight = drawRow(headers, y, true);
    y += headerHeight;

    if (!rows || rows.length === 0) {
      drawRow([`No data available`], y, false);
      doc.end();
      return;
    }

    rows.forEach((row) => {
      const rowCells = headers.map((_, i) => (Array.isArray(row) ? row[i] : ''));
      const bottomLimit = doc.page.height - doc.page.margins.bottom - 10;
      const maxRowHeight = (doc.page.height - doc.page.margins.top - doc.page.margins.bottom) - headerHeight - 10;
      const measuredHeight = measureRowHeight(rowCells, false);
      const needsTruncate = measuredHeight > maxRowHeight;
      const targetHeight = needsTruncate ? maxRowHeight : measuredHeight;

      if (y + targetHeight > bottomLimit) {
        doc.addPage();
        y = doc.page.margins.top;
        const h = drawRow(headers, y, true);
        y += h;
      }

      const safeCells = needsTruncate
        ? rowCells.map((cell) => truncateToHeight(cell, targetHeight))
        : rowCells;

      const drawnHeight = drawRow(safeCells, y, false, targetHeight);
      y += drawnHeight;
    });

    doc.end();
  });

// ─── Read existing data rows from master (strict 24-col header match) ─────────

const getExistingDataRows = (workbook, headers) => {
  const sheetName = workbook.SheetNames[0] || SHEET_NAME;
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) return [];

  const existingRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  if (existingRows.length === 0) return [];

  const headerRowIndex = existingRows.findIndex(
    (row) => Array.isArray(row) && headersMatch(row, headers)
  );

  if (headerRowIndex === -1) {
    throw new Error(
      'Header mismatch in po_master.xlsx. Expected 24-column PO Tracker headers. Please delete or fix po_master.xlsx and try again.'
    );
  }

  return existingRows
    .slice(headerRowIndex + 1)
    .filter((row) => !isRowEmpty(row))
    .map((row) => headers.map((_, index) => row[index] ?? ''));
};

const buildExistingPrSet = (rows) => {
  const prs = new Set();
  rows.forEach((row) => {
    const pr = normalizePrNumber(row?.[3]);
    if (pr) prs.add(pr);
  });
  return prs;
};

const filterOutExistingRows = (rows, existingPrs) =>
  rows.filter((row) => {
    const pr = normalizePrNumber(row?.[3]);
    if (!pr) return true;
    return !existingPrs.has(pr);
  });

const getOutputPaths = (mode) => {
  if (mode === 'existing') {
    return {
      outputFilename: path.basename(MASTER_FILE_PATH),
      outputPath: MASTER_FILE_PATH
    };
  }
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const outputFilename = `po_export_${timestamp}.xlsx`;
  return {
    outputFilename,
    outputPath: path.join(STORAGE_DIR, outputFilename)
  };
};

// ─── HTTP helpers ─────────────────────────────────────────────────────────────

const sendJson = (res, statusCode, payload) => {
  res.writeHead(statusCode, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(payload));
};

const readJsonBody = (req, limitBytes = 5 * 1024 * 1024) =>
  new Promise((resolve, reject) => {
    let raw = '';
    req.on('data', (chunk) => {
      raw += chunk.toString();
      if (raw.length > limitBytes) reject(new Error('Payload too large'));
    });
    req.on('end', () => resolve(raw));
    req.on('error', reject);
  });

// ─── PDF extraction ───────────────────────────────────────────────────────────

const extractPdfText = async (data) => {
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const textParts = [];
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    const pageText = content.items
      .map((item) => (typeof item?.str === 'string' ? item.str : ''))
      .filter(Boolean)
      .join(' ');
    if (pageText.trim()) textParts.push(pageText.trim());
  }
  return textParts.join('\n\n');
};

const buildGeminiPrompt = () => `
  You are an advanced AI specialized in OCR and Document Understanding for Purchase Orders (POs).
  Your task is to extract structured data from the provided PO document with 100% precision.
  You may receive extracted PDF text and/or the original PDF.

  **STRICT EXTRACTION RULES:**
  1. **PO Number**: Look for labels like "PO No", "Order #", "Purchase Order Number".
  2. **Supplier Name (Vendor)**: Look for labels: "Vendor", "Supplier", "To:", "Pay To", "Seller".
  3. **Requestor Name**: Look for labels: "Requestor", "Requested By", "Contact Person", "Buyer", "Prepared By".
  4. **Category**: If not explicitly labeled, infer based on line items. Default to "General" only if ambiguous.
  5. **PR Number**: Look for labels: "PR No", "Requisition #", "Ref No", "RFQ #". Extract EXACTLY as it appears in the document. Do NOT modify or generate PR numbers.
  6. **Dates**: PO Date, PR Date (if available), Delivery Date.
  7. **Line Items**: Extract ALL rows; merge multiline descriptions into one string with \\n.
  8. **Financials**: Tax, Total.
  9. **Remarks & Terms**: Payment terms -> no_of_days_remarks; Negotiation -> negotiation_remarks.

  **OUTPUT REQUIREMENT**:
  - Return strictly valid JSON.
  - Ensure NO fields are null if data exists in the document.
`;

const extractJsonFromResponse = (raw) => {
  if (!raw) throw new Error('No data returned');
  let cleaned = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '');
  try {
    return JSON.parse(cleaned);
  } catch {
    const firstBrace = cleaned.indexOf('{');
    const lastBrace = cleaned.lastIndexOf('}');
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      return JSON.parse(cleaned.slice(firstBrace, lastBrace + 1));
    }
    throw new Error('Failed to parse JSON from Gemini response.');
  }
};

const loadEnvKeyFromFile = () => {
  try {
    const envPath = path.resolve(__dirname, '..', '.env.local');
    if (!fs.existsSync(envPath)) return null;
    const raw = fs.readFileSync(envPath, 'utf8');
    const line = raw.split(/\r?\n/).map((e) => e.trim())
      .find((e) => e && !e.startsWith('#') && e.startsWith('VITE_GEMINI_API_KEY='));
    if (!line) return null;
    const [, value] = line.split('=');
    return (value || '').trim();
  } catch {
    return null;
  }
};

const extractWithGemini = async ({ pdfText, inlineData, fileName }) => {
  const apiKey = process.env.GEMINI_API_KEY || process.env.VITE_GEMINI_API_KEY || loadEnvKeyFromFile();
  if (!apiKey) throw new Error('Missing GEMINI_API_KEY/VITE_GEMINI_API_KEY for server extraction.');

  const ai = new GoogleGenAI({ apiKey });
  const prompt = buildGeminiPrompt();
  const parts = [];
  if (inlineData) parts.push({ inlineData });
  if (pdfText && pdfText.trim().length > 0) {
    parts.push({ text: `${prompt}\n\nPDF_TEXT_START\n${pdfText}\nPDF_TEXT_END` });
  } else {
    parts.push({ text: prompt });
  }

  console.log('Gemini request payload summary:', {
    fileName, hasInlineData: Boolean(inlineData),
    pdfTextLength: pdfText ? pdfText.length : 0, partsCount: parts.length
  });

  const response = await ai.models.generateContent({
    model: 'gemini-2.5-flash',
    contents: { parts },
    config: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          po_number: { type: Type.STRING },
          supplier_name: { type: Type.STRING },
          date: { type: Type.STRING },
          tax: { type: Type.NUMBER },
          total: { type: Type.NUMBER },
          items: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                description: { type: Type.STRING },
                quantity: { type: Type.NUMBER },
                rate: { type: Type.NUMBER },
                amount: { type: Type.NUMBER }
              }
            }
          },
          category: { type: Type.STRING },
          requestor_name: { type: Type.STRING },
          pr_number: { type: Type.STRING },
          pr_date: { type: Type.STRING },
          description: { type: Type.STRING },
          delivery_date_agreed: { type: Type.STRING },
          actual_delivery_date: { type: Type.STRING },
          agreed_vs_actual: { type: Type.STRING },
          remarks: { type: Type.STRING },
          no_of_days_remarks: { type: Type.STRING },
          no_of_days_delay: { type: Type.STRING },
          negotiation_remarks: { type: Type.STRING }
        }
      }
    }
  });

  const rawText = response.text || '';
  console.log('Gemini raw response text (first 1200 chars):', rawText.slice(0, 1200));
  return extractJsonFromResponse(rawText);
};

// ─── Express (Excel management) ───────────────────────────────────────────────

const ensureStorageDir = async () => {
  if (!fs.existsSync(STORAGE_DIR)) {
    await fs.promises.mkdir(STORAGE_DIR, { recursive: true });
  }
};

const writeMasterExcel = async (workbook) => {
  await ensureStorageDir();
  if (!workbook) {
    throw new Error('Workbook is missing.');
  }
  await workbook.xlsx.writeFile(MASTER_FILE_PATH);
  console.log('Master Excel saved to:', MASTER_FILE_PATH);

  const stats = await fs.promises.stat(MASTER_FILE_PATH);
  if (!stats.size) {
    throw new Error('Master Excel file is empty after write.');
  }
};

const upload = multer({
  storage: multer.memoryStorage(),
  fileFilter: (req, file, cb) => {
    const isXlsx = file.originalname?.toLowerCase().endsWith('.xlsx');
    if (!isXlsx) return cb(new Error('Only .xlsx files are allowed.'));
    return cb(null, true);
  }
});

const app = express();

app.get('/api/excel/download-existing', async (req, res) => {
  try {
    await ensureStorageDir();
    const filePath = path.join(STORAGE_DIR, 'po_master.xlsx');
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ success: false, error: 'Master Excel not found' });
    }
    return res.download(filePath, 'po_master.xlsx', (err) => {
      if (err) {
        console.error('Download existing Excel error:', err);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Failed to download master Excel.' });
        }
      }
    });
  } catch (error) {
    console.error('Download existing Excel error:', error);
    return res.status(500).json({ success: false, error: 'Unable to process download.' });
  }
});

const handleDownloadExcel = async (req, res) => {
  try {
    console.log('Excel download requested for:', MASTER_FILE_PATH);
    await ensureStorageDir();
    const exists = fs.existsSync(MASTER_FILE_PATH);
    if (!exists) {
      return res.status(404).json({ success: false, error: 'Master Excel file not found.' });
    }
    const stats = await fs.promises.stat(MASTER_FILE_PATH);
    if (!stats.size) {
      return res.status(500).json({ success: false, error: 'Master Excel file is empty.' });
    }
    return res.download(MASTER_FILE_PATH, 'po_master.xlsx', (err) => {
      if (err) {
        console.error('Download error:', err);
        if (!res.headersSent) res.status(500).json({ success: false, error: 'Failed to download master file.' });
      } else {
        console.log('Excel download served:', MASTER_FILE_PATH);
      }
    });
  } catch (error) {
    console.error('Download Excel error:', error);
    return res.status(500).json({ success: false, error: 'Unable to process download.' });
  }
};

app.get('/api/download-excel', handleDownloadExcel);
// Backward-compatible alias (in case UI is hitting without /api)
app.get('/download-excel', handleDownloadExcel);

const handleUploadExcel = async (req, res) => {
  try {
    await ensureStorageDir();
    const file = req.file;
    if (!file) {
      return res.status(400).json({ success: false, error: 'No file uploaded.' });
    }
    await fs.promises.writeFile(MASTER_FILE_PATH, file.buffer);
    return res.status(200).json({ success: true, message: 'Master Excel file updated.' });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Upload failed.';
    console.error('Upload Excel error:', error);
    return res.status(400).json({ success: false, error: message });
  }
};

app.post('/api/upload-excel', upload.single('file'), handleUploadExcel);
// Backward-compatible alias (in case UI is hitting without /api)
app.post('/upload-excel', upload.single('file'), handleUploadExcel);

app.use((err, req, res, next) => {
  if (!err) return next();
  const message = err instanceof Error ? err.message : 'Request failed.';
  console.error('Excel management error:', err);
  return res.status(400).json({ success: false, error: message });
});

// ─── HTTP Server ──────────────────────────────────────────────────────────────

const server = http.createServer(async (req, res) => {
  const startedAt = Date.now();
  res.on('finish', () => {
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.url} -> ${res.statusCode} (${Date.now() - startedAt}ms)`);
  });

  await new Promise((resolve) => app(req, res, resolve));
  if (res.writableEnded || res.headersSent) return;

  const requestUrl = new URL(req.url || '/', `http://${req.headers.host}`);

  if (req.method === 'GET' && req.url === '/api/health') {
    return sendJson(res, 200, { ok: true, masterFile: MASTER_FILE_PATH });
  }

  if (req.method === 'GET' && requestUrl.pathname.startsWith('/api/download/')) {
    try {
      const filename = path.basename(decodeURIComponent(requestUrl.pathname.replace('/api/download/', '')));
      if (!filename) return sendJson(res, 400, { success: false, error: 'Invalid filename.' });
      const filePath = path.join(STORAGE_DIR, filename);
      if (!fs.existsSync(filePath)) return sendJson(res, 404, { success: false, error: 'File not found.' });
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`
      });
      fs.createReadStream(filePath)
        .on('error', (err) => {
          console.error('Download stream error:', err);
          if (!res.headersSent) sendJson(res, 500, { success: false, error: 'Stream failed.' });
          else res.destroy(err);
        })
        .pipe(res);
      return;
    } catch (error) {
      console.error('Download route error:', error);
      return sendJson(res, 500, { success: false, error: 'Unable to process download.' });
    }
  }

  if (req.method === 'POST' && requestUrl.pathname === '/api/po/extract') {
    try {
      const body = await readJsonBody(req, 15 * 1024 * 1024);
      const payload = JSON.parse(body || '{}');
      const { fileName, mimeType, base64 } = payload || {};

      console.log('PO extract request received:', {
        fileName, mimeType,
        base64Bytes: typeof base64 === 'string' ? Buffer.byteLength(base64, 'base64') : 0
      });

      if (!base64 || typeof base64 !== 'string')
        return sendJson(res, 400, { success: false, error: 'Missing base64 file payload.' });
      if (!mimeType || mimeType !== 'application/pdf')
        return sendJson(res, 400, { success: false, error: 'Only PDF uploads are supported.' });

      const buffer = Buffer.from(base64, 'base64');
      const data = new Uint8Array(buffer);
      let pdfText = '';
      try {
        pdfText = await extractPdfText(data);
        console.log(`PDF text extracted (${pdfText.length} chars) for`, fileName);
        console.log('PDF text (first 2000 chars):', pdfText.slice(0, 2000));
      } catch (pdfError) {
        console.warn('PDF text extraction failed:', pdfError);
      }

      const result = await extractWithGemini({ pdfText, inlineData: { data: base64, mimeType }, fileName });
      return sendJson(res, 200, { success: true, data: result });
    } catch (error) {
      console.error('PO extraction error:', error);
      return sendJson(res, 500, { success: false, error: error instanceof Error ? error.message : 'Unknown error' });
    }
  }

  if (req.method === 'POST' && requestUrl.pathname === '/api/pdf/export') {
    try {
      const body = await readJsonBody(req, 5 * 1024 * 1024);
      const payload = JSON.parse(body || '{}');
      const headers = Array.isArray(payload.headers) ? payload.headers : [];
      const rows = Array.isArray(payload.rows) ? payload.rows : [];

      if (headers.length === 0) return sendJson(res, 400, { error: 'Missing headers in request body.' });
      if (!fs.existsSync(STORAGE_DIR)) fs.mkdirSync(STORAGE_DIR, { recursive: true });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const filename = `po_export_${timestamp}.pdf`;
      const outputPath = path.join(STORAGE_DIR, filename);

      const pdfBuffer = await buildPdfBuffer(headers, rows);
      if (!pdfBuffer || pdfBuffer.length === 0) {
        return sendJson(res, 500, { success: false, error: 'PDF generation failed (empty output).' });
      }

      fs.writeFileSync(outputPath, pdfBuffer);
      console.log('PDF export saved:', { outputPath, filename });

      res.writeHead(200, {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`
      });
      res.end(pdfBuffer);
      return;
    } catch (error) {
      console.error('PDF export error:', error);
      return sendJson(res, 500, { success: false, error: error instanceof Error ? error.message : 'PDF export failed.' });
    }
  }

  if (req.method !== 'POST' || requestUrl.pathname !== '/api/excel/export') {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found');
    return;
  }

  // ─── Excel Export ──────────────────────────────────────────────────────────
  try {
    const body = await readJsonBody(req, 5 * 1024 * 1024);
    const payload = JSON.parse(body || '{}');
    const requestedMode = payload.excelMode ?? payload.mode;
    const mode = requestedMode === 'existing' ? 'existing' : 'new';
    const headers = Array.isArray(payload.headers) ? payload.headers : [];
    const rows = Array.isArray(payload.rows) ? payload.rows : [];

    console.log('Excel export request received:', { mode, headerCount: headers.length, rowCount: rows.length });

    if (headers.length === 0) return sendJson(res, 400, { error: 'Missing headers in request body.' });
    if (!fs.existsSync(STORAGE_DIR)) fs.mkdirSync(STORAGE_DIR, { recursive: true });

    let workbook;
    const { outputFilename, outputPath } = getOutputPaths(mode);

    if (mode === 'existing') {
      if (fs.existsSync(MASTER_FILE_PATH)) {
        const source = fs.readFileSync(MASTER_FILE_PATH);
        const existingWorkbook = XLSX.read(source, { type: 'buffer' });

        // Read existing rows — strict 24-col header match
        const existingRows = getExistingDataRows(existingWorkbook, headers);
        console.log(`Existing master has ${existingRows.length} data rows.`);

        // Build set of existing PR identifiers, then skip any incoming rows whose PR already exists
        const existingPrs = buildExistingPrSet(existingRows);
        const filteredRows = filterOutExistingRows(rows, existingPrs);
        console.log(`Filtered new rows: ${rows.length} -> ${filteredRows.length} (existing PRs skipped).`);

        // Merge existing + filtered rows, dedupe, then sort by PR Number ascending
        // PR numbers are NOT modified — extracted as-is from PDFs
        workbook = buildWorkbook(headers, [...existingRows, ...filteredRows]);
        console.log(`Built workbook: ${existingRows.length} existing + ${filteredRows.length} new rows.`);

      } else {
        // No master file yet — create fresh
        console.log('No existing master file. Creating fresh.');
        workbook = buildWorkbook(headers, rows);
      }

    } else {
      // New file mode
      workbook = buildWorkbook(headers, rows);
    }

    const outputBuffer = Buffer.from(await workbook.xlsx.writeBuffer());

    await ensureStorageDir();
    await fs.promises.writeFile(outputPath, outputBuffer);
    console.log('Excel export saved:', { outputPath, outputFilename });

    // Always update master file as the source of truth
    await writeMasterExcel(workbook);

    sendJson(res, 200, {
      success: true,
      downloadUrl: `/api/download/${encodeURIComponent(outputFilename)}`
    });

  } catch (error) {
    console.error('Excel export error:', error);
    sendJson(res, 500, { success: false, error: error instanceof Error ? error.message : 'Unknown server error' });
  }
});

server.listen(PORT, () => {
  console.log(`Excel server listening on http://localhost:${PORT}`);
  console.log(`Master Excel path: ${MASTER_FILE_PATH}`);
});
