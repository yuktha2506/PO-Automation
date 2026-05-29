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
const PO_TRACKER_HEADERS = [
  "Sl No", "Category", "Requestor Name", "PR Number", "PR Date",
  "Purchase Order No", "PO Date", "Vendor Name", "Description",
  "Qty", "Unit Rate", "Tax", "Grand Total", "Status",
  "Delivery Date agreed as per PO", "Actual Delivery Date",
  "Agreed vs Actual Delivery Date", "Remarks",
  "No of days from PR to PO", "Remarks",
  "No of Days from PO to Delivery", "Remarks",
  "Negotiation Savings", "Remarks"
];

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/legacy/build/pdf.worker.mjs',
  import.meta.url
).toString();

// ─── Utilities ────────────────────────────────────────────────────────────────

const normalizeHeader = (value) => {
  const normalized = String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (normalized === 'si no' || normalized === 's i no' || normalized === 's.no' || normalized === 's no') {
    return 'sl no';
  }
  return normalized;
};

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
