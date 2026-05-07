export interface POItem {
  id: string;
  description: string;
  quantity: number;
  rate: number;
  amount: number;
}

export interface POData {
  id: string; // Internal ID for frontend tracking
  fileName: string;
  po_number: string;
  supplier_name: string;
  date: string;
  items: POItem[];
  tax: number;
  total: number;
  status: 'pending' | 'processing' | 'completed' | 'error';
  errorMessage?: string;
  
  // Mapped Fields
  category: string;
  requestor_name: string;
  pr_number: string;
  pr_date: string;
  description: string; // Main PO description
  
  // New specific fields for Excel Tracker
  delivery_date_agreed?: string;
  actual_delivery_date?: string;
  agreed_vs_actual?: string;
  remarks?: string;
  no_of_days_remarks?: string;
  no_of_days_delay?: string;
  negotiation_remarks?: string;
}

export interface ProcessingStats {
  total: number;
  processed: number;
  success: number;
  failed: number;
}

export type TemplateType = 'default' | 'custom';

export type ExportFormat = 'xlsx' | 'pdf' | 'csv' | 'json';