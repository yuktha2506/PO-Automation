import React, { useState, useRef, useEffect } from 'react';
import { FileUpload } from './FileUpload';
import { TemplateSelector } from './TemplateSelector';
import { ProgressBar } from './ProgressBar';
import { DataPreview } from './DataPreview';
import { api } from '../services/api';
import { POData, TemplateType, ExportFormat } from '../types';
import { Download, Play, RefreshCw, CheckCircle2, FileText, ChevronDown, FileSpreadsheet, FileJson, FileType, Image, AlertTriangle, Upload } from 'lucide-react';

export const Dashboard: React.FC = () => {
  const [files, setFiles] = useState<File[]>([]);
  const [extractedData, setExtractedData] = useState<POData[]>([]);
  const [processing, setProcessing] = useState(false);
  const [progress, setProcessingProgress] = useState(0);
  const [templateType, setTemplateType] = useState<TemplateType>('default');
  const [customTemplate, setCustomTemplate] = useState<File | null>(null);
  const [generationLoading, setGenerationLoading] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);
  const [isExportOpen, setIsExportOpen] = useState(false);
  const [isExcelSubOpen, setIsExcelSubOpen] = useState(false);
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [excelDownloadLoading, setExcelDownloadLoading] = useState(false);
  const [excelUploadLoading, setExcelUploadLoading] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const excelUploadInputRef = useRef<HTMLInputElement>(null);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsExportOpen(false);
        setIsExcelSubOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const executeProcessing = async () => {
    setShowConfirmDialog(false);
    if (files.length === 0) return;
    
    setProcessing(true);
    setExtractedData([]); 
    setProcessingProgress(0);
    setShowSuccess(false);

    const results: POData[] = [];
    
    for (let i = 0; i < files.length; i++) {
      try {
        const data = await api.uploadAndExtract(files[i]);
        results.push(data);
      } catch (error) {
        console.error("Error processing file", files[i].name, error);
        results.push({
          id: Math.random().toString(),
          fileName: files[i].name,
          po_number: 'N/A',
          supplier_name: 'Unknown',
          date: '',
          items: [],
          tax: 0,
          total: 0,
          status: 'error',
          errorMessage: 'OCR Extraction Failed',
          category: '',
          requestor_name: '',
          pr_number: '',
          pr_date: '',
          description: ''
        });
      }
      setProcessingProgress(prev => prev + 1);
      setExtractedData(prev => [...prev, results[results.length - 1]]);
    }

    setProcessing(false);
  };

  const handleProcessRequest = () => {
    if (files.length === 0) return;
    setShowConfirmDialog(true);
  };

  const handleUpdateData = (id: string, newData: Partial<POData>) => {
    setExtractedData(prev => prev.map(item => item.id === id ? { ...item, ...newData } : item));
  };

  const handleGenerateFile = async (
    format: ExportFormat | 'docx' | 'jpeg' | 'jpg',
    excelMode?: 'new' | 'existing'
  ) => {
    if (extractedData.length === 0) return;
    setGenerationLoading(true);
    setIsExportOpen(false);
    setIsExcelSubOpen(false);

    try {
      const validData = extractedData.filter(d => d.status === 'completed');
      const dateStamp = new Date().toISOString().slice(0, 10);
      const suffix = format === 'xlsx' ? `_${excelMode || 'new'}` : '';
      const fileName = `PO_Export_${dateStamp}${suffix}.${format}`;

      let blob: Blob;

      if (format === 'xlsx') {
        // ✅ FIXED: For Excel, get the download URL from server then fetch the actual file
        const downloadUrl = await api.generateExcelDownloadUrl(
          validData,
          templateType,
          customTemplate || undefined,
          { excelMode }
        );
        const fileResponse = await fetch(downloadUrl);
        if (!fileResponse.ok) {
          throw new Error(`Failed to download Excel file: ${fileResponse.statusText}`);
        }
        blob = await fileResponse.blob();
      } else {
        // For all other formats (pdf, docx, jpeg, jpg, csv, json) use generateExport directly
        blob = await api.generateExport(
          validData,
          templateType,
          format as any,
          customTemplate || undefined
        );
      }

      const url = window.URL.createObjectURL(blob);
      if (format === 'pdf') {
        const opened = window.open(url, '_blank');
        if (!opened) {
          const a = document.createElement('a');
          a.href = url;
          a.download = fileName;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        }
        window.URL.revokeObjectURL(url);
      } else {
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      }
      
      setShowSuccess(true);
      setTimeout(() => setShowSuccess(false), 5000);
    } catch (error) {
      console.error("Export failed", error);
      alert("Failed to Generate");
    } finally {
      setGenerationLoading(false);
    }
  };

  const handleDownloadExcel = async () => {
    setExcelDownloadLoading(true);
    try {
      const blob = await api.downloadMasterExcel();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'po_master.xlsx';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (error) {
      console.error('Existing Excel download failed', error);
      alert('Failed to Download');
    } finally {
      setExcelDownloadLoading(false);
    }
  };

  const handleUploadExcelClick = () => {
    excelUploadInputRef.current?.click();
  };

  const handleUploadExcelChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      alert('Please select a valid .xlsx file.');
      return;
    }
    setExcelUploadLoading(true);
    try {
      const result = await api.uploadMasterExcel(file);
      alert(result?.message || 'Excel uploaded successfully.');
    } catch (error: any) {
      console.error('Upload Excel failed', error);
      alert('Failed to Upload');
    } finally {
      setExcelUploadLoading(false);
    }
  };

  const resetAll = () => {
    setFiles([]);
    setExtractedData([]);
    setProcessingProgress(0);
    setShowSuccess(false);
  };

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10 space-y-10 pb-32">
      
      <div className="text-center space-y-4 py-12 bg-gradient-to-b from-blue-50 via-white to-transparent dark:from-slate-800/50 dark:via-slate-900/50 dark:to-transparent rounded-[2.5rem] shadow-sm border border-white/50 dark:border-slate-700/50 backdrop-blur-sm">
        <h1 className="text-4xl md:text-6xl font-extrabold text-gray-900 dark:text-white tracking-tight leading-tight">
          Purchase Order <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-400">Automation</span>
        </h1>
        <p className="text-lg md:text-xl text-gray-600 dark:text-gray-300 max-w-2xl mx-auto leading-relaxed font-light">
          Extract data from multiple document formats and convert them into structured reports instantly.
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        
        <div className="lg:col-span-2 space-y-8">
          <section className="space-y-4">
             <div className="flex items-center justify-between px-1">
                <h2 className="text-2xl font-bold text-gray-800 dark:text-gray-100 flex items-center">
                  <span className="flex items-center justify-center w-8 h-8 rounded-full bg-blue-100 dark:bg-blue-900/50 text-blue-600 dark:text-blue-400 text-sm font-bold mr-3 shadow-sm">1</span>
                  Upload Documents
                </h2>
                {files.length > 0 && (
                  <span className="text-sm font-medium px-4 py-1.5 bg-white dark:bg-slate-800 border border-gray-200 dark:border-slate-700 text-gray-600 dark:text-gray-300 rounded-full shadow-sm">
                    {files.length} {files.length === 1 ? 'file' : 'files'} ready
                  </span>
                )}
             </div>
             <FileUpload 
                files={files} 
                onFilesChange={setFiles} 
                disabled={processing}
             />
          </section>

          {processing && (
            <section className="bg-white/80 dark:bg-slate-800/80 backdrop-blur-md p-8 rounded-2xl shadow-xl border border-blue-100 dark:border-slate-700 animate-in fade-in duration-300">
               <ProgressBar current={progress} total={files.length} />
               <p className="text-center text-sm text-gray-500 dark:text-gray-400 mt-4 animate-pulse">Analyzing documents...</p>
            </section>
          )}

          {extractedData.length > 0 && !processing && (
            <section className="animate-in slide-in-from-bottom-4 duration-500">
              <DataPreview data={extractedData} onUpdate={handleUpdateData} />
            </section>
          )}
        </div>

        <div className="lg:col-span-1 space-y-8 relative">
          
          <section className="space-y-4">
            <h2 className="text-2xl font-bold text-gray-800 dark:text-gray-100 flex items-center px-1">
              <span className="flex items-center justify-center w-8 h-8 rounded-full bg-indigo-100 dark:bg-indigo-900/50 text-indigo-600 dark:text-indigo-400 text-sm font-bold mr-3 shadow-sm">2</span>
              Export Settings
            </h2>
            <TemplateSelector 
              selectedTemplate={templateType}
              onSelect={setTemplateType}
              customTemplateFile={customTemplate}
              onCustomTemplateUpload={setCustomTemplate}
            />
          </section>

          <div className="bg-white/70 dark:bg-slate-800/70 backdrop-blur-xl p-6 rounded-3xl shadow-xl border border-white/40 dark:border-slate-700 space-y-6 relative z-30 overflow-visible">
            
            <h3 className="font-semibold text-gray-800 dark:text-gray-200 text-lg flex items-center">
              Actions
            </h3>
            
            <div className="space-y-4">
              <button
                type="button"
                onClick={handleProcessRequest}
                disabled={files.length === 0 || processing}
                className={`w-full py-4 px-6 rounded-xl flex items-center justify-center font-semibold text-lg transition-all duration-300 transform shadow-lg ${
                  files.length === 0 || processing
                    ? 'bg-gray-100 dark:bg-slate-700 text-gray-400 dark:text-gray-500 cursor-not-allowed shadow-none'
                    : 'bg-gradient-to-r from-blue-600 to-indigo-600 text-white hover:shadow-blue-500/30 hover:-translate-y-0.5'
                }`}
              >
                {processing ? (
                  <RefreshCw className="w-5 h-5 mr-3 animate-spin" />
                ) : (
                  <Play className="w-5 h-5 mr-3 fill-current" />
                )}
                {processing ? 'Processing...' : 'Upload & Process'}
              </button>

              {/* Export Dropdown */}
              <div className="relative" ref={dropdownRef}>
                <button
                  type="button"
                  onClick={() => setIsExportOpen(!isExportOpen)}
                  disabled={extractedData.length === 0 || processing || generationLoading}
                  className={`w-full py-4 px-6 rounded-xl flex items-center justify-between font-semibold text-lg transition-all duration-300 transform shadow-lg ${
                    extractedData.length === 0 || processing
                      ? 'bg-gray-50 dark:bg-slate-800/50 text-gray-300 dark:text-slate-600 cursor-not-allowed border-2 border-dashed border-gray-200 dark:border-slate-700 shadow-none'
                      : 'bg-gradient-to-r from-emerald-500 to-teal-600 text-white hover:shadow-emerald-500/30 hover:-translate-y-0.5'
                  }`}
                >
                  <div className="flex items-center">
                    {generationLoading ? (
                      <RefreshCw className="w-5 h-5 mr-3 animate-spin" />
                    ) : (
                      <Download className="w-5 h-5 mr-3" />
                    )}
                    Generate File
                  </div>
                  {!generationLoading && <ChevronDown className={`w-5 h-5 transition-transform ${isExportOpen ? 'rotate-180' : ''}`} />}
                </button>

                {/* Dropdown Menu */}
                {isExportOpen && !generationLoading && (
                  <div className="absolute top-full left-0 right-0 mt-2 bg-white dark:bg-slate-800 rounded-xl shadow-2xl border border-gray-100 dark:border-slate-700 z-[9999] overflow-hidden animate-in fade-in zoom-in-95 duration-150 ring-1 ring-black/5">
                    <div className="p-1">
                      <div className="relative">
                        <button
                          type="button"
                          onClick={() => setIsExcelSubOpen(prev => !prev)}
                          className="w-full text-left px-4 py-3 hover:bg-emerald-50 dark:hover:bg-slate-700 rounded-lg flex items-center justify-between transition-colors text-gray-700 dark:text-gray-200"
                        >
                          <span className="flex items-center space-x-3">
                            <FileSpreadsheet className="w-5 h-5 text-emerald-500" />
                            <span>Excel (.xlsx)</span>
                          </span>
                          <ChevronDown className={`w-4 h-4 transition-transform ${isExcelSubOpen ? 'rotate-180' : ''}`} />
                        </button>
                        {isExcelSubOpen && (
                          <div className="mt-1 ml-10 mr-2 space-y-1">
                            <button
                              type="button"
                              onClick={() => handleGenerateFile('xlsx', 'new')}
                              className="w-full text-left px-4 py-2.5 hover:bg-emerald-50 dark:hover:bg-slate-700 rounded-lg flex items-center space-x-3 transition-colors text-gray-700 dark:text-gray-200"
                            >
                              <FileSpreadsheet className="w-4 h-4 text-emerald-500" />
                              <span>New Excel</span>
                            </button>
                            <button
                              type="button"
                              onClick={() => handleGenerateFile('xlsx', 'existing')}
                              className="w-full text-left px-4 py-2.5 hover:bg-emerald-50 dark:hover:bg-slate-700 rounded-lg flex items-center space-x-3 transition-colors text-gray-700 dark:text-gray-200"
                            >
                              <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                              <span>Existing Excel</span>
                            </button>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>

            <div className="pt-4 border-t border-gray-100 dark:border-slate-700 space-y-3">
              <h4 className="text-sm font-semibold text-gray-700 dark:text-gray-300 tracking-wide uppercase">
                Excel Management
              </h4>
              <div className="grid grid-cols-1 gap-3">
                <button
                  type="button"
                  onClick={handleDownloadExcel}
                  disabled={excelDownloadLoading}
                  className={`w-full py-3 px-4 rounded-xl flex items-center justify-center font-semibold transition-all duration-300 shadow-md ${
                    excelDownloadLoading
                      ? 'bg-gray-100 dark:bg-slate-700 text-gray-400 dark:text-gray-500 cursor-not-allowed shadow-none'
                      : 'bg-white dark:bg-slate-800 text-gray-700 dark:text-gray-200 border border-gray-200 dark:border-slate-700 hover:shadow-lg'
                  }`}
                >
                  {excelDownloadLoading ? (
                    <RefreshCw className="w-4 h-4 mr-2 animate-spin" />
                  ) : (
                    <Download className="w-4 h-4 mr-2" />
                  )}
                  Download Excel
                </button>

                <button
                  type="button"
                  onClick={handleUploadExcelClick}
                  disabled={excelUploadLoading}
                  className={`w-full py-3 px-4 rounded-xl flex items-center justify-center font-semibold transition-all duration-300 shadow-md ${
                    excelUploadLoading
                      ? 'bg-gray-100 dark:bg-slate-700 text-gray-400 dark:text-gray-500 cursor-not-allowed shadow-none'
                      : 'bg-white dark:bg-slate-800 text-gray-700 dark:text-gray-200 border border-gray-200 dark:border-slate-700 hover:shadow-lg'
                  }`}
                >
                  {excelUploadLoading ? (
                    <RefreshCw className="w-4 h-4 mr-2 animate-spin" />
                  ) : (
                    <Upload className="w-4 h-4 mr-2" />
                  )}
                  Upload Updated Excel
                </button>
                <input
                  ref={excelUploadInputRef}
                  type="file"
                  accept=".xlsx"
                  className="hidden"
                  onChange={handleUploadExcelChange}
                />
              </div>
            </div>

            {showSuccess && (
              <div className="p-4 bg-green-50 dark:bg-green-900/30 border border-green-100 dark:border-green-800 text-green-700 dark:text-green-300 rounded-xl flex items-start animate-in fade-in slide-in-from-top-2">
                <CheckCircle2 className="w-5 h-5 mr-3 mt-0.5 flex-shrink-0" />
                <div>
                   <p className="font-bold">Success!</p>
                   <p className="text-sm opacity-90">File generated and downloaded.</p>
                </div>
              </div>
            )}
            
            {extractedData.length > 0 && (
               <div className="pt-2 text-center">
                 <button 
                    type="button"
                    onClick={resetAll}
                    className="inline-flex items-center text-sm font-medium text-gray-500 dark:text-gray-400 hover:text-red-600 dark:hover:text-red-400 transition-colors"
                  >
                    <RefreshCw className="w-3 h-3 mr-2" />
                    Reset & Start Over
                 </button>
               </div>
            )}
          </div>
        </div>

      </div>

      {/* Confirmation Dialog */}
      {showConfirmDialog && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/40 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-2xl max-w-md w-full p-6 border border-gray-100 dark:border-slate-700 scale-100 animate-in zoom-in-95 duration-200">
            <h3 className="text-xl font-bold text-gray-900 dark:text-white mb-2 flex items-center">
              <CheckCircle2 className="w-6 h-6 text-blue-600 dark:text-blue-400 mr-2" />
              Confirm Processing
            </h3>
            <p className="text-gray-600 dark:text-gray-300 mb-6 leading-relaxed">
              You are about to process <span className="font-bold text-blue-600 dark:text-blue-400 text-lg">{files.length}</span> {files.length === 1 ? 'file' : 'files'}. 
              {files.length >= 5 && (
                <span className="block mt-3 p-3 bg-amber-50 dark:bg-amber-900/20 text-amber-700 dark:text-amber-300 rounded-lg text-sm font-medium border border-amber-100 dark:border-amber-800/30">
                  <AlertTriangle className="inline w-4 h-4 mr-1 mb-0.5"/> 
                  Large batch detected. This may take some time.
                </span>
              )}
              <br/><br/>
              Do you want to continue with the upload and extraction?
            </p>
            <div className="flex justify-end space-x-3">
              <button 
                type="button"
                onClick={() => setShowConfirmDialog(false)}
                className="px-4 py-2.5 text-gray-600 dark:text-gray-300 font-semibold hover:bg-gray-100 dark:hover:bg-slate-700 rounded-xl transition-colors text-sm"
              >
                Cancel
              </button>
              <button 
                type="button"
                onClick={executeProcessing}
                className="px-5 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold rounded-xl shadow-lg hover:shadow-blue-500/25 transition-all transform hover:-translate-y-0.5 text-sm"
              >
                Yes, Start Processing
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
