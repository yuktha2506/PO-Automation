import React, { useRef } from 'react';
import { FileSpreadsheet, Upload, CheckCircle2 } from 'lucide-react';
import { TemplateType } from '../types';

interface TemplateSelectorProps {
  selectedTemplate: TemplateType;
  onSelect: (type: TemplateType) => void;
  customTemplateFile: File | null;
  onCustomTemplateUpload: (file: File) => void;
}

export const TemplateSelector: React.FC<TemplateSelectorProps> = ({ 
  selectedTemplate, 
  onSelect,
  customTemplateFile,
  onCustomTemplateUpload
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onCustomTemplateUpload(e.target.files[0]);
    }
  };

  return (
    <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl border border-gray-100 dark:border-slate-700 shadow-lg">
      <div className="space-y-4">
        {/* Default Template Option */}
        <label className={`relative flex items-start p-4 rounded-xl border-2 cursor-pointer transition-all duration-200 group ${selectedTemplate === 'default' ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20 dark:border-blue-500' : 'border-transparent bg-gray-50 dark:bg-slate-700/50 hover:bg-gray-100 dark:hover:bg-slate-700'}`}>
          <div className="flex items-center h-5">
            <input 
              type="radio" 
              name="template" 
              className="w-4 h-4 text-blue-600 border-gray-300 focus:ring-blue-500"
              checked={selectedTemplate === 'default'}
              onChange={() => onSelect('default')}
            />
          </div>
          <div className="ml-3">
            <span className={`block text-sm font-bold ${selectedTemplate === 'default' ? 'text-blue-700 dark:text-blue-300' : 'text-gray-900 dark:text-gray-200'}`}>Use Default Template</span>
            <span className="block text-xs text-gray-500 dark:text-gray-400 mt-1 leading-relaxed">
              <strong>Full PO Tracker Format (21 Fields):</strong><br/>
              Sl No, Category, Requestor, PR Info, Purchase Dept, PO Date, Vendor, Description, Qty, Rate, Tax, Total, Status, Delivery & Remarks.
            </span>
          </div>
          {selectedTemplate === 'default' && <CheckCircle2 className="absolute top-4 right-4 w-5 h-5 text-blue-500" />}
        </label>

        {/* Custom Template Option */}
        <label className={`relative flex items-start p-4 rounded-xl border-2 cursor-pointer transition-all duration-200 group ${selectedTemplate === 'custom' ? 'border-blue-500 bg-blue-50 dark:bg-blue-900/20 dark:border-blue-500' : 'border-transparent bg-gray-50 dark:bg-slate-700/50 hover:bg-gray-100 dark:hover:bg-slate-700'}`}>
           <div className="flex items-center h-5">
            <input 
              type="radio" 
              name="template" 
              className="w-4 h-4 text-blue-600 border-gray-300 focus:ring-blue-500"
              checked={selectedTemplate === 'custom'}
              onChange={() => onSelect('custom')}
            />
          </div>
          <div className="ml-3 w-full">
            <span className={`block text-sm font-bold ${selectedTemplate === 'custom' ? 'text-blue-700 dark:text-blue-300' : 'text-gray-900 dark:text-gray-200'}`}>Upload Custom Template</span>
            <span className="block text-xs text-gray-500 dark:text-gray-400 mt-1 mb-3">Upload your .xlsx file for precise mapping.</span>
            
            {selectedTemplate === 'custom' && (
              <div className="mt-2 animate-in fade-in zoom-in duration-200">
                <input 
                  type="file" 
                  accept=".xlsx"
                  className="hidden" 
                  ref={fileInputRef}
                  onChange={handleFileChange}
                />
                <button 
                  onClick={(e) => {
                    e.preventDefault();
                    fileInputRef.current?.click();
                  }}
                  className="inline-flex items-center px-4 py-2 border border-gray-300 dark:border-slate-600 text-xs font-semibold rounded-lg text-gray-700 dark:text-gray-200 bg-white dark:bg-slate-800 hover:bg-gray-50 dark:hover:bg-slate-700 focus:outline-none shadow-sm transition-colors"
                >
                  <Upload className="w-3 h-3 mr-2" />
                  {customTemplateFile ? (
                    <span className="truncate max-w-[150px]">{customTemplateFile.name}</span>
                  ) : (
                    'Browse Template...'
                  )}
                </button>
              </div>
            )}
          </div>
          {selectedTemplate === 'custom' && <CheckCircle2 className="absolute top-4 right-4 w-5 h-5 text-blue-500" />}
        </label>
      </div>
    </div>
  );
};