import React, { useCallback, useState } from 'react';
import { UploadCloud, FileText, X, AlertCircle, FilePlus2, FileSpreadsheet, FileImage, File } from 'lucide-react';

interface FileUploadProps {
  files: File[];
  onFilesChange: (files: File[]) => void;
  disabled?: boolean;
}

const SUPPORTED_EXTENSIONS = ".pdf,.jpeg,.jpg,.png,.docx,.xlsx";

export const FileUpload: React.FC<FileUploadProps> = ({ files, onFilesChange, disabled }) => {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) setIsDragging(true);
  }, [disabled]);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) setIsDragging(false);
  }, [disabled]);

  const validateAndAddFiles = (newFiles: File[]) => {
    const validFiles: File[] = [];
    let hasInvalid = false;

    newFiles.forEach(file => {
      const ext = file.name.split('.').pop()?.toLowerCase();
      if (['pdf', 'jpeg', 'jpg', 'png', 'docx', 'xlsx'].includes(ext || '')) {
        validFiles.push(file);
      } else {
        hasInvalid = true;
      }
    });

    if (hasInvalid) {
      setError("Some files were skipped. Supported formats: PDF, JPEG, PNG, DOCX, Excel.");
      setTimeout(() => setError(null), 4000);
    }
    
    if (validFiles.length > 0) {
      onFilesChange([...files, ...validFiles]);
    }
  };

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (disabled) return;

    const droppedFiles = Array.from(e.dataTransfer.files) as File[];
    validateAndAddFiles(droppedFiles);
  }, [files, disabled, onFilesChange]);

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const selectedFiles = Array.from(e.target.files) as File[];
      validateAndAddFiles(selectedFiles);
    }
  };

  const removeFile = (index: number) => {
    const newFiles = [...files];
    newFiles.splice(index, 1);
    onFilesChange(newFiles);
  };

  const getFileIcon = (fileName: string) => {
    const ext = fileName.split('.').pop()?.toLowerCase();
    switch(ext) {
      case 'pdf': return <FileText className="w-5 h-5 text-red-500 dark:text-red-400" />;
      case 'xlsx': 
      case 'xls': return <FileSpreadsheet className="w-5 h-5 text-green-500 dark:text-green-400" />;
      case 'jpg':
      case 'jpeg':
      case 'png': return <FileImage className="w-5 h-5 text-blue-500 dark:text-blue-400" />;
      case 'docx': 
      case 'doc': return <FileText className="w-5 h-5 text-blue-600 dark:text-blue-300" />;
      default: return <File className="w-5 h-5 text-gray-400" />;
    }
  };

  return (
    <div className="w-full space-y-5">
      <div 
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={`
          relative border-[3px] border-dashed rounded-3xl p-10 transition-all duration-300 ease-in-out text-center group backdrop-blur-sm
          ${isDragging 
            ? 'border-blue-500 bg-blue-50/80 dark:bg-blue-900/30 scale-[1.01]' 
            : 'border-gray-200 dark:border-slate-700 bg-white/50 dark:bg-slate-800/50 hover:border-blue-400 dark:hover:border-blue-600 hover:bg-white/80 dark:hover:bg-slate-800/80'}
          ${disabled ? 'opacity-60 cursor-not-allowed grayscale' : 'cursor-pointer shadow-sm hover:shadow-md'}
        `}
      >
        <input 
          type="file" 
          multiple 
          accept={SUPPORTED_EXTENSIONS}
          onChange={handleFileInput}
          disabled={disabled}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer disabled:cursor-not-allowed z-10"
        />
        
        <div className="flex flex-col items-center justify-center space-y-4 pointer-events-none">
          <div className={`
            p-5 rounded-2xl transition-colors duration-300 shadow-sm
            ${isDragging ? 'bg-blue-200 dark:bg-blue-800' : 'bg-white dark:bg-slate-700 group-hover:bg-blue-50 dark:group-hover:bg-slate-600'}
          `}>
            <UploadCloud className={`w-10 h-10 ${isDragging ? 'text-blue-700 dark:text-blue-300' : 'text-blue-500 dark:text-blue-400'}`} />
          </div>
          <div className="space-y-1">
            <p className="text-xl font-bold text-gray-900 dark:text-white">Drag & drop files here</p>
            <p className="text-sm text-gray-500 dark:text-gray-400">PDF, Excel, Word, Images</p>
          </div>
          <div className="pt-2">
            <span className="inline-flex items-center px-3 py-1 rounded-full text-xs font-medium bg-gray-100/80 dark:bg-slate-700/80 text-gray-600 dark:text-gray-300 border border-gray-200 dark:border-slate-600">
               Max 10MB per file
            </span>
          </div>
        </div>
      </div>

      {error && (
        <div className="flex items-center p-4 text-sm text-red-800 dark:text-red-200 rounded-xl bg-red-50 dark:bg-red-900/20 border border-red-100 dark:border-red-800 animate-pulse shadow-sm" role="alert">
          <AlertCircle className="flex-shrink-0 w-5 h-5 mr-3" />
          <span>{error}</span>
        </div>
      )}

      {/* File List */}
      {files.length > 0 && (
        <div className="bg-white/70 dark:bg-slate-800/70 backdrop-blur-md rounded-2xl border border-gray-200 dark:border-slate-700 overflow-hidden shadow-lg">
          <div className="px-5 py-3 bg-gray-50/80 dark:bg-slate-900/50 border-b border-gray-200 dark:border-slate-700 flex justify-between items-center">
            <span className="text-sm font-semibold text-gray-700 dark:text-gray-200 flex items-center">
              <FilePlus2 className="w-4 h-4 mr-2 text-blue-500"/>
              Selected Files ({files.length})
            </span>
            <button 
              onClick={() => onFilesChange([])}
              className="text-xs text-red-600 dark:text-red-400 hover:text-red-800 dark:hover:text-red-300 font-medium px-3 py-1.5 rounded-lg hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors"
              disabled={disabled}
            >
              Clear All
            </button>
          </div>
          <ul className="divide-y divide-gray-100 dark:divide-slate-700 max-h-64 overflow-y-auto scrollbar-hide">
            {files.map((file, index) => (
              <li key={`${file.name}-${index}`} className="px-5 py-3 flex items-center justify-between hover:bg-gray-50/50 dark:hover:bg-slate-700/50 transition-colors">
                <div className="flex items-center space-x-4 overflow-hidden">
                  <div className="bg-white dark:bg-slate-800 p-2.5 rounded-xl shadow-sm border border-gray-100 dark:border-slate-700">
                    {getFileIcon(file.name)}
                  </div>
                  <div className="flex flex-col min-w-0">
                    <span className="text-sm font-semibold text-gray-700 dark:text-gray-200 truncate pr-4">{file.name}</span>
                    <span className="text-xs text-gray-400 dark:text-gray-500 font-mono">{(file.size / 1024).toFixed(1)} KB</span>
                  </div>
                </div>
                <button 
                  onClick={() => removeFile(index)}
                  disabled={disabled}
                  className="text-gray-400 hover:text-red-500 dark:text-gray-500 dark:hover:text-red-400 transition-colors p-2 rounded-full hover:bg-red-50 dark:hover:bg-red-900/20"
                >
                  <X className="w-4 h-4" />
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
};