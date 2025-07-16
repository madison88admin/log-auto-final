import React, { useCallback, useState } from 'react';
import { Upload, File, X, CheckCircle } from 'lucide-react';
import { FileUploadProps } from '../types';

export const FileUpload: React.FC<FileUploadProps> = ({ 
  onFileUpload, 
  uploadedFile, 
  disabled = false 
}) => {
  const [isDragOver, setIsDragOver] = useState(false);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) {
      setIsDragOver(true);
    }
  }, [disabled]);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    
    if (disabled) return;

    const files = Array.from(e.dataTransfer.files);
    const excelFile = files.find(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xlsm')
    );

    if (excelFile) {
      onFileUpload(excelFile);
    }
  }, [onFileUpload, disabled]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onFileUpload(file);
    }
  }, [onFileUpload]);

  const handleRemoveFile = useCallback(() => {
    onFileUpload(null as any);
  }, [onFileUpload]);

  if (uploadedFile) {
    return (
      <div className="border-2 border-green-200 bg-green-50 rounded-lg p-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <CheckCircle className="w-6 h-6 text-green-600" />
            <div>
              <p className="font-medium text-green-900">File uploaded successfully</p>
              <p className="text-sm text-green-700">{uploadedFile.name}</p>
              <p className="text-xs text-green-600 mt-1">
                Size: {(uploadedFile.size / 1024 / 1024).toFixed(2)} MB
              </p>
            </div>
          </div>
          {!disabled && (
            <button
              onClick={handleRemoveFile}
              className="p-2 text-green-600 hover:text-green-800 hover:bg-green-100 rounded-lg transition-colors"
              title="Remove file"
            >
              <X className="w-5 h-5" />
            </button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div
      className={`
        border-2 border-dashed rounded-lg p-8 text-center transition-all duration-200
        ${isDragOver && !disabled
          ? 'border-primary-400 bg-primary-50' 
          : 'border-gray-300 hover:border-gray-400'
        }
        ${disabled ? 'opacity-50 cursor-not-allowed' : 'cursor-pointer'}
      `}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <div className="flex flex-col items-center gap-4">
        <div className={`
          w-16 h-16 rounded-full flex items-center justify-center
          ${isDragOver && !disabled ? 'bg-primary-100' : 'bg-gray-100'}
        `}>
          <Upload className={`
            w-8 h-8 
            ${isDragOver && !disabled ? 'text-primary-600' : 'text-gray-400'}
          `} />
        </div>
        
        <div>
          <h3 className="text-lg font-medium text-gray-900 mb-2">
            Upload Packing List Excel File
          </h3>
          <p className="text-gray-600 mb-4">
            Drag and drop your Excel file here, or click to browse
          </p>
          <p className="text-sm text-gray-500">
            Supports .xlsx and .xlsm files • Must contain "PK (2)" sheet
          </p>
        </div>

        <label className="btn-primary inline-flex items-center gap-2 cursor-pointer">
          <File className="w-4 h-4" />
          Choose File
          <input
            type="file"
            accept=".xlsx,.xlsm"
            onChange={handleFileSelect}
            className="hidden"
            disabled={disabled}
          />
        </label>
      </div>
    </div>
  );
};