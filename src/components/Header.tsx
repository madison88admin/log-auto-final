import React from 'react';
import { FileSpreadsheet, Package } from 'lucide-react';

export const Header: React.FC = () => {
  return (
    <header className="bg-white border-b border-gray-200 shadow-sm">
      <div className="max-w-6xl mx-auto px-4 py-6">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 text-primary-600">
            <Package className="w-8 h-8" />
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-gray-900">
              Packing List Report Generator
            </h1>
            <p className="text-gray-600 mt-1">
              Process Excel packing lists and generate formatted reports with validation
            </p>
          </div>
        </div>
      </div>
    </header>
  );
};