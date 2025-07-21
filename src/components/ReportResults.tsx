import React from 'react';
import { FileText} from 'lucide-react';
import { ProcessingResult } from '../types';

interface ReportResultsProps {
  result: ProcessingResult;
  onReset: () => void;
  onExportSingleReport?: () => void;
  onExportAllAsZip?: () => void;
  singleReportDisabled?: boolean;
  singleReportLabel?: string;
}

export const ReportResults: React.FC<ReportResultsProps> = ({ result }) => {
  const hasUnknownOrder = result.orderReports.some(r => r.orderNumber === 'Unknown');

  return (
    <section className="card">
      {hasUnknownOrder && (
        <div className="bg-yellow-100 text-yellow-900 p-3 rounded mb-4 border border-yellow-300">
          <strong>Warning:</strong> Some orders could not be identified and are labeled as <code>Unknown</code>. Please check your input file for missing or malformed order numbers.
        </div>
      )}
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center gap-3">
          <FileText className="w-6 h-6 text-green-600" />
          <h3 className="text-lg font-medium text-gray-900">
            Reports Generated Successfully
          </h3>
        </div>
        <div className="flex gap-3">
        </div>
      </div>
    </section>
  );
};