import React from 'react';
import { Loader2, FileText, CheckCircle, AlertCircle } from 'lucide-react';

export const ProcessingStatus: React.FC = () => {
  const steps = [
    { id: 'reading', label: 'Reading Excel file', status: 'completed' },
    { id: 'extracting', label: 'Extracting data from row 20', status: 'completed' },
    { id: 'validating', label: 'Validating data integrity', status: 'active' },
    { id: 'processing', label: 'Processing order numbers', status: 'pending' },
    { id: 'generating', label: 'Generating reports', status: 'pending' },
  ];

  const getStepIcon = (status: string) => {
    switch (status) {
      case 'completed':
        return <CheckCircle className="w-5 h-5 text-green-600" />;
      case 'active':
        return <Loader2 className="w-5 h-5 text-primary-600 animate-spin" />;
      case 'error':
        return <AlertCircle className="w-5 h-5 text-red-600" />;
      default:
        return <div className="w-5 h-5 rounded-full border-2 border-gray-300" />;
    }
  };

  return (
    <section className="card">
      <div className="flex items-center gap-3 mb-6">
        <Loader2 className="w-6 h-6 text-primary-600 animate-spin" />
        <h3 className="text-lg font-medium text-gray-900">
          Processing Packing List
        </h3>
      </div>

      <div className="space-y-4">
        {steps.map((step, index) => (
          <div key={step.id} className="flex items-center gap-4">
            {getStepIcon(step.status)}
            <div className="flex-1">
              <p className={`font-medium ${
                step.status === 'active' ? 'text-primary-900' :
                step.status === 'completed' ? 'text-green-900' :
                step.status === 'error' ? 'text-red-900' :
                'text-gray-600'
              }`}>
                {step.label}
              </p>
            </div>
            {step.status === 'active' && (
              <div className="w-32 bg-gray-200 rounded-full h-2">
                <div className="bg-primary-600 h-2 rounded-full animate-pulse" style={{ width: '60%' }} />
              </div>
            )}
          </div>
        ))}
      </div>

      <div className="mt-6 p-4 bg-blue-50 rounded-lg">
        <div className="flex items-start gap-3">
          <FileText className="w-5 h-5 text-blue-600 mt-0.5" />
          <div>
            <p className="text-sm font-medium text-blue-900">
              Processing Details
            </p>
            <p className="text-sm text-blue-700 mt-1">
              Extracting unique order numbers from column E, validating colors in column D, 
              and calculating quantities from column O. This may take a few moments for large files.
            </p>
          </div>
        </div>
      </div>
    </section>
  );
};