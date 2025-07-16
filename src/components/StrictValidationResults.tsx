import React from 'react';
import { ValidationMismatch } from '../types';

interface StrictValidationResultsProps {
  validationResults: ValidationMismatch[];
  onReset: () => void;
}

export const StrictValidationResults: React.FC<StrictValidationResultsProps> = ({
  validationResults,
  onReset
}) => {
  if (validationResults.length === 0) {
    return (
      <section className="card border-green-200 bg-green-50">
        <div className="flex items-start gap-3">
          <div className="w-5 h-5 text-green-500 mt-0.5">
            <svg fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
            </svg>
          </div>
          <div>
            <h3 className="text-lg font-medium text-green-900">
              Strict Validation Passed
            </h3>
            <p className="text-green-700 mt-1">
              All fields in the generated reports match the original TK List data exactly.
            </p>
          </div>
        </div>
      </section>
    );
  }

  // Group mismatches by order number
  const mismatchesByOrder = validationResults.reduce((groups, mismatch) => {
    if (!groups[mismatch.orderNumber]) {
      groups[mismatch.orderNumber] = [];
    }
    groups[mismatch.orderNumber].push(mismatch);
    return groups;
  }, {} as Record<string, ValidationMismatch[]>);

  return (
    <section className="card border-red-200 bg-red-50">
      <div className="flex items-start gap-3 mb-4">
        <div className="w-5 h-5 text-red-500 mt-0.5">
          <svg fill="currentColor" viewBox="0 0 20 20">
            <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
          </svg>
        </div>
        <div>
          <h3 className="text-lg font-medium text-red-900">
            Strict Validation Failed
          </h3>
          <p className="text-red-700 mt-1">
            Found {validationResults.length} field mismatches between generated reports and original TK List data.
          </p>
        </div>
      </div>

      <div className="space-y-4">
        {Object.entries(mismatchesByOrder).map(([orderNumber, mismatches]) => (
          <div key={orderNumber} className="border border-red-200 rounded-lg p-4 bg-white">
            <h4 className="font-medium text-red-900 mb-3">
              Order {orderNumber} - {mismatches.length} mismatches
            </h4>
            
            <div className="overflow-x-auto">
              <table className="min-w-full text-sm">
                <thead>
                  <tr className="border-b border-gray-200">
                    <th className="text-left py-2 px-3 font-medium text-gray-700">Row</th>
                    <th className="text-left py-2 px-3 font-medium text-gray-700">Field</th>
                    <th className="text-left py-2 px-3 font-medium text-gray-700">Original Value</th>
                    <th className="text-left py-2 px-3 font-medium text-gray-700">Generated Value</th>
                  </tr>
                </thead>
                <tbody>
                  {mismatches.map((mismatch, index) => (
                    <tr key={index} className="border-b border-gray-100">
                      <td className="py-2 px-3 text-gray-900">{mismatch.row}</td>
                      <td className="py-2 px-3 text-gray-900 font-medium">{mismatch.fieldName}</td>
                      <td className="py-2 px-3 text-red-600 font-mono">
                        {String(mismatch.originalValue)}
                      </td>
                      <td className="py-2 px-3 text-red-600 font-mono">
                        {String(mismatch.generatedValue)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ))}
      </div>

      <div className="mt-6 flex justify-end">
        <button
          onClick={onReset}
          className="btn-secondary"
        >
          Start Over
        </button>
      </div>
    </section>
  );
}; 