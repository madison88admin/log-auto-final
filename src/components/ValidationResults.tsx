import React, { useState } from 'react';
import { AlertTriangle, CheckCircle, ChevronDown, ChevronUp, Info } from 'lucide-react';
import { ValidationIssue } from '../types';

interface ValidationResultsProps {
  validationLog: ValidationIssue[];
}

export const ValidationResults: React.FC<ValidationResultsProps> = ({ validationLog }) => {
  const [isExpanded, setIsExpanded] = useState(false);

  const hasIssues = validationLog.length > 0;
  const criticalIssues = validationLog.filter(issue => 
    issue.issue.includes('Missing') || issue.issue.includes('Invalid')
  );
  const warningIssues = validationLog.filter(issue => 
    !criticalIssues.includes(issue)
  );

  return (
    <section className="card">
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-3">
          {hasIssues ? (
            <AlertTriangle className="w-6 h-6 text-amber-600" />
          ) : (
            <CheckCircle className="w-6 h-6 text-green-600" />
          )}
          <h3 className="text-lg font-medium text-gray-900">
            Data Validation Results
          </h3>
        </div>
        
        {hasIssues && (
          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="btn-secondary inline-flex items-center gap-2"
          >
            {isExpanded ? (
              <>
                <ChevronUp className="w-4 h-4" />
                Hide Details
              </>
            ) : (
              <>
                <ChevronDown className="w-4 h-4" />
                Show Details
              </>
            )}
          </button>
        )}
      </div>

      {!hasIssues ? (
        <div className="bg-green-50 rounded-lg p-4">
          <div className="flex items-start gap-3">
            <CheckCircle className="w-5 h-5 text-green-600 mt-0.5" />
            <div>
              <p className="font-medium text-green-900">
                All data validation checks passed
              </p>
              <p className="text-sm text-green-700 mt-1">
                No issues found with order numbers, colors, or quantities. 
                All data is ready for report generation.
              </p>
            </div>
          </div>
        </div>
      ) : (
        <>
          {/* Summary */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            {criticalIssues.length > 0 && (
              <div className="bg-red-50 rounded-lg p-4">
                <div className="flex items-center gap-3">
                  <AlertTriangle className="w-6 h-6 text-red-600" />
                  <div>
                    <p className="font-medium text-red-900">
                      {criticalIssues.length} Critical Issues
                    </p>
                    <p className="text-sm text-red-700">
                      Missing or invalid required data
                    </p>
                  </div>
                </div>
              </div>
            )}
            
            {warningIssues.length > 0 && (
              <div className="bg-amber-50 rounded-lg p-4">
                <div className="flex items-center gap-3">
                  <Info className="w-6 h-6 text-amber-600" />
                  <div>
                    <p className="font-medium text-amber-900">
                      {warningIssues.length} Warnings
                    </p>
                    <p className="text-sm text-amber-700">
                      Data quality concerns
                    </p>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Detailed Issues */}
          {isExpanded && (
            <div className="space-y-4">
              <div className="border-t border-gray-200 pt-4">
                <h4 className="font-medium text-gray-900 mb-3">
                  Detailed Validation Issues
                </h4>
                
                <div className="overflow-x-auto">
                  <table className="min-w-full divide-y divide-gray-200">
                    <thead className="bg-gray-50">
                      <tr>
                        <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                          Row
                        </th>
                        <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                          Column
                        </th>
                        <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                          Issue
                        </th>
                        <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                          Value
                        </th>
                      </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-gray-200">
                      {validationLog.map((issue, index) => (
                        <tr key={index} className={
                          criticalIssues.includes(issue) ? 'bg-red-50' : 'bg-amber-50'
                        }>
                          <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-900">
                            {issue.row}
                          </td>
                          <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-900">
                            {issue.column}
                          </td>
                          <td className="px-4 py-3 text-sm text-gray-900">
                            {issue.issue}
                          </td>
                          <td className="px-4 py-3 text-sm text-gray-500 max-w-xs truncate">
                            {issue.value || 'N/A'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </>
      )}
    </section>
  );
};