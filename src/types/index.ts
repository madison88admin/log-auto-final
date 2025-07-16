export interface ValidationIssue {
  row: number;
  column: string;
  issue: string;
  value: string;
}

export interface ValidationMismatch {
  orderNumber: string;
  row: number;
  column: string;
  originalValue: any;
  generatedValue: any;
  fieldName: string;
}

export interface OrderReport {
  orderNumber: string;
  data: PackingListRow[];
  colorSummary: ColorSummary[];
  totalQuantity: number;
  modelName?: string;
  validationResults?: ValidationMismatch[];
}

export interface PackingListRow {
  orderNumber: string;
  color: string;
  quantity: number;
  [key: string]: any;
}

export interface ColorSummary {
  color: string;
  quantity: number;
}

export interface ProcessingResult {
  success: boolean;
  orderReports: OrderReport[];
  excelBuffer?: ArrayBuffer;
  validationLog: ValidationIssue[];
  strictValidationResults: ValidationMismatch[];
  originalFileName: string;
  processedAt: Date;
}

export interface ReportTemplate {
  headers: string[];
  colorStartRow: number;
  totalRow: number;
}

export interface FileUploadProps {
  onFileUpload: (file: File) => void;
  uploadedFile: File | null;
  disabled?: boolean;
}

export interface TKListData {
  headers: string[];
  rows: Record<string, any>[];
  orderGroups: Record<string, Record<string, any>[]>;
}