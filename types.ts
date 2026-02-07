
export interface ExtractedData {
  beneficiaryName: string;
  accountNumber: string;
  swiftCode: string;
  country: string;
  state: string;
  city: string;
  address: string;
  bankName: string;
  goodsDescription?: string;
}

export interface EnrichedData extends ExtractedData {
  companyInfo?: string;
  sources?: { uri: string; title: string }[];
}

export type ProcessingStatus = 'pending' | 'processing' | 'done' | 'error';

export interface ProcessableFile {
  id: string; // Unique ID for key prop
  file: File;
  status: ProcessingStatus;
  data?: ExtractedData;
  error?: string;
}