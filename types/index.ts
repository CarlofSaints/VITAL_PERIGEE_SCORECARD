export interface Rep {
  email: string;
  firstName: string;
  lastName: string;
  fullName: string;
}

export interface StoreVisit {
  repEmail: string;
  repFirstName: string;
  repLastName: string;
  customer: string;
  channel: string;
  store: string;
  storeCode: string;
  province: string;
  date: string; // DD/MM/YYYY as in raw data
  dateObj: Date;
  answers: Record<string, string | null>; // questionId → 'Yes'|'No'|null
  comments: Record<string, string>; // questionId → comment text
  overallComment: string;
}

export interface ParseResult {
  reps: Rep[];
  channels: string[];
  dateRange: { from: string; to: string }; // DD MMM YYYY
  rowCount: number;
  visits: StoreVisit[];
}

export interface RepReport {
  rep: Rep;
  fileName: string;
  folderName: string;
  buffer: Buffer;
  fileSizeKB: number;
}

export interface GenerateResult {
  repName: string;
  fileName: string;
  status: 'ok' | 'error';
  spUrl?: string;
  emailSent?: boolean;
  error?: string;
}
