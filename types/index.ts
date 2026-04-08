export interface Rep {
  email: string;
  firstName: string;
  lastName: string;
  fullName: string;
}

export interface OrphanPhoto {
  /** Base photo column header from the raw export (e.g. "Photo stock pressure"). */
  header: string;
  /** URL from the cell. */
  url: string;
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
  visitUUID?: string;
  answers: Record<string, string | null>; // questionId → 'Yes'|'No'|null
  comments: Record<string, string>; // questionId → comment text
  skus: Record<string, string>; // questionId → pipe-separated SKU list (from "Select…" columns)
  /** questionId → ordered, deduped list of photo URLs (collected from all
   *  matching "Photo Xyz" and "Photo Xyz [N]" columns in the raw export). */
  photoUrls: Record<string, string[]>;
  /** Photos from columns whose base header doesn't map to any scorecard
   *  question (e.g. "Photo stock pressure"). Shown on the "Photos with no
   *  question" sheet of the Exception Report. */
  orphanPhotos: OrphanPhoto[];
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
  spError?: string;
  emailSent?: boolean;
  repEmailSent?: boolean;
  repEmailError?: string;
  error?: string;
}
