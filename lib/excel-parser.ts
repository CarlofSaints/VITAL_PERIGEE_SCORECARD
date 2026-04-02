import * as XLSX from 'xlsx';
import { QUESTIONS } from '@/constants/questions';
import type { ParseResult, Rep, StoreVisit } from '@/types';

function parseDDMMYYYY(dateStr: string): Date {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    const [d, m, y] = parts.map(Number);
    return new Date(y, m - 1, d);
  }
  return new Date(dateStr);
}

function formatDateDisplay(date: Date): string {
  const months = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
  ];
  return `${String(date.getDate()).padStart(2, '0')} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

// Find the next "Comments…" column after a given column index
function findCommentIdx(headers: string[], questionIdx: number): number | null {
  for (let i = questionIdx + 1; i < headers.length; i++) {
    const h = headers[i] ?? '';
    if (/^comments?[,\s]/i.test(h) || /^overall general/i.test(h)) {
      return i;
    }
    // Stop search if we hit another question column
    if (/^(Is |Has |Are |Do |Does )/i.test(h)) {
      return null;
    }
  }
  return null;
}

// Find the "Select…" column after a question (pipe-separated SKU list)
function findSkuIdx(headers: string[], questionIdx: number): number | null {
  for (let i = questionIdx + 1; i < headers.length; i++) {
    const h = headers[i] ?? '';
    if (/^select\b/i.test(h)) return i;
    if (/^comments?[,\s]/i.test(h) || /^overall general/i.test(h)) return null;
    if (/^(Is |Has |Are |Do |Does )/i.test(h)) return null;
  }
  return null;
}

// Find the first "Photo…" column after a question (image URL)
function findPhotoIdx(headers: string[], questionIdx: number): number | null {
  for (let i = questionIdx + 1; i < headers.length; i++) {
    const h = headers[i] ?? '';
    if (/^photo\b/i.test(h)) return i;
    if (/^comments?[,\s]/i.test(h) || /^overall general/i.test(h)) return null;
    if (/^(Is |Has |Are |Do |Does )/i.test(h)) return null;
  }
  return null;
}

export function parsePerigeeExport(buffer: Buffer): ParseResult {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const ws = wb.Sheets['Worksheet'];
  if (!ws) throw new Error('Sheet "Worksheet" not found in uploaded file.');

  const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' }) as string[][];
  if (rows.length < 2) throw new Error('No data rows found.');

  const headers: string[] = rows[0].map((h) => String(h ?? '').trim());

  // Build header → index map (handles duplicates by keeping all positions)
  const headerIdxMap: Record<string, number> = {};
  headers.forEach((h, i) => {
    if (h && !headerIdxMap[h]) headerIdxMap[h] = i; // first occurrence wins
  });

  // Find column indices for each question (case-insensitive partial match fallback)
  const questionColIdx: Record<string, number | null> = {};
  const questionCommentIdx: Record<string, number | null> = {};
  const questionSkuIdx: Record<string, number | null> = {};
  const questionPhotoIdx: Record<string, number | null> = {};

  for (const q of QUESTIONS) {
    // Exact match first
    let idx = headerIdxMap[q.rawHeader] ?? null;

    // Fallback: case-insensitive exact match
    if (idx === null) {
      const lc = q.rawHeader.toLowerCase();
      for (let i = 0; i < headers.length; i++) {
        if (headers[i].toLowerCase() === lc) { idx = i; break; }
      }
    }

    // Fallback: trim trailing ? and match
    if (idx === null) {
      const stripped = q.rawHeader.replace(/\?$/, '').toLowerCase();
      for (let i = 0; i < headers.length; i++) {
        if (headers[i].replace(/\?$/, '').toLowerCase() === stripped) { idx = i; break; }
      }
    }

    questionColIdx[q.id] = idx;
    questionCommentIdx[q.id] = idx !== null ? findCommentIdx(headers, idx) : null;
    questionSkuIdx[q.id] = idx !== null ? findSkuIdx(headers, idx) : null;
    questionPhotoIdx[q.id] = idx !== null ? findPhotoIdx(headers, idx) : null;
  }

  // Overall General Comments column
  const overallCommentIdx = headerIdxMap['Overall General Comments'] ?? null;

  // Essential columns
  const COL = {
    email: headerIdxMap['Email'] ?? 1,
    firstName: headerIdxMap['First Name'] ?? 2,
    lastName: headerIdxMap['Last Name'] ?? 3,
    customer: headerIdxMap['Customer'] ?? 4,
    channel: headerIdxMap['Channel'] ?? 5,
    store: headerIdxMap['Store'] ?? 6,
    storeCode: headerIdxMap['Store Code'] ?? 7,
    province: headerIdxMap['Province'] ?? 8,
    date: headerIdxMap['Date'] ?? 9,
    visitUUID: headerIdxMap['Visit UUID'] ?? null,
  };

  const repsMap: Map<string, Rep> = new Map();
  const channelSet: Set<string> = new Set();
  const visits: StoreVisit[] = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row[COL.email]) continue;

    const email = String(row[COL.email] ?? '').trim().toLowerCase();
    const firstName = String(row[COL.firstName] ?? '').trim();
    const lastName = String(row[COL.lastName] ?? '').trim();
    const channel = String(row[COL.channel] ?? '').trim();
    const store = String(row[COL.store] ?? '').trim();
    const storeCode = String(row[COL.storeCode] ?? '').trim();
    const province = String(row[COL.province] ?? '').trim();
    const dateRaw = String(row[COL.date] ?? '').trim();

    if (!email || !store) continue;

    if (!repsMap.has(email)) {
      repsMap.set(email, {
        email,
        firstName,
        lastName,
        fullName: `${firstName} ${lastName}`.trim(),
      });
    }

    if (channel) channelSet.add(channel);

    // Parse answers, comments, SKUs and photo URLs for each question
    const answers: Record<string, string | null> = {};
    const comments: Record<string, string> = {};
    const skus: Record<string, string> = {};
    const photoUrls: Record<string, string> = {};

    for (const q of QUESTIONS) {
      const ci = questionColIdx[q.id];
      if (ci !== null && ci !== undefined) {
        const raw = String(row[ci] ?? '').trim();
        answers[q.id] = raw || null;
      } else {
        answers[q.id] = null; // column not found
      }

      const cci = questionCommentIdx[q.id];
      if (cci !== null && cci !== undefined) {
        comments[q.id] = String(row[cci] ?? '').trim();
      } else {
        comments[q.id] = '';
      }

      const si = questionSkuIdx[q.id];
      skus[q.id] = si !== null && si !== undefined ? String(row[si] ?? '').trim() : '';

      const pi = questionPhotoIdx[q.id];
      photoUrls[q.id] = pi !== null && pi !== undefined ? String(row[pi] ?? '').trim() : '';
    }

    const overallComment =
      overallCommentIdx !== null ? String(row[overallCommentIdx] ?? '').trim() : '';

    visits.push({
      repEmail: email,
      repFirstName: firstName,
      repLastName: lastName,
      customer: String(row[COL.customer] ?? '').trim(),
      channel,
      store,
      storeCode,
      province,
      date: dateRaw,
      dateObj: parseDDMMYYYY(dateRaw),
      visitUUID: COL.visitUUID !== null ? String(row[COL.visitUUID] ?? '').trim() || undefined : undefined,
      answers,
      comments,
      skus,
      photoUrls,
      overallComment,
    });
  }

  if (visits.length === 0) throw new Error('No valid visit rows found in the file.');

  // Date range across all visits
  const allDates = visits.map((v) => v.dateObj).filter((d) => d.getTime() > 0);
  const minDate = new Date(Math.min(...allDates.map((d) => d.getTime())));
  const maxDate = new Date(Math.max(...allDates.map((d) => d.getTime())));

  return {
    reps: Array.from(repsMap.values()).sort((a, b) =>
      a.fullName.localeCompare(b.fullName)
    ),
    channels: Array.from(channelSet).sort(),
    dateRange: {
      from: formatDateDisplay(minDate),
      to: formatDateDisplay(maxDate),
    },
    rowCount: visits.length,
    visits,
  };
}

export { formatDateDisplay, parseDDMMYYYY };
