/**
 * Vital Score Card — Historical Data Storage
 *
 * Maintains a single Excel file (VITAL_SCORE_HISTORY.xlsx) on SharePoint in the
 * SCORE CARDS root folder. Each row = one unique visit.
 *
 * Deduplication key: repEmail|store|visitDate (composite, case-insensitive)
 * This matches the same deduplication logic used in the scorecard builder
 * (latest visit per rep+store combination wins within a single submission).
 */

import ExcelJS from 'exceljs';
import { QUESTIONS } from '@/constants/questions';
import { downloadHistoryFile, uploadHistoryFile } from '@/lib/sharepoint';
import type { StoreVisit } from '@/types';

// ── Score logic (mirrors calcStoreScores in excel-builder.ts) ─────────────────

function calcScore(visit: StoreVisit): { score: number; outOf: number; pct: number } {
  let score = 0;
  let outOf = 0;
  for (const q of QUESTIONS) {
    const answer = visit.answers[q.id] ?? null;
    if (!answer || answer.toLowerCase() === 'n/a') continue; // no answer or N/A → skip
    const isYes = answer.toLowerCase() === 'yes';
    const good = q.inverted ? !isYes : isYes;
    score += good ? 1 : 0;
    outOf += 1;
  }
  return { score, outOf, pct: outOf > 0 ? score / outOf : 0 };
}

// ── Visit ID ──────────────────────────────────────────────────────────────────

function buildVisitId(v: StoreVisit): string {
  // Prefer UUID as the canonical ID; fall back to composite key
  if (v.visitUUID) return v.visitUUID.toLowerCase();
  return `${v.repEmail.toLowerCase()}|${v.store.toLowerCase()}|${v.date}`;
}

// ── Deduplication (same rule as excel-builder) ────────────────────────────────

function deduplicateVisits(visits: StoreVisit[]): StoreVisit[] {
  const map = new Map<string, StoreVisit>();
  for (const v of visits) {
    const key = `${v.repEmail.toLowerCase()}|${v.store.toLowerCase()}`;
    const existing = map.get(key);
    if (!existing || v.dateObj > existing.dateObj) map.set(key, v);
  }
  return Array.from(map.values());
}

// ── Column definitions ────────────────────────────────────────────────────────

const FIXED_HEADERS = [
  'VisitID',
  'VisitUUID',
  'SubmittedAt',
  'VisitDate',
  'RepEmail',
  'RepFirstName',
  'RepLastName',
  'Province',
  'Channel',
  'StoreName',
  'StoreCode',
  'Customer',
  'Score',
  'OutOf',
  'ScorePct',
];

// Full header row: fixed fields + one column per question + overall comment
const ALL_HEADERS = [
  ...FIXED_HEADERS,
  ...QUESTIONS.map((q) => q.id),
  'OverallComment',
];

function visitToRow(
  v: StoreVisit,
  submittedAt: string
): (string | number)[] {
  const { score, outOf, pct } = calcScore(v);
  return [
    buildVisitId(v),
    v.visitUUID ?? '',
    submittedAt,
    v.date,
    v.repEmail,
    v.repFirstName,
    v.repLastName,
    v.province,
    v.channel,
    v.store,
    v.storeCode,
    v.customer,
    score,
    outOf,
    Math.round(pct * 100), // 0–100 integer
    ...QUESTIONS.map((q) => v.answers[q.id] ?? 'N/A'),
    v.overallComment ?? '',
  ];
}

// ── Main export ───────────────────────────────────────────────────────────────

export interface HistoryResult {
  added: number;
  skipped: number;
}

/**
 * Appends new visits to the SharePoint history file, skipping any visit whose
 * ID (repEmail|store|visitDate) already exists in the file.
 *
 * @param visits        Filtered visits from the current submission (pre-channel/rep filter)
 * @param submissionDate  Display date string e.g. "08 Mar 2026"
 */
export async function appendToHistory(
  visits: StoreVisit[],
  submissionDate: string
): Promise<HistoryResult> {
  // Apply same deduplication as scorecard builder before storing
  const deduped = deduplicateVisits(visits);

  const submittedAt = new Date().toISOString();
  const wb = new ExcelJS.Workbook();
  const existingIds = new Set<string>();

  // ── Load or create the history workbook ──────────────────────────────────
  const existingBuffer = await downloadHistoryFile();

  if (existingBuffer) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await wb.xlsx.load(existingBuffer as any);
    const ws = wb.worksheets[0];
    if (ws) {
      ws.eachRow((row, rowNum) => {
        if (rowNum === 1) return; // skip header
        const cell = row.getCell(1).value;
        if (typeof cell === 'string' && cell) existingIds.add(cell);
      });
    }
  } else {
    // First run — create the workbook with a styled header row
    const ws = wb.addWorksheet('Visit History');
    ws.views = [{ state: 'frozen', ySplit: 1, showGridLines: true }];

    const headerRow = ws.addRow(ALL_HEADERS);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFDA291C' }, // Vital red
    };
    headerRow.height = 18;

    // Set sensible column widths
    ws.getColumn(1).width  = 42; // VisitID
    ws.getColumn(2).width  = 38; // VisitUUID
    ws.getColumn(3).width  = 22; // SubmittedAt
    ws.getColumn(4).width  = 12; // VisitDate
    ws.getColumn(5).width  = 28; // RepEmail
    ws.getColumn(6).width  = 14; // RepFirstName
    ws.getColumn(7).width  = 14; // RepLastName
    ws.getColumn(8).width  = 14; // Province
    ws.getColumn(9).width  = 16; // Channel
    ws.getColumn(10).width = 28; // StoreName
    ws.getColumn(11).width = 12; // StoreCode
    ws.getColumn(12).width = 20; // Customer
    ws.getColumn(13).width = 8;  // Score
    ws.getColumn(14).width = 8;  // OutOf
    ws.getColumn(15).width = 10; // ScorePct
    // Question columns
    for (let i = 16; i <= 15 + QUESTIONS.length; i++) {
      ws.getColumn(i).width = 12;
    }
    // OverallComment
    ws.getColumn(16 + QUESTIONS.length).width = 40;
  }

  // ── Append new rows ───────────────────────────────────────────────────────
  const ws = wb.worksheets[0];
  let added = 0;
  let skipped = 0;

  for (const v of deduped) {
    const visitId = buildVisitId(v);
    if (existingIds.has(visitId)) {
      skipped++;
      continue;
    }
    ws.addRow(visitToRow(v, submittedAt));
    added++;
  }

  // Only write back if something was actually added
  if (added > 0) {
    const raw = await wb.xlsx.writeBuffer();
    await uploadHistoryFile(new Uint8Array(raw as ArrayBuffer));
  }

  return { added, skipped };
}
