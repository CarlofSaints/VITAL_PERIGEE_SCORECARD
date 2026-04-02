/**
 * Vital Score Card — Comparison Report Builder
 *
 * Reads rows from the history Excel buffer and produces a 5-sheet comparison workbook:
 *   Sheet 1 (Field Summary) — Per-question aggregate performance across all visits
 *   Sheet 2 (Summary)       — Channel × Province breakdown with Vital logo
 *   Sheet 3 (By Rep)        — Reps as rows, periods as columns
 *   Sheet 4 (By Store)      — Stores as rows (with Channel + Province), periods as columns
 *   Sheet 5 (By Question)   — Flat table per (question × rep × channel × province × store)
 */

import ExcelJS from 'exceljs';
import { readFileSync } from 'fs';
import { join } from 'path';
import { QUESTIONS } from '@/constants/questions';

export type Period = 'weekly' | 'monthly' | 'quarterly';

// ── Date helpers ─────────────────────────────────────────────────────────────

function parseDDMMYYYY(s: string): Date | null {
  if (!s) return null;
  const parts = s.split('/');
  if (parts.length === 3) {
    const [d, m, y] = parts.map(Number);
    if (!isNaN(d) && !isNaN(m) && !isNaN(y)) return new Date(y, m - 1, d);
  }
  // Try ISO
  const iso = new Date(s);
  return isNaN(iso.getTime()) ? null : iso;
}

/** Next Sunday on or after `d` (week-ending label) */
function nextSunday(d: Date): Date {
  const day = d.getDay(); // 0=Sun
  const diff = day === 0 ? 0 : 7 - day;
  const s = new Date(d);
  s.setDate(s.getDate() + diff);
  return s;
}

function formatDDMMMYYYY(d: Date): string {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

function periodKey(d: Date, period: Period): string {
  if (period === 'weekly') {
    const s = nextSunday(d);
    return s.toISOString().slice(0, 10); // YYYY-MM-DD for sorting
  }
  if (period === 'monthly') {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }
  // quarterly
  const q = Math.floor(d.getMonth() / 3) + 1;
  return `${d.getFullYear()}-Q${q}`;
}

function periodLabel(key: string, period: Period): string {
  if (period === 'weekly') {
    // key is YYYY-MM-DD of the Sunday
    const [y, m, d] = key.split('-').map(Number);
    const dt = new Date(y, m - 1, d);
    return `Wk ending ${formatDDMMMYYYY(dt)}`;
  }
  if (period === 'monthly') {
    const [y, m] = key.split('-').map(Number);
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${months[m - 1]} ${y}`;
  }
  // quarterly: key = "YYYY-Q#"
  const [y, q] = key.split('-');
  return `${q} ${y}`;
}

/** Returns the period label, annotated with "(Xd missing)" if the data doesn't cover the full period. */
function periodLabelWithNote(
  key: string,
  period: Period,
  dateRange: { minDate: Date; maxDate: Date } | undefined
): string {
  const base = periodLabel(key, period);
  if (!dateRange) return base;

  let periodStart: Date;
  let periodEnd: Date;

  if (period === 'weekly') {
    const [y, m, d] = key.split('-').map(Number);
    periodEnd = new Date(y, m - 1, d); // Sunday
    periodStart = new Date(periodEnd);
    periodStart.setDate(periodStart.getDate() - 6); // Monday
  } else if (period === 'monthly') {
    const [y, m] = key.split('-').map(Number);
    periodStart = new Date(y, m - 1, 1);
    periodEnd = new Date(y, m, 0); // last day of month
  } else {
    // quarterly: "YYYY-Q#"
    const [ys, qs] = key.split('-');
    const yr = parseInt(ys);
    const qn = parseInt(qs.slice(1));
    periodStart = new Date(yr, (qn - 1) * 3, 1);
    periodEnd = new Date(yr, qn * 3, 0); // last day of quarter
  }

  // Cap at today — future days of an ongoing period are not "missing"
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const periodEnded = periodEnd <= today;
  const effectiveEnd = periodEnded ? periodEnd : today;

  let missing = 0;
  if (dateRange.minDate > periodStart) {
    missing += Math.round((dateRange.minDate.getTime() - periodStart.getTime()) / 86400000);
  }
  if (dateRange.maxDate < effectiveEnd) {
    missing += Math.round((effectiveEnd.getTime() - dateRange.maxDate.getTime()) / 86400000);
  }
  if (missing > 0) return `${base}\n(${missing}d missing)`;
  if (!periodEnded) return `${base}\n(ongoing)`;
  return base;
}

// ── History row shape ─────────────────────────────────────────────────────────
// Matches columns written by history.ts FIXED_HEADERS (1-indexed):
// 1=VisitID, 2=VisitUUID, 3=SubmittedAt, 4=VisitDate, 5=RepEmail,
// 6=RepFirstName, 7=RepLastName, 8=Province, 9=Channel, 10=StoreName,
// 11=StoreCode, 12=Customer, 13=Score, 14=OutOf, 15=ScorePct, ...questions

interface HistoryRow {
  visitDate: Date;
  repEmail: string;
  repName: string;
  province: string;
  channel: string;
  store: string;
  scorePct: number; // 0–100
}

function parseHistoryBuffer(buf: Buffer): HistoryRow[] {
  // We'll parse using ExcelJS (same lib used to write)
  // But ExcelJS load is async — caller must await buildComparisonReport
  // We handle this inside the async function below.
  // This function is a placeholder — actual parsing in buildComparisonReport.
  void buf;
  return [];
}
void parseHistoryBuffer; // suppress unused warning

// ── Styling constants ─────────────────────────────────────────────────────────

const RED     = 'FFDA291C';
const WHITE   = 'FFFFFFFF';
const DARK    = 'FF32373C';
const LIGHT_RED = 'FFFCE8E6';
const GRAY_BG = 'FFF5F5F5';
const HEADER_FONT = { bold: true, color: { argb: WHITE }, size: 10 };
const RED_FILL = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: RED } };
const GRAY_FILL = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: GRAY_BG } };

function pctColor(pct: number): string {
  if (pct >= 80) return 'FF1B5E20'; // dark green
  if (pct >= 60) return 'FF4CAF50'; // green
  if (pct >= 40) return 'FFF57F17'; // amber
  return 'FFB71C1C'; // dark red
}

function pctBg(pct: number): string {
  if (pct >= 80) return 'FFE8F5E9';
  if (pct >= 60) return 'FFF1F8E9';
  if (pct >= 40) return 'FFFFF9C4';
  return LIGHT_RED;
}

/** N/A-aware per-question score helper (returns 0/0 for N/A — caller checks outOf > 0) */
function getScoreQ(answer: string | null, inverted: boolean): { score: number; outOf: number } {
  if (!answer || answer.toLowerCase() === 'n/a') return { score: 0, outOf: 0 };
  const isYes = answer.toLowerCase() === 'yes';
  const good = inverted ? !isYes : isYes;
  return { score: good ? 1 : 0, outOf: 1 };
}

/** Convert 1-based column number to Excel letter (1=A, 27=AA, …) */
function numToColLetter(n: number): string {
  let r = '';
  while (n > 0) { r = String.fromCharCode(65 + ((n - 1) % 26)) + r; n = Math.floor((n - 1) / 26); }
  return r;
}

// ── Main export ───────────────────────────────────────────────────────────────

export interface ComparisonFilters {
  channels?:  string[];
  provinces?: string[];
  reps?:      string[]; // rep emails
  stores?:    string[]; // store names
}

export interface ComparisonInput {
  historyBuffer: Buffer;
  from: Date;
  to: Date;
  period: Period;
  dateRangeLabel: string; // e.g. "01 Jan 2026 – 08 Mar 2026"
  filters?: ComparisonFilters;
}

// ── Filter options extractor ──────────────────────────────────────────────────

export interface FilterOptions {
  channels:  string[];
  provinces: string[];
  reps:      { email: string; name: string }[];
  stores:    { name: string; channel: string }[];
}

export async function extractFilterOptions(historyBuffer: Buffer): Promise<FilterOptions> {
  const srcWb = new ExcelJS.Workbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await srcWb.xlsx.load(historyBuffer as any);
  const srcWs = srcWb.worksheets[0];
  if (!srcWs) return { channels: [], provinces: [], reps: [], stores: [] };

  const hdrMap: Record<string, number> = {};
  srcWs.getRow(1).eachCell((cell, colNum) => {
    const v = String(cell.value ?? '').trim();
    if (v) hdrMap[v] = colNum;
  });

  const channels  = new Set<string>();
  const provinces = new Set<string>();
  const repMap    = new Map<string, string>(); // email → display name
  const storeMap  = new Map<string, string>(); // storeName → channel

  srcWs.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const channel   = String(row.getCell(hdrMap['Channel']      ?? 9 ).value ?? '').trim();
    const province  = String(row.getCell(hdrMap['Province']     ?? 8 ).value ?? '').trim();
    const repEmail  = String(row.getCell(hdrMap['RepEmail']     ?? 5 ).value ?? '').trim().toLowerCase();
    const firstName = String(row.getCell(hdrMap['RepFirstName'] ?? 6 ).value ?? '').trim();
    const lastName  = String(row.getCell(hdrMap['RepLastName']  ?? 7 ).value ?? '').trim();
    const store     = String(row.getCell(hdrMap['StoreName']    ?? 10).value ?? '').trim();
    if (channel)          channels.add(channel);
    if (province)         provinces.add(province);
    if (repEmail)         repMap.set(repEmail, `${firstName} ${lastName}`.trim() || repEmail);
    if (store && channel) storeMap.set(store, channel);
  });

  return {
    channels:  Array.from(channels).sort(),
    provinces: Array.from(provinces).sort(),
    reps:      Array.from(repMap.entries())
                 .map(([email, name]) => ({ email, name }))
                 .sort((a, b) => a.name.localeCompare(b.name)),
    stores:    Array.from(storeMap.entries())
                 .map(([name, channel]) => ({ name, channel }))
                 .sort((a, b) => a.name.localeCompare(b.name)),
  };
}

export async function buildComparisonReport(input: ComparisonInput): Promise<Buffer> {
  const { historyBuffer, from, to, period, dateRangeLabel, filters } = input;

  // ── Parse history workbook ────────────────────────────────────────────────
  const srcWb = new ExcelJS.Workbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await srcWb.xlsx.load(historyBuffer as any);
  const srcWs = srcWb.worksheets[0];
  if (!srcWs) throw new Error('History file has no worksheets.');

  // Map header names to column indices (1-based)
  const hdrMap: Record<string, number> = {};
  const hdrRow = srcWs.getRow(1);
  hdrRow.eachCell((cell, colNum) => {
    const v = String(cell.value ?? '').trim();
    if (v) hdrMap[v] = colNum;
  });

  const rows: HistoryRow[] = [];

  // Per-question aggregation for "By Question" sheet
  // Key: `${questionId}||${repEmail}||${channel}||${province}||${store}`
  type QDimKey = string;
  const qDimPeriod = new Map<QDimKey, Map<string, { score: number; outOf: number }>>();
  const qDimMeta = new Map<QDimKey, {
    section: string; qText: string; repName: string; channel: string; province: string; store: string;
    overall: { score: number; outOf: number };
  }>();
  // Aggregate per-question data (all visits, no dimension breakdown) — used by Field Summary
  const qAggPeriod = new Map<string, Map<string, { score: number; outOf: number }>>();
  const qAggOverall = new Map<string, { score: number; outOf: number }>();
  // Visit count per period — shown in Field Summary column sub-header
  const periodVisitCount = new Map<string, number>();

  srcWs.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const dateStr = String(row.getCell(hdrMap['VisitDate'] ?? 4).value ?? '').trim();
    const d = parseDDMMYYYY(dateStr);
    if (!d) return;
    if (d < from || d > to) return;

    const repEmail = String(row.getCell(hdrMap['RepEmail'] ?? 5).value ?? '').trim().toLowerCase();
    const firstName = String(row.getCell(hdrMap['RepFirstName'] ?? 6).value ?? '').trim();
    const lastName  = String(row.getCell(hdrMap['RepLastName']  ?? 7).value ?? '').trim();
    const repName   = `${firstName} ${lastName}`.trim() || repEmail;
    const province  = String(row.getCell(hdrMap['Province'] ?? 8).value ?? '').trim();
    const channel   = String(row.getCell(hdrMap['Channel']  ?? 9).value ?? '').trim();
    const store     = String(row.getCell(hdrMap['StoreName'] ?? 10).value ?? '').trim();
    const pk = periodKey(d, period);

    // Apply dimension filters — skip row if it doesn't match selected filters
    if (filters?.channels?.length  && !filters.channels.includes(channel))  return;
    if (filters?.provinces?.length && !filters.provinces.includes(province)) return;
    if (filters?.reps?.length      && !filters.reps.includes(repEmail))      return;
    if (filters?.stores?.length    && !filters.stores.includes(store))       return;

    periodVisitCount.set(pk, (periodVisitCount.get(pk) ?? 0) + 1);

    // Recompute scorePct from per-question columns (N/A-aware; ignores stored ScorePct column)
    let totalScore = 0, totalOutOf = 0;
    for (const q of QUESTIONS) {
      const qColIdx = hdrMap[q.id];
      if (!qColIdx) continue;
      const ans = String(row.getCell(qColIdx).value ?? '').trim() || null;
      const { score: qScore, outOf: qOutOf } = getScoreQ(ans, q.inverted ?? false);
      totalScore += qScore;
      totalOutOf += qOutOf;

      // Accumulate per-question data for By Question sheet
      if (qOutOf > 0) {
        const qKey: QDimKey = `${q.id}||${repEmail}||${channel}||${province}||${store}`;
        if (!qDimPeriod.has(qKey)) qDimPeriod.set(qKey, new Map());
        const qPMap = qDimPeriod.get(qKey)!;
        const qPVal = qPMap.get(pk) ?? { score: 0, outOf: 0 };
        qPVal.score += qScore;
        qPVal.outOf += qOutOf;
        qPMap.set(pk, qPVal);

        if (!qDimMeta.has(qKey)) {
          qDimMeta.set(qKey, {
            section: q.section, qText: q.text, repName, channel, province, store,
            overall: { score: 0, outOf: 0 },
          });
        }
        const qm = qDimMeta.get(qKey)!;
        qm.overall.score += qScore;
        qm.overall.outOf += qOutOf;
        qm.repName = repName;

        // Accumulate for Field Summary aggregate
        if (!qAggPeriod.has(q.id)) qAggPeriod.set(q.id, new Map());
        const qAggPMap = qAggPeriod.get(q.id)!;
        const qAggPVal = qAggPMap.get(pk) ?? { score: 0, outOf: 0 };
        qAggPVal.score += qScore; qAggPVal.outOf += qOutOf;
        qAggPMap.set(pk, qAggPVal);
        const qAggOVal = qAggOverall.get(q.id) ?? { score: 0, outOf: 0 };
        qAggOVal.score += qScore; qAggOVal.outOf += qOutOf;
        qAggOverall.set(q.id, qAggOVal);
      }
    }
    const scorePct = totalOutOf > 0 ? Math.round((totalScore / totalOutOf) * 100) : 0;

    rows.push({ visitDate: d, repEmail, repName, province, channel, store, scorePct });
  });

  if (rows.length === 0) throw new Error('No history rows found in the selected date range.');

  // ── Build period buckets ──────────────────────────────────────────────────
  // Collect all period keys in sorted order
  const periodKeySet = new Set<string>();
  for (const r of rows) periodKeySet.add(periodKey(r.visitDate, period));
  const sortedPeriods = Array.from(periodKeySet).sort();

  // ── Aggregate data ────────────────────────────────────────────────────────

  // Rep × period → { totalPct, count }
  const repPeriodMap = new Map<string, Map<string, { totalPct: number; count: number }>>();
  // Store × period → { totalPct, count }
  const storePeriodMap = new Map<string, Map<string, { totalPct: number; count: number }>>();
  // Channel × Province → { totalPct, count }
  const channelProvinceMap = new Map<string, { totalPct: number; count: number; channel: string; province: string }>();
  // Channel → totals
  const channelTotals = new Map<string, { totalPct: number; count: number }>();
  // Province → totals
  const provinceTotals = new Map<string, { totalPct: number; count: number }>();
  // Province × period → { totalPct, count }
  const provincePeriodMap = new Map<string, Map<string, { totalPct: number; count: number }>>();
  // Channel × period → { totalPct, count }
  const channelPeriodMap = new Map<string, Map<string, { totalPct: number; count: number }>>();
  // Channel × Province × period → { totalPct, count }
  const cpPeriodMap = new Map<string, Map<string, { totalPct: number; count: number }>>();
  // Period → actual min/max visit dates (for partial period detection)
  const periodDateRange = new Map<string, { minDate: Date; maxDate: Date }>();
  // Store metadata
  const storeMeta = new Map<string, { channel: string; province: string }>();
  // Rep metadata
  const repMeta = new Map<string, { name: string }>();

  for (const r of rows) {
    const pk = periodKey(r.visitDate, period);

    // Rep
    if (!repPeriodMap.has(r.repEmail)) repPeriodMap.set(r.repEmail, new Map());
    const repMap = repPeriodMap.get(r.repEmail)!;
    const repPeriod = repMap.get(pk) ?? { totalPct: 0, count: 0 };
    repPeriod.totalPct += r.scorePct;
    repPeriod.count++;
    repMap.set(pk, repPeriod);
    repMeta.set(r.repEmail, { name: r.repName });

    // Store
    const storeKey = `${r.store}||${r.channel}`;
    if (!storePeriodMap.has(storeKey)) storePeriodMap.set(storeKey, new Map());
    const storeMap = storePeriodMap.get(storeKey)!;
    const storePeriod = storeMap.get(pk) ?? { totalPct: 0, count: 0 };
    storePeriod.totalPct += r.scorePct;
    storePeriod.count++;
    storeMap.set(pk, storePeriod);
    storeMeta.set(storeKey, { channel: r.channel, province: r.province });

    // Channel × Province
    const cpKey = `${r.channel}||${r.province}`;
    const cpVal = channelProvinceMap.get(cpKey) ?? { totalPct: 0, count: 0, channel: r.channel, province: r.province };
    cpVal.totalPct += r.scorePct;
    cpVal.count++;
    channelProvinceMap.set(cpKey, cpVal);

    // Channel totals
    const ctVal = channelTotals.get(r.channel) ?? { totalPct: 0, count: 0 };
    ctVal.totalPct += r.scorePct;
    ctVal.count++;
    channelTotals.set(r.channel, ctVal);

    // Province totals
    const ptVal = provinceTotals.get(r.province) ?? { totalPct: 0, count: 0 };
    ptVal.totalPct += r.scorePct;
    ptVal.count++;
    provinceTotals.set(r.province, ptVal);

    // Province × period
    if (!provincePeriodMap.has(r.province)) provincePeriodMap.set(r.province, new Map());
    const ppMap = provincePeriodMap.get(r.province)!;
    const ppVal = ppMap.get(pk) ?? { totalPct: 0, count: 0 };
    ppVal.totalPct += r.scorePct; ppVal.count++;
    ppMap.set(pk, ppVal);

    // Channel × period
    if (!channelPeriodMap.has(r.channel)) channelPeriodMap.set(r.channel, new Map());
    const chpMap = channelPeriodMap.get(r.channel)!;
    const chpVal = chpMap.get(pk) ?? { totalPct: 0, count: 0 };
    chpVal.totalPct += r.scorePct; chpVal.count++;
    chpMap.set(pk, chpVal);

    // Channel × Province × period
    if (!cpPeriodMap.has(cpKey)) cpPeriodMap.set(cpKey, new Map());
    const cpPMap = cpPeriodMap.get(cpKey)!;
    const cpPVal = cpPMap.get(pk) ?? { totalPct: 0, count: 0 };
    cpPVal.totalPct += r.scorePct; cpPVal.count++;
    cpPMap.set(pk, cpPVal);

    // Period date range (for partial period detection)
    const pdr = periodDateRange.get(pk);
    if (!pdr) {
      periodDateRange.set(pk, { minDate: new Date(r.visitDate), maxDate: new Date(r.visitDate) });
    } else {
      if (r.visitDate < pdr.minDate) pdr.minDate = new Date(r.visitDate);
      if (r.visitDate > pdr.maxDate) pdr.maxDate = new Date(r.visitDate);
    }
  }

  // Sort reps by name, stores by store name
  const sortedReps = Array.from(repPeriodMap.keys()).sort((a, b) =>
    (repMeta.get(a)?.name ?? a).localeCompare(repMeta.get(b)?.name ?? b)
  );
  const sortedStores = Array.from(storePeriodMap.keys()).sort((a, b) => {
    const [sa] = a.split('||');
    const [sb] = b.split('||');
    return sa.localeCompare(sb);
  });

  // ── Load logos ────────────────────────────────────────────────────────────
  let logoImageId: number | null = null;
  let perigeeImageId: number | null = null;
  const wb = new ExcelJS.Workbook();
  try {
    const logoData = readFileSync(join(process.cwd(), 'public', 'vital-logo.png'));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    logoImageId = wb.addImage({ buffer: logoData as any, extension: 'png' });
  } catch { /* logo optional */ }
  try {
    const perigeeData = readFileSync(join(process.cwd(), 'public', 'perigee-logo.jpg'));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    perigeeImageId = wb.addImage({ buffer: perigeeData as any, extension: 'jpeg' });
  } catch { /* logo optional */ }

  // Shared constants used by all sheets
  const NUM_PERIODS = sortedPeriods.length;
  const DARK_FILL = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: DARK } };
  const DARK_HDR_FONT = { bold: true, size: 10, color: { argb: WHITE } };
  const PW = period === 'weekly' ? 18 : 14; // period column width

  // ── Sheet 0: Field Summary ───────────────────────────────────────────────
  // Flat filterable table: Section | Question | Target | periods | Overall
  // AutoFilter on row 7 lets users filter by Section or Question text.
  const s0 = wb.addWorksheet('Field Summary');
  const NUM_FIXED_FS = 3; // Section | Question | Target
  const TOTAL_COLS_FS = NUM_FIXED_FS + NUM_PERIODS + 1;
  s0.views = [{ state: 'frozen', xSplit: 3, ySplit: 7, showGridLines: false }];

  // Logos
  if (logoImageId !== null) {
    s0.addImage(logoImageId, { tl: { col: 0, row: 0 }, ext: { width: 160, height: 52 } });
  }
  if (perigeeImageId !== null) {
    s0.addImage(perigeeImageId, { tl: { col: TOTAL_COLS_FS - 1, row: 0 }, ext: { width: 52, height: 52 } });
  }

  // Title block
  s0.getRow(1).height = 30;
  s0.getRow(2).height = 20;
  s0.getRow(3).height = 16;
  s0.getRow(4).height = 12;

  const s0TitleCell = s0.getCell('E1');
  s0TitleCell.value = 'VITAL SCORE CARD — FIELD QUESTION SUMMARY';
  s0TitleCell.font = { bold: true, size: 16, color: { argb: RED } };
  s0TitleCell.alignment = { vertical: 'middle' };

  const s0SubtitleCell = s0.getCell('E2');
  s0SubtitleCell.value = dateRangeLabel;
  s0SubtitleCell.font = { size: 11, color: { argb: DARK }, italic: true };

  const s0PeriodStr = period.charAt(0).toUpperCase() + period.slice(1);
  const s0PeriodCell = s0.getCell('E3');
  s0PeriodCell.value = `Period: ${s0PeriodStr}  ·  ${rows.length} visits  ·  ${sortedPeriods.length} period${sortedPeriods.length !== 1 ? 's' : ''}  ·  Use dropdown filters on row 7 to filter by Section / Question  ·  Filter by Rep/Channel/Province/Store on the 'By Question' tab`;
  s0PeriodCell.font = { size: 9, color: { argb: '888888' } };

  // Red divider row 5
  s0.getRow(5).height = 4;
  for (let c = 1; c <= TOTAL_COLS_FS; c++) {
    s0.getCell(5, c).fill = RED_FILL;
  }

  // Merged red header row 6
  s0.mergeCells(6, 1, 6, TOTAL_COLS_FS);
  const s0HeadCell = s0.getCell(6, 1);
  s0HeadCell.value = 'RETAIL SCORECARD — FIELD QUESTION SUMMARY';
  s0HeadCell.font = { bold: true, size: 12, color: { argb: WHITE } };
  s0HeadCell.fill = RED_FILL;
  s0HeadCell.alignment = { horizontal: 'center', vertical: 'middle' };
  s0.getRow(6).height = 24;

  // Column headers row 7 — AutoFilter enabled
  const s0HdrRow = s0.getRow(7);
  ['Section', 'Question', 'Target'].forEach((h, i) => {
    const cell = s0HdrRow.getCell(i + 1);
    cell.value = h;
    cell.font = DARK_HDR_FONT;
    cell.fill = DARK_FILL;
    cell.alignment = { horizontal: i < 2 ? 'left' : 'center', vertical: 'middle' };
  });
  sortedPeriods.forEach((pk, i) => {
    const vc = periodVisitCount.get(pk) ?? 0;
    const label = periodLabelWithNote(pk, period, periodDateRange.get(pk));
    const cell = s0HdrRow.getCell(NUM_FIXED_FS + i + 1);
    cell.value = `${label}\n(${vc} visit${vc !== 1 ? 's' : ''})`;
    cell.font = DARK_HDR_FONT;
    cell.fill = DARK_FILL;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  const s0OvHdr = s0HdrRow.getCell(TOTAL_COLS_FS);
  s0OvHdr.value = 'Overall';
  s0OvHdr.font = DARK_HDR_FONT;
  s0OvHdr.fill = DARK_FILL;
  s0OvHdr.alignment = { horizontal: 'center', vertical: 'middle' };
  s0HdrRow.height = period === 'weekly' ? 52 : 36;

  // AutoFilter on header row — Section (col A) and Question (col B) dropdowns
  s0.autoFilter = `A7:${numToColLetter(TOTAL_COLS_FS)}7`;

  // Data rows — flat table, no merged section headers; Section repeats in col A
  let s0RowIdx = 8;
  const s0Sections = Array.from(new Set(QUESTIONS.map((q) => q.section)));
  for (const section of s0Sections) {
    const sectionQs = QUESTIONS.filter((q) => q.section === section);

    for (const q of sectionQs) {
      const rowBg = s0RowIdx % 2 === 0 ? 'FFFAFAFA' : WHITE;
      const bgFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: rowBg } };
      const row = s0.getRow(s0RowIdx);

      row.getCell(1).value = section;
      row.getCell(1).font = { bold: true, size: 9, color: { argb: DARK } };
      row.getCell(1).fill = bgFill;
      row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };

      row.getCell(2).value = q.text;
      row.getCell(2).font = { size: 9, color: { argb: DARK } };
      row.getCell(2).fill = bgFill;
      row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

      row.getCell(3).value = q.inverted ? 'No' : 'Yes';
      row.getCell(3).font = { bold: true, size: 9, color: { argb: q.inverted ? RED : 'FF1B5E20' } };
      row.getCell(3).fill = bgFill;
      row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };

      const qAggPMap = qAggPeriod.get(q.id);
      const qAggOVal = qAggOverall.get(q.id);

      sortedPeriods.forEach((pk, pi) => {
        const cell = row.getCell(NUM_FIXED_FS + pi + 1);
        const bucket = qAggPMap?.get(pk);
        if (bucket && bucket.outOf > 0) {
          const avg = Math.round((bucket.score / bucket.outOf) * 100);
          cell.value = avg;
          cell.numFmt = '0"%"';
          cell.font = { bold: true, size: 10, color: { argb: pctColor(avg) } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(avg) } };
        } else {
          cell.value = q.optional ? 'N/A' : '—';
          cell.font = { size: 9, color: { argb: 'AAAAAA' } };
          cell.fill = bgFill;
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      const s0OvCell = row.getCell(TOTAL_COLS_FS);
      if (qAggOVal && qAggOVal.outOf > 0) {
        const avg = Math.round((qAggOVal.score / qAggOVal.outOf) * 100);
        s0OvCell.value = avg;
        s0OvCell.numFmt = '0"%"';
        s0OvCell.font = { bold: true, size: 10, color: { argb: pctColor(avg) } };
        s0OvCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(avg) } };
      } else {
        s0OvCell.value = q.optional ? 'N/A' : '—';
        s0OvCell.font = { size: 9, color: { argb: 'AAAAAA' } };
        s0OvCell.fill = bgFill;
      }
      s0OvCell.alignment = { horizontal: 'center', vertical: 'middle' };
      row.height = 20;
      s0RowIdx++;
    }
  }

  // Column widths for Field Summary
  s0.getColumn(1).width = 24; // Section
  s0.getColumn(2).width = 44; // Question
  s0.getColumn(3).width = 8;  // Target
  sortedPeriods.forEach((_, i) => {
    s0.getColumn(NUM_FIXED_FS + i + 1).width = period === 'weekly' ? 18 : 14;
  });
  s0.getColumn(TOTAL_COLS_FS).width = 13; // Overall

  // ── Sheet 1: Summary ─────────────────────────────────────────────────────
  const s1 = wb.addWorksheet('Summary');
  s1.views = [{ state: 'normal', showGridLines: false }];

  // Vital logo top-left
  if (logoImageId !== null) {
    s1.addImage(logoImageId, { tl: { col: 0, row: 0 }, ext: { width: 160, height: 52 } });
  }
  // Perigee logo top-right — square to preserve aspect ratio
  if (perigeeImageId !== null) {
    s1.addImage(perigeeImageId, { tl: { col: 8, row: 0 }, ext: { width: 52, height: 52 } });
  }

  // Title block
  s1.getRow(1).height = 30;
  s1.getRow(2).height = 20;
  s1.getRow(3).height = 16;
  s1.getRow(4).height = 12;

  const titleCell = s1.getCell('D1');
  titleCell.value = 'VITAL SCORE CARD — COMPARISON REPORT';
  titleCell.font = { bold: true, size: 16, color: { argb: RED } };
  titleCell.alignment = { vertical: 'middle' };

  const subtitleCell = s1.getCell('D2');
  subtitleCell.value = dateRangeLabel;
  subtitleCell.font = { size: 11, color: { argb: DARK }, italic: true };

  const periodCell = s1.getCell('D3');
  const periodStr = period.charAt(0).toUpperCase() + period.slice(1);
  periodCell.value = `Period: ${periodStr}  ·  ${rows.length} visits  ·  ${sortedPeriods.length} period${sortedPeriods.length !== 1 ? 's' : ''}`;
  periodCell.font = { size: 9, color: { argb: '888888' } };

  // Red divider row 5
  s1.getRow(5).height = 4;
  const MAX_S1_COLS = 2 + NUM_PERIODS + 1; // Channel+Province + periods + Overall
  for (let c = 1; c <= MAX_S1_COLS; c++) {
    s1.getCell(5, c).fill = RED_FILL;
  }

  // ── Shared helper: render one period-breakdown table ─────────────────────
  // Returns the row number after a 2-row gap (ready for next table).
  function writeSummaryBlock(
    startRow: number,
    heading: string,
    fixedHeaders: string[],
    tableRows: {
      labels: string[];
      periodMap: Map<string, { totalPct: number; count: number }>;
      overall: { totalPct: number; count: number };
    }[],
  ): number {
    const F = fixedHeaders.length;
    const TOTAL_COLS = F + NUM_PERIODS + 1;

    // Merged heading
    s1.mergeCells(startRow, 1, startRow, TOTAL_COLS);
    const headCell = s1.getCell(startRow, 1);
    headCell.value = heading;
    headCell.font = HEADER_FONT;
    headCell.fill = RED_FILL;
    headCell.alignment = { horizontal: 'center', vertical: 'middle' };
    s1.getRow(startRow).height = 20;
    startRow++;

    // Column headers
    const hdrRow = s1.getRow(startRow);
    fixedHeaders.forEach((h, i) => {
      const cell = hdrRow.getCell(i + 1);
      cell.value = h;
      cell.font = DARK_HDR_FONT;
      cell.fill = DARK_FILL;
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
    });
    sortedPeriods.forEach((pk, i) => {
      const cell = hdrRow.getCell(F + i + 1);
      cell.value = periodLabelWithNote(pk, period, periodDateRange.get(pk));
      cell.font = DARK_HDR_FONT;
      cell.fill = DARK_FILL;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    const ovHdrCell = hdrRow.getCell(F + NUM_PERIODS + 1);
    ovHdrCell.value = 'Overall';
    ovHdrCell.font = DARK_HDR_FONT;
    ovHdrCell.fill = DARK_FILL;
    ovHdrCell.alignment = { horizontal: 'center', vertical: 'middle' };
    hdrRow.height = period === 'weekly' ? 38 : 22;
    startRow++;

    // Data rows
    for (const tr of tableRows) {
      const row = s1.getRow(startRow);
      tr.labels.forEach((lbl, i) => {
        row.getCell(i + 1).value = lbl;
        row.getCell(i + 1).font = i === 0 ? { bold: true, size: 10 } : { size: 10 };
      });
      sortedPeriods.forEach((pk, i) => {
        const cell = row.getCell(F + i + 1);
        const bucket = tr.periodMap.get(pk);
        if (bucket && bucket.count > 0) {
          const avg = Math.round(bucket.totalPct / bucket.count);
          cell.value = avg;
          cell.numFmt = '0"%"';
          cell.font = { bold: true, size: 10, color: { argb: pctColor(avg) } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(avg) } };
        } else {
          cell.value = '—';
          cell.font = { size: 9, color: { argb: 'AAAAAA' } };
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      const ov = tr.overall;
      const ovAvg = ov.count > 0 ? Math.round(ov.totalPct / ov.count) : 0;
      const ovCell = row.getCell(F + NUM_PERIODS + 1);
      if (ov.count > 0) {
        ovCell.value = ovAvg;
        ovCell.numFmt = '0"%"';
        ovCell.font = { bold: true, size: 10, color: { argb: pctColor(ovAvg) } };
        ovCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(ovAvg) } };
      } else {
        ovCell.value = '—';
        ovCell.font = { size: 9, color: { argb: 'AAAAAA' } };
      }
      ovCell.alignment = { horizontal: 'center', vertical: 'middle' };
      row.height = 18;
      startRow++;
    }
    return startRow + 2; // 2-row gap
  }

  // ── Sort summary data ──────────────────────────────────────────────────────
  const sortedProvinces = Array.from(provinceTotals.entries())
    .sort((a, b) => a[0].localeCompare(b[0]));
  const sortedChannels = Array.from(channelTotals.entries())
    .sort((a, b) => a[0].localeCompare(b[0]));
  const sortedCPs = Array.from(channelProvinceMap.values())
    .sort((a, b) => a.channel.localeCompare(b.channel) || a.province.localeCompare(b.province));

  // ── Render three tables vertically ────────────────────────────────────────
  let currentRow = 7;

  currentRow = writeSummaryBlock(
    currentRow,
    'Scorecard Comparison by Province',
    ['Province'],
    sortedProvinces.map(([provName, overall]) => ({
      labels: [provName],
      periodMap: provincePeriodMap.get(provName) ?? new Map(),
      overall,
    })),
  );

  currentRow = writeSummaryBlock(
    currentRow,
    'Scorecard Comparison by Channel',
    ['Channel'],
    sortedChannels.map(([chName, overall]) => ({
      labels: [chName],
      periodMap: channelPeriodMap.get(chName) ?? new Map(),
      overall,
    })),
  );

  writeSummaryBlock(
    currentRow,
    'Channel-Province Score Card Comparison',
    ['Channel', 'Province'],
    sortedCPs.map((cp) => ({
      labels: [cp.channel, cp.province],
      periodMap: cpPeriodMap.get(`${cp.channel}||${cp.province}`) ?? new Map(),
      overall: { totalPct: cp.totalPct, count: cp.count },
    })),
  );

  // Column widths: col 1 = primary label, col 2 = secondary label or 1st period, rest = periods + overall
  s1.getColumn(1).width = 24;
  s1.getColumn(2).width = 20;
  for (let c = 3; c <= 2 + NUM_PERIODS + 1; c++) {
    s1.getColumn(c).width = c === 2 + NUM_PERIODS + 1 ? 13 : PW; // last = Overall
  }

  // ── Sheet 2: By Rep ───────────────────────────────────────────────────────
  const s2 = wb.addWorksheet('By Rep');
  s2.views = [{ state: 'frozen', xSplit: 1, ySplit: 1, showGridLines: false }];

  // Header row
  const s2Hdr = s2.getRow(1);
  s2Hdr.getCell(1).value = 'Rep Name';
  s2Hdr.getCell(1).font = HEADER_FONT;
  s2Hdr.getCell(1).fill = RED_FILL;
  s2Hdr.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
  sortedPeriods.forEach((pk, i) => {
    const cell = s2Hdr.getCell(i + 2);
    cell.value = periodLabel(pk, period);
    cell.font = HEADER_FONT;
    cell.fill = RED_FILL;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  s2Hdr.height = period === 'weekly' ? 36 : 20;

  // Data rows
  sortedReps.forEach((repEmail, ri) => {
    const row = s2.getRow(ri + 2);
    row.getCell(1).value = repMeta.get(repEmail)?.name ?? repEmail;
    row.getCell(1).font = { size: 10 };
    row.getCell(1).fill = ri % 2 === 0 ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAFAFA' } } : { type: 'pattern', pattern: 'solid', fgColor: { argb: WHITE } };

    const repMap = repPeriodMap.get(repEmail)!;
    sortedPeriods.forEach((pk, pi) => {
      const cell = row.getCell(pi + 2);
      const bucket = repMap.get(pk);
      if (bucket && bucket.count > 0) {
        const avg = Math.round(bucket.totalPct / bucket.count);
        cell.value = avg;
        cell.numFmt = '0"%"';
        cell.font = { bold: true, size: 10, color: { argb: pctColor(avg) } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(avg) } };
      } else {
        cell.value = '—';
        cell.font = { size: 9, color: { argb: 'AAAAAA' } };
      }
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    row.height = 18;
  });

  s2.getColumn(1).width = 26;
  sortedPeriods.forEach((_, i) => {
    s2.getColumn(i + 2).width = period === 'weekly' ? 18 : 14;
  });

  // ── Sheet 3: By Store ─────────────────────────────────────────────────────
  const s3 = wb.addWorksheet('By Store');
  s3.views = [{ state: 'frozen', xSplit: 3, ySplit: 1, showGridLines: false }];

  // Header row
  const s3Hdr = s3.getRow(1);
  ['Store Name', 'Channel', 'Province'].forEach((h, i) => {
    const cell = s3Hdr.getCell(i + 1);
    cell.value = h;
    cell.font = HEADER_FONT;
    cell.fill = RED_FILL;
    cell.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
  });
  sortedPeriods.forEach((pk, i) => {
    const cell = s3Hdr.getCell(i + 4);
    cell.value = periodLabel(pk, period);
    cell.font = HEADER_FONT;
    cell.fill = RED_FILL;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  s3Hdr.height = period === 'weekly' ? 36 : 20;

  // Data rows
  sortedStores.forEach((storeKey, si) => {
    const [storeName] = storeKey.split('||');
    const meta = storeMeta.get(storeKey) ?? { channel: '', province: '' };
    const row = s3.getRow(si + 2);
    row.getCell(1).value = storeName;
    row.getCell(2).value = meta.channel;
    row.getCell(3).value = meta.province;
    [1, 2, 3].forEach((c) => {
      row.getCell(c).font = { size: 10 };
      row.getCell(c).fill = si % 2 === 0
        ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAFAFA' } }
        : { type: 'pattern', pattern: 'solid', fgColor: { argb: WHITE } };
    });

    const storeMap = storePeriodMap.get(storeKey)!;
    sortedPeriods.forEach((pk, pi) => {
      const cell = row.getCell(pi + 4);
      const bucket = storeMap.get(pk);
      if (bucket && bucket.count > 0) {
        const avg = Math.round(bucket.totalPct / bucket.count);
        cell.value = avg;
        cell.numFmt = '0"%"';
        cell.font = { bold: true, size: 10, color: { argb: pctColor(avg) } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(avg) } };
      } else {
        cell.value = '—';
        cell.font = { size: 9, color: { argb: 'AAAAAA' } };
      }
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    row.height = 18;
  });

  s3.getColumn(1).width = 30;
  s3.getColumn(2).width = 18;
  s3.getColumn(3).width = 16;
  sortedPeriods.forEach((_, i) => {
    s3.getColumn(i + 4).width = period === 'weekly' ? 18 : 14;
  });

  // ── Sheet 4: By Question ─────────────────────────────────────────────────
  const s4 = wb.addWorksheet('By Question');
  const NUM_FIXED_Q = 6; // Section | Question | Rep | Channel | Province | Store
  const TOTAL_COLS_S4 = NUM_FIXED_Q + sortedPeriods.length + 1;
  s4.views = [{ state: 'frozen', xSplit: NUM_FIXED_Q, ySplit: 1, showGridLines: false }];

  // Header row
  const s4Hdr = s4.getRow(1);
  ['Section', 'Question', 'Rep', 'Channel', 'Province', 'Store'].forEach((h, i) => {
    const cell = s4Hdr.getCell(i + 1);
    cell.value = h;
    cell.font = HEADER_FONT;
    cell.fill = RED_FILL;
    cell.alignment = { horizontal: i < 2 ? 'left' : 'center', vertical: 'middle' };
  });
  sortedPeriods.forEach((pk, i) => {
    const cell = s4Hdr.getCell(NUM_FIXED_Q + i + 1);
    cell.value = periodLabel(pk, period);
    cell.font = HEADER_FONT;
    cell.fill = RED_FILL;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  const s4OvHdr = s4Hdr.getCell(TOTAL_COLS_S4);
  s4OvHdr.value = 'Overall';
  s4OvHdr.font = HEADER_FONT;
  s4OvHdr.fill = RED_FILL;
  s4OvHdr.alignment = { horizontal: 'center', vertical: 'middle' };
  s4Hdr.height = period === 'weekly' ? 36 : 20;

  // AutoFilter on all columns (provides Rep/Channel/Province dropdown slicers)
  s4.autoFilter = `A1:${numToColLetter(TOTAL_COLS_S4)}1`;

  // Data rows: iterate QUESTIONS in order → for each, all (rep, channel, province) combos sorted
  let s4RowIdx = 2;
  for (const q of QUESTIONS) {
    const keysForQ = Array.from(qDimMeta.keys())
      .filter((k) => k.startsWith(`${q.id}||`))
      .sort((a, b) => {
        const ma = qDimMeta.get(a)!;
        const mb = qDimMeta.get(b)!;
        return ma.repName.localeCompare(mb.repName)
          || ma.channel.localeCompare(mb.channel)
          || ma.province.localeCompare(mb.province)
          || ma.store.localeCompare(mb.store);
      });

    for (const qKey of keysForQ) {
      const meta = qDimMeta.get(qKey)!;
      const rowBg = s4RowIdx % 2 === 0 ? 'FFFAFAFA' : WHITE;
      const bgFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: rowBg } };
      const row = s4.getRow(s4RowIdx);

      row.getCell(1).value = meta.section;
      row.getCell(2).value = meta.qText;
      row.getCell(3).value = meta.repName;
      row.getCell(4).value = meta.channel;
      row.getCell(5).value = meta.province;
      row.getCell(6).value = meta.store;
      [1, 2, 3, 4, 5, 6].forEach((c) => {
        row.getCell(c).font = { bold: c === 1, size: 9, color: { argb: DARK } };
        row.getCell(c).fill = bgFill;
        row.getCell(c).alignment = { horizontal: c <= 2 ? 'left' : 'center', vertical: 'middle', wrapText: c === 2 };
      });

      const qPMap = qDimPeriod.get(qKey) ?? new Map<string, { score: number; outOf: number }>();
      sortedPeriods.forEach((pk, pi) => {
        const cell = row.getCell(NUM_FIXED_Q + pi + 1);
        const bucket = qPMap.get(pk);
        if (bucket && bucket.outOf > 0) {
          const avg = Math.round((bucket.score / bucket.outOf) * 100);
          cell.value = avg;
          cell.numFmt = '0"%"';
          cell.font = { bold: true, size: 10, color: { argb: pctColor(avg) } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(avg) } };
        } else {
          cell.value = '—';
          cell.font = { size: 9, color: { argb: 'AAAAAA' } };
          cell.fill = bgFill;
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      const ov = meta.overall;
      const ovCell = row.getCell(TOTAL_COLS_S4);
      if (ov.outOf > 0) {
        const avg = Math.round((ov.score / ov.outOf) * 100);
        ovCell.value = avg;
        ovCell.numFmt = '0"%"';
        ovCell.font = { bold: true, size: 10, color: { argb: pctColor(avg) } };
        ovCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: pctBg(avg) } };
      } else {
        ovCell.value = '—';
        ovCell.font = { size: 9, color: { argb: 'AAAAAA' } };
        ovCell.fill = bgFill;
      }
      ovCell.alignment = { horizontal: 'center', vertical: 'middle' };
      row.height = 18;
      s4RowIdx++;
    }
  }

  // Column widths for By Question sheet
  s4.getColumn(1).width = 30; // Section
  s4.getColumn(2).width = 48; // Question
  s4.getColumn(3).width = 22; // Rep
  s4.getColumn(4).width = 20; // Channel
  s4.getColumn(5).width = 16; // Province
  s4.getColumn(6).width = 26; // Store
  sortedPeriods.forEach((_, i) => {
    s4.getColumn(NUM_FIXED_Q + i + 1).width = period === 'weekly' ? 18 : 14;
  });
  s4.getColumn(TOTAL_COLS_S4).width = 13; // Overall

  // ── Write and return ──────────────────────────────────────────────────────
  const raw = await wb.xlsx.writeBuffer();
  return Buffer.from(raw as ArrayBuffer);
}
