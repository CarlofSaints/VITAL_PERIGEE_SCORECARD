import ExcelJS from 'exceljs';
import { QUESTIONS } from '@/constants/questions';
import { formatDateDisplay } from './excel-parser';
import type { Rep, StoreVisit } from '@/types';

// ── Brand colours ─────────────────────────────────────────────────────────────
const VITAL_RED   = 'FFDA291C';
const WHITE       = 'FFFFFFFF';
const DARK_GRAY   = 'FF32373C';
const LIGHT_GRAY  = 'FFF2F2F2';
const SECTION_BG  = 'FFDA291C';
const HEADER_BG   = 'FF1A1A2E';
const SCORE_GREEN = 'FF00B050';
const SCORE_ORANGE = 'FFFFC000';
const SCORE_RED   = 'FFFF0000';
const CF_RED      = 'FFF8696B';
const CF_GREEN    = 'FF63BE7B';

// ── Helpers ───────────────────────────────────────────────────────────────────

function scoreColor(pct: number): string {
  if (pct > 0.8)  return SCORE_GREEN;
  if (pct >= 0.6) return SCORE_ORANGE;
  return SCORE_RED;
}

function getScore(
  answer: string | null,
  inverted: boolean
): { score: number | null; outOf: number } {
  if (!answer || answer.toLowerCase() === 'n/a') return { score: null, outOf: 0 };
  const isYes = answer.toLowerCase() === 'yes';
  const good = inverted ? !isYes : isYes;
  return { score: good ? 1 : 0, outOf: 1 };
}

function solidFill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

// Safe Excel sheet name (max 31 chars, no special chars)
function safeSheetName(name: string): string {
  return name.slice(0, 31).replace(/[:\\/?*[\]]/g, ' ').trim();
}

// Strip the first " - " separated prefix from a store name (typically the
// channel, e.g. "PICK N PAY - CORPORATE VINCENT PARK" → "CORPORATE VINCENT PARK").
// Falls back to the original if the strip would leave an empty string.
function stripChannelPrefix(storeName: string): string {
  const idx = storeName.indexOf(' - ');
  if (idx < 0) return storeName.trim();
  const stripped = storeName.slice(idx + 3).trim();
  return stripped || storeName.trim();
}

// Build a unique Excel sheet name for a store, in the form
//   "STORE NAME (CODE)"
// while keeping to Excel's 31-char limit and stripping illegal characters.
// Uses parentheses (not square brackets) because Excel disallows [ ] in sheet
// names. Falls back gracefully when the store code is missing.
function buildSheetName(storeName: string, storeCode: string): string {
  const cleanBase = stripChannelPrefix(storeName)
    .replace(/[:\\/?*[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const cleanCode = (storeCode || '').replace(/[:\\/?*[\]()]/g, '').trim();

  if (!cleanCode) return cleanBase.slice(0, 31).trim();

  const suffix = ` (${cleanCode})`;
  // Absurdly long code — fall back to just the truncated code
  if (suffix.length >= 31) return suffix.trim().slice(0, 31);

  const maxBase = 31 - suffix.length;
  const truncated = cleanBase.slice(0, maxBase).trim();
  return `${truncated}${suffix}`;
}

// Column letter: 1=A, 27=AA, …
function colLetter(n: number): string {
  let r = '';
  while (n > 0) {
    r = String.fromCharCode(65 + ((n - 1) % 26)) + r;
    n = Math.floor((n - 1) / 26);
  }
  return r;
}

function repDateRange(visits: StoreVisit[]): { from: string; to: string } {
  const dates = visits.map((v) => v.dateObj).filter((d) => d.getTime() > 0);
  if (!dates.length) return { from: '', to: '' };
  const min = new Date(Math.min(...dates.map((d) => d.getTime())));
  const max = new Date(Math.max(...dates.map((d) => d.getTime())));
  return { from: formatDateDisplay(min), to: formatDateDisplay(max) };
}

function repProvince(visits: StoreVisit[]): string {
  const freq: Record<string, number> = {};
  for (const v of visits) {
    if (v.province) freq[v.province] = (freq[v.province] ?? 0) + 1;
  }
  if (!Object.keys(freq).length) return '';
  return Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0];
}

// ── Store score calculator ────────────────────────────────────────────────────

interface StoreData {
  name: string;        // original full store name (e.g. "PICK N PAY - CORPORATE VINCENT PARK")
  sheetName: string;   // unique Excel sheet tab name (max 31 chars, includes store code)
  visit: StoreVisit;
  score: number;
  outOf: number;
  pct: number;
}

function calcStoreScores(visit: StoreVisit): { score: number; outOf: number; pct: number } {
  let score = 0, outOf = 0;
  for (const q of QUESTIONS) {
    const { score: s, outOf: o } = getScore(visit.answers[q.id] ?? null, q.inverted ?? false);
    score  += s ?? 0;
    outOf  += o;
  }
  return { score, outOf, pct: outOf > 0 ? score / outOf : 0 };
}

// ── Store Sheet ────────────────────────────────────────────────────────────────

function buildStoreSheet(
  wb: ExcelJS.Workbook,
  sheetName: string,
  sd: StoreData,
  submissionDate: string,
  repDateRng: { from: string; to: string }
) {
  const ws = wb.addWorksheet(sheetName);
  ws.views = [{ state: 'frozen', ySplit: 5, showGridLines: false }];
  const visit = sd.visit;

  ws.getColumn(1).width = 52;
  ws.getColumn(2).width = 12;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 8;
  ws.getColumn(5).width = 8;
  ws.getColumn(6).width = 38;
  ws.getColumn(7).width = 32;
  ws.getColumn(8).width = 36;

  const hr = (argb: string) => solidFill(argb);

  // Row 1: Title + HOME button
  ws.getRow(1).height = 30;
  ws.mergeCells('A1:H1');
  const t = ws.getCell('A1');
  t.value = 'RETAIL SCORECARD';
  t.font = { bold: true, size: 16, color: { argb: WHITE } };
  t.fill  = hr(VITAL_RED);
  t.alignment = { horizontal: 'center', vertical: 'middle' };

  ws.getColumn(9).width = 10;
  const homeCell = ws.getCell('I1');
  homeCell.value = { text: '← HOME', hyperlink: '#TOTAL!A1' };
  homeCell.font  = { bold: true, size: 9, color: { argb: WHITE } };
  homeCell.fill  = hr(DARK_GRAY);
  homeCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 2: DATE / DATE RANGE / COMPLETED BY
  ws.getRow(2).height = 20;
  ws.mergeCells('A2:B2');
  ws.getCell('A2').value = `DATE: ${submissionDate}`;
  ws.getCell('A2').font  = { bold: true, size: 10, color: { argb: DARK_GRAY } };
  ws.getCell('A2').fill  = hr(LIGHT_GRAY);
  ws.getCell('A2').alignment = { vertical: 'middle' };

  ws.getCell('C2').value = `${repDateRng.from} – ${repDateRng.to}`;
  ws.getCell('C2').font  = { bold: true, size: 10, color: { argb: VITAL_RED } };
  ws.getCell('C2').fill  = hr(LIGHT_GRAY);
  ws.getCell('C2').alignment = { vertical: 'middle' };

  ws.mergeCells('D2:H2');
  ws.getCell('D2').value = `COMPLETED BY: ${visit.repFirstName} ${visit.repLastName}`;
  ws.getCell('D2').font  = { bold: true, size: 10, color: { argb: DARK_GRAY } };
  ws.getCell('D2').fill  = hr(LIGHT_GRAY);
  ws.getCell('D2').alignment = { vertical: 'middle' };

  // Row 3: STORE
  ws.getRow(3).height = 20;
  ws.mergeCells('A3:B3');
  ws.getCell('A3').value = 'STORE:';
  ws.getCell('A3').font  = { bold: true, size: 10, color: { argb: DARK_GRAY } };
  ws.getCell('A3').fill  = hr(LIGHT_GRAY);
  ws.getCell('A3').alignment = { vertical: 'middle' };
  ws.mergeCells('C3:H3');
  ws.getCell('C3').value = visit.store;
  ws.getCell('C3').font  = { bold: true, size: 10, color: { argb: DARK_GRAY } };
  ws.getCell('C3').fill  = hr(LIGHT_GRAY);
  ws.getCell('C3').alignment = { vertical: 'middle' };

  // Row 4: STORE CODE / PROVINCE
  ws.getRow(4).height = 18;
  ws.mergeCells('A4:B4');
  ws.getCell('A4').value = `STORE CODE: ${visit.storeCode}`;
  ws.getCell('A4').font  = { size: 9, color: { argb: DARK_GRAY } };
  ws.getCell('A4').fill  = hr(LIGHT_GRAY);
  ws.getCell('A4').alignment = { vertical: 'middle' };
  ws.mergeCells('C4:D4');
  ws.getCell('C4').value = `PROVINCE: ${visit.province}`;
  ws.getCell('C4').font  = { size: 9, color: { argb: DARK_GRAY } };
  ws.getCell('C4').fill  = hr(LIGHT_GRAY);
  ws.getCell('C4').alignment = { vertical: 'middle' };
  ws.mergeCells('E4:H4');
  ws.getCell('E4').value = `CHANNEL: ${visit.channel}`;
  ws.getCell('E4').font  = { size: 9, color: { argb: DARK_GRAY } };
  ws.getCell('E4').fill  = hr(LIGHT_GRAY);
  ws.getCell('E4').alignment = { vertical: 'middle' };

  // Row 5: Column headers
  ws.getRow(5).height = 16;
  ws.mergeCells('A5:C5');
  const qhdr = ws.getCell('A5');
  qhdr.value = 'QUESTIONS';
  qhdr.font  = { bold: true, size: 9, color: { argb: WHITE } };
  qhdr.fill  = hr(DARK_GRAY);
  qhdr.alignment = { horizontal: 'center', vertical: 'middle' };
  for (const [col, lbl] of [['D5','MAX'],['E5','SCORE'],['F5','COMMENTS'],['G5',"SKU'S"],['H5','PIC URL']] as const) {
    ws.getCell(col).value = lbl;
    ws.getCell(col).font  = { bold: true, size: 9, color: { argb: WHITE } };
    ws.getCell(col).fill  = hr(DARK_GRAY);
    ws.getCell(col).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // ── Questions ────────────────────────────────────────────────────────────────
  let rowIdx = 6;
  let currentSection = '';

  for (const q of QUESTIONS) {
    if (q.section !== currentSection) {
      currentSection = q.section;
      ws.getRow(rowIdx).height = 18;
      ws.mergeCells(`A${rowIdx}:H${rowIdx}`);
      const s = ws.getCell(`A${rowIdx}`);
      s.value = q.section;
      s.font  = { bold: true, size: 10, color: { argb: WHITE } };
      s.fill  = hr(SECTION_BG);
      s.alignment = { horizontal: 'left', vertical: 'middle' };
      rowIdx++;
    }

    const answer  = visit.answers[q.id] ?? null;
    const { score, outOf } = getScore(answer, q.inverted ?? false);
    const comment = visit.comments[q.id] ?? '';
    const skuRaw  = visit.skus?.[q.id] ?? '';
    const photoUrl = visit.photoUrls?.[q.id] ?? '';
    const rowBg   = rowIdx % 2 === 0 ? LIGHT_GRAY : 'FFFFFFFF';

    ws.getRow(rowIdx).height = 32;
    ws.mergeCells(`A${rowIdx}:B${rowIdx}`);
    const qc = ws.getCell(`A${rowIdx}`);
    qc.value = q.text;
    qc.font  = { size: 9, color: { argb: DARK_GRAY } };
    qc.fill  = hr(rowBg);
    qc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    const ac = ws.getCell(`C${rowIdx}`);
    ac.value = answer ?? 'N/A';
    ac.font  = { size: 9, color: { argb: DARK_GRAY } };
    ac.fill  = hr(rowBg);
    ac.alignment = { horizontal: 'center', vertical: 'middle' };

    const mc = ws.getCell(`D${rowIdx}`);
    mc.value = outOf === 1 ? 1 : '-';
    mc.font  = { size: 9 };
    mc.fill  = hr(rowBg);
    mc.alignment = { horizontal: 'center', vertical: 'middle' };

    const sc = ws.getCell(`E${rowIdx}`);
    sc.value = score !== null ? score : '-';
    sc.font  = {
      size: 9, bold: true,
      color: { argb: score === null ? DARK_GRAY : score === 1 ? SCORE_GREEN : SCORE_RED },
    };
    sc.fill  = hr(rowBg);
    sc.alignment = { horizontal: 'center', vertical: 'middle' };

    const cc = ws.getCell(`F${rowIdx}`);
    cc.value = comment;
    cc.font  = { size: 9, italic: true, color: { argb: DARK_GRAY } };
    cc.fill  = hr(rowBg);
    cc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // SKU's column — replace pipe separators with newlines for readability
    const gc = ws.getCell(`G${rowIdx}`);
    gc.value = skuRaw ? skuRaw.replace(/\|/g, '\n') : '';
    gc.font  = { size: 8, color: { argb: DARK_GRAY } };
    gc.fill  = hr(rowBg);
    gc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

    // PIC URL column — clickable hyperlink if present
    const hc = ws.getCell(`H${rowIdx}`);
    if (photoUrl) {
      hc.value = { text: 'View Photo', hyperlink: photoUrl };
      hc.font  = { size: 8, underline: true, color: { argb: 'FF0563C1' } };
    } else {
      hc.value = '';
      hc.font  = { size: 8, color: { argb: DARK_GRAY } };
    }
    hc.fill  = hr(rowBg);
    hc.alignment = { horizontal: 'center', vertical: 'middle' };

    rowIdx++;
  }

  // ── Score summary ─────────────────────────────────────────────────────────────
  rowIdx++; // spacer
  const addScore = (label: string, val: number | string, pctFmt = false) => {
    ws.getRow(rowIdx).height = 20;
    ws.mergeCells(`A${rowIdx}:D${rowIdx}`);
    const l = ws.getCell(`A${rowIdx}`);
    l.value = label;
    l.font  = { bold: true, size: 10, color: { argb: WHITE } };
    l.fill  = hr(DARK_GRAY);
    l.alignment = { horizontal: 'right', vertical: 'middle' };
    const v = ws.getCell(`E${rowIdx}`);
    v.value = val;
    if (pctFmt) v.numFmt = '0.0%';
    v.font  = { bold: true, size: 10, color: { argb: pctFmt ? scoreColor(sd.pct) : DARK_GRAY } };
    v.fill  = hr(LIGHT_GRAY);
    v.alignment = { horizontal: 'center', vertical: 'middle' };
    // Fill remaining columns for clean look
    for (const col of ['F', 'G', 'H']) {
      ws.getCell(`${col}${rowIdx}`).fill = hr(DARK_GRAY);
    }
    rowIdx++;
  };

  addScore('Score:', sd.score);
  addScore('Out of:', sd.outOf);
  addScore('% Score:', sd.pct, true);

  // ── Overall comment ───────────────────────────────────────────────────────────
  if (visit.overallComment) {
    rowIdx++;
    ws.getRow(rowIdx).height = 16;
    ws.mergeCells(`A${rowIdx}:H${rowIdx}`);
    const cl = ws.getCell(`A${rowIdx}`);
    cl.value = 'GENERAL COMMENTS:';
    cl.font  = { bold: true, size: 10, color: { argb: WHITE } };
    cl.fill  = hr(DARK_GRAY);
    cl.alignment = { vertical: 'middle' };
    rowIdx++;

    ws.getRow(rowIdx).height = 60;
    ws.mergeCells(`A${rowIdx}:H${rowIdx}`);
    const cv = ws.getCell(`A${rowIdx}`);
    cv.value = visit.overallComment;
    cv.font  = { size: 9, italic: true, color: { argb: DARK_GRAY } };
    cv.fill  = hr(LIGHT_GRAY);
    cv.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
  }
}

// ── TOTAL Sheet ────────────────────────────────────────────────────────────────

function buildTotalSheet(
  ws: ExcelJS.Worksheet,
  rep: Rep,
  stores: StoreData[],
  submissionDate: string,
  repDateRng: { from: string; to: string }
) {
  const n = stores.length;
  const overallScore  = stores.reduce((s, st) => s + st.score, 0);
  const overallOutOf  = stores.reduce((s, st) => s + st.outOf, 0);
  const overallPct    = overallOutOf > 0 ? overallScore / overallOutOf : 0;

  ws.views = [{ state: 'frozen', ySplit: 2, showGridLines: false }];

  // Column widths
  ws.getColumn(1).width = 50;
  ws.getColumn(2).width = 4;
  ws.getColumn(3).width = 14;
  for (let i = 4; i <= 3 + n; i++) ws.getColumn(i).width = 7.15;

  const hr = (argb: string) => solidFill(argb);

  // ── Row 1: Header ──────────────────────────────────────────────────────────────
  ws.getRow(1).height = 28;
  ws.mergeCells('A1:B1');
  const a1 = ws.getCell('A1');
  a1.value = `DATE: ${submissionDate}`;
  a1.font  = { bold: true, size: 11, color: { argb: WHITE } };
  a1.fill  = hr(HEADER_BG);
  a1.alignment = { horizontal: 'left', vertical: 'middle' };

  const c1 = ws.getCell('C1');
  c1.value = `${repDateRng.from} – ${repDateRng.to}`;
  c1.font  = { bold: true, size: 11, color: { argb: SCORE_ORANGE } };
  c1.fill  = hr(HEADER_BG);
  c1.alignment = { horizontal: 'center', vertical: 'middle' };

  ws.getCell('D1').value = `COMPLETED BY: ${rep.fullName}`;
  ws.getCell('D1').font  = { bold: true, size: 11, color: { argb: WHITE } };
  ws.getCell('D1').fill  = hr(HEADER_BG);
  ws.getCell('D1').alignment = { horizontal: 'left', vertical: 'middle' };

  // Fill remaining header cells
  for (let i = 5; i <= 3 + n; i++) ws.getCell(1, i).fill = hr(HEADER_BG);

  // ── Row 2: Overall score + store names ────────────────────────────────────────
  ws.getRow(2).height = 150;
  ws.mergeCells('A2:C2');
  const scoreCell = ws.getCell('A2');
  scoreCell.value  = overallPct;
  scoreCell.numFmt = '0.0%';
  scoreCell.font   = { bold: true, size: 60, color: { argb: scoreColor(overallPct) } };
  scoreCell.alignment = { horizontal: 'center', vertical: 'middle' };
  scoreCell.fill  = hr(LIGHT_GRAY);

  stores.forEach((st, i) => {
    const c = ws.getCell(2, 4 + i);
    c.value = { text: st.sheetName, hyperlink: `#'${st.sheetName}'!A1` };
    c.font  = { bold: true, size: 9, color: { argb: WHITE } };
    c.fill  = hr(DARK_GRAY);
    c.alignment = { horizontal: 'center', vertical: 'bottom', textRotation: 90 };
  });

  // ── Rows 3+: Sections + questions ────────────────────────────────────────────
  let rowIdx = 3;
  let currentSection = '';

  for (const q of QUESTIONS) {
    if (q.section !== currentSection) {
      currentSection = q.section;
      ws.getRow(rowIdx).height = 18;
      ws.mergeCells(`A${rowIdx}:B${rowIdx}`);
      const s = ws.getCell(`A${rowIdx}`);
      s.value = q.section;
      s.font  = { bold: true, size: 10, color: { argb: WHITE } };
      s.fill  = hr(SECTION_BG);
      s.alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getCell(`C${rowIdx}`).fill = hr(SECTION_BG);
      for (let i = 0; i < n; i++) ws.getCell(rowIdx, 4 + i).fill = hr(SECTION_BG);
      rowIdx++;
    }

    const rowBg = rowIdx % 2 === 0 ? LIGHT_GRAY : 'FFFFFFFF';
    ws.getRow(rowIdx).height = 28;
    ws.mergeCells(`A${rowIdx}:B${rowIdx}`);
    const qc = ws.getCell(`A${rowIdx}`);
    qc.value = q.text;
    qc.font  = { size: 9, color: { argb: DARK_GRAY } };
    qc.fill  = hr(rowBg);
    qc.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    ws.getCell(`C${rowIdx}`).fill = hr(rowBg);

    stores.forEach((st, i) => {
      const answer = st.visit.answers[q.id] ?? null;
      const { score } = getScore(answer, q.inverted ?? false);
      const c = ws.getCell(rowIdx, 4 + i);
      if (score === null) {
        c.value = '-';
        c.font  = { size: 10, color: { argb: '88888888' } };
      } else {
        c.value = score;
        c.font  = { bold: true, size: 10, color: { argb: score === 1 ? SCORE_GREEN : SCORE_RED } };
      }
      c.fill  = hr(rowBg);
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    rowIdx++;
  }

  // ── Score rows ────────────────────────────────────────────────────────────────
  const addRow = (label: string, vals: (number | string)[], numFmt?: string) => {
    ws.getRow(rowIdx).height = 20;
    ws.mergeCells(`A${rowIdx}:B${rowIdx}`);
    const l = ws.getCell(`A${rowIdx}`);
    l.value = label;
    l.font  = { bold: true, size: 10, color: { argb: WHITE } };
    l.fill  = hr(DARK_GRAY);
    l.alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell(`C${rowIdx}`).fill = hr(DARK_GRAY);
    vals.forEach((v, i) => {
      const c = ws.getCell(rowIdx, 4 + i);
      c.value = v;
      if (numFmt) c.numFmt = numFmt;
      c.font  = { bold: true, size: 10 };
      c.fill  = hr(LIGHT_GRAY);
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    rowIdx++;
  };

  rowIdx++;
  addRow('Score:',   stores.map((s) => s.score));
  addRow('Out of:',  stores.map((s) => s.outOf));

  const pctRowIdx = rowIdx;
  addRow('% Score:', stores.map((s) => s.pct), '0.0%');

  // Conditional colour scale on % row
  if (n > 0) {
    const firstCol = colLetter(4);
    const lastCol  = colLetter(3 + n);
    ws.addConditionalFormatting({
      ref: `${firstCol}${pctRowIdx}:${lastCol}${pctRowIdx}`,
      rules: [
        {
          type: 'colorScale',
          cfvo: [{ type: 'min' }, { type: 'max' }],
          color: [{ argb: CF_RED }, { argb: CF_GREEN }],
          priority: 1,
        },
      ],
    });
  }
}

// ── Main export ───────────────────────────────────────────────────────────────

export async function buildRepReport(
  rep: Rep,
  visits: StoreVisit[],
  submissionDate: string
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vital Score Card Builder';
  wb.created = new Date();

  const repDateRng = repDateRange(visits);

  // Deduplicate: latest visit per store+code composite key. Using the composite
  // key is important because two different stores can share the same name but
  // have different store codes (e.g. "PICK N PAY - CORPORATE VINCENT PARK" in
  // EC vs PC province). Keying on name alone would collapse them.
  const storeMap = new Map<string, StoreVisit>();
  for (const v of visits) {
    const key = `${v.store}||${v.storeCode || ''}`;
    const existing = storeMap.get(key);
    if (!existing || v.dateObj > existing.dateObj) storeMap.set(key, v);
  }

  // Build StoreData with resolved sheet names, then sort alphabetically.
  const stores: StoreData[] = Array.from(storeMap.values())
    .map((visit) => {
      const { score, outOf, pct } = calcStoreScores(visit);
      return {
        name: visit.store,
        sheetName: buildSheetName(visit.store, visit.storeCode),
        visit,
        score,
        outOf,
        pct,
      };
    })
    .sort((a, b) => a.sheetName.localeCompare(b.sheetName));

  // Guarantee uniqueness of sheet names. In the normal case each entry already
  // has a unique "NAME (CODE)" suffix, but if two entries end up identical after
  // truncation we append a numeric discriminator to keep exceljs happy.
  const usedSheetNames = new Set<string>();
  for (const sd of stores) {
    let sn = sd.sheetName;
    if (usedSheetNames.has(sn)) {
      let i = 2;
      // Reserve 4 chars for "-NN" suffix
      const trimmed = sn.slice(0, 28).trim();
      let candidate = `${trimmed}-${i}`;
      while (usedSheetNames.has(candidate)) {
        i++;
        candidate = `${trimmed}-${i}`;
      }
      sn = candidate;
    }
    sd.sheetName = sn;
    usedSheetNames.add(sn);
  }

  // Add TOTAL sheet first (so it's tab 1)
  const totalWs = wb.addWorksheet('TOTAL');
  buildTotalSheet(totalWs, rep, stores, submissionDate, repDateRng);

  // Add individual store sheets using the pre-resolved unique sheet names
  for (const sd of stores) {
    buildStoreSheet(wb, sd.sheetName, sd, submissionDate, repDateRng);
  }

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export function buildFileName(rep: Rep, visits: StoreVisit[]): string {
  const province = repProvince(visits);
  const { from, to } = repDateRange(visits);
  const name = `${province ? province + ' ' : ''}${rep.firstName} ${rep.lastName} ${from} - ${to}_SCORE CARD.xlsx`;
  return name.replace(/[/\\:*?"<>|]/g, '-').trim();
}

export function buildFolderName(allVisits: StoreVisit[]): string {
  const dates = allVisits.map((v) => v.dateObj).filter((d) => d.getTime() > 0);
  if (!dates.length) return 'SCORE CARD';
  const min = new Date(Math.min(...dates.map((d) => d.getTime())));
  const max = new Date(Math.max(...dates.map((d) => d.getTime())));
  return `SCORE CARD - ${formatDateDisplay(min)} - ${formatDateDisplay(max)}`;
}
