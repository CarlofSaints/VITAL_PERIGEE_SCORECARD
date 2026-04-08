import ExcelJS from 'exceljs';
import { QUESTIONS } from '@/constants/questions';
import { formatDateDisplay } from './excel-parser';
import type { StoreVisit } from '@/types';

// ── Brand colours ─────────────────────────────────────────────────────────────
const VITAL_RED  = 'FFDA291C';
const WHITE      = 'FFFFFFFF';
const DARK_GRAY  = 'FF32373C';
const LIGHT_GRAY = 'FFF2F2F2';

function solidFill(argb: string): ExcelJS.Fill {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

// Short, human-friendly sheet names per question (max 31 chars for Excel)
const SHEET_NAMES: Record<string, string> = {
  'osa-sku':          'Ranged SKU Report',
  'osa-flow':         'Merchandising Flow',
  'osa-oos':          'OOS Flagged',
  'osa-npd':          'NPD Report',
  'osa-backup':       'OOS Backup Stock',
  'promo-active':     'Promotion Active',
  'promo-comms':      'Promo Communication',
  'price-pi':         'PI Labels',
  'price-correct':    'Pricing Correct',
  'price-files':      'Pricing Files Shared',
  'price-tracking':   'Pricing Tracking',
  'sh-clean':         'Clean & Presentable',
  'sh-fifo':          'FIFO Rule',
  'sh-shortdated':    'Short Dated Stock',
  'sh-expired':       'Damaged & Expired',
  'sh-expired-check': 'Expired Stock',
  'mp-available':     'Multi Placements',
  'mp-forward':       'Forward Share',
  'mp-identified':    'Multi Place Identified',
  'pos-implemented':  'POS Implemented',
  'gen-mgr-issues':   'Store Manager Issues',
  'gen-opportunities': 'Opportunities',
};

interface ExceptionRow {
  repName: string;
  channel: string;
  store: string;
  storeCode: string;
  province: string;
  date: string;
  answer: string;
  skus: string;
  photoUrls: string[];
  comment: string;
}

interface OrphanPhotoRow {
  repName: string;
  channel: string;
  store: string;
  storeCode: string;
  province: string;
  date: string;
  photoQuestion: string;
  url: string;
}

interface QuestionExceptions {
  questionId: string;
  sheetName: string;
  questionText: string;
  section: string;
  rows: ExceptionRow[];
}

// Name of the catch-all sheet for photo columns that don't map to a scorecard question.
const ORPHAN_SHEET_NAME = 'Photos with no question';

// ── Main builder ──────────────────────────────────────────────────────────────

export async function buildExceptionReport(
  visits: StoreVisit[],
  submissionDate: string
): Promise<{ buffer: Buffer; exceptionCount: number; sheetSummary: { name: string; count: number }[] }> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Vital Exception Report Builder';
  wb.created = new Date();

  const hr = (argb: string) => solidFill(argb);

  // Gather exceptions per question
  const allExceptions: QuestionExceptions[] = [];

  // Gather orphan photo rows across all visits (photos with no question)
  const orphanPhotoRows: OrphanPhotoRow[] = [];
  for (const v of visits) {
    if (!v.orphanPhotos?.length) continue;
    for (const op of v.orphanPhotos) {
      orphanPhotoRows.push({
        repName: `${v.repFirstName} ${v.repLastName}`.trim(),
        channel: v.channel,
        store: v.store,
        storeCode: v.storeCode,
        province: v.province,
        date: v.date,
        photoQuestion: op.header,
        url: op.url,
      });
    }
  }

  for (const q of QUESTIONS) {
    const badAnswer = q.inverted ? 'yes' : 'no';

    const rows: ExceptionRow[] = [];
    for (const v of visits) {
      const answer = v.answers[q.id];
      if (!answer || answer.toLowerCase() !== badAnswer) continue;
      rows.push({
        repName: `${v.repFirstName} ${v.repLastName}`.trim(),
        channel: v.channel,
        store: v.store,
        storeCode: v.storeCode,
        province: v.province,
        date: v.date,
        answer: answer,
        skus: v.skus?.[q.id] ?? '',
        photoUrls: v.photoUrls?.[q.id] ?? [],
        comment: v.comments?.[q.id] ?? '',
      });
    }

    if (rows.length === 0) continue;

    allExceptions.push({
      questionId: q.id,
      sheetName: (SHEET_NAMES[q.id] ?? q.text).slice(0, 31),
      questionText: q.text,
      section: q.section,
      rows,
    });
  }

  const totalExceptions = allExceptions.reduce((sum, qe) => sum + qe.rows.length, 0);

  // ── Summary sheet ──────────────────────────────────────────────────────────
  const summaryWs = wb.addWorksheet('Summary');
  summaryWs.views = [{ showGridLines: false }];

  summaryWs.getColumn(1).width = 42;
  summaryWs.getColumn(2).width = 34;
  summaryWs.getColumn(3).width = 14;

  // Title
  summaryWs.getRow(1).height = 32;
  summaryWs.mergeCells('A1:C1');
  const title = summaryWs.getCell('A1');
  title.value = `VITAL EXCEPTION REPORT — ${submissionDate}`;
  title.font = { bold: true, size: 14, color: { argb: WHITE } };
  title.fill = hr(VITAL_RED);
  title.alignment = { horizontal: 'center', vertical: 'middle' };

  // Stats row
  summaryWs.getRow(2).height = 22;
  summaryWs.mergeCells('A2:C2');
  const stats = summaryWs.getCell('A2');
  const orphanStats = orphanPhotoRows.length > 0
    ? ` · ${orphanPhotoRows.length} orphan photo${orphanPhotoRows.length !== 1 ? 's' : ''}`
    : '';
  stats.value = `${totalExceptions} total exception${totalExceptions !== 1 ? 's' : ''} across ${allExceptions.length} question${allExceptions.length !== 1 ? 's' : ''} · ${visits.length} visits analysed${orphanStats}`;
  stats.font = { size: 10, italic: true, color: { argb: DARK_GRAY } };
  stats.fill = hr(LIGHT_GRAY);
  stats.alignment = { horizontal: 'center', vertical: 'middle' };

  // Headers
  summaryWs.getRow(4).height = 18;
  for (const [ci, lbl] of ['SECTION', 'QUESTION', 'EXCEPTIONS'].entries()) {
    const c = summaryWs.getCell(4, ci + 1);
    c.value = lbl;
    c.font = { bold: true, size: 9, color: { argb: WHITE } };
    c.fill = hr(DARK_GRAY);
    c.alignment = { horizontal: ci === 2 ? 'center' : 'left', vertical: 'middle' };
  }

  // Summary rows
  allExceptions.forEach((qe, i) => {
    const rowIdx = i + 5;
    const rowBg = i % 2 === 0 ? LIGHT_GRAY : WHITE;
    summaryWs.getRow(rowIdx).height = 18;

    const sectionCell = summaryWs.getCell(rowIdx, 1);
    sectionCell.value = qe.section;
    sectionCell.font = { size: 9, color: { argb: DARK_GRAY } };
    sectionCell.fill = hr(rowBg);
    sectionCell.alignment = { vertical: 'middle' };

    const qCell = summaryWs.getCell(rowIdx, 2);
    qCell.value = { text: qe.sheetName, hyperlink: `#'${qe.sheetName}'!A1` };
    qCell.font = { size: 9, underline: true, color: { argb: 'FF0563C1' } };
    qCell.fill = hr(rowBg);
    qCell.alignment = { vertical: 'middle' };

    const cntCell = summaryWs.getCell(rowIdx, 3);
    cntCell.value = qe.rows.length;
    cntCell.font = { bold: true, size: 10, color: { argb: VITAL_RED } };
    cntCell.fill = hr(rowBg);
    cntCell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Orphan photos summary row (if any)
  if (orphanPhotoRows.length > 0) {
    const rowIdx = allExceptions.length + 5;
    const rowBg = allExceptions.length % 2 === 0 ? LIGHT_GRAY : WHITE;
    summaryWs.getRow(rowIdx).height = 18;

    const sectionCell = summaryWs.getCell(rowIdx, 1);
    sectionCell.value = 'Orphan Photos';
    sectionCell.font = { size: 9, color: { argb: DARK_GRAY } };
    sectionCell.fill = hr(rowBg);
    sectionCell.alignment = { vertical: 'middle' };

    const qCell = summaryWs.getCell(rowIdx, 2);
    qCell.value = { text: ORPHAN_SHEET_NAME, hyperlink: `#'${ORPHAN_SHEET_NAME}'!A1` };
    qCell.font = { size: 9, underline: true, color: { argb: 'FF0563C1' } };
    qCell.fill = hr(rowBg);
    qCell.alignment = { vertical: 'middle' };

    const cntCell = summaryWs.getCell(rowIdx, 3);
    cntCell.value = orphanPhotoRows.length;
    cntCell.font = { bold: true, size: 10, color: { argb: VITAL_RED } };
    cntCell.fill = hr(rowBg);
    cntCell.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  if (allExceptions.length === 0 && orphanPhotoRows.length === 0) {
    summaryWs.getRow(5).height = 20;
    summaryWs.mergeCells('A5:C5');
    const noData = summaryWs.getCell('A5');
    noData.value = 'No exceptions found — all responses are positive.';
    noData.font = { size: 10, italic: true, color: { argb: DARK_GRAY } };
    noData.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // ── Per-question sheets ────────────────────────────────────────────────────
  // Helper: 1-based column index → Excel letters (A, B, … Z, AA, AB …)
  const colLetter = (n: number): string => {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };

  for (const qe of allExceptions) {
    const ws = wb.addWorksheet(qe.sheetName);
    ws.views = [{ state: 'frozen', ySplit: 2, showGridLines: true }];

    // Dynamic photo column count — one column per photo up to this question's max.
    const maxPhotoCount = qe.rows.reduce((m, r) => Math.max(m, r.photoUrls.length), 0);
    const firstPhotoCol = 8;                   // after SKUs
    const commentCol = firstPhotoCol + maxPhotoCount;
    const totalCols = commentCol;              // last column

    // Column widths
    ws.getColumn(1).width = 22;  // Rep Name
    ws.getColumn(2).width = 22;  // Channel
    ws.getColumn(3).width = 36;  // Store
    ws.getColumn(4).width = 12;  // Store Code
    ws.getColumn(5).width = 18;  // Province
    ws.getColumn(6).width = 14;  // Date
    ws.getColumn(7).width = 52;  // SKUs
    for (let p = 0; p < maxPhotoCount; p++) {
      ws.getColumn(firstPhotoCol + p).width = 14;
    }
    ws.getColumn(commentCol).width = 40;  // Comments

    const lastColLetter = colLetter(totalCols);

    // Row 1: Title bar with section + question name
    ws.getRow(1).height = 28;
    ws.mergeCells(`A1:${lastColLetter}1`);
    const sheetTitle = ws.getCell('A1');
    sheetTitle.value = `${qe.section}: ${qe.questionText}  (${qe.rows.length} exception${qe.rows.length !== 1 ? 's' : ''})`;
    sheetTitle.font = { bold: true, size: 11, color: { argb: WHITE } };
    sheetTitle.fill = hr(VITAL_RED);
    sheetTitle.alignment = { horizontal: 'left', vertical: 'middle' };

    // Row 2: Column headers — expands Photo into Photo 1 / Photo 2 / … when needed
    ws.getRow(2).height = 20;
    const photoHeaders: string[] =
      maxPhotoCount === 0
        ? []
        : maxPhotoCount === 1
        ? ['Photo']
        : Array.from({ length: maxPhotoCount }, (_, i) => `Photo ${i + 1}`);
    const headers = ['Rep Name', 'Channel', 'Store', 'Store Code', 'Province', 'Date', "SKU's", ...photoHeaders, 'Comments'];
    headers.forEach((label, i) => {
      const c = ws.getCell(2, i + 1);
      c.value = label;
      c.font = { bold: true, size: 9, color: { argb: WHITE } };
      c.fill = hr(DARK_GRAY);
      c.alignment = { horizontal: 'left', vertical: 'middle' };
    });

    // Data rows
    qe.rows.forEach((ex, i) => {
      const rowIdx = i + 3;
      const rowBg = i % 2 === 0 ? LIGHT_GRAY : WHITE;
      const hasMultiLineSkus = ex.skus.includes('|');
      ws.getRow(rowIdx).height = hasMultiLineSkus ? 44 : 20;

      // Standard text cells
      const vals = [ex.repName, ex.channel, ex.store, ex.storeCode, ex.province, ex.date];
      vals.forEach((val, ci) => {
        const c = ws.getCell(rowIdx, ci + 1);
        c.value = val;
        c.font = { size: 9, color: { argb: DARK_GRAY } };
        c.fill = hr(rowBg);
        c.alignment = { horizontal: 'left', vertical: 'middle' };
      });

      // SKUs — replace pipe with newline for readability
      const skuCell = ws.getCell(rowIdx, 7);
      skuCell.value = ex.skus ? ex.skus.replace(/\|/g, '\n') : '';
      skuCell.font = { size: 8, color: { argb: DARK_GRAY } };
      skuCell.fill = hr(rowBg);
      skuCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

      // Photo cells — one hyperlink per column, blank where this row has fewer photos
      const cellLabel = maxPhotoCount === 1 ? 'View Photo' : 'View';
      for (let p = 0; p < maxPhotoCount; p++) {
        const photoCell = ws.getCell(rowIdx, firstPhotoCol + p);
        const url = ex.photoUrls[p];
        if (url) {
          photoCell.value = { text: cellLabel, hyperlink: url };
          photoCell.font = { size: 8, underline: true, color: { argb: 'FF0563C1' } };
        } else {
          photoCell.value = '';
          photoCell.font = { size: 8, color: { argb: DARK_GRAY } };
        }
        photoCell.fill = hr(rowBg);
        photoCell.alignment = { horizontal: 'center', vertical: 'middle' };
      }

      // Comments
      const commentCell = ws.getCell(rowIdx, commentCol);
      commentCell.value = ex.comment;
      commentCell.font = { size: 9, italic: true, color: { argb: DARK_GRAY } };
      commentCell.fill = hr(rowBg);
      commentCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    });

    // Auto-filter on header row
    ws.autoFilter = { from: 'A2', to: `${lastColLetter}${2 + qe.rows.length}` };
  }

  // ── "Photos with no question" sheet ────────────────────────────────────────
  if (orphanPhotoRows.length > 0) {
    const ws = wb.addWorksheet(ORPHAN_SHEET_NAME);
    ws.views = [{ state: 'frozen', ySplit: 2, showGridLines: true }];

    ws.getColumn(1).width = 22;  // Rep Name
    ws.getColumn(2).width = 22;  // Channel
    ws.getColumn(3).width = 36;  // Store
    ws.getColumn(4).width = 12;  // Store Code
    ws.getColumn(5).width = 18;  // Province
    ws.getColumn(6).width = 14;  // Date
    ws.getColumn(7).width = 36;  // Photo Question
    ws.getColumn(8).width = 16;  // Link

    // Row 1: Title bar
    ws.getRow(1).height = 28;
    ws.mergeCells('A1:H1');
    const sheetTitle = ws.getCell('A1');
    sheetTitle.value = `Photos with no question  (${orphanPhotoRows.length} photo${orphanPhotoRows.length !== 1 ? 's' : ''})`;
    sheetTitle.font = { bold: true, size: 11, color: { argb: WHITE } };
    sheetTitle.fill = hr(VITAL_RED);
    sheetTitle.alignment = { horizontal: 'left', vertical: 'middle' };

    // Row 2: Column headers
    ws.getRow(2).height = 20;
    const orphanHeaders = ['Rep Name', 'Channel', 'Store', 'Store Code', 'Province', 'Date', 'Photo Question', 'Link'];
    orphanHeaders.forEach((label, i) => {
      const c = ws.getCell(2, i + 1);
      c.value = label;
      c.font = { bold: true, size: 9, color: { argb: WHITE } };
      c.fill = hr(DARK_GRAY);
      c.alignment = { horizontal: 'left', vertical: 'middle' };
    });

    // Data rows
    orphanPhotoRows.forEach((op, i) => {
      const rowIdx = i + 3;
      const rowBg = i % 2 === 0 ? LIGHT_GRAY : WHITE;
      ws.getRow(rowIdx).height = 20;

      const vals = [op.repName, op.channel, op.store, op.storeCode, op.province, op.date, op.photoQuestion];
      vals.forEach((val, ci) => {
        const c = ws.getCell(rowIdx, ci + 1);
        c.value = val;
        c.font = { size: 9, color: { argb: DARK_GRAY } };
        c.fill = hr(rowBg);
        c.alignment = { horizontal: 'left', vertical: 'middle' };
      });

      const linkCell = ws.getCell(rowIdx, 8);
      linkCell.value = { text: 'View Photo', hyperlink: op.url };
      linkCell.font = { size: 8, underline: true, color: { argb: 'FF0563C1' } };
      linkCell.fill = hr(rowBg);
      linkCell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    ws.autoFilter = { from: 'A2', to: `H${2 + orphanPhotoRows.length}` };
  }

  const buffer = await wb.xlsx.writeBuffer();
  const sheetSummary = allExceptions.map((qe) => ({ name: qe.sheetName, count: qe.rows.length }));
  if (orphanPhotoRows.length > 0) {
    sheetSummary.push({ name: ORPHAN_SHEET_NAME, count: orphanPhotoRows.length });
  }
  return {
    buffer: Buffer.from(buffer),
    exceptionCount: totalExceptions,
    sheetSummary,
  };
}

export function buildExceptionFileName(allVisits: StoreVisit[]): string {
  const dates = allVisits.map((v) => v.dateObj).filter((d) => d.getTime() > 0);
  if (!dates.length) return 'Vital Exception Report.xlsx';
  const min = new Date(Math.min(...dates.map((d) => d.getTime())));
  const max = new Date(Math.max(...dates.map((d) => d.getTime())));
  return `Vital Exception Report - ${formatDateDisplay(min)} - ${formatDateDisplay(max)}.xlsx`
    .replace(/[/\\:*?"<>|]/g, '-')
    .trim();
}

export function buildExceptionFolderName(allVisits: StoreVisit[]): string {
  const dates = allVisits.map((v) => v.dateObj).filter((d) => d.getTime() > 0);
  if (!dates.length) return 'EXCEPTION REPORT';
  const min = new Date(Math.min(...dates.map((d) => d.getTime())));
  const max = new Date(Math.max(...dates.map((d) => d.getTime())));
  return `EXCEPTION REPORT - ${formatDateDisplay(min)} - ${formatDateDisplay(max)}`;
}
