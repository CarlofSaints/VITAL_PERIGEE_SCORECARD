import { NextRequest, NextResponse } from 'next/server';
import { downloadHistoryFile, uploadComparisonReport } from '@/lib/sharepoint';
import { buildComparisonReport, extractFilterOptions, type Period, type ComparisonFilters } from '@/lib/comparison-builder';
import { sendComparisonReport } from '@/lib/mailer';
import { appendActivityLog } from '@/lib/activityLog';

export const maxDuration = 60;

// ── GET: return unique dimension values from history (populates UI filter dropdowns) ──

export async function GET() {
  try {
    let historyBuffer: Buffer | null;
    try {
      historyBuffer = await downloadHistoryFile();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      return NextResponse.json({ error: `Could not read history data: ${msg}` }, { status: 500 });
    }
    if (!historyBuffer) {
      return NextResponse.json({ channels: [], provinces: [], reps: [], stores: [] });
    }
    const options = await extractFilterOptions(historyBuffer);
    return NextResponse.json(options);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}

function formatDateDisplay(d: Date): string {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

function parseInputDate(s: string): Date | null {
  if (!s) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [y, m, d] = s.split('-').map(Number);
    const dt = new Date(y, m - 1, d);
    return isNaN(dt.getTime()) ? null : dt;
  }
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    const [d, m, y] = s.split('/').map(Number);
    const dt = new Date(y, m - 1, d);
    return isNaN(dt.getTime()) ? null : dt;
  }
  return null;
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      from: fromStr,
      to: toStr,
      period: periodRaw,
      action = 'sharepoint',
      emailAddress,
      filters,
    } = body as {
      from: string;
      to: string;
      period: string;
      action?: string;
      emailAddress?: string;
      filters?: ComparisonFilters;
    };

    if (!fromStr || !toStr || !periodRaw) {
      return NextResponse.json({ error: 'from, to, and period are required.' }, { status: 400 });
    }

    const period = periodRaw.toLowerCase() as Period;
    if (!['weekly', 'monthly', 'quarterly'].includes(period)) {
      return NextResponse.json({ error: 'period must be weekly, monthly, or quarterly.' }, { status: 400 });
    }

    const from = parseInputDate(fromStr);
    const to   = parseInputDate(toStr);
    if (!from || !to) {
      return NextResponse.json({ error: 'Invalid date format. Use YYYY-MM-DD.' }, { status: 400 });
    }
    if (from > to) {
      return NextResponse.json({ error: 'From date must be before To date.' }, { status: 400 });
    }

    // Download history file — throws if it exists but can't be read
    let historyBuffer: Buffer | null;
    try {
      console.log('[comparison] Downloading history file…');
      historyBuffer = await downloadHistoryFile();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      return NextResponse.json({ error: `Could not read history data: ${msg}` }, { status: 500 });
    }

    if (!historyBuffer) {
      return NextResponse.json({
        error: 'No history data found yet. Generate at least one batch of score cards first — history is built automatically on each generation.',
      }, { status: 400 });
    }

    console.log(`[comparison] History file downloaded: ${(historyBuffer.length / 1024).toFixed(0)} KB`);

    const dateRangeLabel = `${formatDateDisplay(from)} – ${formatDateDisplay(to)}`;
    const fileName = `VITAL COMPARISON REPORT ${formatDateDisplay(from)} - ${formatDateDisplay(to)}.xlsx`;
    const periodDisplay = period.charAt(0).toUpperCase() + period.slice(1);
    const fromYear = from.getFullYear().toString();

    // Build Excel
    console.log('[comparison] Building comparison report…');
    const excelBuffer = await buildComparisonReport({
      historyBuffer,
      from,
      to,
      period,
      dateRangeLabel,
      filters,
    });

    console.log(`[comparison] Report built: ${(excelBuffer.length / 1024).toFixed(0)} KB`);

    // SharePoint upload
    let spUrl: string | undefined;
    let spError: string | undefined;
    if (action === 'sharepoint' || action === 'both') {
      try {
        const uploaded = await uploadComparisonReport(excelBuffer, fromYear, period, fileName);
        spUrl = uploaded.webUrl;
      } catch (err: unknown) {
        spError = err instanceof Error ? err.message : String(err);
      }
    }

    // Send email
    let emailSent = false;
    let emailError: string | undefined;
    if ((action === 'email' || action === 'both') && emailAddress && emailAddress.includes('@')) {
      try {
        await sendComparisonReport({
          toEmail: emailAddress,
          period: periodDisplay,
          dateRange: dateRangeLabel,
          visitCount: 0,
          periodCount: 0,
          attachment: {
            name: fileName,
            contentBase64: excelBuffer.toString('base64'),
          },
        });
        emailSent = true;
      } catch (err: unknown) {
        emailError = err instanceof Error ? err.message : String(err);
      }
    }

    // Log the activity (non-fatal)
    try {
      await appendActivityLog({
        type: 'comparison',
        dateRange: dateRangeLabel,
        action,
        period: periodDisplay,
        compFileName: fileName,
        compSpUrl: spUrl,
        compSpError: spError,
        compEmailSent: emailSent,
        compEmailTo: emailAddress,
        compEmailError: emailError,
        overallSuccess: !spError && (!emailSent || !emailError),
      });
    } catch (err) {
      console.error('Activity log append failed:', err instanceof Error ? err.message : err);
    }

    return NextResponse.json({
      ok: true,
      fileName,
      spUrl,
      spError,
      emailSent,
      emailError,
      dateRange: dateRangeLabel,
    });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
