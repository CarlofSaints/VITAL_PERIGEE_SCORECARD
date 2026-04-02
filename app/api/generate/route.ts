import { NextRequest, NextResponse } from 'next/server';
import { parsePerigeeExport } from '@/lib/excel-parser';
import { buildRepReport, buildFileName, buildFolderName } from '@/lib/excel-builder';
import { buildExceptionReport, buildExceptionFileName, buildExceptionFolderName } from '@/lib/exception-builder';
import { uploadReport, uploadExceptionReport } from '@/lib/sharepoint';
import { sendReports, sendRepEmail, sendExceptionReport, type RepStat, type ChannelStat } from '@/lib/mailer';
import { appendToHistory } from '@/lib/history';
import { appendActivityLog, type LogReport } from '@/lib/activityLog';
import { QUESTIONS } from '@/constants/questions';
import type { GenerateResult, StoreVisit } from '@/types';

export const maxDuration = 60; // Vercel function max

function todayDisplay(): string {
  const d = new Date();
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

/** Compute N/A-aware average score for a rep's deduped visits */
function calcRepScoreInfo(repVisits: StoreVisit[]): { avgPct: number; visitCount: number; storeCount: number } {
  // Deduplicate by store — latest visit wins (mirrors buildRepReport)
  const storeMap = new Map<string, StoreVisit>();
  for (const v of repVisits) {
    const e = storeMap.get(v.store);
    if (!e || v.dateObj > e.dateObj) storeMap.set(v.store, v);
  }
  const deduped = Array.from(storeMap.values());
  let score = 0, outOf = 0;
  for (const v of deduped) {
    for (const q of QUESTIONS) {
      const ans = v.answers[q.id] ?? null;
      if (!ans || ans.toLowerCase() === 'n/a') continue;
      const isYes = ans.toLowerCase() === 'yes';
      score += (q.inverted ? !isYes : isYes) ? 1 : 0;
      outOf++;
    }
  }
  return {
    avgPct: outOf > 0 ? Math.round((score / outOf) * 100) : 0,
    visitCount: repVisits.length,
    storeCount: deduped.length,
  };
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file provided' }, { status: 400 });

    const selectedRepsRaw        = formData.get('selectedReps') as string | null;
    const selectedChannelsRaw    = formData.get('selectedChannels') as string | null;
    const action                 = (formData.get('action') as string) ?? 'sharepoint';
    const emailAddress           = (formData.get('emailAddress') as string) ?? '';
    const emailToReps            = formData.get('emailToReps') === 'true';
    const ccEmail                = (formData.get('ccEmail') as string) ?? '';
    const includeExceptionReport = formData.get('includeExceptionReport') === 'true';

    const selectedReps     = selectedRepsRaw    ? JSON.parse(selectedRepsRaw)    : null;
    const selectedChannels = selectedChannelsRaw ? JSON.parse(selectedChannelsRaw) : null;

    const buffer = Buffer.from(await file.arrayBuffer());
    const parsed = parsePerigeeExport(buffer);

    // Filter visits
    const filteredVisits = parsed.visits.filter((v) => {
      const repMatch     = !selectedReps    || selectedReps.includes(v.repEmail);
      const channelMatch = !selectedChannels || selectedChannels.includes(v.channel);
      return repMatch && channelMatch;
    });

    if (!filteredVisits.length) {
      return NextResponse.json({ error: 'No visits match the selected filters.' }, { status: 400 });
    }

    const submissionDate = todayDisplay();
    const folderName     = buildFolderName(filteredVisits);

    // Derive year from the latest visit date in the batch (not today's date)
    const maxVisitDate = filteredVisits.reduce(
      (max, v) => v.dateObj > max ? v.dateObj : max,
      filteredVisits[0].dateObj
    );
    const visitYear = maxVisitDate.getFullYear().toString();

    // Group by rep
    const repEmails = [...new Set(filteredVisits.map((v) => v.repEmail))];

    // Compute channel stats (unique stores per channel across all filtered visits)
    const channelStoreMap = new Map<string, Set<string>>();
    for (const v of filteredVisits) {
      if (!channelStoreMap.has(v.channel)) channelStoreMap.set(v.channel, new Set());
      channelStoreMap.get(v.channel)!.add(v.store);
    }
    const channelStats: ChannelStat[] = Array.from(channelStoreMap.entries())
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([channel, stores]) => ({ channel, storeCount: stores.size }));

    const results: GenerateResult[] = [];
    const emailAttachments: { name: string; contentBase64: string }[] = [];
    const emailRepStats: RepStat[] = [];
    const logReports: LogReport[] = [];

    const dateRange = `${parsed.dateRange.from} – ${parsed.dateRange.to}`;

    for (const repEmail of repEmails) {
      const repVisits = filteredVisits.filter((v) => v.repEmail === repEmail);
      const repInfo   = parsed.reps.find((r) => r.email === repEmail);
      if (!repInfo || !repVisits.length) continue;

      const fileName = buildFileName(repInfo, repVisits);

      // Build Excel — if this fails, skip this rep entirely
      let excelBuffer: Buffer;
      try {
        excelBuffer = await buildRepReport(repInfo, repVisits, submissionDate);
      } catch (err: unknown) {
        const errMsg = `Excel build failed: ${err instanceof Error ? err.message : String(err)}`;
        results.push({ repName: repInfo.fullName, fileName, status: 'error', error: errMsg });
        logReports.push({ repName: repInfo.fullName, fileName, spSuccess: false, success: false, error: errMsg });
        continue;
      }

      const result: GenerateResult = { repName: repInfo.fullName, fileName, status: 'ok' };
      const logReport: LogReport = { repName: repInfo.fullName, fileName, spSuccess: false, success: true };

      // SharePoint upload — failure is reported but does NOT block email
      if (action === 'sharepoint' || action === 'both') {
        try {
          const uploaded = await uploadReport(excelBuffer, visitYear, folderName, fileName);
          result.spUrl = uploaded.webUrl;
          logReport.spSuccess = true;
          logReport.spUrl = uploaded.webUrl;
        } catch (err: unknown) {
          result.spError = err instanceof Error ? err.message : String(err);
          logReport.spError = result.spError;
        }
      }

      // Collect for summary email
      if (action === 'email' || action === 'both') {
        emailAttachments.push({ name: fileName, contentBase64: excelBuffer.toString('base64') });
        const storeCount = new Set(repVisits.map((v) => v.store)).size;
        emailRepStats.push({ name: repInfo.fullName, storeCount });
      }

      // Send individual rep email
      if (emailToReps && repInfo.email) {
        const { avgPct, visitCount, storeCount } = calcRepScoreInfo(repVisits);
        try {
          await sendRepEmail({
            toEmail: repInfo.email,
            ccEmail: ccEmail || undefined,
            repFirstName: repInfo.firstName,
            dateRange,
            avgPct,
            visitCount,
            storeCount,
            attachment: { name: fileName, contentBase64: excelBuffer.toString('base64') },
          });
          result.repEmailSent = true;
          logReport.repEmailSent = true;
          logReport.repEmailTo = repInfo.email;
          if (ccEmail) logReport.repEmailCc = ccEmail;
        } catch (err: unknown) {
          result.repEmailError = err instanceof Error ? err.message : String(err);
          logReport.repEmailSent = false;
          logReport.repEmailError = result.repEmailError;
        }
      }

      results.push(result);
      logReports.push(logReport);
    }

    // Send summary email (single email with all attachments)
    type SummaryEmailLog = { sent: boolean; to: string; attachmentCount: number; error?: string };
    let summaryEmailLog: SummaryEmailLog | undefined;
    if ((action === 'email' || action === 'both') && emailAttachments.length && emailAddress) {
      try {
        await sendReports(emailAddress, emailRepStats, channelStats, dateRange, emailAttachments);
        results.forEach((r) => { if (r.status === 'ok') r.emailSent = true; });
        summaryEmailLog = { sent: true, to: emailAddress, attachmentCount: emailAttachments.length };
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        results.forEach((r) => { if (r.status === 'ok') r.error = `Email failed: ${msg}`; });
        summaryEmailLog = { sent: false, to: emailAddress, attachmentCount: emailAttachments.length, error: msg };
      }
    }

    // ── Exception report (optional) ──────────────────────────────────────────
    let exceptionResult: {
      fileName: string;
      exceptionCount: number;
      sheetSummary: { name: string; count: number }[];
      spUrl?: string;
      spError?: string;
      emailSent?: boolean;
      emailError?: string;
    } | undefined;

    if (includeExceptionReport) {
      try {
        const { buffer: exBuf, exceptionCount, sheetSummary } = await buildExceptionReport(filteredVisits, submissionDate);
        const exFileName = buildExceptionFileName(filteredVisits);
        const exFolderName = buildExceptionFolderName(filteredVisits);

        exceptionResult = { fileName: exFileName, exceptionCount, sheetSummary };

        // Upload to SharePoint
        if (action === 'sharepoint' || action === 'both') {
          try {
            const uploaded = await uploadExceptionReport(exBuf, visitYear, exFolderName, exFileName);
            exceptionResult.spUrl = uploaded.webUrl;
          } catch (err: unknown) {
            exceptionResult.spError = err instanceof Error ? err.message : String(err);
          }
        }

        // Send separate email
        if ((action === 'email' || action === 'both') && emailAddress) {
          try {
            await sendExceptionReport({
              toEmail: emailAddress,
              dateRange,
              exceptionCount,
              sheetSummary,
              attachment: { name: exFileName, contentBase64: exBuf.toString('base64') },
            });
            exceptionResult.emailSent = true;
          } catch (err: unknown) {
            exceptionResult.emailError = err instanceof Error ? err.message : String(err);
          }
        }
      } catch (err: unknown) {
        exceptionResult = {
          fileName: 'Exception Report',
          exceptionCount: 0,
          sheetSummary: [],
          spError: `Build failed: ${err instanceof Error ? err.message : String(err)}`,
        };
      }
    }

    // Append visits to the persistent history file on SharePoint
    let historyAdded = 0;
    let historyError: string | undefined;
    try {
      const histResult = await appendToHistory(filteredVisits, submissionDate);
      historyAdded = histResult.added;
    } catch (err) {
      historyError = err instanceof Error ? err.message : String(err);
      console.error('History append failed:', historyError);
    }

    // Log the activity (non-fatal)
    try {
      await appendActivityLog({
        type: 'generate',
        dateRange,
        action,
        reports: logReports,
        summaryEmail: summaryEmailLog ?? undefined,
        historyAdded,
        historyError,
        overallSuccess: results.every((r) => r.status === 'ok'),
      });
    } catch (err) {
      console.error('Activity log append failed:', err instanceof Error ? err.message : err);
    }

    return NextResponse.json({ results, folderName, submissionDate, historyAdded, historyError, exceptionReport: exceptionResult });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
