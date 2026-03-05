import { NextRequest, NextResponse } from 'next/server';
import { parsePerigeeExport } from '@/lib/excel-parser';
import { buildRepReport, buildFileName, buildFolderName } from '@/lib/excel-builder';
import { uploadReport } from '@/lib/sharepoint';
import { sendReports } from '@/lib/mailer';
import type { GenerateResult } from '@/types';

export const maxDuration = 60; // Vercel function max

function todayDisplay(): string {
  const d = new Date();
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file provided' }, { status: 400 });

    const selectedRepsRaw    = formData.get('selectedReps') as string | null;
    const selectedChannelsRaw = formData.get('selectedChannels') as string | null;
    const action              = (formData.get('action') as string) ?? 'sharepoint';
    const emailAddress        = (formData.get('emailAddress') as string) ?? '';

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

    // Group by rep
    const repEmails = [...new Set(filteredVisits.map((v) => v.repEmail))];

    const results: GenerateResult[] = [];
    const emailAttachments: { name: string; contentBase64: string }[] = [];
    const emailRepNames: string[] = [];

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
        results.push({
          repName: repInfo.fullName,
          fileName,
          status:  'error',
          error:   `Excel build failed: ${err instanceof Error ? err.message : String(err)}`,
        });
        continue;
      }

      const result: GenerateResult = { repName: repInfo.fullName, fileName, status: 'ok' };

      // SharePoint upload — failure is reported but does NOT block email
      if (action === 'sharepoint' || action === 'both') {
        try {
          const uploaded = await uploadReport(excelBuffer, folderName, fileName);
          result.spUrl = uploaded.webUrl;
        } catch (err: unknown) {
          result.spError = err instanceof Error ? err.message : String(err);
        }
      }

      // Collect for email — always runs regardless of SP outcome
      if (action === 'email' || action === 'both') {
        emailAttachments.push({ name: fileName, contentBase64: excelBuffer.toString('base64') });
        emailRepNames.push(repInfo.fullName);
      }

      results.push(result);
    }

    // Send email (single email with all attachments)
    if ((action === 'email' || action === 'both') && emailAttachments.length && emailAddress) {
      const dateRange = `${parsed.dateRange.from} – ${parsed.dateRange.to}`;
      try {
        await sendReports(emailAddress, emailRepNames, dateRange, emailAttachments);
        results.forEach((r) => { if (r.status === 'ok') r.emailSent = true; });
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        results.forEach((r) => {
          if (r.status === 'ok') r.error = `Email failed: ${msg}`;
        });
      }
    }

    return NextResponse.json({ results, folderName, submissionDate });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
