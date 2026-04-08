/**
 * Send email with attachments via Microsoft Graph sendMail
 * Sends from OJ_EMAIL_FROM mailbox to the user-entered address
 */

import { readFileSync } from 'fs';
import { join } from 'path';

const TENANT_ID    = process.env.OJ_TENANT_ID!;
const CLIENT_ID    = process.env.OJ_CLIENT_ID!;
const CLIENT_SECRET = process.env.OJ_CLIENT_SECRET!;
const EMAIL_FROM   = process.env.OJ_EMAIL_FROM!;

async function getToken(): Promise<string> {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         'https://graph.microsoft.com/.default',
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error(`Mail auth failed: ${data.error_description}`);
  return data.access_token as string;
}

function loadLogoB64(filename: string): string | null {
  try {
    return readFileSync(join(process.cwd(), 'public', filename)).toString('base64');
  } catch {
    return null;
  }
}

export interface Attachment {
  name: string;
  contentBase64: string;
}

export interface RepStat {
  name: string;
  storeCount: number;
}

export interface ChannelStat {
  channel: string;
  storeCount: number;
}

export interface RepEmailParams {
  toEmail: string;
  /** Optional second TO recipient — used for the rep's alias / new email. */
  aliasEmail?: string;
  ccEmail?: string;
  repFirstName: string;
  dateRange: string;    // "01 Mar 2026 – 31 Mar 2026"
  avgPct: number;       // 0–100 integer
  visitCount: number;
  storeCount: number;
  attachment: Attachment;
}

function repScoreHexColor(pct: number): string {
  if (pct >= 80) return '#1B5E20';
  if (pct >= 60) return '#4CAF50';
  if (pct >= 40) return '#F57F17';
  return '#B71C1C';
}

export async function sendRepEmail(params: RepEmailParams): Promise<void> {
  const { toEmail, aliasEmail, ccEmail, repFirstName, dateRange, avgPct, visitCount, storeCount, attachment } = params;
  const token = await getToken();
  const subject = `Your Vital Score Card Report — ${dateRange}`;
  const vitalLogoB64   = loadLogoB64('vital-logo.png');
  const perigeeLogoB64 = loadLogoB64('perigee-logo.jpg');
  const scoreColor = repScoreHexColor(avgPct);

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#32373C;">
      <table width="600" cellpadding="0" cellspacing="0" style="margin:0 auto;border:1px solid #e5e5e5;">
        <tr>
          <td style="background:#ffffff;padding:16px 30px;border-bottom:3px solid #DA291C;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align:middle;">
                  ${vitalLogoB64 ? `<img src="cid:vital-logo" width="72" height="28" style="height:28px;width:auto;display:block;" alt="Vital" />` : '<strong style="color:#DA291C;font-size:18px;">VITAL</strong>'}
                </td>
                <td style="text-align:right;vertical-align:middle;">
                  ${perigeeLogoB64 ? `<img src="cid:perigee-logo" width="28" height="28" style="height:28px;width:28px;display:block;margin-left:auto;" alt="Perigee" />` : ''}
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="background:#DA291C;padding:16px 30px;">
            <h2 style="color:#fff;margin:0;font-size:20px;">Vital Score Card Report</h2>
          </td>
        </tr>
        <tr>
          <td style="padding:24px 30px;">
            <p style="margin:0 0 16px 0;">Dear <strong>${repFirstName}</strong>,</p>
            <p style="margin:0 0 16px 0;">
              Attached is your Score Card report for <strong>${dateRange}</strong>.
            </p>
            <p style="margin:0 0 8px 0;">Your average score for this period is:</p>
            <div style="margin:12px 0 20px 0;padding:12px 24px;background:#F5F5F5;border-left:4px solid ${scoreColor};border-radius:4px;display:inline-block;">
              <span style="font-size:30px;font-weight:bold;color:${scoreColor};">${avgPct}%</span>
            </div>
            <p style="margin:0 0 8px 0;">In this date range you:</p>
            <ul style="margin:0 0 24px 0;padding-left:20px;">
              <li style="margin-bottom:4px;">Completed <strong>${visitCount}</strong> visit${visitCount !== 1 ? 's' : ''}</li>
              <li style="margin-bottom:4px;">Visited <strong>${storeCount}</strong> unique store${storeCount !== 1 ? 's' : ''}</li>
            </ul>
            <p style="margin:0 0 2px 0;font-weight:bold;">Thank you</p>
            <p style="margin:0 0 2px 0;">Vital Management</p>
            <p style="margin:0 0 24px 0;font-size:11px;color:#888;">Powered by Perigee</p>
            <p style="margin-top:16px;padding-top:12px;border-top:1px solid #eee;color:#888;font-size:11px;">
              Generated by the Vital Score Card Builder.&nbsp;
              For support, contact <a href="mailto:info@outerjoin.co.za" style="color:#DA291C;">info@outerjoin.co.za</a>
            </p>
          </td>
        </tr>
        <tr>
          <td style="background:#F2F2F2;padding:12px 30px;text-align:center;font-size:11px;color:#888;">
            OuterJoin – Visualise Efficiency
          </td>
        </tr>
      </table>
    </div>
  `;

  const inlineAttachments: object[] = [];
  if (vitalLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'vital-logo.png', contentId: 'vital-logo', isInline: true,
      contentType: 'image/png', contentBytes: vitalLogoB64,
    });
  }
  if (perigeeLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'perigee-logo.jpg', contentId: 'perigee-logo', isInline: true,
      contentType: 'image/jpeg', contentBytes: perigeeLogoB64,
    });
  }

  // Build TO recipients — include the alias as an additional address if set
  // and if it's not the same as the data email (case-insensitive).
  const toRecipients: { emailAddress: { address: string } }[] = [
    { emailAddress: { address: toEmail } },
  ];
  if (aliasEmail && aliasEmail.trim().toLowerCase() !== toEmail.trim().toLowerCase()) {
    toRecipients.push({ emailAddress: { address: aliasEmail.trim() } });
  }

  const message: Record<string, unknown> = {
    subject,
    body: { contentType: 'HTML', content: htmlBody },
    toRecipients,
    attachments: [
      ...inlineAttachments,
      {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: attachment.name,
        contentBytes: attachment.contentBase64,
      },
    ],
  };
  if (ccEmail) message.ccRecipients = [{ emailAddress: { address: ccEmail } }];

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(EMAIL_FROM)}/sendMail`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, saveToSentItems: false }),
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`sendRepEmail failed (${res.status}): ${text}`);
  }
}

export interface ComparisonEmailParams {
  toEmail: string;
  period: string;         // 'Weekly' | 'Monthly' | 'Quarterly' (display)
  dateRange: string;      // e.g. "01 Jan 2026 – 08 Mar 2026"
  visitCount: number;
  periodCount: number;
  attachment: Attachment;
}

export async function sendComparisonReport(params: ComparisonEmailParams): Promise<void> {
  const { toEmail, period, dateRange, visitCount, periodCount, attachment } = params;
  const token = await getToken();
  const subject = `Vital Score Card — ${period} Comparison Report · ${dateRange}`;

  const vitalLogoB64   = loadLogoB64('vital-logo.png');
  const perigeeLogoB64 = loadLogoB64('perigee-logo.jpg');

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#32373C;">
      <table width="600" cellpadding="0" cellspacing="0" style="margin:0 auto;border:1px solid #e5e5e5;">
        <tr>
          <td style="background:#ffffff;padding:16px 30px;border-bottom:3px solid #DA291C;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align:middle;">
                  ${vitalLogoB64 ? `<img src="cid:vital-logo" width="72" height="28" style="height:28px;width:auto;display:block;" alt="Vital" />` : '<strong style="color:#DA291C;font-size:18px;">VITAL</strong>'}
                </td>
                <td style="text-align:right;vertical-align:middle;">
                  ${perigeeLogoB64 ? `<img src="cid:perigee-logo" width="28" height="28" style="height:28px;width:28px;display:block;margin-left:auto;" alt="Perigee" />` : ''}
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="background:#DA291C;padding:16px 30px;">
            <h2 style="color:#fff;margin:0;font-size:20px;">Vital Score Card — ${period} Comparison Report</h2>
          </td>
        </tr>
        <tr>
          <td style="padding:24px 30px;">
            <p style="margin:0 0 8px 0;">Please find attached the <strong>${period.toLowerCase()} comparison report</strong> for the following period:</p>
            <table style="margin:12px 0 20px 0;border-collapse:collapse;">
              <tr>
                <td style="padding:4px 12px 4px 0;font-weight:bold;color:#32373C;">Date Range</td>
                <td style="padding:4px 0;color:#555;">${dateRange}</td>
              </tr>
              <tr>
                <td style="padding:4px 12px 4px 0;font-weight:bold;color:#32373C;">Period Type</td>
                <td style="padding:4px 0;color:#555;">${period}</td>
              </tr>
              <tr>
                <td style="padding:4px 12px 4px 0;font-weight:bold;color:#32373C;">Periods Covered</td>
                <td style="padding:4px 0;color:#555;">${periodCount}</td>
              </tr>
              <tr>
                <td style="padding:4px 12px 4px 0;font-weight:bold;color:#32373C;">Total Visits</td>
                <td style="padding:4px 0;color:#555;">${visitCount}</td>
              </tr>
            </table>
            <p style="margin:0 0 8px 0;">The report contains three sheets:</p>
            <ul style="margin:0 0 20px 0;padding-left:20px;">
              <li style="margin-bottom:4px;"><strong>Summary</strong> — Channel &amp; province breakdown with overall averages</li>
              <li style="margin-bottom:4px;"><strong>By Rep</strong> — Average score per rep per ${period.toLowerCase()} period</li>
              <li style="margin-bottom:4px;"><strong>By Store</strong> — Average score per store per ${period.toLowerCase()} period</li>
            </ul>
            <p style="margin-top:24px;color:#888;font-size:12px;">
              Generated by the Vital Score Card Builder.<br/>
              For support, contact <a href="mailto:info@outerjoin.co.za" style="color:#DA291C;">info@outerjoin.co.za</a>
            </p>
          </td>
        </tr>
        <tr>
          <td style="background:#F2F2F2;padding:12px 30px;text-align:center;font-size:11px;color:#888;">
            OuterJoin – Visualise Efficiency
          </td>
        </tr>
      </table>
    </div>
  `;

  const inlineAttachments: object[] = [];
  if (vitalLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'vital-logo.png', contentId: 'vital-logo', isInline: true,
      contentType: 'image/png', contentBytes: vitalLogoB64,
    });
  }
  if (perigeeLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'perigee-logo.jpg', contentId: 'perigee-logo', isInline: true,
      contentType: 'image/jpeg', contentBytes: perigeeLogoB64,
    });
  }

  const message = {
    subject,
    body: { contentType: 'HTML', content: htmlBody },
    toRecipients: [{ emailAddress: { address: toEmail } }],
    attachments: [
      ...inlineAttachments,
      {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: attachment.name,
        contentBytes: attachment.contentBase64,
      },
    ],
  };

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(EMAIL_FROM)}/sendMail`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, saveToSentItems: false }),
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`sendMail failed (${res.status}): ${text}`);
  }
}

export interface ExceptionEmailParams {
  toEmail: string;
  dateRange: string;
  exceptionCount: number;
  sheetSummary: { name: string; count: number }[];
  attachment: Attachment;
}

export async function sendExceptionReport(params: ExceptionEmailParams): Promise<void> {
  const { toEmail, dateRange, exceptionCount, sheetSummary, attachment } = params;
  const token = await getToken();
  const subject = `Vital Exception Report – ${dateRange}`;

  const vitalLogoB64   = loadLogoB64('vital-logo.png');
  const perigeeLogoB64 = loadLogoB64('perigee-logo.jpg');

  const summaryRows = sheetSummary
    .map((s) => `<tr><td style="padding:3px 12px 3px 0;color:#555;">${s.name}</td><td style="padding:3px 0;font-weight:bold;color:#DA291C;">${s.count}</td></tr>`)
    .join('');

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#32373C;">
      <table width="600" cellpadding="0" cellspacing="0" style="margin:0 auto;border:1px solid #e5e5e5;">
        <tr>
          <td style="background:#ffffff;padding:16px 30px;border-bottom:3px solid #DA291C;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align:middle;">
                  ${vitalLogoB64 ? `<img src="cid:vital-logo" width="72" height="28" style="height:28px;width:auto;display:block;" alt="Vital" />` : '<strong style="color:#DA291C;font-size:18px;">VITAL</strong>'}
                </td>
                <td style="text-align:right;vertical-align:middle;">
                  ${perigeeLogoB64 ? `<img src="cid:perigee-logo" width="28" height="28" style="height:28px;width:28px;display:block;margin-left:auto;" alt="Perigee" />` : ''}
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="background:#DA291C;padding:16px 30px;">
            <h2 style="color:#fff;margin:0;font-size:20px;">Vital Exception Report</h2>
          </td>
        </tr>
        <tr>
          <td style="padding:24px 30px;">
            <p style="margin:0 0 8px 0;">Please find attached the <strong>Exception Report</strong> for <strong>${dateRange}</strong>.</p>
            <div style="margin:12px 0 20px 0;padding:12px 24px;background:#F5F5F5;border-left:4px solid #DA291C;border-radius:4px;display:inline-block;">
              <span style="font-size:24px;font-weight:bold;color:#DA291C;">${exceptionCount}</span>
              <span style="font-size:14px;color:#555;margin-left:8px;">total exception${exceptionCount !== 1 ? 's' : ''}</span>
            </div>
            <p style="margin:0 0 8px 0;"><strong>Breakdown by category:</strong></p>
            <table style="margin:0 0 20px 0;border-collapse:collapse;">
              ${summaryRows}
            </table>
            <p style="margin:0 0 8px 0;font-size:12px;color:#888;">
              Each category has its own sheet in the attached Excel file with full store details, SKUs, and photo links.
            </p>
            <p style="margin-top:24px;color:#888;font-size:12px;">
              Generated by the Vital Score Card Builder.<br/>
              For support, contact <a href="mailto:info@outerjoin.co.za" style="color:#DA291C;">info@outerjoin.co.za</a>
            </p>
          </td>
        </tr>
        <tr>
          <td style="background:#F2F2F2;padding:12px 30px;text-align:center;font-size:11px;color:#888;">
            OuterJoin – Visualise Efficiency
          </td>
        </tr>
      </table>
    </div>
  `;

  const inlineAttachments: object[] = [];
  if (vitalLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'vital-logo.png', contentId: 'vital-logo', isInline: true,
      contentType: 'image/png', contentBytes: vitalLogoB64,
    });
  }
  if (perigeeLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'perigee-logo.jpg', contentId: 'perigee-logo', isInline: true,
      contentType: 'image/jpeg', contentBytes: perigeeLogoB64,
    });
  }

  const message = {
    subject,
    body: { contentType: 'HTML', content: htmlBody },
    toRecipients: [{ emailAddress: { address: toEmail } }],
    attachments: [
      ...inlineAttachments,
      {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: attachment.name,
        contentBytes: attachment.contentBase64,
      },
    ],
  };

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(EMAIL_FROM)}/sendMail`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, saveToSentItems: false }),
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`sendExceptionReport failed (${res.status}): ${text}`);
  }
}

export async function sendReports(
  toEmail: string,
  repStats: RepStat[],
  channelStats: ChannelStat[],
  dateRange: string,
  attachments: Attachment[]
): Promise<void> {
  const token = await getToken();

  const subject = `Vital Score Card Reports – ${dateRange}`;

  const vitalLogoB64   = loadLogoB64('vital-logo.png');
  const perigeeLogoB64 = loadLogoB64('perigee-logo.jpg');

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;color:#32373C;">
      <table width="600" cellpadding="0" cellspacing="0" style="margin:0 auto;border:1px solid #e5e5e5;">
        <!-- Logo row -->
        <tr>
          <td style="background:#ffffff;padding:16px 30px;border-bottom:3px solid #DA291C;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align:middle;">
                  ${vitalLogoB64 ? `<img src="cid:vital-logo" width="72" height="28" style="height:28px;width:auto;display:block;" alt="Vital" />` : '<strong style="color:#DA291C;font-size:18px;">VITAL</strong>'}
                </td>
                <td style="text-align:right;vertical-align:middle;">
                  ${perigeeLogoB64 ? `<img src="cid:perigee-logo" width="28" height="28" style="height:28px;width:28px;display:block;margin-left:auto;" alt="Perigee" />` : '<strong style="color:#888;font-size:14px;">PERIGEE</strong>'}
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <!-- Red header -->
        <tr>
          <td style="background:#DA291C;padding:16px 30px;">
            <h2 style="color:#fff;margin:0;font-size:20px;">Vital Score Card Reports</h2>
          </td>
        </tr>
        <!-- Body -->
        <tr>
          <td style="padding:24px 30px;">
            <p style="margin:0 0 8px 0;">Please find attached the score card report(s) for the following rep(s):</p>
            <ul style="margin:0 0 20px 0;padding-left:20px;">
              ${repStats.map((r) => `<li style="margin-bottom:4px;"><strong>${r.name}</strong> &mdash; ${r.storeCount} store${r.storeCount !== 1 ? 's' : ''}</li>`).join('')}
            </ul>
            <p style="margin:0 0 8px 0;"><strong>Date range:</strong> ${dateRange}</p>
            <p style="margin:20px 0 8px 0;"><strong>Channels Covered:</strong></p>
            <ul style="margin:0 0 20px 0;padding-left:20px;">
              ${channelStats.map((c) => `<li style="margin-bottom:4px;">${c.channel} &mdash; ${c.storeCount} store${c.storeCount !== 1 ? 's' : ''}</li>`).join('')}
            </ul>
            <p style="margin-top:24px;color:#888;font-size:12px;">
              Generated by the Vital Score Card Builder.<br/>
              For support, contact <a href="mailto:info@outerjoin.co.za" style="color:#DA291C;">info@outerjoin.co.za</a>
            </p>
          </td>
        </tr>
        <!-- Footer -->
        <tr>
          <td style="background:#F2F2F2;padding:12px 30px;text-align:center;font-size:11px;color:#888;">
            OuterJoin – Visualise Efficiency
          </td>
        </tr>
      </table>
    </div>
  `;

  // Build inline logo attachments
  const inlineAttachments: object[] = [];
  if (vitalLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'vital-logo.png',
      contentId: 'vital-logo',
      isInline: true,
      contentType: 'image/png',
      contentBytes: vitalLogoB64,
    });
  }
  if (perigeeLogoB64) {
    inlineAttachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'perigee-logo.jpg',
      contentId: 'perigee-logo',
      isInline: true,
      contentType: 'image/jpeg',
      contentBytes: perigeeLogoB64,
    });
  }

  const message = {
    subject,
    body: { contentType: 'HTML', content: htmlBody },
    toRecipients: [{ emailAddress: { address: toEmail } }],
    attachments: [
      ...inlineAttachments,
      ...attachments.map((a) => ({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: a.name,
        contentBytes: a.contentBase64,
      })),
    ],
  };

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(EMAIL_FROM)}/sendMail`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ message, saveToSentItems: false }),
    }
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`sendMail failed (${res.status}): ${text}`);
  }
}
