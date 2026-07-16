/**
 * OuterJoin SharePoint – Vital Score Cards
 * Library: Clients
 *
 * Folder structure under BASE_PATH:
 *   Score Card Reports/{visitYear}/{submissionFolder}/{fileName}
 *     visitYear = year of the latest visit date in the batch
 *   Historical Data/VITAL_SCORE_HISTORY.xlsx
 *     Single cumulative file — no year split, so cross-year comparisons work
 *   Comparison Reports/{fromYear}/{period}/{fileName}
 *     fromYear = year of the comparison FROM date
 */

const TENANT_ID     = process.env.OJ_TENANT_ID!;
const CLIENT_ID     = process.env.OJ_CLIENT_ID!;
const CLIENT_SECRET = process.env.OJ_CLIENT_SECRET!;
const SP_HOST       = process.env.OJ_SP_HOST      ?? 'exceler8xl.sharepoint.com';
const LIBRARY_NAME  = process.env.OJ_SP_LIBRARY   ?? 'Clients';
const BASE_PATH     = process.env.OJ_SP_SCORE_CARDS_PATH
  ?? 'VITAL/PERIGEE - FIELD GOOSE/2. EXTERNAL SYNC/SCORE CARDS';

// ── Auth ──────────────────────────────────────────────────────────────────────

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
  if (!data.access_token) {
    throw new Error(`SP auth failed: ${data.error_description ?? JSON.stringify(data)}`);
  }
  return data.access_token as string;
}

function encodePath(path: string): string {
  return path.split('/').map((seg) => encodeURIComponent(seg)).join('/');
}

type DriveCtx = { token: string; driveId: string };

async function getDriveContext(): Promise<DriveCtx> {
  const token = await getToken();

  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SP_HOST}:/`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!siteRes.ok) throw new Error(`SP: could not get site: ${await siteRes.text()}`);
  const site = await siteRes.json();

  const drivesRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const drives = await drivesRes.json();
  const drive = drives.value?.find((d: { name: string }) => d.name === LIBRARY_NAME);
  if (!drive) {
    const names = drives.value?.map((d: { name: string }) => d.name).join(', ');
    throw new Error(`SP: library "${LIBRARY_NAME}" not found. Available: ${names}`);
  }
  return { token, driveId: drive.id as string };
}

// ── Folder creation ───────────────────────────────────────────────────────────

async function ensureFolderExists(
  token: string,
  driveId: string,
  folderPath: string
): Promise<void> {
  const segments = folderPath.split('/');
  let current = '';
  for (const seg of segments) {
    const parent = current ? encodePath(current) : undefined;
    const url = parent
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${parent}:/children`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;

    await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: seg, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
    });
    // We ignore the response — 409 Conflict just means the folder already exists, which is fine.
    current = current ? `${current}/${seg}` : seg;
  }
}

// ── History + Log file helpers ────────────────────────────────────────────────
// Single cumulative files — no year folder. Cross-year comparisons read one file.

const HISTORY_FOLDER    = `${BASE_PATH}/Historical Data`;
const HISTORY_FILENAME  = 'VITAL_SCORE_HISTORY.xlsx';
const LOG_FILENAME      = 'VITAL_ACTIVITY_LOG.json';
const ALIASES_FILENAME  = 'VITAL_EMAIL_ALIASES.json';
const PHOTOMAP_FILENAME = 'VITAL_PHOTO_QUESTION_MAP.json';

export async function downloadLogFile(): Promise<Buffer | null> {
  const { token, driveId } = await getDriveContext();
  const filePath = encodePath(`${HISTORY_FOLDER}/${LOG_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Log download failed (${res.status}): ${await res.text()}`);
  return Buffer.from(await res.arrayBuffer());
}

export async function uploadLogFile(buffer: Buffer): Promise<void> {
  const { token, driveId } = await getDriveContext();
  await ensureFolderExists(token, driveId, HISTORY_FOLDER);
  const filePath = encodePath(`${HISTORY_FOLDER}/${LOG_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: new Uint8Array(buffer),
    }
  );
  if (!res.ok) throw new Error(`Log upload failed (${res.status}): ${await res.text()}`);
}

// ── Email aliases file (JSON) ─────────────────────────────────────────────────

export async function downloadAliasesFile(): Promise<Buffer | null> {
  const { token, driveId } = await getDriveContext();
  const filePath = encodePath(`${HISTORY_FOLDER}/${ALIASES_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Aliases download failed (${res.status}): ${await res.text()}`);
  return Buffer.from(await res.arrayBuffer());
}

export async function uploadAliasesFile(buffer: Buffer): Promise<void> {
  const { token, driveId } = await getDriveContext();
  await ensureFolderExists(token, driveId, HISTORY_FOLDER);
  const filePath = encodePath(`${HISTORY_FOLDER}/${ALIASES_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: new Uint8Array(buffer),
    }
  );
  if (!res.ok) throw new Error(`Aliases upload failed (${res.status}): ${await res.text()}`);
}

// ── Photo-question → sheet mapping file (JSON) ────────────────────────────────

export async function downloadPhotoMapFile(): Promise<Buffer | null> {
  const { token, driveId } = await getDriveContext();
  const filePath = encodePath(`${HISTORY_FOLDER}/${PHOTOMAP_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Photo map download failed (${res.status}): ${await res.text()}`);
  return Buffer.from(await res.arrayBuffer());
}

export async function uploadPhotoMapFile(buffer: Buffer): Promise<void> {
  const { token, driveId } = await getDriveContext();
  await ensureFolderExists(token, driveId, HISTORY_FOLDER);
  const filePath = encodePath(`${HISTORY_FOLDER}/${PHOTOMAP_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: new Uint8Array(buffer),
    }
  );
  if (!res.ok) throw new Error(`Photo map upload failed (${res.status}): ${await res.text()}`);
}

/**
 * Downloads the cumulative history Excel.
 * Returns null ONLY if the file doesn't exist yet (404).
 * Throws on any other error so callers can surface the real problem.
 */
export async function downloadHistoryFile(): Promise<Buffer | null> {
  const { token, driveId } = await getDriveContext();
  const filePath = encodePath(`${HISTORY_FOLDER}/${HISTORY_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`History download failed (${res.status}): ${await res.text()}`);
  return Buffer.from(await res.arrayBuffer());
}

/**
 * Uploads (overwrites) the cumulative history Excel.
 */
export async function uploadHistoryFile(buffer: Uint8Array): Promise<void> {
  const { token, driveId } = await getDriveContext();
  await ensureFolderExists(token, driveId, HISTORY_FOLDER);
  const filePath = encodePath(`${HISTORY_FOLDER}/${HISTORY_FILENAME}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: new Uint8Array(buffer),
    }
  );
  if (!res.ok) throw new Error(`History upload failed (${res.status}): ${await res.text()}`);
}

// ── Comparison report upload ──────────────────────────────────────────────────

export interface UploadResult {
  webUrl: string;
  fileId: string;
}

/**
 * Uploads a comparison report to:
 *   Comparison Reports/{fromYear}/{period}/{fileName}
 *
 * @param year  Year derived from the comparison FROM date (e.g. "2026")
 */
export async function uploadComparisonReport(
  buffer: Buffer,
  year: string,
  period: string,
  fileName: string
): Promise<UploadResult> {
  const { token, driveId } = await getDriveContext();
  const folderPath = `${BASE_PATH}/Comparison Reports/${year}/${period}`;
  await ensureFolderExists(token, driveId, folderPath);

  const filePath = encodePath(`${folderPath}/${fileName}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: new Uint8Array(buffer),
    }
  );
  if (!res.ok) throw new Error(`Comparison upload failed (${res.status}): ${await res.text()}`);
  const uploaded = await res.json();
  return { webUrl: uploaded.webUrl as string, fileId: uploaded.id as string };
}

// ── Exception report upload ──────────────────────────────────────────────────

const EXCEPTION_PATH = process.env.OJ_SP_EXCEPTION_PATH
  ?? 'VITAL/PERIGEE - FIELD GOOSE/2. EXTERNAL SYNC/EXCEPTION REPORTS';

/**
 * Uploads an exception report to:
 *   EXCEPTION REPORTS/{visitYear}/{submissionFolder}/{fileName}
 */
export async function uploadExceptionReport(
  buffer: Buffer,
  year: string,
  submissionFolder: string,
  fileName: string
): Promise<UploadResult> {
  const { token, driveId } = await getDriveContext();
  const folderPath = `${EXCEPTION_PATH}/${year}/${submissionFolder}`;
  await ensureFolderExists(token, driveId, folderPath);

  const filePath = encodePath(`${folderPath}/${fileName}`);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: new Uint8Array(buffer),
    }
  );
  if (!res.ok) throw new Error(`Exception report upload failed (${res.status}): ${await res.text()}`);
  const uploaded = await res.json();
  return { webUrl: uploaded.webUrl as string, fileId: uploaded.id as string };
}

// ── Score card report upload ──────────────────────────────────────────────────

/**
 * Uploads a score card report to:
 *   Score Card Reports/{visitYear}/{submissionFolder}/{fileName}
 *
 * @param year  Year derived from the latest visit date in the batch (e.g. "2026")
 */
export async function uploadReport(
  buffer: Buffer,
  year: string,
  submissionFolder: string,
  fileName: string
): Promise<UploadResult> {
  const { token, driveId } = await getDriveContext();

  const folderPath = `${BASE_PATH}/Score Card Reports/${year}/${submissionFolder}`;
  await ensureFolderExists(token, driveId, folderPath);

  const filePath  = encodePath(`${folderPath}/${fileName}`);
  const uploadRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      body: new Uint8Array(buffer),
    }
  );
  if (!uploadRes.ok) {
    throw new Error(`SP upload failed (${uploadRes.status}): ${await uploadRes.text()}`);
  }
  const uploaded = await uploadRes.json();
  return { webUrl: uploaded.webUrl as string, fileId: uploaded.id as string };
}
