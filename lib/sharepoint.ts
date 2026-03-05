/**
 * OuterJoin SharePoint – Vital Score Cards
 * Library: Clients  /  VITAL/PERIGEE - FIELD GOOSE/2. EXTERNAL SYNC/SCORE CARDS
 */

const TENANT_ID    = process.env.OJ_TENANT_ID!;
const CLIENT_ID    = process.env.OJ_CLIENT_ID!;
const CLIENT_SECRET = process.env.OJ_CLIENT_SECRET!;
const SP_HOST      = process.env.OJ_SP_HOST      ?? 'exceler8xl.sharepoint.com';
const LIBRARY_NAME = process.env.OJ_SP_LIBRARY   ?? 'Clients';
const BASE_PATH    = process.env.OJ_SP_SCORE_CARDS_PATH
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
      body: JSON.stringify({ name: seg, folder: {}, '@microsoft.graph.conflictBehavior': 'replace' }),
    });
    current = current ? `${current}/${seg}` : seg;
  }
}

// ── Upload ────────────────────────────────────────────────────────────────────

export interface UploadResult {
  webUrl: string;
  fileId: string;
}

/**
 * Uploads a report to:
 *   {LIBRARY}/{BASE_PATH}/{submissionFolder}/{fileName}
 */
export async function uploadReport(
  buffer: Buffer,
  submissionFolder: string,
  fileName: string
): Promise<UploadResult> {
  const { token, driveId } = await getDriveContext();

  const folderPath = `${BASE_PATH}/${submissionFolder}`;
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
