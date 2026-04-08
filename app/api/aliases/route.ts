import { NextRequest, NextResponse } from 'next/server';
import { loadAliases, upsertAlias, removeAlias } from '@/lib/aliases';

// Mutable JSON state — never cache
export const dynamic = 'force-dynamic';

const NO_STORE = { 'Cache-Control': 'no-store' };

export async function GET() {
  try {
    const file = await loadAliases();
    return NextResponse.json({ entries: file.entries }, { headers: NO_STORE });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500, headers: NO_STORE });
  }
}

interface UpsertBody {
  dataEmail?: string;
  aliasEmail?: string;
  displayName?: string;
  // Bulk variant
  entries?: { dataEmail: string; aliasEmail?: string; displayName?: string }[];
}

export async function POST(req: NextRequest) {
  try {
    const body = (await req.json()) as UpsertBody;

    if (Array.isArray(body.entries)) {
      const saved = [];
      for (const e of body.entries) {
        if (!e.dataEmail) continue;
        saved.push(
          await upsertAlias({
            dataEmail: e.dataEmail,
            aliasEmail: e.aliasEmail,
            displayName: e.displayName,
          })
        );
      }
      return NextResponse.json({ saved }, { headers: NO_STORE });
    }

    if (!body.dataEmail) {
      return NextResponse.json(
        { error: 'dataEmail is required' },
        { status: 400, headers: NO_STORE }
      );
    }

    const saved = await upsertAlias({
      dataEmail: body.dataEmail,
      aliasEmail: body.aliasEmail,
      displayName: body.displayName,
    });
    return NextResponse.json({ saved }, { headers: NO_STORE });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500, headers: NO_STORE });
  }
}

export async function DELETE(req: NextRequest) {
  try {
    const url = new URL(req.url);
    const dataEmail = url.searchParams.get('dataEmail');
    if (!dataEmail) {
      return NextResponse.json(
        { error: 'dataEmail query param is required' },
        { status: 400, headers: NO_STORE }
      );
    }
    const removed = await removeAlias(dataEmail);
    return NextResponse.json({ removed }, { headers: NO_STORE });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500, headers: NO_STORE });
  }
}
