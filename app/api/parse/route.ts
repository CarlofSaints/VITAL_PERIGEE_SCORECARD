import { NextRequest, NextResponse } from 'next/server';
import { parsePerigeeExport } from '@/lib/excel-parser';
import { autoRegisterEmails } from '@/lib/aliases';

export const dynamic = 'force-dynamic';

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file provided' }, { status: 400 });

    const buffer = Buffer.from(await file.arrayBuffer());
    const result = parsePerigeeExport(buffer);

    // Auto-register any new rep emails into the alias store (without an alias)
    // and tell the client which ones are new so it can warn the user.
    let newEmails: { dataEmail: string; displayName?: string }[] = [];
    try {
      const added = await autoRegisterEmails(
        result.reps.map((r) => ({ dataEmail: r.email, displayName: r.fullName }))
      );
      newEmails = added.map((e) => ({
        dataEmail: e.dataEmail,
        displayName: e.displayName,
      }));
    } catch (err) {
      console.error('[parse] alias auto-register failed:', err instanceof Error ? err.message : err);
      // Non-fatal — continue
    }

    // Don't send visits back to client (can be large) — just metadata
    return NextResponse.json(
      {
        reps:      result.reps,
        channels:  result.channels,
        dateRange: result.dateRange,
        rowCount:  result.rowCount,
        newEmails,
      },
      { headers: { 'Cache-Control': 'no-store' } }
    );
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 422 });
  }
}
