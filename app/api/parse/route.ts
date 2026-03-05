import { NextRequest, NextResponse } from 'next/server';
import { parsePerigeeExport } from '@/lib/excel-parser';

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'No file provided' }, { status: 400 });

    const buffer = Buffer.from(await file.arrayBuffer());
    const result = parsePerigeeExport(buffer);

    // Don't send visits back to client (can be large) — just metadata
    return NextResponse.json({
      reps:      result.reps,
      channels:  result.channels,
      dateRange: result.dateRange,
      rowCount:  result.rowCount,
    });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 422 });
  }
}
