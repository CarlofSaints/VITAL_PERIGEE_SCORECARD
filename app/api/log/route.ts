import { NextResponse } from 'next/server';
import { readActivityLog } from '@/lib/activityLog';

export async function GET() {
  try {
    const entries = await readActivityLog();
    return NextResponse.json({ entries });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
