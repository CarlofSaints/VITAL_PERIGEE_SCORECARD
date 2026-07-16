import { NextRequest, NextResponse } from 'next/server';
import {
  listPhotoMapEntries,
  upsertPhotoMapping,
  upsertPhotoMappings,
  removePhotoMapping,
} from '@/lib/photoMap';
import { SELECTABLE_SHEETS } from '@/constants/sheets';

// Mutable JSON state — never cache
export const dynamic = 'force-dynamic';

const NO_STORE = { 'Cache-Control': 'no-store' };

export async function GET() {
  try {
    const entries = await listPhotoMapEntries();
    return NextResponse.json({ entries, sheets: SELECTABLE_SHEETS }, { headers: NO_STORE });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500, headers: NO_STORE });
  }
}

interface UpsertBody {
  photoQuestion?: string;
  sheetName?: string;
  // Bulk variant
  entries?: { photoQuestion: string; sheetName: string }[];
}

export async function POST(req: NextRequest) {
  try {
    const body = (await req.json()) as UpsertBody;

    if (Array.isArray(body.entries)) {
      const saved = await upsertPhotoMappings(body.entries);
      return NextResponse.json({ saved }, { headers: NO_STORE });
    }

    if (!body.photoQuestion || !body.sheetName) {
      return NextResponse.json(
        { error: 'photoQuestion and sheetName are required' },
        { status: 400, headers: NO_STORE }
      );
    }

    const saved = await upsertPhotoMapping({
      photoQuestion: body.photoQuestion,
      sheetName: body.sheetName,
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
    const photoQuestion = url.searchParams.get('photoQuestion');
    if (!photoQuestion) {
      return NextResponse.json(
        { error: 'photoQuestion query param is required' },
        { status: 400, headers: NO_STORE }
      );
    }
    const removed = await removePhotoMapping(photoQuestion);
    return NextResponse.json({ removed }, { headers: NO_STORE });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500, headers: NO_STORE });
  }
}
