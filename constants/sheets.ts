import { QUESTIONS } from './questions';

/**
 * Short, human-friendly exception-report sheet name per question id
 * (max 31 chars for Excel). Shared by the exception builder (to name sheets)
 * and the UI / photo-question mapping (to offer sheets as targets).
 */
export const SHEET_NAMES: Record<string, string> = {
  'osa-sku':          'Ranged SKU Report',
  'osa-flow':         'Merchandising Flow',
  'osa-oos':          'OOS Flagged',
  'osa-npd':          'NPD Report',
  'osa-backup':       'OOS Backup Stock',
  'promo-active':     'Promotion Active',
  'promo-comms':      'Promo Communication',
  'price-pi':         'PI Labels',
  'price-correct':    'Pricing Correct',
  'price-files':      'Pricing Files Shared',
  'price-tracking':   'Pricing Tracking',
  'sh-clean':         'Clean & Presentable',
  'sh-fifo':          'FIFO Rule',
  'sh-shortdated':    'Short Dated Stock',
  'sh-expired':       'Damaged & Expired',
  'sh-expired-check': 'Expired Stock',
  'mp-available':     'Multi Placements',
  'mp-forward':       'Forward Share',
  'mp-identified':    'Multi Place Identified',
  'pos-implemented':  'POS Implemented',
  'gen-mgr-issues':   'Store Manager Issues',
  'gen-opportunities': 'Opportunities',
};

/** Reverse lookup: sheet name → question id. */
export const SHEET_TO_QUESTION_ID: Record<string, string> = Object.fromEntries(
  Object.entries(SHEET_NAMES).map(([id, name]) => [name, id])
);

/**
 * Ordered, de-duplicated list of every exception sheet a report can produce.
 * Used to populate the target-sheet dropdown when the user maps a Photo
 * Question value to a sheet.
 */
export const SELECTABLE_SHEETS: string[] = (() => {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const q of QUESTIONS) {
    const name = SHEET_NAMES[q.id];
    if (name && !seen.has(name)) {
      seen.add(name);
      out.push(name);
    }
  }
  return out;
})();

/**
 * Normalise a "Photo …" question value into a stable lookup key:
 * collapse whitespace, trim, lowercase. Used both as the map key and when
 * resolving a value found in the raw export.
 */
export function normPhotoQuestion(s: string): string {
  return (s ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * Built-in default mapping of Photo Question value → target sheet name.
 * Keys are already normalised (see normPhotoQuestion). These four are known
 * from historical Perigee exports and never need to be mapped by hand.
 * Any Photo Question value not covered here (or in the persisted store) is
 * treated as "unknown" and prompted for in the UI.
 */
export const DEFAULT_PHOTO_QUESTION_MAP: Record<string, string> = {
  "photo oos sku's":                'OOS Flagged',
  'photo pos':                      'POS Implemented',
  'photo promotional communication':'Promo Communication',
  'photo stock pressure':           'Forward Share',
};
