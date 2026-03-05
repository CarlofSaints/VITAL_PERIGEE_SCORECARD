export interface QuestionDef {
  id: string;
  section: string;
  text: string; // scorecard display text
  rawHeader: string; // exact column header in Perigee export
  inverted?: boolean; // No = good (store manager question)
  optional?: boolean; // not yet in raw data – shows as N/A until added
}

export const SECTIONS = [
  'ON SHELF AVAILABILITY',
  'PROMOTION',
  'PRICING',
  'SHELF HEALTH',
  'MULTI PLACEMENTS / FORWARD SHARE',
  'POS',
  'GENERAL',
] as const;

export const QUESTIONS: QuestionDef[] = [
  // ── ON SHELF AVAILABILITY ─────────────────────────────────────────────────
  {
    id: 'osa-sku',
    section: 'ON SHELF AVAILABILITY',
    text: "Are all ranged SKU's on shelf",
    rawHeader: "Are all ranged SKU's on shelf?",
  },
  {
    id: 'osa-flow',
    section: 'ON SHELF AVAILABILITY',
    text: 'Is the correct merchandising flow implemented',
    rawHeader: 'Is the correct merchandising flow implemented?',
  },
  {
    id: 'osa-oos',
    section: 'ON SHELF AVAILABILITY',
    text: "Has OOS SKU's been flagged to relevant manager",
    rawHeader: "Has OOS SKU's been flagged to relevant manager?",
  },
  {
    id: 'osa-npd',
    section: 'ON SHELF AVAILABILITY',
    text: 'Is the relevant NPD on shelf',
    rawHeader: 'Is the relevant NPD on shelf?',
  },

  // ── PROMOTION ─────────────────────────────────────────────────────────────
  {
    id: 'promo-active',
    section: 'PROMOTION',
    text: 'Is the scheduled Promotion active',
    rawHeader: 'Is the scheduled Promotion active?',
    optional: true,
  },
  {
    id: 'promo-comms',
    section: 'PROMOTION',
    text: 'Is the Promotional communication on shelf and visible',
    rawHeader: 'Is the Promotional communication on shelf and visible?',
    optional: true,
  },

  // ── PRICING ───────────────────────────────────────────────────────────────
  {
    id: 'price-pi',
    section: 'PRICING',
    text: "Do all Vital SKU's have PI labels on shelf",
    rawHeader: "Do all Vital SKU's have PI labels on shelf?",
  },
  {
    id: 'price-correct',
    section: 'PRICING',
    text: 'Is the price updated and correct',
    rawHeader: 'Is the price updated and correct?',
  },
  {
    id: 'price-files',
    section: 'PRICING',
    text: 'Has Pricing files been shared with agents: Vital/Customer',
    rawHeader: 'Has Pricing files been shared with agents: Vital/Customer?',
    optional: true,
  },
  {
    id: 'price-tracking',
    section: 'PRICING',
    text: 'Are the agents order taking being tracked against correct price files - SPAR/WHL',
    rawHeader:
      'Are the agents order taking being tracked against correct price files - SPAR/WHL?',
    optional: true,
  },

  // ── SHELF HEALTH ──────────────────────────────────────────────────────────
  {
    id: 'sh-clean',
    section: 'SHELF HEALTH',
    text: 'Are all product clean and presentable',
    rawHeader: 'Are all product on shelf clean and presentable?',
  },
  {
    id: 'sh-fifo',
    section: 'SHELF HEALTH',
    text: 'Has the FIFO (First In First Out) rule been applied',
    rawHeader: 'Has the FIFO (First In First Out) rule been applied?',
  },
  {
    id: 'sh-shortdated',
    section: 'SHELF HEALTH',
    text: 'Has short dated stock issues been raised with Rep/Vital FSC',
    rawHeader: 'Has short dated stock issues been raised with Rep / Vital FSC?',
  },
  {
    id: 'sh-expired',
    section: 'SHELF HEALTH',
    text: 'Has expired stock been removed and claim raised',
    rawHeader:
      'Has all damages and expired stock been marked down or written up as required',
  },

  // ── MULTI PLACEMENTS / FORWARD SHARE ──────────────────────────────────────
  {
    id: 'mp-available',
    section: 'MULTI PLACEMENTS / FORWARD SHARE',
    text: 'Is there multi placements available in the store',
    rawHeader: 'Is there multi placements available in the store?',
  },
  {
    id: 'mp-forward',
    section: 'MULTI PLACEMENTS / FORWARD SHARE',
    text: 'Has there been any forward share gains in the store on the home shelf space',
    rawHeader:
      'Has there been any forward share gains in the store on the home shelf?',
  },
  {
    id: 'mp-identified',
    section: 'MULTI PLACEMENTS / FORWARD SHARE',
    text: 'Has any multi placements been identified and shared with the relevant manager in store',
    rawHeader: 'Has any multi placements been identified and implemented',
  },

  // ── POS ───────────────────────────────────────────────────────────────────
  {
    id: 'pos-implemented',
    section: 'POS',
    text: 'Has the POS been implemented where required',
    rawHeader: 'Has the POS been implemented where required?',
    optional: true,
  },

  // ── GENERAL ───────────────────────────────────────────────────────────────
  {
    id: 'gen-mgr-issues',
    section: 'GENERAL',
    text: 'Does the store manager have any issues to be addressed',
    rawHeader: 'Does the store manager have any issues to be addressed',
    inverted: true,
  },
  {
    id: 'gen-opportunities',
    section: 'GENERAL',
    text: 'Is there any additional opportunities that has been identified that need to be implemented',
    rawHeader:
      'Is there any additional opportunities that has been identified that need to be implemented',
  },
];

export const TOTAL_QUESTIONS = QUESTIONS.length; // 20
