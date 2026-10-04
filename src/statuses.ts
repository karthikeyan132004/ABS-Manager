/**
 * The Quotation pipeline, in order.
 *
 * These are the 21 items of the "Ref Quotation Status" list, copied exactly as
 * that list spells them - including the "NN. " prefixes and the two spellings
 * that look like mistakes but are the real stored values:
 *   - "02. Inital Client Contact" (not "Initial")
 *   - "20. Quotation Cancelled"   (spelled correctly; an earlier draft of this
 *                                  file wrongly used "Canclled")
 * The status column is a Lookup, so a value that does not match one of these
 * character for character has no id to write and cannot be saved. The dashboard
 * re-checks this list on every load and warns about any drift.
 *
 * The "NN. " prefix is stripped for display only - see displayStatus().
 */
export const ALL_STATUSES: string[] = [
  // --- Sales stage (01-08) ---
  '01. Lead Identified',
  '02. Inital Client Contact',
  '03. Cold Call / Introduction Call',
  '04. Company Profile Sent',
  '05. Scaffolding Services Presentation',
  '06. Client Requirement Discussion',
  '07. Sales Site Visit',
  '08. Scope of Work Reviewed',
  // --- Quotation stage (09-21) ---
  '09. Quotation Draft',
  '10. Price Request',
  '11. Proceed to Submit Quotation',
  '12. Quotation Submitted',
  '13. Pending for Approval',
  '14. Client Meeting',
  '15. Tender / Bid Submitted',
  '16. Commercial Negotiation',
  '17. Purchase Order Follow-up',
  '18. Revised Quotation Submitted',
  '19. Quotation Approved',
  '20. Quotation Cancelled',
  '21. Quotation Lost'
];

/**
 * Where the sales stage ends. 01-08 are the Sales Lead view, 09-21 are the
 * Customer Quotation view, so "09. Quotation Draft" sits on the quotation side.
 */
export const SALES_STAGE_COUNT: number = 8;

/** The 8 statuses behind the Sales section. */
export const SALES_STATUSES: string[] = ALL_STATUSES.slice(0, SALES_STAGE_COUNT);

/** The 13 statuses behind the Customer Quotation section. */
export const QUOTATION_STATUSES: string[] = ALL_STATUSES.slice(SALES_STAGE_COUNT);

/** Terminal statuses. Nothing moves out of these once locking is switched on. */
export const FINAL_STATUSES: string[] = [
  '19. Quotation Approved',
  '20. Quotation Cancelled',
  '21. Quotation Lost'
];

/**
 * Flip to true once approvals go live and Approved / Cancelled / Lost should no
 * longer be editable. Left false for now, as agreed.
 */
export const LOCK_FINAL_STATUSES: boolean = false;

/** Drops the "NN. " ordering prefix for display. Never use this for matching. */
export function displayStatus(status: string): string {
  return status.replace(/^\s*\d+\.\s*/, '');
}

export function isFinalStatus(status: string): boolean {
  return FINAL_STATUSES.indexOf(status) >= 0;
}

/** True when the item sits in a terminal status and locking is switched on. */
export function isLocked(status: string): boolean {
  return LOCK_FINAL_STATUSES && isFinalStatus(status);
}

/**
 * The next or previous status, or undefined when there is nowhere to go.
 * An unrecognised status returns undefined rather than silently resetting the
 * item to the first stage.
 */
export function stepStatus(current: string, step: 1 | -1): string | undefined {
  const index: number = ALL_STATUSES.indexOf(current);
  if (index < 0) {
    return undefined;
  }

  const next: number = index + step;
  return next >= 0 && next < ALL_STATUSES.length ? ALL_STATUSES[next] : undefined;
}
