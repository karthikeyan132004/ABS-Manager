import { IFieldInfo } from '@pnp/sp/fields';
import { IList } from '@pnp/sp/lists';
import { getSP } from '../../../pnpConfig';
import { ALL_STATUSES, SALES_STATUSES } from '../../../statuses';
import { getSiteListTitles } from './DashboardService';

/** A Quotation list item. Id and Title are always present. */
export interface IQuotationItem {
  Id: number;
  Title?: string;
  [internalName: string]: unknown;
}

/** Which list and which internal field names the pipeline works on. */
export interface IQuotationConfig {
  listTitle: string;
  statusField: string;
  amountField: string;
  remarksField: string;
}

/**
 * Mirrors a quotation's sales stage onto the Sales Lead record it is linked to.
 * Both ends are lookups: SalesLead on Quotation points at the lead, and the
 * lead's own stage column points at Ref Quotation Status.
 */
export interface ILeadSyncConfig {
  /** The Sales Lead list. */
  leadListTitle: string;
  /** Lookup on Quotation pointing at Sales Lead, e.g. "SalesLead". */
  quotationLeadField: string;
  /**
   * Lookup on Sales Lead holding the stage, e.g. "CustomerName" - note that a
   * renamed column keeps its original internal name, so this rarely matches
   * what the column is called on screen.
   */
  leadStageField: string;
}

/** How the status column stores its values, which decides how we read and write it. */
export type StatusKind = 'choice' | 'lookup' | 'text';

export interface IQuotationLoad {
  items: IQuotationItem[];
  /** The real values the status column accepts. Empty when it is free text. */
  statusChoices: string[];
  statusKind: StatusKind;
  /** Label -> lookup item id. Only populated when statusKind is 'lookup'. */
  statusIdByLabel: Record<string, number>;
  /** Lookup item id -> label, used to resolve each item's status without $expand. */
  statusLabelById: Record<number, string>;
  /** Sales Lead item id -> its Lead ID, for showing which lead a quotation belongs to. */
  leadLabelById: Record<number, string>;
  /** Status label -> id in the list the lead's stage column points at. */
  leadStageIdByLabel: Record<string, number>;
  /** Configured field names that do not exist on the list. */
  missingFields: string[];
  /** Internal name -> TypeAsString, used to coerce values before saving. */
  fieldTypes: Record<string, string>;
  /** Fatal load error. When set, nothing else on the object is meaningful. */
  error?: string;
  /** Lists that do exist on this web, offered when the configured one does not. */
  availableLists?: string[];
}

const NUMERIC_TYPES: Set<string> = new Set(['Number', 'Currency', 'Integer', 'Counter']);

export const EMPTY_QUOTATION_LOAD: IQuotationLoad = {
  items: [],
  statusChoices: [],
  statusKind: 'text',
  statusIdByLabel: {},
  statusLabelById: {},
  leadLabelById: {},
  leadStageIdByLabel: {},
  missingFields: [],
  fieldTypes: {}
};

function getList(config: IQuotationConfig): IList {
  return getSP().web.lists.getByTitle(config.listTitle.trim());
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

/**
 * Reads the status as a display string. An expanded lookup arrives as an object,
 * a choice or text column as a plain string, so both are flattened here and every
 * caller goes through this rather than reading the field directly.
 */
export function readStatusValue(
  item: IQuotationItem,
  internalName: string,
  labelById?: Record<number, string>
): string {
  if (internalName === '') {
    return '';
  }

  // A lookup is stored on the companion "<name>Id" column. Resolving it against
  // the reference list avoids $expand, which SharePoint ignores alongside
  // $select=*, and does not care what the lookup's display column is called.
  if (labelById !== undefined) {
    const id: unknown = item[`${internalName}Id`];

    if (Array.isArray(id)) {
      return id
        .map((entry: unknown): string => labelById[Number(entry)] ?? '')
        .filter((label: string): boolean => label !== '')
        .join(', ');
    }

    if (typeof id === 'number' || (typeof id === 'string' && id !== '')) {
      return labelById[Number(id)] ?? '';
    }
  }

  const raw: unknown = item[internalName];
  if (raw === null || raw === undefined) {
    return '';
  }

  if (Array.isArray(raw)) {
    return raw
      .map((entry: unknown): string => (isRecord(entry) ? readExpandedLabel(entry) : String(entry)))
      .filter((text: string): boolean => text !== '')
      .join(', ');
  }

  if (isRecord(raw)) {
    return readExpandedLabel(raw);
  }

  return String(raw);
}

/**
 * An expanded lookup arrives as { Id, <display column> }. That column is only
 * called Title by convention, so fall back to the single other property rather
 * than assuming it.
 */
function readExpandedLabel(raw: Record<string, unknown>): string {
  if (typeof raw.Title === 'string' && raw.Title !== '') {
    return raw.Title;
  }
  if (typeof raw.Label === 'string' && raw.Label !== '') {
    return raw.Label;
  }

  for (const key of Object.keys(raw)) {
    if (key === 'Id' || key === 'ID' || key === 'odata.type') {
      continue;
    }
    const value: unknown = raw[key];
    if (typeof value === 'string' && value !== '') {
      return value;
    }
  }

  return '';
}

/**
 * Builds the update payload for a status change. A lookup is written by id on
 * the companion "<name>Id" field; a choice or text column takes the label.
 */
export function buildStatusUpdate(
  config: IQuotationConfig,
  load: IQuotationLoad,
  label: string
): Record<string, unknown> | undefined {
  const internalName: string = config.statusField.trim();

  if (load.statusKind !== 'lookup') {
    return { [internalName]: label };
  }

  const id: number | undefined = load.statusIdByLabel[label];
  if (id === undefined) {
    return undefined; // caller reports which label is missing from the lookup list
  }

  return { [`${internalName}Id`]: id };
}

interface ILookupFieldInfo extends IFieldInfo {
  LookupList?: string;
  LookupField?: string;
}

/**
 * Reads the allowed status values. For a Choice column that is the Choices
 * array; for a Lookup it is the items of the list the column points at, which
 * also gives the label -> id map needed to write the field.
 */
async function loadStatusDomain(
  config: IQuotationConfig,
  kind: StatusKind
): Promise<{
  choices: string[];
  idByLabel: Record<string, number>;
  labelById: Record<number, string>;
  displayField: string;
}> {
  const empty = {
    choices: [] as string[],
    idByLabel: {} as Record<string, number>,
    labelById: {} as Record<number, string>,
    displayField: 'Title'
  };

  try {
    const field: ILookupFieldInfo = await getList(config)
      .fields.getByInternalNameOrTitle(config.statusField.trim())() as ILookupFieldInfo;

    if (kind === 'choice') {
      const choices: unknown = (field as unknown as Record<string, unknown>).Choices;
      return Array.isArray(choices)
        ? { ...empty, choices: choices.map((choice: unknown): string => String(choice)) }
        : empty;
    }

    if (kind !== 'lookup' || field.LookupList === undefined || field.LookupList === '') {
      return empty;
    }

    const displayField: string = field.LookupField !== undefined && field.LookupField !== ''
      ? field.LookupField
      : 'Title';

    // LookupList is a guid, sometimes wrapped in braces.
    const listId: string = field.LookupList.replace(/[{}]/g, '');
    const rows: Record<string, unknown>[] = await getSP().web.lists
      .getById(listId)
      .items.select('Id', displayField).top(500)();

    const idByLabel: Record<string, number> = {};
    const labelById: Record<number, string> = {};
    const choices: string[] = [];
    for (const row of rows) {
      const label: string = String(row[displayField] ?? '');
      if (label !== '') {
        idByLabel[label] = Number(row.Id);
        labelById[Number(row.Id)] = label;
        choices.push(label);
      }
    }

    return { choices, idByLabel, labelById, displayField };
  } catch {
    return empty;
  }
}

/**
 * Loads the quotations plus enough list metadata to tell the manager when the
 * configuration is wrong. Items are fetched with $select=* so a mistyped field
 * name shows up as an empty column and a warning, not a 400; a lookup status is
 * expanded so its label comes back rather than just an id.
 */
export async function loadQuotations(
  config: IQuotationConfig,
  itemLimit: number,
  sync?: ILeadSyncConfig
): Promise<IQuotationLoad> {
  if (config.listTitle.trim() === '') {
    return { ...EMPTY_QUOTATION_LOAD, error: 'No quotation list is configured. Set it in the web part properties.' };
  }

  try {
    const list: IList = getList(config);

    const fields: IFieldInfo[] = await list.fields
      .select('InternalName', 'TypeAsString')
      .filter('Hidden eq false')();

    const fieldTypes: Record<string, string> = {};
    for (const field of fields) {
      fieldTypes[field.InternalName] = field.TypeAsString;
    }

    const statusField: string = config.statusField.trim();
    const configured: string[] = [statusField, config.amountField, config.remarksField]
      .map((name: string): string => name.trim())
      .filter((name: string): boolean => name !== '');

    const missingFields: string[] = configured.filter((name: string): boolean => fieldTypes[name] === undefined);

    const statusType: string | undefined = fieldTypes[statusField];
    const statusKind: StatusKind = statusType === 'Lookup' || statusType === 'LookupMulti'
      ? 'lookup'
      : statusType === 'Choice' || statusType === 'MultiChoice'
        ? 'choice'
        : 'text';

    const domain = missingFields.indexOf(statusField) >= 0
      ? { choices: [], idByLabel: {}, labelById: {}, displayField: 'Title' }
      : await loadStatusDomain(config, statusKind);

    // $select=* returns every column including the lookup's "<name>Id", which is
    // all that is needed; SharePoint ignores lookup projections next to "*".
    const items: IQuotationItem[] = await list.items.top(itemLimit).orderBy('Modified', false)();

    const leadData = sync !== undefined
      ? await loadLeadSyncData(sync)
      : { leadLabelById: {}, leadStageIdByLabel: {} };

    return {
      items,
      leadLabelById: leadData.leadLabelById,
      leadStageIdByLabel: leadData.leadStageIdByLabel,
      statusChoices: domain.choices,
      statusKind,
      statusIdByLabel: domain.idByLabel,
      statusLabelById: domain.labelById,
      missingFields,
      fieldTypes
    };
  } catch (e) {
    const message: string = e instanceof Error ? e.message : String(e);
    // A wrong list title is the usual cause, so offer the ones that do exist.
    const availableLists: string[] = await getSiteListTitles();
    return { ...EMPTY_QUOTATION_LOAD, error: `Could not load "${config.listTitle}": ${message}`, availableLists };
  }
}

export interface IStatusMismatch {
  /** In SharePoint but not in src/statuses.ts, so the pipeline cannot order them. */
  missingFromCode: string[];
  /** In src/statuses.ts but not in SharePoint, so writing them would fail. */
  missingFromList: string[];
}

/** Compares the hard-coded pipeline against the values the column really accepts. */
export function compareStatuses(statusChoices: string[]): IStatusMismatch {
  if (statusChoices.length === 0) {
    return { missingFromCode: [], missingFromList: [] };
  }

  return {
    missingFromCode: statusChoices.filter((choice: string): boolean => ALL_STATUSES.indexOf(choice) < 0),
    missingFromList: ALL_STATUSES.filter((status: string): boolean => statusChoices.indexOf(status) < 0)
  };
}

/**
 * Coerces a form value to what the column expects. Number and Currency columns
 * reject the strings a TextField produces, and an emptied box has to clear the
 * value rather than write 0.
 */
export function coerceFieldValue(
  value: string,
  internalName: string,
  fieldTypes: Record<string, string>
  // null is required here: SharePoint clears a Number column on an explicit null,
  // whereas undefined is dropped from the JSON body and leaves the old value.
  // eslint-disable-next-line @rushstack/no-new-null
): number | string | null {
  const type: string | undefined = fieldTypes[internalName];

  if (type !== undefined && NUMERIC_TYPES.has(type)) {
    const trimmed: string = value.trim();
    if (trimmed === '') {
      return null;
    }
    const parsed: number = Number(trimmed.replace(/,/g, ''));
    return isNaN(parsed) ? null : parsed;
  }

  return value;
}

/** True when the value typed into a numeric box cannot be saved. */
export function isInvalidNumber(value: string, internalName: string, fieldTypes: Record<string, string>): boolean {
  const type: string | undefined = fieldTypes[internalName];
  if (type === undefined || !NUMERIC_TYPES.has(type)) {
    return false;
  }

  const trimmed: string = value.trim();
  return trimmed !== '' && isNaN(Number(trimmed.replace(/,/g, '')));
}

/** Writes the given fields back to one quotation. */
export async function updateQuotation(
  config: IQuotationConfig,
  id: number,
  fields: Record<string, unknown>
): Promise<void> {
  await getList(config).items.getById(id).update(fields);
}

/** Whole days since the item last changed, or undefined when Modified is absent. */
export function daysSinceModified(item: IQuotationItem): number | undefined {
  const raw: unknown = item.Modified;
  if (raw === null || raw === undefined || raw === '') {
    return undefined;
  }

  const modified: Date = new Date(String(raw));
  if (isNaN(modified.getTime())) {
    return undefined;
  }

  const millis: number = Date.now() - modified.getTime();
  return Math.max(0, Math.floor(millis / (24 * 60 * 60 * 1000)));
}

/** A quotation nobody has touched for this long counts as stalled. */
export const STALE_AFTER_DAYS: number = 14;

/** Reads a Currency or Number column off an item. Undefined when not set. */
export function readAmount(item: IQuotationItem, internalName: string): number | undefined {
  if (internalName === '') {
    return undefined;
  }

  const raw: unknown = item[internalName];
  if (raw === null || raw === undefined || raw === '') {
    return undefined;
  }

  const value: number = Number(raw);
  return isNaN(value) ? undefined : value;
}

/** Sums a Currency column across items, skipping the ones with no value. */
export function sumAmount(items: IQuotationItem[], internalName: string): number {
  let total: number = 0;
  for (const item of items) {
    total += readAmount(item, internalName) ?? 0;
  }
  return total;
}

/**
 * Compact money for a tile: 1.2M / 340.5K / 820. Totals on a dashboard are read
 * for magnitude, and the full figure stays available in the grid.
 */
export function formatAmount(value: number): string {
  const abs: number = Math.abs(value);

  if (abs >= 1000000) {
    return `${(value / 1000000).toFixed(1)}M`;
  }
  if (abs >= 1000) {
    return `${(value / 1000).toFixed(1)}K`;
  }
  return Math.round(value).toLocaleString();
}

/** True when both ends of the lead link are configured. */
export function isLeadSyncConfigured(sync: ILeadSyncConfig): boolean {
  return sync.leadListTitle.trim() !== ''
    && sync.quotationLeadField.trim() !== ''
    && sync.leadStageField.trim() !== '';
}

/**
 * Loads what the lead sync needs: the leads themselves (to label the link) and
 * the value domain of the lead's stage column. That domain is read from the
 * list the column actually points at, rather than assuming it is the same
 * reference list the quotation status uses.
 */
export async function loadLeadSyncData(
  sync: ILeadSyncConfig
): Promise<{ leadLabelById: Record<number, string>; leadStageIdByLabel: Record<string, number> }> {
  const empty = { leadLabelById: {} as Record<number, string>, leadStageIdByLabel: {} as Record<string, number> };

  if (!isLeadSyncConfigured(sync)) {
    return empty;
  }

  try {
    const leadList: IList = getSP().web.lists.getByTitle(sync.leadListTitle.trim());

    const leads: Record<string, unknown>[] = await leadList.items.select('Id', 'Title').top(2000)();
    const leadLabelById: Record<number, string> = {};
    for (const lead of leads) {
      leadLabelById[Number(lead.Id)] = String(lead.Title ?? `Lead ${lead.Id}`);
    }

    const stageField: ILookupFieldInfo = await leadList
      .fields.getByInternalNameOrTitle(sync.leadStageField.trim())() as ILookupFieldInfo;

    const leadStageIdByLabel: Record<string, number> = {};
    if (stageField.LookupList !== undefined && stageField.LookupList !== '') {
      const displayField: string = stageField.LookupField !== undefined && stageField.LookupField !== ''
        ? stageField.LookupField
        : 'Title';

      const rows: Record<string, unknown>[] = await getSP().web.lists
        .getById(stageField.LookupList.replace(/[{}]/g, ''))
        .items.select('Id', displayField).top(500)();

      for (const row of rows) {
        const label: string = String(row[displayField] ?? '');
        if (label !== '') {
          leadStageIdByLabel[label] = Number(row.Id);
        }
      }
    }

    return { leadLabelById, leadStageIdByLabel };
  } catch {
    return empty;
  }
}

/** The lead this quotation is linked to, or undefined when the link is empty. */
export function getLinkedLeadId(item: IQuotationItem, sync: ILeadSyncConfig): number | undefined {
  const raw: unknown = item[`${sync.quotationLeadField.trim()}Id`];
  if (raw === null || raw === undefined || raw === '') {
    return undefined;
  }

  const id: number = Number(raw);
  return isNaN(id) || id <= 0 ? undefined : id;
}

export type LeadSyncOutcome =
  | { kind: 'synced'; leadLabel: string }
  | { kind: 'skipped'; reason: string };

/**
 * Mirrors a sales-stage change onto the linked lead.
 *
 * Only stages 01-08 are mirrored: once a quotation passes Quotation Draft the
 * quotation owns the record, and the lead keeps the last sales stage it reached.
 * Every skip returns a reason rather than failing, so a quotation with no lead
 * linked never blocks the status change that already succeeded.
 */
export async function syncLeadStage(
  item: IQuotationItem,
  newStatus: string,
  sync: ILeadSyncConfig,
  load: IQuotationLoad
): Promise<LeadSyncOutcome> {
  if (!isLeadSyncConfigured(sync)) {
    return { kind: 'skipped', reason: 'lead sync is not configured' };
  }

  if (SALES_STATUSES.indexOf(newStatus) < 0) {
    return { kind: 'skipped', reason: 'only sales stages 01-08 are mirrored to the lead' };
  }

  const leadId: number | undefined = getLinkedLeadId(item, sync);
  if (leadId === undefined) {
    return { kind: 'skipped', reason: 'this quotation has no Sales Lead linked' };
  }

  const stageId: number | undefined = load.leadStageIdByLabel[newStatus];
  if (stageId === undefined) {
    return { kind: 'skipped', reason: `"${newStatus}" is not a value the lead's stage column accepts` };
  }

  await getSP().web.lists
    .getByTitle(sync.leadListTitle.trim())
    .items.getById(leadId)
    .update({ [`${sync.leadStageField.trim()}Id`]: stageId });

  return { kind: 'synced', leadLabel: load.leadLabelById[leadId] ?? `Lead ${leadId}` };
}
