import { IFieldInfo } from '@pnp/sp/fields';
import { IList } from '@pnp/sp/lists';
import { getSP } from '../../../pnpConfig';
import { IDashboardField, IKpiSet, IListTabConfig, IListTabData, ListItem } from '../models/IDashboardModels';
import { getDateValue, getNumberValue, isClosed, isExpandable } from './fieldValue';

/** Read-only fields that are still worth showing. */
const ALWAYS_KEEP: Set<string> = new Set(['ID', 'Id', 'Created', 'Modified', 'Author', 'Editor']);

/** Plumbing fields that are visible in the API but meaningless on a dashboard. */
const DENIED_FIELDS: Set<string> = new Set([
  'ContentType', 'ContentTypeId', 'Attachments', 'MetaInfo', 'Order', 'GUID', 'owshiddenversion',
  'FileSystemObjectType', 'ServerRedirectedEmbedUri', 'ServerRedirectedEmbedUrl', 'ComplianceAssetId',
  'FileLeafRef', 'FileDirRef', 'FileRef', 'AppAuthor', 'AppEditor', 'SyncClientId', 'ProgId',
  'ScopeId', 'UniqueId', 'InstanceID', 'WorkflowVersion', 'WorkflowInstanceID', 'ParentVersionString',
  'ParentLeafName', 'SortBehavior', 'NoExecute', 'OriginatorId', 'CheckoutUser', 'PrincipalCount'
]);

/** Field types that cannot be rendered usefully as a grid column. */
const DENIED_TYPES: Set<string> = new Set(['Computed', 'Attachments', 'Invalid', 'File', 'Guid', 'ContentTypeId']);

const AUDIT_FIELDS: Set<string> = new Set(['ID', 'Id', 'Created', 'Modified', 'Author', 'Editor']);

const DEFAULT_AUTO_COLUMNS: number = 7;

export function toDashboardField(field: IFieldInfo): IDashboardField {
  return {
    internalName: field.InternalName,
    title: field.Title !== '' ? field.Title : field.InternalName,
    typeAsString: field.TypeAsString
  };
}

/** Splits a comma or semicolon separated property pane value into clean names. */
export function parseNameList(value: string | undefined): string[] {
  if (value === undefined || value.trim() === '') {
    return [];
  }

  return value
    .split(/[,;]/)
    .map((name: string): string => name.trim())
    .filter((name: string): boolean => name.length > 0);
}

export function isUsable(field: IFieldInfo): boolean {
  if (field.Hidden === true) {
    return false;
  }
  if (DENIED_FIELDS.has(field.InternalName) || DENIED_TYPES.has(field.TypeAsString)) {
    return false;
  }
  if (field.InternalName.indexOf('OData__') === 0) {
    return false;
  }
  return field.ReadOnlyField !== true || ALWAYS_KEEP.has(field.InternalName);
}

/** Title first, then the remaining user columns, then the audit fields. */
function autoColumns(fields: IDashboardField[]): IDashboardField[] {
  const title: IDashboardField[] = fields.filter((f: IDashboardField): boolean => f.internalName === 'Title');
  const custom: IDashboardField[] = fields.filter(
    (f: IDashboardField): boolean => f.internalName !== 'Title' && !AUDIT_FIELDS.has(f.internalName)
  );
  const audit: IDashboardField[] = fields.filter((f: IDashboardField): boolean => AUDIT_FIELDS.has(f.internalName));

  return [...title, ...custom, ...audit].slice(0, DEFAULT_AUTO_COLUMNS);
}

/**
 * Builds $select and $expand so lookup and person columns come back with a
 * display value instead of just an id.
 */
export function buildQuery(fields: IDashboardField[]): { selects: string[]; expands: string[] } {
  const selects: string[] = ['Id'];
  const expands: string[] = [];

  for (const field of fields) {
    if (isExpandable(field)) {
      selects.push(`${field.internalName}/Id`, `${field.internalName}/Title`);
      expands.push(field.internalName);
    } else {
      selects.push(field.internalName);
    }
  }

  return { selects, expands };
}

/**
 * Loads one tab: the usable fields on the list, the columns to show, and the
 * items. A missing list or a bad field name is reported on the returned object
 * instead of thrown, so one broken tab does not take the dashboard down.
 */
export async function loadListTab(config: IListTabConfig, itemLimit: number): Promise<IListTabData> {
  const empty: IListTabData = { columns: [], allFields: [], items: [] };

  if (config.listTitle.trim() === '') {
    return {
      ...empty,
      error: `No list is configured for the "${config.headerText}" tab. Set it in the web part properties.`
    };
  }

  try {
    const list: IList = getSP().web.lists.getByTitle(config.listTitle.trim());

    const rawFields: IFieldInfo[] = await list.fields
      .select('InternalName', 'Title', 'TypeAsString', 'ReadOnlyField', 'Hidden')
      .filter('Hidden eq false')();

    const allFields: IDashboardField[] = rawFields.filter(isUsable).map(toDashboardField);
    const byName: Map<string, IDashboardField> = new Map(
      allFields.map((f: IDashboardField): [string, IDashboardField] => [f.internalName, f])
    );

    const requested: string[] = parseNameList(config.columns);
    const columns: IDashboardField[] = requested.length > 0
      ? requested
        .map((name: string): IDashboardField | undefined => byName.get(name))
        .filter((f: IDashboardField | undefined): f is IDashboardField => f !== undefined)
      : autoColumns(allFields);

    // KPI fields are fetched too, even when they are not shown as columns.
    const extras: IDashboardField[] = [config.dateField, config.statusField, config.probabilityField]
      .map((name: string): IDashboardField | undefined => (name.trim() === '' ? undefined : byName.get(name.trim())))
      .filter((f: IDashboardField | undefined): f is IDashboardField => f !== undefined)
      .filter((f: IDashboardField): boolean =>
        columns.every((c: IDashboardField): boolean => c.internalName !== f.internalName));

    const { selects, expands } = buildQuery([...columns, ...extras]);

    let query = list.items.select(...selects).top(itemLimit).orderBy('Modified', false);
    if (expands.length > 0) {
      query = query.expand(...expands);
    }

    const items: ListItem[] = await query();

    const unknown: string[] = requested.filter((name: string): boolean => !byName.has(name));
    const error: string | undefined = unknown.length > 0
      ? `Ignored unknown column(s): ${unknown.join(', ')}. Use internal names, not display names.`
      : undefined;

    return { columns, allFields, items, error };
  } catch (e) {
    const message: string = e instanceof Error ? e.message : String(e);
    return { ...empty, error: `Could not load "${config.listTitle}": ${message}` };
  }
}

const DAY_MS: number = 24 * 60 * 60 * 1000;

/** Counts the KPI tiles for one tab. Fields that are not configured are skipped. */
export function calculateKpis(data: IListTabData, config: IListTabConfig, probabilityThreshold: number): IKpiSet {
  const startOfToday: Date = new Date();
  startOfToday.setHours(0, 0, 0, 0);
  const weekAhead: Date = new Date(startOfToday.getTime() + 7 * DAY_MS);

  let open: number = 0;
  let overdue: number = 0;
  let dueThisWeek: number = 0;
  let highProbability: number = 0;

  for (const item of data.items) {
    const closed: boolean = isClosed(item, config.statusField.trim());
    if (!closed) {
      open += 1;
    }

    const due: Date | undefined = getDateValue(item, config.dateField.trim());
    if (due !== undefined && !closed) {
      if (due.getTime() < startOfToday.getTime()) {
        overdue += 1;
      } else if (due.getTime() <= weekAhead.getTime()) {
        dueThisWeek += 1;
      }
    }

    const probability: number | undefined = getNumberValue(item, config.probabilityField.trim());
    if (probability !== undefined && !closed) {
      // Percent fields arrive as a fraction, so 0.8 and 80 both mean 80%.
      const percent: number = probability <= 1 ? probability * 100 : probability;
      if (percent >= probabilityThreshold) {
        highProbability += 1;
      }
    }
  }

  return {
    total: data.items.length,
    open,
    overdue,
    dueThisWeek,
    highProbability: config.probabilityField.trim() === '' ? undefined : highProbability
  };
}

/**
 * Titles of the visible lists on the current web, used to help the manager fix
 * a mistyped list name instead of just showing them a 404.
 */
export async function getSiteListTitles(): Promise<string[]> {
  try {
    const lists: { Title: string }[] = await getSP().web.lists
      .select('Title')
      .filter('Hidden eq false')();

    return lists
      .map((list: { Title: string }): string => list.Title)
      .sort((a: string, b: string): number => a.localeCompare(b));
  } catch {
    return [];
  }
}
