import { getSP } from '../../../pnpConfig';

/**
 * The quotation workbooks the ABS Quotation Management web part ("absquot")
 * produces.
 *
 * That web part fills an Excel template and uploads it to a document library on
 * the CRM site, stamping Customer Name, Project Name and Date of Quotation onto
 * the library item. The file is named "<quotation no>-<customer>-<project>.xlsx".
 * Nothing links the workbook back to the Quotation list item, so the dashboard
 * matches them up here instead.
 */

/** Spellings of the library, tried in turn - the same ones absquot uploads to. */
export const GENERATED_LIBRARY_CANDIDATES: string[] = [
  'Generated Quotations',
  'GeneratedQuotations',
  'Generated Quotation'
];

/** Display titles absquot writes the customer and project under. */
const CUSTOMER_TITLES: string[] = ['customer name', 'customername'];
const PROJECT_TITLES: string[] = ['project name', 'projectname'];

export interface IGeneratedQuotation {
  /** File name including the extension, e.g. "Q-1042-Acme-Tower B.xlsx". */
  fileName: string;
  /** Server-relative path, used when no GUID is available. */
  serverRelativeUrl: string;
  /** The file GUID. Preferred, since a name with "#" cannot be addressed by path. */
  uniqueId: string;
  customerName: string;
  projectName: string;
  /** Last modified, used to pick the newest when several files match. */
  modified: string;
}

export interface IGeneratedQuotationIndex {
  items: IGeneratedQuotation[];
  /** The library the files were actually found in, for messages. */
  libraryName: string;
  /** Set when no candidate library could be read. */
  error?: string;
}

export const EMPTY_GENERATED_INDEX: IGeneratedQuotationIndex = { items: [], libraryName: '' };

function normalise(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Drops punctuation so a file name part matches the column it came from. */
function normaliseKey(value: string): string {
  return value.trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

/**
 * The internal names behind the Customer Name and Project Name columns.
 *
 * absquot matches those columns by display title because a renamed column keeps
 * its original internal name, so the same lookup has to happen here rather than
 * assuming the internal names.
 */
async function resolveColumns(libraryName: string): Promise<{ customer: string; project: string }> {
  const fields: { InternalName: string; Title: string }[] = await getSP().web.lists
    .getByTitle(libraryName)
    .fields.select('InternalName', 'Title').filter('Hidden eq false')();

  let customer: string = '';
  let project: string = '';

  for (const field of fields) {
    const title: string = normalise(field.Title ?? '');
    if (customer === '' && CUSTOMER_TITLES.indexOf(title) >= 0) {
      customer = field.InternalName;
    }
    if (project === '' && PROJECT_TITLES.indexOf(title) >= 0) {
      project = field.InternalName;
    }
  }

  return { customer, project };
}

/** Whether a failure says the library is absent rather than that it refused us. */
function isMissingLibrary(message: string): boolean {
  const text: string = message.toLowerCase();
  return text.indexOf('(404)') >= 0 || text.indexOf('does not exist') >= 0 || text.indexOf('not found') >= 0;
}

/** A column value flattened to text; a lookup arrives as { Id, Title }. */
function readColumn(row: Record<string, unknown>, internalName: string): string {
  if (internalName === '') {
    return '';
  }

  const raw: unknown = row[internalName];
  if (raw === null || raw === undefined) {
    return '';
  }

  if (isRecord(raw)) {
    return String(raw.Title ?? raw.Label ?? '');
  }

  return String(raw);
}

/**
 * Reads the generated workbooks. The library spellings are tried in turn, the
 * same way absquot tries them when uploading, so whichever one the site really
 * has is the one that answers.
 */
export async function loadGeneratedQuotations(itemLimit: number): Promise<IGeneratedQuotationIndex> {
  let lastError: string = '';

  for (const libraryName of GENERATED_LIBRARY_CANDIDATES) {
    try {
      const columns = await resolveColumns(libraryName);

      const selects: string[] = ['Id', 'Title', 'Modified', 'FileLeafRef', 'FileRef', 'UniqueId'];
      if (columns.customer !== '') {
        selects.push(columns.customer);
      }
      if (columns.project !== '') {
        selects.push(columns.project);
      }

      const rows: Record<string, unknown>[] = await getSP().web.lists
        .getByTitle(libraryName)
        .items.select(...selects).top(itemLimit).orderBy('Modified', false)();

      const items: IGeneratedQuotation[] = rows.map((row: Record<string, unknown>): IGeneratedQuotation => ({
        fileName: String(row.FileLeafRef ?? ''),
        serverRelativeUrl: String(row.FileRef ?? ''),
        uniqueId: String(row.UniqueId ?? ''),
        customerName: readColumn(row, columns.customer),
        projectName: readColumn(row, columns.project),
        modified: String(row.Modified ?? '')
      }));

      return { items, libraryName };
    } catch (e) {
      lastError = e instanceof Error ? e.message : String(e);
      // Only "no such library" moves on to the next spelling; a library that
      // exists and refused the read is reported rather than skipped past.
      if (!isMissingLibrary(lastError)) {
        break;
      }
    }
  }

  return {
    ...EMPTY_GENERATED_INDEX,
    error: lastError === ''
      ? `None of these libraries exist on this site: ${GENERATED_LIBRARY_CANDIDATES.join(', ')}.`
      : lastError
  };
}

/**
 * The workbook for one quotation.
 *
 * The quotation number is tried first: absquot names the file
 * "<quotation no>-<customer>-<project>.xlsx", so the number is an exact and
 * unique key. Customer plus project is the fallback, which covers a file whose
 * name drifted from the number; the newest wins there, since several
 * quotations can exist for one customer and project.
 */
export function findGeneratedQuotation(
  index: IGeneratedQuotationIndex,
  quotationNumber: string,
  customerName: string,
  projectName: string
): IGeneratedQuotation | undefined {
  const number: string = normaliseKey(quotationNumber);

  if (number !== '') {
    const byNumber: IGeneratedQuotation | undefined = index.items.find(
      (file: IGeneratedQuotation): boolean => {
        const base: string = file.fileName.replace(/\.[^.]+$/, '');
        // A quotation number can itself contain "-", so the name is matched by
        // prefix rather than split on the separator.
        return normaliseKey(base).indexOf(number) === 0;
      }
    );

    if (byNumber !== undefined) {
      return byNumber;
    }
  }

  const customer: string = normaliseKey(customerName);
  const project: string = normaliseKey(projectName);

  if (customer === '' || project === '') {
    return undefined;
  }

  // Items are already newest-first, so the first match is the newest.
  return index.items.find((file: IGeneratedQuotation): boolean =>
    normaliseKey(file.customerName) === customer && normaliseKey(file.projectName) === project);
}

export type ExcelOnlineAction = 'embedview' | 'view' | 'edit' | 'default';

/** Percent-encodes a server-relative URL, keeping the separators literal. */
function encodeServerRelativeUrl(value: string): string {
  return value.split('/').map((segment: string): string => encodeURIComponent(segment)).join('/');
}

/**
 * A URL that renders the workbook with Excel Online, in an iframe
 * ("embedview") or in its own tab ("view" / "edit").
 *
 * Addressed by GUID whenever one is known: sourcedoc as a path cannot resolve a
 * name containing "#" or "%", and generated quote names routinely start with
 * "#". Doc.aspx is the current entry point; WopiFrame.aspx only redirects to it.
 */
export function buildExcelOnlineUrl(
  webAbsoluteUrl: string,
  file: IGeneratedQuotation,
  action: ExcelOnlineAction = 'embedview'
): string {
  const sourceDoc: string = file.uniqueId !== ''
    ? `%7B${file.uniqueId.replace(/[{}]/g, '')}%7D`
    : encodeServerRelativeUrl(file.serverRelativeUrl);

  // Interactivity on: the point is scrolling the sheet and reading every
  // section, so gridlines and headers stay and it reads like the workbook.
  const embedOptions: string = action === 'embedview' ? '&wdAllowInteractivity=True' : '';

  return `${webAbsoluteUrl}/_layouts/15/Doc.aspx?sourcedoc=${sourceDoc}&action=${action}${embedOptions}`;
}
