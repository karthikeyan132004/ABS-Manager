import { IDashboardField, ListItem } from '../models/IDashboardModels';

const EXPANDABLE_TYPES: Set<string> = new Set(['Lookup', 'LookupMulti', 'User', 'UserMulti']);
const DATE_TYPES: Set<string> = new Set(['DateTime']);
const NUMBER_TYPES: Set<string> = new Set(['Number', 'Currency', 'Counter', 'Integer']);

/** True when the field has to be $expand-ed to get a display value back. */
export function isExpandable(field: IDashboardField): boolean {
  return EXPANDABLE_TYPES.has(field.typeAsString);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

/**
 * Lookup and person fields come back either as an array (multi) or as an object
 * (single) once expanded, so both shapes are flattened to a display string.
 */
function formatExpanded(value: unknown): string {
  const entries: unknown[] = Array.isArray(value)
    ? value
    : isRecord(value) && Array.isArray(value.results)
      ? value.results
      : [value];

  return entries
    .map((entry: unknown): string => (isRecord(entry) ? String(entry.Title ?? entry.Id ?? '') : ''))
    .filter((text: string): boolean => text.length > 0)
    .join(', ');
}

function stripHtml(value: string): string {
  return value.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
}

/** Renders any supported field value as plain text for the grid and for search. */
export function formatFieldValue(field: IDashboardField, item: ListItem): string {
  const value: unknown = item[field.internalName];

  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (isExpandable(field)) {
    return formatExpanded(value);
  }

  if (DATE_TYPES.has(field.typeAsString)) {
    const date: Date = new Date(String(value));
    return isNaN(date.getTime()) ? '' : date.toLocaleDateString();
  }

  if (field.typeAsString === 'Boolean') {
    return value === true ? 'Yes' : 'No';
  }

  if (NUMBER_TYPES.has(field.typeAsString)) {
    const num: number = Number(value);
    return isNaN(num) ? '' : num.toLocaleString();
  }

  if (field.typeAsString === 'URL' && isRecord(value)) {
    return String(value.Description ?? value.Url ?? '');
  }

  if (field.typeAsString === 'TaxonomyFieldType' && isRecord(value)) {
    return String(value.Label ?? '');
  }

  if (Array.isArray(value)) {
    return value.map((entry: unknown): string => String(entry)).join(', ');
  }

  if (isRecord(value)) {
    return formatExpanded(value);
  }

  const text: string = String(value);
  return field.typeAsString === 'Note' ? stripHtml(text) : text;
}

/** Parses a field value as a date, for the overdue / due-soon calculations. */
export function getDateValue(item: ListItem, internalName: string): Date | undefined {
  if (internalName === '') {
    return undefined;
  }

  const raw: unknown = item[internalName];
  if (raw === null || raw === undefined || raw === '') {
    return undefined;
  }

  const date: Date = new Date(String(raw));
  return isNaN(date.getTime()) ? undefined : date;
}

/** Parses a field value as a number, for the probability KPI. */
export function getNumberValue(item: ListItem, internalName: string): number | undefined {
  if (internalName === '') {
    return undefined;
  }

  const raw: unknown = item[internalName];
  if (raw === null || raw === undefined || raw === '') {
    return undefined;
  }

  const num: number = Number(raw);
  return isNaN(num) ? undefined : num;
}

const CLOSED_STATUS: RegExp = /closed|won|lost|complete|completed|cancel|cancelled|done|archive/i;

/** Reads the configured status field as text. Empty when not configured or not set. */
export function getStatusValue(item: ListItem, internalName: string): string {
  if (internalName === '') {
    return '';
  }

  const raw: unknown = item[internalName];
  if (raw === null || raw === undefined) {
    return '';
  }

  if (Array.isArray(raw)) {
    return raw.map((entry: unknown): string => String(entry)).join(', ');
  }

  if (isRecord(raw)) {
    return String(raw.Title ?? raw.Label ?? '');
  }

  return String(raw);
}

/**
 * Treats a record as closed when its status reads like a terminal state, so
 * finished work stops counting towards Overdue.
 */
export function isClosed(item: ListItem, statusFieldInternalName: string): boolean {
  const status: string = getStatusValue(item, statusFieldInternalName);
  return status !== '' && CLOSED_STATUS.test(status);
}
