/** A SharePoint list item as returned by the REST API. */
export type ListItem = Record<string, unknown>;

/** The subset of SharePoint field metadata the dashboard needs. */
export interface IDashboardField {
  internalName: string;
  title: string;
  typeAsString: string;
}

/** Per-tab configuration, all of it editable from the property pane. */
export interface IListTabConfig {
  key: string;
  headerText: string;
  listTitle: string;
  /** Comma separated internal names. Empty means "pick the first few automatically". */
  columns: string;
  /** Date field driving the Overdue / Due this week KPIs. */
  dateField: string;
  /** Choice or text field driving the status filter and the open/closed split. */
  statusField: string;
  /** Numeric field driving the High probability KPI. Only used on the leads tab. */
  probabilityField: string;
}

/** Everything one tab needs to render, or the reason it could not be loaded. */
export interface IListTabData {
  /** Columns actually shown in the grid. */
  columns: IDashboardField[];
  /** Every usable field on the list, for the "Fields" diagnostics panel. */
  allFields: IDashboardField[];
  items: ListItem[];
  error?: string;
}

export interface IKpiSet {
  total: number;
  open: number;
  overdue: number;
  dueThisWeek: number;
  /** Undefined when no probability field is configured for the tab. */
  highProbability?: number;
}
