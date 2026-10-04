import * as React from 'react';
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  SearchBox,
  SelectionMode,
  Spinner,
  SpinnerSize,
  Stack,
  Text
} from '@fluentui/react';
import styles from './ManagerMaster.module.scss';
import KpiCard from './KpiCard';
import { IDashboardField, IKpiSet, IListTabConfig, IListTabData, ListItem } from '../models/IDashboardModels';
import { formatFieldValue, getDateValue, getNumberValue, getStatusValue } from '../services/fieldValue';

export interface IListTabProps {
  config: IListTabConfig;
  data: IListTabData;
  kpis: IKpiSet;
  loading: boolean;
  probabilityThreshold: number;
  onRefresh: () => void;
}

interface ISortState {
  internalName: string;
  descending: boolean;
}

const ALL_STATUSES: string = '__all__';

/** Sorts on the raw value for dates and numbers, on the rendered text otherwise. */
function compareItems(a: ListItem, b: ListItem, field: IDashboardField): number {
  if (field.typeAsString === 'DateTime') {
    const left: number = getDateValue(a, field.internalName)?.getTime() ?? 0;
    const right: number = getDateValue(b, field.internalName)?.getTime() ?? 0;
    return left - right;
  }

  if (['Number', 'Currency', 'Counter', 'Integer'].indexOf(field.typeAsString) >= 0) {
    return (getNumberValue(a, field.internalName) ?? 0) - (getNumberValue(b, field.internalName) ?? 0);
  }

  return formatFieldValue(field, a).localeCompare(formatFieldValue(field, b));
}

const ListTab: React.FunctionComponent<IListTabProps> = (props: IListTabProps) => {
  const { config, data, kpis, loading, probabilityThreshold, onRefresh } = props;

  const [search, setSearch] = React.useState<string>('');
  const [status, setStatus] = React.useState<string>(ALL_STATUSES);
  const [sort, setSort] = React.useState<ISortState | undefined>(undefined);
  const [fieldsOpen, setFieldsOpen] = React.useState<boolean>(false);

  const hasStatusField: boolean = config.statusField.trim() !== '';

  const statusOptions: IDropdownOption[] = React.useMemo(() => {
    if (!hasStatusField) {
      return [];
    }

    const seen: Set<string> = new Set<string>();
    for (const item of data.items) {
      const value: string = getStatusValue(item, config.statusField.trim());
      if (value !== '') {
        seen.add(value);
      }
    }

    return [
      { key: ALL_STATUSES, text: 'All statuses' },
      ...Array.from(seen).sort().map((value: string): IDropdownOption => ({ key: value, text: value }))
    ];
  }, [data.items, config.statusField, hasStatusField]);

  const visibleItems: ListItem[] = React.useMemo(() => {
    const needle: string = search.trim().toLowerCase();

    let rows: ListItem[] = data.items;

    if (status !== ALL_STATUSES) {
      rows = rows.filter((item: ListItem): boolean => getStatusValue(item, config.statusField.trim()) === status);
    }

    if (needle !== '') {
      rows = rows.filter((item: ListItem): boolean =>
        data.columns.some((field: IDashboardField): boolean =>
          formatFieldValue(field, item).toLowerCase().indexOf(needle) >= 0));
    }

    if (sort !== undefined) {
      const field: IDashboardField | undefined = data.columns.filter(
        (f: IDashboardField): boolean => f.internalName === sort.internalName
      )[0];

      if (field !== undefined) {
        rows = [...rows].sort((a: ListItem, b: ListItem): number => {
          const result: number = compareItems(a, b, field);
          return sort.descending ? -result : result;
        });
      }
    }

    return rows;
  }, [data.items, data.columns, search, status, sort, config.statusField]);

  const onColumnClick = (_ev?: unknown, column?: IColumn): void => {
    if (column === undefined) {
      return;
    }
    setSort((current: ISortState | undefined): ISortState => ({
      internalName: column.key,
      descending: current !== undefined && current.internalName === column.key ? !current.descending : false
    }));
  };

  const columns: IColumn[] = data.columns.map((field: IDashboardField): IColumn => ({
    key: field.internalName,
    name: field.title,
    fieldName: field.internalName,
    minWidth: 90,
    maxWidth: field.typeAsString === 'Note' ? 320 : 200,
    isResizable: true,
    isSorted: sort !== undefined && sort.internalName === field.internalName,
    isSortedDescending: sort !== undefined && sort.internalName === field.internalName && sort.descending,
    onColumnClick: onColumnClick,
    onRender: (item: ListItem): string => formatFieldValue(field, item)
  }));

  if (loading) {
    return <Spinner size={SpinnerSize.large} label={`Loading ${config.headerText}...`} className={styles.spinner} />;
  }

  return (
    <div className={styles.tab}>
      {data.error !== undefined && (
        <MessageBar
          messageBarType={data.items.length > 0 ? MessageBarType.warning : MessageBarType.error}
          className={styles.messageBar}
        >
          {data.error}
        </MessageBar>
      )}

      <div className={styles.kpiRow}>
        <KpiCard label={`Total ${config.headerText}`} value={kpis.total} iconName="NumberSymbol" tone="neutral" />
        {hasStatusField && (
          <KpiCard label="Open" value={kpis.open} iconName="OpenFolderHorizontal" tone="success" hint={config.statusField} />
        )}
        {config.dateField.trim() !== '' && (
          <>
            <KpiCard label="Overdue" value={kpis.overdue} iconName="Warning" tone="danger" hint={config.dateField} />
            <KpiCard label="Due in 7 days" value={kpis.dueThisWeek} iconName="Clock" tone="warning" hint={config.dateField} />
          </>
        )}
        {kpis.highProbability !== undefined && (
          <KpiCard
            label={`Probability >= ${probabilityThreshold}%`}
            value={kpis.highProbability}
            iconName="Trending12"
            tone="success"
            hint={config.probabilityField}
          />
        )}
      </div>

      <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end" className={styles.toolbar}>
        <SearchBox
          placeholder={`Search ${config.headerText}`}
          value={search}
          onChange={(_ev, newValue?: string): void => setSearch(newValue ?? '')}
          onClear={(): void => setSearch('')}
          className={styles.searchBox}
        />
        {statusOptions.length > 1 && (
          <Dropdown
            selectedKey={status}
            options={statusOptions}
            onChange={(_ev, option?: IDropdownOption): void => setStatus(String(option?.key ?? ALL_STATUSES))}
            className={styles.statusFilter}
          />
        )}
        <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={onRefresh} />
        <DefaultButton text="Fields" iconProps={{ iconName: 'Info' }} onClick={(): void => setFieldsOpen(true)} />
        <Text variant="small" className={styles.resultCount}>
          {visibleItems.length === data.items.length
            ? `${data.items.length} item(s)`
            : `${visibleItems.length} of ${data.items.length} item(s)`}
        </Text>
      </Stack>

      {data.columns.length === 0 ? (
        <MessageBar messageBarType={MessageBarType.info}>
          No columns to show yet. Open <strong>Fields</strong> to see the internal names on this list, then set them in the
          web part properties.
        </MessageBar>
      ) : (
        <DetailsList
          items={visibleItems}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          getKey={(item: ListItem): string => String(item.Id ?? item.ID ?? '')}
          setKey={config.key}
          isHeaderVisible={true}
        />
      )}

      <Panel
        isOpen={fieldsOpen}
        onDismiss={(): void => setFieldsOpen(false)}
        type={PanelType.medium}
        headerText={`Fields on "${config.listTitle}"`}
        closeButtonAriaLabel="Close"
      >
        <Text variant="small">
          Paste the internal names you want, comma separated, into the Columns box for this tab in the web part properties.
        </Text>
        <table className={styles.fieldTable}>
          <thead>
            <tr>
              <th>Display name</th>
              <th>Internal name</th>
              <th>Type</th>
            </tr>
          </thead>
          <tbody>
            {data.allFields.map((field: IDashboardField) => (
              <tr key={field.internalName}>
                <td>{field.title}</td>
                <td><code>{field.internalName}</code></td>
                <td>{field.typeAsString}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Panel>
    </div>
  );
};

export default ListTab;
