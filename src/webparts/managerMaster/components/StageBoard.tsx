import * as React from 'react';
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  IContextualMenuItem,
  IDropdownOption,
  Link,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  PrimaryButton,
  SearchBox,
  SelectionMode,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField,
  Toggle
} from '@fluentui/react';
import styles from './ManagerMaster.module.scss';
import KpiCard from './KpiCard';
import StageChips, { IStageChip } from './StageChips';
import { ALL_STATUSES, FINAL_STATUSES, LOCK_FINAL_STATUSES, displayStatus, isLocked, stepStatus } from '../../../statuses';
import {
  ILeadSyncConfig,
  IQuotationConfig,
  IQuotationItem,
  IQuotationLoad,
  buildStatusUpdate,
  coerceFieldValue,
  isInvalidNumber,
  readStatusValue,
  updateQuotation,
  STALE_AFTER_DAYS,
  LeadSyncOutcome,
  daysSinceModified,
  formatAmount,
  getLinkedLeadId,
  isLeadSyncConfigured,
  sumAmount,
  syncLeadStage
} from '../services/QuotationService';

export interface IStageBoardProps {
  config: IQuotationConfig;
  leadSync: ILeadSyncConfig;
  load: IQuotationLoad;
  /** The statuses this board owns, in pipeline order. */
  stages: string[];
  heading: string;
  description: string;
  onReload: () => Promise<void>;
  /** A status to pre-filter on, set when arriving from the Overview funnel. */
  focusStatus?: string;
  onFocusHandled?: () => void;
}

interface IDraft {
  status: string;
  amount: string;
  remarks: string;
}

const EMPTY_DRAFT: IDraft = { status: '', amount: '', remarks: '' };

const StageBoard: React.FunctionComponent<IStageBoardProps> = (props: IStageBoardProps) => {
  const { config, leadSync, load, stages, heading, description, onReload, focusStatus, onFocusHandled } = props;

  const syncEnabled: boolean = isLeadSyncConfigured(leadSync);

  const [filter, setFilter] = React.useState<string | undefined>(undefined);
  const [search, setSearch] = React.useState<string>('');
  const [showAllStages, setShowAllStages] = React.useState<boolean>(false);
  const [showStaleOnly, setShowStaleOnly] = React.useState<boolean>(false);
  const [selected, setSelected] = React.useState<IQuotationItem | undefined>(undefined);
  const [draft, setDraft] = React.useState<IDraft>(EMPTY_DRAFT);
  const [busyId, setBusyId] = React.useState<number | undefined>(undefined);
  const [message, setMessage] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');

  const statusField: string = config.statusField.trim();
  const amountField: string = config.amountField.trim();
  const remarksField: string = config.remarksField.trim();

  // A name that is not on the list is treated as unset, so a stale or mistyped
  // setting hides the field instead of showing a permanently empty column.
  const hasAmount: boolean = amountField !== '' && load.fieldTypes[amountField] !== undefined;
  const hasRemarks: boolean = remarksField !== '' && load.fieldTypes[remarksField] !== undefined;

  const mounted: React.MutableRefObject<boolean> = React.useRef<boolean>(true);
  React.useEffect(() => {
    return (): void => {
      mounted.current = false;
    };
  }, []);

  // Arriving from the Overview funnel pre-selects that stage, once.
  React.useEffect(() => {
    if (focusStatus !== undefined) {
      setFilter(focusStatus);
      if (onFocusHandled !== undefined) {
        onFocusHandled();
      }
    }
  }, [focusStatus, onFocusHandled]);

  const flash = (text: string): void => {
    setMessage(text);
    setTimeout((): void => {
      if (mounted.current) {
        setMessage('');
      }
    }, 3000);
  };

  const statusOf = (item: IQuotationItem): string =>
    readStatusValue(item, statusField, load.statusLabelById);

  const applyStatus = async (item: IQuotationItem, newStatus: string): Promise<void> => {
    const payload: Record<string, unknown> | undefined = buildStatusUpdate(config, load, newStatus);
    if (payload === undefined) {
      setError(`"${newStatus}" is not one of the values the status column accepts, so it cannot be saved.`);
      return;
    }

    setError('');
    setBusyId(item.Id);
    try {
      await updateQuotation(config, item.Id, payload);

      // The quotation is already saved at this point, so a sync problem is
      // reported as a note on a successful move, never as a failed move.
      let note: string = '';
      if (syncEnabled) {
        try {
          const outcome: LeadSyncOutcome = await syncLeadStage(item, newStatus, leadSync, load);
          note = outcome.kind === 'synced'
            ? ` and updated lead ${outcome.leadLabel}`
            : ` (lead not updated: ${outcome.reason})`;
        } catch (syncError) {
          note = ` (lead not updated: ${syncError instanceof Error ? syncError.message : String(syncError)})`;
        }
      }

      await onReload();
      flash(`Moved ${item.Title ?? item.Id} to "${displayStatus(newStatus)}"${note}`);
    } catch (e) {
      setError(`Could not update ${item.Title ?? item.Id}: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      if (mounted.current) {
        setBusyId(undefined);
      }
    }
  };

  const move = (item: IQuotationItem, step: 1 | -1): void => {
    const current: string = statusOf(item);
    const next: string | undefined = stepStatus(current, step);

    if (next === undefined) {
      setError(
        current === ''
          ? `${item.Title ?? item.Id} has no status set, so it cannot be moved. Use "Jump to..." to set one.`
          : `"${current}" is not in the pipeline in src/statuses.ts, so there is no ${step === 1 ? 'next' : 'previous'} stage. Fix the spelling there or use "Jump to...".`
      );
      return;
    }

    applyStatus(item, next).catch((e: unknown): void => setError(String(e)));
  };

  const openItem = (item: IQuotationItem): void => {
    setSelected(item);
    setDraft({
      status: statusOf(item),
      amount: hasAmount ? String(item[amountField] ?? '') : '',
      remarks: hasRemarks ? String(item[remarksField] ?? '') : ''
    });
  };

  const amountInvalid: boolean = hasAmount && isInvalidNumber(draft.amount, amountField, load.fieldTypes);

  const save = async (): Promise<void> => {
    if (selected === undefined || amountInvalid) {
      return;
    }

    const statusPayload: Record<string, unknown> | undefined = buildStatusUpdate(config, load, draft.status);
    if (statusPayload === undefined) {
      setError(`"${draft.status}" is not one of the values the status column accepts, so it cannot be saved.`);
      return;
    }

    const fields: Record<string, unknown> = { ...statusPayload };
    if (hasAmount) {
      fields[amountField] = coerceFieldValue(draft.amount, amountField, load.fieldTypes);
    }
    if (hasRemarks) {
      fields[remarksField] = draft.remarks;
    }

    setError('');
    setBusyId(selected.Id);
    try {
      await updateQuotation(config, selected.Id, fields);
      await onReload();
      if (mounted.current) {
        setSelected(undefined);
      }
      flash(`Saved ${selected.Title ?? selected.Id}`);
    } catch (e) {
      setError(`Save failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      if (mounted.current) {
        setBusyId(undefined);
      }
    }
  };

  const countFor = (status: string): number =>
    load.items.filter((item: IQuotationItem): boolean => statusOf(item) === status).length;

  const stageItems: IQuotationItem[] = React.useMemo(() => {
    if (showAllStages) {
      return load.items;
    }
    return load.items.filter((item: IQuotationItem): boolean =>
      stages.indexOf(readStatusValue(item, statusField, load.statusLabelById)) >= 0);
  }, [load.items, stages, statusField, showAllStages]);

  const chips: IStageChip[] = stages.map((status: string): IStageChip => ({
    status,
    count: countFor(status)
  }));

  const busiest: IStageChip = chips.reduce(
    (best: IStageChip, chip: IStageChip): IStageChip => (chip.count > best.count ? chip : best),
    { status: '', count: 0 }
  );

  const isStale = (item: IQuotationItem): boolean => {
    // Approved, Cancelled and Lost are finished, not stalled, so they never
    // carry the idle warning however long ago they were last touched.
    if (FINAL_STATUSES.indexOf(statusOf(item)) >= 0) {
      return false;
    }
    const days: number | undefined = daysSinceModified(item);
    return days !== undefined && days >= STALE_AFTER_DAYS;
  };

  const staleCount: number = stageItems.filter(isStale).length;

  const visibleItems: IQuotationItem[] = React.useMemo(() => {
    let rows: IQuotationItem[] = stageItems;

    if (showStaleOnly) {
      rows = rows.filter(isStale);
    }

    if (filter !== undefined) {
      rows = rows.filter((item: IQuotationItem): boolean =>
        readStatusValue(item, statusField, load.statusLabelById) === filter);
    }

    const needle: string = search.trim().toLowerCase();
    if (needle !== '') {
      rows = rows.filter((item: IQuotationItem): boolean =>
        [item.Title, readStatusValue(item, statusField, load.statusLabelById), hasAmount ? item[amountField] : '', hasRemarks ? item[remarksField] : '']
          .map((value: unknown): string => String(value ?? '').toLowerCase())
          .some((text: string): boolean => text.indexOf(needle) >= 0));
    }

    return rows;
  }, [stageItems, filter, search, statusField, amountField, remarksField, showStaleOnly, hasAmount, hasRemarks]);

  const allColumns: IColumn[] = [
    {
      key: 'title',
      name: 'Quotation ID',
      fieldName: 'Title',
      minWidth: 110,
      maxWidth: 160,
      isResizable: true,
      onRender: (item: IQuotationItem): React.ReactNode => (
        <Link onClick={(): void => openItem(item)}>{item.Title ?? `Item ${item.Id}`}</Link>
      )
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: statusField,
      minWidth: 230,
      maxWidth: 330,
      isResizable: true,
      onRender: (item: IQuotationItem): React.ReactNode => {
        const status: string = statusOf(item);
        const position: number = ALL_STATUSES.indexOf(status) + 1;
        const days: number | undefined = daysSinceModified(item);
        const stale: boolean = isStale(item);

        return (
          <div className={styles.statusCell}>
            <span className={styles.statusPill}>{status === '' ? 'No status' : displayStatus(status)}</span>
            {position > 0 && (
              <span className={styles.progressWrap} title={`Stage ${position} of ${ALL_STATUSES.length}`}>
                <span className={styles.progressTrack}>
                  <span
                    className={styles.progressBar}
                    style={{ width: `${Math.round((position / ALL_STATUSES.length) * 100)}%` }}
                  />
                </span>
                <span className={styles.progressText}>{position}/{ALL_STATUSES.length}</span>
              </span>
            )}
            {stale && (
              <span className={styles.staleTag} title={`Last changed ${days} days ago`}>
                {days}d idle
              </span>
            )}
          </div>
        );
      }
    },
    {
      key: 'amount',
      name: 'Amount',
      fieldName: amountField,
      minWidth: 90,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: IQuotationItem): React.ReactNode => {
        const raw: unknown = item[amountField];
        const numeric: number = Number(raw);
        return raw === null || raw === undefined || raw === '' || isNaN(numeric)
          ? String(raw ?? '')
          : numeric.toLocaleString();
      }
    },
    {
      key: 'lead',
      name: 'Sales Lead',
      minWidth: 110,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IQuotationItem): React.ReactNode => {
        const leadId: number | undefined = getLinkedLeadId(item, leadSync);
        return leadId === undefined
          ? <span className={styles.unlinkedTag}>Not linked</span>
          : (load.leadLabelById[leadId] ?? `Lead ${leadId}`);
      }
    },
    {
      key: 'actions',
      name: 'Change stage',
      minWidth: 190,
      maxWidth: 220,
      onRender: (item: IQuotationItem): React.ReactNode => {
        const locked: boolean = isLocked(statusOf(item));
        const busy: boolean = busyId === item.Id;
        const disabled: boolean = locked || busy;

        return (
          <Stack
            horizontal
            tokens={{ childrenGap: 4 }}
            verticalAlign="center"
            onClick={(ev: React.MouseEvent<HTMLElement>): void => ev.stopPropagation()}
          >
            <DefaultButton
              text="&#9664;"
              title="Previous stage"
              ariaLabel="Previous stage"
              disabled={disabled}
              onClick={(): void => move(item, -1)}
            />
            <PrimaryButton text="Next stage &#9654;" disabled={disabled} onClick={(): void => move(item, 1)} />
            <DefaultButton
              iconProps={{ iconName: 'MoreVertical' }}
              title="Jump to any stage"
              ariaLabel="Jump to any stage"
              className={styles.jumpButton}
              disabled={disabled}
              menuProps={{
                items: ALL_STATUSES.map((status: string): IContextualMenuItem => ({
                  key: status,
                  text: displayStatus(status),
                  canCheck: true,
                  checked: statusOf(item) === status,
                  onClick: (): void => {
                    applyStatus(item, status).catch((e: unknown): void => setError(String(e)));
                  }
                }))
              }}
            />
            {busy && <Spinner size={SpinnerSize.xSmall} />}
          </Stack>
        );
      }
    }
  ];

  const columns: IColumn[] = allColumns.filter((column: IColumn): boolean => {
    if (column.key === 'amount') {
      return hasAmount;
    }
    if (column.key === 'lead') {
      return syncEnabled;
    }
    return true;
  });

  const statusOptions: IDropdownOption[] = React.useMemo(() => {
    const known: string[] = draft.status !== '' && ALL_STATUSES.indexOf(draft.status) < 0
      ? [draft.status, ...ALL_STATUSES] // keep an unrecognised current value visible
      : ALL_STATUSES;

    return known.map((status: string): IDropdownOption => ({ key: status, text: displayStatus(status) }));
  }, [draft.status]);

  return (
    <div className={styles.page}>
      <header className={styles.pageHead}>
        <Text variant="xLarge" className={styles.pageTitle} block>{heading}</Text>
        <Text variant="small" className={styles.pageSubtitle}>{description}</Text>
      </header>

      {message !== '' && (
        <MessageBar messageBarType={MessageBarType.success} className={styles.messageBar}>{message}</MessageBar>
      )}

      {error !== '' && (
        <MessageBar
          messageBarType={MessageBarType.error}
          className={styles.messageBar}
          onDismiss={(): void => setError('')}
          isMultiline
        >
          {error}
        </MessageBar>
      )}

      <div className={styles.kpiRow}>
        <KpiCard
          label="In this stage group"
          value={stageItems.length}
          iconName="BulletedList"
          tone="brand"
        />
        <KpiCard
          label={`No movement in ${STALE_AFTER_DAYS}+ days`}
          value={staleCount}
          iconName="Clock"
          tone={staleCount > 0 ? 'warning' : 'neutral'}
          onClick={(): void => setShowStaleOnly(!showStaleOnly)}
          active={showStaleOnly}
          hint="Click to filter"
        />
        <KpiCard
          label="Busiest stage"
          value={busiest.count}
          iconName="Trending12"
          tone="neutral"
          hint={busiest.count > 0 ? displayStatus(busiest.status) : 'Nothing here yet'}
        />
      </div>

      <StageChips
        chips={chips}
        selected={filter}
        onSelect={(status: string | undefined): void => setFilter(status)}
      />

      <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="center" className={styles.toolbar}>
        <SearchBox
          placeholder="Search quotations"
          value={search}
          onChange={(_ev, newValue?: string): void => setSearch(newValue ?? '')}
          onClear={(): void => setSearch('')}
          className={styles.searchBox}
        />
        {filter !== undefined && (
          <DefaultButton
            text={`Clear: ${displayStatus(filter)}`}
            iconProps={{ iconName: 'ClearFilter' }}
            onClick={(): void => setFilter(undefined)}
          />
        )}
        <Toggle
          label="All stages"
          inlineLabel
          checked={showAllStages}
          onChange={(_ev, checked?: boolean): void => setShowAllStages(checked === true)}
          className={styles.inlineToggle}
        />
        <DefaultButton
          text="Refresh"
          iconProps={{ iconName: 'Refresh' }}
          onClick={(): void => {
            onReload().catch((e: unknown): void => setError(String(e)));
          }}
        />
        <Text variant="small" className={styles.resultCount}>
          {visibleItems.length === stageItems.length
            ? `${stageItems.length} quotation(s)`
            : `${visibleItems.length} of ${stageItems.length} quotation(s)`}
          {hasAmount && ` - ${formatAmount(sumAmount(visibleItems, amountField))}`}
        </Text>
      </Stack>

      <div className={styles.panelCard}>
        <DetailsList
          items={visibleItems}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          getKey={(item: IQuotationItem): string => String(item.Id)}
          setKey={heading}
          className={styles.grid}
          onRenderRow={(rowProps, defaultRender): JSX.Element | null => {
            if (rowProps === undefined || defaultRender === undefined) {
              return null;
            }
            // The whole row opens the panel; the stage buttons stop the bubble.
            return (
              <div
                className={styles.gridRow}
                onClick={(): void => openItem(rowProps.item as IQuotationItem)}
              >
                {defaultRender(rowProps)}
              </div>
            );
          }}
        />
        {visibleItems.length === 0 && (
          <div className={styles.emptyState}>
            <Text variant="mediumPlus" block className={styles.emptyTitle}>Nothing here right now</Text>
            <Text variant="small" block className={styles.emptyNote}>
              {filter !== undefined
                ? `No quotations sit in "${displayStatus(filter)}".`
                : showStaleOnly
                  ? `Nothing has been idle for ${STALE_AFTER_DAYS}+ days. That is good news.`
                  : search.trim() !== ''
                    ? `Nothing matches "${search.trim()}".`
                    : 'Turn on "All stages" to see quotations from every status.'}
            </Text>
          </div>
        )}
      </div>

      <Panel
        isOpen={selected !== undefined}
        onDismiss={(): void => setSelected(undefined)}
        type={PanelType.medium}
        headerText={selected !== undefined ? `Quotation ${selected.Title ?? selected.Id}` : ''}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={(): React.ReactElement => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Save"
              disabled={amountInvalid || busyId !== undefined}
              onClick={(): void => {
                save().catch((e: unknown): void => setError(String(e)));
              }}
            />
            <DefaultButton text="Cancel" onClick={(): void => setSelected(undefined)} />
            {busyId !== undefined && <Spinner size={SpinnerSize.small} />}
          </Stack>
        )}
        isFooterAtBottom
      >
        {selected !== undefined && isLocked(statusOf(selected)) && (
          <MessageBar messageBarType={MessageBarType.info} className={styles.messageBar}>
            This quotation is in a final status and locking is switched on, so the status cannot be changed.
          </MessageBar>
        )}

        <Dropdown
          label="Status"
          selectedKey={draft.status !== '' ? draft.status : undefined}
          placeholder="Choose a status"
          options={statusOptions}
          disabled={selected !== undefined && isLocked(statusOf(selected))}
          onChange={(_ev, option?: IDropdownOption): void => {
            if (option !== undefined) {
              setDraft({ ...draft, status: String(option.key) });
            }
          }}
        />

        {hasAmount && (
          <TextField
            label="Amount"
            value={draft.amount}
            errorMessage={amountInvalid ? 'Enter a number.' : undefined}
            onChange={(_ev, newValue?: string): void => setDraft({ ...draft, amount: newValue ?? '' })}
          />
        )}

        {hasRemarks && (
          <TextField
            label="Remarks"
            multiline
            rows={4}
            value={draft.remarks}
            onChange={(_ev, newValue?: string): void => setDraft({ ...draft, remarks: newValue ?? '' })}
          />
        )}

        {!LOCK_FINAL_STATUSES && (
          <Text variant="small" className={styles.panelHint} block>
            Every status is editable. To lock Approved / Canclled / Lost later, set LOCK_FINAL_STATUSES to true in
            src/statuses.ts.
          </Text>
        )}
      </Panel>
    </div>
  );
};

export default StageBoard;
