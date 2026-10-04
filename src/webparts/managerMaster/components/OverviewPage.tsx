import * as React from 'react';
import { MessageBar, MessageBarType, Text } from '@fluentui/react';
import styles from './ManagerMaster.module.scss';
import KpiCard from './KpiCard';
import StageFunnel, { IStageCount } from './StageFunnel';
import { ALL_STATUSES, FINAL_STATUSES, QUOTATION_STATUSES, SALES_STATUSES } from '../../../statuses';
import { IListTabConfig, IListTabData } from '../models/IDashboardModels';
import {
  IQuotationConfig,
  IQuotationItem,
  IQuotationLoad,
  IStatusMismatch,
  STALE_AFTER_DAYS,
  compareStatuses,
  daysSinceModified,
  formatAmount,
  readStatusValue,
  sumAmount
} from '../services/QuotationService';

export interface IOverviewPageProps {
  quotationConfig: IQuotationConfig;
  load: IQuotationLoad;
  tabs: IListTabConfig[];
  listData: Record<string, IListTabData>;
  /** Jumps to the Sales or Quotation board with that stage pre-filtered. */
  onSelectStage: (status: string) => void;
}

const OverviewPage: React.FunctionComponent<IOverviewPageProps> = (props: IOverviewPageProps) => {
  const { quotationConfig, load, tabs, listData, onSelectStage } = props;

  const statusField: string = quotationConfig.statusField.trim();
  const statusOf = (item: IQuotationItem): string =>
    readStatusValue(item, statusField, load.statusLabelById);

  const countIn = (statuses: string[]): number =>
    load.items.filter((item: IQuotationItem): boolean => statuses.indexOf(statusOf(item)) >= 0).length;

  const toStageCounts = (statuses: string[]): IStageCount[] =>
    statuses.map((status: string): IStageCount => ({
      status,
      count: load.items.filter((item: IQuotationItem): boolean => statusOf(item) === status).length
    }));

  const approved: number = load.items.filter((item: IQuotationItem): boolean =>
    statusOf(item) === 'Quotation Approved').length;

  const isOpen = (item: IQuotationItem): boolean => FINAL_STATUSES.indexOf(statusOf(item)) < 0;

  const stalled: number = load.items.filter((item: IQuotationItem): boolean => {
    const days: number | undefined = daysSinceModified(item);
    return isOpen(item) && days !== undefined && days >= STALE_AFTER_DAYS;
  }).length;

  const movedThisWeek: number = load.items.filter((item: IQuotationItem): boolean => {
    const days: number | undefined = daysSinceModified(item);
    return days !== undefined && days <= 7;
  }).length;

  const awaitingDecision: number = load.items.filter((item: IQuotationItem): boolean =>
    ['13. Pending for Approval', '12. Quotation Submitted', '18. Revised Quotation Submitted']
      .indexOf(statusOf(item)) >= 0).length;

  const amountField: string = quotationConfig.amountField.trim();
  const hasAmount: boolean = amountField !== '' && load.fieldTypes[amountField] !== undefined;

  const itemsIn = (statuses: string[]): IQuotationItem[] =>
    load.items.filter((item: IQuotationItem): boolean => statuses.indexOf(statusOf(item)) >= 0);

  const openValue: number = hasAmount
    ? sumAmount(load.items.filter(isOpen), amountField)
    : 0;
  const salesValue: number = hasAmount ? sumAmount(itemsIn(SALES_STATUSES), amountField) : 0;
  const quotationValue: number = hasAmount ? sumAmount(itemsIn(QUOTATION_STATUSES), amountField) : 0;
  const wonValue: number = hasAmount
    ? sumAmount(itemsIn(['19. Quotation Approved']), amountField)
    : 0;

  const mismatch: IStatusMismatch = compareStatuses(load.statusChoices);
  const hasMismatch: boolean = mismatch.missingFromCode.length > 0 || mismatch.missingFromList.length > 0;

  return (
    <div className={styles.page}>
      <header className={styles.pageHead}>
        <Text variant="xLarge" className={styles.pageTitle} block>Overview</Text>
        <Text variant="small" className={styles.pageSubtitle}>
          Live counts across {quotationConfig.listTitle} and the CRM lists on this site.
        </Text>
      </header>

      {load.availableLists !== undefined && load.availableLists.length > 0 && (
        <MessageBar messageBarType={MessageBarType.severeWarning} className={styles.messageBar} isMultiline>
          <strong>&quot;{quotationConfig.listTitle}&quot; was not found on this site.</strong> Lists that do exist
          here: {load.availableLists.join(', ')}. Put the exact title into <em>Quotation pipeline tab &rarr; List
          title</em> in the web part properties, or point the web part at the site that holds it.
        </MessageBar>
      )}

      {hasMismatch && (
        <MessageBar messageBarType={MessageBarType.warning} className={styles.messageBar} isMultiline>
          <strong>
            The status column has {load.statusChoices.length} choices; src/statuses.ts lists {ALL_STATUSES.length}.
          </strong>
          {mismatch.missingFromCode.length > 0 && (
            <div>In SharePoint but not in the code (cannot be ordered): {mismatch.missingFromCode.join(', ')}</div>
          )}
          {mismatch.missingFromList.length > 0 && (
            <div>In the code but not in SharePoint (saving these will fail): {mismatch.missingFromList.join(', ')}</div>
          )}
        </MessageBar>
      )}

      {hasAmount && (
        <>
          <Text variant="mediumPlus" className={styles.sectionLabel} block>Pipeline value</Text>
          <div className={styles.kpiRow}>
            <KpiCard
              label="Open pipeline"
              value={openValue}
              display={formatAmount(openValue)}
              iconName="Money"
              tone="brand"
              hint="Everything not yet closed"
            />
            <KpiCard
              label="In sales stage"
              value={salesValue}
              display={formatAmount(salesValue)}
              iconName="Trending12"
              tone="neutral"
            />
            <KpiCard
              label="In quotation stage"
              value={quotationValue}
              display={formatAmount(quotationValue)}
              iconName="Script"
              tone="neutral"
            />
            <KpiCard
              label="Approved value"
              value={wonValue}
              display={formatAmount(wonValue)}
              iconName="CheckMark"
              tone="success"
            />
          </div>
        </>
      )}

      <Text variant="mediumPlus" className={styles.sectionLabel} block>Needs attention</Text>
      <div className={styles.kpiRow}>
        <KpiCard
          label={`Stalled ${STALE_AFTER_DAYS}+ days`}
          value={stalled}
          iconName="Clock"
          tone={stalled > 0 ? 'danger' : 'success'}
          hint="Open quotations with no recent change"
        />
        <KpiCard
          label="Awaiting client or approval"
          value={awaitingDecision}
          iconName="Hourglass"
          tone={awaitingDecision > 0 ? 'warning' : 'neutral'}
          hint="Submitted, revised or pending approval"
        />
        <KpiCard
          label="Moved in last 7 days"
          value={movedThisWeek}
          iconName="Rocket"
          tone="success"
          hint="Quotations that changed recently"
        />
      </div>

      <Text variant="mediumPlus" className={styles.sectionLabel} block>Pipeline volume</Text>
      <div className={styles.kpiRow}>
        <KpiCard
          label="Quotations"
          value={load.items.length}
          iconName="Script"
          tone="brand"
          hint={quotationConfig.listTitle}
        />
        <KpiCard label="In sales stage" value={countIn(SALES_STATUSES)} iconName="Trending12" tone="neutral" />
        <KpiCard label="In quotation stage" value={countIn(QUOTATION_STATUSES)} iconName="QuickNote" tone="neutral" />
        <KpiCard label="Approved" value={approved} iconName="CheckMark" tone="success" />
        <KpiCard label="Closed" value={countIn(FINAL_STATUSES)} iconName="Archive" tone="warning" />
        {tabs.map((tab: IListTabConfig) => (
          <KpiCard
            key={tab.key}
            label={tab.headerText}
            value={(listData[tab.key]?.items ?? []).length}
            iconName="BulletedList"
            tone="neutral"
            hint={tab.headerText === tab.listTitle ? undefined : tab.listTitle}
          />
        ))}
      </div>

      <div className={styles.funnelGrid}>
        <StageFunnel
          title={`Sales stage (${SALES_STATUSES.length})`}
          stages={toStageCounts(SALES_STATUSES)}
          onSelect={onSelectStage}
        />
        <StageFunnel
          title={`Quotation stage (${QUOTATION_STATUSES.length})`}
          stages={toStageCounts(QUOTATION_STATUSES)}
          onSelect={onSelectStage}
        />
      </div>
    </div>
  );
};

export default OverviewPage;
