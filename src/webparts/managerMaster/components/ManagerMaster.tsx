import * as React from 'react';
import { MessageBar, MessageBarType, Persona, PersonaSize, Spinner, SpinnerSize, Text } from '@fluentui/react';
import styles from './ManagerMaster.module.scss';
import {
  EMPTY_GENERATED_INDEX,
  IGeneratedQuotationIndex,
  loadGeneratedQuotations
} from '../services/GeneratedQuotationService';
import type { IManagerMasterProps } from './IManagerMasterProps';
import SideNav, { INavItem } from './SideNav';
import OverviewPage from './OverviewPage';
import StageBoard from './StageBoard';
import ListTab from './ListTab';
import { QUOTATION_STATUSES, SALES_STATUSES } from '../../../statuses';
import { IKpiSet, IListTabConfig, IListTabData } from '../models/IDashboardModels';
import { calculateKpis, loadListTab } from '../services/DashboardService';
import {
  EMPTY_QUOTATION_LOAD,
  ILeadSyncConfig,
  IQuotationConfig,
  IQuotationItem,
  IQuotationLoad,
  loadQuotations,
  readStatusValue
} from '../services/QuotationService';

const OVERVIEW_KEY: string = 'overview';
const SALES_KEY: string = 'sales';
const QUOTATION_KEY: string = 'quotation';

const EMPTY_TAB: IListTabData = { columns: [], allFields: [], items: [] };

const ManagerMaster: React.FunctionComponent<IManagerMasterProps> = (props: IManagerMasterProps) => {
  const {
    context, title, environmentMessage, userDisplayName, tabs, itemLimit, probabilityThreshold,
    quotation, quotationPageUrl, leadSync
  } = props;

  const [listData, setListData] = React.useState<Record<string, IListTabData>>({});
  const [quotationLoad, setQuotationLoad] = React.useState<IQuotationLoad>(EMPTY_QUOTATION_LOAD);
  const [generated, setGenerated] = React.useState<IGeneratedQuotationIndex>(EMPTY_GENERATED_INDEX);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [selectedKey, setSelectedKey] = React.useState<string>(OVERVIEW_KEY);
  const [focusStatus, setFocusStatus] = React.useState<string | undefined>(undefined);
  const [collapsed, setCollapsed] = React.useState<boolean>(false);

  const hasQuotationList: boolean = quotation.listTitle.trim() !== '';
  const statusField: string = quotation.statusField.trim();

  // Config comes from the property pane, so compare by value rather than identity.
  const tabSignature: string = JSON.stringify(tabs);
  const quotationSignature: string = JSON.stringify(quotation);
  const leadSyncSignature: string = JSON.stringify(leadSync);

  const mounted: React.MutableRefObject<boolean> = React.useRef<boolean>(true);
  React.useEffect(() => {
    return (): void => {
      mounted.current = false;
    };
  }, []);

  const reload = React.useCallback(async (): Promise<void> => {
    const configuredTabs: IListTabConfig[] = JSON.parse(tabSignature) as IListTabConfig[];
    const quotationConfig: IQuotationConfig = JSON.parse(quotationSignature) as IQuotationConfig;
    const syncConfig: ILeadSyncConfig = JSON.parse(leadSyncSignature) as ILeadSyncConfig;

    const [quotationResult, generatedResult, listResults] = await Promise.all([
      quotationConfig.listTitle.trim() === ''
        ? Promise.resolve(EMPTY_QUOTATION_LOAD)
        : loadQuotations(quotationConfig, itemLimit, syncConfig),
      // The workbooks absquot generates. A missing library is reported on the
      // result, so the dashboard still loads without one.
      loadGeneratedQuotations(itemLimit),
      Promise.all(
        configuredTabs.map((tab: IListTabConfig): Promise<[string, IListTabData]> =>
          loadListTab(tab, itemLimit).then((result: IListTabData): [string, IListTabData] => [tab.key, result]))
      )
    ]);

    if (!mounted.current) {
      return;
    }

    const nextListData: Record<string, IListTabData> = {};
    for (const [key, value] of listResults) {
      nextListData[key] = value;
    }

    setQuotationLoad(quotationResult);
    setGenerated(generatedResult);
    setListData(nextListData);
  }, [tabSignature, quotationSignature, leadSyncSignature, itemLimit]);

  React.useEffect(() => {
    setLoading(true);
    reload()
      .catch((e: unknown): void => {
        console.error('Manager dashboard failed to load', e);
      })
      .then((): void => {
        if (mounted.current) {
          setLoading(false);
        }
      })
      .catch((): void => undefined);
  }, [reload]);

  const kpisByTab: Record<string, IKpiSet> = {};
  for (const tab of tabs) {
    kpisByTab[tab.key] = calculateKpis(listData[tab.key] ?? EMPTY_TAB, tab, probabilityThreshold);
  }

  const countInStages = (stages: string[]): number =>
    quotationLoad.items.filter((item: IQuotationItem): boolean =>
      stages.indexOf(readStatusValue(item, statusField, quotationLoad.statusLabelById)) >= 0).length;

  /** Nav order, as agreed: Overview, Sales, Customer Quotation, then the lists. */
  const navItems: INavItem[] = [
    { key: OVERVIEW_KEY, label: 'Overview', icon: 'ViewDashboard', group: 'Dashboard' }
  ];

  if (hasQuotationList) {
    navItems.push(
      { key: SALES_KEY, label: 'Sales', icon: 'Trending12', group: 'Pipeline', badge: countInStages(SALES_STATUSES) },
      {
        key: QUOTATION_KEY,
        label: 'Customer Quotation',
        icon: 'Script',
        group: 'Pipeline',
        badge: countInStages(QUOTATION_STATUSES)
      }
    );
  }

  for (const tab of tabs) {
    navItems.push({
      key: tab.key,
      label: tab.headerText,
      icon: tab.key === 'leads' ? 'ContactList' : tab.key === 'activity' ? 'TaskManager' : 'ProjectCollection',
      group: 'Lists',
      badge: kpisByTab[tab.key].total
    });
  }

  /** Sends the manager from an Overview funnel bar to the board that owns it. */
  const onSelectStage = (status: string): void => {
    setSelectedKey(SALES_STATUSES.indexOf(status) >= 0 ? SALES_KEY : QUOTATION_KEY);
    setFocusStatus(status);
  };

  const activeTab: IListTabConfig | undefined = tabs.filter(
    (tab: IListTabConfig): boolean => tab.key === selectedKey
  )[0];

  const renderContent = (): React.ReactNode => {
    if (loading) {
      return <Spinner size={SpinnerSize.large} label="Loading dashboard..." className={styles.spinner} />;
    }

    if (selectedKey === OVERVIEW_KEY) {
      return (
        <OverviewPage
          quotationConfig={quotation}
          load={quotationLoad}
          tabs={tabs}
          listData={listData}
          onSelectStage={onSelectStage}
        />
      );
    }

    if (selectedKey === SALES_KEY || selectedKey === QUOTATION_KEY) {
      const isSales: boolean = selectedKey === SALES_KEY;
      return (
        <StageBoard
          config={quotation}
          quotationPageUrl={quotationPageUrl}
          generated={generated}
          webUrl={context.pageContext.web.absoluteUrl}
          leadSync={leadSync}
          load={quotationLoad}
          stages={isSales ? SALES_STATUSES : QUOTATION_STATUSES}
          heading={isSales ? 'Sales' : 'Customer Quotation'}
          description={
            isSales
              ? `The first ${SALES_STATUSES.length} stages, from Lead Identified to Quotation Draft.`
              : `The remaining ${QUOTATION_STATUSES.length} stages, from Quotation Prepared to Quotation Lost.`
          }
          onReload={reload}
          focusStatus={focusStatus}
          onFocusHandled={(): void => setFocusStatus(undefined)}
        />
      );
    }

    if (activeTab !== undefined) {
      return (
        <div className={styles.page}>
          <header className={styles.pageHead}>
            <Text variant="xLarge" className={styles.pageTitle} block>{activeTab.headerText}</Text>
            <Text variant="small" className={styles.pageSubtitle}>{activeTab.listTitle}</Text>
          </header>
          <ListTab
            config={activeTab}
            data={listData[activeTab.key] ?? EMPTY_TAB}
            kpis={kpisByTab[activeTab.key]}
            loading={false}
            probabilityThreshold={probabilityThreshold}
            onRefresh={(): void => {
              reload().catch((e: unknown): void => console.error(e));
            }}
          />
        </div>
      );
    }

    return undefined;
  };

  if (!hasQuotationList && tabs.length === 0) {
    return (
      <div className={styles.managerMaster}>
        <MessageBar messageBarType={MessageBarType.warning}>
          Nothing is configured yet. Open the web part properties and set the quotation list, or at least one list
          title.
        </MessageBar>
      </div>
    );
  }

  return (
    <div className={`${styles.managerMaster} ${styles.shell}`}>
      <SideNav
        items={navItems}
        selectedKey={selectedKey}
        onSelect={setSelectedKey}
        brand={title}
        collapsed={collapsed}
        onToggleCollapsed={(): void => setCollapsed(!collapsed)}
      />

      <div className={styles.main}>
        <header className={styles.topBar}>
          <div className={styles.topBarTitle}>
            <Text variant="large" block>{title}</Text>
            <Text variant="small" className={styles.topBarMeta}>{environmentMessage}</Text>
          </div>
          <Persona text={userDisplayName} size={PersonaSize.size32} className={styles.topBarPersona} />
        </header>

        {quotationLoad.error !== undefined && (
          <MessageBar messageBarType={MessageBarType.error} className={styles.messageBar} isMultiline>
            {quotationLoad.error}
          </MessageBar>
        )}

        <div className={styles.content}>{renderContent()}</div>
      </div>
    </div>
  );
};

export default ManagerMaster;
