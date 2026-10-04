import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IListTabConfig } from '../models/IDashboardModels';
import { ILeadSyncConfig, IQuotationConfig } from '../services/QuotationService';

export interface IManagerMasterProps {
  context: WebPartContext;
  title: string;
  /** The Customer Quotation pipeline, shown as the default tab. */
  quotation: IQuotationConfig;
  /** Mirrors sales stages onto the linked Sales Lead record. */
  leadSync: ILeadSyncConfig;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  /** One entry per pivot tab, in display order. */
  tabs: IListTabConfig[];
  itemLimit: number;
  probabilityThreshold: number;
}
