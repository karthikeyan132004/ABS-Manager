import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ManagerMasterWebPartStrings';
import ManagerMaster from './components/ManagerMaster';
import { IManagerMasterProps } from './components/IManagerMasterProps';
import { IListTabConfig } from './models/IDashboardModels';
import { ILeadSyncConfig, IQuotationConfig } from './services/QuotationService';
import { getSP } from '../../pnpConfig';

export interface IManagerMasterWebPartProps {
  title: string;
  itemLimit: number;
  probabilityThreshold: number;

  quotationListTitle: string;
  quotationStatusField: string;
  quotationAmountField: string;
  quotationRemarksField: string;
  quotationLeadField: string;
  leadStageField: string;

  leadsListTitle: string;
  leadsColumns: string;
  leadsDateField: string;
  leadsStatusField: string;
  leadsProbabilityField: string;

  activityListTitle: string;
  activityColumns: string;
  activityDateField: string;
  activityStatusField: string;

  projectsListTitle: string;
  projectsColumns: string;
  projectsDateField: string;
  projectsStatusField: string;
}

export default class ManagerMasterWebPart extends BaseClientSideWebPart<IManagerMasterWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IManagerMasterProps> = React.createElement(
      ManagerMaster,
      {
        context: this.context,
        title: this.properties.title !== undefined && this.properties.title !== ''
          ? this.properties.title
          : strings.DefaultDashboardTitle,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        quotation: this._buildQuotationConfig(),
        leadSync: this._buildLeadSyncConfig(),
        tabs: this._buildTabs(),
        itemLimit: this.properties.itemLimit ?? 500,
        probabilityThreshold: this.properties.probabilityThreshold ?? 70
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _buildQuotationConfig(): IQuotationConfig {
    return {
      listTitle: this.properties.quotationListTitle ?? '',
      statusField: this.properties.quotationStatusField ?? 'QuotStatus',
      amountField: this.properties.quotationAmountField ?? 'field_4',
      remarksField: this.properties.quotationRemarksField ?? 'Description'
    };
  }

  private _buildLeadSyncConfig(): ILeadSyncConfig {
    return {
      leadListTitle: this.properties.leadsListTitle ?? '',
      quotationLeadField: this.properties.quotationLeadField ?? 'SalesLead',
      leadStageField: this.properties.leadStageField ?? 'CustomerName'
    };
  }

  /** Only lists that have actually been named make it into the pivot. */
  private _buildTabs(): IListTabConfig[] {
    const candidates: IListTabConfig[] = [
      {
        key: 'leads',
        headerText: this.properties.leadsListTitle ?? '',
        listTitle: this.properties.leadsListTitle ?? '',
        columns: this.properties.leadsColumns ?? '',
        dateField: this.properties.leadsDateField ?? '',
        statusField: this.properties.leadsStatusField ?? '',
        probabilityField: this.properties.leadsProbabilityField ?? ''
      },
      {
        key: 'activity',
        headerText: this.properties.activityListTitle ?? '',
        listTitle: this.properties.activityListTitle ?? '',
        columns: this.properties.activityColumns ?? '',
        dateField: this.properties.activityDateField ?? '',
        statusField: this.properties.activityStatusField ?? '',
        probabilityField: ''
      },
      {
        key: 'projects',
        headerText: this.properties.projectsListTitle ?? '',
        listTitle: this.properties.projectsListTitle ?? '',
        columns: this.properties.projectsColumns ?? '',
        dateField: this.properties.projectsDateField ?? '',
        statusField: this.properties.projectsStatusField ?? '',
        probabilityField: ''
      }
    ];

    return candidates.filter((tab: IListTabConfig): boolean => tab.listTitle.trim() !== '');
  }

  protected onInit(): Promise<void> {
    // Hand the web part context to PnPjs once, before the first render.
    getSP(this.context);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
                PropertyPaneSlider('itemLimit', {
                  label: strings.ItemLimitFieldLabel,
                  min: 50,
                  max: 2000,
                  step: 50
                }),
                PropertyPaneSlider('probabilityThreshold', {
                  label: strings.ProbabilityThresholdFieldLabel,
                  min: 10,
                  max: 100,
                  step: 5
                })
              ]
            },
            {
              groupName: strings.QuotationGroupName,
              groupFields: [
                PropertyPaneTextField('quotationListTitle', { label: strings.ListTitleFieldLabel }),
                PropertyPaneTextField('quotationStatusField', {
                  label: strings.StatusFieldLabel,
                  description: strings.QuotationStatusFieldDescription
                }),
                PropertyPaneTextField('quotationAmountField', {
                  label: strings.AmountFieldLabel,
                  description: strings.OptionalFieldDescription
                }),
                PropertyPaneTextField('quotationRemarksField', {
                  label: strings.RemarksFieldLabel,
                  description: strings.OptionalFieldDescription
                }),
                PropertyPaneTextField('quotationLeadField', {
                  label: strings.LeadLinkFieldLabel,
                  description: strings.LeadLinkFieldDescription
                }),
                PropertyPaneTextField('leadStageField', {
                  label: strings.LeadStageFieldLabel,
                  description: strings.LeadStageFieldDescription
                })
              ]
            },
            {
              groupName: strings.LeadsGroupName,
              groupFields: [
                PropertyPaneTextField('leadsListTitle', { label: strings.ListTitleFieldLabel }),
                PropertyPaneTextField('leadsColumns', {
                  label: strings.ColumnsFieldLabel,
                  description: strings.ColumnsFieldDescription,
                  multiline: true
                }),
                PropertyPaneTextField('leadsDateField', {
                  label: strings.DateFieldLabel,
                  description: strings.DateFieldDescription
                }),
                PropertyPaneTextField('leadsStatusField', {
                  label: strings.StatusFieldLabel,
                  description: strings.StatusFieldDescription
                }),
                PropertyPaneTextField('leadsProbabilityField', {
                  label: strings.ProbabilityFieldLabel,
                  description: strings.ProbabilityFieldDescription
                })
              ]
            },
            {
              groupName: strings.ActivityGroupName,
              groupFields: [
                PropertyPaneTextField('activityListTitle', { label: strings.ListTitleFieldLabel }),
                PropertyPaneTextField('activityColumns', {
                  label: strings.ColumnsFieldLabel,
                  description: strings.ColumnsFieldDescription,
                  multiline: true
                }),
                PropertyPaneTextField('activityDateField', {
                  label: strings.DateFieldLabel,
                  description: strings.DateFieldDescription
                }),
                PropertyPaneTextField('activityStatusField', {
                  label: strings.StatusFieldLabel,
                  description: strings.StatusFieldDescription
                })
              ]
            },
            {
              groupName: strings.ProjectsGroupName,
              groupFields: [
                PropertyPaneTextField('projectsListTitle', { label: strings.ListTitleFieldLabel }),
                PropertyPaneTextField('projectsColumns', {
                  label: strings.ColumnsFieldLabel,
                  description: strings.ColumnsFieldDescription,
                  multiline: true
                }),
                PropertyPaneTextField('projectsDateField', {
                  label: strings.DateFieldLabel,
                  description: strings.DateFieldDescription
                }),
                PropertyPaneTextField('projectsStatusField', {
                  label: strings.StatusFieldLabel,
                  description: strings.StatusFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
