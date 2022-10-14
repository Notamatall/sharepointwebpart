import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'FilterWebPartStrings';
import Filter from './components/Filter';
import { IFilterProps } from './components/IFilterProps';
import { SPHttpClient } from '@microsoft/sp-http';

export type ButtonClickedCallback = () => void;
export default class FilterWebPart extends BaseClientSideWebPart<{}> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
  }

  public async getListItems() {
    // const response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/Lists/GetByTitle('Documents')/Items", SPHttpClient.configurations.v1, {});
    const response = await this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "/_api/web/Lists/GetByTitle('Documents')/GetItems(query=@v1)?@v1={}&$select=*,Editor,File_x0020_Type,FileRef,Modified_x0020_By,FileLeafRef, EncodedAbsUrl", SPHttpClient.configurations.v1, {});
    //console.log(await response.json());

    return (await response.json());
  }

  public async render(): Promise<void> {

    const list = await this.getListItems();

    const element: React.ReactElement<IFilterProps> = React.createElement(
      Filter,
      {
        context: this.context,
        list: list,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
