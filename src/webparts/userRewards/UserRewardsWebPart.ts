import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';
import * as strings from 'UserRewardsWebPartStrings';
import UserRewards from './components/UserRewards';
import { IUserRewardsProps } from './components/IUserRewardsProps';
import { sp } from '@pnp/pnpjs';
import RetrieveListDataService, { IRetrieveListDataService } from '../../services/RetrieveListDataService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IUserRewardsWebPartProps {
  description: string;
  selectedList: string;
  numberOfItemsToDisplay: number;
  showAll: boolean;
}

export default class UserRewardsWebPart extends BaseClientSideWebPart<IUserRewardsWebPartProps> {

  private _service: IRetrieveListDataService;
  private listsDropdownOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<IUserRewardsProps> = React.createElement(
      UserRewards,
      {
        service: this._service,
        context: this.context,
        
      }
    );

    ReactDom.render(element, this.domElement);
  }
  
  protected onPropertyPaneConfigurationStart(): void {
    this.loadLists();
  }
  private loadLists(): void {
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.error(`Failed to load lists. Error: ${response.statusText}`);
        }
      })
      .then((data: any) => {
        if (data && data.value) {
          this.listsDropdownOptions = data.value.map((list: any) => {
            return {
              key: list.Title,
              text: list.Title,
            };
          });
          this.context.propertyPane.refresh();
          this.render();
        }
      })
      .catch((error: any) => {
        console.error(`Error loading lists: ${error}`);
      });
  }

  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context as any
    });
    this._service = new RetrieveListDataService(this.context, this.properties.selectedList);
    return super.onInit();
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
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select a List',
                  options: this.listsDropdownOptions,
                }),
                PropertyPaneCheckbox("showAll", {
                  text: "Display All Items"
                }),
                PropertyPaneSlider("numberOfItemsToDisplay", {
                  label: "Number Of Items To Display",
                  min: 1,
                  max: 10,
                  value: 3,
                  showValue: true,
                  step: 1,
                  disabled: this.properties.showAll
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
