import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneDropdownOptionType,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'GroupPeopleWebPartStrings';
import GroupPeople from './components/GroupPeople';
import { IGroupPeopleProps } from './components/IGroupPeopleProps';

import { sp } from "@pnp/sp";
import { SiteGroup } from '@pnp/sp/src/sitegroups';

export interface IGroupPeopleWebPartProps {
  SPGroups: string;
}

/** Groupe People WebPart
 * @class
 * @extends BaseClientSideWebPart
 */
export default class GroupPeopleWebPart extends BaseClientSideWebPart<IGroupPeopleWebPartProps> {

  /**
   * @type {IPropertyPaneDropdownOption[]}
   * @property
   * @private
   */
  private _spSiteGrps: IPropertyPaneDropdownOption[] = [];

  /** Init WebPart
   * @protected
   */
  protected onInit(): Promise<void> {

    sp.setup({
      spfxContext: this.context
    });

    this.fetchGrpOptions().then((g) => {
      this._spSiteGrps = g;
    });

    return super.onInit();
  }

  /** Default Render
   * @public
   */
  public render(): void {
    const element: React.ReactElement<IGroupPeopleProps> = React.createElement(
      GroupPeople,
      {
        title: 'toto'
      }
    );

    sp.web.siteGroups.getById(parseInt(this.properties.SPGroups)).users.get().then((users) => {
      console.log(users);
    });

    ReactDom.render(element, this.domElement);
  }

  /** On Dispose
   * @protected
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /** Version
   * @returns {Version}
   * @protected
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Property Pane Configuration
   * @returns {IPropertyPaneConfiguration}
   * @property
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('SPGroups', {
                  label: strings.DropdownGroupLabel,
                  options: this._spSiteGrps
                }),
                PropertyPaneChoiceGroup('layout', {
                  label: strings.LayoutGroupLabel,
                  options: [
                    { key: 'compact', text: 'Compact', imageSize: { width: 32, height: 32 }, iconProps: { officeFabricIconFontName: 'DockLeft' } },
                    { key: 'descriptive', text: 'Descriptive', imageSize: { width: 32, height: 32 }, iconProps: { officeFabricIconFontName: 'DiffInline' } }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * Get all groups of SharePoint Site
   * @returns {Array<SiteGroup>}
   * @async
   * @private
   */
  private async fetchSPGroups(): Promise<any> {
    return sp.web.siteGroups.get().then((grps) => { return grps; });
  }

  /**
   * Convert list of SharePoint groups to PropertyPaneDropdown options
   * @returns {Array<IPropertyPaneDropdownOption>}
   * @async
   * @private
   */
  private async fetchGrpOptions(): Promise<IPropertyPaneDropdownOption[]> {
    return this.fetchSPGroups().then((r) => {
      var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      r.map((g: any) => {
        options.push({ key: g.Id, text: g.Title });
      });
      return options;
    });
  }
}
