import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'GroupPeopleWebPartStrings';
import GroupPeople from './components/GroupPeople';
import GroupPeoplePlaceHolder from './components/GroupPeoplePlaceholder';
import { IGroupPeopleProps } from './components/IGroupPeopleProps';

import { sp } from "@pnp/sp";

export interface IGroupPeopleWebPartProps {
  SPGroups: string;
  Layout: string;
  CustomTitle: string;
}

/** Groupe People WebPart
 * @class
 * @extends BaseClientSideWebPart
 */
export default class GroupPeopleWebPart extends BaseClientSideWebPart<IGroupPeopleWebPartProps> {

  private _spSiteGrps: any = null;

  private _spGrpUsers: Array<any>[] = new Array;

  /** Init WebPart
   * @protected
   */
  protected onInit(): Promise<void> {

    sp.setup({
      spfxContext: this.context
    });

    this.fetchSPGroups().then((spGroups) => {
      this._spSiteGrps = spGroups;
    });

    return super.onInit();
  }

  /** Default Render
   * @public
   */
  public render(): void {
    // Check if a SharePoint group was selected. If not, display the PlaceHolder
    if (this.properties.SPGroups) {
      // Get Users from selected group
      this.fetchUsersGroup().then((users) => {
        return users;
      }).then((u: any) => {
        if (u.length > 0) {
          // Get for each user, their profile information
          u.forEach(user => {
            this.getUserProfile(user.LoginName).then((r) => {
              // Store the information to a property
              this._spGrpUsers.push(r);

              // Once all profile parsed, start the render
              if (this._spGrpUsers.length == u.length) {
                this.postRender();
              }
            });
          });
        } else {
          this._spGrpUsers = new Array;
          this.postRender();
        }
      });
    } else {
      const element: React.ReactElement<GroupPeoplePlaceHolder> = React.createElement(GroupPeoplePlaceHolder);
      ReactDom.render(element, this.domElement);
    }
  }

  private postRender() {
    const element: React.ReactElement<IGroupPeopleProps> = React.createElement(GroupPeople, {
      title: this.properties.CustomTitle.length > 0 ? this.properties.CustomTitle : this._spSiteGrps.find(g => g.Id === this.properties.SPGroups).Title,
      users: this._spGrpUsers
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
                  options: this.convertGrpToOptions(this._spSiteGrps)
                }),
                PropertyPaneTextField('CustomTitle', {
                  label: strings.CustomTitleLabel,
                  description: strings.CustomTitleDescription
                }),
                PropertyPaneChoiceGroup('Layout', {
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

  private async fetchSPGroups(): Promise<any> {
    return sp.web.siteGroups.get().then((grps) => { return grps; });
  }

  private async fetchUsersGroup(): Promise<any> {
    return sp.web.siteGroups.getById(parseInt(this.properties.SPGroups)).users.get().then((users) => { return users.filter(function (u) { return u.UserId != null && u.Email != null && u.Email.length > 0; }); });
  }

  private async getUserProfile(login) {
    return sp.profiles.getPropertiesFor(login).then((r) => { return r; });
  }

  private convertGrpToOptions(grp): IPropertyPaneDropdownOption[] {
    var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
    grp.map((g: any) => {
      options.push({ key: g.Id, text: g.Title });
    });
    return options;
  }
}
