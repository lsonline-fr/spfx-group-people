import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import * as strings from 'GroupPeopleWebPartStrings';
import GroupPeople from './components/GroupPeople';
import GroupPeoplePlaceHolder from './components/GroupPeoplePlaceholder';
import { IGroupPeopleProps } from './components/IGroupPeopleProps';

import { sp } from "@pnp/sp";

export interface IGroupPeopleWebPartProps {
  SPGroups: string;
  Layout: string;
  CustomTitle: string;
  ToggleTitle: boolean;
}

/** Groupe People WebPart
 * @class
 * @extends
 */
export default class GroupPeopleWebPart extends BaseClientSideWebPart<IGroupPeopleWebPartProps> {

  /** SharePoint site groups
   * @private
   */
  private _spSiteGrps: any = null;

  /** Members of SharePoint group
   * @private
   */
  private _spGrpUsers: Array<any>[] = new Array;

  /** Title statement
   * Detect if only the title is edited from the property pane
   * @private
   */
  private _changeTitleState: boolean = false;

  /** Init WebPart
   * @returns
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
    if (!this._changeTitleState) {
      // reset userGroup
      this._spGrpUsers = [];
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
    this._changeTitleState = false;
  }

  /** Render the compact users layouts
   * @private
   */
  private postRender() {
    const element: React.ReactElement<IGroupPeopleProps> = React.createElement(GroupPeople, {
      title: this.properties.CustomTitle.length > 0 ? this.properties.CustomTitle : (this.properties.SPGroups !== undefined && this._spSiteGrps !== null) ? this._spSiteGrps.find(g => g.Id == this.properties.SPGroups).Title : '',
      users: this._spGrpUsers.sort((a:any,b:any) => (a.DisplayName > b.DisplayName) ? 1 : ((b.DisplayName > a.DisplayName) ? -1 : 0)),
      displayTitle: this.properties.ToggleTitle
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
   * @returns
   * @protected
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Customize the behavior of property pane change
   * @param targetProperty 
   * @param newValue 
   * @protected
   */
  protected onPropertyPaneFieldChanged(targetProperty: string, newValue: any) {
    if (targetProperty == 'CustomTitle' || targetProperty == 'ToggleTitle') {
      this._changeTitleState = true;
      this.postRender();
    }
  }

  /** Property Pane Configuration
   * @returns
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
                PropertyPaneToggle('ToggleTitle', {
                  label: strings.ToggleTitleLabel
                }),
                PropertyPaneTextField('CustomTitle', {
                  label: strings.CustomTitleLabel,
                  description: strings.CustomTitleDescription,
                  disabled: !this.properties.ToggleTitle
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /** Get all SharePoint Groups
   * This function exclude 'SharingLinks' groups
   * @return SharePoint groups
   * @async
   * @private
   */
  private async fetchSPGroups(): Promise<any> {
    return sp.web.siteGroups.get().then((grps) => { return grps.filter((g) => { return !/^SharingLinks./.test(g.LoginName) }); });
  }

  /** Get members of selected SharePoint group
   * This function ensure that users have PrincipalType to 1 and an email
   * @return SharePoint Group Members
   * @async
   * @private
   */
  private async fetchUsersGroup(): Promise<any> {
    // PrincipalType.User = 1 (SP.User)
    return sp.web.siteGroups.getById(parseInt(this.properties.SPGroups)).users.get().then((users) => { return users.filter((u) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }); });
  }

  /** Get User Profile specified by his LoginName
   * @param login LoginName of user
   * @return User Profile Properties
   * @async
   * @private
   */
  private async getUserProfile(login) {
    return sp.profiles.getPropertiesFor(login).then((r) => { return r; });
  }

  /** Convert array of SharePoint groups to array of DropDown options
   * @param grp Array of SharePoint groups
   * @return DropDown options
   * @private
   */
  private convertGrpToOptions(grp): IPropertyPaneDropdownOption[] {
    var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
    if (grp && grp.length > 0){
      grp.map((g: any) => {
        options.push({ key: g.Id, text: g.Title });
      });
    }
    return options;
  }
}
