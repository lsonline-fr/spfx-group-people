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
 * @extends BaseClientSideWebPart
 */
export default class GroupPeopleWebPart extends BaseClientSideWebPart<IGroupPeopleWebPartProps> {

  private _spSiteGrps: any = null;

  private _spGrpUsers: Array<any>[] = new Array;

  private _changeTitleState: boolean = false;

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

  /**
   * Render of compact users
   * @param {Array<SP.UserProfiles.PersonProperties>} grpUsers Array of user profile properties
   * @private
   */
  private postRender(grpUsers = undefined) {
    const element: React.ReactElement<IGroupPeopleProps> = React.createElement(GroupPeople, {
      title: this.properties.CustomTitle.length > 0 ? this.properties.CustomTitle : (this.properties.SPGroups !== undefined) ? this._spSiteGrps.find(g => g.Id == this.properties.SPGroups).Title : '',
      users: grpUsers !== undefined ? grpUsers : this._spGrpUsers,
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
   * @returns {Version}
   * @protected
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Customize the behavior of property pane change
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
                PropertyPaneToggle('ToggleTitle', {
                  label: strings.ToggleTitleLabel
                }),
                PropertyPaneTextField('CustomTitle', {
                  label: strings.CustomTitleLabel,
                  description: strings.CustomTitleDescription,
                  disabled: !this.properties.ToggleTitle
                })/*,
                PropertyPaneChoiceGroup('Layout', {
                  label: strings.LayoutGroupLabel,
                  options: [
                    { key: 'compact', text: 'Compact', imageSize: { width: 32, height: 32 }, iconProps: { officeFabricIconFontName: 'DockLeft' } },
                    { key: 'descriptive', text: 'Descriptive', imageSize: { width: 32, height: 32 }, iconProps: { officeFabricIconFontName: 'DiffInline' } }
                  ]
                })*/
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * Get all SharePoint Groups
   * @return {Array<SP.Group>} SharePoint groups
   * @async
   * @private
   */
  private async fetchSPGroups(): Promise<any> {
    return sp.web.siteGroups.get().then((grps) => { return grps; });
  }

  /**
   * Get members of selected SharePoint group
   * @return {Array<SP.User>} SharePoint Group Members
   * @async
   * @private
   */
  private async fetchUsersGroup(): Promise<any> {
    // PrincipalType.User = 1 (SP.User)
    return sp.web.siteGroups.getById(parseInt(this.properties.SPGroups)).users.get().then((users) => { return users.filter((u) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }); });
  }

  /**
   * Get User Profile specified by his LoginName
   * @param {string} login LoginName of user
   * @return {SP.UserProfiles.PersonProperties} User Profile Properties
   * @async
   * @private
   */
  private async getUserProfile(login) {
    return sp.profiles.getPropertiesFor(login).then((r) => { return r; });
  }

  /**
   * Convert array of SharePoint groups to array of DropDown options
   * @param {Array<SP.Group>} grp Array of SharePoint groups
   * @return {Array<IPropertyPaneDropdownOption>} DropDown options
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
