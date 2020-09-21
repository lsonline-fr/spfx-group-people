import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';

import * as strings from 'GroupPeopleWebPartStrings';

import GroupPeople from './components/GroupPeople';
import GroupPeoplePlaceHolder from './components/GroupPeoplePlaceholder';
import { IGroupPeopleProps } from './components/IGroupPeopleProps';

import SPGroupService from './services/SPGroupService';
import MockSPGroupService from './mocks/MockSPGroupService';

import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';

import PeopleCard from './models/PeopleCard';
import ISPGroupService from './models/ISPGroupService';
import { ISiteGroupInfo } from './models/ISiteGroupInfo';
import { ISiteUserInfo } from './models/ISiteUserInfo';
import { Utils } from './GroupPeopleUtils';

export interface IGroupPeopleWebPartProps {
  SPGroups: string;
  Layout: string;
  CustomTitle: string;
  ToggleTitle: boolean;
  PictureSize: string;
  HideWebPart: boolean;
  PictureUrl: string;
  LineOne: string;
  LineTwo: string;
  LineThree: string;
}

/** Groupe People WebPart
 * @class
 * @extends
 */
export default class GroupPeopleWebPart extends BaseClientSideWebPart<IGroupPeopleWebPartProps> {

  /** List of available SharePoint site groups
   * @private
   */
  private _spSiteGrps: ISiteGroupInfo[];

  /** List of thte members available into the selected SharePoint group
   * @private
   */
  private _spGrpUsers: Array<PeopleCard>;

  /** Partial update statement
   * Detect if the webpart must be render partially in accordance with somes properties pane
   * @private
   */
  private _partialUpdateRender: boolean = false;

  /** SharePoint Group Service
   * @private
   */
  private _spGrpSvc: ISPGroupService;

  /**
   * SharePoint selected group title
   * @private
   */
  private _grpTitle: string;

  /** Init WebPart
   * @returns
   * @protected
   */
  protected onInit(): Promise<void> {
    this._spGrpSvc = Environment.type == EnvironmentType.Local ? new MockSPGroupService(this.context.pageContext.site.absoluteUrl) : new SPGroupService(this.context.pageContext.site.absoluteUrl);

    if (DisplayMode.Edit == this.displayMode) { /* Get all SharePoint groups only in edit mode */
      this._spGrpSvc.fetchSPGroups().then((spGroups: Array<ISiteGroupInfo>) => {
        this._spSiteGrps = spGroups;
        if (this.properties.SPGroups) {
          this.postRender();
        } else {
          this.render();
        }
      });
    } 
    if (DisplayMode.Read == this.displayMode && this.properties.SPGroups) {
      this._spGrpSvc.getSPGroup(parseInt(this.properties.SPGroups)).then((grp: ISiteGroupInfo) => {
        this._grpTitle = grp.Title;
        this.postRender();
      });
    }
    return super.onInit();
  }

  /** Default Render
   * @public
   */
  public render(): void {
    if (!this.properties.SPGroups || null == this.properties.SPGroups) {
      const element: React.ReactElement<GroupPeoplePlaceHolder> = React.createElement(GroupPeoplePlaceHolder);
      ReactDom.render(element, this.domElement);
    } else if (!this._partialUpdateRender && this.properties.SPGroups) {
      this.getUsersGroup();
    }
    this._partialUpdateRender = false;
  }

  /** Render the compact users layouts
   * @private
   */
  private postRender() {
    if ((this._spGrpUsers && this._spGrpUsers.length > 0) || (this._spGrpUsers && this._spGrpUsers.length == 0 && (undefined == this.properties.HideWebPart || false == this.properties.HideWebPart || DisplayMode.Edit == this.displayMode))) {
      this._grpTitle = this._grpTitle ? this._grpTitle : (undefined !== this.properties.SPGroups && null != this._spSiteGrps) ? this._spSiteGrps.find(g => g.Id == parseInt(this.properties.SPGroups)).Title : '';
      const element: React.ReactElement<IGroupPeopleProps> = React.createElement(GroupPeople, {
        title: (this.properties.CustomTitle && this.properties.CustomTitle.length > 0) ? this.properties.CustomTitle : this._grpTitle,
        users: (this._spGrpUsers && this._spGrpUsers.length > 0) ? this._spGrpUsers.sort((a: PeopleCard, b: PeopleCard) => (a.lineOne > b.lineOne) ? 1 : ((b.lineOne > a.lineOne) ? -1 : 0)) : new Array,
        size: PersonaSize[this.properties.PictureSize],
        displayTitle: this.properties.ToggleTitle,
        hide: (DisplayMode.Read == this.displayMode && this.properties.HideWebPart) ? true : false
      });
      ReactDom.render(element, this.domElement);
    } else {
      this.onDispose();
    }
  }

  /**
   * Get all users from the selected SharePoint group and then populate the people cards
   */
  private getUsersGroup() {
    this._spGrpSvc.fetchUsersGroup(parseInt(this.properties.SPGroups)).then((users: Array<ISiteUserInfo>) => {
      return users;
    }).then((u: Array<ISiteUserInfo>) => {
      this._spGrpUsers = new Array;
      if (null != u && u.length > 0) {
        this.populatePeopleCards(u);
      } else {
        this.postRender();
      }
    });
  }

  /**
   * Get the user profiles and populate the list of People Cards
   * @param u List of users informations members of the selected SharePoint groups
   */
  private populatePeopleCards(u: ISiteUserInfo[]) {
    let uCount = 0;
    u.forEach(user => {
      this._spGrpSvc.getUserProfile(user.LoginName).then((r) => {
        try {
          if (null != r && undefined != r) { // Ensure at least one user profile was found
            this._spGrpUsers.push(new PeopleCard(
              r.find(props => props.Key == 'AccountName').Value,
              r.find(props => props.Key == this.properties.PictureUrl) ? r.find(props => props.Key == this.properties.PictureUrl).Value : '',
              r.find(props => props.Key == this.properties.LineOne) ? r.find(props => props.Key == this.properties.LineOne).Value : '',
              r.find(props => props.Key == this.properties.LineTwo) ? r.find(props => props.Key == this.properties.LineTwo).Value : '',
              r.find(props => props.Key == this.properties.LineThree) ? r.find(props => props.Key == this.properties.LineThree).Value : ''
            ));
          }
        } catch (e) { /*console.log(e);*/ }
        uCount++;
        // Once all profiles are parsed, start the render
        if (uCount == u.length) {
          this.postRender();
        }
      });
    });
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
    if ('CustomTitle' == targetProperty || 'ToggleTitle' == targetProperty) {
      this._partialUpdateRender = true;
      this.postRender();
    } else if ('PictureSize' == targetProperty) {
      this._partialUpdateRender = true;
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
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('SPGroups', {
                  label: strings.DropdownGroupLabel,
                  options: Utils.convertGrpToOptions(this._spSiteGrps)
                }),
                PropertyPaneToggle('ToggleTitle', {
                  label: strings.ToggleTitleLabel
                }),
                PropertyPaneTextField('CustomTitle', {
                  label: strings.CustomTitleLabel,
                  description: strings.CustomTitleDescription,
                  disabled: !this.properties.ToggleTitle
                }),
                PropertyPaneDropdown('PictureSize', {
                  label: strings.PictureSize,
                  options: Utils.enumSizesToOptions(),
                  selectedKey: Utils.enumSizesToOptions()[0].key // Select first value by default
                }),
                PropertyPaneLabel('LabelSeparator', {
                  text: ' '
                }),
                PropertyPaneCheckbox('HideWebPart', {
                  text: strings.HideWebPart
                })
              ]
            },
            {
              groupName: strings.FieldsGroupLabel,
              groupFields: [
                PropertyPaneTextField('PictureUrl', {
                  label: strings.PictureUrl
                }),
                PropertyPaneTextField('LineOne', {
                  label: strings.LineOne
                }),
                PropertyPaneTextField('LineTwo', {
                  label: strings.LineTwo
                }),
                PropertyPaneTextField('LineThree', {
                  label: strings.LineThree,
                  description: strings.LineThreeDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
