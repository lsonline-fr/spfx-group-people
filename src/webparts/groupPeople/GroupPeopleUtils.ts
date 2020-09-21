import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

import * as strings from 'GroupPeopleWebPartStrings';

import { PictureSizes } from './models/PictureSizes';
import { ISiteGroupInfo } from './models/ISiteGroupInfo';

/**
 * Utils
 * @class
 */
export class Utils {

    /** Convert array of SharePoint groups into an array of DropDown options
     * @param grp Array of SharePoint groups
     * @return DropDown options of SharePoint groups
     */
    public static convertGrpToOptions(grp: Array<ISiteGroupInfo>): Array<IPropertyPaneDropdownOption> {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        if (grp && grp.length > 0) {
            grp.map((g: ISiteGroupInfo) => {
                options.push({ key: g.Id, text: g.Title });
            });
        }
        return options;
    }

    /** Convert ENUMS of People Card sizes into Dropdown options
     * @return DropDown options of People Card sizes
     */
    public static enumSizesToOptions(): Array<IPropertyPaneDropdownOption> {
        let sizeOpt: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        let sizes: Array<String> = Object.keys(PictureSizes).filter(x => !(parseInt(x) >= 0));
        sizes.forEach((s) => {
            sizeOpt.push({ key: s.toString(), text: strings[s.toString()] ? strings[s.toString()] : s.toString() });
        });
        return sizeOpt;
    }
}