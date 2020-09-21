/// <reference types="jest" />

import { Utils } from './GroupPeopleUtils';

import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

/**
 * @see https://github.com/Voitanos/jest-preset-spfx-react16/issues/7
 */
jest.mock('GroupPeopleWebPartStrings', () => {
    return {
        "size48": "Regular",
        "size72": "Large"
    }
},
    { virtual: true },
);

describe('GroupPeopleUtils', () => {

    it('should return an array of available picture sizes', () => {
        expect(Utils.enumSizesToOptions()).toEqual([
            { "key": "size48", "text": "Regular" },
            { "key": "size72", "text": "Large" },
            { "key": "size100", "text": "size100" }
        ]);
    });

    it('should return an empty array of SharePoint group as PropertyPaneDropdownOption', () => {
        const grp = null;
        expect(Utils.convertGrpToOptions(grp).length).toEqual(0);
        expect(Utils.convertGrpToOptions(grp)).toEqual(new Array<IPropertyPaneDropdownOption>());
    });

    it('should return an array of SharePoint groups as PropertyPaneDropdownOption', () => {
        const grp = [
            {
                AllowMembersEditMembership: false,
                AllowRequestToJoinLeave: false,
                AutoAcceptRequestToJoinLeave: false,
                Description: "Limited Access System Group",
                Id: 15,
                IsHiddenInUI: false,
                LoginName: "Limited Access System Group",
                OnlyAllowMembersViewMembership: true,
                OwnerTitle: "System Account",
                PrincipalType: 8,
                RequestToJoinLeaveEmailSetting: null,
                Title: "Limited Access System Group",
                "odata.editLink": "Web/SiteGroups/GetById(15)",
                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/SiteGroups/GetById(15)",
                "odata.type": "SP.Group"
            },
            {
                AllowMembersEditMembership: true,
                AllowRequestToJoinLeave: false,
                AutoAcceptRequestToJoinLeave: false,
                Description: null,
                Id: 5,
                IsHiddenInUI: false,
                LoginName: "Contoso Project - Group People Members",
                OnlyAllowMembersViewMembership: false,
                OwnerTitle: "Contoso Project - Group People Owners",
                PrincipalType: 8,
                RequestToJoinLeaveEmailSetting: "",
                Title: "Contoso Project - Group People Members",
                "odata.editLink": "Web/SiteGroups/GetById(5)",
                "odata.id": "https://ltsrdev.sharepoint.com/sites/project/_api/Web/SiteGroups/GetById(5)",
                "odata.type": "SP.Group"
            },
            {
                AllowMembersEditMembership: false,
                AllowRequestToJoinLeave: false,
                AutoAcceptRequestToJoinLeave: false,
                Description: null,
                Id: 3,
                IsHiddenInUI: false,
                LoginName: "Contoso Project - Group People Owners",
                OnlyAllowMembersViewMembership: false,
                OwnerTitle: "Contoso Project - Group People Owners",
                PrincipalType: 8,
                RequestToJoinLeaveEmailSetting: "",
                Title: "Contoso Project - Group People Owners",
                "odata.editLink": "Web/SiteGroups/GetById(3)",
                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/SiteGroups/GetById(3)",
                "odata.type": "SP.Group"
            },
            {
                AllowMembersEditMembership: false,
                AllowRequestToJoinLeave: false,
                AutoAcceptRequestToJoinLeave: false,
                Description: null,
                Id: 4,
                IsHiddenInUI: false,
                LoginName: "Contoso Project - Group People Visitors",
                OnlyAllowMembersViewMembership: false,
                OwnerTitle: "Contoso Project - Group People Owners",
                PrincipalType: 8,
                RequestToJoinLeaveEmailSetting: "",
                Title: "Contoso Project - Group People Visitors",
                "odata.editLink": "Web/SiteGroups/GetById(4)",
                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/SiteGroups/GetById(4)",
                "odata.type": "SP.Group"
            },
            {
                AllowMembersEditMembership: false,
                AllowRequestToJoinLeave: false,
                AutoAcceptRequestToJoinLeave: false,
                Description: "This group is for Flexible sharing links on item 'Shared Documents/Document.docx'",
                Id: 10,
                IsHiddenInUI: false,
                LoginName: "SharingLinks.00000000-1111-2222-3333-444444444444.Flexible.00000000-0000-0000-0000-000000000000",
                OnlyAllowMembersViewMembership: true,
                OwnerTitle: "System Account",
                PrincipalType: 8,
                RequestToJoinLeaveEmailSetting: null,
                Title: "SharingLinks.00000000-1111-2222-3333-444444444444.Flexible.00000000-0000-0000-0000-000000000000",
                "odata.editLink": "Web/SiteGroups/GetById(10)",
                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/SiteGroups/GetById(10)",
                "odata.type": "SP.Group"
            }
        ];
        expect(Utils.convertGrpToOptions(grp).length).toEqual(5);
    });
});