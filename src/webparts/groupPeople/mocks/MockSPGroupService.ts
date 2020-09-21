import ISPGroupService from "../models/ISPGroupService";
import { ISiteGroupInfo } from "../models/ISiteGroupInfo";
import { ISiteUserInfo } from "../models/ISiteUserInfo";

export default class SPGroupService implements ISPGroupService {

    private _siteUrl: string;

    /** Default constructor
     * @param u SharePoint site URL
     */
    constructor(u: string) {
        if (null == u || u.trim().length < 10) {
            throw new TypeError('SharePoint site URL can not be null or empty.');
        }
        this._siteUrl = u;
    }

    public fetchSPGroups(): Promise<Array<ISiteGroupInfo>> {
        return new Promise<Array<ISiteGroupInfo>>((resolve, reject) => {
            setTimeout(() => {
                resolve([
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
                ].filter((g: ISiteGroupInfo) => { return !/^SharingLinks./.test(g.LoginName); }) as unknown as Array<ISiteGroupInfo>);
            }, 500);
        });
    }

    public getSPGroup(i: number): Promise<ISiteGroupInfo> {
        return new Promise<any>((resolve, reject) => {
            setTimeout(() => {
                switch (i) {
                    case 15:
                        resolve({
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
                        } as any);
                        break;
                    case 5:
                        resolve({
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
                        } as any);
                        break;
                    case 3:
                        resolve({
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
                        } as any);
                        break;
                    case 4:
                        resolve({
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
                        } as any);
                        break;
                    default:
                        resolve(null as any);
                }
            }, 500);
        });
    }

    public fetchUsersGroup(i: number): Promise<Array<ISiteUserInfo>> {
        return new Promise<Array<ISiteUserInfo>>((resolve, reject) => {
            setTimeout(() => {
                switch (i) {
                    case 5:
                        resolve([
                            {
                                Email: "user1@contoso.onmicrosoft.com",
                                Expiration: "",
                                Id: 10,
                                IsEmailAuthenticationGuestUser: false,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: false,
                                IsSiteAdmin: false,
                                LoginName: "i:0#.f|membership|user1@contoso.onmicrosoft.com",
                                PrincipalType: 1,
                                Title: "User 01",
                                UserPrincipalName: "user1@contoso.onmicrosoft.com",
                                "odata.editLink": "Web/GetUserById(10)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(10)",
                                "odata.type": "SP.User"
                            },
                            {
                                Email: "user2@contoso.onmicrosoft.com",
                                Expiration: "",
                                Id: 11,
                                IsEmailAuthenticationGuestUser: false,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: false,
                                IsSiteAdmin: false,
                                LoginName: "i:0#.f|membership|user2@contoso.onmicrosoft.com",
                                PrincipalType: 1,
                                Title: "User 02",
                                UserPrincipalName: "user2@contoso.onmicrosoft.com",
                                "odata.editLink": "Web/GetUserById(11)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(11)",
                                "odata.type": "SP.User"
                            },
                            {
                                Email: "",
                                Expiration: "",
                                Id: 12,
                                IsEmailAuthenticationGuestUser: false,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: false,
                                IsSiteAdmin: false,
                                LoginName: "c:0t.c|tenant|00000000-0000-0000-0000-000000000000",
                                PrincipalType: 4,
                                Title: "Security Group Only",
                                UserPrincipalName: null,
                                "odata.editLink": "Web/GetUserById(12)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(12)",
                                "odata.type": "SP.User"
                            },
                            {
                                Email: "user3@contoso.onmicrosoft.com",
                                Expiration: "",
                                Id: 13,
                                IsEmailAuthenticationGuestUser: false,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: false,
                                IsSiteAdmin: false,
                                LoginName: "i:0#.f|membership|user3@contoso.onmicrosoft.com",
                                PrincipalType: 1,
                                Title: "User 03",
                                UserPrincipalName: "user3@contoso.onmicrosoft.com",
                                "odata.editLink": "Web/GetUserById(13)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(13)",
                                "odata.type": "SP.User"
                            },
                            {
                                Email: "external@outlook.com",
                                Expiration: "",
                                Id: 15,
                                IsEmailAuthenticationGuestUser: true,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: true,
                                IsSiteAdmin: false,
                                LoginName: "i:0#.f|membership|urn%3aspo%3aguest#external@outlook.com",
                                PrincipalType: 1,
                                Title: "external@outlook.com",
                                UserPrincipalName: null,
                                "odata.editLink": "Web/GetUserById(15)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(15)",
                                "odata.type": "SP.User"
                            }
                        ].filter((u) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }) as unknown as Array<ISiteUserInfo>);
                        break;
                    case 3:
                        resolve([
                            {
                                Email: "user1@contoso.onmicrosoft.com",
                                Expiration: "",
                                Id: 10,
                                IsEmailAuthenticationGuestUser: false,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: false,
                                IsSiteAdmin: false,
                                LoginName: "i:0#.f|membership|user1@contoso.onmicrosoft.com",
                                PrincipalType: 1,
                                Title: "User 01",
                                UserPrincipalName: "user1@contoso.onmicrosoft.com",
                                "odata.editLink": "Web/GetUserById(10)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(10)",
                                "odata.type": "SP.User"
                            }
                        ].filter((u) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }) as unknown as Array<ISiteUserInfo>);
                        break;
                    case 4:
                        resolve([].filter((u) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }) as unknown as Array<ISiteUserInfo>);
                        break;
                    case 10:
                        resolve([
                            {
                                Email: "external@outlook.com",
                                Expiration: "",
                                Id: 15,
                                IsEmailAuthenticationGuestUser: true,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: true,
                                IsSiteAdmin: false,
                                LoginName: "i:0#.f|membership|urn%3aspo%3aguest#external@outlook.com",
                                PrincipalType: 1,
                                Title: "external@outlook.com",
                                UserPrincipalName: null,
                                "odata.editLink": "Web/GetUserById(15)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(15)",
                                "odata.type": "SP.User"
                            }
                        ].filter((u) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }) as unknown as Array<ISiteUserInfo>);
                        break;
                    case 15:
                        resolve([
                            {
                                Email: "external@outlook.com",
                                Expiration: "",
                                Id: 15,
                                IsEmailAuthenticationGuestUser: true,
                                IsHiddenInUI: false,
                                IsShareByEmailGuestUser: true,
                                IsSiteAdmin: false,
                                LoginName: "i:0#.f|membership|urn%3aspo%3aguest#external@outlook.com",
                                PrincipalType: 1,
                                Title: "external@outlook.com",
                                UserPrincipalName: null,
                                "odata.editLink": "Web/GetUserById(15)",
                                "odata.id": "https://contoso.sharepoint.com/sites/project/_api/Web/GetUserById(15)",
                                "odata.type": "SP.User"
                            }
                        ].filter((u) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }) as unknown as Array<ISiteUserInfo>);
                        break;
                }
            }, 500);
        });
    }

    public getUserProfile(login: string): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            setTimeout(() => {
                switch (login) {
                    case 'i:0#.f|membership|user1@contoso.onmicrosoft.com':
                        resolve([
                            { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000000", ValueType: "Edm.String" },
                            { Key: "SID", Value: "i:0h.f|membership|1486325dfzvd54@live.com", ValueType: "Edm.String" },
                            { Key: "ADGuid", Value: "System.Byte[]", ValueType: "Edm.String" },
                            { Key: "AccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "FirstName", Value: "Annie", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                            { Key: "LastName", Value: "Lindqvist", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticLastName", Value: "", ValueType: "Edm.String" },
                            { Key: "PreferredName", Value: "Annie Lindqvist", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticDisplayName", Value: "", ValueType: "Edm.String" },
                            { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                            { Key: "Department", Value: "", ValueType: "Edm.String" },
                            { Key: "Title", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                            { Key: "SPS-Department", Value: "Information Technology", ValueType: "Edm.String" },
                            { Key: "Manager", Value: "", ValueType: "Edm.String" },
                            { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                            { Key: "PersonalSpace", Value: "/personal/user1_contoso_onmicrosoft_com/", ValueType: "Edm.String" },
                            { Key: "PictureURL", Value: "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png", ValueType: "Edm.String" },
                            { Key: "UserName", Value: "user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "QuickLinks", Value: "", ValueType: "Edm.String" },
                            { Key: "WebSite", Value: "", ValueType: "Edm.String" },
                            { Key: "PublicSiteRedirect", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-JobTitle", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                            { Key: "SPS-Dotted-line", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Peers", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Responsibility", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-SipAddress", Value: "user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-MySiteUpgrade", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ProxyAddresses", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-HireDate", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-DisplayOrder", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimID", Value: "user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimProviderID", Value: "membership", ValueType: "Edm.String" },
                            { Key: "SPS-ResourceSID", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ResourceAccountName", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-MasterAccountName", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-UserPrincipalName", Value: "user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-O15FirstRunExperience", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteInstantiationState", Value: "2", ValueType: "Edm.String" },
                            { Key: "SPS-DistinguishedName", Value: "CN=00000000-0000-0000-0000-000000000000,OU=dsfsf51…Tenants,OU=MSOnline,DC=SPOCONTOSO,DC=msft,DC=net", ValueType: "Edm.String" },
                            { Key: "SPS-SourceObjectDN", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimProviderType", Value: "Forms", ValueType: "Edm.String" },
                            { Key: "SPS-SavedAccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-SavedSID", Value: "System.Byte[]", ValueType: "Edm.String" },
                            { Key: "SPS-ObjectExists", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteCapabilities", Value: "4", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteFirstCreationTime", Value: "2/24/2020 2:04:23 PM", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteLastCreationTime", Value: "2/24/2020 2:04:23 PM", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteNumberOfRetries", Value: "1", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteFirstCreationError", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-FeedIdentifier", Value: "", ValueType: "Edm.String" },
                            { Key: "WorkEmail", Value: "user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "CellPhone", Value: "", ValueType: "Edm.String" },
                            { Key: "Fax", Value: "", ValueType: "Edm.String" },
                            { Key: "HomePhone", Value: "", ValueType: "Edm.String" },
                            { Key: "Office", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Location", Value: "", ValueType: "Edm.String" },
                            { Key: "Assistant", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PastProjects", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Skills", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-School", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Birthday", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-StatusNotes", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Interests", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-HashTags", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-EmailOptin", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PrivacyPeople", Value: "True", ValueType: "Edm.String" },
                            { Key: "SPS-PrivacyActivity", Value: "4095", ValueType: "Edm.String" },
                            { Key: "SPS-PictureTimestamp", Value: "63720225810", ValueType: "Edm.String" },
                            { Key: "SPS-PicturePlaceholderState", Value: "0", ValueType: "Edm.String" },
                            { Key: "SPS-PictureExchangeSyncState", Value: "1", ValueType: "Edm.String" },
                            { Key: "SPS-TimeZone", Value: "", ValueType: "Edm.String" },
                            { Key: "OfficeGraphEnabled", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-UserType", Value: "0", ValueType: "Edm.String" },
                            { Key: "SPS-HideFromAddressLists", Value: "False", ValueType: "Edm.String" },
                            { Key: "SPS-RecipientTypeDetails", Value: "", ValueType: "Edm.String" },
                            { Key: "DelveFlags", Value: "", ValueType: "Edm.String" },
                            { Key: "msOnline-ObjectId", Value: "00000000-0000-0000-0000-00000000000", ValueType: "Edm.String" },
                            { Key: "SPS-PointPublishingUrl", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-TenantInstanceId", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-SharePointHomeExperienceState", Value: "134671", ValueType: "Edm.String" },
                            { Key: "SPS-MultiGeoFlags", Value: "", ValueType: "Edm.String" },
                            { Key: "PreferredDataLocation", Value: "", ValueType: "Edm.String" }
                        ] as any);
                        break;
                    case 'i:0#.f|membership|user2@contoso.onmicrosoft.com':
                        resolve([
                            { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000001", ValueType: "Edm.String" },
                            { Key: "SID", Value: "i:0h.f|membership|6545646s5d4fdfs@live.com", ValueType: "Edm.String" },
                            { Key: "ADGuid", Value: "System.Byte[]", ValueType: "Edm.String" },
                            { Key: "AccountName", Value: "i:0#.f|membership|user2@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "FirstName", Value: "Ted", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                            { Key: "LastName", Value: "Randall", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticLastName", Value: "", ValueType: "Edm.String" },
                            { Key: "PreferredName", Value: "Ted Randall", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticDisplayName", Value: "", ValueType: "Edm.String" },
                            { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                            { Key: "Department", Value: "", ValueType: "Edm.String" },
                            { Key: "Title", Value: "Microsoft 365 Developer", ValueType: "Edm.String" },
                            { Key: "SPS-Department", Value: "Software Development", ValueType: "Edm.String" },
                            { Key: "Manager", Value: "", ValueType: "Edm.String" },
                            { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                            { Key: "PersonalSpace", Value: "/personal/user2_contoso_onmicrosoft_com/", ValueType: "Edm.String" },
                            { Key: "PictureURL", Value: "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-male.png", ValueType: "Edm.String" },
                            { Key: "UserName", Value: "user2@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "QuickLinks", Value: "", ValueType: "Edm.String" },
                            { Key: "WebSite", Value: "", ValueType: "Edm.String" },
                            { Key: "PublicSiteRedirect", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-JobTitle", Value: "Microsoft 365 Developer", ValueType: "Edm.String" },
                            { Key: "SPS-Dotted-line", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Peers", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Responsibility", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-SipAddress", Value: "user2@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-MySiteUpgrade", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ProxyAddresses", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-HireDate", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-DisplayOrder", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimID", Value: "user2@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimProviderID", Value: "membership", ValueType: "Edm.String" },
                            { Key: "SPS-ResourceSID", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ResourceAccountName", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-MasterAccountName", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-UserPrincipalName", Value: "user2@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-O15FirstRunExperience", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteInstantiationState", Value: "2", ValueType: "Edm.String" },
                            { Key: "SPS-DistinguishedName", Value: "CN=00000000-0000-0000-0000-000000000001,OU=dsfsf51…Tenants,OU=MSOnline,DC=SPOCONTOSO,DC=msft,DC=net", ValueType: "Edm.String" },
                            { Key: "SPS-SourceObjectDN", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimProviderType", Value: "Forms", ValueType: "Edm.String" },
                            { Key: "SPS-SavedAccountName", Value: "i:0#.f|membership|user2@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-SavedSID", Value: "System.Byte[]", ValueType: "Edm.String" },
                            { Key: "SPS-ObjectExists", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteCapabilities", Value: "4", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteFirstCreationTime", Value: "2/24/2020 2:04:23 PM", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteLastCreationTime", Value: "2/24/2020 2:04:23 PM", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteNumberOfRetries", Value: "1", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteFirstCreationError", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-FeedIdentifier", Value: "", ValueType: "Edm.String" },
                            { Key: "WorkEmail", Value: "user2@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "CellPhone", Value: "", ValueType: "Edm.String" },
                            { Key: "Fax", Value: "", ValueType: "Edm.String" },
                            { Key: "HomePhone", Value: "", ValueType: "Edm.String" },
                            { Key: "Office", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Location", Value: "", ValueType: "Edm.String" },
                            { Key: "Assistant", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PastProjects", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Skills", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-School", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Birthday", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-StatusNotes", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Interests", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-HashTags", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-EmailOptin", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PrivacyPeople", Value: "True", ValueType: "Edm.String" },
                            { Key: "SPS-PrivacyActivity", Value: "4095", ValueType: "Edm.String" },
                            { Key: "SPS-PictureTimestamp", Value: "63720225810", ValueType: "Edm.String" },
                            { Key: "SPS-PicturePlaceholderState", Value: "0", ValueType: "Edm.String" },
                            { Key: "SPS-PictureExchangeSyncState", Value: "1", ValueType: "Edm.String" },
                            { Key: "SPS-TimeZone", Value: "", ValueType: "Edm.String" },
                            { Key: "OfficeGraphEnabled", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-UserType", Value: "0", ValueType: "Edm.String" },
                            { Key: "SPS-HideFromAddressLists", Value: "False", ValueType: "Edm.String" },
                            { Key: "SPS-RecipientTypeDetails", Value: "", ValueType: "Edm.String" },
                            { Key: "DelveFlags", Value: "", ValueType: "Edm.String" },
                            { Key: "msOnline-ObjectId", Value: "00000000-0000-0000-0000-00000000001", ValueType: "Edm.String" },
                            { Key: "SPS-PointPublishingUrl", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-TenantInstanceId", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-SharePointHomeExperienceState", Value: "134671", ValueType: "Edm.String" },
                            { Key: "SPS-MultiGeoFlags", Value: "", ValueType: "Edm.String" },
                            { Key: "PreferredDataLocation", Value: "", ValueType: "Edm.String" }
                        ] as any);
                        break;
                    case 'i:0#.f|membership|user3@contoso.onmicrosoft.com':
                        resolve([
                            { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000002", ValueType: "Edm.String" },
                            { Key: "SID", Value: "i:0h.f|membership|gferzgfe56g4684@live.com", ValueType: "Edm.String" },
                            { Key: "ADGuid", Value: "System.Byte[]", ValueType: "Edm.String" },
                            { Key: "AccountName", Value: "i:0#.f|membership|user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "FirstName", Value: "Maor", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                            { Key: "LastName", Value: "Sharett", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticLastName", Value: "", ValueType: "Edm.String" },
                            { Key: "PreferredName", Value: "Maor Sharett", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticDisplayName", Value: "", ValueType: "Edm.String" },
                            { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                            { Key: "Department", Value: "", ValueType: "Edm.String" },
                            { Key: "Title", Value: "Microsoft 365 Developer", ValueType: "Edm.String" },
                            { Key: "SPS-Department", Value: "Software Development", ValueType: "Edm.String" },
                            { Key: "Manager", Value: "", ValueType: "Edm.String" },
                            { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                            { Key: "PersonalSpace", Value: "/personal/user3_contoso_onmicrosoft_com/", ValueType: "Edm.String" },
                            { Key: "PictureURL", Value: "", ValueType: "Edm.String" },
                            { Key: "UserName", Value: "user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "QuickLinks", Value: "", ValueType: "Edm.String" },
                            { Key: "WebSite", Value: "", ValueType: "Edm.String" },
                            { Key: "PublicSiteRedirect", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-JobTitle", Value: "Microsoft 365 Developer", ValueType: "Edm.String" },
                            { Key: "SPS-Dotted-line", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Peers", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Responsibility", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-SipAddress", Value: "user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-MySiteUpgrade", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ProxyAddresses", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-HireDate", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-DisplayOrder", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimID", Value: "user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimProviderID", Value: "membership", ValueType: "Edm.String" },
                            { Key: "SPS-ResourceSID", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ResourceAccountName", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-MasterAccountName", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-UserPrincipalName", Value: "user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-O15FirstRunExperience", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteInstantiationState", Value: "2", ValueType: "Edm.String" },
                            { Key: "SPS-DistinguishedName", Value: "CN=00000000-0000-0000-0000-000000000002,OU=dsfsf51…Tenants,OU=MSOnline,DC=SPOCONTOSO,DC=msft,DC=net", ValueType: "Edm.String" },
                            { Key: "SPS-SourceObjectDN", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-ClaimProviderType", Value: "Forms", ValueType: "Edm.String" },
                            { Key: "SPS-SavedAccountName", Value: "i:0#.f|membership|user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "SPS-SavedSID", Value: "System.Byte[]", ValueType: "Edm.String" },
                            { Key: "SPS-ObjectExists", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteCapabilities", Value: "4", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteFirstCreationTime", Value: "2/24/2020 2:04:23 PM", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteLastCreationTime", Value: "2/24/2020 2:04:23 PM", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteNumberOfRetries", Value: "1", ValueType: "Edm.String" },
                            { Key: "SPS-PersonalSiteFirstCreationError", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-FeedIdentifier", Value: "", ValueType: "Edm.String" },
                            { Key: "WorkEmail", Value: "user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "CellPhone", Value: "", ValueType: "Edm.String" },
                            { Key: "Fax", Value: "", ValueType: "Edm.String" },
                            { Key: "HomePhone", Value: "", ValueType: "Edm.String" },
                            { Key: "Office", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Location", Value: "", ValueType: "Edm.String" },
                            { Key: "Assistant", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PastProjects", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Skills", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-School", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Birthday", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-StatusNotes", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-Interests", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-HashTags", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-EmailOptin", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-PrivacyPeople", Value: "True", ValueType: "Edm.String" },
                            { Key: "SPS-PrivacyActivity", Value: "4095", ValueType: "Edm.String" },
                            { Key: "SPS-PictureTimestamp", Value: "63720225810", ValueType: "Edm.String" },
                            { Key: "SPS-PicturePlaceholderState", Value: "0", ValueType: "Edm.String" },
                            { Key: "SPS-PictureExchangeSyncState", Value: "1", ValueType: "Edm.String" },
                            { Key: "SPS-TimeZone", Value: "", ValueType: "Edm.String" },
                            { Key: "OfficeGraphEnabled", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-UserType", Value: "0", ValueType: "Edm.String" },
                            { Key: "SPS-HideFromAddressLists", Value: "False", ValueType: "Edm.String" },
                            { Key: "SPS-RecipientTypeDetails", Value: "", ValueType: "Edm.String" },
                            { Key: "DelveFlags", Value: "", ValueType: "Edm.String" },
                            { Key: "msOnline-ObjectId", Value: "00000000-0000-0000-0000-00000000002", ValueType: "Edm.String" },
                            { Key: "SPS-PointPublishingUrl", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-TenantInstanceId", Value: "", ValueType: "Edm.String" },
                            { Key: "SPS-SharePointHomeExperienceState", Value: "134671", ValueType: "Edm.String" },
                            { Key: "SPS-MultiGeoFlags", Value: "", ValueType: "Edm.String" },
                            { Key: "PreferredDataLocation", Value: "", ValueType: "Edm.String" }
                        ] as any);
                        break;
                    default:
                        resolve(null as any);
                }
            }, 500);
        });
    }

    get siteUrl(): string {
        return this._siteUrl;
    }

    set siteUrl(value: string) {
        if (null == value || value.trim().length < 10) {
            throw new TypeError('SharePoint site URL can not be null or empty.');
        }
        this._siteUrl = value;
    }
}