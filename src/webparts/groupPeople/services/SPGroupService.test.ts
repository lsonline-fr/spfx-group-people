/// <reference types="jest" />

import SPGroupService from './SPGroupService';
import sinon from 'sinon';

describe('SPGroupService', () => {

    let mySPGroupService: SPGroupService;
    let server: any;

    let getLocalStorageStub = sinon.stub(localStorage, 'getItem');
    let rmLocalStorageStub = sinon.stub(localStorage, 'removeItem');

    beforeEach(() => {
        server = sinon.fakeServer.create();
    });

    afterEach(() => {
        server.restore();
    });

    it('should create a SPGroupService instance with the SharePoint URL', () => {
        mySPGroupService = new SPGroupService('https://contoso.sharepoint.com/sites/project');
        expect(mySPGroupService).toEqual({
            _siteUrl: 'https://contoso.sharepoint.com/sites/project'
        });
    });

    it('should throw an error when creating a SPGroupService instance with an empty parameter', () => {
        expect(() => {
            new SPGroupService(' ')
        }).toThrow(TypeError);
    });

    it('should throw an error when creating a SPGroupService instance with a null parameter', () => {
        let url = null;
        expect(() => {
            new SPGroupService(url);
        }).toThrow(TypeError);
    });

    it('should set a new SPGroupService._siteUrl to the passed argument \'https://contoso.sharepoint.com/sites/project2\'', () => {
        mySPGroupService = new SPGroupService('https://contoso.sharepoint.com/sites/project');
        mySPGroupService.siteUrl = 'https://contoso.sharepoint.com/sites/project2';
        expect(mySPGroupService.siteUrl).toBe('https://contoso.sharepoint.com/sites/project2');
    });

    it('should throw an error when setting a new SPGroupService._siteUrl to the passed argument \'\'', () => {
        expect(() => {
            mySPGroupService.siteUrl = ' ';
        }).toThrow(TypeError);
    });

    it('should throw an error when setting a new SPGroupService._siteUrl to the passed argument \'null\'', () => {
        mySPGroupService = new SPGroupService('https://contoso.sharepoint.com/sites/project');
        let url = null;
        expect(() => {
            mySPGroupService.siteUrl = url;
        }).toThrow(TypeError);
    });

    describe('Fetch SharePoint groups', () => {
        beforeEach(() => {
            mySPGroupService = new SPGroupService('https://contoso.sharepoint.com/sites/project');
        });

        it('should return the list of 4/5 SharePoint groups of the current site collection', (done) => {
            server.respondWith('GET', /\/_api\/Web\/SiteGroups$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": [
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
                        ]
                    }
                })
            ]);

            mySPGroupService.fetchSPGroups().then((grps) => {
                expect(grps.length).toEqual(4);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should handle when no SharePoint group is returned', (done) => {
            server.respondWith('GET', /\/_api\/Web\/SiteGroups$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": null
                })
            ]);

            mySPGroupService.fetchSPGroups().then((grps) => {
                expect(grps).toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return a \'null\' object for the current SharePoint site', (done) => {
            server.respondWith('GET', /\/_api\/Web\/SiteGroups$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": null
                })
            ]);

            mySPGroupService.fetchSPGroups().then((grps) => {
                expect(grps).toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should handle a reject promise', (done) => {
            server.respondWith('GET', /\/_api\/Web\/SiteGroups$/, [
                500, {
                    'Content-Type': 'application/json'
                }, ''
            ]);

            mySPGroupService.fetchSPGroups().catch((c) => {
                try {
                    expect(c).toBeDefined();
                    done();
                } catch (e) {
                    done(e);
                }
            });
            server.respond();
        });
    });

    describe('Get group information', () => {
        beforeEach(() => {
            mySPGroupService = new SPGroupService('https://contoso.sharepoint.com/sites/project');
        });

        afterEach(() => {
            rmLocalStorageStub.restore();
        });

        it('should return the SharePoint group information from the API', (done) => {
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d":
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
                        Title: "Contoso Project - Group People Members"
                    }
                })
            ]);

            mySPGroupService.getSPGroup(5).then((grp) => {
                expect(grp).not.toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the SharePoint group information from the API when cache is incomplete (with no group)', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d":
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
                        Title: "Contoso Project - Group People Members"
                    }
                })
            ]);

            mySPGroupService.getSPGroup(5).then((grp) => {
                expect(grp).not.toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the SharePoint group information from the API when cache is incomplete (with wrong group)', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                group: {
                    "https://contoso.sharepoint.com/sites/otherproject": {
                        5: {
                            AllowMembersEditMembership: false,
                            AllowRequestToJoinLeave: true,
                            AutoAcceptRequestToJoinLeave: false,
                            Description: null,
                            Id: 5,
                            IsHiddenInUI: false,
                            LoginName: "Contoso Project - Group People Members",
                            OnlyAllowMembersViewMembership: false,
                            OwnerTitle: "Contoso Project - Group People Owners",
                            PrincipalType: 8,
                            RequestToJoinLeaveEmailSetting: "",
                            Title: "Contoso Project - Group People Members"
                        }
                    }
                },
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d":
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
                        Title: "Contoso Project - Group People Members"
                    }
                })
            ]);

            mySPGroupService.getSPGroup(5).then((grp) => {
                expect(grp).not.toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the SharePoint group information from API when cache is expired', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                group: {
                    "https://contoso.sharepoint.com/sites/project": {
                        5: {
                            AllowMembersEditMembership: false,
                            AllowRequestToJoinLeave: true,
                            AutoAcceptRequestToJoinLeave: false,
                            Description: null,
                            Id: 5,
                            IsHiddenInUI: false,
                            LoginName: "Contoso Project - Group People Members",
                            OnlyAllowMembersViewMembership: false,
                            OwnerTitle: "Contoso Project - Group People Owners",
                            PrincipalType: 8,
                            RequestToJoinLeaveEmailSetting: "",
                            Title: "Contoso Project - Group People Members"
                        }
                    }
                },
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d":
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
                        Title: "Contoso Project - Group People Members"
                    }
                })
            ]);

            mySPGroupService.getSPGroup(5).then((grp) => {
                expect(grp).not.toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the SharePoint group information from cache', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                group: {
                    "https://contoso.sharepoint.com/sites/project": {
                        5: {
                            AllowMembersEditMembership: false,
                            AllowRequestToJoinLeave: true,
                            AutoAcceptRequestToJoinLeave: false,
                            Description: null,
                            Id: 5,
                            IsHiddenInUI: false,
                            LoginName: "Contoso Project - Group People Members",
                            OnlyAllowMembersViewMembership: false,
                            OwnerTitle: "Contoso Project - Group People Owners",
                            PrincipalType: 8,
                            RequestToJoinLeaveEmailSetting: "",
                            Title: "Contoso Project - Group People Members"
                        }
                    }
                },
                expiry: new Date().getTime() + 60000
            }));

            mySPGroupService.getSPGroup(5).then((grp) => {
                expect(grp).not.toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
        });

        it('should handle when no group information found', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)$/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": null
                })
            ]);

            mySPGroupService.getSPGroup(5).then((grp) => {
                expect(grp).toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should handle a reject promise', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)$/, [
                500, {
                    'Content-Type': 'application/json'
                }, ''
            ]);

            mySPGroupService.getSPGroup(5).catch((c) => {
                try {
                    expect(c).toBeDefined();
                    done();
                } catch (e) {
                    done(e);
                }
            });
            server.respond();
        });
    });

    describe('Fetch group members', () => {
        beforeEach(() => {
            mySPGroupService = new SPGroupService('https://contoso.sharepoint.com/sites/project');
        });

        afterEach(() => {
            rmLocalStorageStub.restore();
        });

        it('should return the list of members (users only) for the specified SharePoint group id from API', (done) => {
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": [
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
                        ]
                    }
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(4);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the list of members (users only) for the specified SharePoint group id from API when cache is incomplete (with no members)', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": [
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
                        ]
                    }
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(4);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the list of members (users only) for the specified SharePoint group id from API when cache is incomplete (with wrong group)', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                members: {
                    "https://contoso.sharepoint.com/sites/project": {
                        3: [
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
                        ]
                    }
                },
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": [
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
                        ]
                    }
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(4);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the list of members (users only) for the specified SharePoint group id from API when cache is expired', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                members: {
                    "https://contoso.sharepoint.com/sites/project": {
                        5: [
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
                        ]
                    }
                },
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": [
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
                        ]
                    }
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(4);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the list of members (users only) for the specified SharePoint group id from cache', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                members: {
                    "https://contoso.sharepoint.com/sites/project": {
                        5: [
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
                        ]
                    }
                },
                expiry: new Date().getTime() + 60000
            }));

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(4);
                done();
            }).catch((c) => {
                done(c);
            });
        });

        it('should return an empty list of members for the specified SharePoint group id that contains one security group only', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": [{
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
                        }]
                    }
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(0);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return an empty list of members for the specified SharePoint group id that contains security group only', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": [
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
                            }
                        ]
                    }
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(0);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return an empty list of members for the specified SharePoint group id', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "results": []
                    }
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users.length).toEqual(0);
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return a \'null\' object for the specified SharePoint group id', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": null
                })
            ]);

            mySPGroupService.fetchUsersGroup(5).then((users) => {
                expect(users).toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should handle a reject promise', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/Web\/SiteGroups\/GetById\([0-9]+\)\/Users/, [
                500, {
                    'Content-Type': 'application/json'
                }, ''
            ]);

            mySPGroupService.fetchUsersGroup(5).catch((c) => {
                try {
                    expect(c).toBeDefined();
                    done();
                } catch (e) {
                    done(e);
                }
            });
            server.respond();
        });
    });

    describe('Get user profile', () => {
        beforeEach(() => {
            mySPGroupService = new SPGroupService('https://contoso.sharepoint.com/sites/project');
        });

        afterEach(() => {
            rmLocalStorageStub.restore();
        });

        it('should throw an error when the passed argument is \'\'', () => {
            expect(() => {
                mySPGroupService.getUserProfile(' ');
            }).toThrow(TypeError);
        });

        it('should throw an error when the passed argument is \'null\'', () => {
            let loginName = null;
            expect(() => {
                mySPGroupService.getUserProfile(loginName);
            }).toThrow(TypeError);
        });

        it('should return the specified user profile properties from API', (done) => {
            server.respondWith('GET', /\/_api\/SP.UserProfiles.PeopleManager\/GetPropertiesFor/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "UserProfileProperties": {
                            "results": [
                                { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000000", ValueType: "Edm.String" },
                                { Key: "AccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                                { Key: "FirstName", Value: "Annie", ValueType: "Edm.String" },
                                { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                                { Key: "LastName", Value: "Lindqvist", ValueType: "Edm.String" },
                                { Key: "PreferredName", Value: "Annie Lindqvist", ValueType: "Edm.String" },
                                { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                                { Key: "Department", Value: "", ValueType: "Edm.String" },
                                { Key: "Title", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                                { Key: "SPS-Department", Value: "Information Technology", ValueType: "Edm.String" },
                                { Key: "Manager", Value: "", ValueType: "Edm.String" },
                                { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                            ]
                        }
                    }
                })
            ]);

            mySPGroupService.getUserProfile('i:0#.f|membership|user1@contoso.onmicrosoft.com').then((p) => {
                expect(p).not.toBeNull();
                let accountName: string = p.find(props => props.Key == 'AccountName') ? p.find(props => props.Key == 'AccountName').Value : '';
                expect(accountName).not.toBeNull();
                expect(accountName).toEqual('i:0#.f|membership|user1@contoso.onmicrosoft.com');
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the specified user profile properties from API when cache is incomplete (with no user)', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/SP.UserProfiles.PeopleManager\/GetPropertiesFor/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "UserProfileProperties": {
                            "results": [
                                { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000000", ValueType: "Edm.String" },
                                { Key: "AccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                                { Key: "FirstName", Value: "Annie", ValueType: "Edm.String" },
                                { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                                { Key: "LastName", Value: "Lindqvist", ValueType: "Edm.String" },
                                { Key: "PreferredName", Value: "Annie Lindqvist", ValueType: "Edm.String" },
                                { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                                { Key: "Department", Value: "", ValueType: "Edm.String" },
                                { Key: "Title", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                                { Key: "SPS-Department", Value: "Information Technology", ValueType: "Edm.String" },
                                { Key: "Manager", Value: "", ValueType: "Edm.String" },
                                { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                            ]
                        }
                    }
                })
            ]);

            mySPGroupService.getUserProfile('i:0#.f|membership|user1@contoso.onmicrosoft.com').then((p) => {
                expect(p).not.toBeNull();
                let accountName: string = p.find(props => props.Key == 'AccountName') ? p.find(props => props.Key == 'AccountName').Value : '';
                expect(accountName).not.toBeNull();
                expect(accountName).toEqual('i:0#.f|membership|user1@contoso.onmicrosoft.com');
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the specified user profile properties from API when cache is incomplete (with wrong user)', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                users: {
                    "i:0#.f|membership|user3@contoso.onmicrosoft.com":
                        [
                            { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000002", ValueType: "Edm.String" },
                            { Key: "AccountName", Value: "i:0#.f|membership|user3@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "FirstName", Value: "Maor", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                            { Key: "LastName", Value: "Sharett", ValueType: "Edm.String" },
                            { Key: "PreferredName", Value: "Maor Sharett", ValueType: "Edm.String" },
                            { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                            { Key: "Department", Value: "", ValueType: "Edm.String" },
                            { Key: "Title", Value: "Microsoft 365 Developer", ValueType: "Edm.String" },
                            { Key: "SPS-Department", Value: "Software Development", ValueType: "Edm.String" },
                            { Key: "Manager", Value: "", ValueType: "Edm.String" },
                            { Key: "AboutMe", Value: "", ValueType: "Edm.String" }
                        ]
                },
                expiry: new Date().getTime() + 60000
            }));
            server.respondWith('GET', /\/_api\/SP.UserProfiles.PeopleManager\/GetPropertiesFor/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "UserProfileProperties": {
                            "results": [
                                { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000000", ValueType: "Edm.String" },
                                { Key: "AccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                                { Key: "FirstName", Value: "Annie", ValueType: "Edm.String" },
                                { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                                { Key: "LastName", Value: "Lindqvist", ValueType: "Edm.String" },
                                { Key: "PreferredName", Value: "Annie Lindqvist", ValueType: "Edm.String" },
                                { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                                { Key: "Department", Value: "", ValueType: "Edm.String" },
                                { Key: "Title", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                                { Key: "SPS-Department", Value: "Information Technology", ValueType: "Edm.String" },
                                { Key: "Manager", Value: "", ValueType: "Edm.String" },
                                { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                            ]
                        }
                    }
                })
            ]);

            mySPGroupService.getUserProfile('i:0#.f|membership|user1@contoso.onmicrosoft.com').then((p) => {
                expect(p).not.toBeNull();
                let accountName: string = p.find(props => props.Key == 'AccountName') ? p.find(props => props.Key == 'AccountName').Value : '';
                expect(accountName).not.toBeNull();
                expect(accountName).toEqual('i:0#.f|membership|user1@contoso.onmicrosoft.com');
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the specified user profile properties from API when cache is expired', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                users: {
                    "i:0#.f|membership|user1@contoso.onmicrosoft.com":
                        [
                            { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000000", ValueType: "Edm.String" },
                            { Key: "AccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "FirstName", Value: "Annie", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                            { Key: "LastName", Value: "Lindqvist", ValueType: "Edm.String" },
                            { Key: "PreferredName", Value: "Annie Lindqvist", ValueType: "Edm.String" },
                            { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                            { Key: "Department", Value: "", ValueType: "Edm.String" },
                            { Key: "Title", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                            { Key: "SPS-Department", Value: "Information Technology", ValueType: "Edm.String" },
                            { Key: "Manager", Value: "", ValueType: "Edm.String" },
                            { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                        ]
                },
                expiry: new Date().getTime() - 60000
            }));
            server.respondWith('GET', /\/_api\/SP.UserProfiles.PeopleManager\/GetPropertiesFor/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
                        "UserProfileProperties": {
                            "results": [
                                { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000000", ValueType: "Edm.String" },
                                { Key: "AccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                                { Key: "FirstName", Value: "Annie", ValueType: "Edm.String" },
                                { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                                { Key: "LastName", Value: "Lindqvist", ValueType: "Edm.String" },
                                { Key: "PreferredName", Value: "Annie Lindqvist", ValueType: "Edm.String" },
                                { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                                { Key: "Department", Value: "", ValueType: "Edm.String" },
                                { Key: "Title", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                                { Key: "SPS-Department", Value: "Information Technology", ValueType: "Edm.String" },
                                { Key: "Manager", Value: "", ValueType: "Edm.String" },
                                { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                            ]
                        }
                    }
                })
            ]);

            mySPGroupService.getUserProfile('i:0#.f|membership|user1@contoso.onmicrosoft.com').then((p) => {
                expect(p).not.toBeNull();
                let accountName: string = p.find(props => props.Key == 'AccountName') ? p.find(props => props.Key == 'AccountName').Value : '';
                expect(accountName).not.toBeNull();
                expect(accountName).toEqual('i:0#.f|membership|user1@contoso.onmicrosoft.com');
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should return the specified user profile properties from cache', (done) => {
            getLocalStorageStub.returns(JSON.stringify({
                users: {
                    "i:0#.f|membership|user1@contoso.onmicrosoft.com":
                        [
                            { Key: "UserProfile_GUID", Value: "00000000-0000-0000-0000-000000000000", ValueType: "Edm.String" },
                            { Key: "AccountName", Value: "i:0#.f|membership|user1@contoso.onmicrosoft.com", ValueType: "Edm.String" },
                            { Key: "FirstName", Value: "Annie", ValueType: "Edm.String" },
                            { Key: "SPS-PhoneticFirstName", Value: "", ValueType: "Edm.String" },
                            { Key: "LastName", Value: "Lindqvist", ValueType: "Edm.String" },
                            { Key: "PreferredName", Value: "Annie Lindqvist", ValueType: "Edm.String" },
                            { Key: "WorkPhone", Value: "4250000000", ValueType: "Edm.String" },
                            { Key: "Department", Value: "", ValueType: "Edm.String" },
                            { Key: "Title", Value: "Microsoft 365 Architect", ValueType: "Edm.String" },
                            { Key: "SPS-Department", Value: "Information Technology", ValueType: "Edm.String" },
                            { Key: "Manager", Value: "", ValueType: "Edm.String" },
                            { Key: "AboutMe", Value: "", ValueType: "Edm.String" },
                        ]
                },
                expiry: new Date().getTime() + 60000
            }));

            mySPGroupService.getUserProfile('i:0#.f|membership|user1@contoso.onmicrosoft.com').then((p) => {
                expect(p).not.toBeNull();
                let accountName: string = p.find(props => props.Key == 'AccountName') ? p.find(props => props.Key == 'AccountName').Value : '';
                expect(accountName).not.toBeNull();
                expect(accountName).toEqual('i:0#.f|membership|user1@contoso.onmicrosoft.com');
                done();
            }).catch((c) => {
                done(c);
            });
        });

        it('should handle an user without UserProfileProperties (external user basicaly)', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/SP.UserProfiles.PeopleManager\/GetPropertiesFor/, [
                200, {
                    'Content-Type': 'application/json'
                }, JSON.stringify({
                    "d": {
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
                })
            ]);

            mySPGroupService.getUserProfile('i:0#.f|membership|urn%3aspo%3aguest#external@outlook.com').then((p) => {
                expect(p).toBeNull();
                done();
            }).catch((c) => {
                done(c);
            });
            server.respond();
        });

        it('should handle a reject promise', (done) => {
            getLocalStorageStub.returns(null);
            server.respondWith('GET', /\/_api\/SP.UserProfiles.PeopleManager\/GetPropertiesFor/, [
                500, {
                    'Content-Type': 'application/json'
                }, ''
            ]);

            mySPGroupService.getUserProfile('i:0#.f|membership|urn%3aspo%3aguest#external@outlook.com').catch((c) => {
                try {
                    expect(c).toBeDefined();
                    done();
                } catch (e) {
                    done(e);
                }
            });
            server.respond();
        });
    });
});