import ISPGroupService from "../models/ISPGroupService";

import { ISiteGroupInfo } from "../models/ISiteGroupInfo";
import { ISiteUserInfo } from "../models/ISiteUserInfo";

/**
 * TTL used to simulate a local storage expiration (5 min)
 */
const CACHEEXPIRACY: number = 30000;

/**
 * Local Storage Key to store My Properties
 */
const CACHEKEY: string = 'grpPeopleWPSPGrpSvc';

/**
 * SharePoint Group Service
 * 
 * REST API
 * @implements ISPGroupService
 * @class
 */
export default class SPGroupService implements ISPGroupService {

    /**
     * SharePoint site URL used to perform the REST API requests
     */
    private _siteUrl: string;

    /**
     * @constructor
     * @param u SharePoint site URL
     * ```typescript
     * new SPGroupService(this.context.pageContext.site.absoluteUrl);
     * ```
     * @throws URL cannot be null or empty
     */
    constructor(u: string) {
        if (null == u || u.trim().length < 10) {
            throw new TypeError('SharePoint site URL can not be null or empty.');
        }
        this._siteUrl = u;
    }

    /** Get all SharePoint Groups
     * This function exclude 'SharingLinks' groups
     * @return SharePoint groups
     */
    public fetchSPGroups(): Promise<Array<ISiteGroupInfo>> {
        const apiUrl = this._siteUrl + '/_api/Web/SiteGroups';
        return new Promise((resolve, reject) => {
            this.sendRequest(apiUrl).then((r) => {
                resolve(r ? r.results.filter((g: ISiteGroupInfo) => { return !/^SharingLinks./.test(g.LoginName); }) : null);
            }).catch((e) => { reject(e); });
        });
    }

    /**
     * Get group informations
     * @param i SharePoint group ID
     * @returns SharePoint group information
     */
    public getSPGroup(i: number): Promise<ISiteGroupInfo> {
        const apiUrl = this._siteUrl + '/_api/Web/SiteGroups/GetById(' + i + ')';
        return new Promise((resolve, reject) => {
            if (this.getGroupCache(i)) {
                resolve(this.getGroupCache(i));
            } else {
                this.sendRequest(apiUrl).then((r) => {
                    const result: ISiteGroupInfo = r ? r : null;
                    this.setGroupCache(i, result);
                    resolve(result);
                }).catch((e) => { reject(e); });
            }
        });
    }

    /** Get members of selected SharePoint group
    * This function ensure that users have PrincipalType to 1 and an email
     * @param i SharePoint site collection User ID
     * @return SharePoint Group Members
     * @see https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-csom/ee541430(v=office.15)
     */
    public fetchUsersGroup(i: number): Promise<Array<ISiteUserInfo>> {
        const apiUrl = this._siteUrl + '/_api/Web/SiteGroups/GetById(' + i + ')/Users?$select=Email,LoginName,PrincipalType';
        return new Promise((resolve, reject) => {
            if (this.getUsersGroupCache(i)) {
                resolve(this.getUsersGroupCache(i));
            } else {
                this.sendRequest(apiUrl).then((r) => {
                    // PrincipalType.User = 1 (SP.User)
                    const result = r ? r.results.filter((u: ISiteUserInfo) => { return u.PrincipalType == 1 && u.Email != null && u.Email.length > 0; }) : null;
                    this.setUsersGroupCache(i, result);
                    resolve(result);
                }).catch((e) => { reject(e); });
            }
        });
    }

    /** Get User Profile specified by his LoginName
     * @param login LoginName of user
     * @return User Profile Properties
     * @throws The login cannot be null or empty
     */
    public getUserProfile(login: string): Promise<any> {
        if (undefined === login || null == login || login.trim().length < 1) {
            throw new TypeError('Login can not be null or empty.');
        }
        const apiUrl = this._siteUrl + `/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(@v)?@v='${encodeURIComponent(login)}'`;
        return new Promise((resolve, reject) => {
            if (this.getUsersProfileCache(login)) {
                resolve(this.getUsersProfileCache(login));
            } else {
                this.sendRequest(apiUrl).then((r) => {
                    const result = (r && r.UserProfileProperties) ? r.UserProfileProperties.results : null;
                    this.setUsersProfileCache(login, result);
                    resolve(result);
                }).catch((e) => { reject(e); });
            }
        });
    }

    /**
     * Get the site URL used to perform the REST API requests
     */
    get siteUrl(): string {
        return this._siteUrl;
    }

    /**
     * Set the site URL used to perform the REST API requests
     * @throws Site URL cannot be null or empty
     */
    set siteUrl(value: string) {
        if (null == value || value.trim().length < 10) {
            throw new TypeError('SharePoint site URL can not be null or empty.');
        }
        this._siteUrl = value;
    }

    /**
     * Send SharePoint REST API requests
     * @param u URL of the API REST
     * @returns THe SharePoint REST API results
     */
    private sendRequest(u: string): Promise<any> {
        return new Promise((resolve, reject) => {
            var xhr = new XMLHttpRequest();
            /* Pass function parameters to control XHR properties */
            xhr.open('GET', u, true);
            xhr.setRequestHeader("Accept", "application/json;odata=verbose");

            xhr.onreadystatechange = () => {
                if (xhr.readyState == 4) { // `DONE`
                    var status = xhr.status;
                    if (status === 0 || (status >= 200 && status < 400)) {
                        resolve(JSON.parse(xhr.responseText).d);
                    } else {
                        reject({ status: xhr.status, statusText: xhr.statusText, responseText: xhr.responseText });
                    }
                }
            };
            xhr.send();
        });
    }

    /**
     * Get SharePoint group information from cache
     * @param i SharePoint group ID
     * @returns SharePoint group information if exist else return null
     */
    private getGroupCache(i: number): ISiteGroupInfo {
        let c = this.getCache();
        if (c && c['group'] && c['group'][this._siteUrl] && c['group'][this._siteUrl][i]) {
            return c['group'][this._siteUrl][i];
        } else {
            return null;
        }
    }

    /**
     * Set group information into the cache
     * @param id SharePoint group ID
     * @param r SharePoint group information
     */
    private setGroupCache(id: number, r: ISiteGroupInfo) {
        let c = this.getCache();
        if (c && c['group']) {
            c['group'][this._siteUrl] = {};
            c['group'][this._siteUrl][id] = r;
        }
        else if (c && !c['group']) {
            c['group'] = {};
            c['group'][this._siteUrl] = {};
            c['group'][this._siteUrl][id] = r;
        } else {
            c = {};
            c['group'] = {};
            c['group'][this._siteUrl] = {};
            c['group'][this._siteUrl][id] = r;
            c['expiry'] = new Date().getTime() + CACHEEXPIRACY;
        }
        localStorage.setItem(CACHEKEY, JSON.stringify(c));
    }

    /**
     * Get the users of a SharePoint group from the cache
     * @param i SharePoint group ID
     * @returns List of users if exist else return null
     */
    private getUsersGroupCache(i: number): Array<ISiteUserInfo> {
        let c = this.getCache();
        if (c && c['members'] && c['members'][this._siteUrl] && c['members'][this._siteUrl][i]) {
            return c['members'][this._siteUrl][i];
        } else {
            return null;
        }
    }

    /**
     * Set the users of the group into the cache
     * @param id SharePoint group ID
     * @param r List of users
     */
    private setUsersGroupCache(id: number, r: Array<ISiteUserInfo>) {
        let c = this.getCache();
        if (c && c['members']) {
            c['members'][this._siteUrl] = {};
            c['members'][this._siteUrl][id] = r;
        }
        else if (c && !c['members']) {
            c['members'] = {};
            c['members'][this._siteUrl] = {};
            c['members'][this._siteUrl][id] = r;
        } else {
            c = {};
            c['members'] = {};
            c['members'][this._siteUrl] = {};
            c['members'][this._siteUrl][id] = r;
            c['expiry'] = new Date().getTime() + CACHEEXPIRACY;
        }
        localStorage.setItem(CACHEKEY, JSON.stringify(c));
    }

    /**
     * Get the user profile from the cache
     * @param l User login name
     * @returns User profile information else null
     */
    private getUsersProfileCache(l: string): Array<ISiteUserInfo> {
        let c = this.getCache();
        if (c && c['users'] && c['users'][l]) {
            return c['users'][l];
        } else {
            return null;
        }
    }

    /**
     * Set the users profile into the cache
     * @param l User login name
     * @param r User profile
     */
    private setUsersProfileCache(l: string, r: any) {
        let c = this.getCache();
        if (c && c['users']) {
            c['users'][l] = r;
        }
        else if (c && !c['users']) {
            c['users'] = {};
            c['users'][l] = r;
        } else {
            c = {};
            c['users'] = {};
            c['users'][l] = r;
            c['expiry'] = new Date().getTime() + CACHEEXPIRACY;
        }
        localStorage.setItem(CACHEKEY, JSON.stringify(c));
    }

    /**
     * Get value from the local storage
     * If the expiracy value is outdated, the local storage key is removed and return null (to simulate an expiration)
     */
    private getCache(): Object {
        const cacheProps = localStorage.getItem(CACHEKEY);
        if (!cacheProps) {
            return null;
        }
        const props = JSON.parse(cacheProps);
        if (new Date().getTime() > props.expiry) {
            // If the item is expired, delete the item from storage
            // and return null
            localStorage.removeItem(CACHEKEY);
            return null;
        }
        return props;
    }
}