import { ISiteGroupInfo } from "./ISiteGroupInfo";
import { ISiteUserInfo } from "./ISiteUserInfo";

export default interface ISPGroupService {
    fetchSPGroups(): Promise<Array<ISiteGroupInfo>>;
    getSPGroup(i: number): Promise<ISiteGroupInfo>;
    fetchUsersGroup(i: number): Promise<Array<ISiteUserInfo>>;
    getUserProfile(login: string): Promise<any>;
}