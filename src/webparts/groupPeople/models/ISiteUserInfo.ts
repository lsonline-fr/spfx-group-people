/**
 * @see https://github.com/pnp/pnpjs/blob/version-2/packages/sp/types.ts
 */
export declare const enum PrincipalType {
    /**
     * Enumeration whose value specifies no principal type.
     */
    None = 0,
    /**
     * Enumeration whose value specifies a user as the principal type.
     */
    User = 1,
    /**
     * Enumeration whose value specifies a distribution list as the principal type.
     */
    DistributionList = 2,
    /**
     * Enumeration whose value specifies a security group as the principal type.
     */
    SecurityGroup = 4,
    /**
     * Enumeration whose value specifies a group as the principal type.
     */
    SharePointGroup = 8,
    /**
     * Enumeration whose value specifies all principal types.
     */
    All = 15
}

/**
 * @see https://github.com/pnp/pnpjs/blob/version-2/packages/sp/site-users/types.ts
 */
export interface ISiteUserInfo {
    Email: string;
    Id: number;
    IsHiddenInUI: boolean;
    IsShareByEmailGuestUser: boolean;
    IsSiteAdmin: boolean;
    LoginName: string;
    PrincipalType: number | PrincipalType;
    Title: string;
    Expiration: string;
    IsEmailAuthenticationGuestUser: boolean;
    UserId: {
        NameId: string;
        NameIdIssuer: string;
    };
    UserPrincipalName: string | null;
}