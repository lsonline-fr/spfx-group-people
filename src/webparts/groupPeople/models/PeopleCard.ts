/**
 * People Card
 * @class
 */
export default class PeopleCard {

    private _loginName: string;
    private _image: string;

    private _lineOne: string;
    private _lineTwo: string;
    private _lineThree: string;

    /**
     * @constructor
     * @param loginName Login Name of the user
     * @param img User Picture
     * @param fl Primary text to display, usually the name of the person.
     * @param sl Secondary text to display, usually the role of the user.
     * @param tl Tertiary text to display, usually the status of the user. The tertiary text will only be shown when using size72 or size100.
     * ```typescript
     * new PeopleCard(
     *     'i:0#.f|membership|user1@contoso.onmicrosoft.com',
     *     'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-female.png',
     *     'Annie Lindqvist',
     *     'Microsoft 365 Architect',
     *     'Information Technology');
     * ```
     * @throws If login name is null or empty
     */
    constructor(loginName: string, img: string = null, fl: string = null, sl: string = null, tl: string = null) {
        if (undefined === loginName || null == loginName || loginName.trim().length < 1) {
            throw new TypeError('Login name can not be null or empty.');
        }
        this._loginName = loginName;
        this._image = img;
        this._lineOne = fl;
        this._lineTwo = sl;
        this._lineThree = tl;
    }

    //#region Getters / Setters
    /**
     * Get unique key of the PeopleCard based on the login name
     * @returns Unique Persona Key
     */
    get key(): number {
        return this.hashKey();
    }

    /**
     * Get the login name of the user
     */
    get loginName(): string {
        return this._loginName;
    }

    /**
     * Get the user profile picture based on the user profile properties
     * @returns URL of the profile picture
     */
    get image(): string {
        return this._image;
    }

    /**
     * Set an user profile picture
     */
    set image(value: string) {
        this._image = value;
    }

    /**
     * Get the content for the first line of the persona control
     */
    get lineOne(): string {
        return this._lineOne;
    }

    /**
     * Set the content for the first line of the persona control
     */
    set lineOne(value: string) {
        this._lineOne = value;
    }

    /**
     * Get the content for the second line of the persona control
     */
    get lineTwo(): string {
        return this._lineTwo;
    }

    /**
     * Set the content for the second line of the persona control
     */
    set lineTwo(value: string) {
        this._lineTwo = value;
    }

    /**
     * Get the content for the third line of the persona control
     */
    get lineThree(): string {
        return this._lineThree;
    }

    /**
     * Set the content for the third line of the persona control
     */
    set lineThree(value: string) {
        this._lineThree = value;
    }
    //#endregion

    /**
     * Create an unique key that can be used by the render
     * @returns the loginName property hashed
     */
    private hashKey(): number {
        var hash = 0;
        for (var i = 0; i < this._loginName.length; i++) {
            var char = this._loginName.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash;
        }
        return hash;
    }
}