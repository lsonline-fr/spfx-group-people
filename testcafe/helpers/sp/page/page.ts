import { t, Selector } from 'testcafe';
import { configLayout, pageLayout, background } from './section/layouts';

/**
 * SharePoint Page Management for TestCafé
 */
class SPPage {

    /**
     * Edit age
     */
    public async edit() {
        await t.click(
            Selector('button')
                .withAttribute('data-automation-id', 'pageCommandBarEditButton')
        );
    }

    /**
     * Save page
     */
    public async saveAsDraft() {
        await t.click(
            Selector('button')
                .withAttribute('data-automation-id', 'pageCommandBarSaveButton')
        );
    }

    /**
     * Publish or Republish page
     */
    public async publish() {
        await t.click(
            Selector('button')
                .withAttribute('data-automation-id', 'pageCommandBarPublishButton')
        );
    }

    /**
     * Discard page changes
     * @param ensure Valide discard changes. Yes by default.
     */
    public async discardChanges(ensure: boolean = true) {
        await t.click(
            Selector('button')
                .withAttribute('data-automation-id', 'discardButton')
        );
        // Alert
        const yesBtn = Selector('button').withAttribute('data-automation-id', 'yesButton');
        const noBtn = Selector('button').withAttribute('data-automation-id', 'noButton');
        if (yesBtn && true == ensure) {
            await t.click(yesBtn);
        } else if (noBtn) {
            await t.click(noBtn);
        }
    }

    //#region Section
    /**
     * Add page section
     * @param p Position of the section
     * @param layout Type of section layout
     */
    public async addSection(p: number, layout: pageLayout) {
        await t
            .click(
                Selector('.CanvasToolboxHint-plusButton').nth(p),
                { offsetX: 10, offsetY: 10 }
            )
            .click(
                Selector('button').withAttribute('data-automation-id', layout)
            );
    }

    /**
     * Edit a section from the page
     * @param s Section number (position into the page)
     */
    public async editSection(s: number) {
        this.selectSection(s);
        await t.click(
            Selector('.CanvasZoneToolbar-sticky')
                .find('button')
                .withAttribute('data-automation-id', 'configureButton')
        );
    }

    /**
     * Remove a section from the page
     * @param s Section number (position into the page)
     */
    public async removeSection(s: number) {
        this.selectSection(s);
        await t.click(
            Selector('.CanvasZoneToolbar-sticky')
                .find('button')
                .withAttribute('data-automation-id', 'deleteButton'),
            { offsetX: 10, offsetY: 10 }
        );
    }

    /**
     * Change selected section layout
     * @param layout Section Layout
     */
    public async setSectionLayout(layout: configLayout) {
        await t.click(
            Selector('div')
                .withAttribute('data-automation-id', 'propertyPanePageContent')
                .find('input')
                .withAttribute('data-automation-id', layout)
        );
    }

    /**
     * Change selected section bacground color
     * @param b Background color
     */
    public async setSectionBackground(b: background) {
        await t.click(
            Selector('div')
                .withAttribute('data-automation-id', 'propertyPanePageContent')
                .find('button')
                .withAttribute('data-automation-id', b)
        );
    }
    //#endregion

    //#region Web Part
    /**
     * Insert a Web Part
     * @param wp Web Part number (position into the page)
     * @param title Title of the Web Part
     */
    public async addWebPart(wp: number, title: string) {
        // Open web part toolbox
        const toolboxHint = Selector('button').withAttribute('data-automation-id', 'toolboxHint-webPart').nth(wp);
        await t
            .hover(toolboxHint)
            .click(toolboxHint);

        // Search web part
        await t.typeText(
            Selector('input').withAttribute('data-automation-id', 'toolbox-searchBox'),
            title
        );

        // Click on web part
        await t.click(
            Selector('div').withAttribute('data-automation-id', 'less-text').nth(0).withText(title)
        );
    }

    /**
     * Open Property panel of a Web Part
     * @param wp Web Part number (position into the page)
     */
    public async editWebPart(wp: number) {
        this.selectWebPart(wp);
        await t.click(
            Selector('.CanvasControlToolbar')
                .find('button')
                .withAttribute('data-automation-id', 'configureButton')
        );
    }

    /**
     * Remove Web Part
     * @param wp Web Part number (position into the page)
     */
    public async removeWebPart(wp: number) {
        this.selectWebPart(wp);
        await t.click(
            Selector('.CanvasControlToolbar').nth(0)
                .find('button')
                .withAttribute('data-automation-id', 'deleteButton')
        );
    }
    //#endregion

    //#region Reusable private methods
    /**
     * Select a page section
     * @param s Section number
     */
    private async selectSection(s: number) {
        await t.click(
            Selector('.CanvasZoneContainer').nth(s), { offsetX: 10, offsetY: 10 }
        );
    }

    /**
     * Select a page Web Part
     * @param wp Web Part number (position into the page)
     */
    private async selectWebPart(wp: number) {
        await t.click(
            Selector('div').withAttribute('data-automation-id', 'CanvasControl').nth(wp)
        );
    }
    //#endregion
}

export default new SPPage();