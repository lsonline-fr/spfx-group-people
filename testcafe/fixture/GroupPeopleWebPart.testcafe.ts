import { Selector } from 'testcafe';
import SPPage from '../helpers/sp/page/page';
import { configLayout } from '../helpers/sp/page/section/layouts';
import { msalAuth } from '../roles/auth/msal';
import Config from '../config';

const PAGE_URL = '/sites/spfx-grppeople/SitePages/unit-tests.aspx';

fixture
    .disablePageReloads('SharePoint Group People')
    .page(new URL(PAGE_URL, Config.get().var('sp_baseurl')).toString());

test('Navigate to the SharePoint Site Page', async t => {
    await t.navigateTo(new URL(PAGE_URL, Config.get().var('sp_baseurl')).toString());
}).before(async t => {
    await t.useRole(msalAuth)
});

test('Should Web Part exists in the SharePoint picker', async t => {
    await SPPage.edit();
    await SPPage.addWebPart(0, 'SharePoint Group People');
    const wpCount = Selector('.ControlZone').count;
    await t.expect(wpCount).gt(0);
});

test('Should init Web Part by changing the SharePoint group property', async t => {
    await SPPage.editWebPart(0);
    const groupsDd = Selector('#spPropertyPaneContainer .ms-Dropdown').child(0);
    await t.click(groupsDd);
    const groupsDdId = await groupsDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${groupsDdId.replace('-option', '')}-list1`));
    const noItem = Selector('.grpPeopleNoItem').exists;
    await t.expect(noItem).ok();
});

test('Should change the Web Part title', async t => {
    await t.typeText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(0).find('input'), 'New title');
    await t.expect(Selector('div[class^=groupPeople]').nth(0).find('h2').nth(0).textContent).eql('New title');
    await t.
        selectText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(0).find('input'))
        .pressKey('delete');
});

test('Should hide the Web Part title', async t => {
    await t.click(Selector('#spPropertyPaneContainer .ms-Toggle').nth(0).find('label').nth(0));
    const hiddenTitle = Selector('div[class^=groupPeople]').nth(0).find('h2[class*=hidden_]').nth(0).exists;
    await t.expect(hiddenTitle).ok();
});

test('Should show the Web Part title', async t => {
    await t.click(Selector('#spPropertyPaneContainer .ms-Toggle').nth(0).find('label').nth(0));
    const hiddenTitle = Selector('div[class^=groupPeople]').nth(0).find('h2:not([class*=hidden_])').nth(0).exists;
    await t.expect(hiddenTitle).ok();
});

test('Should display people inside a no blank SharePoint group', async t => {
    const groupsDd = Selector('#spPropertyPaneContainer .ms-Dropdown').nth(0);
    await t.click(groupsDd);
    const groupsDdId = await groupsDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${groupsDdId.replace('-option', '')}-list2`));
    const personaTiles = Selector('div[class*=personaTile_]').exists;
    await t.expect(personaTiles).ok();
});

test('Should display regular (48) picture size with 2 fields by default', async t => {
    const regularSizePicture = Selector('.ms-Persona--size48').exists;
    await t.expect(regularSizePicture).ok();
    const visibleDetails = Selector('.ms-Persona--size48:first-child .ms-Persona-details>div').filterVisible().count;
    await t.expect(visibleDetails).eql(2);
});

test('Should change picture size to \'Large\' (72)', async t => {
    const sizeDd = Selector('#spPropertyPaneContainer .ms-Dropdown').nth(1);
    await t.click(sizeDd);
    const sizeDdId = await sizeDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${sizeDdId.replace('-option', '')}-list1`));
    const largeSizePicture = Selector('.ms-Persona--size72').exists;
    await t.expect(largeSizePicture).ok();
});

test('Should display 3 detail fields when picture size set to \'Large\' (72)', async t => {
    const visibleDetails = Selector('.ms-Persona--size72:first-child .ms-Persona-details>div').filterVisible().count;
    await t.expect(visibleDetails).eql(3);
});

test('Should change picture size to \'Extralarge\' (100)', async t => {
    const sizeDd = Selector('#spPropertyPaneContainer .ms-Dropdown').nth(1);
    await t.click(sizeDd);
    const sizeDdId = await sizeDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${sizeDdId.replace('-option', '')}-list2`));
    const largeSizePicture = Selector('.ms-Persona--size100').exists;
    await t.expect(largeSizePicture).ok();
});

test('Should discard changes', async t => {
    await SPPage.discardChanges();
});