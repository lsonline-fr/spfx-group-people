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
    const groupsDd = Selector('#spPropertyPaneContainer .ms-Dropdown').nth(0);
    await t.click(groupsDd);
    const groupsDdId = await groupsDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${groupsDdId}-list1`));
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
    await t.click(Selector('button').withAttribute('id', `${groupsDdId}-list2`));
    const personaTiles = Selector('div[class*=personaTile_]').exists;
    await t.expect(personaTiles).ok();
});

test('Should display regular (48) picture size with 2 fields by default', async t => {
    const regularSizePicture = Selector('.ms-Persona--size48').exists;
    await t.expect(regularSizePicture).ok();
    const visibleDetails = Selector('.ms-Persona--size48').nth(0).find('.ms-Persona-details>div').filterVisible().count;
    await t.expect(visibleDetails).eql(2);
});

test('Should change picture size to \'Large\' (72)', async t => {
    const sizeDd = Selector('#spPropertyPaneContainer .ms-Dropdown').nth(1);
    await t.click(sizeDd);
    const sizeDdId = await sizeDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${sizeDdId}-list1`));
    const largeSizePicture = Selector('.ms-Persona--size72').exists;
    await t.expect(largeSizePicture).ok();
});

test('Should display 3 detail fields when picture size set to \'Large\' (72)', async t => {
    const visibleDetails = Selector('.ms-Persona--size72').nth(0).find('.ms-Persona-details>div').filterVisible().count;
    await t.expect(visibleDetails).eql(3);
});

test('Should change picture size to \'Extralarge\' (100)', async t => {
    const sizeDd = Selector('#spPropertyPaneContainer .ms-Dropdown').nth(1);
    await t.click(sizeDd);
    const sizeDdId = await sizeDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${sizeDdId}-list2`));
    const largeSizePicture = Selector('.ms-Persona--size100').exists;
    await t.expect(largeSizePicture).ok();
});

// User Profile Properties available: /_api/SP.UserProfiles.PeopleManager/GetMyProperties
test('Should display user profile picture', async t => {
    const personaImage = Selector('div[class*=personaTile_] .ms-Persona-image').exists;
    await t.expect(personaImage).ok();
});

test('Should display user initials when no picture specified', async t => {
    await t.
        selectText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(1).find('input'))
        .pressKey('delete');
    const personaInitials = Selector('div[class*=personaTile_] .ms-Persona-initials').exists;
    await t.expect(personaInitials).ok();
});

test('Should display user email at \'Line 1\'', async t => {
    await t.typeText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(2).find('input'), 'UserName', { replace: true });
    const primaryText = Selector('div[class*=personaTile_]').nth(0).find('.ms-Persona-primaryText > .ms-TooltipHost').nth(0).textContent;
    await t.expect(/^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/.test(await primaryText)).eql(true);
});

test('Should display user email at \'Line 2\'', async t => {
    await t.typeText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(3).find('input'), 'UserName', { replace: true });
    const secondaryText = Selector('div[class*=personaTile_]').nth(0).find('.ms-Persona-secondaryText > .ms-TooltipHost').nth(0).textContent;
    await t.expect(/^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/.test(await secondaryText)).eql(true);
});

test('Should display user email at \'Line 3\'', async t => {
    await t.typeText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(4).find('input'), 'UserName', { replace: true });
    const tertiaryText = Selector('div[class*=personaTile_]').nth(0).find('.ms-Persona-tertiaryText > .ms-TooltipHost').nth(0).textContent;
    await t.expect(/^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/.test(await tertiaryText)).eql(true);
});

test('Should display an empty \'Line 3\' when no property defined', async t => {
    await t.
        selectText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(4).find('input'))
        .pressKey('delete');
    const tertiaryText = Selector('div[class*=personaTile_]').nth(0).find('.ms-Persona-tertiaryText').textContent;
    await t.expect(tertiaryText).eql('');
});

test('Should display an empty \'Line 2\' when no property defined', async t => {
    await t.
        selectText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(3).find('input'))
        .pressKey('delete');
    const secondaryText = Selector('div[class*=personaTile_]').nth(0).find('.ms-Persona-secondaryText').textContent;
    await t.expect(secondaryText).eql('');
});

test('Should display an empty \'Line 1\' when no property defined', async t => {
    await t.
        selectText(Selector('#spPropertyPaneContainer .ms-TextField-wrapper').nth(2).find('input'))
        .pressKey('delete');
    const primaryText = Selector('div[class*=personaTile_]').nth(0).find('.ms-Persona-primaryText').textContent;
    await t.expect(primaryText).eql('');
});

test('Should display Web Part and default message when no people existing into the selected group', async t => {
    const groupsDd = Selector('#spPropertyPaneContainer .ms-Dropdown').child(0);
    await t.click(groupsDd);
    const groupsDdId = await groupsDd.getAttribute('id');
    await t.click(Selector('button').withAttribute('id', `${groupsDdId.replace('-option', '')}-list1`));
    
    await SPPage.saveAsDraft();
    const displayedWP = Selector('div[class^=groupPeople_]').filterVisible().count;
    await t.expect(displayedWP).eql(1);
});

test('Should hide Web Part when no people existing into the selected group', async t => {
    await SPPage.edit();
    await SPPage.editWebPart(0);
    
    const hideWP = Selector('#spPropertyPaneContainer .ms-Checkbox').nth(0);
    await t.click(hideWP);

    await SPPage.saveAsDraft();
    const displayedWP = Selector('div[class^=groupPeople_]').filterVisible().count;
    await t.expect(displayedWP).eql(0);
});

test('Should discard changes', async t => {
    await SPPage.edit();
    await SPPage.removeWebPart(0);
    await SPPage.publish();
});