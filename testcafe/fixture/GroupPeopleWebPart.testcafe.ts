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

    await SPPage.editWebPart(0);
});

test('Should discard changes', async t => {
    await SPPage.discardChanges();
});
