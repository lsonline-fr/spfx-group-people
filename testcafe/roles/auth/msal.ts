import { Role, Selector } from 'testcafe';
import Config from '../../config';

const mslogin_url = 'https://login.microsoftonline.com';

export const msalAuth = Role(mslogin_url, async t => {
    await t
        .typeText(
            Selector('input').withAttribute('type', 'email'),
            Config.get().var('sp_username')
        )
        .click(Selector('#idSIButton9'))
        .typeText(
            Selector('input').withAttribute('type', 'password'),
            Config.get().var('sp_password')
        )
        .click(Selector('#idSIButton9'))
        .click(Selector('#idBtn_Back'));
}, { preserveUrl: true });
