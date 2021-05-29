import { t } from 'testcafe';

export async function reload() {
    await t.eval(() => location.reload());
};