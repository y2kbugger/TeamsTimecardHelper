export const config = Object.freeze({
    clientId: '80dd4c35-6e75-4ec5-96ee-d97f88d077d8',
    authorityTenant: 'organizations',
    scopes: Object.freeze([
        'Team.ReadBasic.All',
        'Schedule.ReadWrite.All',
    ]),
    appPagePath: './index.html',
    loginPagePath: './login.html',
    // Both must be registered redirect URIs of the app (Azure portal, Single-page application),
    // on localhost:8999 and on the production host.
    authDonePagePath: './auth-done.html',
});
