cfg = {
    WEBSITE_HOSTNAME: process.env.host || '********.ngrok.io',
    BOTAUTH_SECRET: 'hIB0tS3crets',
    AZUREAD_APP_ID:  process.env.azAppId ||'guid-from-apps-dev-microsoft-com',
    AZUREAD_APP_PASSWORD:  process.env.azAppPassword || 'password-from-apps-dev-microsoft-com',
    AZUREAD_APP_REALM: 'common'
}
module.exports = cfg;