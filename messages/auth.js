const cfg = require('./config');
const WEBSITE_HOSTNAME = cfg.WEBSITE_HOSTNAME;
const BOTAUTH_SECRET = cfg.BOTAUTH_SECRET;
const AZUREAD_APP_ID = cfg.AZUREAD_APP_ID;
const AZUREAD_APP_PASSWORD = cfg.AZUREAD_APP_PASSWORD;
const AZUREAD_APP_REALM = cfg.AZUREAD_APP_REALM;

const botauth = require('botauth');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
// const envx = require("envx");

// Environment information
// const WEBSITE_HOSTNAME = envx("WEBSITE_HOSTNAME");
// const PORT = envx("PORT", 3978);
// const BOTAUTH_SECRET = envx("BOTAUTH_SECRET");


//oauth details for AAD
// const AZUREAD_APP_ID = envx("AZUREAD_APP_ID");
// const AZUREAD_APP_PASSWORD = envx("AZUREAD_APP_PASSWORD");
// const AZUREAD_APP_REALM = envx("AZUREAD_APP_REALM");

const AuthHelper = {
    configure: function(server, bot) {
        var ba = new botauth.BotAuthenticator(server, bot, {
            session: true,
            baseUrl: `https://${WEBSITE_HOSTNAME}`,
            secret : BOTAUTH_SECRET,
            successRedirect: '/code'
        });
        ba.provider("aadv2", (options) => {
            // Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
            // For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
            let oidStrategyv2 = {
                redirectUrl: options.callbackURL, //  redirect: /botauth/aadv2/callback
                realm: AZUREAD_APP_REALM,
                clientID: AZUREAD_APP_ID,
                clientSecret: AZUREAD_APP_PASSWORD,
                identityMetadata: 'https://login.microsoftonline.com/' + AZUREAD_APP_REALM + '/v2.0/.well-known/openid-configuration',
                skipUserProfile: false,
                validateIssuer: false, // true causes us to fail....
                responseType: 'code',
                responseMode: 'query',
                scope: ['Calendars.ReadWrite', 'User.Read', 'offline_access', 'https://graph.microsoft.com/mail.read'],
                passReqToCallback: true
            };

            let strategy = oidStrategyv2;
            return new OIDCStrategy(strategy,
                (req, iss, sub, profile, accessToken, refreshToken, done) => {
                if (!profile.displayName) {
                    return done(new Error("No oid found"), null);
                }
                profile.accessToken = accessToken;
                profile.refreshToken = refreshToken;
                done(null, profile);
            });
        });

        return ba;
    }
}
module.exports = AuthHelper;