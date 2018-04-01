const path = require('path');
const restify = require('restify');
const expressSession = require('express-session');
const builder = require('botbuilder');
const azure = require('botbuilder-azure');
const azureStorage = require('azure-storage');
const AuthHelper = require('./auth');
const cfg = require('./config');
const useEmulator = (process.env.NODE_ENV == 'development');
/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework.
-----------------------------------------------------------------------------*/
'use strict';

const storageConnectionString = process.env.storageConnectionString;


var documentDbOptions = {
    host: process.env.docDbHost,
    masterKey: process.env.docDbKey,
    database: process.env.docDbName || 'botdocdb',
    collection: process.env.docDbCollection ||'botdata'
};

//This is an old version of the node docDbClient, the new version is not compatible with the AzureBotStorage library.
var botDocDbClient = new azure.DocumentDbClient(documentDbOptions);
var tableStorage = new azure.AzureBotStorage({ gzipData: false }, botDocDbClient);


// Create chat connector for communicating with the Bot Framework Service
// We don't need to have these environment variables set in the development environment
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    stateEndpoint: process.env.BotStateEndpoint,
    openIdMetadata: process.env.BotOpenIdMetadata
});

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3979, function () {
    console.log('%s listening to %s', server.name, server.url);
});
// Listen for messages from users
server.post('/api/messages', connector.listen());

// Serve the magic code page for login.
server.get('/code', restify.plugins.serveStatic({
    'directory': path.join(__dirname, 'public'),
    'file': 'code.html'
}));


server.use(restify.plugins.queryParser());
server.use(restify.plugins.bodyParser());

// Setup the server with the bot secret
server.use(expressSession({
    secret: cfg.BOTAUTH_SECRET,
    resave: true,
    saveUninitialized: false
}));

const auth = AuthHelper.configure(server, bot);

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot.
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

// Create your bot with a function to receive messages from the user
// Using document DB for bot storage
var bot = new builder.UniversalBot(connector)
                        .set('storage',tableStorage);

// Intercept trigger event (ActivityTypes.Trigger)
bot.on('trigger', function (message) {
    // handle message from trigger function
    const payload = message.value;
    // This is handy to see the full message that you got back from the Azure Function.
    // const msg = new builder.Message()
    //                         .address(payload.address)
    //                         .text('Here\s the raw message: ' + JSON.stringify(payload));
    // bot.send(msg);
    switch(payload.action)
    {
        case 'WEBHOOK_REGISTERED':
            bot.beginDialog(payload.address, '/hookRegistered', payload.subscriptionId);
            break;
        case 'NOTIFICATION_RECIEVED':
            bot.beginDialog(payload.address, '/meetingRequested', payload);
            break;
        default:
            var reply = new builder.Message()
                                    .address(message.address)
                                    .text('default message: ' + message.text);
            bot.send(reply);
    }
});

bot.on('conversationUpdate', function(update) {
    if(update.membersAdded) {
        update.membersAdded.forEach(member => {
            if(member.id !== update.address.bot.id) {
                bot.beginDialog(update.address, '/intro');
            }
        });
    }
});

bot.dialog("/intro", (session) => {
    var reply = new builder.Message()
        .address(session.message.address)
        .text(`Hey, I can't do much, I suggest you start my signing in. Say intro to get this dialog again`)
        .suggestedActions(
            new builder.SuggestedActions()
                        .addAction(new builder.CardAction()
                                            .title('Logout')
                                            .type('postBack')
                                            .value('logout'))
                        .addAction(new builder.CardAction()
                                            .title('Signin to get help with your calendar')
                                            .type('imBack')
                                            .value('signin'))
        );
    bot.send(reply);
    session.endDialog();
});

// Handle message from user
bot.dialog("/", new builder.IntentDialog()
    .matches(/logout/, "/logout")
    .matches(/signin/, "/signin")
    .matches(/intro/, '/intro')
    .onDefault((session, args) => {
        session.endDialog("welcome");
    })
);

bot.dialog('/hookRegistered', (session, args) => {
    console.log('Hook registration completed');
    session.userData.subscriptionId = args;
    session.send(`I'm now watching for new meetings and will help you book travel time`);
    session.endDialog();
});

bot.dialog('/meetingRequested', [
    (session, args) => {
        console.log('Meeting request received');
        session.userData.meetingRequest = args;
        // Use this to debug your messages sent to the bot
        // session.send('here\s what I know about that meeting');
        // session.send(JSON.stringify(args));
        const start = new Date(args.start);
        const end = new Date(args.end);
        session.send(`${args.organizer} wants to have a meeting on ${start.toDateString()} from ${start.toTimeString()} until ${end.toTimeString()} at ${args.location} about ${args.subject}` );
        builder.Prompts.choice(session, 'Would you like to accept?', ['Yes', 'No']);
    },
    (session, results) => {
        if (results.response.entity !== 'Yes') {
            session.send('Ok, have a nice day');
            session.endDialog();
        }
        session.send(`I'll accept that meeting for you.`);

        // enqueue a message to accept the meeting
        var queueSvc = azureStorage.createQueueService(storageConnectionString);
        queueSvc.createQueueIfNotExists('accept-meeting', function(err, result, response){
            const message = {
                accessToken: session.userData.accessToken,
                meeting: session.userData.meetingRequest.resource
            }
            const msg = JSON.stringify(message);
            const queueMessageBuffer = new Buffer(msg).toString('base64');
            queueSvc.createMessage('accept-meeting', queueMessageBuffer, function(err, result, response){
                if (err) {
                    console.error(err);
                }
            });
        });

        builder.Prompts.choice(session, 'Would you like to block some travel time?', ['Yes', 'No']);
    },
    (session, results) => {
        if (results.response.entity !== 'Yes') {
            session.send('Ok, have a nice day');
            session.endDialog();
        }
        session.send(`Ok, let's find out about how much time to book.`)
        session.beginDialog('/requestTravelTime');
    },
    (session, results) => {
        session.send(`Ok, I'm blocking ${results.response} minutes of travel time for you`);
        const durationInMinutes = results.response;
        // enqueue messages to block some time.
        var queueSvc = azureStorage.createQueueService(storageConnectionString);
        queueSvc.createQueueIfNotExists('add-travel-meeting', function(err, result, response){
            const MS_PER_MINUTE = 60000;
            const start = new Date(session.userData.meetingRequest.start);
            const messages =[ {
                accessToken: session.userData.accessToken,
                start: session.userData.meetingRequest.end,
                durationInMins: durationInMinutes
            },{
                accessToken: session.userData.accessToken,
                meeting: session.userData.meetingRequest.resource,
                start: new Date(start - durationInMinutes * MS_PER_MINUTE).toISOString(),
                durationInMins: durationInMinutes
            }]
            for (let message of messages) {
                const msg = JSON.stringify(message);
                console.log(msg);
                const queueMessageBuffer = new Buffer(msg).toString('base64');
                queueSvc.createMessage('add-travel-meeting', queueMessageBuffer, function(err, result, response){
                    if (err) {
                        console.error(err);
                    }
                });
            }
        });

        session.endDialog();
    }
])
.beginDialogAction('meetingRequestedHelpAction', '/meetingRequestedHelp', { matches: /^help$/i });;

bot.dialog('/meetingRequestedHelp', function(session, args, next) {
    session.endDialog('Contextual help for this dialog');
})
// Once triggered, will restart the dialog.
.reloadAction('startOver', 'Ok, starting over.', {
    matches: /^start over$/i
});

bot.dialog('/requestTravelTime',[
    session => {
        builder.Prompts.choice(session, 'How much time? ', ['15 minutes', '30 minutes', '1 hour', 'Other']);
    },
    (session, results) => {
        if (results.response.entity === 'Other') {
            // use a dialog here to gather the infomation on how long.
            session.beginDialog('/customTravelLength');
        }
        if (results.response.index === 0) {
            session.endDialogWithResult({ response: 15 });
        }
        if (results.response.index === 1) {
            session.endDialogWithResult({ response: 30 });
        }
        if (results.response.index === 2) {
            session.endDialogWithResult({ response: 60 });
        }
    },
    (session, results) => {
        // this returns the results of the /customTravelLength Dialog.
        session.endDialogWithResult(results);
    }
]);

bot.dialog('/customTravelLength', [
    (session, args) => {
        if (args && args.reprompt) {
            builder.Prompts.text(session, `Sorry, can you try telling me in hours or minutes, that's all I know about yet`);
        } else {
            builder.Prompts.text(session, 'How long should I block for travel time?');
        }
    },
    (session, results) => {
        const userInput = results.response;
        let minutes = 0;
        if ('string' === (typeof userInput)) {
            if (userInput.indexOf('min') > 1) {
                minutes = userInput.split(' ')[0];
            }
            if (userInput.indexOf('hour') > 1) {
                minutes = userInput.split(' ')[0]*60;
            }
        }
        if (minutes === 0) {
            session.replaceDialog('/customTravelLength', {reprompt: true});
        } else {
            session.endDialogWithResult({ response: minutes});
        }
    }
]);

bot.dialog("/logout", (session) => {
    auth.logout(session, "aadv2");
    session.endDialog("logged_out");
});

bot.dialog("/signin", [].concat(
    auth.authenticate("aadv2"),
    (session, args, skip) => {
        let user = auth.profile(session, "aadv2");
        // persist some userful means of finding a user.
        session.userData.oid = user.oid;
        session.userData.address= session.message.address;
        session.userData.accessToken = user.accessToken;
        session.userData.refreshToken = user.refreshToken;
        session.sendTyping();
        if (session.userData.subscriptionId) {
            session.send(`All set up and ready to go, I'll let you know when you get some meeting requests` );
            session.endDialog();
            return;
        }
        var queueSvc = azureStorage.createQueueService(storageConnectionString);
        queueSvc.createQueueIfNotExists('bot-hook-registration', function(err, result, response){
            if(!err){
                // enqueue a message to register a webhook.
                const raw = JSON.parse(user._raw);
                let message = {
                    accessToken: user.accessToken,
                    userId: user.oid,
                    address: session.message.address
                }
                let msg = JSON.stringify(message);
                console.log(msg)
                var queueMessageBuffer = new Buffer(msg).toString('base64');
                queueSvc.createMessage('bot-hook-registration', queueMessageBuffer, function(err, result, response){
                    if(!err){
                        // Message inserted
                        session.send(`I'm just getting setup to watch your calendar for changes`);
                    } else {
                        // this should be a log for the dev, not a message to the user
                        session.send('There was an error inserting your message into queue: '+ err);
                        console.error(err);
                    }
                    session.endDialog();
                });
            } else {
                // this should be a log for the dev, not a message to the user
                session.send('There was an error creating your queue: '+ err);
                console.error(err);
                session.endDialog();
            }
        });

        session.send(`Hi ${user.displayName}`);
    }
));

bot.dialog('help', function (session, args, next) {
    session.send("I'll watch your calendar and let you know when you get new meeting requests.");
    session.send("If you want to start over a section, try saying restart");
    session.endDialog("If you need to exit a conversation try using the cancel or exit command.");
})
.triggerAction({
    matches: /^help$/i,
    onSelectAction: (session, args, next) => {
        // Add the help dialog to the dialog stack
        // (override the default behavior of replacing the stack)
        session.beginDialog(args.action, args);
    }
});



bot.dialog('exit', function (session, args, next) {
    session.endConversation('exiting...');
})
.triggerAction({
    matches: /^(exit)|(cancel)|(quit)$/i,
    onSelectAction: (session, args, next) => {
        // Add the exit dialog to the dialog stack
        session.beginDialog(args.action, args);
    }
});
