"use strict";
var Alexa = require("alexa-sdk");
var MicrosoftGraph = require("msgraph-sdk-javascript");
var voiceInsights = require("voice-insights-sdk");

// Helpers
// var notebookHelper = require("./notebookHelper.js");
// var sectionHelper = require("./sectionHelper.js");
// var pageHelper = require("./pageHelper.js");
var speechHelper = require("./speechHelper.js");

var VI_APP_TOKEN = process.env.VI_APP_TOKEN;

var currLogLevel = process.env.LOG_LEVEL != null ? process.env.LOG_LEVEL : 'info';

var client = {};

// print the log statement, only if the requested log level is greater than the current log level
function log(statement, logLevel) {
    // no loglevel, set to debug
    if(!logLevel){
        logLevel = logLevels.debug;
    }
    // output log if greater then log level
    if(logLevel >= logLevels[currLogLevel] ) {
        console.log(statement);
    }
}

// Entry point for Alexa
exports.handler = function(event, context, callback) {
    // Initialize Telemetry
    voiceInsights.initialize(event.session, VI_APP_TOKEN);

    // DEBUG: return all the environment varibles
    log(process.env, logLevels.debug);
    // DEBUG: return the Node version info
    log(process.versions, logLevels.debug);

    // Verify that the Request is Intended for Your Service
    // This value is set on the server or in the .env file
    // SETUP NOTE: Add Lambda Environment Variable called APPLICATION_ID
    var appId = process.env.APPLICATION_ID;

    var alexa = Alexa.handler(event, context);
    alexa.appId = appId;
    alexa.registerHandlers(handlers);

    // Get the OAuth 2.0 Bearer Token from the linked account
    var token = event.session.user.accessToken;

    // validate the Auth Token
    if (token) {
        log("Auth Token: " + token, logLevels.debug);
        // TODO: validate the token

        // Initialize the Microsoft Graph client
        client = MicrosoftGraph.Client.init({
            authProvider: (done) => {
                done(null, token);
            }
        });

        // Handle the intent
        alexa.execute();

    } else {
        // no token! display card and let user know they need to sign in
        log("No Auth Token", logLevels.warn);
        var speechOutput = linkAccout();

        voiceInsights.track("NoAuthToken", null, null, (error, response) => {
            alexa.emit(":tellWithLinkAccountCard", speechOutput);
        });
    }
};

var handlers = {
    "LaunchRequest": function () {
        log("LaunchRequest", logLevels.info);
        var welcome = welcomeMessages();
        var speechOutput = wecome.message;
        var repromptSpeech = welcome.message;
        var cardTitle = welcome.messageCardTitle;
        var cardContent = welcome.messageCard;
        VoiceInsights.track("LaunchRequest", null, null, (error, response) => {
            this.emit(":askWithCard", speechOutput, repromptSpeech, cardTitle, cardContent);
        });
    },
    "SessionEndedRequest": function () {
        log("SessionEndedRequest", logLevels.info);
        var speechOutput = shutdownMessage();
        VoiceInsights.track("SessionEndedRequest", null, null, (error, response) => {
            this.emit(":tell", speechOutput);
        });
    },
    "ListNoteBooksIntent": function () {
        log("ListNoteBooksIntent", logLevels.info);

        ListNoteBooksIntent(this);

    },
    "ListSectionsIntent": function () {
        log("ListSectionsIntent", logLevels.info);

        ListSectionsIntent(this);

    },
    "ListPagesIntent": function () {
        log("ListPagesIntent", logLevels.info);

        ListPagesIntent(this);

    },
    "WhoAmIIntent": function () {
        log("WhoAmIIntent", logLevels.info);

        WhoAmIIntent(this);

    },
    "AMAZON.StopIntent": function() {
        log("StopIntent", logLevels.info);
        var speechOutput = shutdownMessage();
        VoiceInsights.track("StopIntent", null, null, (error, response) => {
            this.emit(":tell", speechOutput);
        });
    },
    // Let the user completely exit the skill
    "AMAZON.CancelIntent": function() {
        log("CancelIntent", logLevels.info);
        VoiceInsights.track("CancelIntent", null, null, (error, response) => {
            this.emit(":tell", shutdownMessage());
        });
    },
    // Provide help about how to use the skill
    "AMAZON.HelpIntent": function () {
        log("HelpIntent", logLevels.info);
        var Help = helpMessage();
        var speechOutput = Help.message;
        var repromptSpeech = Help.message;
        var cardTitle = Help.messageCardTitle;
        var cardContent = Help.messageCard;
        VoiceInsights.track("HelpIntent", null, null, (error, response) => {
           this.emit(":askWithCard", speechOutput, repromptSpeech, cardTitle, cardContent);
        });
    },
    // Catch everything else
    "Unhandled": function () {
        log("UnhandledIntent", logLevels.info);
        var Help = helpMessage();
        var speechOutput = Help.message;
        var repromptSpeech = Help.message;
        VoiceInsights.track("UnhandledIntent", null, null, (error, response) => {
            this.emit(":ask", speechOutput, repromptSpeech);
        });
    }
};

function WhoAmIIntent(alexaResponse){
        // get the authenticated user info
        getUser(alexaResponse)
        // **
        // handle the getUser results
        .then(function(user){
            //check if the user is valid
            if(!user) throw "There is no user returned ";

            var displayName = user.displayName;

            //return the results to Alexa
            VoiceInsights.track("WhoAmIIntent", null, null, (error, response) => {
                return alexaResponse.emit(":tell", "The linked account belongs to " + displayName);
            });
        })
        .catch(function(err){
            log("WhoAmIIntent getUser Error: " + JSON.stringify(err), logLevels.error);
            alexaResponse.emit(":tell", "There was an error. " + err.message)
            // re-throw the error so the chain of promises don't continue
            throw "There was a getuser catch error: " + JSON.stringify(err);
        })
}

function ListNoteBooksIntent(alexaResponse){
        // get the authenticated user info
        getUser(alexaResponse)
        // **
        // handle the getUser results
        .then(function(user){
            //check if the user is valid
            if(!user) throw "There is no user returned ";

            // then send a mail to the current user           
            //return sendMail(user);
        })
        .catch(function(err){
            log("ListNoteBooksIntent getUser Error: " + JSON.stringify(err), logLevels.error);
            alexaResponse.emit(":tell", "There was an error. " + err.message)
            // re-throw the error so the chain of promises don't continue
            throw "There was a getuser catch error: " + JSON.stringify(err);
        })

        // handle the sendMail results
        .then(function(mail){
            // check if the sendmail succeded
            if(!mail) throw "There was an error sending mail";

            // then send confirmation back to alexa
            var mailSubject = mail.Message.Subject;
            log("Mail Sent: " + JSON.stringify(mail), logLevels.debug);
            //return the results to Alexa
            VoiceInsights.track("sendMailIntent", null, null, (error, response) => {
                return alexaResponse.emit(":tell", "Mail sent to you with a subject of " + mailSubject);
            });
        })
        .catch(function(err){
            log("sendMail Error: " + JSON.stringify(err), logLevels.error);
            alexaResponse.emit(":tell", "There was an error sending the mail");
            // re-throw the error so the chain of promises don't continue
            throw "There was an sendmail catch error: " + JSON.stringify(err);
        })
    getNoteBooks()
    	.then(function(noteBooks){
    		if(!noteBooks) throw "There were no notebooks returned";

    		var listOfNotebooks = [];
    		for (var i = noteBooks.value.length - 1; i >= 0; i--) {
    			var noteBook = noteBooks.value[i];
    			listOfNotebooks.push(noteBook.displayName);
    		}

			// VoiceInsights.track("ListNoteBooksIntent", null, null, (error, response) => {
			// 	return alexaResponse.emit(":tell", "The linked account belongs to " + displayName);
			// });

            var speechOutput = 'Your notebooks are ' + sayArray(listOfNotebooks,  'and')
                    + '. please choose either ' +  sayArray(listOfNotebooks,  'or');
        	this.response.speak(speechOutput);
        	this.emit(':responseReady');

    	})
    	.catch(function(err){
    		log("ListNoteBooksIntent getNoteBooks Error: " + JSON.stringify(err), logLevels.error);
    		alexaResponse.emit(":tell", "There was an error. " + err.message);
    		throw "There was a getNoteBooks catch error: " + JSON.stringify(err);
    	})
}

function ListSectionsIntent(alexaResponse){
	getSections()
    	.then(function(sections){
    		if(!sections) throw "There were no sections returned";

    		var listOfSections = [];
    		for (var i = sections.value.length - 1; i >= 0; i--) {
    			var section = sections.value[i];
    			listOfSections.push(section.displayName);
    		}

			// VoiceInsights.track("ListNoteBooksIntent", null, null, (error, response) => {
			// 	return alexaResponse.emit(":tell", "The linked account belongs to " + displayName);
			// });

            var speechOutput = 'Your notebooks are ' + sayArray(listOfSections,  'and')
                    + '. please choose either ' +  sayArray(listOfSections,  'or');
        	this.response.speak(speechOutput);
        	this.emit(':responseReady');
    	})
    	.catch(function(err){
    		log("ListSectionsIntent getSections Error: " + JSON.stringify(err), logLevels.error);
    		alexaResponse.emit(":tell", "There was an error. " + err.message);
    		throw "There was a getSections catch error: " + JSON.stringify(err);
    	})
}

function ListPagesintent(alexaResponse){
	getPages()
    	.then(function(pages){
    		if(!pages) throw "There were no pages returned";

    		var listOfPages = [];
    		for (var i = pages.value.length - 1; i >= 0; i--) {
    			var page = pages.value[i];
    			listOfPages.push(page.title);
    		}

			// VoiceInsights.track("ListNoteBooksIntent", null, null, (error, response) => {
			// 	return alexaResponse.emit(":tell", "The linked account belongs to " + displayName);
			// });

            var speechOutput = 'Your notebooks are ' + sayArray(listOfPages,  'and')
                    + '. please choose either ' +  sayArray(listOfPages,  'or');
        	this.response.speak(speechOutput);
        	this.emit(':responseReady');
    	})
    	.catch(function(err){
    		log("ListPagesintent getPages Error: " + JSON.stringify(err), logLevels.error);
    		alexaResponse.emit(":tell", "There was an error. " + err.message);
    		throw "There was a getPages catch error: " + JSON.stringify(err);
    	})
}

function getNoteBooks() {
	log("getNoteBooks", logLevels.debug);

	return client
		.api("/me/onenote/notebooks")
		.get();
}

function getSections() {
	log("getNoteBooks", logLevels.debug);

	return client
		.api("/me/onenote/sections")
		.get();
}

function getPages() {
	log("getNoteBooks", logLevels.debug);

	return client
		.api("/me/onenote/pages")
		.get();
}

function getUser(){
    log("getUser", logLevels.debug)
    //Make a call to the Graph API, this returns a Promise
    return client
            .api("/me")
            .get();
}

function sayArray(myData, andor) {
    // the first argument is an array [] of items
    // the second argument is the list penultimate word; and/or/nor etc.

    var listString = '';

    if (myData.length == 1) {
        listString = myData[0];
    } else {
        if (myData.length == 2) {
            listString = myData[0] + ' ' + andor + ' ' + myData[1];
        } else {

            for (var i = 0; i < myData.length; i++) {
                if (i < myData.length - 2) {
                    listString = listString + myData[i] + ', ';
                    if (i = myData.length - 2) {
                        listString = listString + myData[i] + ', ' + andor + ' ';
                    }

                } else {
                    listString = listString + myData[i];
                }

            }
        }

    }

    return(listString);
}