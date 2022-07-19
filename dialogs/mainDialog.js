// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TimexProperty } = require('@microsoft/recognizers-text-data-types-timex-expression');
const { MessageFactory, InputHints, CardFactory } = require('botbuilder');
const { LuisRecognizer } = require('botbuilder-ai');
const { ComponentDialog, DialogSet, DialogTurnStatus, TextPrompt, WaterfallDialog } = require('botbuilder-dialogs');
const moment = require('moment-timezone');

var chatConnection = require('../bots/resources/chatConnection.json');

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';

class MainDialog extends ComponentDialog {
    constructor(luisRecognizer, departmentDialog) {
        super('MainDialog');

        if (!luisRecognizer) throw new Error('[MainDialog]: Missing parameter \'luisRecognizer\' is required');
        this.luisRecognizer = luisRecognizer;

        if (!departmentDialog) throw new Error('[MainDialog]: Missing parameter \'departmentDialog\' is required');

        // Define the main dialog and its related components.
        this.addDialog(new TextPrompt('TextPrompt'))
            .addDialog(departmentDialog)
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
                this.introStep.bind(this),
                this.actStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} turnContext
     * @param {*} accessor
     */
    async run(turnContext, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async introStep(stepContext) {
        if (!this.luisRecognizer.isConfigured) {
            const messageText = 'NOTE: LUIS is not configured. To enable all capabilities, add `LuisAppId`, `LuisAPIKey` and `LuisAPIHostName` to the .env file.';
            await stepContext.context.sendActivity(messageText, null, InputHints.IgnoringInput);
            return await stepContext.next();
        }

        const messageText = stepContext.options.restartMsg ? stepContext.options.restartMsg : `Try asking: "Can you connect me with someone from the Computer Science department?"`;
        const promptMessage = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);
        return await stepContext.prompt('TextPrompt', { prompt: promptMessage });
    }

    /**
     * Second step in the waterfall.  This will use LUIS to attempt to extract the origin, destination and travel dates.
     * Then, it hands off to the departmentDialog child dialog to collect any remaining details.
     */
    async actStep(stepContext) {
        const departmentDetails = {};

        if (!this.luisRecognizer.isConfigured) {
            // LUIS is not configured, we just run the departmentDialog path.
            return await stepContext.beginDialog('departmentDialog', departmentDetails);
        }

        const luisResult = await this.luisRecognizer.executeLuisQuery(stepContext.context);
        switch (LuisRecognizer.topIntent(luisResult)) {
        case 'SelectDepartmentMember': {
            // Extract the values for the composite entities from the LUIS result.
            const departmentEntities = this.luisRecognizer.getDepartmentEntities(luisResult);

            // Initialize departmentDetails with any entities we may have found in the response.
            departmentDetails.departmentName = departmentEntities.departmentName;
            departmentDetails.facultyName = departmentEntities.facultyName;

            console.log('LUIS extracted these department details:', JSON.stringify(departmentDetails));

            if (departmentDetails.departmentName) {
                return await stepContext.beginDialog('departmentDialog', departmentDetails);
            }
        }

        default: {
            // Catch all for unhandled intents
            const didntUnderstandMessageText = `Sorry, I didn't get that.`;
            await stepContext.context.sendActivity(didntUnderstandMessageText, didntUnderstandMessageText, InputHints.IgnoringInput);
        }
        }

        return await stepContext.next();
    }

    async finalStep(stepContext) {
        // If the child dialog ("departmentDialog") was cancelled or the user failed to confirm, the Result here will be null.
        if (stepContext.result) {
            const result = stepContext.result;
            const msg = `Okay. I am connecting you to ${ result.facultyName }...`;
            await stepContext.context.sendActivity(msg, msg, InputHints.IgnoringInput);
            chatConnection.body[0].columns[1].items[0].text = `You are now chatting with ${result.facultyName}.`;
            chatConnection.body[0].columns[1].items[1].text = (new Date()).toUTCString();
            return await stepContext.context.sendActivity({
                attachments: [CardFactory.adaptiveCard(chatConnection)]
            });
        }
        
        // Restart the main dialog with a different message the second time around
        return await stepContext.replaceDialog(this.initialDialogId, { restartMsg: 'What else can I do for you?' });
    }
}

module.exports.MainDialog = MainDialog;
