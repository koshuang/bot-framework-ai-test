// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { InputHints, MessageFactory } = require('botbuilder');
const { ConfirmPrompt, ChoicePrompt, TextPrompt, WaterfallDialog, ListStyle } = require('botbuilder-dialogs');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');
const FACULTY_MEMBERS = require('../bots/resources/facultyMembers.json');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'choicePrompt';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class DepartmentDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'departmentDialog');

        this.addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.facultyMemberStep.bind(this),
                this.confirmStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    async facultyMemberStep(stepContext) {
        const departmentDetails = stepContext.options;

        if (!departmentDetails.facultyName && departmentDetails.departmentName) {
            const messageText = `Who would you like to speak with from the ${departmentDetails.departmentName} department? Here are the faculty members online.`;
            const msg = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);

            return await stepContext.prompt(CHOICE_PROMPT, {
                prompt: msg,
                retryPrompt: 'Please choose an option from the list.',
                choices: FACULTY_MEMBERS[departmentDetails.departmentName.toLowerCase()],
                style: ListStyle.heroCard,
            });
        }

        return await stepContext.next(departmentDetails.facultyName);
    }

    async confirmStep(stepContext) {
        const departmentDetails = stepContext.options;

        // Capture the results of the previous step
        departmentDetails.facultyName = departmentDetails.facultyName !== undefined ? departmentDetails.facultyName : stepContext.result.value;
        const messageText = `Please confirm you want to speak with ${ departmentDetails.facultyName } from the ${ departmentDetails.departmentName } department.`;
        const msg = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);

        // Offer a YES/NO prompt.
        return await stepContext.prompt(CONFIRM_PROMPT, { prompt: msg });
    }

    async finalStep(stepContext) {
        if (stepContext.result === true) {
            const departmentDetails = stepContext.options;
            return await stepContext.endDialog(departmentDetails);
        }
        return await stepContext.endDialog();
    }
}

module.exports.DepartmentDialog = DepartmentDialog;
