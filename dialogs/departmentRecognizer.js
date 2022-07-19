// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { LuisRecognizer } = require('botbuilder-ai');

class DepartmentRecognizer {
    constructor(config) {
        const luisIsConfigured = config && config.applicationId && config.endpointKey && config.endpoint;
        if (luisIsConfigured) {
            // Set the recognizer options depending on which endpoint version you want to use e.g v2 or v3.
            // More details can be found in https://docs.microsoft.com/en-gb/azure/cognitive-services/luis/luis-migration-api-v3
            const recognizerOptions = {
                apiVersion: 'v3'
            };

            this.recognizer = new LuisRecognizer(config, recognizerOptions);
        }
    }

    get isConfigured() {
        return (this.recognizer !== undefined);
    }

    /**
     * Returns an object with preformatted LUIS results for the bot's dialogs to consume.
     * @param {TurnContext} context
     */
    async executeLuisQuery(context) {
        return await this.recognizer.recognize(context);
    }

    getDepartmentEntities(result) {
        let _departmentName, _facultyName;
        
        if (result.entities.$instance.DepartmentName) {
            _departmentName = result.entities.$instance.DepartmentName[0].text;
        }

        if (result.entities.$instance.FacultyName) {
            _facultyName = result.entities.$instance.FacultyName[0].text;
        }

        return { departmentName: _departmentName, facultyName: _facultyName };
    }
}

module.exports.DepartmentRecognizer = DepartmentRecognizer;
