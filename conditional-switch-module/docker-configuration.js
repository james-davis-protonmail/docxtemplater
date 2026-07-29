"use strict";

const ConditionalSwitchModule = require("./index.js");

/*
 * Merge this hook into the Docker image's configuration.js. Docxtemplater
 * Docker calls it before compilation, which is early enough for custom tags
 * to participate in parsing and in the retrieve-tags endpoints.
 */
module.exports = {
	configureDocxtemplater(doc) {
		doc.attachModule(new ConditionalSwitchModule());
		return doc;
	},
};
