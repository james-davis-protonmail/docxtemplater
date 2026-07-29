"use strict";

const path = require("path");
const Mocha = require("mocha");

const mocha = new Mocha({ fullTrace: true });
mocha.addFile(path.join(__dirname, "test.js"));
mocha.run((failures) => {
	process.exitCode = failures ? 1 : 0;
});
