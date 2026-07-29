"use strict";

const { expect } = require("chai");
const PizZip = require("pizzip");
const Docxtemplater = require("../es6/docxtemplater.js");
const expressionParser = require("../es6/expressions.js");
const inspectModuleFactory = require("../es6/inspect-module.js");
const ConditionalSwitchModule = require("./index.js");

const CONTENT_TYPES = `<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;

function escapeXml(value) {
	return value
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;");
}

function paragraph(value, runProperties = "") {
	return `<w:p><w:r>${runProperties}<w:t xml:space="preserve">${escapeXml(
		value
	)}</w:t></w:r></w:p>`;
}

function documentXml(body) {
	return `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:headerReference r:id="rIdHeader"/><w:footerReference r:id="rIdFooter"/></w:sectPr></w:body></w:document>`;
}

function createZip(body, { header, footer } = {}) {
	const zip = new PizZip();
	zip.file("[Content_Types].xml", CONTENT_TYPES);
	zip.file("_rels/.rels", ROOT_RELS);
	zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS, {
		createFolders: true,
	});
	zip.file("word/document.xml", documentXml(body), { createFolders: true });
	if (header != null) {
		zip.file(
			"word/header1.xml",
			`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${header}</w:hdr>`
		);
	}
	if (footer != null) {
		zip.file(
			"word/footer1.xml",
			`<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${footer}</w:ftr>`
		);
	}
	return zip;
}

function createDoc(template, options = {}) {
	const modules = [new ConditionalSwitchModule(), ...(options.modules || [])];
	return new Docxtemplater(createZip(template, options.parts), {
		parser: options.parser || expressionParser,
		paragraphLoop: true,
		errorLogging: false,
		...options.docOptions,
		modules,
	});
}

function xmlText(xml) {
	return xml
		.replace(/<[^>]+>/g, "")
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&amp;/g, "&");
}

function renderedPart(doc, path = "word/document.xml") {
	return doc.getZip().files[path].asText();
}

function renderText(template, data, options) {
	const doc = createDoc(template, options);
	doc.render(data);
	return xmlText(renderedPart(doc));
}

function getErrors(error) {
	return error.properties.errors || [error];
}

describe("private conditional switch module", () => {
	it("renders the documented inline example", () => {
		const template = paragraph(
			"The product has a {!switch}{!case price < 10}low{!case price < 100}medium{!default}high{/!switch} price."
		);
		expect(renderText(template, { price: 50 })).to.equal(
			"The product has a medium price."
		);
	});

	it("does not affect templates without switch tags", () => {
		expect(
			renderText(paragraph("Hello {name}!"), { name: "Ada" })
		).to.equal("Hello Ada!");
	});

	it("selects first, later, default, and empty branches correctly", () => {
		const template = paragraph(
			"{!switch}{!case first}A{!case second}B{!default}D{/!switch}"
		);
		expect(renderText(template, { first: true, second: true })).to.equal(
			"A"
		);
		expect(renderText(template, { first: false, second: true })).to.equal(
			"B"
		);
		expect(renderText(template, { first: false, second: false })).to.equal(
			"D"
		);
		expect(
			renderText(paragraph("{!switch}{!case first}A{/!switch}"), {
				first: false,
			})
		).to.equal("");
	});

	it("does not evaluate a later case after a match", () => {
		let evaluations = 0;
		const parser = expressionParser.configure({
			postEvaluate(value, tag) {
				if (tag === "later") {
					evaluations++;
				}
				return value;
			},
		});
		const template = paragraph(
			"{!switch}{!case first}A{!case later}B{/!switch}"
		);
		expect(
			renderText(template, { first: true, later: true }, { parser })
		).to.equal("A");
		expect(evaluations).to.equal(0);
	});

	it("supports multiline switches and removes control-only paragraphs", () => {
		const template = [
			paragraph("Before"),
			paragraph("{!switch}"),
			paragraph("{!case first}"),
			paragraph("First"),
			paragraph("{!case second}"),
			paragraph("Second"),
			paragraph("{!default}"),
			paragraph("Default"),
			paragraph("{/!switch}"),
			paragraph("After"),
		].join("");
		const doc = createDoc(template);
		doc.render({ first: false, second: true });
		const xml = renderedPart(doc);
		expect(xmlText(xml)).to.equal("BeforeSecondAfter");
		expect((xml.match(/<w:p>/g) || []).length).to.equal(3);
	});

	it("supports comparisons, booleans, object access, negation, and missing values", () => {
		const template = paragraph(
			"{!switch}{!case customer.active && amount >= 1000 && !cancelled}approved{!case missing.value}missing{!default}other{/!switch}"
		);
		expect(
			renderText(template, {
				customer: { active: true },
				amount: 1000,
				cancelled: false,
			})
		).to.equal("approved");
		expect(
			renderText(template, {
				customer: { active: false },
				amount: 5000,
			})
		).to.equal("other");
	});

	it("uses registered Angular filters", () => {
		expressionParser.filters.switchCheap = (price, limit) => price < limit;
		const template = paragraph(
			"{!switch}{!case price | switchCheap:threshold}cheap{!default}expensive{/!switch}"
		);
		expect(renderText(template, { price: 5, threshold: 10 })).to.equal(
			"cheap"
		);
	});

	it("respects loop-local scope", () => {
		const template = paragraph(
			"{#items}{!switch}{!case active}{name}:on{!default}{name}:off{/!switch};{/items}"
		);
		expect(
			renderText(template, {
				items: [
					{ name: "A", active: true },
					{ name: "B", active: false },
				],
			})
		).to.equal("A:on;B:off;");
	});

	it("supports nested switches and multiple switches in a paragraph", () => {
		const template = paragraph(
			"{!switch}{!case customer.active}{!switch}{!case customer.vip}VIP{!default}ACTIVE{/!switch}{!default}INACTIVE{/!switch} / {!switch}{!case paid}PAID{!default}DUE{/!switch}"
		);
		expect(
			renderText(template, {
				customer: { active: true, vip: true },
				paid: false,
			})
		).to.equal("VIP / DUE");
	});

	it("supports renderAsync", async () => {
		const doc = createDoc(
			paragraph(
				"{!switch}{!case enabled}{value}{!default}disabled{/!switch}"
			)
		);
		await doc.renderAsync({
			enabled: true,
			value: Promise.resolve("async value"),
		});
		expect(xmlText(renderedPart(doc))).to.equal("async value");
	});

	it("works in tables, headers, and footers", () => {
		const switchText = "{!switch}{!case enabled}YES{!default}NO{/!switch}";
		const body = `<w:tbl><w:tr><w:tc>${paragraph(
			switchText
		)}</w:tc></w:tr></w:tbl>`;
		const doc = createDoc(body, {
			parts: {
				header: paragraph(`H:${switchText}`),
				footer: paragraph(`F:${switchText}`),
			},
		});
		doc.render({ enabled: true });
		expect(xmlText(renderedPart(doc))).to.equal("YES");
		expect(xmlText(renderedPart(doc, "word/header1.xml"))).to.equal(
			"H:YES"
		);
		expect(xmlText(renderedPart(doc, "word/footer1.xml"))).to.equal(
			"F:YES"
		);
	});

	it("handles control tags split across Word runs", () => {
		const body =
			"<w:p><w:r><w:t>{!sw</w:t></w:r><w:r><w:t>itch}{!ca</w:t></w:r>" +
			"<w:r><w:t>se enabled}selected{!default}other{/!switch}</w:t></w:r></w:p>";
		expect(renderText(body, { enabled: true })).to.equal("selected");
	});

	it("preserves selected branch run formatting", () => {
		const body =
			"<w:p><w:r><w:t>{!switch}{!case enabled}</w:t></w:r>" +
			"<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>" +
			"<w:r><w:t>{!default}Plain{/!switch}</w:t></w:r></w:p>";
		const doc = createDoc(body);
		doc.render({ enabled: true });
		const xml = renderedPart(doc);
		expect(xmlText(xml)).to.equal("Bold");
		expect(xml).to.include("<w:b/>");
	});

	it("keeps conditions and placeholders from every branch discoverable", () => {
		const inspect = inspectModuleFactory();
		const template = paragraph(
			"{!switch}{!case price < threshold}{product.name}{!default}{fallbackText}{/!switch}"
		);
		createDoc(template, { modules: [inspect] });
		const tags = inspect.getAllTags();
		expect(tags).to.have.property("price");
		expect(tags).to.have.property("threshold");
		expect(tags).to.have.property("product.name");
		expect(tags).to.have.property("fallbackText");
		const structured = inspect.getAllStructuredTags();
		expect(JSON.stringify(structured)).to.include("price < threshold");
	});

	const invalidTemplates = [
		[
			"missing closing switch",
			"{!switch}{!case ok}yes",
			"conditional_switch_missing_closing_tag",
		],
		[
			"closing without opening",
			"{/!switch}",
			"conditional_switch_unmatched_closing_tag",
		],
		[
			"case outside switch",
			"{!case ok}",
			"conditional_switch_branch_outside_switch",
		],
		[
			"default outside switch",
			"{!default}",
			"conditional_switch_branch_outside_switch",
		],
		[
			"switch without cases",
			"{!switch}{!default}x{/!switch}",
			"conditional_switch_no_cases",
		],
		[
			"empty case",
			"{!switch}{!case}x{/!switch}",
			"conditional_switch_empty_case",
		],
		[
			"multiple defaults",
			"{!switch}{!case ok}x{!default}y{!default}z{/!switch}",
			"conditional_switch_multiple_defaults",
		],
		[
			"case after default",
			"{!switch}{!case one}x{!default}y{!case two}z{/!switch}",
			"conditional_switch_case_after_default",
		],
	];

	for (const [name, template, expectedId] of invalidTemplates) {
		it(`reports ${name}`, () => {
			expect(() => createDoc(paragraph(template)))
				.to.throw()
				.and.satisfy((error) =>
					getErrors(error).some(
						(item) => item.properties.id === expectedId
					)
				);
		});
	}

	it("identifies invalid Angular case expressions", () => {
		let thrown;
		try {
			createDoc(
				paragraph("{!switch}{!case customer.}x{!default}y{/!switch}")
			);
		} catch (error) {
			thrown = error;
		}
		expect(thrown).to.not.equal(undefined);
		const invalid = getErrors(thrown).find(
			(error) =>
				error.properties.id === "conditional_switch_invalid_expression"
		);
		expect(invalid).to.not.equal(undefined);
		expect(invalid.properties.expression).to.equal("customer.");
		expect(invalid.properties.file).to.equal("word/document.xml");
	});
});
