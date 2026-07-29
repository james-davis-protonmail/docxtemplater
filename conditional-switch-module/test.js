"use strict";

const fs = require("fs");
const path = require("path");
const { DOMParser } = require("@xmldom/xmldom");
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

function createPptxDoc(templateParagraphs, replacementFactory) {
	const zip = new PizZip(
		fs.readFileSync(
			path.resolve(__dirname, "../examples/simple-example.pptx")
		)
	);
	const slidePath = "ppt/slides/slide1.xml";
	const slideXml = zip.file(slidePath).asText();
	const paragraphMatch = slideXml.match(/<a:p>[\s\S]*?<\/a:p>/);
	if (!paragraphMatch) {
		throw new Error("PowerPoint fixture has no paragraph");
	}
	const paragraphTemplate = paragraphMatch[0];
	const replacement = replacementFactory
		? replacementFactory(paragraphTemplate)
		: templateParagraphs
				.map((value) =>
					paragraphTemplate.replace("Hello {name}", escapeXml(value))
				)
				.join("");
	zip.file(slidePath, slideXml.replace(paragraphTemplate, replacement));
	return new Docxtemplater(zip, {
		parser: expressionParser,
		errorLogging: false,
		modules: [new ConditionalSwitchModule()],
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

function expectWellFormedXml(xml) {
	expect(() =>
		new DOMParser({
			onError(level, message) {
				if (level !== "warning") {
					throw new Error(message);
				}
			},
		}).parseFromString(xml, "text/xml")
	).to.not.throw();
}

function directElementChildNames(xml, parentTag) {
	const document = new DOMParser().parseFromString(xml, "text/xml");
	const parent = document.getElementsByTagName(parentTag)[0];
	const names = [];
	for (let child = parent.firstChild; child; child = child.nextSibling) {
		if (child.nodeType === 1) {
			names.push(child.nodeName);
		}
	}
	return names;
}

function runContaining(xml, text) {
	return (xml.match(/<([wa]):r(?:\s[^>]*)?>[\s\S]*?<\/\1:r>/g) || []).find(
		(run) => xmlText(run).includes(text)
	);
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

	it("removes styled control-only paragraphs", () => {
		function styledControl(value) {
			return `<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(
				value
			)}</w:t></w:r></w:p>`;
		}
		const template = [
			styledControl("{!switch}"),
			styledControl("{!case enabled}"),
			paragraph("YES"),
			styledControl("{!default}"),
			paragraph("NO"),
			styledControl("{/!switch}"),
		].join("");

		for (const [enabled, expected] of [
			[true, "YES"],
			[false, "NO"],
		]) {
			const doc = createDoc(template);
			doc.render({ enabled });
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect((xml.match(/<w:p>/g) || []).length).to.equal(1);
			expect(xml).to.not.include("<w:pPr>");
			expect(xml).to.not.include("<w:b/>");
			expectWellFormedXml(xml);
		}
	});

	it("removes proofing and cached-layout metadata with control paragraphs", () => {
		function metadataControl(value) {
			return (
				'<w:p><w:proofErr w:type="spellStart"/>' +
				`<w:r><w:t>${escapeXml(
					value
				)}</w:t><w:lastRenderedPageBreak/></w:r>` +
				'<w:proofErr w:type="spellEnd"/></w:p>'
			);
		}
		const template = [
			metadataControl("{!switch}"),
			metadataControl("{!case enabled}"),
			paragraph("YES"),
			metadataControl("{!default}"),
			paragraph("NO"),
			metadataControl("{/!switch}"),
		].join("");

		for (const [enabled, expected] of [
			[true, "YES"],
			[false, "NO"],
		]) {
			const doc = createDoc(template);
			doc.render({ enabled });
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect((xml.match(/<w:p>/g) || []).length).to.equal(1);
			expect(xml).to.not.include("<w:proofErr");
			expect(xml).to.not.include("<w:lastRenderedPageBreak");
			expectWellFormedXml(xml);
		}
	});

	it("does not confuse paragraph tab-stop formatting with a tab character", () => {
		function tabStyledControl(value) {
			return (
				'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>' +
				`<w:r><w:t>${escapeXml(value)}</w:t></w:r></w:p>`
			);
		}
		const template = [
			tabStyledControl("{!switch}"),
			tabStyledControl("{!case enabled}"),
			paragraph("YES"),
			tabStyledControl("{!default}"),
			paragraph("NO"),
			tabStyledControl("{/!switch}"),
		].join("");
		const doc = createDoc(template);
		doc.render({ enabled: true });
		const xml = renderedPart(doc);
		expect(xmlText(xml)).to.equal("YES");
		expect((xml.match(/<w:p>/g) || []).length).to.equal(1);
		expect(xml).to.not.include("<w:tabs");
		expectWellFormedXml(xml);
	});

	it("supports styled multiline switches in PowerPoint", () => {
		const templateParagraphs = [
			"{!switch}",
			"{!case enabled}",
			"YES",
			"{!default}",
			"NO",
			"{/!switch}",
		];

		for (const [enabled, expected] of [
			[true, "YES"],
			[false, "NO"],
		]) {
			const doc = createPptxDoc(templateParagraphs);
			expect(doc.fileType).to.equal("pptx");
			doc.render({ enabled });
			const xml = renderedPart(doc, "ppt/slides/slide1.xml");
			expect(doc.getFullText()).to.equal(expected);
			expect((xml.match(/<a:p(?:\s[^>]*)?>/g) || []).length).to.equal(1);
			expect((xml.match(/<a:pPr(?:\s[^>]*)?>/g) || []).length).to.equal(
				1
			);
			expect(
				(xml.match(/<a:endParaRPr(?:\s[^>]*)?>/g) || []).length
			).to.equal(1);
			expectWellFormedXml(xml);
		}
	});

	it("keeps a styled PowerPoint break in its original branch", () => {
		const paragraphXml =
			'<a:p><a:pPr algn="ctr"/>' +
			'<a:r><a:rPr lang="en-US"/><a:t>{!switch}{!case first}A</a:t></a:r>' +
			'<a:br><a:rPr lang="fr-FR" sz="3200"/></a:br>' +
			'<a:r><a:rPr b="1" lang="en-US"/><a:t>{!case second}B{/!switch}</a:t></a:r>' +
			'<a:endParaRPr lang="en-US"/></a:p>';

		const firstDoc = createPptxDoc([], () => paragraphXml);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc, "ppt/slides/slide1.xml");
		expect(firstDoc.getFullText()).to.equal("A");
		expect(firstXml).to.include(
			'<a:br><a:rPr lang="fr-FR" sz="3200"/></a:br>'
		);
		expectWellFormedXml(firstXml);

		const secondDoc = createPptxDoc([], () => paragraphXml);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc, "ppt/slides/slide1.xml");
		expect(secondDoc.getFullText()).to.equal("B");
		expect(secondXml).to.not.include("<a:br");
		expect(secondXml).to.not.include('lang="fr-FR"');
		expect(runContaining(secondXml, "B")).to.include('b="1"');
		expectWellFormedXml(secondXml);
	});

	it("keeps a complete PowerPoint field in its original branch", () => {
		const paragraphXml =
			"<a:p><a:pPr/>" +
			"<a:r><a:t>{!switch}{!case first}A</a:t></a:r>" +
			'<a:fld id="{00000000-0000-0000-0000-000000000001}" type="datetime"><a:rPr lang="en-US"/><a:t>FIELD</a:t></a:fld>' +
			'<a:r><a:rPr i="1"/><a:t>{!case second}B{/!switch}</a:t></a:r>' +
			"<a:endParaRPr/></a:p>";

		const firstDoc = createPptxDoc([], () => paragraphXml);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc, "ppt/slides/slide1.xml");
		expect(firstDoc.getFullText()).to.equal("AFIELD");
		expect(firstXml).to.include("<a:fld ");
		expectWellFormedXml(firstXml);

		const secondDoc = createPptxDoc([], () => paragraphXml);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc, "ppt/slides/slide1.xml");
		expect(secondDoc.getFullText()).to.equal("B");
		expect(secondXml).to.not.include("<a:fld ");
		expect(runContaining(secondXml, "B")).to.include('i="1"');
		expectWellFormedXml(secondXml);
	});

	it("keeps a semantic later-case PowerPoint paragraph with that branch", () => {
		function makeTemplate(paragraphTemplate) {
			function controlParagraph(value) {
				return paragraphTemplate.replace(
					"Hello {name}",
					escapeXml(value)
				);
			}
			const laterCase = controlParagraph("{!case second}").replace(
				"</a:r><a:endParaRPr",
				'</a:r><a:br><a:rPr lang="fr-FR" sz="3200"/></a:br><a:endParaRPr'
			);
			return [
				controlParagraph("{!switch}"),
				controlParagraph("{!case first}"),
				controlParagraph("A"),
				laterCase,
				controlParagraph("B"),
				controlParagraph("{/!switch}"),
			].join("");
		}

		const firstDoc = createPptxDoc([], makeTemplate);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc, "ppt/slides/slide1.xml");
		expect(firstDoc.getFullText()).to.equal("A");
		expect(firstXml).to.not.include("<a:br");
		expect((firstXml.match(/<a:p(?:\s[^>]*)?>/g) || []).length).to.equal(1);
		expectWellFormedXml(firstXml);

		const secondDoc = createPptxDoc([], makeTemplate);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc, "ppt/slides/slide1.xml");
		expect(secondDoc.getFullText()).to.equal("B");
		expect(secondXml).to.include(
			'<a:br><a:rPr lang="fr-FR" sz="3200"/></a:br>'
		);
		expect((secondXml.match(/<a:p(?:\s[^>]*)?>/g) || []).length).to.equal(
			2
		);
		expect(
			directElementChildNames(secondXml, "p:txBody").every((name) =>
				["a:bodyPr", "a:lstStyle", "a:p"].includes(name)
			)
		).to.equal(true);
		expectWellFormedXml(secondXml);
	});

	it("does not leak a prior PowerPoint paragraph into a later branch", () => {
		function makeTemplate(paragraphTemplate) {
			function controlParagraph(value) {
				return paragraphTemplate.replace(
					"Hello {name}",
					escapeXml(value)
				);
			}
			const priorBreak = controlParagraph("").replace(
				"</a:r><a:endParaRPr",
				'</a:r><a:br><a:rPr lang="de-DE" sz="2800"/></a:br><a:endParaRPr'
			);
			const laterCase = controlParagraph("{!case second}").replace(
				"</a:r><a:endParaRPr",
				'</a:r><a:br><a:rPr lang="fr-FR" sz="3200"/></a:br><a:endParaRPr'
			);
			return [
				controlParagraph("{!switch}"),
				controlParagraph("{!case first}"),
				controlParagraph("A"),
				priorBreak,
				laterCase,
				controlParagraph("B"),
				controlParagraph("{/!switch}"),
			].join("");
		}

		const firstDoc = createPptxDoc([], makeTemplate);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc, "ppt/slides/slide1.xml");
		expect(firstDoc.getFullText()).to.equal("A");
		expect(firstXml).to.include('lang="de-DE"');
		expect(firstXml).to.not.include('lang="fr-FR"');
		expect((firstXml.match(/<a:p(?:\s[^>]*)?>/g) || []).length).to.equal(2);
		expectWellFormedXml(firstXml);

		const secondDoc = createPptxDoc([], makeTemplate);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc, "ppt/slides/slide1.xml");
		expect(secondDoc.getFullText()).to.equal("B");
		expect(secondXml).to.not.include('lang="de-DE"');
		expect(secondXml).to.include('lang="fr-FR"');
		expect((secondXml.match(/<a:p(?:\s[^>]*)?>/g) || []).length).to.equal(
			2
		);
		expect(
			directElementChildNames(secondXml, "p:txBody").every((name) =>
				["a:bodyPr", "a:lstStyle", "a:p"].includes(name)
			)
		).to.equal(true);
		expectWellFormedXml(secondXml);
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

	it("supports nested multiline switches with control-only paragraphs", () => {
		const template = [
			paragraph("{!switch}"),
			paragraph("{!case outer}"),
			paragraph("{!switch}"),
			paragraph("{!case inner}"),
			paragraph("INNER"),
			paragraph("{!default}"),
			paragraph("INNER DEFAULT"),
			paragraph("{/!switch}"),
			paragraph("{!default}"),
			paragraph("OUTER DEFAULT"),
			paragraph("{/!switch}"),
		].join("");
		const doc = createDoc(template);
		doc.render({ outer: true, inner: false });
		const xml = renderedPart(doc);
		expect(xmlText(xml)).to.equal("INNER DEFAULT");
		expect((xml.match(/<w:p>/g) || []).length).to.equal(1);
	});

	it("removes control paragraphs nested inside a Word text box", () => {
		const textBoxContent = [
			paragraph("{!switch}"),
			paragraph("{!case enabled}"),
			paragraph("YES"),
			paragraph("{!default}"),
			paragraph("NO"),
			paragraph("{/!switch}"),
		].join("");
		const template =
			"<w:p><w:r><w:drawing><w:txbxContent>" +
			textBoxContent +
			"</w:txbxContent></w:drawing></w:r></w:p>";

		for (const [enabled, expected] of [
			[true, "YES"],
			[false, "NO"],
		]) {
			const doc = createDoc(template);
			doc.render({ enabled });
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect((xml.match(/<w:p>/g) || []).length).to.equal(2);
			expect((xml.match(/<w:drawing>/g) || []).length).to.equal(1);
			expectWellFormedXml(xml);
		}
	});

	it("supports structurally balanced mixed inline and multiline branches", () => {
		const template = [
			paragraph("{!switch}"),
			paragraph("{!case first}FIRST"),
			paragraph("{!case second}"),
			paragraph("SECOND"),
			paragraph("{/!switch}"),
		].join("");

		for (const [data, expected] of [
			[{ first: true, second: true }, "FIRST"],
			[{ first: false, second: true }, "SECOND"],
		]) {
			const doc = createDoc(template);
			doc.render(data);
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect((xml.match(/<w:p>/g) || []).length).to.equal(1);
			expectWellFormedXml(xml);
		}
	});

	it("rejects mixed layouts that cannot produce balanced XML", () => {
		const template = [
			paragraph("{!switch}"),
			paragraph("{!case enabled}YES{!default}NO{/!switch}"),
			paragraph("After"),
		].join("");
		let thrown;
		try {
			createDoc(template);
		} catch (error) {
			thrown = error;
		}
		expect(thrown).to.not.equal(undefined);
		const invalid = getErrors(thrown).find(
			(error) =>
				error.properties.id === "conditional_switch_invalid_structure"
		);
		expect(invalid).to.not.equal(undefined);
		expect(invalid.name).to.equal("TemplateError");
		expect(invalid.properties.xtag).to.equal("!switch");
		expect(invalid.properties.file).to.equal("word/document.xml");
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

	it("uses the same parser context for render and renderAsync", async () => {
		const parser = expressionParser.configure({
			postEvaluate(value, tag, scope, context) {
				if (tag !== "name") {
					return value;
				}
				const path = context.scopePathItem;
				return `${path[path.length - 1]}:${value}`;
			},
		});
		const template = paragraph(
			"{#items}{!switch}{!case enabled}{name}{!default}off{/!switch};{/items}"
		);
		const data = {
			items: [
				{ name: "A", enabled: true },
				{ name: "B", enabled: true },
			],
		};
		const syncDoc = createDoc(template, { parser });
		syncDoc.render(data);
		const asyncDoc = createDoc(template, { parser });
		await asyncDoc.renderAsync(data);
		const syncText = xmlText(renderedPart(syncDoc));
		const asyncText = xmlText(renderedPart(asyncDoc));
		expect(syncText).to.equal("0:A;1:B;");
		expect(asyncText).to.equal(syncText);
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

	it("allows a switch contained by one hyperlink", () => {
		const body =
			'<w:p><w:hyperlink w:anchor="target"><w:r><w:t>{!switch}{!case enabled}YES{!default}NO{/!switch}</w:t></w:r></w:hyperlink></w:p>';
		for (const [enabled, expected] of [
			[true, "YES"],
			[false, "NO"],
		]) {
			const doc = createDoc(body);
			doc.render({ enabled });
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect(xml).to.include('<w:hyperlink w:anchor="target">');
			expectWellFormedXml(xml);
		}
	});

	it("rejects switches that cross hyperlink containers", () => {
		const body =
			'<w:p><w:hyperlink w:anchor="target"><w:r><w:t>{!switch}{!case enabled}YES</w:t></w:r></w:hyperlink>' +
			"<w:r><w:t>{!default}NO{/!switch}</w:t></w:r></w:p>";
		let thrown;
		try {
			createDoc(body);
		} catch (error) {
			thrown = error;
		}
		expect(thrown).to.not.equal(undefined);
		const invalid = getErrors(thrown).find(
			(error) =>
				error.properties.id === "conditional_switch_invalid_structure"
		);
		expect(invalid).to.not.equal(undefined);
		expect(invalid.properties.xtag).to.equal("!switch");
		expect(invalid.properties.file).to.equal("word/document.xml");
	});

	it("rejects switches that cross table cells", () => {
		const body =
			`<w:tbl><w:tr><w:tc>${paragraph(
				"{!switch}{!case enabled}YES"
			)}</w:tc>` +
			`<w:tc>${paragraph(
				"{!default}NO{/!switch}"
			)}</w:tc></w:tr></w:tbl>`;
		let thrown;
		try {
			createDoc(body);
		} catch (error) {
			thrown = error;
		}
		expect(thrown).to.not.equal(undefined);
		expect(
			getErrors(thrown).some(
				(error) =>
					error.properties.id ===
					"conditional_switch_invalid_structure"
			)
		).to.equal(true);
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

	it("preserves the run formatting of a later default branch", () => {
		const body =
			"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{!switch}{!case enabled}YES</w:t></w:r>" +
			"<w:r><w:rPr><w:i/></w:rPr><w:t>{!default}NO{/!switch}</w:t></w:r></w:p>";
		const doc = createDoc(body);
		doc.render({ enabled: false });
		const xml = renderedPart(doc);
		const selectedRun = runContaining(xml, "NO");
		expect(xmlText(xml)).to.equal("NO");
		expect(selectedRun).to.be.a("string");
		expect(selectedRun).to.include("<w:i/>");
		expect(selectedRun).to.not.include("<w:b/>");
	});

	it("restores the physical post-switch Word run", () => {
		const body =
			'<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr>' +
			"<w:t>PRE{!switch}{!case first}A</w:t></w:r>" +
			"<w:r><w:rPr><w:b/></w:rPr><w:t>{!case second}B</w:t></w:r>" +
			"<w:r><w:rPr><w:i/></w:rPr><w:t>{/!switch}POST</w:t></w:r></w:p>";

		for (const [data, expected, expectedRuns] of [
			[{ first: true, second: false }, "PREAPOST", 2],
			[{ first: false, second: true }, "PREBPOST", 3],
			[{ first: false, second: false }, "PREPOST", 2],
		]) {
			const doc = createDoc(body);
			doc.render(data);
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect((xml.match(/<w:r>/g) || []).length).to.equal(expectedRuns);
			expect(runContaining(xml, "PRE")).to.include(
				'<w:color w:val="FF0000"/>'
			);
			expect(runContaining(xml, "POST")).to.include("<w:i/>");
			expectWellFormedXml(xml);
		}
	});

	it("restores the physical post-switch Word paragraph", () => {
		const body =
			'<w:p><w:pPr><w:pStyle w:val="One"/></w:pPr>' +
			"<w:r><w:t>PRE{!switch}{!case first}A</w:t></w:r></w:p>" +
			'<w:p><w:pPr><w:pStyle w:val="Two"/></w:pPr>' +
			"<w:r><w:t>{!case second}B{/!switch}POST</w:t></w:r></w:p>";

		for (const [data, expected] of [
			[{ first: true, second: false }, "PREAPOST"],
			[{ first: false, second: true }, "PREBPOST"],
			[{ first: false, second: false }, "PREPOST"],
		]) {
			const doc = createDoc(body);
			doc.render(data);
			const xml = renderedPart(doc);
			const paragraphs = xml.match(/<w:p>[\s\S]*?<\/w:p>/g) || [];
			expect(xmlText(xml)).to.equal(expected);
			expect(paragraphs).to.have.length(2);
			expect(paragraphs[0]).to.include('<w:pStyle w:val="One"/>');
			expect(paragraphs[1]).to.include('<w:pStyle w:val="Two"/>');
			expect(xmlText(paragraphs[1])).to.include("POST");
			expectWellFormedXml(xml);
		}
	});

	it("restores the physical post-switch PowerPoint run", () => {
		const paragraphXml =
			"<a:p><a:pPr/>" +
			'<a:r><a:rPr i="1"/><a:t>PRE{!switch}{!case first}A</a:t></a:r>' +
			'<a:r><a:rPr b="1"/><a:t>{!case second}B</a:t></a:r>' +
			'<a:r><a:rPr u="sng"/><a:t>{/!switch}POST</a:t></a:r>' +
			"<a:endParaRPr/></a:p>";

		for (const [data, expected, expectedRuns] of [
			[{ first: true, second: false }, "PREAPOST", 2],
			[{ first: false, second: true }, "PREBPOST", 3],
			[{ first: false, second: false }, "PREPOST", 2],
		]) {
			const doc = createPptxDoc([], () => paragraphXml);
			doc.render(data);
			const xml = renderedPart(doc, "ppt/slides/slide1.xml");
			expect(doc.getFullText()).to.equal(expected);
			expect((xml.match(/<a:r(?:\s[^>]*)?>/g) || []).length).to.equal(
				expectedRuns
			);
			expect(runContaining(xml, "PRE")).to.include('i="1"');
			expect(runContaining(xml, "POST")).to.include('u="sng"');
			expectWellFormedXml(xml);
		}
	});

	it("preserves a page break in a selected control paragraph", () => {
		const body = [
			paragraph("{!switch}"),
			'<w:p><w:r><w:t>{!case enabled}</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>',
			paragraph("{!default}"),
			paragraph("No break"),
			paragraph("{/!switch}"),
		].join("");
		const doc = createDoc(body);
		doc.render({ enabled: true });
		const xml = renderedPart(doc);
		expect(xml).to.include('<w:br w:type="page"/>');
		expectWellFormedXml(xml);
	});

	it("preserves generated Word run content in a selected branch", () => {
		const generatedContent =
			"<w:p><w:r><w:t>{!case enabled}</w:t>" +
			"<w:pgNum/><w:footnoteRef/><w:endnoteRef/><w:dayShort/>" +
			'<w:instrText>DATE</w:instrText><w:contentPart r:id="rIdContent"/>' +
			'</w:r><w:subDoc r:id="rIdSubDoc"/></w:p>';
		const body = [
			paragraph("{!switch}"),
			generatedContent,
			paragraph("Selected"),
			paragraph("{!default}"),
			paragraph("Other"),
			paragraph("{/!switch}"),
		].join("");

		const selectedDoc = createDoc(body);
		selectedDoc.render({ enabled: true });
		const selectedXml = renderedPart(selectedDoc);
		for (const tag of [
			"w:pgNum",
			"w:footnoteRef",
			"w:endnoteRef",
			"w:dayShort",
			"w:instrText",
			"w:contentPart",
			"w:subDoc",
		]) {
			expect(selectedXml).to.include(`<${tag}`);
		}
		expect(xmlText(selectedXml)).to.equal("DATESelected");
		expectWellFormedXml(selectedXml);

		const otherDoc = createDoc(body);
		otherDoc.render({ enabled: false });
		const otherXml = renderedPart(otherDoc);
		expect(xmlText(otherXml)).to.equal("Other");
		for (const tag of [
			"w:pgNum",
			"w:footnoteRef",
			"w:endnoteRef",
			"w:dayShort",
			"w:instrText",
			"w:contentPart",
			"w:subDoc",
		]) {
			expect(otherXml).to.not.include(`<${tag}`);
		}
		expectWellFormedXml(otherXml);
	});

	it("keeps generated page numbers and run tabs in their physical branch", () => {
		const body =
			"<w:p><w:r><w:t>{!switch}{!case first}A</w:t>" +
			"<w:pgNum/><w:tab/><w:t>{!case second}B{/!switch}</w:t></w:r></w:p>";

		const firstDoc = createDoc(body);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc);
		expect(xmlText(firstXml)).to.equal("A");
		expect(firstXml).to.include("<w:pgNum/>");
		expect(firstXml).to.include("<w:tab/>");
		expectWellFormedXml(firstXml);

		const secondDoc = createDoc(body);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc);
		expect(xmlText(secondXml)).to.equal("B");
		expect(secondXml).to.not.include("<w:pgNum");
		expect(secondXml).to.not.include("<w:tab");
		expectWellFormedXml(secondXml);
	});

	it("retains tab-stop formatting for a later branch", () => {
		const body =
			"<w:p><w:r><w:t>{!switch}{!case first}A</w:t></w:r></w:p>" +
			'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>' +
			"<w:r><w:t>{!case second}B{/!switch}</w:t></w:r></w:p>";
		const doc = createDoc(body);
		doc.render({ first: false, second: true });
		const xml = renderedPart(doc);
		expect(xmlText(xml)).to.equal("B");
		expect(xml).to.include(
			'<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>'
		);
		expectWellFormedXml(xml);
	});

	it("preserves page-break-before only in its selected branch", () => {
		const pageBreakCase =
			"<w:p><w:pPr><w:pageBreakBefore/></w:pPr>" +
			"<w:r><w:t>{!case enabled}</w:t></w:r></w:p>";
		const body = [
			paragraph("{!switch}"),
			pageBreakCase,
			paragraph("Selected"),
			paragraph("{!default}"),
			paragraph("Other"),
			paragraph("{/!switch}"),
		].join("");

		const selectedDoc = createDoc(body);
		selectedDoc.render({ enabled: true });
		const selectedXml = renderedPart(selectedDoc);
		expect(xmlText(selectedXml)).to.equal("Selected");
		expect(selectedXml).to.include("<w:pageBreakBefore/>");
		expectWellFormedXml(selectedXml);

		const otherDoc = createDoc(body);
		otherDoc.render({ enabled: false });
		const otherXml = renderedPart(otherDoc);
		expect(xmlText(otherXml)).to.equal("Other");
		expect(otherXml).to.not.include("<w:pageBreakBefore");
		expectWellFormedXml(otherXml);
	});

	it("keeps a semantic later-case Word paragraph with that branch", () => {
		const laterCase =
			"<w:p><w:pPr><w:pageBreakBefore/>" +
			'<w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr>' +
			"<w:r><w:pgNum/><w:t>{!case second}</w:t>" +
			'<w:br w:type="page"/></w:r></w:p>';
		const body = [
			paragraph("{!switch}"),
			paragraph("{!case first}"),
			paragraph("A"),
			laterCase,
			paragraph("B"),
			paragraph("{/!switch}"),
		].join("");

		const firstDoc = createDoc(body);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc);
		expect(xmlText(firstXml)).to.equal("A");
		for (const tag of ["w:pageBreakBefore", "w:type", "w:pgNum", "w:br"]) {
			expect(firstXml).to.not.include(`<${tag}`);
		}
		expect((firstXml.match(/<w:p>/g) || []).length).to.equal(1);
		expectWellFormedXml(firstXml);

		const secondDoc = createDoc(body);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc);
		expect(xmlText(secondXml)).to.equal("B");
		expect(secondXml).to.include("<w:pageBreakBefore/>");
		expect(secondXml).to.include('<w:type w:val="continuous"/>');
		expect(secondXml).to.include("<w:pgNum/>");
		expect(secondXml).to.include('<w:br w:type="page"/>');
		expect((secondXml.match(/<w:p>/g) || []).length).to.equal(2);
		expectWellFormedXml(secondXml);
	});

	it("does not leak a prior Word run into a later branch", () => {
		const priorPayload =
			"<w:p><w:r><w:t>A</w:t></w:r>" +
			"<w:r><w:rPr><w:b/></w:rPr><w:br/></w:r></w:p>";
		const laterCase =
			"<w:p><w:pPr><w:pageBreakBefore/></w:pPr>" +
			"<w:r><w:t>{!case second}</w:t></w:r></w:p>";
		const body = [
			paragraph("{!switch}"),
			paragraph("{!case first}"),
			priorPayload,
			laterCase,
			paragraph("B"),
			paragraph("{/!switch}"),
		].join("");

		const firstDoc = createDoc(body);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc);
		expect(xmlText(firstXml)).to.equal("A");
		expect(firstXml).to.include("<w:b/>");
		expect(firstXml).to.include("<w:br/>");
		expect((firstXml.match(/<w:p>/g) || []).length).to.equal(1);
		expectWellFormedXml(firstXml);

		const secondDoc = createDoc(body);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc);
		expect(xmlText(secondXml)).to.equal("B");
		expect(secondXml).to.not.include("<w:b/>");
		expect(secondXml).to.not.include("<w:br/>");
		expect((secondXml.match(/<w:p>/g) || []).length).to.equal(2);
		expect(
			directElementChildNames(secondXml, "w:body").every((name) =>
				["w:p", "w:sectPr"].includes(name)
			)
		).to.equal(true);
		expectWellFormedXml(secondXml);
	});

	it("keeps a layout-only paragraph only in its physical branch", () => {
		const priorLayout =
			"<w:p><w:pPr><w:keepNext/>" +
			'<w:spacing w:before="999"/></w:pPr></w:p>';
		const laterCase =
			"<w:p><w:pPr><w:pageBreakBefore/></w:pPr>" +
			"<w:r><w:t>{!case second}</w:t></w:r></w:p>";
		const body = [
			paragraph("{!switch}"),
			paragraph("{!case first}"),
			paragraph("A"),
			priorLayout,
			laterCase,
			paragraph("B"),
			paragraph("{/!switch}"),
		].join("");

		const firstDoc = createDoc(body);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc);
		expect(xmlText(firstXml)).to.equal("A");
		expect(firstXml).to.include("<w:keepNext/>");
		expect(firstXml).to.include('<w:spacing w:before="999"/>');
		expect((firstXml.match(/<w:p>/g) || []).length).to.equal(2);
		expectWellFormedXml(firstXml);

		const secondDoc = createDoc(body);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc);
		expect(xmlText(secondXml)).to.equal("B");
		expect(secondXml).to.not.include("<w:keepNext");
		expect(secondXml).to.not.include("<w:spacing");
		expect((secondXml.match(/<w:p>/g) || []).length).to.equal(2);
		expectWellFormedXml(secondXml);
	});

	it("keeps breaks and drawings in the branch that contains them", () => {
		const body =
			"<w:p><w:r><w:t>{!switch}{!case first}A</w:t></w:r>" +
			'<w:r><w:br/><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/></w:drawing></w:r>' +
			"<w:r><w:rPr><w:i/></w:rPr><w:t>{!case second}B{/!switch}</w:t></w:r></w:p>";

		const firstDoc = createDoc(body);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc);
		expect(xmlText(firstXml)).to.equal("A");
		expect(firstXml).to.include("<w:br/>");
		expect(firstXml).to.include("<w:drawing>");
		expectWellFormedXml(firstXml);

		const secondDoc = createDoc(body);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc);
		const selectedRun = runContaining(secondXml, "B");
		expect(xmlText(secondXml)).to.equal("B");
		expect(secondXml).to.not.include("<w:br/>");
		expect(secondXml).to.not.include("<w:drawing>");
		expect(selectedRun).to.include("<w:i/>");
		expectWellFormedXml(secondXml);
	});

	it("keeps complete inline content containers in their original branch", () => {
		const body =
			"<w:p><w:r><w:t>{!switch}{!case first}A</w:t></w:r>" +
			'<w:hyperlink w:anchor="target"><w:r><w:t>LINK</w:t></w:r></w:hyperlink>' +
			'<w:ins w:id="7"><w:r><w:t>INSERTED</w:t></w:r></w:ins>' +
			'<w:fldSimple w:instr="DATE"><w:r><w:t>FIELD</w:t></w:r></w:fldSimple>' +
			'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>MATH</m:t></m:r></m:oMath>' +
			"<w:r><w:rPr><w:i/></w:rPr><w:t>{!case second}B{/!switch}</w:t></w:r></w:p>";

		const firstDoc = createDoc(body);
		firstDoc.render({ first: true, second: false });
		const firstXml = renderedPart(firstDoc);
		expect(xmlText(firstXml)).to.equal("ALINKINSERTEDFIELDMATH");
		for (const tag of ["w:hyperlink", "w:ins", "w:fldSimple", "m:oMath"]) {
			expect(firstXml).to.include(`<${tag}`);
		}
		expectWellFormedXml(firstXml);

		const secondDoc = createDoc(body);
		secondDoc.render({ first: false, second: true });
		const secondXml = renderedPart(secondDoc);
		expect(xmlText(secondXml)).to.equal("B");
		for (const tag of ["w:hyperlink", "w:ins", "w:fldSimple", "m:oMath"]) {
			expect(secondXml).to.not.include(`<${tag}`);
		}
		expect(runContaining(secondXml, "B")).to.include("<w:i/>");
		expectWellFormedXml(secondXml);
	});

	it("preserves Office Math run properties in a later branch", () => {
		const body =
			'<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
			"<m:r><m:t>{!switch}{!case first}A</m:t></m:r>" +
			'<m:r><m:rPr><m:sty m:val="b"/></m:rPr>' +
			"<m:t>{!case second}B{/!switch}POST</m:t></m:r></m:oMath></w:p>";

		for (const [data, expected] of [
			[{ first: true, second: false }, "APOST"],
			[{ first: false, second: true }, "BPOST"],
		]) {
			const doc = createDoc(body);
			doc.render(data);
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect(xml).to.include('<m:rPr><m:sty m:val="b"/></m:rPr>');
			expectWellFormedXml(xml);
		}
	});

	it("preserves structural XML mixed with visible branch markers", () => {
		const body =
			"<w:p><w:r><w:t>{!switch}{!case first}A</w:t></w:r>" +
			'<w:proofErr w:type="spellStart"/><w:proofErr w:type="spellEnd"/>' +
			'<w:hyperlink w:anchor="target"><w:r></w:r></w:hyperlink>' +
			"<w:r><w:t>{!case second}B{/!switch}</w:t></w:r></w:p>";

		for (const [data, expected, proofErrors, hyperlinks] of [
			[{ first: true, second: false }, "A", 2, 1],
			[{ first: false, second: true }, "B", 0, 0],
		]) {
			const doc = createDoc(body);
			doc.render(data);
			const xml = renderedPart(doc);
			expect(xmlText(xml)).to.equal(expected);
			expect((xml.match(/<w:proofErr\b/g) || []).length).to.equal(
				proofErrors
			);
			expect((xml.match(/<w:hyperlink\b/g) || []).length).to.equal(
				hyperlinks
			);
			expectWellFormedXml(xml);
		}
	});

	it("preserves matching bookmark tags in a selected branch", () => {
		const body = [
			paragraph("{!switch}"),
			'<w:p><w:r><w:t>{!case enabled}</w:t></w:r><w:bookmarkStart w:id="7" w:name="conditional"/></w:p>',
			'<w:p><w:r><w:t>Bookmarked</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
			paragraph("{!default}"),
			paragraph("No bookmark"),
			paragraph("{/!switch}"),
		].join("");
		const doc = createDoc(body);
		doc.render({ enabled: true });
		const xml = renderedPart(doc);
		expect(xmlText(xml)).to.equal("Bookmarked");
		expect(xml).to.include(
			'<w:bookmarkStart w:id="7" w:name="conditional"/>'
		);
		expect(xml).to.include('<w:bookmarkEnd w:id="7"/>');
		expectWellFormedXml(xml);
	});

	it("preserves a comment reference in a selected branch", () => {
		const body = [
			paragraph("{!switch}"),
			'<w:p><w:r><w:t>{!case enabled}</w:t></w:r><w:r><w:commentReference w:id="7"/></w:r></w:p>',
			paragraph("Selected"),
			paragraph("{!default}"),
			paragraph("Other"),
			paragraph("{/!switch}"),
		].join("");
		const doc = createDoc(body);
		doc.render({ enabled: true });
		const xml = renderedPart(doc);
		expect(xmlText(xml)).to.equal("Selected");
		expect(xml).to.include('<w:commentReference w:id="7"/>');
		expect((xml.match(/<w:p>/g) || []).length).to.equal(2);
		expectWellFormedXml(xml);
	});

	it("preserves section and permission markers only in their branch", () => {
		const markedCase =
			'<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr>' +
			'<w:permStart w:id="9" w:edGrp="everyone"/><w:r><w:t>{!case enabled}</w:t></w:r><w:permEnd w:id="9"/></w:p>';
		const body = [
			paragraph("{!switch}"),
			markedCase,
			paragraph("Selected"),
			paragraph("{!default}"),
			paragraph("Other"),
			paragraph("{/!switch}"),
		].join("");

		const selectedDoc = createDoc(body);
		selectedDoc.render({ enabled: true });
		const selectedXml = renderedPart(selectedDoc);
		expect(xmlText(selectedXml)).to.equal("Selected");
		expect(selectedXml).to.include('<w:type w:val="continuous"/>');
		expect(selectedXml).to.include(
			'<w:permStart w:id="9" w:edGrp="everyone"/>'
		);
		expect(selectedXml).to.include('<w:permEnd w:id="9"/>');
		expectWellFormedXml(selectedXml);

		const otherDoc = createDoc(body);
		otherDoc.render({ enabled: false });
		const otherXml = renderedPart(otherDoc);
		expect(xmlText(otherXml)).to.equal("Other");
		expect(otherXml).to.not.include('<w:type w:val="continuous"/>');
		expect(otherXml).to.not.include("<w:permStart");
		expect(otherXml).to.not.include("<w:permEnd");
		expectWellFormedXml(otherXml);
	});

	it("compiles each placeholder in branch content once", () => {
		let compilations = 0;
		function parser(tag, meta) {
			if (tag === "value") {
				compilations++;
			}
			return expressionParser(tag, meta);
		}
		createDoc(
			paragraph(
				"{!switch}{!case enabled}{value}{!default}fallback{/!switch}"
			),
			{ parser }
		);
		expect(compilations).to.equal(1);
	});

	it("reports an invalid branch placeholder only once", () => {
		let thrown;
		try {
			createDoc(
				paragraph(
					"{!switch}{!case enabled}{customer.}{!default}fallback{/!switch}"
				)
			);
		} catch (error) {
			thrown = error;
		}
		expect(thrown).to.not.equal(undefined);
		const compilationErrors = getErrors(thrown).filter(
			(error) =>
				error.properties.id === "scopeparser_compilation_failed" &&
				error.properties.xtag === "customer."
		);
		expect(compilationErrors).to.have.length(1);
	});

	it("keeps parser error attribution isolated between document parts", () => {
		let thrown;
		try {
			createDoc(paragraph("123456789{broken.}"), {
				parts: {
					header: paragraph(
						"{!switch}{!case headerCondition}H{/!switch}"
					),
				},
			});
		} catch (error) {
			thrown = error;
		}
		expect(thrown).to.not.equal(undefined);
		const bodyError = getErrors(thrown).find(
			(error) => error.properties.file === "word/document.xml"
		);
		expect(bodyError).to.not.equal(undefined);
		expect(bodyError.properties.id).to.equal(
			"scopeparser_compilation_failed"
		);
		expect(bodyError.properties.xtag).to.equal("broken.");
	});

	it("accepts outer whitespace around case tags", () => {
		expect(
			renderText(
				paragraph(
					"{ !switch }{ !case enabled }YES{ !default }NO{ /!switch }"
				),
				{ enabled: true }
			)
		).to.equal("YES");
	});

	it("allows a case expression matching the former empty-case sentinel", () => {
		const identifier = "__EMPTY_CONDITIONAL_SWITCH_CASE__";
		expect(
			renderText(
				paragraph(
					`{!switch}{!case ${identifier}}YES{!default}NO{/!switch}`
				),
				{ [identifier]: true }
			)
		).to.equal("YES");
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
