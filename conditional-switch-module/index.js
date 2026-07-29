"use strict";

const MODULE_NAME = "private-conditional-switch";

const RELOCATABLE_XML_CONTAINERS = new Set([
	"w:p",
	"w:r",
	"w:t",
	"a:p",
	"a:r",
	"a:t",
	"m:r",
	"m:t",
]);

const PROPERTY_XML_PARENTS = new Map([
	["a:endParaRPr", new Set(["a:p"])],
	["a:pPr", new Set(["a:fld", "a:p"])],
	["a:rPr", new Set(["a:br", "a:fld", "a:r"])],
	["m:oMathParaPr", new Set(["m:oMathPara"])],
	["m:rPr", new Set(["m:r"])],
	["w:pPr", new Set(["w:p"])],
	["w:rPr", new Set(["w:r"])],
]);

const PARAGRAPH_SEMANTIC_XML_TAGS = new Set([
	"a:br",
	"a:fld",
	"a14:m",
	"w:annotationRef",
	"w:bookmarkEnd",
	"w:bookmarkStart",
	"w:br",
	"w:commentRangeEnd",
	"w:commentRangeStart",
	"w:commentReference",
	"w:contentPart",
	"w:continuationSeparator",
	"w:cr",
	"w:customXmlDelRangeEnd",
	"w:customXmlDelRangeStart",
	"w:customXmlInsRangeEnd",
	"w:customXmlInsRangeStart",
	"w:customXmlMoveFromRangeEnd",
	"w:customXmlMoveFromRangeStart",
	"w:customXmlMoveToRangeEnd",
	"w:customXmlMoveToRangeStart",
	"w:dayLong",
	"w:dayShort",
	"w:delInstrText",
	"w:delText",
	"w:drawing",
	"w:endnoteRef",
	"w:endnoteReference",
	"w:fldChar",
	"w:fldSimple",
	"w:footnoteRef",
	"w:footnoteReference",
	"w:instrText",
	"w:monthLong",
	"w:monthShort",
	"w:moveFromRangeEnd",
	"w:moveFromRangeStart",
	"w:moveToRangeEnd",
	"w:moveToRangeStart",
	"w:noBreakHyphen",
	"w:object",
	"w:permEnd",
	"w:permStart",
	"w:pict",
	"w:pgNum",
	"w:ptab",
	"w:ruby",
	"w:separator",
	"w:sectPr",
	"w:softHyphen",
	"w:subDoc",
	"w:sym",
	"w:yearLong",
	"w:yearShort",
]);

function exactTag(expected) {
	return (tag) => (tag.trim() === expected ? "." : false);
}

function caseTag(tag) {
	const match = tag.trim().match(/^!case(?:\s+([\s\S]*))?$/);
	if (!match) {
		return false;
	}
	const expression = (match[1] || "").trim();
	return { expression, emptyCase: expression === "" };
}

function makeTemplateError(id, message, part, extra = {}) {
	const error = new Error(message);
	error.name = "TemplateError";
	error.properties = {
		id,
		explanation: message,
		xtag: part && part.raw,
		offset: part && part.offset,
		lIndex: part && part.lIndex,
		...extra,
	};
	return error;
}

function isControl(part, kind) {
	return (
		part &&
		part.type === "placeholder" &&
		part.module === MODULE_NAME &&
		(kind == null || part.switchKind === kind)
	);
}

function isParagraphStart(part) {
	return (
		part &&
		part.type === "tag" &&
		part.position === "start" &&
		(part.tag === "w:p" || part.tag === "a:p")
	);
}

function isParagraphEnd(part) {
	return (
		part &&
		part.type === "tag" &&
		part.position === "end" &&
		(part.tag === "w:p" || part.tag === "a:p")
	);
}

function findXmlTagEnd(value, start) {
	let quote = null;
	for (let index = start + 1; index < value.length; index++) {
		const character = value[index];
		if (quote) {
			if (character === quote) {
				quote = null;
			}
			continue;
		}
		if (character === '"' || character === "'") {
			quote = character;
		} else if (character === ">") {
			return index;
		}
	}
	return -1;
}

function getRawXmlEvents(value) {
	const events = [];
	let index = 0;
	while (index < value.length) {
		const start = value.indexOf("<", index);
		if (start === -1) {
			break;
		}
		if (value.startsWith("<!--", start)) {
			const end = value.indexOf("-->", start + 4);
			index = end === -1 ? value.length : end + 3;
			continue;
		}
		if (value.startsWith("<![CDATA[", start)) {
			const end = value.indexOf("]]>", start + 9);
			index = end === -1 ? value.length : end + 3;
			continue;
		}

		const end = findXmlTagEnd(value, start);
		if (end === -1) {
			break;
		}
		let rawTag = value.slice(start + 1, end).trim();
		index = end + 1;
		if (!rawTag || rawTag[0] === "!" || rawTag[0] === "?") {
			continue;
		}

		let position = "start";
		if (rawTag[0] === "/") {
			position = "end";
			rawTag = rawTag.slice(1).trim();
		} else if (rawTag.endsWith("/")) {
			position = "selfclosing";
			rawTag = rawTag.slice(0, -1).trim();
		}
		const tag = rawTag.split(/\s/, 1)[0];
		if (/^[A-Za-z_][A-Za-z0-9_.:-]*$/.test(tag)) {
			events.push({ tag, position, start, end: end + 1 });
		}
	}
	return events;
}

function getXmlEvents(part) {
	if (part.type === "tag") {
		return [{ tag: part.tag, position: part.position }];
	}
	if (part.type === "content" && part.position === "outsidetag") {
		return getRawXmlEvents(part.value || "");
	}
	return [];
}

function getXmlEventValue(part, event) {
	if (part.type === "content" && part.position === "outsidetag") {
		return (part.value || "").slice(event.start, event.end);
	}
	return part.value || "";
}

function updateXmlStack(stack, event) {
	if (event.position === "start") {
		stack.push(event.tag);
		return;
	}
	if (event.position === "end" && stack[stack.length - 1] === event.tag) {
		stack.pop();
	}
}

function isParagraphSemanticXmlEvent(event, xmlStack) {
	if (event.position === "end") {
		return false;
	}
	const parent = xmlStack[xmlStack.length - 1];
	if (event.tag === "w:tab") {
		return parent === "w:r";
	}
	if (event.tag === "w:pageBreakBefore") {
		return parent === "w:pPr";
	}
	return PARAGRAPH_SEMANTIC_XML_TAGS.has(event.tag);
}

function updateParagraphFrame(frame, part, xmlStack) {
	if (isControl(part)) {
		frame.hasControl = true;
	} else if (
		part.type === "placeholder" ||
		(part.type === "content" &&
			part.position === "insidetag" &&
			(part.value || "").trim() !== "")
	) {
		frame.hasSemanticContent = true;
	}
	for (const event of getXmlEvents(part)) {
		if (isParagraphSemanticXmlEvent(event, xmlStack)) {
			frame.hasSemanticContent = true;
		}
		updateXmlStack(xmlStack, event);
	}
}

/*
 * Office places a tag typed on its own line inside a complete paragraph/run/text
 * scaffold. Mark that scaffold so it can be omitted together with the control
 * tag, without expanding inline switches to paragraph boundaries.
 */
function markControlOnlyParagraphs(parsed) {
	const frames = [];
	const xmlStack = [];

	for (let i = 0; i < parsed.length; i++) {
		const part = parsed[i];
		if (isParagraphStart(part)) {
			frames.push({
				tag: part.tag,
				start: i,
				hasControl: false,
				hasSemanticContent: false,
			});
		}
		const frame = frames[frames.length - 1];
		if (frame) {
			updateParagraphFrame(frame, part, xmlStack);
		} else {
			for (const event of getXmlEvents(part)) {
				updateXmlStack(xmlStack, event);
			}
		}
		if (!isParagraphEnd(part) || !frame || frame.tag !== part.tag) {
			continue;
		}

		frames.pop();
		if (frame.hasControl && !frame.hasSemanticContent) {
			for (let j = frame.start; j <= i; j++) {
				parsed[j].switchControlScaffold = true;
			}
		}
		const parent = frames[frames.length - 1];
		if (parent) {
			parent.hasControl ||= frame.hasControl;
			parent.hasSemanticContent ||= frame.hasSemanticContent;
		}
	}
}

function flattenObjectIdentifiers(tree, prefix = "", result = []) {
	if (!tree || typeof tree !== "object") {
		return result;
	}
	for (const key of Object.keys(tree)) {
		const path = prefix ? `${prefix}.${key}` : key;
		result.push(path);
		flattenObjectIdentifiers(tree[key], path, result);
	}
	return result;
}

function unique(values) {
	return [...new Set(values.filter(Boolean))];
}

function isPropertyForOwner(tag, ownerTag) {
	const parents = PROPERTY_XML_PARENTS.get(tag);
	if (parents) {
		return parents.has(ownerTag);
	}
	return (
		tag.startsWith("m:") && tag.endsWith("Pr") && ownerTag.startsWith("m:")
	);
}

function extractPropertyXml(value, ownerTag) {
	let result = "";
	let captureStart = -1;
	let captureDepth = 0;

	for (const event of getRawXmlEvents(value)) {
		if (captureStart !== -1) {
			if (event.position === "start") {
				captureDepth++;
			} else if (event.position === "end") {
				captureDepth--;
				if (captureDepth === 0) {
					result += value.slice(captureStart, event.end);
					captureStart = -1;
				}
			}
			continue;
		}
		if (!isPropertyForOwner(event.tag, ownerTag)) {
			continue;
		}
		if (event.position === "selfclosing") {
			result += value.slice(event.start, event.end);
		} else if (event.position === "start") {
			captureStart = event.start;
			captureDepth = 1;
		}
	}
	return result;
}

function commonPhysicalContext(left, right) {
	let length = 0;
	while (
		length < left.length &&
		length < right.length &&
		left[length].tag === right[length].tag &&
		left[length].id === right[length].id &&
		left[length].source === right[length].source
	) {
		length++;
	}
	return left.slice(0, length);
}

/*
 * Equal tag stacks are not enough: two w:r/a:r elements can have different
 * formatting. Close the selected branch's exact physical context, then reopen
 * the context containing {/!switch} with that context's paragraph/run
 * properties so content after the switch stays in its original container.
 */
function makeDirectContextBridge(currentContext, exitContext, anchorPart) {
	const common = commonPhysicalContext(currentContext, exitContext);
	const currentSuffix = currentContext.slice(common.length);
	const exitSuffix = exitContext.slice(common.length);
	if (
		[...currentSuffix, ...exitSuffix].some(
			(frame) => !RELOCATABLE_XML_CONTAINERS.has(frame.tag)
		)
	) {
		return [];
	}

	let value = "";
	for (let index = currentSuffix.length - 1; index >= 0; index--) {
		value += `</${currentSuffix[index].tag}>`;
	}
	for (let index = 0; index < exitSuffix.length; index++) {
		const frame = exitSuffix[index];
		value += frame.openXml;
		const next = exitSuffix[index + 1];
		if (next && frame.source === next.source) {
			value += extractPropertyXml(
				frame.source.xml.slice(frame.openEnd, next.openStart),
				frame.tag
			);
		}
	}
	if (!value) {
		return [];
	}
	return [
		{
			type: "content",
			position: "outsidetag",
			value,
			lIndex: anchorPart.lIndex,
			offset: anchorPart.offset,
		},
	];
}

function updateReverseContainerStack(events, stack) {
	for (let index = events.length - 1; index >= 0; index--) {
		const event = events[index];
		if (
			event.position === "selfclosing" ||
			RELOCATABLE_XML_CONTAINERS.has(event.tag)
		) {
			continue;
		}
		if (event.position === "end") {
			stack.push(event.tag);
		} else if (
			event.position === "start" &&
			stack[stack.length - 1] === event.tag
		) {
			stack.pop();
		}
	}
}

function findTrailingTransitionIndex(parts) {
	const containerStack = [];
	for (let index = parts.length - 1; index >= 0; index--) {
		const part = parts[index];
		updateReverseContainerStack(getXmlEvents(part), containerStack);
		if (
			containerStack.length === 0 &&
			(part.type === "placeholder" ||
				(part.type === "content" && part.position === "insidetag"))
		) {
			return index + 1;
		}
	}
	return 0;
}

function findTargetPrefix(parts) {
	const stack = [];
	for (let index = 0; index < parts.length; index++) {
		for (const event of getXmlEvents(parts[index])) {
			if (
				event.position === "selfclosing" ||
				!RELOCATABLE_XML_CONTAINERS.has(event.tag)
			) {
				continue;
			}
			if (event.position === "start") {
				stack.push({ tag: event.tag, index, event });
			} else if (
				event.position === "end" &&
				stack[stack.length - 1] &&
				stack[stack.length - 1].tag === event.tag
			) {
				stack.pop();
			}
		}
	}
	if (stack.length === 0) {
		return {
			priorBridge: parts.slice(),
			targetPrefix: [],
			openCount: 0,
		};
	}

	const first = stack[0];
	const priorBridge = parts.slice(0, first.index);
	const targetPrefix = parts.slice(first.index + 1);
	const boundaryPart = parts[first.index];
	if (
		boundaryPart.type === "content" &&
		boundaryPart.position === "outsidetag"
	) {
		const before = boundaryPart.value.slice(0, first.event.start);
		const after = boundaryPart.value.slice(first.event.start);
		if (before) {
			priorBridge.push({ ...boundaryPart, value: before });
		}
		if (after) {
			targetPrefix.unshift({ ...boundaryPart, value: after });
		}
	} else {
		targetPrefix.unshift(boundaryPart);
	}
	return {
		priorBridge,
		targetPrefix,
		openCount: stack.length,
	};
}

/*
 * A previous branch can end by closing containers that were already open at
 * the switch, then pass through complete paragraphs/runs before the next case.
 * The next branch needs those initial closing tags, but none of the balanced
 * detours owned by the previous branch.
 */
function clonePriorBridge(parts, initialXmlStack) {
	const result = [];
	const stack = initialXmlStack.map((tag) => ({ tag, initial: true }));

	function shouldKeep(event) {
		if (event.position === "start") {
			stack.push({ tag: event.tag, initial: false });
			return false;
		}
		if (event.position !== "end") {
			return false;
		}
		const current = stack[stack.length - 1];
		if (!current || current.tag !== event.tag) {
			return false;
		}
		stack.pop();
		return current.initial;
	}

	for (const part of parts) {
		if (part.type === "content" && part.position === "outsidetag") {
			let value = "";
			for (const event of getRawXmlEvents(part.value || "")) {
				if (shouldKeep(event)) {
					value += part.value.slice(event.start, event.end);
				}
			}
			if (value) {
				result.push({ ...part, value });
			}
			continue;
		}

		if (getXmlEvents(part).some(shouldKeep)) {
			result.push({ ...part });
		}
	}
	return result;
}

function takeTrailingTransition(parts, entryXmlStack) {
	const index = findTrailingTransitionIndex(parts);
	const transition = parts.slice(index);
	if (!entryXmlStack) {
		return { leadingContent: [], openCount: 0 };
	}

	const { priorBridge, targetPrefix, openCount } =
		findTargetPrefix(transition);
	parts.splice(index, transition.length, ...priorBridge);
	return {
		leadingContent: [
			...clonePriorBridge(priorBridge, entryXmlStack),
			...targetPrefix,
		],
		openCount,
	};
}

function findResolvedPart(scopeManager, part) {
	let { resolved } = scopeManager;
	for (
		let index = scopeManager.resolveOffset;
		index < scopeManager.scopePath.length;
		index++
	) {
		const lIndex = scopeManager.scopeLindex[index];
		const entry = resolved.find((item) => item.lIndex === lIndex);
		if (!entry) {
			return null;
		}
		resolved = entry.value[scopeManager.scopePathItem[index]];
	}
	return resolved.find((item) => item.lIndex === part.lIndex) || null;
}

function findMatchingSwitchEndIndex(parsed, startIndex) {
	let depth = 0;
	for (let index = startIndex + 1; index < parsed.length; index++) {
		if (isControl(parsed[index], "switch")) {
			depth++;
		} else if (isControl(parsed[index], "end")) {
			if (depth === 0) {
				return index;
			}
			depth--;
		}
	}
	return -1;
}

class ConditionalSwitchModule {
	constructor() {
		this.name = "ConditionalSwitchModule";
		this.priority = 100;
		this.requiredAPIVersion = "3.47.2";
		this.supportedFileTypes = ["docx", "pptx"];
		this.xmlContexts = new WeakMap();
	}

	clone() {
		return new ConditionalSwitchModule();
	}

	optionsTransformer(options, docxtemplater) {
		this.docxtemplater = docxtemplater;
		return options;
	}

	matchers() {
		return [
			[
				exactTag("!switch"),
				MODULE_NAME,
				{
					switchKind: "switch",
					value: ".",
					dataBound: false,
					priority: 100,
				},
			],
			[
				caseTag,
				MODULE_NAME,
				([, match]) => ({
					switchKind: "case",
					value: match.expression || ".",
					emptyCase: match.emptyCase,
					dataBound: false,
					priority: 100,
				}),
			],
			[
				exactTag("!default"),
				MODULE_NAME,
				{
					switchKind: "default",
					value: ".",
					dataBound: false,
					priority: 100,
				},
			],
			[
				exactTag("/!switch"),
				MODULE_NAME,
				{
					switchKind: "end",
					value: ".",
					dataBound: false,
					priority: 100,
				},
			],
		];
	}

	compileCondition(expression, part, options, errors) {
		try {
			const parser = this.docxtemplater.parser(expression, { tag: part });
			options.cachedParsers[part.lIndex] = parser;
			const identifiers =
				typeof parser.getIdentifiers === "function"
					? parser.getIdentifiers()
					: [];
			const objectIdentifiers =
				typeof parser.getObjectIdentifiers === "function"
					? flattenObjectIdentifiers(parser.getObjectIdentifiers())
					: [];
			return unique([...identifiers, ...objectIdentifiers]);
		} catch (rootError) {
			errors.push(
				makeTemplateError(
					"conditional_switch_invalid_expression",
					`Invalid conditional switch case expression "${expression}".`,
					part,
					{ expression, rootError, xtag: expression }
				)
			);
			return [];
		}
	}

	makeDiscoveryPart(identifier, casePart, index) {
		return {
			type: "placeholder",
			value: identifier,
			raw: identifier,
			lIndex: `${casePart.lIndex}-switch-identifier-${index}`,
			offset: casePart.offset,
		};
	}

	postparse(parsed, options) {
		markControlOnlyParagraphs(parsed);
		this.ensureXmlContexts(parsed);
		const errors = [];
		const result = this.groupLevel(parsed, options, errors, false);
		return { postparsed: result.parts, errors };
	}

	ensureXmlContexts(parsed) {
		const controls = parsed.filter((part) => isControl(part));
		if (
			controls.length === 0 ||
			controls.some((part) => this.xmlContexts.has(part))
		) {
			return;
		}

		const stack = [];
		const source = { chunks: [], length: 0, xml: "" };
		let nextId = 0;
		for (const part of parsed) {
			if (isControl(part)) {
				this.xmlContexts.set(part, stack.slice());
			}
			if (part.switchControlScaffold) {
				continue;
			}
			for (const event of getXmlEvents(part)) {
				const openXml = getXmlEventValue(part, event);
				const openStart = source.length;
				source.chunks.push(openXml);
				source.length += openXml.length;
				const openEnd = source.length;
				if (event.position === "start") {
					stack.push({
						tag: event.tag,
						id: nextId++,
						openXml,
						openStart,
						openEnd,
						source,
					});
				} else if (event.position === "end") {
					const current = stack[stack.length - 1];
					if (current && current.tag === event.tag) {
						stack.pop();
					}
				}
			}
		}
		source.xml = source.chunks.join("");
		delete source.chunks;
		delete source.length;
	}

	hasSameProtectedContext(left, right) {
		const leftProtected = left.filter(
			(item) => !RELOCATABLE_XML_CONTAINERS.has(item.tag)
		);
		const rightProtected = right.filter(
			(item) => !RELOCATABLE_XML_CONTAINERS.has(item.tag)
		);
		if (leftProtected.length !== rightProtected.length) {
			return false;
		}
		return leftProtected.every(
			(item, index) =>
				item.tag === rightProtected[index].tag &&
				item.id === rightProtected[index].id
		);
	}

	branchMatchesContext(branch, entryContext, exitContext) {
		const stack = entryContext.map((item) => item.tag);
		for (const part of branch.rawContent) {
			if (part.switchControlScaffold) {
				continue;
			}
			for (const event of getXmlEvents(part)) {
				if (event.position === "start") {
					stack.push(event.tag);
				} else if (event.position === "end") {
					if (stack[stack.length - 1] !== event.tag) {
						return false;
					}
					stack.pop();
				}
			}
		}
		const expected = exitContext.map((item) => item.tag);
		return (
			stack.length === expected.length &&
			stack.every((tag, index) => tag === expected[index])
		);
	}

	validateSwitchStructure(switchPart, endPart, branches, errors) {
		const entryContext = this.xmlContexts.get(switchPart);
		const exitContext = this.xmlContexts.get(endPart);
		const controls = [
			switchPart,
			...branches.map((branch) => branch.controlPart),
			endPart,
		];
		if (
			!entryContext ||
			!exitContext ||
			controls.some((part) => !this.xmlContexts.has(part))
		) {
			return;
		}

		const crossesContainer = controls.some((part) => {
			const context = this.xmlContexts.get(part);
			return !this.hasSameProtectedContext(entryContext, context);
		});
		const invalidBranch = branches.find(
			(branch) =>
				!this.branchMatchesContext(branch, entryContext, exitContext)
		);
		if (!crossesContainer && !invalidBranch) {
			return;
		}

		errors.push(
			makeTemplateError(
				"conditional_switch_invalid_structure",
				"A {!switch} block crosses incompatible document XML containers. Keep the switch and all branch controls within the same table cell, hyperlink, and content container.",
				switchPart,
				{
					branch:
						invalidBranch && invalidBranch.kind === "case"
							? invalidBranch.expression
							: invalidBranch && invalidBranch.kind,
				}
			)
		);
	}

	groupLevel(parsed, options, errors, insideSwitch) {
		const output = [];

		for (let index = 0; index < parsed.length; index++) {
			const part = parsed[index];
			if (!isControl(part)) {
				if (!part.switchControlScaffold) {
					output.push(part);
				}
				continue;
			}

			if (part.switchKind === "switch") {
				if (part.branches) {
					output.push(part);
					continue;
				}
				const grouped = this.groupSwitch(
					parsed,
					index,
					options,
					errors
				);
				if (grouped.part) {
					output.push(grouped.part);
				}
				index = grouped.endIndex;
				continue;
			}

			if (part.switchKind === "end") {
				errors.push(
					makeTemplateError(
						"conditional_switch_unmatched_closing_tag",
						"Closing {/!switch} has no matching {!switch}.",
						part
					)
				);
			} else {
				errors.push(
					makeTemplateError(
						"conditional_switch_branch_outside_switch",
						`{${part.raw}} must appear inside a {!switch} block.`,
						part
					)
				);
			}
		}

		return { parts: output, insideSwitch };
	}

	// eslint-disable-next-line complexity
	groupSwitch(parsed, startIndex, options, errors) {
		const switchPart = parsed[startIndex];
		const matchingEndIndex = findMatchingSwitchEndIndex(parsed, startIndex);
		const entryContext = this.xmlContexts.get(switchPart);
		const exitContext =
			matchingEndIndex === -1
				? null
				: this.xmlContexts.get(parsed[matchingEndIndex]);
		const entryXmlStack =
			entryContext && entryContext.map((item) => item.tag);
		const branches = [];
		let currentBranch = null;
		let defaultSeen = false;
		let depth = 0;
		let endIndex = parsed.length - 1;
		let foundEnd = false;
		let pendingContent = [];

		for (let index = startIndex + 1; index < parsed.length; index++) {
			const part = parsed[index];
			if (isControl(part, "switch")) {
				depth++;
				if (currentBranch) {
					currentBranch.rawContent.push(part);
				} else if (!part.switchControlScaffold) {
					pendingContent.push(part);
				}
				continue;
			}
			if (isControl(part, "end")) {
				if (depth > 0) {
					depth--;
					if (currentBranch) {
						currentBranch.rawContent.push(part);
					} else if (!part.switchControlScaffold) {
						pendingContent.push(part);
					}
					continue;
				}
				endIndex = index;
				foundEnd = true;
				break;
			}

			if (depth > 0) {
				if (currentBranch) {
					if (isControl(part) || !part.switchControlScaffold) {
						currentBranch.rawContent.push(part);
					}
				} else if (isControl(part) || !part.switchControlScaffold) {
					pendingContent.push(part);
				}
				continue;
			}

			if (depth === 0 && isControl(part, "case")) {
				const controlContext = this.xmlContexts.get(part);
				if (part.emptyCase) {
					part.value = "";
				}
				if (!part.value) {
					errors.push(
						makeTemplateError(
							"conditional_switch_empty_case",
							"A {!case} tag must contain a non-empty expression.",
							part,
							{ expression: part.value }
						)
					);
				}
				if (defaultSeen) {
					errors.push(
						makeTemplateError(
							"conditional_switch_case_after_default",
							`Case expression "${part.value}" appears after {!default}.`,
							part,
							{ expression: part.value }
						)
					);
				}
				const transition = currentBranch
					? takeTrailingTransition(
							currentBranch.rawContent,
							entryXmlStack
						)
					: null;
				if (currentBranch && controlContext) {
					currentBranch.endContext = controlContext.slice(
						0,
						Math.max(
							0,
							controlContext.length - transition.openCount
						)
					);
				}
				const leadingContent = transition
					? transition.leadingContent
					: pendingContent;
				pendingContent = [];
				currentBranch = {
					kind: "case",
					expression: part.value,
					controlPart: part,
					controlContext,
					rawContent: leadingContent,
				};
				branches.push(currentBranch);
				continue;
			}

			if (depth === 0 && isControl(part, "default")) {
				const controlContext = this.xmlContexts.get(part);
				if (defaultSeen) {
					errors.push(
						makeTemplateError(
							"conditional_switch_multiple_defaults",
							"A switch may contain at most one {!default} branch.",
							part
						)
					);
				}
				defaultSeen = true;
				const transition = currentBranch
					? takeTrailingTransition(
							currentBranch.rawContent,
							entryXmlStack
						)
					: null;
				if (currentBranch && controlContext) {
					currentBranch.endContext = controlContext.slice(
						0,
						Math.max(
							0,
							controlContext.length - transition.openCount
						)
					);
				}
				const leadingContent = transition
					? transition.leadingContent
					: pendingContent;
				pendingContent = [];
				currentBranch = {
					kind: "default",
					expression: null,
					controlPart: part,
					controlContext,
					rawContent: leadingContent,
				};
				branches.push(currentBranch);
				continue;
			}

			if (currentBranch && !part.switchControlScaffold) {
				currentBranch.rawContent.push(part);
			} else if (!currentBranch && !part.switchControlScaffold) {
				pendingContent.push(part);
			}
		}

		if (!foundEnd) {
			errors.push(
				makeTemplateError(
					"conditional_switch_missing_closing_tag",
					"A {!switch} block is missing its closing {/!switch} tag.",
					switchPart
				)
			);
		}

		const caseCount = branches.filter(
			(branch) => branch.kind === "case"
		).length;
		if (caseCount === 0) {
			errors.push(
				makeTemplateError(
					"conditional_switch_no_cases",
					"A {!switch} block must contain at least one {!case expression}.",
					switchPart
				)
			);
		}

		if (foundEnd) {
			if (currentBranch) {
				currentBranch.endContext = exitContext;
			}
			for (const branch of branches.slice(0, -1)) {
				if (branch.endContext && exitContext) {
					branch.rawContent.push(
						...makeDirectContextBridge(
							branch.endContext,
							exitContext,
							parsed[endIndex]
						)
					);
				}
			}
			this.validateSwitchStructure(
				switchPart,
				parsed[endIndex],
				branches,
				errors
			);
		}

		const emptyContent =
			foundEnd && entryContext && exitContext
				? options.postparse(
						makeDirectContextBridge(
							entryContext,
							exitContext,
							parsed[endIndex]
						),
						{ basePart: switchPart }
					)
				: [];
		for (const branch of branches) {
			branch.content = options.postparse(branch.rawContent, {
				basePart: switchPart,
			});
			delete branch.rawContent;
			const identifiers =
				branch.kind === "case"
					? this.compileCondition(
							branch.expression,
							branch.controlPart,
							options,
							errors
						)
					: [];
			branch.identifiers = identifiers;
			branch.discoveryPart = {
				...branch.controlPart,
				dataBound: false,
				subparsed: [
					...identifiers.map((identifier, index) =>
						this.makeDiscoveryPart(
							identifier,
							branch.controlPart,
							index
						)
					),
					...branch.content,
				],
			};
		}

		const groupedPart = {
			...switchPart,
			dataBound: false,
			branches,
			emptyContent,
			subparsed: branches.map((branch) => branch.discoveryPart),
			endLindex: foundEnd ? parsed[endIndex].lIndex : switchPart.lIndex,
		};
		return { part: groupedPart, endIndex };
	}

	resolve(part, options) {
		if (!isControl(part, "switch") || !part.branches) {
			return null;
		}
		return this.resolveSwitch(part, options);
	}

	async resolveSwitch(part, options) {
		const conditionResults = [];
		let selected = null;

		for (const branch of part.branches) {
			if (branch.kind === "default") {
				selected = branch;
				break;
			}
			const value = await options.scopeManager.getValueAsync(
				branch.expression,
				{ part: branch.controlPart }
			);
			conditionResults.push({
				tag: branch.expression,
				lIndex: branch.controlPart.lIndex,
				value,
			});
			if (value) {
				selected = branch;
				break;
			}
		}

		let resolved = [];
		if (selected) {
			const result = await options.resolve({
				...options,
				compiled: selected.content,
				tags: {},
				scopeManager: options.scopeManager,
			});
			if (result.errors.length) {
				throw result.errors;
			}
			resolved = result.resolved;
		}
		return [[...conditionResults, ...resolved]];
	}

	getSelectedBranch(part, scopeManager, resolvedEntries) {
		for (const branch of part.branches) {
			if (branch.kind === "default") {
				return branch;
			}
			const resolved = resolvedEntries
				? resolvedEntries.find(
						(item) => item.lIndex === branch.controlPart.lIndex
					)
				: null;
			const value = resolvedEntries
				? resolved && resolved.value
				: scopeManager.getValue(branch.expression, {
						part: branch.controlPart,
					});
			if (value) {
				return branch;
			}
		}
		return null;
	}

	render(part, options) {
		if (!isControl(part, "switch") || !part.branches) {
			return null;
		}
		const finishedResolving =
			options.scopeManager.root.finishedResolving === true;
		const switchResolution = finishedResolving
			? findResolvedPart(options.scopeManager, part)
			: null;
		const resolvedEntries =
			switchResolution && switchResolution.value
				? switchResolution.value[0]
				: null;
		const selected = this.getSelectedBranch(
			part,
			options.scopeManager,
			resolvedEntries
		);

		let branchScope = options.scopeManager;
		if (resolvedEntries) {
			branchScope = Object.create(options.scopeManager);
			branchScope.resolved = resolvedEntries;
			branchScope.resolveOffset = branchScope.scopePath.length;
		}
		const subRendered = options.render({
			...options,
			compiled: selected ? selected.content : part.emptyContent,
			tags: {},
			scopeManager: branchScope,
		});
		return {
			value: options.joinUncorrupt(subRendered.parts, {
				...options,
				basePart: part,
			}),
			errors: subRendered.errors,
		};
	}
}

module.exports = ConditionalSwitchModule;
module.exports.MODULE_NAME = MODULE_NAME;
