"use strict";

const MODULE_NAME = "private-conditional-switch";

function exactTag(expected) {
	return (tag) => (tag.trim() === expected ? "." : false);
}

function caseTag(tag) {
	const match = tag.match(/^!case(?:\s+([\s\S]*))?$/);
	if (!match) {
		return false;
	}
	return (match[1] || "").trim() || "__EMPTY_CONDITIONAL_SWITCH_CASE__";
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

/*
 * Word places a tag typed on its own line inside a complete paragraph/run/text
 * scaffold. Mark that scaffold so it can be omitted together with the control
 * tag, without expanding inline switches to paragraph boundaries.
 */
// eslint-disable-next-line complexity
function markControlOnlyParagraphs(parsed) {
	let start = -1;
	let depth = 0;

	for (let i = 0; i < parsed.length; i++) {
		const part = parsed[i];
		if (isParagraphStart(part)) {
			if (depth === 0) {
				start = i;
			}
			depth++;
		}
		if (!isParagraphEnd(part)) {
			continue;
		}
		depth--;
		if (depth !== 0 || start === -1) {
			continue;
		}

		const paragraph = parsed.slice(start, i + 1);
		let controls = 0;
		let hasVisibleContent = false;
		for (const paragraphPart of paragraph) {
			if (isControl(paragraphPart)) {
				controls++;
			} else if (paragraphPart.type === "placeholder") {
				hasVisibleContent = true;
			} else if (
				paragraphPart.type === "content" &&
				paragraphPart.value &&
				paragraphPart.value.trim() !== ""
			) {
				hasVisibleContent = true;
			}
		}
		if (controls > 0 && !hasVisibleContent) {
			for (let j = start; j <= i; j++) {
				parsed[j].switchControlScaffold = true;
			}
		}
		start = -1;
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

class ConditionalSwitchModule {
	constructor() {
		this.name = "ConditionalSwitchModule";
		this.requiredAPIVersion = "3.47.2";
		this.caseExpressions = new Map();
		this.caseExpressionsByOffset = new Map();
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
				([, expression]) => ({
					switchKind: "case",
					value: expression,
					emptyCase:
						expression === "__EMPTY_CONDITIONAL_SWITCH_CASE__",
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

	getConditionIdentifiers(expression, part) {
		try {
			const parser = this.docxtemplater.parser(expression, { tag: part });
			const identifiers =
				typeof parser.getIdentifiers === "function"
					? parser.getIdentifiers()
					: [];
			const objectIdentifiers =
				typeof parser.getObjectIdentifiers === "function"
					? flattenObjectIdentifiers(parser.getObjectIdentifiers())
					: [];
			return unique([...identifiers, ...objectIdentifiers]);
		} catch {
			/*
			 * The core Render module reports the parser compilation error. Keeping
			 * discovery best-effort here avoids producing a duplicate error.
			 */
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
		const errors = [];
		const result = this.groupLevel(parsed, options, errors, false);
		return { postparsed: result.parts, errors };
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
		const branches = [];
		let currentBranch = null;
		let defaultSeen = false;
		let depth = 0;
		let endIndex = parsed.length - 1;
		let foundEnd = false;

		for (let index = startIndex + 1; index < parsed.length; index++) {
			const part = parsed[index];
			if (isControl(part, "switch")) {
				depth++;
				if (currentBranch) {
					currentBranch.rawContent.push(part);
				}
				continue;
			}
			if (isControl(part, "end")) {
				if (depth > 0) {
					depth--;
					if (currentBranch) {
						currentBranch.rawContent.push(part);
					}
					continue;
				}
				endIndex = index;
				foundEnd = true;
				break;
			}

			if (depth === 0 && isControl(part, "case")) {
				if (part.emptyCase) {
					part.value = "";
				}
				this.caseExpressions.set(part.lIndex, part.value);
				this.caseExpressionsByOffset.set(part.offset, part.value);
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
				currentBranch = {
					kind: "case",
					expression: part.value,
					controlPart: part,
					rawContent: [],
				};
				branches.push(currentBranch);
				continue;
			}

			if (depth === 0 && isControl(part, "default")) {
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
				currentBranch = {
					kind: "default",
					expression: null,
					controlPart: part,
					rawContent: [],
				};
				branches.push(currentBranch);
				continue;
			}

			if (currentBranch && !part.switchControlScaffold) {
				currentBranch.rawContent.push(part);
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

		for (const branch of branches) {
			branch.content = options.postparse(branch.rawContent, {
				basePart: switchPart,
			});
			delete branch.rawContent;
			const identifiers =
				branch.kind === "case"
					? this.getConditionIdentifiers(
							branch.expression,
							branch.controlPart
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
			subparsed: branches.map((branch) => branch.discoveryPart),
			endLindex: foundEnd ? parsed[endIndex].lIndex : switchPart.lIndex,
		};
		return { part: groupedPart, endIndex };
	}

	errorsTransformer(errors) {
		for (const error of errors) {
			const lIndex = error && error.properties && error.properties.lIndex;
			const offset = error && error.properties && error.properties.offset;
			const expression =
				this.caseExpressions.get(lIndex) ||
				this.caseExpressionsByOffset.get(offset);
			if (
				error &&
				error.properties &&
				error.properties.id === "scopeparser_compilation_failed" &&
				expression != null
			) {
				error.properties.id = "conditional_switch_invalid_expression";
				error.properties.expression = expression;
				error.properties.xtag = expression;
				error.properties.explanation = `Invalid conditional switch case expression "${expression}": ${error.properties.explanation}`;
			}
		}
		return errors;
	}

	createBranchScope(scopeManager, part) {
		const scope = scopeManager.scopeList[scopeManager.scopeList.length - 1];
		return scopeManager.createSubScopeManager(
			scope,
			part.value,
			0,
			part,
			1
		);
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

	getSelectedBranch(part, scopeManager, branchScope) {
		for (const branch of part.branches) {
			if (branch.kind === "default") {
				return branch;
			}
			const manager = scopeManager.root.finishedResolving
				? branchScope
				: scopeManager;
			const value = manager.getValue(branch.expression, {
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
		const branchScope = this.createBranchScope(options.scopeManager, part);
		const selected = this.getSelectedBranch(
			part,
			options.scopeManager,
			branchScope
		);
		if (!selected) {
			return { value: "" };
		}

		const subRendered = options.render({
			...options,
			compiled: selected.content,
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
