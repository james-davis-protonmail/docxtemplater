"use strict";

const MODULE_NAME = "private-conditional-switch";

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
				paragraphPart.type === "tag" &&
				paragraphPart.position === "selfclosing" &&
				[
					"w:br",
					"w:bookmarkStart",
					"w:bookmarkEnd",
					"w:commentRangeStart",
					"w:commentRangeEnd",
				].includes(paragraphPart.tag)
			) {
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

function isVisibleBranchPart(part) {
	return (
		part.type === "placeholder" ||
		(part.type === "content" && part.position === "insidetag")
	);
}

function takeTrailingTransition(parts) {
	const lastPart = parts[parts.length - 1];
	if (!lastPart || lastPart.type !== "tag" || lastPart.position !== "start") {
		return [];
	}
	let index = parts.length - 1;
	while (index >= 0 && !isVisibleBranchPart(parts[index])) {
		index--;
	}
	return parts.splice(index + 1);
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

class ConditionalSwitchModule {
	constructor() {
		this.name = "ConditionalSwitchModule";
		this.priority = 100;
		this.requiredAPIVersion = "3.47.2";
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
				const leadingContent = currentBranch
					? takeTrailingTransition(currentBranch.rawContent)
					: pendingContent;
				pendingContent = [];
				currentBranch = {
					kind: "case",
					expression: part.value,
					controlPart: part,
					rawContent: leadingContent,
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
				const leadingContent = currentBranch
					? takeTrailingTransition(currentBranch.rawContent)
					: pendingContent;
				pendingContent = [];
				currentBranch = {
					kind: "default",
					expression: null,
					controlPart: part,
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
		if (!selected) {
			return { value: "" };
		}

		let branchScope = options.scopeManager;
		if (resolvedEntries) {
			branchScope = Object.create(options.scopeManager);
			branchScope.resolved = resolvedEntries;
			branchScope.resolveOffset = branchScope.scopePath.length;
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
