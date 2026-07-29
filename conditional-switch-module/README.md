# Conditional switch module

Private Docxtemplater module implementing ordered `switch` / `case` /
`default` blocks without modifying Docxtemplater core.

## Supported versions

- Standalone Node integration: `docxtemplater@3.69.3`
- Docxtemplater Docker image: `3.46.5` (bundles Docxtemplater `3.69.0`)
- Node.js: active LTS releases from 18 through 24
- File formats: DOCX and PPTX

These versions are deliberately pinned. Upgrade the module and rerun its test
suite before changing either Docxtemplater or the Docker image.

## Installation

Copy this directory into the application and install it as a private package:

```sh
npm install ./conditional-switch-module
```

Configure Docxtemplater with the Angular expression parser already used by the
application:

```js
const Docxtemplater = require("docxtemplater");
const expressions = require("docxtemplater/expressions.js");
const ConditionalSwitchModule = require("@private/docxtemplater-conditional-switch");

const doc = new Docxtemplater(zip, {
    parser: expressions,
    modules: [new ConditionalSwitchModule()],
});
```

Do not create a second expression parser for this module. Case conditions are
compiled and evaluated through the parser configured on the Docxtemplater
instance, so its filters, missing-value behavior, and loop-local scopes apply.

## Docker image integration

Pin the enterprise image to `3.46.5`, copy `index.js` into the image next to
`configuration.js`, and merge this hook into the existing configuration
object:

```js
const ConditionalSwitchModule = require("./conditional-switch-module/index.js");

module.exports = {
    // Keep the image's other configuration hooks here.
    configureDocxtemplater(doc) {
        doc.attachModule(new ConditionalSwitchModule());
        return doc;
    },
};
```

An equivalent ready-to-merge example is provided in
`docker-configuration.js`. The hook runs before compilation, so `/generate`,
`/retrieve-tags`, and `/retrieve-structured-tags` all see the module. The
current Docker image still supports `configureDocxtemplater`; its documentation
recommends `transformModules` for new integrations, but this hook is retained
here to meet the deployment contract.

If filters are registered through the Docker image's
`transformExpressionParserOptions`, case expressions use those same filters.

## Syntax and behavior

```text
The product has a {!switch}{!case price < 10}low{!case price < 100}medium{!default}high{/!switch} price.
```

```text
{!switch}
{!case customer.active && amount >= 1000}
Approved
{!case status == "PENDING"}
Pending
{!default}
Other
{/!switch}
```

Cases are evaluated from left to right. The first truthy case is rendered and
later cases are not evaluated. `default` is optional; without a match or
default, the block renders nothing. Switches may be nested and may appear
inline, across paragraphs, in table cells, headers, footers, and loop-local
scopes.

Control-only paragraphs are removed. Inline control tags are removed without
expanding to paragraph boundaries, preserving surrounding text and physical
run/paragraph formatting in Word and PowerPoint.

All controls in one switch must remain in the same table cell, hyperlink, and
other structural content container. Layouts that cannot render every branch as
balanced document XML fail compilation with
`conditional_switch_invalid_structure`.

## Validation errors

Compilation reports module-specific error IDs for:

- `conditional_switch_missing_closing_tag`
- `conditional_switch_unmatched_closing_tag`
- `conditional_switch_branch_outside_switch`
- `conditional_switch_no_cases`
- `conditional_switch_empty_case`
- `conditional_switch_multiple_defaults`
- `conditional_switch_case_after_default`
- `conditional_switch_invalid_expression`
- `conditional_switch_invalid_structure`

Errors include the raw tag or expression, offset/lIndex where available, and
the document part added by Docxtemplater (for example
`word/document.xml` or `word/header1.xml`).

## Tag discovery

Every branch remains in the postparsed inspection tree. In addition, parser
identifiers from case expressions are inserted as inspection-only
placeholders. For:

```text
{!switch}{!case price < threshold}{product.name}{!default}{fallbackText}{/!switch}
```

tag inspection exposes `price`, `threshold`, `product.name`, and
`fallbackText`, while structured inspection also retains the original
`price < threshold` case expression.

## Tests

From the Docxtemplater checkout:

```sh
npm run test:switch
```

The suite covers synchronous and asynchronous rendering, ordered selection,
Angular expressions and filters, missing values, loops, nesting, inline and
multiline Word/PowerPoint layout, tables, headers, footers, split runs,
formatting, inspection, and all structural validation rules.
