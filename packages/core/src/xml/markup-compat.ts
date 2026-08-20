/**
 * PptxGenJS: Markup Compatibility (ECMA-376 Part 3)
 *
 * Optional markup — anything outside the base ECMA-376 schemas — may only reach a package through
 * one of three mechanisms (MS-PPTX 2.2):
 *
 * 1. an `extLst` extension, keyed by a `p:ext@uri` (see `OOXML_EXT` in core-enums)
 * 2. `mc:AlternateContent`, when the optional markup *replaces* a standard element and therefore
 *    needs a fallback for consumers that do not understand it
 * 3. `@mc:Ignorable` on the part root, when optional markup appears as *attributes* on otherwise
 *    standard elements. No PresentationML feature needs this yet - every optional attribute we emit
 *    (`p14:dur`) sits inside an `mc:Choice` - so it is deliberately not implemented until one does.
 *    `test/ms-pptx-profile.test.ts` fails if optional markup ever appears outside a wrapper.
 *
 * This module owns 2, so that every call site produces the same, verifiable shape:
 * - `xmlns:mc` is declared on `mc:AlternateContent`
 * - every optional namespace is declared on `mc:Choice`, keeping it out of the fallback's scope
 * - the `Requires` prefixes are exactly the prefixes declared there, so a consumer can always
 *   evaluate the condition
 * - a fallback is always present, because an `mc:AlternateContent` without one silently drops
 *   content in consumers that reject the choice
 */

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

/** Namespace declarations for the optional markup inside an `mc:Choice`, keyed by prefix */
export type NamespaceMap = Record<string, string>

/**
 * Offer optional markup with a fallback (ECMA-376 Part 3, 10.2)
 *
 * @param {object} props - alternate-content parts
 * @param {NamespaceMap} props.namespaces - prefix -> namespace for the optional markup
 * @param {string} props.choice - the optional markup; must only use the declared prefixes
 * @param {string} props.fallback - standard markup for consumers that reject the choice; an empty
 *   string yields `<mc:Fallback/>`, which is legal and means "render nothing"
 * @param {string[]} props.requires - prefixes a consumer must understand; defaults to every
 *   declared prefix
 * @returns {string} XML string
 */
export function alternateContent (props: { namespaces: NamespaceMap, choice: string, fallback: string, requires?: string[] }): string {
	const prefixes = Object.keys(props.namespaces)
	if (prefixes.length === 0) {
		// Without an optional namespace there is nothing to be compatible about; emitting the
		// wrapper would only add a level of indirection around the choice.
		return props.choice
	}

	const requires = props.requires ?? prefixes
	const undeclared = requires.filter(prefix => !prefixes.includes(prefix))
	if (undeclared.length > 0) {
		// A `Requires` prefix that is not in scope makes the condition unevaluable, so a consumer
		// must treat the whole block as unsupported - the choice would never be used
		console.warn(`[pptxgenjs] mc:Choice requires undeclared namespace prefix(es) "${undeclared.join(', ')}" - declaring them alongside`)
	}

	const declarations = prefixes.map(prefix => `xmlns:${prefix}="${props.namespaces[prefix]}"`).join(' ')

	return (
		`<mc:AlternateContent xmlns:mc="${MC_NS}">` +
		`<mc:Choice ${declarations} Requires="${[...new Set([...requires, ...undeclared])].join(' ')}">${props.choice}</mc:Choice>` +
		(props.fallback ? `<mc:Fallback>${props.fallback}</mc:Fallback>` : '<mc:Fallback/>') +
		'</mc:AlternateContent>'
	)
}
