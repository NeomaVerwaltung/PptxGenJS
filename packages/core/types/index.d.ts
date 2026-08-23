// Type definitions for pptxgenjs 4.2.0
// Project: https://gitbrent.github.io/PptxGenJS/
// Definitions by: Brent Ely <https://github.com/gitbrent/>
//                 Michael Beaumont <https://github.com/michaelbeaumont>
//                 Nicholas Tietz-Sokolsky <https://github.com/ntietz>
//                 David Adams <https://github.com/iota-pi>
//                 Stephen Cronin <https://github.com/cronin4392>
// TypeScript Version: 3.x

export as namespace PptxGenJS

export default PptxGenJS

declare class PptxGenJS {
	/**
	 * PptxGenJS Library Version
	 * @type {string}
	 */
	readonly version: string

	// Exposed prop types
	readonly presLayout: PptxGenJS.PresLayout
	readonly AlignH: typeof PptxGenJS.AlignH
	readonly AlignV: typeof PptxGenJS.AlignV
	readonly ChartType: typeof PptxGenJS.ChartType
	readonly OutputType: typeof PptxGenJS.OutputType
	readonly SchemeColor: typeof PptxGenJS.SchemeColor
	readonly ShapeType: typeof PptxGenJS.ShapeType
	readonly PlaceholderType: typeof PptxGenJS.PLACEHOLDER_TYPES

	// Presentation Props

	/**
	 * Presentation layout name.
	 * Standard layouts:
	 * - 'LAYOUT_4x3'   (10" x 7.5")
	 * - 'LAYOUT_16x9'  (10" x 5.625")
	 * - 'LAYOUT_16x10' (10" x 6.25")
	 * - 'LAYOUT_WIDE'  (13.33" x 7.5")
	 *
	 * Custom layouts:
	 * - Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
	 *
	 * @type {string}
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	layout: string
	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 * @type {boolean}
	 */
	rtlMode: boolean

	/**
	 * Zip compression for exported files - applied by every export method
	 * - 'none': store uncompressed (fastest), 'fast': DEFLATE level 1, 'best': DEFLATE level 9
	 * @default 'none'
	 * @since v4.1.0
	 */
	compression: PptxGenJS.CompressionLevel

	/**
	 * Whether charts track data references rather than positions (MS-PPTX 2.2.12 `p15:chartTrackingRefBased`)
	 * - PowerPoint sets this on the presentations it creates, so it is on by default here too
	 * - affects no rendering; set `false` to leave the presentation properties part empty
	 * @default true
	 */
	chartTrackingRefBased: boolean

	/**
	 * Comment authors, with the identity metadata PowerPoint shows
	 * - authors named by `addComment()` that are not listed here are added automatically
	 */
	commentAuthors?: PptxGenJS.CommentAuthorProps[]

	/**
	 * Image quality preference PowerPoint applies to inserted pictures (MS-PPTX 2.2.7)
	 * - `220` is PowerPoint's own default; `0` means "do not compress"
	 */
	defaultImageDpi?: number

	/**
	 * Whether PowerPoint discards picture crop/edit data on save (MS-PPTX 2.2.7)
	 * @default false
	 */
	discardImageEditData: boolean

	/**
	 * Whether PowerPoint recommends opening the file read-only (MS-PPTX 2.2.16)
	 * @default false
	 */
	readonlyRecommended: boolean

	/**
	 * Slide-show options (`p:showPr`)
	 */
	slideShow?: PptxGenJS.SlideShowProps

	/** Drawing guides shown in the slide editing view */
	guides?: PptxGenJS.GuideProps[]

	/** Drawing guides shown in the notes view */
	notesGuides?: PptxGenJS.GuideProps[]

	// Presentation Metadata
	/**
	 * Author name
	 * @type {string}
	 */
	author: string
	/**
	 * Comapny name
	 * @type {string}
	 */
	company: string
	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PowerPoint will throw errors upon opening!)
	 */
	revision: string
	/**
	 * Presentation subject
	 * @type {string}
	 */
	subject: string
	/**
	 * Presentation theme (default fonts)
	 * @type {ThemeProps}
	 */
	theme: PptxGenJS.ThemeProps
	/**
	 * Presentation name
	 * @type {string}
	 */
	title: string

	// Methods

	/**
	 * Export the current Presentation to stream
	 * @param {WriteBaseProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file stream
	 */
	stream(props?: PptxGenJS.WriteBaseProps): Promise<string | ArrayBuffer | Blob | Uint8Array>
	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {WriteProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file content in selected type
	 */
	write(props?: PptxGenJS.WriteProps): Promise<string | ArrayBuffer | Blob | Uint8Array>
	/**
	 * Export the current Presentation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
	 * @param {WriteFileProps} props output file properties
	 * @example pptx.writeFile({ fileName:'CustomerReport.pptx' }) // export presentation as "CustomerReport.pptx"
	 * @example pptx.writeFile({ fileName:'CustomerReport.pptx', compression:true }) // export presentation as "CustomerReport.pptx" compressed (can save up to 30%)
	 * @returns {Promise<string>} the presentation name
	 */
	writeFile(props?: PptxGenJS.WriteFileProps): Promise<string>
	/**
	 * Add a new Section to Presentation
	 * @param {SectionProps} props section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(props: PptxGenJS.SectionProps): void
	/**
	 * Add a new Slide to Presentation
	 * @param {AddSlideProps} props slide options
	 * @returns {Slide} the new Slide
	 */
	/**
	 * Embed a font in the presentation
	 * - opt-in: without any `addFont()` call the package has no font parts and is unchanged
	 * - PowerPoint stores embedded fonts as Embedded OpenType (EOT); convert TTF/OTF/WOFF first
	 * - you must hold a licence that permits embedding the font
	 * @param options - font props
	 * @example pptx.addFont({ fontFace: 'Custom Sans', data: eotBase64 })
	 */
	addFont(options: PptxGenJS.AddFontProps): void

	addSlide(props?: PptxGenJS.AddSlideProps): PptxGenJS.Slide
	/**
	 * Add a new Slide to Presentation
	 * @param {string} masterName master slide name
	 * @returns {Slide} the new Slide
	 * @deprecated use `addSlide(IAddSlideOptions)`
	 */
	addSlide(masterName?: string): PptxGenJS.Slide
	/**
	 * Create a custom Slide Layout in any size
	 * @param {PresLayout} layout an object with user-defined w/h
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout(layout: PptxGenJS.DefineLayoutProps): void
	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {SlideMasterProps} props layout definition
	 */
	defineSlideMaster(props: PptxGenJS.SlideMasterProps): void
	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} eleId table HTML element ID
	 * @param {TableToSlidesProps} props generation options
	 * @deprecated vNEXT - built on the estimating auto-pager; removed in v5.0. Migrate to `tableFromHtml` from @neo-ma/pptxgenjs-std/tables (see DEPRECATION-PLAN.md F10)
	 */
	tableToSlides(eleId: string, props?: PptxGenJS.TableToSlidesProps): void
}

declare namespace PptxGenJS {
	// Exported enums for module apps
	// @example: pptxgen.ShapeType.rect
	export enum AlignH {
		'left' = 'left',
		'center' = 'center',
		'right' = 'right',
		'justify' = 'justify',
	}
	export enum AlignV {
		'top' = 'top',
		'middle' = 'middle',
		'bottom' = 'bottom',
	}
	export enum ChartType {
		'area' = 'area',
		'bar' = 'bar',
		'bar3d' = 'bar3D',
		'bubble' = 'bubble',
		'bubble3d' = 'bubble3D',
		'doughnut' = 'doughnut',
		'line' = 'line',
		'pie' = 'pie',
		'radar' = 'radar',
		'scatter' = 'scatter',
		'boxWhisker' = 'boxWhisker',
		'funnel' = 'funnel',
		'histogram' = 'histogram',
		'sunburst' = 'sunburst',
		'treemap' = 'treemap',
		'waterfall' = 'waterfall',
	}
	export enum OutputType {
		'arraybuffer' = 'arraybuffer',
		'base64' = 'base64',
		'binarystring' = 'binarystring',
		'blob' = 'blob',
		'nodebuffer' = 'nodebuffer',
		'uint8array' = 'uint8array',
	}
	/**
	 * TODO: FUTURE: v4.0: rename to `SchemeColor`
	 */
	export enum SchemeColor {
		'text1' = 'tx1',
		'text2' = 'tx2',
		'background1' = 'bg1',
		'background2' = 'bg2',
		'accent1' = 'accent1',
		'accent2' = 'accent2',
		'accent3' = 'accent3',
		'accent4' = 'accent4',
		'accent5' = 'accent5',
		'accent6' = 'accent6',
	}
	export enum ShapeType {
		'accentBorderCallout1' = 'accentBorderCallout1',
		'accentBorderCallout2' = 'accentBorderCallout2',
		'accentBorderCallout3' = 'accentBorderCallout3',
		'accentCallout1' = 'accentCallout1',
		'accentCallout2' = 'accentCallout2',
		'accentCallout3' = 'accentCallout3',
		'actionButtonBackPrevious' = 'actionButtonBackPrevious',
		'actionButtonBeginning' = 'actionButtonBeginning',
		'actionButtonBlank' = 'actionButtonBlank',
		'actionButtonDocument' = 'actionButtonDocument',
		'actionButtonEnd' = 'actionButtonEnd',
		'actionButtonForwardNext' = 'actionButtonForwardNext',
		'actionButtonHelp' = 'actionButtonHelp',
		'actionButtonHome' = 'actionButtonHome',
		'actionButtonInformation' = 'actionButtonInformation',
		'actionButtonMovie' = 'actionButtonMovie',
		'actionButtonReturn' = 'actionButtonReturn',
		'actionButtonSound' = 'actionButtonSound',
		'arc' = 'arc',
		'bentArrow' = 'bentArrow',
		'bentUpArrow' = 'bentUpArrow',
		'bevel' = 'bevel',
		'blockArc' = 'blockArc',
		'borderCallout1' = 'borderCallout1',
		'borderCallout2' = 'borderCallout2',
		'borderCallout3' = 'borderCallout3',
		'bracePair' = 'bracePair',
		'bracketPair' = 'bracketPair',
		'callout1' = 'callout1',
		'callout2' = 'callout2',
		'callout3' = 'callout3',
		'can' = 'can',
		'chartPlus' = 'chartPlus',
		'chartStar' = 'chartStar',
		'chartX' = 'chartX',
		'chevron' = 'chevron',
		'chord' = 'chord',
		'circularArrow' = 'circularArrow',
		'cloud' = 'cloud',
		'cloudCallout' = 'cloudCallout',
		'corner' = 'corner',
		'cornerTabs' = 'cornerTabs',
		'cube' = 'cube',
		'curvedDownArrow' = 'curvedDownArrow',
		'curvedLeftArrow' = 'curvedLeftArrow',
		'curvedRightArrow' = 'curvedRightArrow',
		'curvedUpArrow' = 'curvedUpArrow',
		'decagon' = 'decagon',
		'diagStripe' = 'diagStripe',
		'diamond' = 'diamond',
		'dodecagon' = 'dodecagon',
		'donut' = 'donut',
		'doubleWave' = 'doubleWave',
		'downArrow' = 'downArrow',
		'downArrowCallout' = 'downArrowCallout',
		'ellipse' = 'ellipse',
		'ellipseRibbon' = 'ellipseRibbon',
		'ellipseRibbon2' = 'ellipseRibbon2',
		'flowChartAlternateProcess' = 'flowChartAlternateProcess',
		'flowChartCollate' = 'flowChartCollate',
		'flowChartConnector' = 'flowChartConnector',
		'flowChartDecision' = 'flowChartDecision',
		'flowChartDelay' = 'flowChartDelay',
		'flowChartDisplay' = 'flowChartDisplay',
		'flowChartDocument' = 'flowChartDocument',
		'flowChartExtract' = 'flowChartExtract',
		'flowChartInputOutput' = 'flowChartInputOutput',
		'flowChartInternalStorage' = 'flowChartInternalStorage',
		'flowChartMagneticDisk' = 'flowChartMagneticDisk',
		'flowChartMagneticDrum' = 'flowChartMagneticDrum',
		'flowChartMagneticTape' = 'flowChartMagneticTape',
		'flowChartManualInput' = 'flowChartManualInput',
		'flowChartManualOperation' = 'flowChartManualOperation',
		'flowChartMerge' = 'flowChartMerge',
		'flowChartMultidocument' = 'flowChartMultidocument',
		'flowChartOfflineStorage' = 'flowChartOfflineStorage',
		'flowChartOffpageConnector' = 'flowChartOffpageConnector',
		'flowChartOnlineStorage' = 'flowChartOnlineStorage',
		'flowChartOr' = 'flowChartOr',
		'flowChartPredefinedProcess' = 'flowChartPredefinedProcess',
		'flowChartPreparation' = 'flowChartPreparation',
		'flowChartProcess' = 'flowChartProcess',
		'flowChartPunchedCard' = 'flowChartPunchedCard',
		'flowChartPunchedTape' = 'flowChartPunchedTape',
		'flowChartSort' = 'flowChartSort',
		'flowChartSummingJunction' = 'flowChartSummingJunction',
		'flowChartTerminator' = 'flowChartTerminator',
		'folderCorner' = 'folderCorner',
		'frame' = 'frame',
		'funnel' = 'funnel',
		'gear6' = 'gear6',
		'gear9' = 'gear9',
		'halfFrame' = 'halfFrame',
		'heart' = 'heart',
		'heptagon' = 'heptagon',
		'hexagon' = 'hexagon',
		'homePlate' = 'homePlate',
		'horizontalScroll' = 'horizontalScroll',
		'irregularSeal1' = 'irregularSeal1',
		'irregularSeal2' = 'irregularSeal2',
		'leftArrow' = 'leftArrow',
		'leftArrowCallout' = 'leftArrowCallout',
		'leftBrace' = 'leftBrace',
		'leftBracket' = 'leftBracket',
		'leftCircularArrow' = 'leftCircularArrow',
		'leftRightArrow' = 'leftRightArrow',
		'leftRightArrowCallout' = 'leftRightArrowCallout',
		'leftRightCircularArrow' = 'leftRightCircularArrow',
		'leftRightRibbon' = 'leftRightRibbon',
		'leftRightUpArrow' = 'leftRightUpArrow',
		'leftUpArrow' = 'leftUpArrow',
		'lightningBolt' = 'lightningBolt',
		'line' = 'line',
		'lineInv' = 'lineInv',
		'mathDivide' = 'mathDivide',
		'mathEqual' = 'mathEqual',
		'mathMinus' = 'mathMinus',
		'mathMultiply' = 'mathMultiply',
		'mathNotEqual' = 'mathNotEqual',
		'mathPlus' = 'mathPlus',
		'moon' = 'moon',
		'nonIsoscelesTrapezoid' = 'nonIsoscelesTrapezoid',
		'noSmoking' = 'noSmoking',
		'notchedRightArrow' = 'notchedRightArrow',
		'octagon' = 'octagon',
		'parallelogram' = 'parallelogram',
		'pentagon' = 'pentagon',
		'pie' = 'pie',
		'pieWedge' = 'pieWedge',
		'plaque' = 'plaque',
		'plaqueTabs' = 'plaqueTabs',
		'plus' = 'plus',
		'quadArrow' = 'quadArrow',
		'quadArrowCallout' = 'quadArrowCallout',
		'rect' = 'rect',
		'ribbon' = 'ribbon',
		'ribbon2' = 'ribbon2',
		'rightArrow' = 'rightArrow',
		'rightArrowCallout' = 'rightArrowCallout',
		'rightBrace' = 'rightBrace',
		'rightBracket' = 'rightBracket',
		'round1Rect' = 'round1Rect',
		'round2DiagRect' = 'round2DiagRect',
		'round2SameRect' = 'round2SameRect',
		'roundRect' = 'roundRect',
		'rtTriangle' = 'rtTriangle',
		'smileyFace' = 'smileyFace',
		'snip1Rect' = 'snip1Rect',
		'snip2DiagRect' = 'snip2DiagRect',
		'snip2SameRect' = 'snip2SameRect',
		'snipRoundRect' = 'snipRoundRect',
		'squareTabs' = 'squareTabs',
		'star10' = 'star10',
		'star12' = 'star12',
		'star16' = 'star16',
		'star24' = 'star24',
		'star32' = 'star32',
		'star4' = 'star4',
		'star5' = 'star5',
		'star6' = 'star6',
		'star7' = 'star7',
		'star8' = 'star8',
		'stripedRightArrow' = 'stripedRightArrow',
		'sun' = 'sun',
		'swooshArrow' = 'swooshArrow',
		'teardrop' = 'teardrop',
		'trapezoid' = 'trapezoid',
		'triangle' = 'triangle',
		'upArrow' = 'upArrow',
		'upArrowCallout' = 'upArrowCallout',
		'upDownArrow' = 'upDownArrow',
		'upDownArrowCallout' = 'upDownArrowCallout',
		'uturnArrow' = 'uturnArrow',
		'verticalScroll' = 'verticalScroll',
		'wave' = 'wave',
		'wedgeEllipseCallout' = 'wedgeEllipseCallout',
		'wedgeRectCallout' = 'wedgeRectCallout',
		'wedgeRoundRectCallout' = 'wedgeRoundRectCallout',
	}
	// used by charts, shape, text
	export interface BorderOptions {
		/**
		 * Border type
		 */
		type?: 'none' | 'dash' | 'solid'
		/**
		 * Border color (hex)
		 * @example 'FF3399'
		 */
		color?: HexColor
		/**
		 * Border size (points)
		 */
		pt?: number
	}
	// These are used by browser/script clients and have been named like this since v0.1.
	// Desc: charts and shapes for `pptxgen.charts.` `pptxgen.shapes.`
	// Note: "charts" and "shapes" are manually created by cloning
	export enum charts {
		'AREA' = 'area',
		'BAR' = 'bar',
		'BAR3D' = 'bar3D',
		'BUBBLE' = 'bubble',
		'DOUGHNUT' = 'doughnut',
		'LINE' = 'line',
		'PIE' = 'pie',
		'RADAR' = 'radar',
		'SCATTER' = 'scatter',
	}
	export enum shapes {
		ACTION_BUTTON_BACK_OR_PREVIOUS = 'actionButtonBackPrevious',
		ACTION_BUTTON_BEGINNING = 'actionButtonBeginning',
		ACTION_BUTTON_CUSTOM = 'actionButtonBlank',
		ACTION_BUTTON_DOCUMENT = 'actionButtonDocument',
		ACTION_BUTTON_END = 'actionButtonEnd',
		ACTION_BUTTON_FORWARD_OR_NEXT = 'actionButtonForwardNext',
		ACTION_BUTTON_HELP = 'actionButtonHelp',
		ACTION_BUTTON_HOME = 'actionButtonHome',
		ACTION_BUTTON_INFORMATION = 'actionButtonInformation',
		ACTION_BUTTON_MOVIE = 'actionButtonMovie',
		ACTION_BUTTON_RETURN = 'actionButtonReturn',
		ACTION_BUTTON_SOUND = 'actionButtonSound',
		ARC = 'arc',
		BALLOON = 'wedgeRoundRectCallout',
		BENT_ARROW = 'bentArrow',
		BENT_UP_ARROW = 'bentUpArrow',
		BEVEL = 'bevel',
		BLOCK_ARC = 'blockArc',
		CAN = 'can',
		CHART_PLUS = 'chartPlus',
		CHART_STAR = 'chartStar',
		CHART_X = 'chartX',
		CHEVRON = 'chevron',
		CHORD = 'chord',
		CIRCULAR_ARROW = 'circularArrow',
		CLOUD = 'cloud',
		CLOUD_CALLOUT = 'cloudCallout',
		CORNER = 'corner',
		CORNER_TABS = 'cornerTabs',
		CROSS = 'plus',
		CUBE = 'cube',
		CURVED_DOWN_ARROW = 'curvedDownArrow',
		CURVED_DOWN_RIBBON = 'ellipseRibbon',
		CURVED_LEFT_ARROW = 'curvedLeftArrow',
		CURVED_RIGHT_ARROW = 'curvedRightArrow',
		CURVED_UP_ARROW = 'curvedUpArrow',
		CURVED_UP_RIBBON = 'ellipseRibbon2',
		DECAGON = 'decagon',
		DIAGONAL_STRIPE = 'diagStripe',
		DIAMOND = 'diamond',
		DODECAGON = 'dodecagon',
		DONUT = 'donut',
		DOUBLE_BRACE = 'bracePair',
		DOUBLE_BRACKET = 'bracketPair',
		DOUBLE_WAVE = 'doubleWave',
		DOWN_ARROW = 'downArrow',
		DOWN_ARROW_CALLOUT = 'downArrowCallout',
		DOWN_RIBBON = 'ribbon',
		EXPLOSION1 = 'irregularSeal1',
		EXPLOSION2 = 'irregularSeal2',
		FLOWCHART_ALTERNATE_PROCESS = 'flowChartAlternateProcess',
		FLOWCHART_CARD = 'flowChartPunchedCard',
		FLOWCHART_COLLATE = 'flowChartCollate',
		FLOWCHART_CONNECTOR = 'flowChartConnector',
		FLOWCHART_DATA = 'flowChartInputOutput',
		FLOWCHART_DECISION = 'flowChartDecision',
		FLOWCHART_DELAY = 'flowChartDelay',
		FLOWCHART_DIRECT_ACCESS_STORAGE = 'flowChartMagneticDrum',
		FLOWCHART_DISPLAY = 'flowChartDisplay',
		FLOWCHART_DOCUMENT = 'flowChartDocument',
		FLOWCHART_EXTRACT = 'flowChartExtract',
		FLOWCHART_INTERNAL_STORAGE = 'flowChartInternalStorage',
		FLOWCHART_MAGNETIC_DISK = 'flowChartMagneticDisk',
		FLOWCHART_MANUAL_INPUT = 'flowChartManualInput',
		FLOWCHART_MANUAL_OPERATION = 'flowChartManualOperation',
		FLOWCHART_MERGE = 'flowChartMerge',
		FLOWCHART_MULTIDOCUMENT = 'flowChartMultidocument',
		FLOWCHART_OFFLINE_STORAGE = 'flowChartOfflineStorage',
		FLOWCHART_OFFPAGE_CONNECTOR = 'flowChartOffpageConnector',
		FLOWCHART_OR = 'flowChartOr',
		FLOWCHART_PREDEFINED_PROCESS = 'flowChartPredefinedProcess',
		FLOWCHART_PREPARATION = 'flowChartPreparation',
		FLOWCHART_PROCESS = 'flowChartProcess',
		FLOWCHART_PUNCHED_TAPE = 'flowChartPunchedTape',
		FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 'flowChartMagneticTape',
		FLOWCHART_SORT = 'flowChartSort',
		FLOWCHART_STORED_DATA = 'flowChartOnlineStorage',
		FLOWCHART_SUMMING_JUNCTION = 'flowChartSummingJunction',
		FLOWCHART_TERMINATOR = 'flowChartTerminator',
		FOLDED_CORNER = 'folderCorner',
		FRAME = 'frame',
		FUNNEL = 'funnel',
		GEAR_6 = 'gear6',
		GEAR_9 = 'gear9',
		HALF_FRAME = 'halfFrame',
		HEART = 'heart',
		HEPTAGON = 'heptagon',
		HEXAGON = 'hexagon',
		HORIZONTAL_SCROLL = 'horizontalScroll',
		ISOSCELES_TRIANGLE = 'triangle',
		LEFT_ARROW = 'leftArrow',
		LEFT_ARROW_CALLOUT = 'leftArrowCallout',
		LEFT_BRACE = 'leftBrace',
		LEFT_BRACKET = 'leftBracket',
		LEFT_CIRCULAR_ARROW = 'leftCircularArrow',
		LEFT_RIGHT_ARROW = 'leftRightArrow',
		LEFT_RIGHT_ARROW_CALLOUT = 'leftRightArrowCallout',
		LEFT_RIGHT_CIRCULAR_ARROW = 'leftRightCircularArrow',
		LEFT_RIGHT_RIBBON = 'leftRightRibbon',
		LEFT_RIGHT_UP_ARROW = 'leftRightUpArrow',
		LEFT_UP_ARROW = 'leftUpArrow',
		LIGHTNING_BOLT = 'lightningBolt',
		LINE_CALLOUT_1 = 'borderCallout1',
		LINE_CALLOUT_1_ACCENT_BAR = 'accentCallout1',
		LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 'accentBorderCallout1',
		LINE_CALLOUT_1_NO_BORDER = 'callout1',
		LINE_CALLOUT_2 = 'borderCallout2',
		LINE_CALLOUT_2_ACCENT_BAR = 'accentCallout2',
		LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 'accentBorderCallout2',
		LINE_CALLOUT_2_NO_BORDER = 'callout2',
		LINE_CALLOUT_3 = 'borderCallout3',
		LINE_CALLOUT_3_ACCENT_BAR = 'accentCallout3',
		LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 'accentBorderCallout3',
		LINE_CALLOUT_3_NO_BORDER = 'callout3',
		LINE_CALLOUT_4 = 'borderCallout4',
		LINE_CALLOUT_4_ACCENT_BAR = 'accentCallout4',
		LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 'accentBorderCallout4',
		LINE_CALLOUT_4_NO_BORDER = 'callout4',
		LINE = 'line',
		LINE_INVERSE = 'lineInv',
		MATH_DIVIDE = 'mathDivide',
		MATH_EQUAL = 'mathEqual',
		MATH_MINUS = 'mathMinus',
		MATH_MULTIPLY = 'mathMultiply',
		MATH_NOT_EQUAL = 'mathNotEqual',
		MATH_PLUS = 'mathPlus',
		MOON = 'moon',
		NON_ISOSCELES_TRAPEZOID = 'nonIsoscelesTrapezoid',
		NOTCHED_RIGHT_ARROW = 'notchedRightArrow',
		NO_SYMBOL = 'noSmoking',
		OCTAGON = 'octagon',
		OVAL = 'ellipse',
		OVAL_CALLOUT = 'wedgeEllipseCallout',
		PARALLELOGRAM = 'parallelogram',
		PENTAGON = 'homePlate',
		PIE = 'pie',
		PIE_WEDGE = 'pieWedge',
		PLAQUE = 'plaque',
		PLAQUE_TABS = 'plaqueTabs',
		QUAD_ARROW = 'quadArrow',
		QUAD_ARROW_CALLOUT = 'quadArrowCallout',
		RECTANGLE = 'rect',
		RECTANGULAR_CALLOUT = 'wedgeRectCallout',
		REGULAR_PENTAGON = 'pentagon',
		RIGHT_ARROW = 'rightArrow',
		RIGHT_ARROW_CALLOUT = 'rightArrowCallout',
		RIGHT_BRACE = 'rightBrace',
		RIGHT_BRACKET = 'rightBracket',
		RIGHT_TRIANGLE = 'rtTriangle',
		ROUNDED_RECTANGLE = 'roundRect',
		ROUNDED_RECTANGULAR_CALLOUT = 'wedgeRoundRectCallout',
		ROUND_1_RECTANGLE = 'round1Rect',
		ROUND_2_DIAG_RECTANGLE = 'round2DiagRect',
		ROUND_2_SAME_RECTANGLE = 'round2SameRect',
		SMILEY_FACE = 'smileyFace',
		SNIP_1_RECTANGLE = 'snip1Rect',
		SNIP_2_DIAG_RECTANGLE = 'snip2DiagRect',
		SNIP_2_SAME_RECTANGLE = 'snip2SameRect',
		SNIP_ROUND_RECTANGLE = 'snipRoundRect',
		SQUARE_TABS = 'squareTabs',
		STAR_10_POINT = 'star10',
		STAR_12_POINT = 'star12',
		STAR_16_POINT = 'star16',
		STAR_24_POINT = 'star24',
		STAR_32_POINT = 'star32',
		STAR_4_POINT = 'star4',
		STAR_5_POINT = 'star5',
		STAR_6_POINT = 'star6',
		STAR_7_POINT = 'star7',
		STAR_8_POINT = 'star8',
		STRIPED_RIGHT_ARROW = 'stripedRightArrow',
		SUN = 'sun',
		SWOOSH_ARROW = 'swooshArrow',
		TEAR = 'teardrop',
		TRAPEZOID = 'trapezoid',
		UP_ARROW = 'upArrow',
		UP_ARROW_CALLOUT = 'upArrowCallout',
		UP_DOWN_ARROW = 'upDownArrow',
		UP_DOWN_ARROW_CALLOUT = 'upDownArrowCallout',
		UP_RIBBON = 'ribbon2',
		U_TURN_ARROW = 'uturnArrow',
		VERTICAL_SCROLL = 'verticalScroll',
		WAVE = 'wave',
	}

	// @source `core-enums.ts`
	export type JSZIP_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	export type WRITE_OUTPUT_TYPE = JSZIP_OUTPUT_TYPE | 'STREAM'
	export enum CHART_TYPE {
		'AREA' = 'area',
		'BAR' = 'bar',
		'BAR3D' = 'bar3D',
		'BUBBLE' = 'bubble',
		'DOUGHNUT' = 'doughnut',
		'LINE' = 'line',
		'PIE' = 'pie',
		'RADAR' = 'radar',
		'SCATTER' = 'scatter',
	}
	export enum SCHEME_COLOR_NAMES {
		'TEXT1' = 'tx1',
		'TEXT2' = 'tx2',
		'BACKGROUND1' = 'bg1',
		'BACKGROUND2' = 'bg2',
		'ACCENT1' = 'accent1',
		'ACCENT2' = 'accent2',
		'ACCENT3' = 'accent3',
		'ACCENT4' = 'accent4',
		'ACCENT5' = 'accent5',
		'ACCENT6' = 'accent6',
	}

	// @source `core-interfaces.d.ts` (via import)
	// @code `import { CHART_NAME, PLACEHOLDER_TYPES, SHAPE_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN, WRITE_OUTPUT_TYPE } from './core-enums'`
	export type CHART_NAME = 'area' | 'bar' | 'bar3D' | 'bubble' | 'doughnut' | 'line' | 'pie' | 'radar' | 'scatter'
		| CHARTEX_NAME
	/** PowerPoint 2016+ chart types; emitted as a `cx:chartSpace` part (MS-ODRAWXML 2.1) rather than ECMA-376 `c:chartSpace` */
	export type CHARTEX_NAME = 'boxWhisker' | 'funnel' | 'histogram' | 'sunburst' | 'treemap' | 'waterfall'
	export enum PLACEHOLDER_TYPES {
		'title' = 'title',
		'body' = 'body',
		'image' = 'pic',
		'chart' = 'chart',
		'table' = 'tbl',
		'media' = 'media',
	}
	export type PLACEHOLDER_TYPE = 'title' | 'body' | 'pic' | 'chart' | 'tbl' | 'media'

	export type SHAPE_NAME =
		| 'accentBorderCallout1'
		| 'accentBorderCallout2'
		| 'accentBorderCallout3'
		| 'accentCallout1'
		| 'accentCallout2'
		| 'accentCallout3'
		| 'actionButtonBackPrevious'
		| 'actionButtonBeginning'
		| 'actionButtonBlank'
		| 'actionButtonDocument'
		| 'actionButtonEnd'
		| 'actionButtonForwardNext'
		| 'actionButtonHelp'
		| 'actionButtonHome'
		| 'actionButtonInformation'
		| 'actionButtonMovie'
		| 'actionButtonReturn'
		| 'actionButtonSound'
		| 'arc'
		| 'bentArrow'
		| 'bentUpArrow'
		| 'bevel'
		| 'blockArc'
		| 'borderCallout1'
		| 'borderCallout2'
		| 'borderCallout3'
		| 'bracePair'
		| 'bracketPair'
		| 'callout1'
		| 'callout2'
		| 'callout3'
		| 'can'
		| 'chartPlus'
		| 'chartStar'
		| 'chartX'
		| 'chevron'
		| 'chord'
		| 'circularArrow'
		| 'cloud'
		| 'cloudCallout'
		| 'corner'
		| 'cornerTabs'
		| 'cube'
		| 'curvedDownArrow'
		| 'curvedLeftArrow'
		| 'curvedRightArrow'
		| 'curvedUpArrow'
		| 'decagon'
		| 'diagStripe'
		| 'diamond'
		| 'dodecagon'
		| 'donut'
		| 'doubleWave'
		| 'downArrow'
		| 'downArrowCallout'
		| 'ellipse'
		| 'ellipseRibbon'
		| 'ellipseRibbon2'
		| 'flowChartAlternateProcess'
		| 'flowChartCollate'
		| 'flowChartConnector'
		| 'flowChartDecision'
		| 'flowChartDelay'
		| 'flowChartDisplay'
		| 'flowChartDocument'
		| 'flowChartExtract'
		| 'flowChartInputOutput'
		| 'flowChartInternalStorage'
		| 'flowChartMagneticDisk'
		| 'flowChartMagneticDrum'
		| 'flowChartMagneticTape'
		| 'flowChartManualInput'
		| 'flowChartManualOperation'
		| 'flowChartMerge'
		| 'flowChartMultidocument'
		| 'flowChartOfflineStorage'
		| 'flowChartOffpageConnector'
		| 'flowChartOnlineStorage'
		| 'flowChartOr'
		| 'flowChartPredefinedProcess'
		| 'flowChartPreparation'
		| 'flowChartProcess'
		| 'flowChartPunchedCard'
		| 'flowChartPunchedTape'
		| 'flowChartSort'
		| 'flowChartSummingJunction'
		| 'flowChartTerminator'
		| 'folderCorner'
		| 'frame'
		| 'funnel'
		| 'gear6'
		| 'gear9'
		| 'halfFrame'
		| 'heart'
		| 'heptagon'
		| 'hexagon'
		| 'homePlate'
		| 'horizontalScroll'
		| 'irregularSeal1'
		| 'irregularSeal2'
		| 'leftArrow'
		| 'leftArrowCallout'
		| 'leftBrace'
		| 'leftBracket'
		| 'leftCircularArrow'
		| 'leftRightArrow'
		| 'leftRightArrowCallout'
		| 'leftRightCircularArrow'
		| 'leftRightRibbon'
		| 'leftRightUpArrow'
		| 'leftUpArrow'
		| 'lightningBolt'
		| 'line'
		| 'lineInv'
		| 'mathDivide'
		| 'mathEqual'
		| 'mathMinus'
		| 'mathMultiply'
		| 'mathNotEqual'
		| 'mathPlus'
		| 'moon'
		| 'noSmoking'
		| 'nonIsoscelesTrapezoid'
		| 'notchedRightArrow'
		| 'octagon'
		| 'parallelogram'
		| 'pentagon'
		| 'pie'
		| 'pieWedge'
		| 'plaque'
		| 'plaqueTabs'
		| 'plus'
		| 'quadArrow'
		| 'quadArrowCallout'
		| 'rect'
		| 'ribbon'
		| 'ribbon2'
		| 'rightArrow'
		| 'rightArrowCallout'
		| 'rightBrace'
		| 'rightBracket'
		| 'round1Rect'
		| 'round2DiagRect'
		| 'round2SameRect'
		| 'roundRect'
		| 'rtTriangle'
		| 'smileyFace'
		| 'snip1Rect'
		| 'snip2DiagRect'
		| 'snip2SameRect'
		| 'snipRoundRect'
		| 'squareTabs'
		| 'star10'
		| 'star12'
		| 'star16'
		| 'star24'
		| 'star32'
		| 'star4'
		| 'star5'
		| 'star6'
		| 'star7'
		| 'star8'
		| 'stripedRightArrow'
		| 'sun'
		| 'swooshArrow'
		| 'teardrop'
		| 'trapezoid'
		| 'triangle'
		| 'upArrow'
		| 'upArrowCallout'
		| 'upDownArrow'
		| 'upDownArrowCallout'
		| 'uturnArrow'
		| 'verticalScroll'
		| 'wave'
		| 'wedgeEllipseCallout'
		| 'wedgeRectCallout'
		| 'wedgeRoundRectCallout'

	export enum SLIDE_OBJECT_TYPES {
		'chart' = 'chart',
		'hyperlink' = 'hyperlink',
		'image' = 'image',
		'media' = 'media',
		'online' = 'online',
		'placeholder' = 'placeholder',
		'table' = 'table',
		'tablecell' = 'tablecell',
		'text' = 'text',
		'notes' = 'notes',
	}
	export enum TEXT_HALIGN {
		'left' = 'left',
		'center' = 'center',
		'right' = 'right',
		'justify' = 'justify',
	}
	export enum TEXT_VALIGN {
		'b' = 'b',
		'ctr' = 'ctr',
		't' = 't',
	}

	// @source `core-interfaces.d.ts` (direct)
	// Core Types
	// ==========

	/**
	 * Coordinate number - either:
	 * - Inches (0-n)
	 * - Percentage (0-100)
	 *
	 * @example 10.25 // coordinate in inches
	 * @example '2.5cm' // coordinate with an explicit unit (in/cm/mm/pt)
	 * @example '75%' // coordinate as percentage of slide size
	 */
	export type UnitLength = `${number}in` | `${number}cm` | `${number}mm` | `${number}pt`
	export type Coord = number | `${number}%` | UnitLength
	export interface PositionProps {
		/**
		 * Horizontal position
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		x?: Coord
		/**
		 * Vertical position
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		y?: Coord
		/**
		 * Height
		 * - inches or percentage
		 * @example 10.25 // height in inches
		 * @example '75%' // height as percentage of slide size
		 */
		h?: Coord
		/**
		 * Width
		 * - inches or percentage
		 * @example 10.25 // width in inches
		 * @example '75%' // width as percentage of slide size
		 */
		w?: Coord
	}
	/**
	 * Either `data` or `path` is required
	 */
	export interface DataOrPathProps {
		/**
		 * URL or relative path
		 *
		 * @example 'https://onedrives.com/myimg.png` // retrieve image via URL
		 * @example '/home/gitbrent/images/myimg.png` // retrieve image via local path
		 */
		path?: string
		/**
		 * base64-encoded string
		 * - Useful for avoiding potential path/server issues
		 *
		 * @example 'image/png;base64,iVtDafDrBF[...]=' // pre-encoded image in base-64
		 */
		data?: string
	}
	export interface BackgroundProps extends DataOrPathProps, ShapeFillProps {
		/**
		 * Color (hex format)
		 * @deprecated v3.6.0 - use `ShapeFillProps` instead
		 */
		fill?: HexColor

		/**
		 * source URL
		 * @deprecated v3.6.0 - use `DataOrPathProps` instead - remove in v4.0.0
		 */
		src?: string
		/**
		 * Recolour and correction effects applied to a background image
		 * @example { grayscale: true, brightness: -20 }
		 */
		recolor?: ImageRecolorProps
	}
	/**
	 * Color in Hex format
	 * @example 'FF3399'
	 */
	export type HexColor = string
	export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
	export type SchemeColorValue =
		| 'bg1' | 'tx1' | 'bg2' | 'tx2' | 'dk1' | 'lt1' | 'dk2' | 'lt2'
		| 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
		| 'hlink' | 'folHlink' | 'phClr'
	/**
	 * Color transforms (ECMA-376 20.1.2.3, `EG_ColorTransform`)
	 * - percentages are 0-100; offsets are -100-100; `hueOff` is in degrees
	 * - any number of transforms may be combined on one color
	 */
	export interface ColorTransformProps {
		/** lighten toward white (percent) */
		tint?: number
		/** darken toward black (percent) */
		shade?: number
		/** opacity (percent); 100 is opaque */
		alpha?: number
		/** shift opacity (percent, -100 to 100) */
		alphaOff?: number
		/** scale opacity (percent) */
		alphaMod?: number
		/** scale luminance (percent) */
		lumMod?: number
		/** shift luminance (percent, -100 to 100) */
		lumOff?: number
		/** scale saturation (percent) */
		satMod?: number
		/** shift saturation (percent, -100 to 100) */
		satOff?: number
		/** scale hue (percent) */
		hueMod?: number
		/** shift hue (degrees, -360 to 360) */
		hueOff?: number
		/** use the complement */
		complement?: boolean
		/** use the inverse */
		inverse?: boolean
		/** convert to grayscale */
		grayscale?: boolean
		/** apply gamma */
		gamma?: boolean
		/** apply inverse gamma */
		inverseGamma?: boolean
	}
	/**
	 * A color given as an object, so transforms and the non-hex DrawingML color kinds are reachable
	 * - exactly one specification field identifies the color; that is what distinguishes a color
	 *   object from a fill object at runtime
	 * @example { hex: 'FF0000', alpha: 50 }
	 * @example { scheme: 'accent1', lumMod: 60, lumOff: 40 }
	 * @example { preset: 'cornflowerBlue' }
	 * @example { hsl: { hue: 210, sat: 80, lum: 50 }, shade: 25 }
	 */
	export type ColorProps =
		/** 6-digit hex, as the string form accepts (`a:srgbClr`) */
		| ({ hex: HexColor } & ColorTransformProps)
		/** theme slot (`a:schemeClr`) */
		| ({ scheme: SchemeColorValue } & ColorTransformProps)
		/** system color such as `windowText`, with an optional last-known value (`a:sysClr`) */
		| ({ system: string, lastColor?: HexColor } & ColorTransformProps)
		/** one of the 140 preset color names (`a:prstClr`) */
		| ({ preset: string } & ColorTransformProps)
		/** hue/saturation/luminance; hue in degrees, the rest percent (`a:hslClr`) */
		| ({ hsl: { hue: number, sat: number, lum: number } } & ColorTransformProps)
		/** linear-gamma RGB percentages (`a:scrgbClr`) */
		| ({ scrgb: { r: number, g: number, b: number } } & ColorTransformProps)
	export type Color = HexColor | ThemeColor | ColorProps
	export type Margin = number | [number, number, number, number]
	export type HAlign = 'left' | 'center' | 'right' | 'justify'
	export type VAlign = 'top' | 'middle' | 'bottom'

	// used by charts, shape, text
	export interface BorderProps {
		/**
		 * Border type
		 * @default solid
		 */
		type?: 'none' | 'dash' | 'solid'
		/**
		 * Border color (hex)
		 * @example 'FF3399'
		 * @default '666666'
		 */
		color?: HexColor

		// TODO: add `transparency` prop to Borders (0-100%)

		/**
		 * Border width (points)
		 * - same name and unit as `ShapeLineProps.width`
		 * @default 1
		 */
		width?: number
		/**
		 * Border size (points)
		 * @deprecated v4.1.0 - use `width`
		 * @default 1
		 */
		pt?: number
	}
	// used by: image, object, text,
	export interface HyperlinkSoundProps {
		/**
		 * WAV audio data (base64), with a mime header
		 * - one of `data` or `path` is required
		 * @example 'audio/wav;base64,UklGRi...'
		 */
		data?: string
		/**
		 * WAV file path or URL
		 * - one of `data` or `path` is required
		 */
		path?: string
		/**
		 * Sound name PowerPoint shows in the Action dialog
		 * @default 'sound.wav'
		 */
		name?: string
	}
	export interface HyperlinkProps {
		//_rId: number
		/**
		 * Slide number to link to
		 */
		slide?: number
		/**
		 * Url to link to
		 */
		url?: string
		/**
		 * Hyperlink Tooltip
		 */
		tooltip?: string
			/**
			 * Highlight the link when clicked
			 * @default false
			 */
			highlightClick?: boolean
			/**
			 * Stop any playing sounds when the link is clicked
			 * @default false
			 */
			stopSoundsOnClick?: boolean
			/**
			 * Sound played when the link is triggered (`a:snd`)
			 * - must be WAV data; ECMA-376 20.1.2.2.32 allows no other format here
			 */
			sound?: HyperlinkSoundProps
	}
	// used by: chart, text, image
	export interface ShadowProps {
		/**
		 * shadow type
		 * @default 'none'
		 */
		type: 'outer' | 'inner' | 'none' | 'preset'
		/**
		 * Preset shadow name, required when `type` is `preset` (`a:prstShdw@prst`)
		 * - PowerPoint's twenty built-in shadow presets
		 */
		preset?: 'shdw1' | 'shdw2' | 'shdw3' | 'shdw4' | 'shdw5' | 'shdw6' | 'shdw7' | 'shdw8' | 'shdw9' | 'shdw10' | 'shdw11' | 'shdw12' | 'shdw13' | 'shdw14' | 'shdw15' | 'shdw16' | 'shdw17' | 'shdw18' | 'shdw19' | 'shdw20'
		/**
		 * opacity (percent)
		 * - range: 0.0-1.0
		 * @example 0.5 // 50% opaque
		 */
		opacity?: number // TODO: "Transparency (0-100%)" in PPT // TODO: deprecate and add `transparency`
		/**
		 * blur (points)
		 * - range: 0-100
		 * @default 0
		 */
		blur?: number
		/**
		 * angle (degrees)
		 * - range: 0-359
		 * @default 0
		 */
		angle?: number
		/**
		 * shadow offset (points)
		 * - range: 0-200
		 * @default 0
		 */
		offset?: number // TODO: "Distance" in PPT
		/**
		 * shadow color (hex format)
		 * @example 'FF3399'
		 */
		color?: HexColor
		/**
		 * whether to rotate shadow with shape
		 * @default false
		 */
		rotateWithShape?: boolean
	}
	// used by: shape, table, text
	export interface ShapeGradientStopProps {
		/**
		 * Stop color
		 * - `HexColor` or `ThemeColor`
		 */
		color: Color
		/**
		 * Stop position along the gradient (percent)
		 * - range: 0-100
		 */
		position: number
		/**
		 * Transparency (percent)
		 * - range: 0-100
		 * @default 0
		 */
		transparency?: number
	}
	export interface ShapeGradientProps {
		/**
		 * Gradient geometry
		 * @default 'linear'
		 */
		type?: 'linear' | 'radial'
		/**
		 * Linear gradient angle (degrees, clockwise)
		 * - 0 = left-to-right, 90 = top-to-bottom
		 * - normalized into the range 0-359
		 * - ignored when `type` is `'radial'`
		 * @default 90
		 */
		angle?: number
		/**
		 * Whether the linear gradient angle scales with the fill region
		 * @default false
		 */
		scaled?: boolean
		/**
		 * Whether the gradient rotates with its shape
		 * @default true
		 */
		rotateWithShape?: boolean
		/**
		 * Gradient color stops
		 * - MS-PPT requires **at least 2 stops**; fewer falls back to a solid fill
		 * - stops are sorted by `position` before they are written
		 */
		stops: ShapeGradientStopProps[]
	}
	export type AnimationPreset =
		// entrance
		| 'appear' | 'fadeIn' | 'wipeIn' | 'zoomIn'
		// exit
		| 'disappear' | 'fadeOut' | 'wipeOut' | 'zoomOut'
	export interface AnimationProps {
		/**
		 * Animation preset
		 * - entrance: `appear`, `fadeIn`, `wipeIn`, `zoomIn`
		 * - exit: `disappear`, `fadeOut`, `wipeOut`, `zoomOut`
		 * - motion-path and emphasis effects are not supported yet
		 */
		type: AnimationPreset
		/**
		 * What starts the effect
		 * - `withPrevious` and `afterPrevious` join the click group before them
		 * @default 'onClick'
		 */
		trigger?: 'onClick' | 'withPrevious' | 'afterPrevious'
		/**
		 * Effect direction
		 * - `wipeIn`/`wipeOut`: `up` | `right` | `down` | `left` (default `up`)
		 * - `zoomIn`/`zoomOut`: `in` | `out` (default matches the preset)
		 * - ignored with a warning for presets that take no direction
		 */
		direction?: string
		/**
		 * Delay before the effect starts (milliseconds)
		 * @default 0
		 */
		delay?: number
		/**
		 * Effect length (milliseconds)
		 * @default 500
		 */
		duration?: number
	}
	export type PatternType =
		// ECMA-376 20.1.10.51 ST_PresetPatternVal - all 54 preset patterns
		| 'pct5' | 'pct10' | 'pct20' | 'pct25' | 'pct30' | 'pct40' | 'pct50' | 'pct60' | 'pct70' | 'pct75'
		| 'pct80' | 'pct90'
		| 'horz' | 'vert' | 'ltHorz' | 'ltVert' | 'dkHorz' | 'dkVert' | 'narHorz' | 'narVert'
		| 'dashHorz' | 'dashVert'
		| 'cross' | 'dnDiag' | 'upDiag' | 'ltDnDiag' | 'ltUpDiag' | 'dkDnDiag' | 'dkUpDiag'
		| 'wdDnDiag' | 'wdUpDiag' | 'dashDnDiag' | 'dashUpDiag' | 'diagCross'
		| 'smCheck' | 'lgCheck' | 'smGrid' | 'lgGrid' | 'dotGrid' | 'smConfetti' | 'lgConfetti'
		| 'horzBrick' | 'diagBrick' | 'solidDmnd' | 'openDmnd' | 'dotDmnd'
		| 'plaid' | 'sphere' | 'weave' | 'divot' | 'shingle' | 'wave' | 'trellis' | 'zigZag'
	export interface ShapePatternProps {
		/**
		 * Preset pattern (ECMA-376 20.1.8.47 `a:pattFill@prst`)
		 */
		preset: PatternType
		/**
		 * Foreground (pattern line) color
		 * - `HexColor` or `ThemeColor`
		 * @default '000000'
		 */
		color?: Color
		/**
		 * Background color behind the pattern
		 * - `HexColor` or `ThemeColor`
		 * @default 'FFFFFF'
		 */
		backColor?: Color
	}
	export interface ShapeImageFillProps {
		/**
		 * Image data (base64), with a mime header
		 * - one of `data` or `path` is required
		 * @example 'image/png;base64,iVBORw0KGgo...'
		 */
		data?: string
		/**
		 * Image path or URL
		 * - one of `data` or `path` is required
		 */
		path?: string
		/**
		 * How the image fills the shape
		 * - `stretch` scales it to the shape (ECMA-376 20.1.8.56)
		 * - `tile` repeats it at its natural size (20.1.8.58)
		 * @default 'stretch'
		 */
		sizing?: 'stretch' | 'tile'
		/**
		 * `tile` only: scale applied to each tile (percent)
		 * @default 100
		 */
		scale?: number
		/**
		 * `tile` only: where tiling starts
		 * @default 'tl'
		 */
		alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br'
		/**
		 * Whether the fill rotates with the shape
		 * @default true
		 */
		rotateWithShape?: boolean
	}
	export interface ShapeFillProps {
		/**
		 * Fill color
		 * - `HexColor` or `ThemeColor`
		 * @example 'FF0000' // hex color (red)
		 * @example pptx.SchemeColor.text1 // Theme color (Text1)
		 */
		color?: Color
		/**
		 * Transparency (percent)
		 * - MS-PPT > Format Shape > Fill & Line > Fill > Transparency
		 * - range: 0-100
		 * @default 0
		 */
		transparency?: number
		/**
		 * Fill type
		 * @default 'solid'
		 */
		type?: 'none' | 'solid' | 'gradient' | 'pattern' | 'image' | 'group'
		/**
		 * Pattern fill definition
		 * - required when `type` is `'pattern'`
		 * @example { type:'pattern', pattern:{ preset:'diagCross', color:'0000FF', backColor:'FFFFFF' } }
		 */
		pattern?: ShapePatternProps
		/**
		 * Picture fill definition
		 * - required when `type` is `'image'`
		 * @example { type:'image', image:{ data:'image/png;base64,iV[...]', sizing:'tile' } }
		 */
		image?: ShapeImageFillProps
		/**
		 * Gradient fill definition
		 * - required when `type` is `'gradient'`
		 * @example { type:'gradient', gradient:{ angle:90, stops:[{ color:'FF0000', position:0 }, { color:'0000FF', position:100 }] } }
		 */
		gradient?: ShapeGradientProps

		/**
		 * Transparency (percent)
		 * @deprecated v3.3.0 - use `transparency`
		 */
		alpha?: number
	}
	export interface ShapeLineProps extends ShapeFillProps {
		/**
		 * Line width (pt)
		 * @default 1
		 */
		width?: number
		/**
		 * Dash type
		 * @default 'solid'
		 */
		dashType?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * Begin arrow type
		 * @since v3.3.0
		 */
		beginArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * End arrow type
		 * @since v3.3.0
		 */
		endArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * Compound line type (`a:ln@cmpd`)
		 * @default 'sng'
		 */
		compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri'
		/**
		 * How line segments join at a corner (`a:round` / `a:bevel` / `a:miter`)
		 * @default 'round'
		 */
		join?: 'round' | 'bevel' | 'miter'
		/**
		 * `join: 'miter'` only: how far the miter may extend, as a percent of line width (`a:miter@lim`)
		 * @default 800
		 */
		miterLimit?: number
		/**
		 * Custom dash pattern (`a:custDash`), overriding `dashType`
		 * - each stop is a dash length and the gap after it, as a percent of line width
		 * @example [{ dash: 400, space: 300 }, { dash: 100, space: 300 }]
		 */
		customDash?: Array<{ dash: number, space: number }>
		/**
		 * Arrow head size (`a:headEnd@w` / `@len`)
		 * - `beginArrowType` selects the shape; these size it
		 */
		beginArrowSize?: { width?: 'sm' | 'med' | 'lg', length?: 'sm' | 'med' | 'lg' }
		/**
		 * Arrow tail size (`a:tailEnd@w` / `@len`)
		 */
		endArrowSize?: { width?: 'sm' | 'med' | 'lg', length?: 'sm' | 'med' | 'lg' }

		/**
		 * Dash type
		 * @deprecated v3.3.0 - use `dashType`
		 */
		lineDash?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * @deprecated v3.3.0 - use `beginArrowType`
		 */
		lineHead?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * @deprecated v3.3.0 - use `endArrowType`
		 */
		lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * Line width (pt)
		 * @deprecated v3.3.0 - use `width`
		 */
		pt?: number
		/**
		 * Line size (pt)
		 * @deprecated v3.3.0 - use `width`
		 */
		size?: number
	}
	// used by: chart, slide, table, text
	export interface TextBaseProps {
		/**
		 * Horizontal alignment
		 * @default 'left'
		 */
		align?: HAlign
		/**
		 * Bold style
		 * @default false
		 */
		bold?: boolean
		/**
		 * Add a line-break
		 * @default false
		 */
		breakLine?: boolean
		/**
		 * Add standard or custom bullet
		 * - use `true` for standard bullet
		 * - pass object options for custom bullet
		 * @default false
		 */
		bullet?:
		| boolean
		| {
			/**
			 * Bullet type
			 * @default bullet
			 */
			type?: 'bullet' | 'number'
			/**
			 * Bullet colour, independent of the run's text colour (`a:buClr`)
			 */
			color?: Color
			/**
			 * Bullet size as a percent of the text size (`a:buSzPct`)
			 * - range: 25-400
			 * @default 100
			 */
			size?: number
			/**
			 * Bullet size in points (`a:buSzPts`), instead of a percent
			 */
			sizePts?: number
			/**
			 * Typeface for the bullet glyph (`a:buFont`)
			 * - needed for Wingdings-style character bullets
			 */
			fontFace?: string
			/**
			 * Picture bullet (`a:buBlip`) - base64 image data with a mime header
			 * - takes precedence over a character or number bullet
			 */
			image?: string
			/**
			 * Bullet character code (unicode)
			 * @since v3.3.0
			 * @example '25BA' // 'BLACK RIGHT-POINTING POINTER' (U+25BA)
			 */
			characterCode?: string
			/**
			 * Indentation (space between bullet and text) (points)
			 * @since v3.3.0
			 * @default 27 // DEF_BULLET_MARGIN
			 * @example 10 // Indents text 10 points from bullet
			 */
			indent?: number
			/**
			 * Number type
			 * @since v3.3.0
			 * @example 'romanLcParenR' // roman numerals lower-case with paranthesis right
			 */
			numberType?:
			| 'alphaLcParenBoth'
			| 'alphaLcParenR'
			| 'alphaLcPeriod'
			| 'alphaUcParenBoth'
			| 'alphaUcParenR'
			| 'alphaUcPeriod'
			| 'arabicParenBoth'
			| 'arabicParenR'
			| 'arabicPeriod'
			| 'arabicPlain'
			| 'romanLcParenBoth'
			| 'romanLcParenR'
			| 'romanLcPeriod'
			| 'romanUcParenBoth'
			| 'romanUcParenR'
			| 'romanUcPeriod'
			/**
			 * Number bullets start at
			 * @since v3.3.0
			 * @default 1
			 * @example 10 // numbered bullets start with 10
			 */
			numberStartAt?: number

			// DEPRECATED

			/**
			 * Bullet code (unicode)
			 * @deprecated v3.3.0 - use `characterCode`
			 */
			code?: string
			/**
			 * Margin between bullet and text
			 * @since v3.2.1
			 * @deplrecated v3.3.0 - use `indent`
			 */
			marginPt?: number
			/**
			 * Number to start with (only applies to type:number)
			 * @deprecated v3.3.0 - use `numberStartAt`
			 */
			startAt?: number
			/**
			 * Number type
			 * @deprecated v3.3.0 - use `numberType`
			 */
			style?: string
		}
		/**
		 * Text color
		 * - `HexColor` or `ThemeColor`
		 * - MS-PPT > Format Shape > Text Options > Text Fill & Outline > Text Fill > Color
		 * @example 'FF0000' // hex color (red)
		 * @example pptx.SchemeColor.text1 // Theme color (Text1)
		 */
		color?: Color
		/**
		 * Font face name
		 * @example 'Arial' // Arial font
		 */
		fontFace?: string
		/**
		 * Font size
		 * @example 12 // Font size 12
		 */
		fontSize?: number
		/**
		 * Text highlight color (hex format)
		 * @example 'FFFF00' // yellow
		 */
		highlight?: HexColor
		/**
		 * italic style
		 * @default false
		 */
		italic?: boolean
		/**
		 * language
		 * - ISO 639-1 standard language code
		 * @default 'en-US' // english US
		 * @example 'fr-CA' // french Canadian
		 */
		lang?: string
		/**
		 * Add a soft line-break (shift+enter) before line text content
		 * @default false
		 * @since v3.5.0
		 */
		softBreakBefore?: boolean
		/**
		 * tab stops
		 * - PowerPoint: Paragraph > Tabs > Tab stop position
		 * @example [{ position:1 }, { position:3 }] // Set first tab stop to 1 inch, set second tab stop to 3 inches
		 */
		tabStops?: Array<{ position: number, alignment?: 'l' | 'r' | 'ctr' | 'dec' }>
		/**
		 * text direction
		 * `horz` = horizontal
		 * `vert` = rotate 90^
		 * `vert270` = rotate 270^
		 * `wordArtVert` = stacked
		 * @default 'horz'
		 */
		textDirection?: 'horz' | 'vert' | 'vert270' | 'wordArtVert'
		/**
		 * Transparency (percent)
		 * - MS-PPT > Format Shape > Text Options > Text Fill & Outline > Text Fill > Transparency
		 * - range: 0-100
		 * @default 0
		 */
		transparency?: number
		/**
		 * underline properties
		 * - PowerPoint: Font > Color & Underline > Underline Style/Underline Color
		 * @default (none)
		 */
		underline?: {
			style?:
			| 'dash'
			| 'dashHeavy'
			| 'dashLong'
			| 'dashLongHeavy'
			| 'dbl'
			| 'dotDash'
			| 'dotDashHeave'
			| 'dotDotDash'
			| 'dotDotDashHeavy'
			| 'dotted'
			| 'dottedHeavy'
			| 'heavy'
			| 'none'
			| 'sng'
			| 'wavy'
			| 'wavyDbl'
			| 'wavyHeavy'
			color?: Color
		}
		/**
		 * vertical alignment
		 * @default 'top'
		 */
		valign?: VAlign
	}
	export interface PlaceholderProps extends PositionProps, TextBaseProps {
		name: string
		type: PLACEHOLDER_TYPE
		/**
		 * margin (points)
		 */
		margin?: Margin
		/**
		 * Text direction inside the placeholder (`p:ph@orient`)
		 * @default horz
		 */
		orient?: 'horz' | 'vert'
		/**
		 * How much of the layout the placeholder covers (`p:ph@sz`)
		 * @default full
		 */
		sz?: 'full' | 'half' | 'quarter'
		/**
		 * Mark the placeholder as drawn by the author rather than inherited layout furniture
		 * (`p:nvPr@userDrawn`)
		 * @default false
		 */
		userDrawn?: boolean
	}
	/**
	 * Editing locks (ECMA-376 20.1.2.2.34 `a:spLocks` and its siblings)
	 * - each is omitted unless set, so the values the library emits today are unchanged
	 */
	export interface ShapeLockProps {
		/** prevent grouping with other shapes */
		noGroup?: boolean
		/** prevent selection */
		noSelect?: boolean
		/** prevent rotation */
		noRotate?: boolean
		/** prevent changing the aspect ratio */
		noChangeAspect?: boolean
		/** prevent moving */
		noMove?: boolean
		/** prevent resizing */
		noResize?: boolean
		/** prevent editing the geometry points */
		noEditPoints?: boolean
		/** prevent dragging the adjust handles */
		noAdjustHandles?: boolean
		/** prevent changing arrowheads */
		noChangeArrowheads?: boolean
		/** prevent changing the preset geometry */
		noChangeShapeType?: boolean
		/** prevent editing the text */
		noTextEdit?: boolean
		/** pictures only: prevent cropping */
		noCrop?: boolean
		/** pictures only: prefer resizing relative to the original size */
		preferRelativeResize?: boolean
	}
	/**
	 * Non-visual drawing properties beyond the name (ECMA-376 19.3.1.12 `p:cNvPr`)
	 */
	export interface NonVisualProps {
		/**
		 * Alt-text *title*, distinct from the description
		 * - PowerPoint: right-click > Edit Alt Text
		 */
		title?: string
		/**
		 * Hide the shape
		 * - it stays in the file and can be re-shown from the selection pane
		 * @default false
		 */
		hidden?: boolean
		/**
		 * Editing locks for this shape
		 */
		lock?: ShapeLockProps
	}
	export interface ObjectNameProps extends NonVisualProps {
		/**
		 * Object name
		 * - used instead of default "Object N" name
		 * - PowerPoint: Home > Arrange > Selection Pane...
		 * @since v3.10.0
		 * @default 'Object 1'
		 * @example 'Antenna Design 9'
		 */
		objectName?: string
	}
	export interface ThemeProps {
		/**
		 * Headings font face name
		 * @example 'Arial Narrow'
		 * @default 'Calibri Light'
		 */
		headFontFace?: string
		/**
		 * Body font face name
		 * @example 'Arial'
		 * @default 'Calibri'
		 */
		bodyFontFace?: string
	}

	// image / media ==================================================================================
	export type MediaType = 'audio' | 'online' | 'video' | 'audioCd' | 'wav'

	/**
	 * A point on an audio CD (`a:st`/`a:end`, ECMA-376 Part 1 §20.1.3.4 CT_AudioCDTime)
	 */
	export interface AudioCdTimeProps {
		/** CD track number (0-255) - required by the schema */
		track: number
		/**
		 * Offset into the track, in seconds
		 * @default 0
		 */
		time?: number
	}

	/**
	 * CD audio source (`a:audioCd`, ECMA-376 Part 1 §20.1.3.3 CT_AudioCD)
	 * - references the listener's CD drive, so it embeds nothing and needs no relationship
	 */
	export interface AudioCdProps {
		start: AudioCdTimeProps
		end: AudioCdTimeProps
	}

	export interface ImageProps extends PositionProps, DataOrPathProps, ObjectNameProps {
		/**
		 * Glow effect
		 */
		glow?: TextGlowProps
		/**
		 * Soft-edge effect
		 */
		softEdge?: SoftEdgeProps
		/**
		 * Reflection effect
		 */
		reflection?: ReflectionProps
		/**
		 * Blur options (`a:blur`)
		 */
		blur?: BlurProps
		/**
		 * Fill blended over the object's own fill (`a:fillOverlay`)
		 * @example { blend: 'mult', fill: { color: 'FF0000', transparency: 50 } }
		 */
		fillOverlay?: FillOverlayProps
		/**
		 * Emit the effects as a composed effect graph (`a:effectDag`) rather than an `a:effectLst`
		 * - the two are alternatives in the schema, so this replaces the effect list
		 */
		effectDag?: EffectDagProps
		/**
		 * Alpha (transparency) effects applied to the image itself
		 * @example { invert: true }
		 */
		alphaEffects?: ImageAlphaEffectProps
		/**
		 * Recolour and correction effects applied to the image itself
		 * @example { grayscale: true, brightness: 20, contrast: -10 }
		 */
		recolor?: ImageRecolorProps
		/**
		 * Theme style references (`p:style`)
		 * - lets the object restyle when the user swaps the presentation theme in PowerPoint
		 * @example { fill: 1, line: 2, effect: 0, font: 'minor' }
		 */
		styleRef?: ShapeStyleProps
		/**
		 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
		 * - PowerPoint: [right-click on an image] > "Edit Alt Text..."
		 */
		altText?: string
		/**
		 * Flip horizontally?
		 * @default false
		 */
		flipH?: boolean
		/**
		 * Flip vertical?
		 * @default false
		 */
		flipV?: boolean
		hyperlink?: HyperlinkProps
		/**
		 * Mouse-over action, configured like `hyperlink` but triggered on hover
		 * - PowerPoint's Insert > Action > Mouse Over tab
		 * @example { hyperlinkHover: { slide: 3, tooltip: 'Jump to results' } }
		 */
		hyperlinkHover?: HyperlinkProps
		/**
		 * Image outline/border (a picture frame)
		 * @example { color: '696969', width: 2 } // 2pt dim-gray border
		 */
		line?: ShapeLineProps
		/**
		 * Placeholder type
		 * - values: 'body' | 'header' | 'footer' | 'title' | et. al.
		 * @example 'body'
		 * @see https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
		 */
		placeholder?: string
		/**
		 * Image rotation (degrees)
		 * - range: -360 to 360
		 * @default 0
		 * @example 180 // rotate image 180 degrees
		 */
		rotate?: number
		/**
		 * Enable image rounding
		 * @default false
		 */
		rounding?: boolean
		/**
		 * Shadow Props
		 * - MS-PPT > Format Picture > Shadow
		 * @example
		 * { type: 'outer', color: '000000', opacity: 0.5, blur: 20,  offset: 20, angle: 270 }
		 */
		shadow?: ShadowProps
		/**
		 * Image sizing options
		 */
		sizing?: {
			/**
			 * Sizing type
			 */
			type: 'contain' | 'cover' | 'crop'
			/**
			 * Image width
			 * - inches or percentage
			 * @example 10.25 // position in inches
			 * @example '75%' // position as percentage of slide size
			 */
			w: Coord
			/**
			 * Image height
			 * - inches or percentage
			 * @example 10.25 // position in inches
			 * @example '75%' // position as percentage of slide size
			 */
			h: Coord
			/**
			 * Offset from left to crop image
			 * - `crop` only
			 * - inches or percentage
			 * @example 10.25 // position in inches
			 * @example '75%' // position as percentage of slide size
			 */
			x?: Coord
			/**
			 * Offset from top to crop image
			 * - `crop` only
			 * - inches or percentage
			 * @example 10.25 // position in inches
			 * @example '75%' // position as percentage of slide size
			 */
			y?: Coord
		}
		/**
		 * Transparency (percent)
		 * - MS-PPT > Format Picture > Picture > Picture Transparency > Transparency
		 * - range: 0-100
		 * @default 0
		 * @example 25 // 25% transparent
		 */
		transparency?: number
		/**
		 * Animation applied to this object
		 * @example { animation: { type: 'fadeIn', trigger: 'afterPrevious', duration: 800 } }
		 */
		animation?: AnimationProps
	}
	/**
	 * Add media (audio/video) to slide
	 * @requires either `link` or `path`
	 */
	export interface MediaProps extends PositionProps, DataOrPathProps, ObjectNameProps {
		/**
		 * Media type
		 * - Use 'online' to embed a YouTube video (only supported in recent versions of PowerPoint)
		 */
		type: MediaType
		/**
		 * Cover image
		 * @since 3.9.0
		 * @default "play button" image, gray background
		 */
		cover?: string
		/**
		 * media file extension
		 * - use when the media file path does not already have an extension, ex: "/folder/SomeSong"
		 * @since 3.9.0
		 * @default extension from file provided
		 */
		extn?: string
		/**
		 * video embed link
		 * - works with YouTube
		 * - other sites may not show correctly in PowerPoint
		 * @example 'https://www.youtube.com/embed/Dph6ynRVyUc' // embed a youtube video
		 */
		link?: string
		/**
		 * full or local path
		 * @example 'https://freesounds/simpsons/bart.mp3' // embed mp3 audio clip from server
		 * @example '/sounds/simpsons_haha.mp3' // embed mp3 audio clip from local directory
		 */
		path?: string
		/**
		 * Start playing as soon as the slide is shown, without a click
		 * - emitted as a media node in the slide timing tree with a `delay="0"` start condition
		 *   (ECMA-376 19.5.30 `CT_TLCommonMediaNodeData`)
		 * @default false (plays on click)
		 */
		autoplay?: boolean
		/**
		 * Repeat until the slide advances ("Loop until Stopped" in PowerPoint's Playback tab)
		 * - sets `repeatCount="indefinite"` on the media node
		 * @default false
		 */
		loop?: boolean
		/**
		 * Play video full-screen (`p:video@fullScrn`)
		 * - video only; ignored with a warning for `type: 'audio'`
		 * @default false
		 */
		fullScreen?: boolean
		/**
		 * Mute the media's audio (`p:cMediaNode@mute`)
		 * @default false
		 */
		mute?: boolean
		/**
		 * Animation applied to this object
		 * @example { animation: { type: 'fadeIn', trigger: 'afterPrevious', duration: 800 } }
		 */
		animation?: AnimationProps
		/**
		 * MIME type of the referenced media (`a:audioFile@contentType`, `a:videoFile@contentType`)
		 * @example 'video/mp4'
		 */
		contentType?: string
		/**
		 * CD audio track range - required when `type` is `audioCd`
		 * @example { start: { track: 1 }, end: { track: 1, time: 30 } }
		 */
		audioCd?: AudioCdProps
		/**
		 * Mark the media frame as a photo (`p:nvPr@isPhoto`)
		 * @default false
		 */
		isPhoto?: boolean
		/**
		 * Mark the media frame as author-placed rather than layout furniture (`p:nvPr@userDrawn`)
		 * @default false
		 */
		userDrawn?: boolean
	}

	// shapes =========================================================================================

	export interface ShapeProps extends PositionProps, ObjectNameProps, TextBodyProps {
		/**
		 * Glow effect
		 */
		glow?: TextGlowProps
		/**
		 * Soft-edge effect
		 */
		softEdge?: SoftEdgeProps
		/**
		 * Reflection effect
		 */
		reflection?: ReflectionProps
		/**
		 * Blur options (`a:blur`)
		 */
		blur?: BlurProps
		/**
		 * Fill blended over the object's own fill (`a:fillOverlay`)
		 * @example { blend: 'mult', fill: { color: 'FF0000', transparency: 50 } }
		 */
		fillOverlay?: FillOverlayProps
		/**
		 * Emit the effects as a composed effect graph (`a:effectDag`) rather than an `a:effectLst`
		 * - the two are alternatives in the schema, so this replaces the effect list
		 */
		effectDag?: EffectDagProps
		/**
		 * Theme style references (`p:style`)
		 * - lets the object restyle when the user swaps the presentation theme in PowerPoint
		 * @example { fill: 1, line: 2, effect: 0, font: 'minor' }
		 */
		styleRef?: ShapeStyleProps
		/**
		 * Horizontal alignment
		 * @default 'left'
		 */
		align?: HAlign
		/**
		 * Radius (only for pptx.shapes.PIE, pptx.shapes.ARC, pptx.shapes.BLOCK_ARC)
		 * - In the case of pptx.shapes.BLOCK_ARC you have to setup the arcThicknessRatio
		 * - values: [0-359, 0-359]
		 * @since v3.4.0
		 * @default [270, 0]
		 */
		angleRange?: [number, number]
		/**
		 * Radius (only for pptx.shapes.BLOCK_ARC)
		 * - You have to setup the angleRange values too
		 * - values: 0.0-1.0
		 * @since v3.4.0
		 * @default 0.5
		 */
		arcThicknessRatio?: number
		/**
		 * Shape fill color properties
		 * @example { color:'FF0000' } // hex color (red)
		 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
		 * @example { color:pptx.SchemeColor.accent1 } // Theme color Accent1
		 */
		fill?: ShapeFillProps
		/**
		 * Flip shape horizontally?
		 * @default false
		 */
		flipH?: boolean
		/**
		 * Flip shape vertical?
		 * @default false
		 */
		flipV?: boolean
		/**
		 * Add hyperlink to shape
		 * @example hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" },
		 */
		hyperlink?: HyperlinkProps
		/**
		 * Mouse-over action, configured like `hyperlink` but triggered on hover
		 * - PowerPoint's Insert > Action > Mouse Over tab
		 * @example { hyperlinkHover: { slide: 3, tooltip: 'Jump to results' } }
		 */
		hyperlinkHover?: HyperlinkProps
		/**
		 * Line options
		 */
		line?: ShapeLineProps
		/**
		 * Points (only for pptx.shapes.CUSTOM_GEOMETRY)
		 * - type: 'arc'
		 * - `hR` Shape Arc Height Radius
		 * - `wR` Shape Arc Width Radius
		 * - `stAng` Shape Arc Start Angle
		 * - `swAng` Shape Arc Swing Angle
		 * @see http://www.datypic.com/sc/ooxml/e-a_arcTo-1.html
		 * @example [{ x: 0, y: 0 }, { x: 10, y: 10 }] // draw a line between those two points
		 */
		points?: Array<
			| { x: Coord, y: Coord, moveTo?: boolean }
			| { x: Coord, y: Coord, curve: { type: 'arc', hR: Coord, wR: Coord, stAng: number, swAng: number } }
			| { x: Coord, y: Coord, curve: { type: 'cubic', x1: Coord, y1: Coord, x2: Coord, y2: Coord } }
			| { x: Coord, y: Coord, curve: { type: 'quadratic', x1: Coord, y1: Coord } }
			| { close: true }
		>
		/**
		 * Rounded rectangle radius (only for pptx.shapes.ROUNDED_RECTANGLE)
		 * - values: 0.0 to 1.0
		 * @default 0
		 */
		rectRadius?: number
		/**
		 * Rotation (degrees)
		 * - range: -360 to 360
		 * @default 0
		 * @example 180 // rotate 180 degrees
		 */
		rotate?: number
		/**
		 * Shadow options
		 * TODO: need new demo.js entry for shape shadow
		 */
		shadow?: ShadowProps

		/**
		 * @deprecated v3.3.0
		 */
		lineSize?: number
		/**
		 * @deprecated v3.3.0
		 */
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		/**
		 * @deprecated v3.3.0
		 */
		lineHead?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
		/**
		 * @deprecated v3.3.0
		 */
		lineTail?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
		/**
		 * Shape name (used instead of default "Shape N" name)
		 * @deprecated v3.10.0 - use `objectName`
		 */
		shapeName?: string
		/**
		 * Animation applied to this object
		 * @example { animation: { type: 'fadeIn', trigger: 'afterPrevious', duration: 800 } }
		 */
		animation?: AnimationProps
	}

	// tables =========================================================================================

	export interface TableToSlidesProps extends TableProps {
		//_arrObjTabHeadRows?: TableRow[]
		// _masterSlide?: SlideLayout

		/**
		 * Add an image to slide(s) created during autopaging
		 * - `image` prop requires either `path` or `data`
		 * - see `DataOrPathProps` for details on `image` props
		 * - see `PositionProps` for details on `options` props
		 */
		addImage?: { image: DataOrPathProps, options: PositionProps }
		/**
		 * Add a shape to slide(s) created during autopaging
		 */
		addShape?: { shapeName: SHAPE_NAME, options: ShapeProps }
		/**
		 * Add a table to slide(s) created during autopaging
		 */
		addTable?: { rows: TableRow[], options: TableProps }
		/**
		 * Add a text object to slide(s) created during autopaging
		 */
		addText?: { text: TextProps[], options: TextPropsOptions }
		/**
		 * Whether to enable auto-paging
		 * - auto-paging creates new slides as content overflows a slide
		 * @deprecated vNEXT - estimates where text wraps instead of measuring it; removed in v5.0. Migrate to `paginateTable` from @neo-ma/pptxgenjs-std/tables (see DEPRECATION-PLAN.md F10)
		 * @default true
		 */
		autoPage?: boolean
		/**
		 * Auto-paging character weight
		 * - adjusts how many characters are used before lines wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
		 * @deprecated vNEXT - tuning knob for the built-in wrap estimate; removed in v5.0 with the auto-paging engine (see DEPRECATION-PLAN.md F10)
		 */
		autoPageCharWeight?: number
		/**
		 * Auto-paging line weight
		 * - adjusts how many lines are used before slides wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
		 * @deprecated vNEXT - tuning knob for the built-in line-height estimate; removed in v5.0 with the auto-paging engine (see DEPRECATION-PLAN.md F10)
		 */
		autoPageLineWeight?: number
		/**
		 * Whether to repeat head row(s) on new tables created by autopaging
		 * @since v3.3.0
		 * @default false
		 */
		autoPageRepeatHeader?: boolean
		/**
		 * The `y` location to use on subsequent slides created by autopaging
		 * @default (top margin of Slide)
		 */
		autoPageSlideStartY?: number
		/**
		 * Column widths (inches)
		 */
		colW?: number | number[]
		/**
		 * Master slide name
		 * - define a master slide to have your auto-paged slides have corporate design, etc.
		 * @see https://gitbrent.github.io/PptxGenJS/docs/masters.html
		 */
		masterSlideName?: string
		/**
		 * Slide margin
		 * - this margin will be across all slides created by auto-paging
		 */
		slideMargin?: Margin

		/**
		 * @deprecated v3.3.0 - use `autoPageRepeatHeader`
		 */
		addHeaderToEach?: boolean
		/**
		 * @deprecated v3.3.0 - use `autoPageSlideStartY`
		 */
		newSlideStartY?: number
	}
	/**
	 * 3-D bevel applied to a table cell (`a:bevel`, ECMA-376 Part 1 §20.1.5.3 CT_Bevel)
	 */
	export interface CellBevelProps {
		/**
		 * Bevel shape
		 * @default circle
		 */
		preset?: 'relaxedInset' | 'circle' | 'slope' | 'cross' | 'angle' | 'softRound' | 'convex' | 'coolSlant' | 'divot' | 'riblet' | 'hardEdge' | 'artDeco'
		/**
		 * Bevel width (inches)
		 * @default 0.083
		 */
		width?: number
		/**
		 * Bevel height (inches)
		 * @default 0.083
		 */
		height?: number
	}

	/**
	 * Light rig for a 3-D table cell (`a:lightRig`, ECMA-376 Part 1 §20.1.5.5 CT_LightRig)
	 * - both properties are required by the schema, so a partial value is not emitted
	 */
	export interface CellLightRigProps {
		rig: 'legacyFlat1' | 'legacyFlat2' | 'legacyFlat3' | 'legacyFlat4' | 'legacyNormal1' | 'legacyNormal2' | 'legacyNormal3' | 'legacyNormal4' | 'legacyHarsh1' | 'legacyHarsh2' | 'legacyHarsh3' | 'legacyHarsh4' | 'threePt' | 'balanced' | 'soft' | 'harsh' | 'flood' | 'contrasting' | 'morning' | 'sunrise' | 'sunset' | 'chilly' | 'freezing' | 'flat' | 'twoPt' | 'glow' | 'brightRoom'
		dir: 'tl' | 't' | 'tr' | 'l' | 'r' | 'bl' | 'b' | 'br'
	}

	/**
	 * 3-D properties of a table cell (`a:cell3D`, ECMA-376 Part 1 §21.1.3.1 CT_Cell3D)
	 */
	export interface Cell3DProps {
		/**
		 * Cell bevel
		 * - `a:bevel` is required by the schema, so an `a:bevel` with schema defaults is written
		 *   whenever `cell3D` is set
		 */
		bevel?: CellBevelProps
		/**
		 * Surface material
		 * @default plastic
		 */
		material?: 'legacyMatte' | 'legacyPlastic' | 'legacyMetal' | 'legacyWireframe' | 'matte' | 'plastic' | 'metal' | 'warmMatte' | 'translucentPowder' | 'powder' | 'dkEdge' | 'softEdge' | 'clear' | 'flat' | 'softmetal'
		/** Light rig - dropped unless both `rig` and `dir` are set */
		lightRig?: CellLightRigProps
	}

	export interface TableCellProps extends TextBaseProps {
		/**
		 * Auto-paging character weight
		 * - adjusts how many characters are used before lines wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
		 * @deprecated vNEXT - tuning knob for the built-in wrap estimate; removed in v5.0 with the auto-paging engine (see DEPRECATION-PLAN.md F10)
		 */
		autoPageCharWeight?: number
		/**
		 * Auto-paging line weight
		 * - adjusts how many lines are used before slides wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
		 * @deprecated vNEXT - tuning knob for the built-in line-height estimate; removed in v5.0 with the auto-paging engine (see DEPRECATION-PLAN.md F10)
		 */
		autoPageLineWeight?: number
		/**
		 * Whether text is centered both horizontally and vertically in the cell
		 * @default false
		 */
		anchorCtr?: boolean
		/**
		 * Cell border
		 */
		border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
		/**
		 * Diagonal border running bottom-left to top-right (`a:lnBlToTr`)
		 * - independent of `border`, which only covers the four edges
		 */
		borderDiagonalUp?: BorderProps
		/**
		 * Diagonal border running top-left to bottom-right (`a:lnTlToBr`)
		 * - independent of `border`, which only covers the four edges
		 */
		borderDiagonalDown?: BorderProps
		/**
		 * 3-D bevel and lighting for the cell
		 * @example { bevel: { preset: 'circle', width: 0.05, height: 0.05 }, material: 'metal' }
		 */
		cell3D?: Cell3DProps
		/**
		 * Cell colspan
		 */
		colspan?: number
		/**
		 * Whether text wider than the cell is clipped or allowed to overflow it
		 * @default clip
		 */
		horzOverflow?: 'clip' | 'overflow'
		/**
		 * Fill color
		 * @example { color:'FF0000' } // hex color (red)
		 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
		 * @example { color:pptx.SchemeColor.accent1 } // theme color Accent1
		 */
		fill?: ShapeFillProps
		hyperlink?: HyperlinkProps
		/**
		 * Mouse-over action, configured like `hyperlink` but triggered on hover
		 * - PowerPoint's Insert > Action > Mouse Over tab
		 * @example { hyperlinkHover: { slide: 3, tooltip: 'Jump to results' } }
		 */
		hyperlinkHover?: HyperlinkProps
		/**
		 * Cell margin (inches)
		 * @default 0
		 */
		margin?: Margin
		/**
		 * Cell rowspan
		 */
		rowspan?: number
	}
	export interface TableProps extends PositionProps, TextBaseProps, ObjectNameProps {
		//_arrObjTabHeadRows?: TableRow[]

		/**
		 * Whether to enable auto-paging
		 * - auto-paging creates new slides as content overflows a slide
		 * @deprecated vNEXT - estimates where text wraps instead of measuring it; removed in v5.0. Migrate to `paginateTable` from @neo-ma/pptxgenjs-std/tables (see DEPRECATION-PLAN.md F10)
		 * @default false
		 */
		autoPage?: boolean
		/**
		 * Auto-paging character weight
		 * - adjusts how many characters are used before lines wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
		 * @deprecated vNEXT - tuning knob for the built-in wrap estimate; removed in v5.0 with the auto-paging engine (see DEPRECATION-PLAN.md F10)
		 */
		autoPageCharWeight?: number
		/**
		 * Auto-paging line weight
		 * - adjusts how many lines are used before slides wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
		 * @deprecated vNEXT - tuning knob for the built-in line-height estimate; removed in v5.0 with the auto-paging engine (see DEPRECATION-PLAN.md F10)
		 */
		autoPageLineWeight?: number
		/**
		 * Whether table header row(s) should be repeated on each new slide creating by autoPage.
		 * Use `autoPageHeaderRows` to designate how many rows comprise the table header (1+).
		 * @default false
		 * @since v3.3.0
		 */
		autoPageRepeatHeader?: boolean
		/**
		 * Number of rows that comprise table headers
		 * - required when `autoPageRepeatHeader` is set to true.
		 * @example 2 - repeats the first two table rows on each new slide created
		 * @default 1
		 * @since v3.3.0
		 */
		autoPageHeaderRows?: number
		/**
		 * The `y` location to use on subsequent slides created by autopaging
		 * @default (top margin of Slide)
		 */
		autoPageSlideStartY?: number
		/**
		 * Table border
		 * - single value is applied to all 4 sides
		 * - array of values in TRBL order for individual sides
		 */
		border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
		/**
		 * Width of table columns (inches)
		 * - single value is applied to every column equally based upon `w`
		 * - array of values in applied to each column in order
		 * @default columns of equal width based upon `w`
		 */
		colW?: number | number[]
		/**
		 * Cell background color
		 * @example { color:'FF0000' } // hex color (red)
		 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
		 * @example { color:pptx.SchemeColor.accent1 } // theme color Accent1
		 */
		fill?: ShapeFillProps
		/**
		 * Cell margin (inches)
		 * - affects all table cells, is superceded by cell options
		 */
		margin?: Margin
		/**
		 * Height of table rows (inches)
		 * - single value is applied to every row equally based upon `h`
		 * - array of values in applied to each row in order
		 * @default rows of equal height based upon `h`
		 */
		rowH?: number | number[]
		/**
		 * Apply special formatting to the first row (header emphasis)
		 * - only renders when a table style is in effect (see `tableStyleId`)
		 * @default false
		 */
		firstRow?: boolean
		/**
		 * Apply special formatting to the last row (totals emphasis)
		 * @default false
		 */
		lastRow?: boolean
		/**
		 * Apply special formatting to the first column
		 * @default false
		 */
		firstCol?: boolean
		/**
		 * Apply special formatting to the last column
		 * @default false
		 */
		lastCol?: boolean
		/**
		 * Band (alternate the fill of) the rows
		 * @default false
		 */
		bandRow?: boolean
		/**
		 * Band (alternate the fill of) the columns
		 * @default false
		 */
		bandCol?: boolean
		/**
		 * Lay the columns out right-to-left (`a:tblPr@rtl`)
		 * - reverses column order for the whole table; cell text direction is set per cell
		 * @default false
		 */
		rtl?: boolean
		/**
		 * Table style id (GUID of a built-in PowerPoint table style)
		 * - required for `bandRow`/`firstRow`/etc. to have a visible effect
		 * @example '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' // "Medium Style 2 - Accent 1"
		 */
		tableStyleId?: string
		/**
		 * DEV TOOL: Verbose Mode (to console)
		 * - tell the library to provide an almost ridiculous amount of detail during auto-paging calculations
		 * @default false // obviously
		 */
		verbose?: boolean // Undocumented; shows verbose output

		/**
		 * @deprecated v3.3.0 - use `autoPageSlideStartY`
		 */
		newSlideStartY?: number
		/**
		 * Animation applied to this object
		 * @example { animation: { type: 'fadeIn', trigger: 'afterPrevious', duration: 800 } }
		 */
		animation?: AnimationProps
	}
	export interface TableCell {
		text?: string | TableCell[]
		options?: TableCellProps
	}
	export interface TableRowSlide {
		rows: TableRow[]
	}
	export type TableRow = TableCell[]

	// text ===========================================================================================
	/**
	 * Soft-edge effect (`a:softEdge`).
	 * - MS-PPT > Format Shape/Picture > Effects > Soft Edges
	 */
	/**
	 * Blur effect (`a:blur`, ECMA-376 Part 1 §20.1.8.15 CT_BlurEffect)
	 */
	export interface BlurProps {
		/** Blur radius (points) */
		radius: number
		/**
		 * Whether the blur may grow beyond the shape's bounds
		 * @default true
		 */
		grow?: boolean
	}

	/**
	 * Fill-overlay effect (`a:fillOverlay`, ECMA-376 Part 1 §20.1.8.29 CT_FillOverlayEffect)
	 * - blends a second fill over the shape's own fill
	 * - both properties are required by the schema, so a partial value is not emitted
	 */
	export interface FillOverlayProps {
		/** Blend mode */
		blend: 'over' | 'mult' | 'screen' | 'darken' | 'lighten'
		/** The fill blended over the shape */
		fill: ShapeFillProps
	}

	/**
	 * Composed effect graph (`a:effectDag`, ECMA-376 Part 1 §20.1.8.25 CT_EffectContainer)
	 * - `a:effectLst` and `a:effectDag` are alternatives in the schema, so setting this emits the
	 *   shape's effects inside `a:effectDag` instead of `a:effectLst`
	 */
	export interface EffectDagProps {
		/**
		 * Whether the contained effects apply to siblings or to the whole tree
		 * @default sib
		 */
		type?: 'sib' | 'tree'
	}

	/**
	 * Recolour and correction effects applied to an image's `a:blip`
	 * - PowerPoint's Picture Format > Color and > Corrections galleries
	 */
	export interface ImageRecolorProps {
		/**
		 * Map the image onto two colours (`a:duotone`)
		 * - the schema requires exactly two, so anything else is dropped
		 * @example ['000000', 'FFFFFF']
		 */
		duotone?: [Color, Color]
		/** Convert to greyscale (`a:grayscl`) */
		grayscale?: boolean
		/**
		 * Brightness, percent -100 to 100 (`a:lum@bright`)
		 * @default 0
		 */
		brightness?: number
		/**
		 * Contrast, percent -100 to 100 (`a:lum@contrast`)
		 * @default 0
		 */
		contrast?: number
		/**
		 * Reduce to two tones at this threshold, percent 0 to 100 (`a:biLevel@thresh`)
		 * - required by the schema, so this is what turns the effect on
		 */
		blackWhiteThreshold?: number
		/**
		 * Replace one colour with another (`a:clrChange`) - PowerPoint's Set Transparent Color
		 * - both colours are required by the schema
		 */
		colorChange?: {
			from: Color
			to: Color
			/**
			 * Whether the alpha channel takes part in the comparison (`@useA`)
			 * @default true
			 */
			useAlpha?: boolean
		}
	}

	/**
	 * Alpha (transparency) effects applied to an image's `a:blip`
	 * - these are image effects, not shape effects: the schema allows them on `a:blip` and inside
	 *   `a:effectDag`, but not in `a:effectLst`
	 */
	export interface ImageAlphaEffectProps {
		/**
		 * Replace every alpha value with this one (`a:alphaRepl`)
		 * - percent, 0 (fully transparent) to 100 (fully opaque)
		 */
		replace?: number
		/** Invert the alpha channel (`a:alphaInv`) */
		invert?: boolean
		/** Force alpha values below 100% to fully transparent (`a:alphaFloor`) */
		floor?: boolean
		/** Force alpha values above 0% to fully opaque (`a:alphaCeiling`) */
		ceiling?: boolean
	}

	/**
	 * Theme style references for a shape (`p:style`, ECMA-376 Part 1 §20.1.4.1.25 CT_ShapeStyle)
	 * - each index points into the matching list in the theme's `a:fmtScheme`, 1-based; omit a
	 *   property to reference nothing, which is what an unstyled shape does today
	 * - a referenced property is the one that changes when the user swaps the theme in PowerPoint, so
	 *   setting `fill` here and no explicit `fill` on the shape leaves the fill to the theme
	 */
	export interface ShapeStyleProps {
		/** Index into the theme's `a:lnStyleLst` (1-based) */
		line?: number
		/** Index into the theme's `a:fillStyleLst` (1-based) */
		fill?: number
		/** Index into the theme's `a:effectStyleLst` (1-based) */
		effect?: number
		/**
		 * Which of the theme's font collections the shape's text follows
		 * @default none
		 */
		font?: 'major' | 'minor' | 'none'
		/**
		 * Colour the referenced line, fill and effect resolve the theme's `phClr` against
		 * @default 'accent1'
		 */
		color?: Color
		/**
		 * Colour the referenced font resolves against
		 * - defaults to `lt1` because a font sharing the fill's colour renders as invisible text
		 * @default 'lt1'
		 */
		fontColor?: Color
	}

	/**
	 * Chart colour style (`cs:colorStyle` in `ppt/charts/colorsN.xml`)
	 * - the palette PowerPoint's Change Colors gallery offers for the chart
	 */
	export interface ChartColorStyleProps {
		/**
		 * How the palette is walked
		 * @default cycle
		 */
		method?: 'cycle' | 'withinLinear' | 'acrossLinear' | 'withinLinearReversed' | 'acrossLinearReversed'
		/**
		 * Which entry of the Change Colors gallery is selected
		 * @default 10
		 */
		id?: number
		/**
		 * The palette itself
		 * @default the theme's six accent colours
		 */
		colors?: Color[]
	}

	/**
	 * Borders of one part of a table style (`a:tcBdr`, ECMA-376 Part 1 §21.1.3.11)
	 */
	export interface TableStyleBorderProps {
		left?: ShapeLineProps
		right?: ShapeLineProps
		top?: ShapeLineProps
		bottom?: ShapeLineProps
		/** horizontal borders between rows */
		insideH?: ShapeLineProps
		/** vertical borders between columns */
		insideV?: ShapeLineProps
	}

	/**
	 * How one part of a table looks in a table style (`a:wholeTbl`, `a:band1H`, `a:firstRow`, ...)
	 */
	export interface TableStylePartProps {
		/** bold text - omitted leaves it to the theme */
		bold?: boolean
		/** italic text - omitted leaves it to the theme */
		italic?: boolean
		/** text colour */
		color?: Color
		/** cell fill */
		fill?: ShapeFillProps
		/** cell borders */
		borders?: TableStyleBorderProps
	}

	/**
	 * A custom table style, written into `ppt/tableStyles.xml` (ECMA-376 Part 1 §14.2.9)
	 * - reference it from a table with `tableStyleId: '<the same id>'`
	 * - `id` and `name` are both required by the schema
	 */
	export interface TableStyleProps {
		/** Style GUID, braced, e.g. `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` */
		id: string
		/** Name shown in PowerPoint's table-style gallery */
		name: string
		/** applies to every cell */
		wholeTable?: TableStylePartProps
		/** first banded row */
		band1H?: TableStylePartProps
		/** second banded row */
		band2H?: TableStylePartProps
		/** first banded column */
		band1V?: TableStylePartProps
		/** second banded column */
		band2V?: TableStylePartProps
		firstRow?: TableStylePartProps
		lastRow?: TableStylePartProps
		firstCol?: TableStylePartProps
		lastCol?: TableStylePartProps
		/** top-left cell */
		nwCell?: TableStylePartProps
		/** top-right cell */
		neCell?: TableStylePartProps
		/** bottom-left cell */
		swCell?: TableStylePartProps
		/** bottom-right cell */
		seCell?: TableStylePartProps
	}

	export interface SoftEdgeProps {
		/**
		 * Soft-edge radius (points)
		 * @example 4
		 */
		radius: number
	}
	/**
	 * Reflection effect (`a:reflection`).
	 * - MS-PPT > Format Shape/Picture > Effects > Reflection
	 */
	export interface ReflectionProps {
		/**
		 * Reflection blur radius (points)
		 * @default 0
		 */
		blur?: number
		/**
		 * Reflection distance from the shape (points)
		 * @default 0
		 */
		distance?: number
		/**
		 * Reflection direction (degrees)
		 * - range: 0-359
		 * @default 0
		 */
		direction?: number
		/**
		 * Starting opacity
		 * - range: 0.0-1.0
		 * @default 0.5
		 */
		opacity?: number
		/**
		 * Vertical scale multiplier
		 * @default -1
		 */
		scaleY?: number
	}
	export interface TextGlowProps {
		/**
		 * Border color (hex format)
		 * @example 'FF3399'
		 */
		color?: HexColor
		/**
		 * opacity (0.0 - 1.0)
		 * @example 0.5
		 * 50% opaque
		 */
		opacity: number
		/**
		 * size (points)
		 */
		size: number
	}

	/**
	 * Text-body attributes beyond wrap/insets/anchor (`a:bodyPr`, ECMA-376 21.1.2.1.1)
	 * - every one is optional and omitted when unset, so default output does not change
	 */
	export type TextFieldType =
		// ECMA-376 21.1.2.2.4 a:fld@type
		| 'slidenum' | 'datetime' | 'datetimeFigureOut'
		| 'datetime1' | 'datetime2' | 'datetime3' | 'datetime4' | 'datetime5' | 'datetime6' | 'datetime7'
		| 'datetime8' | 'datetime9' | 'datetime10' | 'datetime11' | 'datetime12' | 'datetime13'
		| 'headerfooter' | 'hdr' | 'ftr'
	export interface TextBodyProps {
		/**
		 * Lay columns out right-to-left (`a:bodyPr@rtlCol`)
		 * - defaults to the presentation's `rtlMode`, so an RTL deck flows its columns correctly
		 */
		rtlColumns?: boolean
		/**
		 * Keep text upright when the shape is rotated
		 * @default false
		 */
		upright?: boolean
		/**
		 * Rotate the text body independently of the shape (degrees)
		 * - distinct from the shape's own `rotate`
		 */
		textRotate?: number
		/**
		 * Centre the anchor point as well as the text
		 * @default false
		 */
		anchorCenter?: boolean
		/**
		 * Respect paragraph spacing before the first and after the last line
		 * @default false
		 */
		spaceFirstLastPara?: boolean
		/**
		 * Use legacy line-spacing rules
		 * @default false
		 */
		compatLineSpacing?: boolean
		/**
		 * Force anti-aliasing regardless of size
		 * @default false
		 */
		forceAntiAlias?: boolean
		/**
		 * How text overflowing the body behaves
		 * - `horzOverflow` / `vertOverflow` in the schema
		 */
		horizontalOverflow?: 'overflow' | 'clip'
		verticalOverflow?: 'overflow' | 'ellipsis' | 'clip'
	}
	/**
	 * Paragraph attributes beyond margins/alignment (`a:pPr`, ECMA-376 21.1.2.2.7)
	 */
	export interface ParagraphProps {
		/**
		 * Right margin (inches)
		 * - the left margin is `indentLevel`/`marL` handling already present
		 */
		marginRight?: number
		/**
		 * Default tab stop interval (inches)
		 */
		defaultTabSize?: number
		/**
		 * Baseline alignment of glyphs within the line
		 */
		fontAlign?: 'auto' | 't' | 'ctr' | 'base' | 'b'
		/**
		 * Apply East Asian line-breaking rules
		 * @default true
		 */
		eastAsianLineBreak?: boolean
		/**
		 * Allow Latin words to break across lines
		 * @default true
		 */
		latinLineBreak?: boolean
		/**
		 * Apply hanging punctuation
		 * @default true
		 */
		hangingPunctuation?: boolean
	}
	/**
	 * Run attributes beyond bold/italic/size (`a:rPr`, ECMA-376 21.1.2.3.9)
	 */
	export interface TextRunProps {
		/**
		 * Capitalisation applied to the run
		 */
		capitalization?: 'none' | 'small' | 'all'
		/**
		 * Normalise glyph heights
		 * @default false
		 */
		normalizeHeight?: boolean
		/**
		 * Exclude the run from spelling and grammar checking
		 * @default false
		 */
		noProof?: boolean
		/**
		 * Mark the run as needing re-inspection by the consumer
		 * - written as `dirty`, which was previously hardcoded to `0`
		 * @default false
		 */
		dirty?: boolean
		/**
		 * Underline line properties (`a:uLn`), distinct from the underline colour
		 * - pass `'text'` to follow the run's own line (`a:uLnTx`)
		 */
		underlineLine?: 'text' | ShapeLineProps
		/**
		 * Symbol font for the run (`a:sym`), for Wingdings-style glyphs
		 */
		symbolFontFace?: string
		/**
		 * Per-script typefaces (`a:latin` / `a:ea` / `a:cs`)
		 * - `fontFace` sets all three; these override individual scripts
		 */
		latinFontFace?: string
		eastAsianFontFace?: string
		complexScriptFontFace?: string
	}
	export interface TextPropsOptions extends PositionProps, DataOrPathProps, TextBaseProps, ObjectNameProps, TextBodyProps, ParagraphProps, TextRunProps {
		/** text-body, paragraph, and run attribute completeness (ECMA-376 21.1.2)
		 */
		baseline?: number
		/**
		 * Character spacing
		 */
		charSpacing?: number
		/**
		 * Text fit options
		 *
		 * MS-PPT > Format Shape > Shape Options > Text Box > "[unlabeled group]": [3 options below]
		 * - 'none' = Do not Autofit
		 * - 'shrink' = Shrink text on overflow
		 * - 'resize' = Resize shape to fit text
		 *
		 * **Note** 'shrink' and 'resize' only take effect after editing text/resize shape.
		 * Both PowerPoint and Word dynamically calculate a scaling factor and apply it when edit/resize occurs.
		 *
		 * There is no way for this library to trigger that behavior, sorry.
		 * @since v3.3.0
		 * @default "none"
		 */
		fit?: 'none' | 'shrink' | 'resize'
		/**
		 * Shape fill
		 * @example { color:'FF0000' } // hex color (red)
		 * @example { color:'0088CC', transparency:50 } // hex color, 50% transparent
		 * @example { color:pptx.SchemeColor.accent1 } // theme color Accent1
		 */
		fill?: ShapeFillProps
		/**
		 * Flip shape horizontally?
		 * @default false
		 */
		flipH?: boolean
		/**
		 * Flip shape vertical?
		 * @default false
		 */
		flipV?: boolean
		glow?: TextGlowProps
		hyperlink?: HyperlinkProps
		/**
		 * Mouse-over action, configured like `hyperlink` but triggered on hover
		 * - PowerPoint's Insert > Action > Mouse Over tab
		 * @example { hyperlinkHover: { slide: 3, tooltip: 'Jump to results' } }
		 */
		hyperlinkHover?: HyperlinkProps
		indentLevel?: number
		isTextBox?: boolean
		line?: ShapeLineProps
		/**
		 * Line spacing (pt)
		 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Exactly"
		 * @example 28 // 28pt
		 */
		lineSpacing?: number
		/**
		 * line spacing multiple (percent)
		 * - range: 0.0-9.99
		 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Multiple"
		 * @example 1.5 // 1.5X line spacing
		 * @since v3.5.0
		 */
		lineSpacingMultiple?: number
		// TODO: [20220219] powerpoint uses inches but library has always been pt... @future @deprecated - update in v4.0? [range: 0.0-22.0]
		/**
		 * Margin (points)
		 * - PowerPoint: Format Shape > Shape Options > Size & Properties > Text Box > Left/Right/Top/Bottom margin
		 * @default "Normal" margin in PowerPoint [3.5, 7.0, 3.5, 7.0] // (this library sets no value, but PowerPoint defaults to "Normal" [0.05", 0.1", 0.05", 0.1"])
		 * @example 0 // Top/Right/Bottom/Left margin 0 [0.0" in powerpoint]
		 * @example 10 // Top/Right/Bottom/Left margin 10 [0.14" in powerpoint]
		 * @example [10,5,10,5] // Top margin 10, Right margin 5, Bottom margin 10, Left margin 5
		 */
		margin?: Margin
		outline?: { color: Color, size: number }
		paraSpaceAfter?: number
		paraSpaceBefore?: number
		placeholder?: string
		/**
		 * Rounded rectangle radius (only for pptx.shapes.ROUNDED_RECTANGLE)
		 * - values: 0.0 to 1.0
		 * @default 0
		 */
		rectRadius?: number
		/**
		 * Rotation (degrees)
		 * - range: -360 to 360
		 * @default 0
		 * @example 180 // rotate 180 degrees
		 */
		rotate?: number
		/**
		 * Whether to enable right-to-left mode
		 * @default false
		 */
		rtlMode?: boolean
		shadow?: ShadowProps
		shape?: SHAPE_NAME
		strike?: boolean | 'dblStrike' | 'sngStrike'
		subscript?: boolean
		superscript?: boolean
		/**
		 * Vertical alignment
		 * @default middle
		 */
		valign?: VAlign
		vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
		/**
		 * Text wrap
		 * @since v3.3.0
		 * @default true
		 */
		wrap?: boolean
		/**
		 * Split the text box into newspaper-style columns (`a:bodyPr@numCol`)
		 * - range: 1-16; PowerPoint's Format Shape > Text Options > Columns
		 * @default 1
		 */
		columns?: number
		/**
		 * Space between columns, in inches (`a:bodyPr@spcCol`)
		 * - only meaningful with `columns` > 1
		 * @default 0
		 */
		columnSpacing?: number
		/**
		 * Render this run as a field the consumer refreshes, rather than literal text (`a:fld`)
		 * - the run's `text` becomes the cached value, so consumers that do not refresh still show something
		 * @example { text: '22/08/2026', options: { field: 'datetime1' } }
		 */
		field?: TextFieldType
		/**
		 * Horizontal-in-vertical numerals for East Asian vertical text (`a:rPr@kumimoji`)
		 * @default false
		 */
		kumimoji?: boolean
		/**
		 * Office Math (OMML) markup for this run
		 * - the run is emitted as a math zone instead of a plain text run; `text` becomes the
		 *   `mc:Fallback` shown by consumers that do not understand Office math
		 * - the value must be **well-formed OMML** (`m:oMath`, `m:oMathPara`, or an inner fragment
		 *   such as `<m:f>...</m:f>`); convert from LaTeX/MathML with a library of your choice
		 * - PptxGenJS supplies the `m:oMath` root and namespace declarations when they are missing
		 * @example { text: 'a/b', options: { omml: '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>' } }
		 */
		omml?: string

		/**
		 * Whether "Fit to Shape?" is enabled
		 * @deprecated v3.3.0 - use `fit`
		 */
		autoFit?: boolean
		/**
		 * Whather "Shrink Text on Overflow?" is enabled
		 * @deprecated v3.3.0 - use `fit`
		 */
		shrinkText?: boolean
		/**
		 * Inset
		 * @deprecated v3.10.0 - use `margin`
		 */
		inset?: number
		/**
		 * Dash type
		 * @deprecated v3.3.0 - use `line.dashType`
		 */
		lineDash?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * @deprecated v3.3.0 - use `line.beginArrowType`
		 */
		lineHead?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * @deprecated v3.3.0 - use `line.width`
		 */
		lineSize?: number
		/**
		 * @deprecated v3.3.0 - use `line.endArrowType`
		 */
		lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * Animation applied to this object
		 * @example { animation: { type: 'fadeIn', trigger: 'afterPrevious', duration: 800 } }
		 */
		animation?: AnimationProps
	}
	export interface TextProps {
		text?: string
		options?: TextPropsOptions
	}

	// charts =========================================================================================
	// FUTURE: BREAKING-CHANGE: (soln: use `OptsDataLabelPosition|string` until 3.5/4.0)
	/*
	export interface OptsDataLabelPosition {
		pie: 'ctr' | 'inEnd' | 'outEnd' | 'bestFit'
		scatter: 'b' | 'ctr' | 'l' | 'r' | 't'
		// TODO: add all othere chart types
	}
	*/

	export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
	export type ChartLineCap = 'flat' | 'round' | 'square'

	export interface OptsChartData {
		//_dataIndex?: number

		/**
		 * category labels
		 * @example ['Year 2000', 'Year 2010', 'Year 2020'] // single-level category axes labels
		 * @example [['Year 2000', 'Year 2010', 'Year 2020'], ['Decades', '', '']] // multi-level category axes labels
		 * @since `labels` string[][] type added v3.11.0
		 */
		labels?: string[] | string[][]
		/**
		 * series name
		 * @example 'Locations'
		 */
		name?: string
		/**
		 * bubble sizes
		 * @example [5, 1, 5, 1]
		 */
		sizes?: number[]
		/**
		 * category values
		 * @example [2000, 2010, 2020]
		 */
		values?: number[]
		/**
		 * Series color - overrides the `chartColors` cycle for this series only
		 * - hex color or the string `'transparent'`
		 * - pie/doughnut charts colour each data point rather than each series: use `chartColors` for those
		 * @example 'FF0000' // this series is red, the rest follow `chartColors`
		 */
		color?: string
		/**
		 * Per-point custom data label text. Defined entries replace the numeric label at that index.
		 */
		dataLabels?: string[]
	}
	export interface OptsChartGridLine {
		/**
		 * MS-PPT > Chart format > Format Major Gridlines > Line > Cap type
		 * - line cap type
		 * @default flat
		 */
		cap?: ChartLineCap
		/**
		 * Gridline color (hex)
		 * @example 'FF3399'
		 */
		color?: Color
		/**
		 * Gridline size (points)
		 */
		size?: number
		/**
		 * Gridline style
		 */
		style?: 'solid' | 'dash' | 'dot' | 'none'
	}
	// TODO: 202008: chart types remain with predicated with "I" in v3.3.0 (ran out of time!)
	export interface IChartMulti {
		type: CHART_NAME
		data: OptsChartData[]
		options: IChartOpts
	}
	export interface IChartPropsFillLine {
		/**
		 * PowerPoint: Format Chart Area/Plot > Border ["Line"]
		 * @example border: {color: 'FF0000', pt: 1} // hex RGB color, 1 pt line
		 */
		border?: BorderProps
		/**
		 * PowerPoint: Format Chart Area/Plot Area > Fill
		 * @example fill: {color: '696969'} // hex RGB color value
		 * @example fill: {color: pptx.SchemeColor.background2} // Theme color value
		 * @example fill: {transparency: 50} // 50% transparency
		 */
		fill?: ShapeFillProps
	}
	export interface IChartAreaProps extends IChartPropsFillLine {
		/**
		 * Whether the chart area has rounded corners
		 * - only applies when either `fill` or `border` is used
		 * @default true
		 * @since v3.11
		 */
		roundedCorners?: boolean
	}
	export interface IChartPropsBase {
		/**
		 * Axis position
		 */
		axisPos?: 'b' | 'l' | 'r' | 't'
		chartColors?: HexColor[]
		/**
		 * opacity (0 - 100)
		 * @example 50 // 50% opaque
		 */
		chartColorsOpacity?: number
		dataBorder?: BorderProps
		displayBlanksAs?: string
		invertedColors?: HexColor[]
		lang?: string
		layout?: PositionProps
		shadow?: ShadowProps
		/**
		 * @default false
		 */
		showLabel?: boolean
		showLeaderLines?: boolean
		/**
		 * @default false
		 */
		showLegend?: boolean
		/**
		 * @default false
		 */
		showPercent?: boolean
		/**
		 * @default false
		 */
		showSerName?: boolean
		/**
		 * @default false
		 */
		showTitle?: boolean
		/**
		 * @default false
		 */
		showValue?: boolean
		/**
		 * 3D Perspecitve
		 * - range: 0-120
		 * @default 30
		 */
		v3DPerspective?: number
		/**
		 * Right Angle Axes
		 * - Shows chart from first-person perspective
		 * - Overrides `v3DPerspective` when true
		 * - PowerPoint: Chart Options > 3-D Rotation
		 * @default false
		 */
		v3DRAngAx?: boolean
		/**
		 * X Rotation
		 * - PowerPoint: Chart Options > 3-D Rotation
		 * - range: 0-359.9
		 * @default 30
		 */
		v3DRotX?: number
		/**
		 * Y Rotation
		 * - range: 0-359.9
		 * @default 30
		 */
		v3DRotY?: number

		/**
		 * PowerPoint: Format Chart Area (Fill & Border/Line)
		 * @since v3.11
		 */
		chartArea?: IChartAreaProps
		/**
		 * PowerPoint: Format Plot Area (Fill & Border/Line)
		 * @since v3.11
		 */
		plotArea?: IChartPropsFillLine

		/**
		 * @deprecated v3.11.0 - use `plotArea.border`
		 */
		border?: BorderProps
		/**
		 * @deprecated v3.11.0 - use `plotArea.fill`
		 */
		fill?: HexColor
	}
	export interface IChartPropsAxisCat {
		/**
		 * Multi-Chart prop: array of cat axes
		 */
		catAxes?: IChartPropsAxisCat[]
		catAxisBaseTimeUnit?: string
		catAxisCrossesAt?: number | 'autoZero'
		catAxisHidden?: boolean
		catAxisLabelColor?: string
		catAxisLabelFontBold?: boolean
		catAxisLabelFontFace?: string
		catAxisLabelFontItalic?: boolean
		catAxisLabelFontSize?: number
		catAxisLabelFrequency?: string
		catAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		catAxisLabelRotate?: number
		catAxisLineColor?: string
		catAxisLineShow?: boolean
		catAxisLineSize?: number
		catAxisLineStyle?: 'solid' | 'dash' | 'dot'
		catAxisMajorTickMark?: ChartAxisTickMark
		catAxisMajorTimeUnit?: string
		catAxisMajorUnit?: number
		catAxisMaxVal?: number
		catAxisMinorTickMark?: ChartAxisTickMark
		catAxisMinorTimeUnit?: string
		catAxisMinorUnit?: number
		catAxisMinVal?: number
		/** @since v3.11.0 */
		catAxisMultiLevelLabels?: boolean
		catAxisOrientation?: 'minMax'
		catAxisTitle?: string
		catAxisTitleColor?: string
		catAxisTitleFontFace?: string
		catAxisTitleFontSize?: number
		catAxisTitleRotate?: number
		catGridLine?: OptsChartGridLine
		catLabelFormatCode?: string
		/**
		 * Whether data should use secondary category axis (instead of primary)
		 * @default false
		 */
		secondaryCatAxis?: boolean
		showCatAxisTitle?: boolean
	}
	export interface IChartPropsAxisSer {
		serAxisBaseTimeUnit?: string
		serAxisHidden?: boolean
		serAxisLabelColor?: string
		serAxisLabelFontBold?: boolean
		serAxisLabelFontFace?: string
		serAxisLabelFontItalic?: boolean
		serAxisLabelFontSize?: number
		serAxisLabelFrequency?: string
		serAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		serAxisLineColor?: string
		serAxisLineShow?: boolean
		serAxisMajorTimeUnit?: string
		serAxisMajorUnit?: number
		serAxisMinorTimeUnit?: string
		serAxisMinorUnit?: number
		serAxisOrientation?: string
		serAxisTitle?: string
		serAxisTitleColor?: string
		serAxisTitleFontFace?: string
		serAxisTitleFontSize?: number
		serAxisTitleRotate?: number
		serGridLine?: OptsChartGridLine
		serLabelFormatCode?: string
		showSerAxisTitle?: boolean
	}
	export interface IChartPropsAxisVal {
		/**
		 * Whether data should use secondary value axis (instead of primary)
		 * @default false
		 */
		secondaryValAxis?: boolean
		showValAxisTitle?: boolean
		/**
		 * Multi-Chart prop: array of val axes
		 */
		valAxes?: IChartPropsAxisVal[]
		valAxisCrossesAt?: number | 'autoZero'
		valAxisDisplayUnit?: 'billions' | 'hundredMillions' | 'hundreds' | 'hundredThousands' | 'millions' | 'tenMillions' | 'tenThousands' | 'thousands' | 'trillions'
		valAxisDisplayUnitLabel?: boolean
		valAxisHidden?: boolean
		valAxisLabelColor?: string
		valAxisLabelFontBold?: boolean
		valAxisLabelFontFace?: string
		valAxisLabelFontItalic?: boolean
		valAxisLabelFontSize?: number
		valAxisLabelFormatCode?: string
		valAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		valAxisLabelRotate?: number
		valAxisLineColor?: string
		valAxisLineShow?: boolean
		valAxisLineSize?: number
		valAxisLineStyle?: 'solid' | 'dash' | 'dot'
		/**
		 * PowerPoint: Format Axis > Axis Options > Logarithmic scale - Base
		 * - range: 2-99
		 * @since v3.5.0
		 */
		valAxisLogScaleBase?: number
		valAxisMajorTickMark?: ChartAxisTickMark
		valAxisMajorUnit?: number
		valAxisMaxVal?: number
		valAxisMinorTickMark?: ChartAxisTickMark
		valAxisMinVal?: number
		valAxisOrientation?: 'minMax'
		valAxisTitle?: string
		valAxisTitleColor?: string
		valAxisTitleFontFace?: string
		valAxisTitleFontSize?: number
		valAxisTitleRotate?: number
		valGridLine?: OptsChartGridLine
		/**
		 * Value label format code
		 * - this also directs Data Table formatting
		 * @since v3.3.0
		 * @example '#%' // round percent
		 * @example '0.00%' // shows values as '0.00%'
		 * @example '$0.00' // shows values as '$0.00'
		 */
		valLabelFormatCode?: string
	}
	export interface IChartPropsChartBar {
		/**
		 * 3D bar shape
		 * @default 'box'
		 */
		bar3DShape?: 'box' | 'cone' | 'coneToMax' | 'cylinder' | 'pyramid' | 'pyramidToMax'
		/**
		 * Bar direction - horizontal bars or vertical columns
		 * @default 'col'
		 */
		barDir?: 'bar' | 'col'
		barGapDepthPct?: number
		/**
		 * MS-PPT > Format chart > Format Data Point > Series Options >  "Gap Width"
		 * - width (percent)
		 * - range: `0`-`500`
		 * @default 150
		 */
		barGapWidthPct?: number
		/**
		 * Bar grouping
		 * @default 'clustered'
		 */
		barGrouping?: 'clustered' | 'percentStacked' | 'stacked' | 'standard'
		/**
		 * MS-PPT > Format chart > Format Data Point > Series Options >  "Series Overlap"
		 * - overlap (percent)
		 * - range: `-100`-`100`
		 * @since v3.9.0
		 * @default 0
		 */
		barOverlapPct?: number
	}
	export interface IChartPropsChartDoughnut {
		dataNoEffects?: boolean
		holeSize?: number
	}
	export interface IChartPropsChartLine {
		/**
		 * MS-PPT > Chart format > Format Data Series > Line > Cap type
		 * - line cap type
		 * @default flat
		 */
		lineCap?: ChartLineCap
		/**
		 * MS-PPT > Chart format > Format Data Series > Marker Options > Built-in > Type
		 * - line dash type
		 * @default solid
		 */
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		/**
		 * MS-PPT > Chart format > Format Data Series > Marker Options > Built-in > Type
		 * - marker type
		 * @default circle
		 */
		lineDataSymbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle'
		/**
		 * MS-PPT > Chart format > Format Data Series > [Marker Options] > Border > Color
		 * - border color
		 * @default circle
		 */
		lineDataSymbolLineColor?: string
		/**
		 * MS-PPT > Chart format > Format Data Series > [Marker Options] > Border > Width
		 * - border width (points)
		 * @default 0.75
		 */
		lineDataSymbolLineSize?: number
		/**
		 * MS-PPT > Chart format > Format Data Series > Marker Options > Built-in > Size
		 * - marker size
		 * - range: 2-72
		 * @default 6
		 */
		lineDataSymbolSize?: number
		/**
		 * MS-PPT > Chart format > Format Data Series > Line > Width
		 * - line width (points)
		 * - range: 0-1584
		 * @default 2
		 */
		lineSize?: number
		/**
		 * MS-PPT > Chart format > Format Data Series > Line > Smoothed line
		 * - "Smoothed line"
		 * @default false
		 */
		lineSmooth?: boolean
	}
	export interface IChartPropsChartPie {
		dataNoEffects?: boolean
		/**
		 * MS-PPT > Format chart > Format Data Series > Series Options >  "Angle of first slice"
		 * - angle (degrees)
		 * - range: 0-359
		 * @since v3.4.0
		 * @default 0
		 */
		firstSliceAng?: number
	}
	export interface IChartPropsChartRadar {
		/**
		 * MS-PPT > Chart Type > Waterfall
		 * - radar chart type
		 * @default standard
		 */
		radarStyle?: 'standard' | 'marker' | 'filled' // TODO: convert to 'radar'|'markers'|'filled' in 4.0 (verbatim with PPT app UI)
	}
	/**
	 * Options for the PowerPoint 2016+ chartex layouts (waterfall, funnel, treemap, sunburst, histogram,
	 * box & whisker). Each is ignored by every other chart type.
	 * @see MS-ODRAWXML 2.1
	 */
	export interface IChartPropsChartEx {
		/**
		 * MS-PPT > Chart Type > Histogram > Format Axis > "Number of bins"
		 * - fixed bin count; ignored when `chartExBinSize` is set
		 * - omit both to let PowerPoint choose the bins
		 * @example 8
		 */
		chartExBinCount?: number
		/**
		 * MS-PPT > Chart Type > Histogram > Format Axis > "Bin width"
		 * - fixed bin width, in value-axis units; takes precedence over `chartExBinCount`
		 * @example 5
		 */
		chartExBinSize?: number
		/**
		 * MS-PPT > Chart Type > Box & Whisker > Format Data Series > "Show mean line"
		 * @default false
		 */
		chartExMeanLine?: boolean
		/**
		 * MS-PPT > Chart Type > Treemap > Format Data Series > "Treemap Label Options"
		 * - how a parent category label is drawn over its children
		 * @default overlapping
		 */
		chartExParentLabels?: 'none' | 'overlapping' | 'banner'
		/**
		 * MS-PPT > Chart Type > Waterfall > "Set as Total"
		 * - zero-based indexes of the data points drawn as absolute totals rather than deltas
		 * @example [0, 4] // first and fifth bars are totals
		 */
		chartExSubtotals?: number[]
	}
	export interface IChartPropsDataLabel {
		dataLabelBkgrdColors?: boolean
		dataLabelColor?: string
		dataLabelFontBold?: boolean
		dataLabelFontFace?: string
		dataLabelFontItalic?: boolean
		dataLabelFontSize?: number
		/**
		 * Data label format code
		 * @example '#%' // round percent
		 * @example '0.00%' // shows values as '0.00%'
		 * @example '$0.00' // shows values as '$0.00'
		 */
		dataLabelFormatCode?: string
		dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY'
		/**
		 * Data label position
		 * - friendly names are translated to their OOXML codes; the codes themselves are still accepted
		 * - valid values differ per chart type: a value the chart type does not support is dropped with a
		 *   console warning rather than producing a file PowerPoint asks to repair
		 * @example 'outsideEnd' // bar/column, pie
		 * @example 'top' // line, scatter, radar
		 */
		dataLabelPosition?:
		| 'bottom' | 'center' | 'left' | 'right' | 'top' | 'insideEnd' | 'insideBase' | 'outsideEnd' | 'bestFit'
		| 'b' | 'ctr' | 'l' | 'r' | 't' | 'inEnd' | 'inBase' | 'outEnd'
	}
	export interface IChartPropsDataTable {
		dataTableFontSize?: number
		/**
		 * Data table format code
		 * @since v3.3.0
		 * @example '#%' // round percent
		 * @example '0.00%' // shows values as '0.00%'
		 * @example '$0.00' // shows values as '$0.00'
		 */
		dataTableFormatCode?: string
		/**
		 * Whether to show a data table adjacent to the chart
		 * @default false
		 */
		showDataTable?: boolean
		showDataTableHorzBorder?: boolean
		showDataTableKeys?: boolean
		showDataTableOutline?: boolean
		showDataTableVertBorder?: boolean
	}
	export interface IChartPropsLegend {
		legendColor?: string
		legendFontFace?: string
		legendFontSize?: number
		legendPos?: 'b' | 'l' | 'r' | 't' | 'tr'
	}
	export interface IChartPropsTitle extends TextBaseProps {
		title?: string
		titleAlign?: string
		titleBold?: boolean
		titleColor?: string
		titleFontFace?: string
		titleFontSize?: number
		titleItalic?: boolean
		titlePos?: { x: number, y: number }
		titleRotate?: number
	}
	export interface IChartOpts
		extends IChartPropsAxisCat,
		IChartPropsAxisSer,
		IChartPropsAxisVal,
		IChartPropsBase,
		IChartPropsChartBar,
		IChartPropsChartDoughnut,
		IChartPropsChartEx,
		IChartPropsChartLine,
		IChartPropsChartPie,
		IChartPropsChartRadar,
		IChartPropsDataLabel,
		IChartPropsDataTable,
		IChartPropsLegend,
		IChartPropsTitle,
		ObjectNameProps,
		OptsChartGridLine,
		PositionProps {
		/**
		 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
		 * - PowerPoint: [right-click on a chart] > "Edit Alt Text..."
		 */
		altText?: string
		/**
		 * Which entry of PowerPoint's Chart Styles gallery is selected (`cs:chartStyle@id`)
		 * - the style *definitions* written are Office's defaults regardless of this id, so it selects
		 *   the gallery entry rather than changing the formatting
		 * @default 201
		 */
		chartStyle?: number
		/**
		 * Chart colour style (`ppt/charts/colorsN.xml`)
		 * - distinct from `chartColors`, which sets the series colours directly on the chart
		 */
		chartColorStyle?: ChartColorStyleProps
	}
	export interface ISlideRelChart extends OptsChartData {
		type: CHART_NAME | IChartMulti[]
		opts: IChartOpts
		data: OptsChartData[]
		// internal below
		//rId: number
		//Target: string
		//globalId: number
		//fileName: string
	}

	// Core
	// ====
	/**
	 * Zip compression applied to the exported file
	 * - 'none': store uncompressed (fastest)
	 * - 'fast': DEFLATE level 1 (quick, decent savings)
	 * - 'best': DEFLATE level 9 (smallest file, slowest)
	 * @since v4.1.0
	 */
	export type CompressionLevel = 'none' | 'fast' | 'best'

	export interface WriteBaseProps {
		/**
		 * Whether to compress export (can save substantial space, but takes a bit longer to export)
		 * @default false
		 * @since v3.5.0
		 * @deprecated v4.1.0 - set `compression` on the presentation instead (`pptx.compression = 'best'`)
		 */
		compression?: boolean
	}
	export interface WriteProps extends WriteBaseProps {
		/**
		 * Output type
		 * - values: 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array' | 'STREAM'
		 * @default 'blob'
		 */
		outputType?: WRITE_OUTPUT_TYPE
	}
	export interface WriteFileProps extends WriteBaseProps {
		/**
		 * Export file name
		 * @default 'Presentation.pptx'
		 */
		fileName?: string
	}
	export interface SectionProps {
		//_type: 'user' | 'default'
		//_slides: PresSlide[]

		/**
		 * Section title
		 */
		title: string
		/**
		 * Section order - uses to add section at any index
		 * - values: 1-n
		 */
		order?: number
	}
	export interface PresLayout {
		//_sizeW?: number
		//_sizeH?: number

		/**
		 * Layout Name
		 * @example 'LAYOUT_WIDE'
		 */
		name: string
		width: number
		height: number
	}
	/**
	 * Argument to `defineLayout()` - dimensions may be given as `width`/`height` or as `w`/`h`
	 */
	export interface DefineLayoutProps {
		/**
		 * Layout name
		 * @example 'A3'
		 */
		name: string
		/** Layout width (inches) - or use `w` */
		width?: number
		/** Layout height (inches) - or use `h` */
		height?: number
		/** Layout width (inches) - alias of `width` */
		w?: number
		/** Layout height (inches) - alias of `height` */
		h?: number
	}
	export interface SlideNumberProps extends PositionProps, TextBaseProps {
		/**
		 * margin (points)
		 */
		margin?: Margin // TODO: convert to inches in 4.0 (valid values are 0-22)
	}
	/**
	 * ECMA-376 20.1.10.14 ST_ColorSchemeIndex - a slot in the theme's colour scheme
	 */
	export type ColorSchemeIndex = 'dk1' | 'lt1' | 'dk2' | 'lt2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink'

	/**
	 * Per-layout colour map override (`p:clrMapOvr` > `a:overrideClrMapping`)
	 * - all twelve attributes are required by the schema, so anything left unset is filled from the
	 *   identity map (which is what inheriting the master's mapping means)
	 */
	export interface ColorMapOverrideProps {
		bg1?: ColorSchemeIndex
		tx1?: ColorSchemeIndex
		bg2?: ColorSchemeIndex
		tx2?: ColorSchemeIndex
		accent1?: ColorSchemeIndex
		accent2?: ColorSchemeIndex
		accent3?: ColorSchemeIndex
		accent4?: ColorSchemeIndex
		accent5?: ColorSchemeIndex
		accent6?: ColorSchemeIndex
		hlink?: ColorSchemeIndex
		folHlink?: ColorSchemeIndex
	}

	/**
	 * ECMA-376 19.7.15 ST_SlideLayoutType - the placeholder arrangement a layout describes
	 * - PowerPoint's layout gallery and its Reset Layout command read this
	 */
	export type SlideLayoutType = 'title' | 'tx' | 'twoColTx' | 'tbl' | 'txAndChart' | 'chartAndTx' | 'dgm' | 'chart' | 'txAndClipArt' | 'clipArtAndTx' | 'titleOnly' | 'blank' | 'txAndObj' | 'objAndTx' | 'objOnly' | 'obj' | 'txAndMedia' | 'mediaAndTx' | 'objOverTx' | 'txOverObj' | 'txAndTwoObj' | 'twoObjAndTx' | 'twoObjOverTx' | 'fourObj' | 'vertTx' | 'clipArtAndVertTx' | 'vertTitleAndTx' | 'vertTitleAndTxOverChart' | 'twoObj' | 'objAndTwoObj' | 'twoObjAndObj' | 'cust' | 'secHead' | 'twoTxTwoObj' | 'objTx' | 'picTx'

	export interface SlideMasterProps {
		/**
		 * Unique name for this master
		 */
		title: string
		/**
		 * Which placeholder arrangement this layout describes (`p:sldLayout@type`)
		 * - PowerPoint's New Slide gallery groups layouts by this, and Reset Layout trusts it
		 * @default cust
		 */
		layoutType?: SlideLayoutType
		/**
		 * Name shown for the layout in PowerPoint's gallery (`@matchingName`)
		 * - unset writes nothing; PowerPoint then falls back to the layout name from `title`
		 */
		matchingName?: string
		/**
		 * Keep the layout in the deck even when no slide uses it (`@preserve`)
		 * @default true
		 */
		preserve?: boolean
		/**
		 * Draw the master's shapes behind slides using this layout (`@showMasterSp`)
		 * @default true
		 */
		showMasterShapes?: boolean
		/**
		 * Play the master's placeholder animations on slides using this layout (`@showMasterPhAnim`)
		 * @default true
		 */
		showMasterPlaceholderAnimation?: boolean
		/**
		 * Mark the layout as drawn by the author rather than generated (`@userDrawn`)
		 * @default false
		 */
		userDrawn?: boolean
		/**
		 * Remap the theme's colour slots for this layout only (`p:clrMapOvr`)
		 * - with none set the layout inherits the master's mapping, which is today's behaviour
		 */
		colorMapOverride?: ColorMapOverrideProps
		/**
		 * Transition applied to slides using this layout (`p:transition`)
		 */
		transition?: SlideTransitionProps
		background?: BackgroundProps
		margin?: Margin
		slideNumber?: SlideNumberProps
		objects?: Array<| { chart: IChartOpts }
			| { image: ImageProps }
			| { line: ShapeProps }
			| { rect: ShapeProps }
			| { text: TextProps }
			/** any of the 180+ shape types (`line`/`rect` above are shorthands) */
			| { shape: { type: SHAPE_NAME, options?: ShapeProps } }
			| { table: { rows: TableRow[], options?: TableProps } }
			| { media: MediaProps }
			| {
				placeholder: {
					options: PlaceholderProps
					/**
					 * Text to be shown in placeholder (shown until user focuses textbox or adds text)
					 * - Leave blank to have powerpoint show default phrase (ex: "Click to add title")
					 */
					text?: string
				}
			}>

		/**
		 * @deprecated v3.3.0 - use `background`
		 */
		bkgd?: string | BackgroundProps
	}
	export interface ObjectOptions extends ImageProps, PositionProps, ShapeProps, TableCellProps, TextPropsOptions {
		//_placeholderIdx?: number
		//_placeholderType?: PLACEHOLDER_TYPE

		cx?: Coord
		cy?: Coord
		margin?: Margin
		colW?: number | number[] // table
		rowH?: number | number[] // table
	}
	export interface PresSlide {
		addChart: Function
		addImage: Function
		addMedia: Function
		addNotes: Function
		addShape: Function
		addTable: Function
		addText: Function

		/**
		 * Background color or image (`color` | `path` | `data`)
		 * @example { color: 'FF3399' } - hex color
		 * @example { color: 'FF3399', transparency:50 } - hex color with 50% transparency
		 * @example { path: 'https://onedrives.com/myimg.png` } - retrieve image via URL
		 * @example { path: '/home/gitbrent/images/myimg.png` } - retrieve image via local path
		 * @example { data: 'image/png;base64,iVtDaDrF[...]=' } - base64 string
		 * @since v3.3.0
		 */
		background?: BackgroundProps
		/**
		 * Default text color (hex format)
		 * @example 'FF3399'
		 * @default '000000' (DEF_FONT_COLOR)
		 */
		color?: HexColor
		/**
		 * Whether slide is hidden
		 * @default false
		 */
		hidden?: boolean
		/** Slide transition (`<p:transition>`) */
		transition?: SlideTransitionProps
		/**
		 * Stable identity for this slide across saves (MS-PPTX 2.2.9 `p14:creationId`)
		 * - `true` assigns one, or pass your own unsigned 32-bit value
		 */
		creationId?: boolean | number
		/**
		 * Slide number options
		 */
		slideNumber?: SlideNumberProps
	}
	export type TransitionType =
		// ECMA-376 19.3.1.50 base transitions
		| 'blinds' | 'checker' | 'circle' | 'comb' | 'cover' | 'cut' | 'diamond' | 'dissolve' | 'fade'
		| 'newsflash' | 'none' | 'plus' | 'pull' | 'push' | 'random' | 'randomBar' | 'split' | 'strips'
		| 'wedge' | 'wheel' | 'wipe' | 'zoom'
		// PowerPoint 2010+ transitions, emitted through `mc:AlternateContent` with a base fallback
		| 'conveyor' | 'doors' | 'ferris' | 'flash' | 'flip' | 'flythrough' | 'gallery' | 'glitter'
		| 'honeycomb' | 'morph' | 'pan' | 'prism' | 'reveal' | 'ripple' | 'shred' | 'switch' | 'vortex'
		| 'warp' | 'wheelReverse' | 'window'
	export interface SlideTransitionProps {
		/**
		 * Transition effect
		 * - PowerPoint 2010+ effects are written with a base-transition fallback, so other consumers
		 *   still show a sensible transition instead of failing to open the file
		 */
		type: TransitionType
		/**
		 * Transition direction - OOXML tokens or friendly aliases (`left`, `up`, `horizontal`, ...)
		 * - `push`/`wipe`/`pan`/`reveal`/`vortex`: `l` | `r` | `u` | `d`
		 * - `blinds`/`checker`/`comb`/`randomBar`/`doors`/`window`: `horz` | `vert`
		 * - `cover`/`pull`/`conveyor`/`ferris`/`flip`/`gallery`/`glitter`/`prism`/`switch`: adds `lu` | `ru` | `ld` | `rd`
		 * - `strips`: `lu` | `ru` | `ld` | `rd`
		 * - `zoom`/`split`/`flythrough`/`warp`: `in` | `out`
		 * - a direction the effect does not accept is ignored with a warning
		 */
		direction?: string
		/**
		 * `split` only: which axis splits (`CT_SplitTransition@orient`)
		 * @default 'horz'
		 */
		orient?: 'horz' | 'vert' | 'horizontal' | 'vertical'
		/**
		 * Transition speed
		 * - ignored when `duration` is set
		 */
		speed?: 'slow' | 'med' | 'fast'
		/**
		 * Transition duration (milliseconds)
		 * - PowerPoint 2010+ (`p14:dur`); takes precedence over `speed`
		 */
		duration?: number
		/**
		 * Advance the slide on mouse click
		 * @default true
		 */
		advClick?: boolean
		/**
		 * Advance the slide automatically after this many milliseconds
		 */
		advTm?: number
		/**
		 * `wheel` only: number of spokes
		 * @default 4
		 */
		spokes?: 1 | 2 | 3 | 4 | 8
		/**
		 * `fade` and `cut` only: transition through black
		 * @default false
		 */
		thruBlk?: boolean
	}
	export type FontEmbedStyle = 'regular' | 'bold' | 'italic' | 'boldItalic'
	export interface AddFontProps {
		/**
		 * Typeface name, exactly as used in `fontFace` on text options
		 * @example 'Custom Sans'
		 */
		fontFace: string
		/**
		 * Font data as **Embedded OpenType (EOT)**
		 * - base64 (with or without a data-URI header), `ArrayBuffer`, or `Uint8Array`
		 * - PowerPoint stores embedded fonts as EOT; convert TTF/OTF/WOFF before calling
		 *   (ex: `ttf2eot`, `fonteditor-core`) - PptxGenJS does not convert font files
		 * - you must hold a licence that permits embedding the font
		 */
		data: string | ArrayBuffer | Uint8Array
		/**
		 * Which style of the typeface this file is
		 * - register several styles of one typeface with repeated `addFont()` calls
		 * @default 'regular'
		 */
		style?: FontEmbedStyle
	}
	export interface ZoomBaseProps extends PositionProps, ObjectNameProps {
		/**
		 * Thumbnail shown for the zoom (base64 image)
		 * - PowerPoint replaces it with a live thumbnail of the target when the file is opened; this is
		 *   what every other consumer shows, and what the `mc:Fallback` picture uses
		 * @default a 1x1 transparent placeholder
		 * @example 'image/png;base64,iVBORw0KGgo...'
		 */
		cover?: string
		/**
		 * Return to the parent slide when the zoom finishes
		 * @default true
		 */
		returnToParent?: boolean
		/**
		 * Keep the parent slide's background while zooming
		 * @default true
		 */
		showBg?: boolean
		/**
		 * Zoom transition duration (milliseconds)
		 */
		transitionDur?: number
		/**
		 * Alt text for the zoom object
		 */
		altText?: string
	}
	export interface SlideZoomProps extends ZoomBaseProps {
		/**
		 * Slide to zoom to, 1-based
		 */
		slideNumber: number
	}
	export interface SectionZoomProps extends ZoomBaseProps {
		/**
		 * Title of the section to zoom to, as passed to `addSection()`
		 */
		sectionTitle: string
	}
	export interface SummaryZoomProps extends ZoomBaseProps {
		/**
		 * Titles of the sections the summary links to, in order
		 * - at least one is required
		 */
		sectionTitles: string[]
	}
	export interface CommentAuthorProps {
		/**
		 * Author name shown on the comment
		 */
		name: string
		/**
		 * Initials shown in the comment avatar
		 */
		initials?: string
		/**
		 * Author id (GUID)
		 * - derived from the author's position when omitted, so output stays reproducible
		 */
		id?: string
		/**
		 * Identity-provider user id
		 * @default '' (an anonymous author)
		 */
		userId?: string
		/**
		 * Identity provider
		 * @default 'None'
		 */
		providerId?: string
	}
	export interface CommentReplyProps {
		/**
		 * Reply text
		 */
		text: string
		/**
		 * Author name; an author not passed to `pptx.commentAuthors` is added automatically
		 */
		author: string
		/**
		 * When the reply was written, ISO 8601
		 * - defaults to the time of export; pass it to keep generated packages reproducible
		 */
		created?: string
		/**
		 * Reply id (GUID) - derived from its position when omitted
		 */
		id?: string
	}
	export interface CommentProps {
		/**
		 * Comment text
		 */
		text: string
		/**
		 * Author name; an author not passed to `pptx.commentAuthors` is added automatically
		 */
		author: string
		/**
		 * Anchor position on the slide (inches)
		 * - both are needed for an anchored comment; omit both for a slide-level comment
		 */
		x?: number
		y?: number
		/**
		 * When the comment was written, ISO 8601
		 * - defaults to the time of export; pass it to keep generated packages reproducible
		 */
		created?: string
		/**
		 * Comment id (GUID) - derived from its position when omitted
		 */
		id?: string
		/**
		 * Whether the comment thread is resolved
		 * @default false
		 */
		resolved?: boolean
		/**
		 * Replies in the comment thread
		 */
		replies?: CommentReplyProps[]
	}
	export interface ContentPartProps extends PositionProps, ObjectNameProps {
		/**
		 * Payload markup for the embedded part
		 * - for ink this is the InkML document; for anything else, that format's markup
		 */
		data: string
		/**
		 * Content type of the payload, declared in `[Content_Types].xml`
		 * - belongs to the format being embedded, so it is required rather than guessed
		 * @example 'application/inkml+xml'
		 */
		contentType: string
		/**
		 * Relationship type linking the slide to the payload part
		 * - also format-specific and therefore required
		 * @example 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
		 */
		relationshipType: string
		/**
		 * File name for the payload inside the package
		 * @default 'contentPart<n>.xml'
		 */
		fileName?: string
		/**
		 * Whether this content part holds ink
		 * - ink must fall back to a raster picture, so `cover` becomes required (MS-PPTX 2.2.3.1)
		 * @default false
		 */
		ink?: boolean
		/**
		 * Raster preview shown by consumers that cannot render the payload (base64 image)
		 * - required when `ink` is true
		 */
		cover?: string
		/**
		 * Alt text for the fallback shape
		 */
		altText?: string
	}
	export interface AddSlideProps {
		masterName?: string // TODO: 20200528: rename to "masterTitle" (createMaster uses `title` so lets be consistent)
		sectionTitle?: string
		/** Slide transition applied at creation - same as setting `slide.transition` */
		transition?: SlideTransitionProps
	}
	export interface SlideShowProps {
		/**
		 * How the slide show runs (`p:showPr` choice, ECMA-376 19.2.1.30)
		 * - `present`: full screen, presenter-driven
		 * - `browse`: windowed, viewer-driven
		 * - `kiosk`: full screen, self-running
		 * @default 'present'
		 */
		mode?: 'present' | 'browse' | 'kiosk'
		/**
		 * Restart the show after the last slide
		 * @default false
		 */
		loop?: boolean
		/**
		 * Play recorded narration
		 * @default true
		 */
		showNarration?: boolean
		/**
		 * Play animations
		 * @default true
		 */
		showAnimation?: boolean
		/**
		 * Use the slide timings recorded in the file
		 * @default true
		 */
		useTimings?: boolean
		/**
		 * Show the browse-mode UI (MS-PPTX 2.2.6 `p14:browseMode`)
		 * - only meaningful with `mode: 'browse'`
		 */
		browseMode?: boolean
		/**
		 * Laser-pointer color (MS-PPTX 2.2.6 `p14:laserClr`)
		 * - `HexColor` or `ThemeColor`
		 * @example 'FF0000'
		 */
		laserColor?: Color
	}
	export interface GuideProps {
		/**
		 * Guide orientation
		 * - `horz` is a horizontal guide positioned by its distance from the top
		 * - `vert` is a vertical guide positioned by its distance from the left
		 */
		orientation: 'horz' | 'vert'
		/**
		 * Distance from the top (`horz`) or left (`vert`), in inches
		 */
		position: number
		/**
		 * Guide color
		 * - `HexColor` or `ThemeColor`
		 * @default 'A4A3A4' (PowerPoint's guide gray)
		 */
		color?: Color
	}
	/**
	 * Additional document properties (ECMA-376 15.2 and the OPC core properties)
	 * - each is omitted when unset, so `docProps` output is unchanged by default
	 */
	export interface DocumentProps {
		/** `dc:description` - the Comments field in PowerPoint's Info pane */
		description?: string
		/** `dc:language`, ex: 'en-US' */
		language?: string
		/** `dc:identifier` */
		identifier?: string
		/** `cp:keywords`, comma-separated */
		keywords?: string
		/** `cp:category` */
		category?: string
		/** `cp:contentStatus`, ex: 'Draft' */
		contentStatus?: string
		/** `cp:version` */
		version?: string
		/** `cp:lastPrinted`, ISO 8601 */
		lastPrinted?: string
		/** `Manager` in app.xml */
		manager?: string
		/** `Template` in app.xml */
		template?: string
		/** `HyperlinkBase` in app.xml - the base for relative hyperlinks */
		hyperlinkBase?: string
		/** `TotalTime` in app.xml - editing time in minutes */
		totalEditTime?: number
	}
	/**
	 * Slide-size preset (`p:sldSz@type`)
	 */
	export type SlideSizeType =
		| 'screen4x3' | 'screen16x9' | 'screen16x10' | 'letter' | 'ledger' | 'a3' | 'a4' | 'b4ISO'
		| 'b5ISO' | 'b4JIS' | 'b5JIS' | 'hagakiCard' | '35mm' | 'overhead' | 'banner' | 'custom'
	/**
	 * Photo-album mode (`p:photoAlbum`)
	 */
	export interface PhotoAlbumProps {
		/** render in black and white */
		blackWhite?: boolean
		/** show captions below each picture */
		showCaptions?: boolean
		/** pictures per slide and orientation */
		layout?: 'fitToSlide' | '1pic' | '2pic' | '4pic' | '1picTitle' | '2picTitle' | '4picTitle'
		/** frame style drawn around each picture */
		frame?: 'frameStyle1' | 'frameStyle2' | 'frameStyle3' | 'frameStyle4' | 'frameStyle5' | 'frameStyle6' | 'frameStyle7'
	}
	/**
	 * East Asian line-breaking rules (`p:kinsoku`)
	 */
	export interface KinsokuProps {
		/** language the rules apply to, ex: 'ja-JP' */
		lang?: string
		/** characters that may not start a line */
		invalidStartChars?: string
		/** characters that may not end a line */
		invalidEndChars?: string
	}
	/**
	 * Print defaults stored with the deck (`p:prnPr`)
	 */
	export interface PrintProps {
		/** what to print */
		what?: 'slides' | 'handouts1' | 'handouts2' | 'handouts3' | 'handouts4' | 'handouts6' | 'handouts9' | 'notes' | 'outline'
		/** colour mode */
		colorMode?: 'bw' | 'gray' | 'clr'
		/** include hidden slides */
		hiddenSlides?: boolean
		/** scale to fit the paper */
		scaleToFitPaper?: boolean
		/** draw a frame around each slide */
		frameSlides?: boolean
	}
	/**
	 * View properties stored in `ppt/viewProps.xml` (`p:viewPr`)
	 * - unset values keep the literal prior versions wrote
	 */
	export interface ViewProps {
		/** zoom percent in the normal slide view @default 136 */
		zoom?: number
		/** snap objects to the grid @default false */
		snapToGrid?: boolean
		/** snap objects to each other @default true */
		snapToObjects?: boolean
		/** show drawing guides */
		showGuides?: boolean
		/** show comments and ink markup */
		showComments?: boolean
		/** grid spacing in inches @default 0.0833 (76200 EMU) */
		gridSpacing?: number
		/** view PowerPoint opens the deck in */
		lastView?: 'sldView' | 'sldMasterView' | 'notesView' | 'handoutView' | 'notesMasterView' | 'outlineView' | 'sldSorterView' | 'sldThumbnailView'
		/** classic drawing guides (`p:guideLst`), positioned in inches */
		guides?: GuideProps[]
	}
	export interface PresentationProps {
		author: string
		company: string
		layout: string
		masterSlide: PresSlide
		/**
		 * Presentation's layout
		 * read-only
		 */
		presLayout: PresLayout
		revision: string
		/**
		 * Whether to enable right-to-left mode
		 * @default false
		 */
		rtlMode: boolean
		subject: string
		theme: ThemeProps
		title: string
		/**
		 * Image quality preference PowerPoint applies to inserted pictures
		 * (MS-PPTX 2.2.7 `p14:defaultImageDpi`)
		 * - `220` is PowerPoint's own default; `0` means "do not compress"
		 * - opt-in: unset leaves the presentation properties part unchanged
		 */
		defaultImageDpi?: number
		/**
		 * Whether PowerPoint discards crop/edit data on save (MS-PPTX 2.2.7 `p14:discardImageEditData`)
		 * - shrinks the file at the cost of no longer being able to undo picture edits
		 * @default false
		 */
		discardImageEditData?: boolean
		/**
		 * Whether PowerPoint recommends opening the file read-only
		 * (MS-PPTX 2.2.16 `p1710:readonlyRecommended`)
		 * @default false
		 */
		readonlyRecommended?: boolean
		/**
		 * Slide-show options (`p:showPr`)
		 * - opt-in: unset leaves the presentation properties part unchanged
		 */
		slideShow?: SlideShowProps
		/**
		 * Drawing guides shown in the slide editing view (MS-PPTX 2.2.11 `p15:sldGuideLst`)
		 * - opt-in: unset writes no guide list
		 * @example pptx.guides = [{ orientation: 'vert', position: 5 }, { orientation: 'horz', position: 3.75 }]
		 */
		guides?: GuideProps[]
		/**
		 * Drawing guides shown in the notes view (MS-PPTX 2.2.11 `p15:notesGuideLst`)
		 * - opt-in: unset writes no guide list
		 */
		notesGuides?: GuideProps[]
		/**
		 * Additional document properties written to `docProps`
		 */
		documentProps?: DocumentProps
		/**
		 * Slide-size preset (`p:sldSz@type`)
		 * - the dimensions come from the layout; this records which preset they match
		 */
		slideSizeType?: SlideSizeType
		/**
		 * Photo-album mode (`p:photoAlbum`)
		 */
		photoAlbum?: PhotoAlbumProps
		/**
		 * East Asian line-breaking rules (`p:kinsoku`)
		 */
		kinsoku?: KinsokuProps
		/**
		 * Print defaults stored with the deck (`p:prnPr`)
		 */
		printProps?: PrintProps
		/**
		 * Recently-used colours shown in the colour picker (`p:clrMru`)
		 */
		recentColors?: Color[]
		/**
		 * View properties (`ppt/viewProps.xml`)
		 */
		/**
		 * Custom table style definitions written into `ppt/tableStyles.xml`
		 * - reference one from a table with `tableStyleId`
		 */
		tableStyles?: TableStyleProps[]
		viewProps?: ViewProps
	}

	// LAST: Slide
	/**
	 * `slide.d.ts`
	 */
	export class Slide {
		/**
		 * Background color or image (`color` | `path` | `data`)
		 * @example { color: 'FF3399' } - hex color
		 * @example { color: 'FF3399', transparency:50 } - hex color with 50% transparency
		 * @example { path: 'https://onedrives.com/myimg.png` } - retrieve image via URL
		 * @example { path: '/home/gitbrent/images/myimg.png` } - retrieve image via local path
		 * @example { data: 'image/png;base64,iVtDaDrF[...]=' } - base64 string
		 * @since 3.3.0
		 */
		background: BackgroundProps
		/**
		 * Default text color (hex format)
		 * @example 'FF3399'
		 * @default '000000' (DEF_FONT_COLOR)
		 */
		color: HexColor
		/**
		 * Whether slide is hidden
		 * @default false
		 */
		hidden: boolean
		/** Slide transition (`<p:transition>`) */
		transition?: SlideTransitionProps
		/**
		 * Stable identity for this slide across saves (MS-PPTX 2.2.9 `p14:creationId`)
		 */
		creationId?: boolean | number
		/**
		 * Slide number options
		 */
		slideNumber: SlideNumberProps
		/**
		 * New slides added by an auto paged table
		 */
		newAutoPagedSlides: PresSlide[]
		/**
		 * Add chart to Slide
		 * @param {CHART_NAME|IChartMulti[]} type - chart type
		 * @param {object[]} data - data object
		 * @param {IChartOpts} options - chart options
		 * @return {Slide} this Slide
		 * @type {Function}
		 */
		addChart(type: CHART_NAME | IChartMulti[], data: any[], options?: IChartOpts): Slide
		/**
		 * Add image to Slide
		 * @param {ImageProps} options - image options
		 * @return {Slide} this Slide
		 */
		addImage(options: ImageProps): Slide
		/**
		 * Add media (audio/video) to Slide
		 * @param {MediaProps} options - media options
		 * @return {Slide} this Slide
		 */
		addMedia(options: MediaProps): Slide
		/**
		 * Add a threaded comment to this Slide
		 * @param options - comment props
		 */
		addComment(options: CommentProps): Slide
		/**
		 * Embed markup PresentationML does not define as its own package part
		 * @param options - content-part props
		 */
		addContentPart(options: ContentPartProps): Slide
		/**
		 * Add a zoom to another slide
		 * @param options - zoom props
		 */
		addZoom(options: SlideZoomProps): Slide
		/**
		 * Add a zoom to a section
		 * @param options - zoom props
		 */
		addSectionZoom(options: SectionZoomProps): Slide
		/**
		 * Add a summary zoom linking several sections
		 * @param options - zoom props
		 */
		addSummaryZoom(options: SummaryZoomProps): Slide
		/**
		 * Add speaker notes to Slide
		 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
		 * @param {string} notes - notes to add to slide
		 * @return {Slide} this Slide
		 */
		addNotes(notes: string): Slide
		/**
		 * Add shape to Slide
		 * @param {SHAPE_NAME} shapeName - shape name
		 * @param {ShapeProps} options - shape options
		 * @return {Slide} this Slide
		 */
		addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide
		/**
		 * Add table to Slide
		 * @param {TableRow[]} tableRows - table rows
		 * @param {TableProps} options - table options
		 * @return {Slide} this Slide
		 */
		addTable(tableRows: TableRow[], options?: TableProps): Slide
		/**
		 * Add text to Slide
		 * @param {string|TextProps[]} text - text string or complex object
		 * @param {TextPropsOptions} options - text options
		 * @return {Slide} this Slide
		 */
		addText(text: string | TextProps[], options?: TextPropsOptions): Slide

		/**
		 * Background color
		 * @deprecated in 3.3.0 - use `background` instead
		 */
		bkgd: string
	}
}
