/**
 * PptxGenJS: Slide Class
 */

import { CHART_NAME, SHAPE_NAME } from './core-enums'
import {
	AddSlideProps,
	BackgroundProps,
	HexColor,
	IChartMulti,
	IChartOpts,
	IOptsChartData,
	ISlideObject,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ImageProps,
	CommentProps,
	ContentPartProps,
	MediaProps,
	PresLayout,
	PresSlide,
	ShapeProps,
	SlideLayout,
	SectionZoomProps,
	SlideNumberProps,
	SlideZoomProps,
	SummaryZoomProps,
	SlideTransitionProps,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
} from './core-interfaces'
import * as genObj from './gen-objects'

/**
 * Copy caller-supplied options so the library never mutates objects it does not own (issue #20)
 * @note shallow by design - the generators rewrite top-level props only (nested props are read, not written)
 */
function cloneOpts<T extends object> (options?: T): T {
	return { ...(options ?? {}) } as T
}

export default class Slide {
	private readonly _setSlideNum: (value: SlideNumberProps) => void

	public addSlide: (options?: AddSlideProps) => PresSlide
	public getSlide: (slideNum: number) => PresSlide
	public _name: string
	public _presLayout: PresLayout
	public _rels: ISlideRel[]
	public _relsChart: ISlideRelChart[]
	public _relsMedia: ISlideRelMedia[]
	public _rId: number
	public _slideId: number
	public _slideLayout: SlideLayout
	public _slideNum: number
	public _slideNumberProps?: SlideNumberProps
	public _slideObjects: ISlideObject[]
	public _newAutoPagedSlides?: PresSlide[]

	constructor(params: {
		addSlide: (options?: AddSlideProps) => PresSlide
		getSlide: (slideNum: number) => PresSlide
		presLayout: PresLayout
		setSlideNum: (value: SlideNumberProps) => void
		slideId: number
		slideRId: number
		slideNumber: number
		slideLayout: SlideLayout
	}) {
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
		this._name = `Slide ${params.slideNumber}`
		this._presLayout = params.presLayout
		this._rId = params.slideRId
		this._rels = []
		this._relsChart = []
		this._relsMedia = []
		this._setSlideNum = params.setSlideNum
		this._slideId = params.slideId
		this._slideLayout = params.slideLayout
		this._slideNum = params.slideNumber
		this._slideObjects = []
		/** NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		 * `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		 * so, lastly, add to the Slide now.
		 */
		this._slideNumberProps = this._slideLayout?._slideNumberProps ? this._slideLayout._slideNumberProps : undefined
	}

	/**
	 * Background color
	 * @type {string|BackgroundProps}
	 * @deprecated in v3.3.0 - use `background` instead
	 */
	private _bkgd?: string | BackgroundProps
	public set bkgd(value: string | BackgroundProps) {
		this._bkgd = value
		if (!this._background || !this._background.color) {
			if (!this._background) this._background = {}
			if (typeof value === 'string') this._background.color = value
		}
	}

	public get bkgd(): string | BackgroundProps | undefined {
		return this._bkgd
	}

	/**
	 * Background color or image
	 * @type {BackgroundProps}
	 * @example solid color `background: { color:'FF0000' }`
	 * @example color+trans `background: { color:'FF0000', transparency:0.5 }`
	 * @example base64 `background: { data:'image/png;base64,ABC[...]123' }`
	 * @example url `background: { path:'https://some.url/image.jpg'}`
	 * @since v3.3.0
	 */
	private _background?: BackgroundProps
	public set background(props: BackgroundProps) {
		this._background = props
		// Add background (image data/path must be captured before `exportPresentation()` is called)
		if (props) genObj.addBackgroundDefinition(props, this)
	}

	public get background(): BackgroundProps | undefined {
		return this._background
	}

	/**
	 * Default font color
	 * @type {HexColor}
	 */
	private _color?: HexColor
	public set color(value: HexColor) {
		this._color = value
	}

	public get color(): HexColor | undefined {
		return this._color
	}

	/**
	 * @type {boolean}
	 */
	private _hidden = false
	public set hidden(value: boolean) {
		this._hidden = value
	}

	public get hidden(): boolean {
		return this._hidden
	}

	/**
	 * Comments on this slide
	 * @type {CommentProps[]}
	 */
	private _comments?: CommentProps[]
	public get comments(): CommentProps[] | undefined {
		return this._comments
	}

	/**
	 * Slide transition
	 * @type {SlideTransitionProps}
	 */
	private _transition?: SlideTransitionProps
	public set transition(value: SlideTransitionProps | undefined) {
		this._transition = value
	}

	public get transition(): SlideTransitionProps | undefined {
		return this._transition
	}

	/**
	 * Stable identity for this slide across saves (MS-PPTX `p14:creationId`)
	 * @type {boolean | number}
	 */
	private _creationId?: boolean | number
	public set creationId(value: boolean | number | undefined) {
		this._creationId = value
	}

	public get creationId(): boolean | number | undefined {
		return this._creationId
	}

	/**
	 * @type {SlideNumberProps}
	 */
	public set slideNumber(value: SlideNumberProps) {
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		this._slideNumberProps = value
		this._setSlideNum(value)
	}

	public get slideNumber(): SlideNumberProps | undefined {
		return this._slideNumberProps
	}

	public get newAutoPagedSlides(): PresSlide[] | undefined {
		return this._newAutoPagedSlides
	}

	/**
	 * Add chart to Slide
	 * @param {CHART_NAME|IChartMulti[]} type - chart type
	 * @param {object[]} data - data object
	 * @param {IChartOpts} options - chart options
	 * @return {Slide} this Slide
	 */
	addChart(type: CHART_NAME | IChartMulti[], data: IOptsChartData[], options?: IChartOpts): Slide {
		// FUTURE: TODO-VERSION-4: Remove first arg - only take data and opts, with "type" required on opts
		genObj.addChartDefinition(this, type, Array.isArray(data) ? data.map(item => ({ ...item })) : data, cloneOpts(options))
		return this
	}

	/**
	 * Add image to Slide
	 * @param {ImageProps} options - image options
	 * @return {Slide} this Slide
	 */
	addImage(options: ImageProps): Slide {
		genObj.addImageDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Add media (audio/video) to Slide
	 * @param {MediaProps} options - media options
	 * @return {Slide} this Slide
	 */
	addMedia(options: MediaProps): Slide {
		genObj.addMediaDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Add a threaded comment to this Slide (MS-PPTX 2.16)
	 * - an author not listed in `pptx.commentAuthors` is added to the author table automatically
	 * - pass `created` (ISO 8601) to keep generated packages byte-reproducible
	 * @param {CommentProps} options - comment props
	 * @return {Slide} this Slide
	 * @example slide.addComment({ text: 'Check this figure', author: 'Ada Lovelace', x: 4, y: 2 })
	 */
	addComment(options: CommentProps): Slide {
		const comment = cloneOpts(options)
		if (typeof comment?.text !== 'string' || !comment.text.trim()) {
			console.warn('[pptxgenjs] addComment: `text` is required - comment ignored')
			return this
		}
		if (typeof comment.author !== 'string' || !comment.author.trim()) {
			console.warn('[pptxgenjs] addComment: `author` is required - comment ignored')
			return this
		}
		this._comments = this._comments ?? []
		this._comments.push(comment)
		return this
	}

	/**
	 * Embed markup PresentationML does not define as its own package part (MS-PPTX 2.2.3)
	 * - `contentType` and `relationshipType` belong to the format being embedded and are required:
	 *   a wrong value produces a package PowerPoint reports as damaged
	 * - set `ink: true` for InkML, which must also supply a `cover` for its picture fallback
	 * @param {ContentPartProps} options - content-part props
	 * @return {Slide} this Slide
	 * @example slide.addContentPart({ data: inkml, contentType: 'application/inkml+xml', relationshipType: '...', ink: true, cover: pngBase64 })
	 */
	addContentPart(options: ContentPartProps): Slide {
		genObj.addContentPartDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Add a zoom to another slide (MS-PPTX 2.2.15)
	 * - the target may be added after this call; it is resolved during export
	 * @param {SlideZoomProps} options - zoom props
	 * @return {Slide} this Slide
	 * @example slide.addZoom({ slideNumber: 4, x: 1, y: 1, w: 3, h: 2 })
	 */
	addZoom(options: SlideZoomProps): Slide {
		genObj.addZoomDefinition(this, 'slide', cloneOpts(options))
		return this
	}

	/**
	 * Add a zoom to a section (MS-PPTX 2.2.15)
	 * @param {SectionZoomProps} options - zoom props
	 * @return {Slide} this Slide
	 * @example slide.addSectionZoom({ sectionTitle: 'Results', x: 1, y: 1, w: 3, h: 2 })
	 */
	addSectionZoom(options: SectionZoomProps): Slide {
		genObj.addZoomDefinition(this, 'section', cloneOpts(options))
		return this
	}

	/**
	 * Add a summary zoom linking several sections (MS-PPTX 2.2.15)
	 * @param {SummaryZoomProps} options - zoom props
	 * @return {Slide} this Slide
	 * @example slide.addSummaryZoom({ sectionTitles: ['Intro', 'Results'], x: 1, y: 1, w: 8, h: 4 })
	 */
	addSummaryZoom(options: SummaryZoomProps): Slide {
		genObj.addZoomDefinition(this, 'summary', cloneOpts(options))
		return this
	}

	/**
	 * Add speaker notes to Slide
	 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
	 * @param {string} notes - notes to add to slide
	 * @return {Slide} this Slide
	 */
	addNotes(notes: string): Slide {
		genObj.addNotesDefinition(this, notes)
		return this
	}

	/**
	 * Add shape to Slide
	 * @param {SHAPE_NAME} shapeName - shape name
	 * @param {ShapeProps} options - shape options
	 * @return {Slide} this Slide
	 */
	addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide {
		// NOTE: As of v3.1.0, <script> users are passing the old shape object from the shapes file (orig to the project)
		// But React/TypeScript users are passing the shapeName from an enum, which is a simple string, so lets cast
		// <script./> => `pptx.shapes.RECTANGLE` [string] "rect" ... shapeName['name'] = 'rect'
		// TypeScript => `pptxgen.shapes.RECTANGLE` [string] "rect" ... shapeName = 'rect'
		// let shapeNameDecode = typeof shapeName === 'object' && shapeName['name'] ? shapeName['name'] : shapeName
		genObj.addShapeDefinition(this, shapeName, cloneOpts(options))
		return this
	}

	/**
	 * Add table to Slide
	 * @param {TableRow[]} tableRows - table rows
	 * @param {TableProps} options - table options
	 * @return {Slide} this Slide
	 */
	addTable(tableRows: TableRow[], options?: TableProps): Slide {
		// FUTURE: we pass `this` - we dont need to pass layouts - they can be read from this!
		this._newAutoPagedSlides = genObj.addTableDefinition(this, tableRows, cloneOpts(options), this._slideLayout, this._presLayout, this.addSlide, this.getSlide)
		return this
	}

	/**
	 * Add text to Slide
	 * @param {string|TextProps[]} text - text string or complex object
	 * @param {TextPropsOptions} options - text options
	 * @return {Slide} this Slide
	 */
	addText(text: string | TextProps[], options?: TextPropsOptions): Slide {
		const textParam = typeof text === 'string' || typeof text === 'number' ? [{ text, options }] : text
		genObj.addTextDefinition(this, textParam, cloneOpts(options), false)
		return this
	}
}
