export interface ColorDef {
	rgb?: string
	theme?: number
	tint?: number
	indexed?: number
	auto?: boolean
}

export interface FontDef {
	sz?: number
	name?: string
	b?: boolean
	i?: boolean
	u?: boolean
	color?: string | ColorDef
}

export interface BorderSideDef {
	style?: string
	color?: string | ColorDef
}

export type BorderSide = string | BorderSideDef | null

export interface BorderSides {
	left?: BorderSide
	right?: BorderSide
	top?: BorderSide
	bottom?: BorderSide
	diagonal?: BorderSide
}

export interface FillDef {
	fgColor?: string
	bgColor?: string
	pattern?: string
}

export interface StyleDef {
	font?: FontDef
	border?: string | BorderSides
	fill?: string | FillDef
}

export interface CellObject {
	value?: string | number | boolean | Date | null
	style?: string
	format?: "date" | "datetime"
}

export type CellValue = string | number | boolean | Date | CellObject | null | undefined

export interface ColObject {
	width?: number | string
	min?: number
	max?: number
	bestFit?: number | boolean
	customWidth?: number | boolean
}

export type ColDef = string | number | ColObject | 0 | false | null | undefined

export interface RowObject {
	hidden?: boolean
	height?: number
	data: CellValue[]
}

export type Row = CellValue[] | RowObject

export interface FreezeDef {
	rows?: number
	cols?: number
}

export interface SheetObject {
	name?: string
	cols?: string | ColDef[]
	data: (Row | null)[]
	freeze?: FreezeDef
}

export type Sheet = CellValue[][] | SheetObject

export interface Workbook {
	styles?: Record<string, StyleDef>
	sheets: (Sheet | null | undefined)[]
}

export interface XlsxFile {
	name: string
	content: string
}

export interface ZipOptions {
	deflate?: (data: Uint8Array) => Uint8Array
	comment?: string
}

export type Callback = (err: Error | null, data: Uint8Array) => void

export function createFiles(workbook: Workbook): XlsxFile[]

export function createXlsx(workbook: Workbook, opts: ZipOptions, next: Callback): void
export function createXlsx(workbook: Workbook, next: Callback): void
export function createXlsx(workbook: Workbook, opts?: ZipOptions): Promise<Uint8Array>
