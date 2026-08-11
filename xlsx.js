

;((exports, Object) => {
	var UNDEF
	, createZip = exports.createZip || require('@litejs/zip').createZip
	, createFiles = workbook => {
		var xmlHead = '<?xml version="1.0" encoding="UTF-8"?>'
		, nsPackage = 'http://schemas.openxmlformats.org/package/2006/'
		, nsRels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
		// Excel's epoch is January 1, 1900 (with a bug treating 1900 as leap year)
		, excelEpoch = Date.UTC(1899, 11, 30)
		, types = [
			{ PartName: '/xl/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' },
			{ PartName: '/xl/workbook.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' }
		]
		, rels = [{ Id: 'rId0', Type: nsRels + 'styles', Target: 'styles.xml' }]
		, relsFile = (name, Relationship) => ({
			name,
			content: xmlHead + toXml('Relationships', { xmlns: nsPackage + 'relationships' }, { Relationship })
		})
		, sheets = ''
		, assign = Object.assign
		, isArr = Array.isArray
		, dataArr = arr => isArr(arr) ? { data: arr } : arr
		, isNum = num => num === num && typeof num === 'number'
		, isObj = obj => !!obj && obj.constructor === Object
		, isStr = str => typeof str === 'string'
		, isTruthy = s => s
		, mapEntries = (obj, fn, separator) => !obj ? '' : isStr(obj) ? obj : Object.entries(obj).map(fn).filter(isTruthy).join(separator)
		, toCol = num => (num > 25 ? toCol((0 | num / 26) - 1) : '') + String.fromCharCode(65 + num % 26)
		, toVal = (b, key) => Object.entries(b).reduce((accum, arr) => (accum[arr[0]] = [arr[1] === true ? {} : { [key]: arr[1] }], accum), {})
		, esc = val => ('' + val).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
		, toXml = (name, attrs, childs, wrapVal) => (
			attrs = mapEntries(attrs, a => a[1] != null ? a[0] + '="' + esc(a[1]) + '"' : '', ' '),
			childs = mapEntries(childs, a => a[1] && a[1].map(wrapVal ? b => toXml(a[0], 0, toVal(b, wrapVal)) : b => toXml(a[0], b)).join(''), ''),
			'<' + (attrs ? name + ' ' + attrs : name) + (childs ? '>' + childs + '</' + name + '>' : '/>')
		)
		, numFmt = [
			{ numFmtId: 164, formatCode: 'yyyy-mm-dd' },
			{ numFmtId: 165, formatCode: 'yyyy-mm-dd hh:mm:ss' },
		]
		, font = [
			{ sz: 11, name: 'Calibri' },
			{ sz: 11, name: 'Calibri', b: true },
		]
		, fill = [
			{ pattern: 'none' },
			{ pattern: 'gray125' },
		]
		, border = [
			{}
		]
		, xf = [
			{ fontId: 0, applyFont: 1 },
			{ numFmtId: 164, applyNumberFormat: 1 },
			{ numFmtId: 165, applyNumberFormat: 1 },
			{ numFmtId: 0, fontId: 1, applyFont: 1 },
		]
		, styles = Object.entries(workbook.styles||{}).reduce((accum, a) => {
			var newBorder = a[1].border
			, newFill = a[1].fill
			, newFont = a[1].font
			if (isStr(newBorder)) newBorder = { left: newBorder, right: newBorder, top: newBorder, bottom: newBorder }
			if (isStr(newFill)) newFill = { fgColor: newFill }
			if (newFill) newFill.pattern = newFill.pattern || 'solid'
			a[1] = xf.push({
				fontId: newFont ? font.push(newFont) - 1 : 0,
				applyFont: newFont ? 1 : UNDEF,
				borderId: newBorder ? border.push(newBorder) - 1 : UNDEF,
				applyBorder: newBorder ? 1 : UNDEF,
				fillId: newFill ? fill.push(newFill) - 1 : UNDEF,
				applyFill: newFill ? 1 : UNDEF,
			}) - 1
			accum[a[0]] = a[1]
			return accum
		}, {})
		, xfCache = {}
		, mergeXf = (styleId, fmtId) => xfCache[styleId + '.' + fmtId] ||
			(xfCache[styleId + '.' + fmtId] = xf.push(assign({}, xf[styleId], {
				numFmtId: xf[fmtId].numFmtId,
				applyNumberFormat: 1
			})) - 1)
		, getXf = (val, isDate) => {
			var style = val.style
			, format = val.format
			, styleId = isNum(styles[style]) ? styles[style] : style === 'bold' ? 3 : 0
			, fmtId = format === 'date' ? 1 : format === 'datetime' || isDate ? 2 : 0
			, attr = fmtId && styleId ? mergeXf(styleId, fmtId) : fmtId || styleId
			return attr ? '" s="' + attr : ''
		}
		, files = workbook.sheets.filter(isTruthy).map(
			(sheet, i) => {
				i++
				sheet = dataArr(sheet)
				var cols = sheet.cols
				, width = 0
				, rowIndex = 0
				, freeze = sheet.freeze
				, freezeRows = freeze && freeze.rows || 0
				, freezeCols = freeze && freeze.cols || 0
				, freezePane = freeze && (freezeRows ? 'bottom' : 'top') + (freezeCols ? 'Right' : 'Left')
				, name = 'worksheets/sheet' + i + '.xml'
				, sheetData = sheet.data.map(
					row => (++rowIndex, row = dataArr(row)) ? toXml('row', {
						r: rowIndex,
						hidden: row.hidden ? 1 : UNDEF,
						ht: row.height,
						customHeight: row.height ? 1 : UNDEF,
					}, (row = row.data, row.length > width && (width = row.length), row).map(
						(val, col, tmp) => val != null ? '<c r="' + toCol(col) + rowIndex + (
							isObj(val) ? (tmp = val, val = val.value, tmp = getXf(tmp, val instanceof Date)) : (tmp = '')
						) + (
							val && isStr(val) ? (
								val[0] === '=' ? '"><f>' + esc(val.slice(1)) + '</f>' :
								'" t="inlineStr"><is><t' + (/^\s|\s$/.test(val = esc(val)) ? ' xml:space="preserve"' : '') + '>' + val + '</t></is>'
							) :
							val !== val || isNum(val) ? (isFinite(val) ? '"><v>' + val + '</v>' : '" t="e"><v>#NUM!</v>') :
							typeof val === 'boolean' ? '" t="b"><v>' + (val ? 1 : 0) + '</v>' :
							val instanceof Date ? (tmp ? '' : '" s="2') + '"><v>' + ((val - excelEpoch)/(24 * 60 * 60 * 1000)).toFixed(6) + '</v>' :
							'">'
						) + '</c>' : ''
					).join('')) : ''
				).join('')

				types.push({ PartName: '/xl/' + name, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' })
				rels.push({ Id: 'rId' + i, Type: nsRels + 'worksheet', Target: name })
				sheets += toXml('sheet', { name: sheet.name || 'Sheet' + i, sheetId: i, 'r:id': 'rId' + i })

				return {
					name: 'xl/' + name,
					content: xmlHead +
						'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
						(width ? '<dimension ref="A1:' + toCol(width - 1) + sheet.data.length + '"/>' : '') +
						(freeze ?
						'<sheetViews>' + toXml('sheetView', { workbookViewId: 0 }, {
							pane: [{
								xSplit: freezeCols || UNDEF,
								ySplit: freezeRows || UNDEF,
								topLeftCell: toCol(freezeCols) + (freezeRows + 1),
								activePane: freezePane,
								state: 'frozen'
							}],
							selection: [{ pane: freezePane }]}) + '</sheetViews>' : '') +
						(cols ? toXml('cols', 0, { col: (isStr(cols) ? cols.split(',') : cols).map(
							(w, col) => w ? assign({ min: col + 1, max: col + 1 }, isStr(w) ? { width: w, customWidth: 1 } : w) : 0
						).filter(isTruthy)}) : '') +
						'<sheetData>' + sheetData + '</sheetData></worksheet>'
				}
			}
		)
		files.unshift(
			{
				name: '[Content_Types].xml',
				content: xmlHead + toXml('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' }, {
					Default: [
						{ Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' },
						{ Extension: 'xml', ContentType: 'application/xml' }
					],
					Override: types
				})
			},
			relsFile('_rels/.rels', [{ Id: 'rId1', Type: nsRels + 'officeDocument', Target: 'xl/workbook.xml' }]),
			relsFile('xl/_rels/workbook.xml.rels', rels),
			{
				name: 'xl/styles.xml',
				content: xmlHead + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
				toXml('numFmts', { count: numFmt.length }, { numFmt }) +
				toXml('fonts', { count: font.length }, { font }, 'val') +
				toXml('fills', { count: fill.length }, fill.map(
					f => '<fill>' + toXml('patternFill', { patternType: f.pattern }, {
						fgColor: f.fgColor ? [{ rgb: f.fgColor }] : UNDEF,
						bgColor: f.bgColor ? [{ rgb: f.bgColor }] : UNDEF,
					}) + '</fill>'
				).join('')) +
				toXml('borders', { count: border.length }, { border }, 'style') +
				toXml('cellXfs', { count: xf.length }, { xf }) +
				'</styleSheet>'
			},
			{
				name: 'xl/workbook.xml',
				content: xmlHead + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets + '</sheets></workbook>'
			}
		)
		return files
	}


	exports.createFiles = createFiles
	exports.createXlsx = (workbook, opts, next) => createZip(createFiles(workbook), opts, next)

// this is `exports` in module and `window` in browser
})(this, Object) // jshint ignore:line

