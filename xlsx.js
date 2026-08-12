

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
		, dataArr = arr => Array.isArray(arr) ? { data: arr } : arr
		, isNum = num => num === num && typeof num === 'number'
		, isObj = obj => !!obj && obj.constructor === Object
		, isStr = str => typeof str === 'string'
		, isTruthy = s => s
		, mapEntries = (obj, fn, sep) => !obj ? '' : isStr(obj) ? obj : Object.entries(obj).map(fn).filter(isTruthy).join(sep)
		, normalizeRgb = rgb => isStr(rgb) ? (rgb = rgb.replace(/^#/, '').toUpperCase(), rgb.length === 6 ? 'FF' + rgb : rgb) : rgb
		, toCol = num => (num > 25 ? toCol((0 | num / 26) - 1) : '') + String.fromCharCode(65 + num % 26)
		, toColor = c => isStr(c) ? { rgb: normalizeRgb(c) } : assign({}, c, { rgb: normalizeRgb(c.rgb) })
		, esc = val => ('' + val).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
		, toXml = (name, attrs, childs) => (
			attrs = mapEntries(attrs, a => a[1] != null ? a[0] + '="' + esc(a[1]) + '"' : '', ' '),
			childs = mapEntries(childs, a => a[1] && a[1].map(b => toXml(a[0], b)).join(''), ''),
			'<' + (attrs ? name + ' ' + attrs : name) + (childs ? '>' + childs + '</' + name + '>' : '/>')
		)
		, toOrderXml = (name, child, arr, order, fn) => toXml(name, { count: arr.length }, arr.map(
			row => toXml(child, 0, order.map(key => row[key] === UNDEF ? '' : fn(key, row[key])).join(''))).join('')
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
			, newFont = a[1].font && assign({ sz: 11, name: 'Calibri' }, a[1].font)
			if (newFont && newFont.color) newFont.color = toColor(newFont.color)
			if (isStr(newBorder)) newBorder = { left: newBorder, right: newBorder, top: newBorder, bottom: newBorder }
			if (isStr(newFill)) newFill = { fgColor: newFill }
			if (newFill) newFill = {
				pattern: newFill.pattern || 'solid',
				fgColor: normalizeRgb(newFill.fgColor),
				bgColor: normalizeRgb(newFill.bgColor)
			}
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
							val instanceof Date ? (tmp ? '' : '" s="2') + '"><v>' + (+((val - excelEpoch)/(24 * 60 * 60 * 1000)).toFixed(11)) + '</v>' :
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
				toOrderXml(
					'fonts', 'font', font,
					['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'],
					(k, v) => toXml(k, v === true ? 0 : isObj(v) ? v : { val: v })
				) +
				toXml('fills', { count: fill.length }, fill.map(
					f => '<fill>' + toXml('patternFill', { patternType: f.pattern }, {
						fgColor: f.fgColor ? [{ rgb: f.fgColor }] : UNDEF,
						bgColor: f.bgColor ? [{ rgb: f.bgColor }] : UNDEF,
					}) + '</fill>'
				).join('')) +
				toOrderXml(
					'borders', 'border', border,
					['left', 'right', 'top', 'bottom', 'diagonal'],
					(k, v) => toXml(k, { style: v && v.style || v }, v && v.color ? { color: [toColor(v.color)] } : 0)
				) +
				'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
				toXml('cellXfs', { count: xf.length }, { xf }) +
				'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
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

