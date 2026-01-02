

;((exports, Object) => {
	var UNDEF
	, createZip = exports.createZip || require('@litejs/zip').createZip
	, createFiles = workbook => {
		var xmlHead = '<?xml version="1.0" encoding="UTF-8"?>'
		, nsPackage = 'http://schemas.openxmlformats.org/package/2006/'
		, nsRels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
		// Excel's epoch is January 1, 1900 (with a bug treating 1900 as leap year)
		, excelEpoch = new Date(1899, 11, 30)
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
		, toXml = (name, attrs, childs, wrapVal) => (
			attrs = mapEntries(attrs, a => a[1] != null ? a[0] + '="' + a[1] + '"' : '', ' '),
			childs = mapEntries(childs, a => a[1].map(wrapVal ? b => toXml(a[0], 0, toVal(b, wrapVal)) : b => toXml(a[0], b)).join(''), ''),
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
		, border = [
			{}
		]
		, xf = [
			{ fontId: 0, applyFont: 1 },
			{ numFmtId: 164, applyNumberFormat: 1 },
			{ numFmtId: 165, applyNumberFormat: 1 },
			{ numFmtId: 0, fontId: 1, applyFont: 1 },
		]
		, styles = Object.fromEntries(Object.entries(workbook.styles||{}).map(a => {
			var newBorder = a[1].border
			if (isStr(newBorder)) newBorder = { left: newBorder, right: newBorder, top: newBorder, bottom: newBorder }
			a[1] = xf.push({
				fontId: a[1].font ? font.push(a[1].font) - 1 : 0,
				borderId: newBorder ? border.push(newBorder) - 1 : UNDEF,
				applyBorder: newBorder ? 1 : UNDEF,
			}) - 1
			return a
		}))
		, getXf = val => {
			var style = val.style
			, format = val.format
			, attr = isNum(styles[style]) ? styles[style] : style === 'bold' ? 3 : format === 'date' ? 1 : format === 'datetime' ? 2 : 0
			return attr ? '" s="' + attr : ''
		}
		, files = workbook.sheets.filter(isTruthy).map(
			(sheet, i) => {
				i++
				sheet = dataArr(sheet)
				var cols = sheet.cols
				, rowIndex = 0
				, name = 'worksheets/sheet' + i + '.xml'

				types.push({ PartName: '/xl/' + name, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' })
				rels.push({ Id: 'rId' + i, Type: nsRels + 'worksheet', Target: name })
				sheets += toXml('sheet', { name: sheet.name || 'Sheet' + i, sheetId: i, 'r:id': 'rId' + i })

				return {
					name: 'xl/' + name,
					content: xmlHead +
						'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
						(sheet.data[0] ? '<dimension ref="A1:' + toCol(sheet.data[0].length - 1) + sheet.data.length + '"/>' : '') +
						(cols ? toXml('cols', 0, { col: (isStr(cols) ? cols.split(',') : cols).map(
							(w, col) => w ? assign({ min: col + 1, max: col + 1 }, isStr(w) ? { width: w, customWidth: 1 } : w) : 0
						).filter(isTruthy)}) : '') +
						'<sheetData>' + sheet.data.map(
							row => (row = dataArr(row)) ? toXml('row', {
								r: ++rowIndex,
								hidden: row.hidden ? 1 : UNDEF,
								ht: row.height,
								customHeight: row.height ? 1 : UNDEF,
							}, row.data.map(
								(val, col, tmp) => val != null ? '<c r="' + toCol(col) + rowIndex + (
									isObj(val) ? (tmp = getXf(val), val = val.value, tmp) : ''
								) + (
									val && isStr(val) ? (
										val[0] === '=' ? '"><f>' + val.slice(1) + '</f>' :
										'" t="inlineStr"><is><t>' + val.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</t></is>'
									) :
									isNum(val) ? '"><v>' + val + '</v>' :
									typeof val === 'boolean' ? '" t="b"><v>' + (val ? 1 : 0) + '</v>' :
									val instanceof Date ? (isArr(tmp) ? '" s="2' : '') +'"><v>' + ((val - excelEpoch)/(24 * 60 * 60 * 1000)).toFixed(6) + '</v>' :
									'">'
								) + '</c>' : ''
							).join('')) : ''
						).join('') + '</sheetData></worksheet>'
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
				'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
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


