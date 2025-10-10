

var createZip = require("@litejs/zip").createZip
, createFiles = workbook => {
    var xmlHead = '<?xml version="1.0" encoding="UTF-8"?>'
	, nsPackage = 'http://schemas.openxmlformats.org/package/2006/'
	, nsRels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
	// Excel's epoch is January 1, 1900 (with a bug treating 1900 as leap year)
	, excelEpoch = new Date(1899, 11, 30)
	, types = [
		{ PartName: "/styles.xml", ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' },
		{ PartName: "/workbook.xml", ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' }
	]
	, rels = [{ Id: 'rId0', Type: nsRels + 'styles', Target: 'styles.xml' }]
	, relsFile = (name, Relationship) => ({
		name,
		content: xmlHead + toXml('Relationships', { xmlns: nsPackage + 'relationships' }, { Relationship })
	})
	, sheets = ''
	, isObj = obj => !!obj && obj.constructor === Object
	, isStr = str => typeof str === "string"
	, mapEntries = (obj, fn, separator) => obj && Object.entries(obj).map(fn).join(separator)
	, toCol = num => (num > 25 ? toCol((0 | num / 26) - 1) : '') + String.fromCharCode(65 + num % 26)
	, toXml = (name, attrs, childs) => (
		attrs = mapEntries(attrs, a => a[0] + '="' + a[1] + '"', ' '),
		childs = mapEntries(childs, a => a[1].map(b => toXml(a[0], b)).join(''), ''),
		'<' + (attrs ? name + ' ' + attrs : name) + (childs ? '>' + childs + '</' + name + '>' : '/>')
	)
	, files = workbook.sheets.filter(s => s).map(
		(sheet, i, name) => {
			sheet = Array.isArray(sheet)? { data: sheet } : sheet
			i++
			name = 'sheet' + i + '.xml'
			types.push({ PartName: name, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' })
			rels.push({ Id: 'rId' + i, Type: nsRels + 'worksheet', Target: name })
			sheets += '<sheet name="' + (sheet.name || 'Sheet' + i) + '" sheetId="' + i + '" r:id="rId' + i + '"/>'
			var cols = sheet.cols
			if (cols) cols = (isStr(cols) ? cols.split(",") : cols).map(
				(w, i) => w ? toXml('col', { min: (i + 1), max: (i + 1), ...(isStr(w) ? {width:w, customWidth:1} : w)}) : ''
			).join('')
			, rowIndex = 0

			return {
				name,
				content: xmlHead +
					'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
					(sheet.data[0] ? '<dimension ref="A1:' + toCol(sheet.data[0].length) + sheet.data.length + '"/>' : '') +
					(cols ? '<cols>' + cols + '</cols>' : '') +
					'<sheetData>' + sheet.data.map(
						row => row ? '<row r="' + (++rowIndex) + '">' + row.map(
							(val, col, tmp) => '<c r="' + toCol(col) + rowIndex + (
								isObj(val) ? (tmp = val.style === 'bold' ? '" s="2' : '', val = val.value, tmp) : ''
							) + (
								val && isStr(val) ? (
									val[0] === "=" ? '"><f>' + val.slice(1) + '</f>' :
									'" t="inlineStr"><is><t>' + val.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</t></is>'
								) :
								typeof val === "number" ? '"><v>' + val + '</v>' :
								typeof val === "boolean" ? '" t="b"><v>' + (val ? 1 : 0) + '</v>' :
								val instanceof Date ? '" s="1"><v>' + ((val - excelEpoch)/(24 * 60 * 60 * 1000)).toFixed(6) + '</v>' :
								'">'
							) + '</c>'
						).join('') + '</row>' : ''
					).join('') + `</sheetData></worksheet>`
			}
		}
	)
	files.unshift(
		{
			name: '[Content_Types].xml',
			content: xmlHead + toXml('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' }, {
				Default: [
					{ Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" },
					{ Extension: "xml", ContentType: "application/xml" }
				],
				Override: types
			})
		},
		relsFile('_rels/.rels', [{ Id: 'rId1', Type: nsRels + 'officeDocument', Target: 'workbook.xml' }]),
		relsFile('_rels/workbook.xml.rels', rels),
		{
			name: 'styles.xml',
			content: xmlHead + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><!-- Date format --><numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/><!-- DateTime format --></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Calibri"/><b/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="3"><xf fontId="0" applyFont="1"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" applyFont="1"/></cellXfs></styleSheet>'
		},
		{
			name: 'workbook.xml',
			content: xmlHead + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets + '</sheets></workbook>'
		}
	)
	//console.log(files)
	return files
}


exports.createFiles = createFiles
exports.createXlsx = (workbook, opts, next) => createZip(createFiles(workbook), opts, next)






