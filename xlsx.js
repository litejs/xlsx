

var createZip = require("@litejs/zip").createZip
, createFiles = workbook => {
    var xmlHead = '<?xml version="1.0" encoding="UTF-8"?>'
	// Excel's epoch is January 1, 1900 (with a bug treating 1900 as leap year)
	, excelEpoch = new Date(1899, 11, 30)
	, types = ''
	, relations = ''
	, sheets = ''
	, isObj = obj => !!obj && obj.constructor === Object
	, toCol = num => (num > 25 ? toCol((0 | num / 26) - 1) : '') + String.fromCharCode(65 + num % 26)
	, files = workbook.sheets.map(
		(sheet, i, name) => {
			sheet = Array.isArray(sheet)? { data: sheet } : sheet
			i++
			name = 'sheet' + i + '.xml'
			types += '<Override PartName="' + name + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
			relations += '<Relationship Id="rId' + i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="' + name + '"/>'
			sheets += '<sheet name="' + (sheet.name || 'Sheet' + i) + '" sheetId="' + i + '" r:id="rId' + i + '"/>'
			var cols = sheet.cols ? sheet.cols.split(",").map(
				(w, i) => w ? '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>' : ''
			).join('') : ''
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
								typeof val === "string" && val[0] ? (
									val[0] === "=" ? '"><f>' + val.slice(1) + '</f>' :
									'" t="inlineStr"><is><t>' + val.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</t></is>'
								) :
								typeof val === "number" ? '"><v>' + val + '</v>' :
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
			content: xmlHead + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + types + '</Types>'
		},
		{
			name: '_rels/.rels',
			content: xmlHead + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="workbook.xml"/></Relationships>'
		},
		{
			name: '_rels/workbook.xml.rels',
			content: xmlHead + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + relations + '</Relationships>'
		},
		// Add this to your file list
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


exports.createXlsx = (workbook, opts, next) => createZip(createFiles(workbook), opts, next)






