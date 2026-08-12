
describe("xlsx", function() {
	require("@litejs/cli/snapshot.js")
	var { createFiles, createXlsx } = require("..")
	, compressionSuported = typeof CompressionStream !== "undefined" && typeof Response !== "undefined"

	function sheet1(data, sheet) {
		return createFiles({ sheets: [{ data, ...sheet }] }).find(f => f.name === 'xl/worksheets/sheet1.xml').content
	}

	test("Readme", function(assert, mock) {
		mock.swap(Date, "now", mock.fn(1514900750001))
		var workbook = {
			sheets: [
				{
					name: 'Products',
					cols: [{width:20,bestFit:1,customWidth:1},0,'15'],
					data: [
						['Apple', 1.99, 10],
						['Banana', 0.99, 15],
						['Orange', 2.49, 8],
						null,
						['Totals', '=SUM(B1:B3)', {style: 'bold', value: '=SUM(C1:C3)'}]
					]
				},
				null,
				[
					['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA']
				],
				{
					name: 'Empty Sheet',
					data: []
				},
				{
					name: 'Types',
					cols: '20,40',
					freeze: { rows: 1, cols: 0 },
					data: [
						['null', null],
						['true', true],
						['false', false],
						['Empty string', ''],
						['Empty object', {}],
						['Object as value', { value: {} }],
						['Empty array', {}],
						['Default Date', new Date(1514900750001)],
						['Datetime', { format: 'datetime', value: new Date(0) }],
						['Date', { format: 'date', value: new Date(1514900750001) }],
					]
				},
			]
		}
		assert.matchSnapshot("test/snap/readme.json", JSON.stringify(createFiles(workbook), null, 2))
		if (!compressionSuported) return assert.end()
		createXlsx(workbook)
		.then(uint8 => {
			assert.matchSnapshot("test/snap/readme.xlsx", uint8)
			assert.end()
		})
	})
	test("styles", function(assert, mock) {
		mock.swap(Date, "now", mock.fn(1514900750001))
		var workbook = {
			styles: {
				My1: {
					font: { sz: 15, name: "Calibri" },
				},
				Plain: {},
				Border1: {
					border: 'thin',
				},
				Border2: {
					border: { top: 'double' },
				},
				Border3: {
					border: { right: { style: 'medium', color: 'CC0000' }, bottom: { style: 'thin', color: { theme: 4 } } },
				},
				Fill1: {
					fill: 'FFFF00',
				},
				Fill2: {
					fill: { bgColor: 'FF9900', pattern: 'solid' },
				},
				Color1: {
					font: { b: true, color: 'CC0000' },
				},
				Color2: {
					font: { color: { theme: 1, tint: -0.25 } },
				}
			},
			sheets: [
				{
					cols: [{ width: null }],
					name: 'Styles',
					freeze: { rows: 0, cols: 1 },
					data: [
						[{style: 'My1', value: 'Apple My1'}, { style: "Plain", value: "Banana Plain" }],
						{ hidden: true, data: ['Hidden Row', 1] },
						{ height: 25, data: ['Sized Row', { style: 'Border1', value: 1 }] },
						[{ style: 'Fill1', value: 'Filled' }, { style: 'Fill2', value: 2 }],
						[{ style: 'Color1', value: 'Red bold' }, { style: 'Color2', value: 'Themed' }],
						[{ style: 'Border3', value: 'Colored border' }],
					]
				},
			]
		}
		assert.matchSnapshot("test/snap/styles.json", JSON.stringify(createFiles(workbook), null, 2))
		if (!compressionSuported) return assert.end()
		createXlsx(workbook)
		.then(uint8 => {
			assert.matchSnapshot("test/snap/styles.xlsx", uint8)
			assert.end()
		})
	})
	test("xml escaping and whitespace preservation", function(assert) {
		var files = createFiles({
			sheets: [
				{ name: 'SV11 & SV12', data: [] },
				{ name: 'A < B', data: [] },
				{ name: 'A "B"', data: [
					['=IF(A1<5,"yes","no")'],
					['  padded  ', 'plain', '\ttabbed', 'has inner space'],
					['cdata ]]> end', 'a > b']
				] },
			]
		})
		var workbook = files.find(function(f) { return f.name === 'xl/workbook.xml' }).content
		var sheet = files.find(function(f) { return f.name === 'xl/worksheets/sheet3.xml' }).content
		assert.ok(workbook.indexOf('name="SV11 &amp; SV12"') > -1, 'ampersand escaped')
		assert.ok(workbook.indexOf('name="A &lt; B"') > -1, 'less-than escaped')
		assert.ok(workbook.indexOf('name="A &quot;B&quot;"') > -1, 'double-quote escaped')
		assert.ok(sheet.indexOf('<f>IF(A1&lt;5,&quot;yes&quot;,&quot;no&quot;)</f>') > -1, 'formula escaped')
		assert.ok(sheet.indexOf('<t xml:space="preserve">  padded  </t>') > -1, 'surrounding spaces preserved')
		assert.ok(sheet.indexOf('<t xml:space="preserve">\ttabbed</t>') > -1, 'leading tab preserved')
		assert.ok(sheet.indexOf('<t>plain</t>') > -1, 'no xml:space when not needed')
		assert.ok(sheet.indexOf('<t>has inner space</t>') > -1, 'inner space needs no xml:space')
		assert.ok(sheet.indexOf('<t>cdata ]]&gt; end</t>') > -1, 'CDATA-close sequence escaped')
		assert.ok(sheet.indexOf('<t>a &gt; b</t>') > -1, 'greater-than escaped')
		assert.end()
	})
	test("strips XML-illegal control characters", function(assert) {
		var files = createFiles({
			sheets: [
				{ name: 'Tab\u0007Bell', data: [
					['a\u0000b\u0001c\u001Fd'],
					['keep\ttab\nand newline']
				] }
			]
		})
		var workbook = files.find(function(f) { return f.name === 'xl/workbook.xml' }).content
		var sheet = files.find(function(f) { return f.name === 'xl/worksheets/sheet1.xml' }).content
		assert.ok(sheet.indexOf('<t>abcd</t>') > -1, 'control chars removed from cell text')
		assert.ok(sheet.indexOf('<t>keep\ttab\nand newline</t>') > -1, 'tab and newline preserved')
		assert.ok(workbook.indexOf('name="TabBell"') > -1, 'control chars removed from attributes')
		assert.notOk(/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(sheet + workbook), 'no illegal chars anywhere in output')
		assert.end()
	})
	test("cells resolve to a complete cellXfs entry", function(assert) {
		var date = new Date(1514900750001)
		// resolve the cellXfs entry that cell A1 actually points at
		function cellXf(cell) {
			var files = createFiles({
				styles: { My1: { font: { sz: 15, name: 'Calibri' } } },
				sheets: [{ data: [[cell]] }]
			})
			var sheet = files.find(function(f) { return f.name === 'xl/worksheets/sheet1.xml' }).content
			var styles = files.find(function(f) { return f.name === 'xl/styles.xml' }).content
			var s = /<c r="A1"(?: s="(\d+)")?/.exec(sheet)[1] || '0'
			return /<cellXfs[^>]*>([\s\S]*)<\/cellXfs>/.exec(styles)[1].match(/<xf[^>]*\/>/g)[+s]
		}
		var styled = cellXf({ style: 'My1', value: date })
		var plain = cellXf({ style: 'My1', value: 'x' })
		assert.ok(cellXf({ value: date }).indexOf('numFmtId="165"') > -1, 'wrapped Date gets datetime format')
		assert.ok(cellXf(date).indexOf('numFmtId="165"') > -1, 'bare Date gets datetime format')
		assert.ok(styled.indexOf('numFmtId="165"') > -1, 'styled Date keeps datetime format')
		assert.ok(styled.indexOf('fontId="2"') > -1, 'styled Date keeps the custom font')
		assert.ok(cellXf({ style: 'My1', format: 'date', value: date }).indexOf('numFmtId="164"') > -1, 'format applies alongside style')
		assert.ok(plain.indexOf('applyFont="1"') > -1, 'custom font is flagged as applied')
		assert.ok(styled.indexOf('applyFont="1"') > -1, 'merged xf keeps the applyFont flag')
		assert.end()
	})
	test("non-finite numbers become error cells", function(assert) {
		var sheet = sheet1([[1.5, Infinity, -Infinity, NaN, 1e21]])

		assert.ok(sheet.indexOf('<c r="A1"><v>1.5</v></c>') > -1, 'finite number unchanged')
		assert.ok(sheet.indexOf('<c r="B1" t="e"><v>#NUM!</v></c>') > -1, 'Infinity becomes #NUM!')
		assert.ok(sheet.indexOf('<c r="C1" t="e"><v>#NUM!</v></c>') > -1, '-Infinity becomes #NUM!')
		assert.ok(sheet.indexOf('<c r="D1" t="e"><v>#NUM!</v></c>') > -1, 'NaN becomes #NUM!')
		assert.ok(sheet.indexOf('<c r="E1"><v>1e+21</v></c>') > -1, 'large exponent stays numeric')
		assert.end()
	})
	test("border sides follow schema order", function(assert) {
		function sides(def) {
			var styles = createFiles({
				styles: { B: { border: def } },
				sheets: [{ data: [[{ style: 'B', value: 1 }]] }]
			}).find(function(f) { return f.name === 'xl/styles.xml' }).content
			return /<borders[^>]*>([\s\S]*)<\/borders>/.exec(styles)[1]
			.match(/<border>[\s\S]*?<\/border>|<border\/>/g)[1]
			.match(/<(left|right|top|bottom|diagonal)\b/g)
		}
		assert.equal(sides({ bottom: 'thin', top: 'double', left: 'thin' }), ['<left', '<top', '<bottom'], 'object form reordered to left, top, bottom')
		assert.equal(sides('thin'), ['<left', '<right', '<top', '<bottom'], 'string form stays ordered')
		assert.equal(sides({ bottom: 'thin', top: null }), ['<top', '<bottom'], 'sides with no style keep their position')
		assert.end()
	})
	test("fill colors are 8 digit ARGB", function(assert) {
		function fills(def) {
			return /<fills[^>]*>([\s\S]*?)<\/fills>/.exec(createFiles({
				styles: { F: { fill: def } },
				sheets: [{ data: [[{ style: 'F', value: 1 }]] }]
			}).find(function(f) { return f.name === 'xl/styles.xml' }).content)[1]
		}
		assert.ok(fills('FFFF00').indexOf('<fgColor rgb="FFFFFF00"/>') > -1, 'six digits gain an FF alpha')
		assert.ok(fills('#ffff00').indexOf('<fgColor rgb="FFFFFF00"/>') > -1, 'leading hash dropped, case raised')
		assert.ok(fills('FFFFFF00').indexOf('<fgColor rgb="FFFFFF00"/>') > -1, 'eight digits pass through')
		assert.ok(fills({ bgColor: 'ff9900' }).indexOf('<bgColor rgb="FFFF9900"/>') > -1, 'bgColor normalized too')
		var caller = { fgColor: 'FFFF00' }
		createFiles({ styles: { F: { fill: caller } }, sheets: [{ data: [[1]] }] })
		assert.equal(caller, { fgColor: 'FFFF00' }, 'caller style object left untouched')
		assert.end()
	})
	test("custom fonts keep a size and family", function(assert) {
		function fonts(def) {
			return /<fonts[^>]*>([\s\S]*?)<\/fonts>/.exec(createFiles({
				styles: { F: { font: def } },
				sheets: [{ data: [[{ style: 'F', value: 1 }]] }]
			}).find(function(f) { return f.name === 'xl/styles.xml' }).content)[1]
			.match(/<font>[\s\S]*?<\/font>/g).pop()
		}
		assert.equal(fonts({ b: true }), '<font><b/><sz val="11"/><name val="Calibri"/></font>', 'defaults fill in around a bare bold')
		assert.equal(fonts({ sz: 15, name: 'Arial' }), '<font><sz val="15"/><name val="Arial"/></font>', 'caller values win')
		// CT_Font is a sequence, b/i/u come before sz and name
		assert.equal(fonts({ name: 'Arial', u: true, sz: 14, i: true }), '<font><i/><u/><sz val="14"/><name val="Arial"/></font>', 'children reordered to schema order')
		var caller = { b: true }
		createFiles({ styles: { F: { font: caller } }, sheets: [{ data: [[1]] }] })
		assert.equal(caller, { b: true }, 'caller font object left untouched')
		assert.end()
	})
	test("border side color is a child element", function(assert) {
		function border(def) {
			return /<borders[^>]*>[\s\S]*?<\/borders>/.exec(createFiles({
				styles: { B: { border: def } },
				sheets: [{ data: [[{ style: 'B', value: 1 }]] }]
			}).find(function(f) { return f.name === 'xl/styles.xml' }).content)[0]
			.match(/<border>[\s\S]*?<\/border>|<border\/>/g).pop()
		}
		assert.equal(border({ left: { style: 'thin', color: 'FF0000' } }), '<border><left style="thin"><color rgb="FFFF0000"/></left></border>', 'color is a child of the side, not an attribute')
		assert.equal(border({ diagonal: { style: 'thin', color: { theme: 2 } } }), '<border><diagonal style="thin"><color theme="2"/></diagonal></border>', 'CT_Color attributes carried through')
		assert.equal(border({ left: { style: 'thin' } }), '<border><left style="thin"/></border>', 'object side without a color')
		assert.equal(border('thin'), '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>', 'string form unchanged')
		assert.end()
	})
	test("font color is a CT_Color, never a val attribute", function(assert) {
		function color(def) {
			return /<color[^>]*\/>/.exec(/<fonts[^>]*>([\s\S]*?)<\/fonts>/.exec(createFiles({
				styles: { F: { font: { color: def } } },
				sheets: [{ data: [[{ style: 'F', value: 1 }]] }]
			}).find(function(f) { return f.name === 'xl/styles.xml' }).content)[1])[0]
		}
		assert.equal(color('FF0000'), '<color rgb="FFFF0000"/>', 'string color becomes an ARGB rgb')
		assert.equal(color('#ff0000'), '<color rgb="FFFF0000"/>', 'hash dropped, case raised')
		assert.equal(color('FFFF0000'), '<color rgb="FFFF0000"/>', 'eight digits pass through')
		assert.equal(color({ rgb: 'FF0000' }), '<color rgb="FFFF0000"/>', 'rgb inside an object normalized too')
		assert.equal(color({ theme: 1, tint: -0.25 }), '<color theme="1" tint="-0.25"/>', 'other CT_Color attributes kept')
		// color sits between sz and name in the CT_Font sequence
		var font = /<font><sz[\s\S]*?<\/font>/.exec(createFiles({
			styles: { F: { font: { name: 'Arial', color: 'FF0000', sz: 12 } } },
			sheets: [{ data: [[{ style: 'F', value: 1 }]] }]
		}).find(function(f) { return f.name === 'xl/styles.xml' }).content.match(/<font>[\s\S]*?<\/font>/g).pop())[0]
		assert.equal(font, '<font><sz val="12"/><color rgb="FFFF0000"/><name val="Arial"/></font>', 'color ordered after sz, before name')
		var caller = { color: { rgb: 'FF0000' } }
		createFiles({ styles: { F: { font: caller } }, sheets: [{ data: [[1]] }] })
		assert.equal(caller, { color: { rgb: 'FF0000' } }, 'caller color object left untouched')
		assert.end()
	})
	test("styleSheet carries the named style tables", function(assert) {
		var styles = createFiles({ sheets: [{ data: [[1]] }] })
		.find(function(f) { return f.name === 'xl/styles.xml' }).content
		assert.ok(styles.indexOf('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>') > -1, 'cellStyleXfs present')
		assert.ok(styles.indexOf('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>') > -1, 'cellStyles present')
		assert.ok(styles.indexOf('<cellStyleXfs') < styles.indexOf('<cellXfs'), 'cellStyleXfs sits before cellXfs')
		assert.ok(styles.indexOf('<cellXfs') < styles.indexOf('<cellStyles'), 'cellStyles sits after cellXfs')
		assert.end()
	})
	test("null rows preserve row positions", function(assert) {
		var files = createFiles({
			sheets: [{ data: [['A'], null, ['C']] }]
		})
		var sheet = files.find(function(f) { return f.name === 'xl/worksheets/sheet1.xml' }).content
		assert.ok(sheet.indexOf('r="3"') > -1, 'third row has r=3')
		assert.end()
	})
	test("dimension ref covers the widest row", function(assert) {
		assert.ok(sheet1([[], ['a', 'b']]).indexOf('ref="A1:B2"') > -1, 'empty first row does not shrink the ref')
		assert.ok(sheet1([['a'], ['a', 'b', 'c']]).indexOf('ref="A1:C2"') > -1, 'jagged rows use the widest row')
		assert.ok(sheet1([[], []]).indexOf('<dimension') === -1, 'no cells means no dimension')
		assert.end()
	})
	test("partial freeze defaults the missing axis", function(assert) {
		assert.ok(sheet1([['a']], { freeze: { rows: 1 } }).indexOf('topLeftCell="A2"') > -1, 'missing cols freezes from column A')
		assert.ok(sheet1([['a']], { freeze: { cols: 1 } }).indexOf('topLeftCell="B1"') > -1, 'missing rows freezes from row 1')
		assert.end()
	})
	test("dimension ref correct when first row is object", function(assert) {
		var files = createFiles({
			sheets: [{ data: [{ hidden: true, data: ['a', 'b'] }, ['c', 'd']] }]
		})
		var sheet = files.find(function(f) { return f.name === 'xl/worksheets/sheet1.xml' }).content
		assert.ok(sheet.indexOf('ref="A1:B2"') > -1, 'dimension uses column count from row data')
		assert.end()
	})
})
