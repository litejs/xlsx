
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
						['Datetime', { format: 'datetime', value: new Date(1514900750001) }],
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
				Fill1: {
					fill: 'FFFF00',
				},
				Fill2: {
					fill: { bgColor: 'FF9900', pattern: 'solid' },
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
	test("wrapped Date without format gets datetime style", function(assert) {
		var files = createFiles({
			sheets: [{ data: [[{ value: new Date(1514900750001) }]] }]
		})
		var sheet = files.find(function(f) { return f.name === 'xl/worksheets/sheet1.xml' }).content
		assert.ok(sheet.indexOf(' s="2"') > -1, 'datetime style applied')
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
