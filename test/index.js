
describe("xlsx", function() {
	require("@litejs/cli/snapshot.js")
	var { createFiles, createXlsx } = require("..")
	, compressionSuported = typeof CompressionStream !== "undefined" && typeof Response !== "undefined"

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
})
