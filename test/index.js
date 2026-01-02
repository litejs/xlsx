
describe("xlsx", function() {
	require("@litejs/cli/snapshot.js")
	var { createFiles, createXlsx } = require("..")

	test("Readme", typeof CompressionStream !== "undefined" && typeof Response !== "undefined" && function(assert, mock) {
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
		createXlsx(workbook)
		.then(uint8 => {
			assert.matchSnapshot("test/snap/readme.xlsx", uint8)
			assert.end()
		})
	})
	test("styles", typeof CompressionStream !== "undefined" && typeof Response !== "undefined" && function(assert, mock) {
		mock.swap(Date, "now", mock.fn(1514900750001))
		var workbook = {
			styles: {
				My1: {
					font: { sz: 15, name: "Calibri" },
				},
				Plain: {},
			},
			sheets: [
				{
					cols: [{ width: null }],
					name: 'Styles',
					data: [
						[{style: 'My1', value: 'Apple My1'}, { style: "Plain", value: "Banana Plain" }],
						{ hidden: true, data: ['Hidden Row', 1] },
						{ height: 25, data: ['Sized Row', 1] },
					]
				},
			]
		}
		assert.matchSnapshot("test/snap/styles.json", JSON.stringify(createFiles(workbook), null, 2))
		createXlsx(workbook)
		.then(uint8 => {
			assert.matchSnapshot("test/snap/styles.xlsx", uint8)
			assert.end()
		})
	})
})

