
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
						['Apple', 1.99, false],
						['Banana', 0.99, null],
						['Orange', { value: 2.49 }, true],
						null,
						['Sum of all fruits', {style: 'bold', value: '=SUM(B1:B3)'}, new Date(1514900750001)]
					]
				},
				null,
				[
					['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA']
				],
				{
					name: 'Empty Sheet',
					data: []
				}
			]
		}
		assert.matchSnapshot("test/snap/readme.json", JSON.stringify(createFiles(workbook), null, 2))
		createXlsx(workbook)
		.then(uint8 => {
			assert.matchSnapshot("test/snap/readme.xlsx", uint8)
			assert.end()
		})
	})
})

