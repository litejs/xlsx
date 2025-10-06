
describe("xlsx", function() {
	require("@litejs/cli/snapshot.js")
	var createXlsx = require("..").createXlsx

	test("Readme", typeof CompressionStream !== "undefined" && typeof Response !== "undefined" && function(assert, mock) {
		mock.swap(Date, "now", mock.fn(1514900750001))
		createXlsx({
			sheets: [
				{
					name: 'Products',
					cols: '20,,15',
					data: [
						['Apple', 1.99, false],
						['Banana', 0.99, null],
						['Orange', { value: 2.49 }, true],
						null,
						['Sum', {style: 'bold', value: '=SUM(B1:B3)'}, new Date(1514900750001)]
					]
				},
				[
					['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA']
				],
				{
					name: 'Empty Sheet',
					data: []
				}
			]
		})
		.then(uint8 => {
			assert.matchSnapshot("test/snap/readme.xlsx", uint8)
			assert.end()
		})
	})
})

