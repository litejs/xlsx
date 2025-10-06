
describe("xlsx", function() {
	require("@litejs/cli/snapshot.js")
	var createXlsx = require("..").createXlsx

	test("Readme", typeof CompressionStream !== "undefined" && typeof Response !== "undefined" && function(assert, mock) {
		mock.swap(Date, "now", mock.fn(1514900750001))
		createXlsx({
			sheets: [
				{
					name: 'Products',
					widths: '20,10,10',
					data: [
						['Apple', 1.99, 10],
						['Banana', 0.99, null],
						['Orange', 2.49, 8],
						null,
						['Sum', '=SUM(B1:B3)', 8]
					]
				},
			]
		})
		.then(uint8 => {
			assert.matchSnapshot("test/snap/readme.xlsx", uint8)
			assert.end()
		})
	})
})

