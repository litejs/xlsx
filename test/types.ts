import { createFiles, createXlsx, Workbook, XlsxFile } from ".."

var workbook: Workbook = {
	styles: {
		bold: { font: { sz: 11, name: "Calibri", b: true } },
		border: { border: "thin" },
		fill: { fill: "FFFF00" },
	},
	sheets: [
		{
			name: "Products",
			cols: [{ width: 20, bestFit: true, customWidth: true }, 0, "15"],
			freeze: { rows: 1, cols: 0 },
			data: [
				["Apple", 1.99, 10],
				["Banana", 0.99, null],
				{ hidden: true, data: ["hidden row"] },
				{ height: 25, data: [{ style: "bold", value: "Total" }, { format: "date", value: new Date() }] },
				null,
			]
		},
		null,
		[["A", "B", "C"]],
	]
}

var files: XlsxFile[] = createFiles(workbook)
files[0].name satisfies string
files[0].content satisfies string

// Promise form
createXlsx(workbook) satisfies Promise<Uint8Array>
createXlsx(workbook, { comment: "test" }) satisfies Promise<Uint8Array>

// Callback form
createXlsx(workbook, (err, data) => {
	data satisfies Uint8Array
})
createXlsx(workbook, {}, (err, data) => {
	data satisfies Uint8Array
})
