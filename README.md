
[1]: https://badgen.net/coveralls/c/github/litejs/xlsx
[2]: https://coveralls.io/r/litejs/xlsx
[3]: https://badgen.net/packagephobia/install/@litejs/xlsx
[4]: https://packagephobia.now.sh/result?p=@litejs/xlsx
[5]: https://badgen.net/badge/icon/Buy%20Me%20A%20Tea/orange?icon=kofi&label
[6]: https://www.buymeacoffee.com/lauriro


LiteJS Xlsx &ndash; [![Coverage][1]][2] [![Size][3]][4] [![Buy Me A Tea][5]][6]
===========

JavaScript library for creating XLSX files in Browser or Server.


Examples
--------

```javascript
const { createXlsx } = require("@litejs/xlsx");
const fileAsUint8Array = await createXlsx({
    sheets: [
		{
			name: 'Products',
			data: [
				['Apple', 1.99, 10, new Date()],
				['Banana', 0.99, 15],
				['Orange', 2.49, 8]
			]
		},
    ]
})
```

## Contributing

Follow [Coding Style Guide](https://github.com/litejs/litejs/wiki/Style-Guide),
run tests `npm install; npm test`.


> Copyright (c) 2025 Lauri Rooden &lt;lauri@rooden.ee&gt;  
[MIT License](https://litejs.com/MIT-LICENSE.txt) |
[GitHub repo](https://github.com/litejs/xlsx) |
[npm package](https://npmjs.org/package/@litejs/xlsx) |
[Buy Me A Tea][6]


