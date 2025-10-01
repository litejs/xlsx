
import { createXlsx } from "../xlsx.js"

describe("Run as ESM module", () => {
	it("should have function createXlsx", assert => {
		assert.type(createXlsx, "function")
		assert.end()
	})
})

