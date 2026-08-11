'use strict'
var cp = require('child_process')
, fs = require('fs')
, path = require('path')
, validator = '@xarsh/ooxml-validator@0.3.0'
, npx = process.platform === 'win32' ? 'npx.cmd' : 'npx'
, snapDir = path.join(__dirname, 'snap')
, versions = ['Microsoft365', 'Office2007']
, fails = 0

fs.readdirSync(snapDir).forEach(function (file) {
	if (file.endsWith('.snap.xlsx')) versions.forEach(function (ver) {
		var out
		try {
			out = cp.execFileSync(npx, ['--yes', validator, path.join(snapDir, file), '--officeVersion', ver], { encoding: 'utf8' })
		} catch (e) {
			// an invalid file exits 1, the report is still on stdout, keep going
			if (!(out = e.stdout)) throw e
		}
		var result = JSON.parse(out)
		console.log('# validate %s (%s) %s', file, ver, result.ok ? 'ok' : 'FAIL')
		if (!result.ok) {
			fails++
			result.errors.forEach(function (e) {
				console.error('  %s %s: %s', e.path, e.xPath, e.description)
			})
		}
	})
})

process.exit(fails ? 1 : 0)

