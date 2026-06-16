'use strict'
var cp = require('child_process')
var fs = require('fs')
var os = require('os')
var path = require('path')

var root = path.join(__dirname, '..')
var snaps = fs.readdirSync(path.join(__dirname, 'snap'))
	.filter(function (f) { return f.endsWith('.snap.xlsx') })
	.map(function (f) { return path.join(__dirname, 'snap', f) })
var versions = ['Microsoft365', 'Office2007']

function platformId() {
	var plat = { linux: 'linux', darwin: 'macos', win32: 'windows' }[process.platform] || 'linux'
	var arch = { x64: 'x64', arm64: 'arm64', ia32: 'x86' }[process.arch] || process.arch
	return 'ooxml-validator-' + plat + '-' + arch
}

function findBin() {
	var caches = [process.env.npm_config_cache, process.env.NPM_CONFIG_CACHE, path.join(os.homedir(), '.npm')]
	if (process.env.LOCALAPPDATA) caches.push(path.join(process.env.LOCALAPPDATA, 'npm-cache'))
	var id = platformId()
	for (var i = 0; i < caches.length; i++) {
		if (!caches[i]) continue
		var npxDir = path.join(caches[i], '_npx')
		if (!fs.existsSync(npxDir)) continue
		var hit = walk(npxDir, id)
		if (hit) return hit
	}
	return null
}

function walk(dir, id) {
	var entries
	try { entries = fs.readdirSync(dir, { withFileTypes: true }) } catch (e) { return null }
	for (var i = 0; i < entries.length; i++) {
		var ent = entries[i]
		var p = path.join(dir, ent.name)
		if (!ent.isDirectory()) continue
		if (ent.name === id) {
			var bin = path.join(p, 'ooxml-validator' + (process.platform === 'win32' ? '.exe' : ''))
			if (fs.existsSync(bin)) return bin
		}
		var nested = walk(p, id)
		if (nested) return nested
	}
	return null
}

function ensureBin() {
	var bin = findBin()
	if (bin) return bin
	cp.execFileSync('npx', ['--yes', '@xarsh/ooxml-validator', snaps[0]], { stdio: 'pipe', cwd: root })
	bin = findBin()
	if (!bin) throw new Error('ooxml-validator binary not found after npx install')
	return bin
}

var bin = ensureBin()
var fails = 0

snaps.forEach(function (file) {
	versions.forEach(function (ver) {
		var out = cp.execFileSync(bin, [file, ver], { encoding: 'utf8' })
		var result = JSON.parse(out)
		console.log('# validate %s (%s) %s', path.basename(file), ver, result.ok ? 'ok' : 'FAIL')
		if (!result.ok) {
			fails++
			result.errors.forEach(function (e) {
				console.error('  %s: %s', e.path, e.description)
			})
		}
	})
})

process.exit(fails ? 1 : 0)
