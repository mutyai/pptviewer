#!/usr/bin/env node

const fs = require("fs")
const path = require("path")
const { execFileSync, execSync } = require("child_process")

// Config
const repoRoot = path.resolve(__dirname, "..")
const packageJsonPath = path.join(repoRoot, "package.json")
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, "utf8"))
const version = packageJson.version

const vsixName = `muty-pptviewer-${version}.vsix`
const vsixPath = path.join(repoRoot, vsixName)
const targetVsixPath = `/Users/terence/Project/ainote/extensions/vsix/muty-pptviewer-${version}.vsix`
const builtInExtDir = "/Users/terence/Project/ainote/.build/builtInExtensions/muty-pptviewer"
const productJsonPath = "/Users/terence/Project/ainote/product.json"

function log(msg) {
	console.log(msg)
}

function fail(msg) {
	console.error(`❌ ${msg}`)
	process.exit(1)
}

function ensureFileExists(file) {
	if (!fs.existsSync(file)) fail(`未找到文件: ${file}`)
}

function runVscePackage() {
	log("▶️  运行 npx vsce package ...")
	try {
		// 让 vsce 直接继承 stdio：可以看到 validate-extension.js 输出
		execFileSync("npx", ["vsce", "package"], { stdio: "inherit", cwd: repoRoot })
	} catch (e) {
		fail("vsce package 执行失败（语法检查未通过或构建失败）")
	}
	ensureFileExists(vsixPath)
	log("✅ VSIX 构建完成并通过语法检查")
}

function copyVsix() {
	log("▶️  覆盖拷贝 VSIX 到 extensions/vsix ...")
	const targetDir = path.dirname(targetVsixPath)
	fs.mkdirSync(targetDir, { recursive: true })
	if (fs.existsSync(targetVsixPath)) {
		fs.rmSync(targetVsixPath, { force: true })
	}
	fs.copyFileSync(vsixPath, targetVsixPath)
	log("✅ 拷贝完成")
}

function removeBuiltInExtensionDir() {
	log("▶️  删除内置扩展目录 ...")
	if (fs.existsSync(builtInExtDir)) {
		fs.rmSync(builtInExtDir, { recursive: true, force: true })
		log("✅ 已删除旧的内置扩展目录")
	} else {
		log("ℹ️  目录不存在，跳过删除")
	}
}

function computeSha256() {
	log("▶️  计算 VSIX 的 sha256 ...")
	try {
		const out = execSync(`shasum -a 256 ${JSON.stringify(vsixPath)}`, { encoding: "utf8" })
		const sha = (out.split(/\s+/)[0] || "").trim()
		if (!sha || sha.length !== 64) fail("计算 sha256 失败")
		log(`✅ sha256: ${sha}`)
		return sha
	} catch (e) {
		fail(`shasum 命令执行失败: ${e.message}`)
	}
}

function updateProductJson(newSha) {
	log("▶️  更新 product.json 中的 sha256 ...")
	ensureFileExists(productJsonPath)
	let json
	try {
		json = JSON.parse(fs.readFileSync(productJsonPath, "utf8"))
	} catch (e) {
		fail(`读取/解析 product.json 失败: ${e.message}`)
	}

	// product.json 里可能有数组或对象，按用户提供的结构查找并更新
	let updated = false
	const updateEntry = (entry) => {
		if (entry && entry.name === "muty-pptviewer") {
			entry.version = version
			entry.vsix = `extensions/vsix/muty-pptviewer-${version}.vsix`
			entry.sha256 = newSha
			updated = true
		}
	}

	// 兼容 builtInExtensions（你的 product.json 使用该键）
	if (Array.isArray(json.builtInExtensions)) {
		json.builtInExtensions.forEach(updateEntry)
	}
	// 兼容 extensions（其他结构）
	if (Array.isArray(json.extensions)) {
		json.extensions.forEach(updateEntry)
	}
	// 兼容顶层就是对象的情况
	updateEntry(json)

	if (!updated) fail("未在 product.json 中找到需要更新的条目")

	try {
		fs.writeFileSync(productJsonPath, JSON.stringify(json, null, 2) + "\n", "utf8")
		log("✅ product.json 已更新")
	} catch (e) {
		fail(`写入 product.json 失败: ${e.message}`)
	}
}

async function main() {
	runVscePackage()
	copyVsix()
	removeBuiltInExtensionDir()
	const sha = computeSha256()
	updateProductJson(sha)
	log("🎉 全部完成")
}

main().catch((e) => fail(e.message))
