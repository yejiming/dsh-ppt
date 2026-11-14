import { i as throwIfAborted, n as PptError, r as asPptError } from "./errors-BdtilPdq.mjs";
import { homedir } from "node:os";
import { basename, delimiter, dirname, extname, isAbsolute, join, posix, relative, resolve, sep, win32 } from "node:path";
import { access, copyFile, link, lstat, mkdir, open, opendir, readFile, readdir, realpath, rename, rm, stat, unlink } from "node:fs/promises";
import { constants } from "node:fs";
import { createHash, randomUUID } from "node:crypto";
import * as fontkit from "fontkit";
import { fileURLToPath, pathToFileURL } from "node:url";
import { lookup } from "node:dns/promises";
import { isIP } from "node:net";
import sharp from "sharp";
import { decompressSync, deflateSync, strFromU8, strToU8, unzipSync } from "fflate";
import { z } from "zod";
import PptxGenJS from "pptxgenjs";
//#region src/platform.ts
function isSupportedPlatform(platform = process.platform) {
	return platform === "darwin" || platform === "linux" || platform === "win32";
}
function systemFontDirectories(platform = process.platform, home = homedir(), env = process.env) {
	if (platform === "darwin") return [
		"/System/Library/Fonts",
		"/System/Library/Fonts/Supplemental",
		"/System/Library/AssetsV2/com_apple_MobileAsset_Font7",
		"/System/Library/AssetsV2/com_apple_MobileAsset_Font8",
		"/Library/Fonts",
		posix.join(home, "Library/Fonts")
	];
	if (platform === "win32") {
		const windows = env.SystemRoot ?? env.WINDIR ?? "C:\\Windows";
		const local = env.LOCALAPPDATA;
		return [win32.join(windows, "Fonts"), ...local === void 0 ? [] : [win32.join(local, "Microsoft", "Windows", "Fonts")]];
	}
	return [
		"/usr/share/fonts",
		"/usr/local/share/fonts",
		posix.join(home, ".local/share/fonts"),
		posix.join(home, ".fonts")
	];
}
function libreOfficeCandidates(platform = process.platform, env = process.env) {
	if (platform === "darwin") return [
		"soffice",
		"/Applications/LibreOffice.app/Contents/MacOS/soffice",
		"/Applications/LibreOfficeDev.app/Contents/MacOS/soffice"
	];
	if (platform === "win32") return [
		"soffice.exe",
		"soffice",
		...[
			env.ProgramFiles,
			env["ProgramFiles(x86)"],
			env.LOCALAPPDATA
		].filter((value) => typeof value === "string" && value.length > 0).map((root) => win32.join(root, "LibreOffice", "program", "soffice.exe"))
	];
	return [
		"soffice",
		"libreoffice",
		"/usr/bin/soffice",
		"/usr/bin/libreoffice"
	];
}
function pdfToPpmCandidates(platform = process.platform) {
	return platform === "win32" ? ["pdftoppm.exe", "pdftoppm"] : ["pdftoppm"];
}
function appleScriptCandidates(platform = process.platform) {
	return platform === "darwin" ? ["/usr/bin/osascript", "osascript"] : [];
}
function screenCaptureCandidates(platform = process.platform) {
	return platform === "darwin" ? ["/usr/sbin/screencapture", "screencapture"] : [];
}
function keynoteCandidates(platform = process.platform, home = homedir()) {
	return platform === "darwin" ? ["/Applications/Keynote.app", posix.join(home, "Applications/Keynote.app")] : [];
}
function powerShellCandidates(platform = process.platform) {
	return platform === "win32" ? [
		"powershell.exe",
		"powershell",
		"pwsh.exe",
		"pwsh"
	] : [];
}
function powerPointCandidates(platform = process.platform, env = process.env, home = homedir()) {
	if (platform === "darwin") return ["/Applications/Microsoft PowerPoint.app", posix.join(home, "Applications/Microsoft PowerPoint.app")];
	if (platform !== "win32") return [];
	return [
		env.ProgramFiles,
		env["ProgramFiles(x86)"],
		env.LOCALAPPDATA
	].filter((value) => typeof value === "string" && value.length > 0).flatMap((root) => [win32.join(root, "Microsoft Office", "root", "Office16", "POWERPNT.EXE"), win32.join(root, "Microsoft Office", "Office16", "POWERPNT.EXE")]);
}
function pptImageBackendOrder(platform = process.platform, requested = "auto") {
	if (requested !== "auto") return [requested];
	if (platform === "darwin") return [
		"keynote",
		"libreoffice",
		"powerpoint"
	];
	if (platform === "win32") return ["powerpoint", "libreoffice"];
	return platform === "linux" ? ["libreoffice"] : [];
}
function browserSystemCandidates(platform = process.platform, home = homedir(), env = process.env) {
	if (platform === "darwin") return [
		"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
		"/Applications/Chromium.app/Contents/MacOS/Chromium",
		posix.join(home, "Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
	];
	if (platform === "win32") return [
		env.ProgramFiles,
		env["ProgramFiles(x86)"],
		env.LOCALAPPDATA
	].filter((value) => typeof value === "string" && value.length > 0).flatMap((root) => [
		win32.join(root, "Google", "Chrome", "Application", "chrome.exe"),
		win32.join(root, "Chromium", "Application", "chrome.exe"),
		win32.join(root, "Microsoft", "Edge", "Application", "msedge.exe")
	]);
	return [
		"/usr/bin/google-chrome",
		"/usr/bin/google-chrome-stable",
		"/usr/bin/chromium",
		"/usr/bin/chromium-browser",
		"/snap/bin/chromium"
	];
}
//#endregion
//#region src/fonts.ts
const FONT_LAYERS = [
	"portable",
	"system",
	"custom"
];
const FONT_ROLES = [
	"latin-sans",
	"latin-serif",
	"cjk-sans",
	"cjk-serif",
	"display",
	"code"
];
const ALL_PLATFORMS = [
	"darwin",
	"win32",
	"linux"
];
const FONT_REGISTRY = Object.freeze([
	{
		name: "Arial",
		aliases: ["ArialMT"],
		language: "Multi-language",
		style: "Sans-serif",
		characteristics: "Portable Office-safe sans-serif",
		layer: "portable",
		platforms: ALL_PLATFORMS,
		roles: ["latin-sans"]
	},
	{
		name: "Times New Roman",
		aliases: ["TimesNewRomanPSMT"],
		language: "Multi-language",
		style: "Serif",
		characteristics: "Portable Office-safe serif",
		layer: "portable",
		platforms: ALL_PLATFORMS,
		roles: ["latin-serif"]
	},
	{
		name: "Segoe UI",
		aliases: ["SegoeUI"],
		language: "Western",
		style: "Sans-serif",
		characteristics: "Windows interface sans-serif",
		layer: "system",
		platforms: ["win32"],
		roles: ["latin-sans"]
	},
	{
		name: "Microsoft YaHei",
		aliases: [
			"Microsoft YaHei UI",
			"MicrosoftYaHei",
			"MicrosoftYaHeiUI"
		],
		language: "Chinese + Western",
		style: "Sans-serif",
		characteristics: "Windows ClearType Chinese sans-serif",
		layer: "system",
		platforms: ["win32"],
		roles: ["cjk-sans", "latin-sans"]
	},
	{
		name: "DengXian",
		aliases: ["Deng"],
		language: "Chinese + Western",
		style: "Sans-serif",
		characteristics: "Windows modern Chinese sans-serif",
		layer: "system",
		platforms: ["win32"],
		roles: ["cjk-sans", "latin-sans"]
	},
	{
		name: "SimSun",
		aliases: ["NSimSun"],
		language: "Chinese + Western",
		style: "Serif",
		characteristics: "Windows Song-style Chinese serif",
		layer: "system",
		platforms: ["win32"],
		roles: ["cjk-serif", "latin-serif"]
	},
	{
		name: "Helvetica Neue",
		aliases: ["HelveticaNeue"],
		language: "Western",
		style: "Sans-serif",
		characteristics: "macOS system sans-serif",
		layer: "system",
		platforms: ["darwin"],
		roles: ["latin-sans"]
	},
	{
		name: "PingFang SC",
		aliases: ["PingFangSC"],
		language: "Chinese + Western",
		style: "Sans-serif",
		characteristics: "macOS Simplified Chinese system sans-serif",
		layer: "system",
		platforms: ["darwin"],
		roles: ["cjk-sans", "latin-sans"]
	},
	{
		name: "Hiragino Sans GB",
		aliases: ["HiraginoSansGB"],
		language: "Chinese + Western",
		style: "Sans-serif",
		characteristics: "macOS Simplified Chinese humanist sans-serif",
		layer: "system",
		platforms: ["darwin"],
		roles: ["cjk-sans", "latin-sans"]
	},
	{
		name: "Songti SC",
		aliases: ["SongtiSC"],
		language: "Chinese + Western",
		style: "Serif",
		characteristics: "macOS Simplified Chinese Song serif",
		layer: "system",
		platforms: ["darwin"],
		roles: ["cjk-serif", "latin-serif"]
	},
	{
		name: "Liberation Sans",
		aliases: ["LiberationSans"],
		language: "Western",
		style: "Sans-serif",
		characteristics: "Linux metric-compatible Arial alternative",
		layer: "system",
		platforms: ["linux"],
		roles: ["latin-sans"]
	},
	{
		name: "Liberation Serif",
		aliases: ["LiberationSerif"],
		language: "Western",
		style: "Serif",
		characteristics: "Linux metric-compatible Times alternative",
		layer: "system",
		platforms: ["linux"],
		roles: ["latin-serif"]
	},
	{
		name: "DejaVu Sans",
		aliases: ["DejaVuSans"],
		language: "Multi-language",
		style: "Sans-serif",
		characteristics: "Widely available Linux sans-serif with broad glyph coverage",
		layer: "system",
		platforms: ["linux"],
		roles: ["latin-sans", "code"]
	},
	{
		name: "DejaVu Serif",
		aliases: ["DejaVuSerif"],
		language: "Multi-language",
		style: "Serif",
		characteristics: "Widely available Linux serif",
		layer: "system",
		platforms: ["linux"],
		roles: ["latin-serif"]
	},
	{
		name: "Noto Sans CJK SC",
		aliases: ["NotoSansCJKsc"],
		language: "Chinese + Multi-language",
		style: "Sans-serif",
		characteristics: "Linux-oriented open CJK sans-serif",
		layer: "system",
		platforms: ["linux"],
		roles: ["cjk-sans", "latin-sans"]
	},
	{
		name: "Noto Serif CJK SC",
		aliases: ["NotoSerifCJKsc"],
		language: "Chinese + Multi-language",
		style: "Serif",
		characteristics: "Linux-oriented open CJK serif",
		layer: "system",
		platforms: ["linux"],
		roles: ["cjk-serif", "latin-serif"]
	},
	{
		name: "Liter",
		language: "English",
		style: "Sans-serif",
		characteristics: "Modern geometric, low contrast, balanced and rational",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["latin-sans", "display"]
	},
	{
		name: "HedvigLettersSans",
		aliases: ["Hedvig Letters Sans"],
		language: "English",
		style: "Sans-serif",
		characteristics: "Slightly irregular with a distinctive brand character",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["latin-sans", "display"]
	},
	{
		name: "Oranienbaum",
		language: "English",
		style: "High-contrast serif",
		characteristics: "Geometric, elegant and classical",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["latin-serif", "display"]
	},
	{
		name: "QuattrocentoSans",
		aliases: ["Quattrocento Sans"],
		language: "English",
		style: "Classical sans-serif",
		characteristics: "Gentle, readable and sharp at small sizes",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["latin-sans"]
	},
	{
		name: "SortsMillGoudy",
		aliases: ["Sorts Mill Goudy"],
		language: "English",
		style: "Serif",
		characteristics: "Goudy Old Style revival with soft, legible serifs",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["latin-serif"]
	},
	{
		name: "Unna",
		language: "English",
		style: "Neoclassical serif",
		characteristics: "Pronounced vertical rhythm and elegant power",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["latin-serif", "display"]
	},
	{
		name: "Coda",
		language: "English",
		style: "Sans-serif",
		characteristics: "Round, friendly and open",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["latin-sans", "display"]
	},
	{
		name: "Jersey15",
		aliases: ["Jersey 15"],
		language: "English + Numbers",
		style: "Pixel",
		characteristics: "Sports jersey geometry with a strong grid",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["display"]
	},
	{
		name: "Jersey20Charted",
		aliases: ["Jersey 20 Charted"],
		language: "English + Numbers",
		style: "Pixel",
		characteristics: "Grid-textured sports number style",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["display"]
	},
	{
		name: "MiSans",
		aliases: ["Mi Sans"],
		language: "Chinese + Multi-language",
		style: "Sans-serif",
		characteristics: "Clean modern variable system font",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: [
			"cjk-sans",
			"latin-sans",
			"display"
		]
	},
	{
		name: "Noto Sans SC",
		aliases: ["NotoSansSC"],
		language: "Chinese + Multi-language",
		style: "Sans-serif",
		characteristics: "Neutral standardized Source Han Sans structure",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-sans", "latin-sans"]
	},
	{
		name: "siyuanSongti",
		aliases: ["Source Han Serif SC", "Source Han Serif CN"],
		language: "Chinese + Multi-language",
		style: "Serif",
		characteristics: "Refined Song structure with contrasting strokes",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-serif", "latin-serif"]
	},
	{
		name: "alimamadaoliti",
		language: "Chinese",
		style: "Clerical",
		characteristics: "Knife-edge strokes with power and antiquity",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-serif", "display"]
	},
	{
		name: "alimamashuheiti",
		language: "Chinese",
		style: "Geometric sans-serif",
		characteristics: "Orderly commercial geometry",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-sans", "display"]
	},
	{
		name: "zhankuwenyiti",
		language: "Chinese",
		style: "Handwritten",
		characteristics: "Simple, fresh and lightly artistic",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-sans", "display"]
	},
	{
		name: "feibozhengdianti",
		language: "Chinese",
		style: "Brush",
		characteristics: "Thick and powerful brush strokes",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-serif", "display"]
	},
	{
		name: "deyihei",
		language: "Chinese",
		style: "Sans-serif",
		characteristics: "Thin slanted humanist geometry",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-sans", "display"]
	},
	{
		name: "jingpindianzhenTi",
		language: "Chinese + Western",
		style: "Pixel",
		characteristics: "9x9 retro-electronic bitmap style",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["display", "code"]
	},
	{
		name: "LXGW Bright",
		language: "Chinese + Western",
		style: "Song/Kai",
		characteristics: "Gentle, clear and legible",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: [
			"cjk-serif",
			"latin-serif",
			"display"
		]
	},
	{
		name: "ZCOOL KuaiLe",
		language: "Chinese + Western",
		style: "Display",
		characteristics: "Lively, playful and youthful",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["display"]
	},
	{
		name: "xiawuxinzhisong",
		language: "Chinese",
		style: "Serif",
		characteristics: "Bright and elegant Mincho-derived structure",
		layer: "custom",
		platforms: ALL_PLATFORMS,
		roles: ["cjk-serif", "display"]
	}
]);
const REGISTERED_BY_KEY = /* @__PURE__ */ new Map();
for (const font of FONT_REGISTRY) for (const name of [font.name, ...font.aliases ?? []]) REGISTERED_BY_KEY.set(fontKey(name), font);
const REGISTERED_ALIASES = [...REGISTERED_BY_KEY.entries()].sort((a, b) => b[0].length - a[0].length);
function fontKey(value) {
	return value.toLocaleLowerCase().replace(/[\s_-]+/g, "");
}
function registeredFace(value) {
	const key = fontKey(value);
	const exact = REGISTERED_BY_KEY.get(key);
	if (exact !== void 0) return exact;
	for (const [alias, font] of REGISTERED_ALIASES) {
		if (!key.startsWith(alias)) continue;
		const suffix = key.slice(alias.length);
		if (/^(?:(?:extra|ultra|semi|demi)?(?:light|bold)|thin|regular|book|medium|heavy|black|italic|oblique|w\d+)+$/u.test(suffix)) return font;
	}
}
async function fontFiles(roots, maxFiles = 1e4) {
	const files = [];
	const queue = [...new Set(roots)];
	while (queue.length > 0 && files.length < maxFiles) {
		const directory = queue.shift();
		let handle;
		try {
			handle = await opendir(directory);
		} catch {
			continue;
		}
		for await (const entry of handle) {
			const path = join(directory, entry.name);
			if (entry.isDirectory()) queue.push(path);
			else if (entry.isFile() && /^\.(?:ttf|otf|ttc)$/i.test(extname(entry.name))) files.push(path);
			if (files.length >= maxFiles) break;
		}
	}
	return files.sort();
}
function faces(value) {
	return Array.isArray(value.fonts) ? value.fonts : [value];
}
async function discoverRegisteredFonts(extraDirs = [], platform = process.platform) {
	const discovered = [];
	for (const file of await fontFiles([...systemFontDirectories(platform), ...extraDirs])) {
		let opened;
		try {
			opened = await fontkit.open(file);
		} catch {
			continue;
		}
		let hash;
		for (const face of faces(opened)) {
			const familyName = typeof face.familyName === "string" ? face.familyName : "";
			const postscriptName = typeof face.postscriptName === "string" ? face.postscriptName : null;
			const registered = registeredFace(familyName) ?? (postscriptName === null ? void 0 : registeredFace(postscriptName));
			if (registered === void 0) continue;
			hash ??= createHash("sha256").update(await readFile(file)).digest("hex");
			const points = new Set(Array.isArray(face.characterSet) ? face.characterSet.filter((value) => Number.isInteger(value)) : []);
			discovered.push({
				name: registered.name,
				file,
				sha256: hash,
				familyName,
				postscriptName,
				weight: typeof face.subfamilyName === "string" ? face.subfamilyName : "Regular",
				glyphCount: points.size,
				supportsLatin: [..."AaZz09"].every((character) => points.has(character.codePointAt(0))),
				supportsCjk: points.has("中".codePointAt(0)) && points.has("文".codePointAt(0)),
				codePoints: points
			});
		}
	}
	return discovered.sort((a, b) => a.name.localeCompare(b.name) || a.weight.localeCompare(b.weight) || a.file.localeCompare(b.file));
}
const FONT_FALLBACKS = Object.freeze({
	Liter: [
		"HedvigLettersSans",
		"QuattrocentoSans",
		"Arial"
	],
	HedvigLettersSans: [
		"Liter",
		"QuattrocentoSans",
		"Arial"
	],
	MiSans: ["alimamashuheiti", "Noto Sans SC"],
	"Noto Sans SC": ["MiSans", "alimamashuheiti"],
	siyuanSongti: ["xiawuxinzhisong", "LXGW Bright"],
	xiawuxinzhisong: ["siyuanSongti", "LXGW Bright"]
});
const PLATFORM_ROLE_FALLBACKS = Object.freeze({
	darwin: Object.freeze({
		"latin-sans": [
			"Arial",
			"Helvetica Neue",
			"PingFang SC",
			"Liter",
			"Noto Sans SC"
		],
		"latin-serif": [
			"Times New Roman",
			"Songti SC",
			"siyuanSongti"
		],
		"cjk-sans": [
			"PingFang SC",
			"Hiragino Sans GB",
			"Noto Sans SC",
			"MiSans"
		],
		"cjk-serif": [
			"Songti SC",
			"siyuanSongti",
			"LXGW Bright"
		],
		display: [
			"Liter",
			"PingFang SC",
			"Arial",
			"Noto Sans SC"
		],
		code: [
			"Arial",
			"Helvetica Neue",
			"PingFang SC",
			"Noto Sans SC"
		]
	}),
	win32: Object.freeze({
		"latin-sans": [
			"Arial",
			"Segoe UI",
			"Microsoft YaHei",
			"Liter",
			"Noto Sans SC"
		],
		"latin-serif": [
			"Times New Roman",
			"SimSun",
			"siyuanSongti"
		],
		"cjk-sans": [
			"Microsoft YaHei",
			"DengXian",
			"Noto Sans SC",
			"MiSans"
		],
		"cjk-serif": [
			"SimSun",
			"siyuanSongti",
			"LXGW Bright"
		],
		display: [
			"Liter",
			"Microsoft YaHei",
			"Arial",
			"Noto Sans SC"
		],
		code: [
			"Arial",
			"Segoe UI",
			"Microsoft YaHei",
			"Noto Sans SC"
		]
	}),
	linux: Object.freeze({
		"latin-sans": [
			"Liberation Sans",
			"DejaVu Sans",
			"Arial",
			"Liter",
			"Noto Sans SC"
		],
		"latin-serif": [
			"Liberation Serif",
			"DejaVu Serif",
			"Times New Roman",
			"Noto Serif CJK SC"
		],
		"cjk-sans": [
			"Noto Sans CJK SC",
			"Noto Sans SC",
			"MiSans"
		],
		"cjk-serif": [
			"Noto Serif CJK SC",
			"siyuanSongti",
			"LXGW Bright"
		],
		display: [
			"Liberation Sans",
			"DejaVu Sans",
			"Noto Sans CJK SC",
			"Noto Sans SC"
		],
		code: [
			"DejaVu Sans",
			"Liberation Sans",
			"Noto Sans CJK SC"
		]
	})
});
function supportedPlatform(platform) {
	return platform === "win32" || platform === "linux" ? platform : "darwin";
}
function containsCjk(text) {
	return /[\u3400-\u9FFF\uF900-\uFAFF]/u.test(text);
}
function semanticRole(font, text) {
	const serif = font.roles.includes("latin-serif") || font.roles.includes("cjk-serif");
	if (containsCjk(text)) return serif ? "cjk-serif" : "cjk-sans";
	if (font.roles.includes("code")) return "code";
	if (font.roles.includes("display") && !font.roles.includes("latin-sans") && !font.roles.includes("latin-serif")) return "display";
	return serif ? "latin-serif" : "latin-sans";
}
/**
* Adapts the machine-wide font inventory onto the shape the outline resolver
* already understands, so a deck may name any font that is actually installed
* instead of only the families in the built-in registry. Coverage is rebuilt
* from the face flags: a code point enters the set only when the face claims the
* script covering it, which keeps `supportsText` meaningful without re-reading
* every glyph table of every installed font. Pass `wanted` to restrict the
* conversion to the families a deck actually names.
*/
function installedFontsAsDiscovered(faces, wanted) {
	const discovered = [];
	for (const face of faces) {
		if (wanted !== void 0 && !wanted.has(fontKey(face.family))) continue;
		const codePoints = /* @__PURE__ */ new Set();
		if (face.supportsLatin) for (let point = 32; point <= 126; point += 1) codePoints.add(point);
		if (face.supportsCjk) for (let point = 19968; point <= 40869; point += 1) codePoints.add(point);
		discovered.push({
			name: face.family,
			file: face.file,
			sha256: face.sha256,
			familyName: face.family,
			postscriptName: face.postscriptName,
			weight: face.subfamily,
			glyphCount: face.glyphCount,
			supportsLatin: face.supportsLatin,
			supportsCjk: face.supportsCjk,
			codePoints
		});
	}
	return discovered;
}
function registeredFont(name) {
	return REGISTERED_BY_KEY.get(fontKey(name));
}
function fontFallbackCandidates(name, text, platform = process.platform) {
	const requested = registeredFont(name);
	if (requested === void 0) throw new PptError("PPT_OUTLINE_INVALID", `font is not registered: ${name}`);
	const role = semanticRole(requested, text);
	return [.../* @__PURE__ */ new Set([
		requested.name,
		...FONT_FALLBACKS[requested.name] ?? [],
		...PLATFORM_ROLE_FALLBACKS[supportedPlatform(platform)][role]
	])];
}
function supportsText(font, text) {
	for (const character of text) {
		const point = character.codePointAt(0);
		if (!/\s/u.test(character) && !font.codePoints.has(point)) return false;
	}
	return true;
}
/**
* A font name this build does not know about — for example a deck authored
* against a newer registry than the one the running host has loaded — must not
* abort the whole pipeline. Fall back in registry order so the outcome stays
* deterministic, and leave a warning the caller can surface.
*/
function resolveUnregisteredFont(name, text, discovered, platform) {
	const supported = supportedPlatform(platform);
	for (const candidate of FONT_REGISTRY) {
		if (!candidate.platforms.includes(supported)) continue;
		const font = discovered.find((item) => item.name === candidate.name && supportsText(item, text));
		if (font !== void 0) return {
			requested: name,
			resolved: font,
			fallback: true,
			warning: `font ${name} is not registered in this build; replaced with ${supported} fallback ${candidate.name}`
		};
	}
	throw new PptError("PPT_DEPENDENCY_MISSING", `no installed approved font covers the requested text for ${name}`, { details: {
		requested: name,
		platform: supported
	} });
}
function resolveRegisteredFont(name, text, discovered, platform = process.platform) {
	const requested = registeredFont(name);
	if (requested === void 0) return resolveUnregisteredFont(name, text, discovered, platform);
	const candidates = fontFallbackCandidates(requested.name, text, platform);
	for (const candidate of candidates) {
		const font = discovered.find((item) => item.name === candidate && supportsText(item, text));
		if (font !== void 0) return {
			requested: name,
			resolved: font,
			fallback: candidate !== requested.name,
			...candidate === requested.name ? {} : { warning: `font ${requested.name} was replaced with installed ${supportedPlatform(platform)} fallback ${candidate}` }
		};
	}
	throw new PptError("PPT_DEPENDENCY_MISSING", `no installed approved font covers the requested text for ${requested.name}`, { details: {
		requested: requested.name,
		platform: supportedPlatform(platform),
		candidates
	} });
}
function summarizeFontAvailability(discovered, platform = process.platform) {
	const currentPlatform = supportedPlatform(platform);
	const availableNames = new Set(discovered.map((font) => font.name));
	const layers = Object.fromEntries(FONT_LAYERS.map((layer) => {
		const registered = FONT_REGISTRY.filter((font) => font.layer === layer);
		const families = registered.filter((font) => availableNames.has(font.name)).map((font) => font.name).sort();
		return [layer, {
			registered: registered.length,
			available: families.length,
			families
		}];
	}));
	const roles = Object.fromEntries(FONT_ROLES.map((role) => {
		const families = [.../* @__PURE__ */ new Set([...PLATFORM_ROLE_FALLBACKS[currentPlatform][role], ...FONT_REGISTRY.filter((font) => font.roles.includes(role)).map((font) => font.name)])].filter((name) => availableNames.has(name)).sort();
		return [role, {
			available: families.length > 0,
			families
		}];
	}));
	return {
		scope: "approved_registry",
		platform: currentPlatform,
		registryFamilies: FONT_REGISTRY.length,
		availableFamilies: availableNames.size,
		availableFaces: discovered.length,
		layers,
		roles
	};
}
function fontRecommendations(discovered, platform = process.platform, text) {
	const currentPlatform = supportedPlatform(platform);
	const available = new Set(discovered.map((font) => font.name));
	return Object.fromEntries(FONT_ROLES.map((role) => {
		return [role, [.../* @__PURE__ */ new Set([...PLATFORM_ROLE_FALLBACKS[currentPlatform][role], ...FONT_REGISTRY.filter((font) => font.roles.includes(role)).map((font) => font.name)])].filter((name) => available.has(name) && (text === void 0 || discovered.some((font) => font.name === name && supportsText(font, text))))];
	}));
}
function buildFontCatalog(discovered, options = {}) {
	const platform = supportedPlatform(options.platform ?? process.platform);
	const role = options.role ?? "all";
	const layer = options.layer ?? "all";
	const includeUnavailable = options.includeUnavailable ?? false;
	const text = options.text;
	const availability = summarizeFontAvailability(discovered, platform);
	const recommendations = fontRecommendations(discovered, platform, text);
	const entries = FONT_REGISTRY.map((descriptor) => {
		const faces = discovered.filter((font) => font.name === descriptor.name);
		const installed = faces.length > 0;
		const coversText = text === void 0 ? void 0 : faces.some((font) => supportsText(font, text));
		const recommendedFor = FONT_ROLES.filter((candidate) => recommendations[candidate].includes(descriptor.name));
		return {
			name: descriptor.name,
			layer: descriptor.layer,
			platforms: [...descriptor.platforms],
			roles: [...descriptor.roles],
			recommended_for: recommendedFor,
			language: descriptor.language,
			style: descriptor.style,
			characteristics: descriptor.characteristics,
			installed,
			weights: [...new Set(faces.map((font) => font.weight))].sort(),
			supports_latin: faces.some((font) => font.supportsLatin),
			supports_cjk: faces.some((font) => font.supportsCjk),
			...coversText === void 0 ? {} : { covers_text: coversText }
		};
	}).filter((font) => {
		if (!includeUnavailable && !font.installed) return false;
		if (!includeUnavailable && text !== void 0 && font.covers_text !== true) return false;
		if (layer !== "all" && font.layer !== layer) return false;
		if (role !== "all" && !font.roles.includes(role) && !font.recommended_for.includes(role)) return false;
		return true;
	});
	const warnings = [];
	if (entries.length === 0) warnings.push("No approved font matches the requested filters on this host.");
	if (role !== "all" && recommendations[role].length === 0) warnings.push(`No installed approved ${role} font covers the requested text.`);
	return {
		scope: "approved_registry",
		scope_note: "This is the installed subset of the plugin approved registry, not the host-wide font inventory.",
		platform,
		registry_families: availability.registryFamilies,
		available_families: availability.availableFamilies,
		available_faces: availability.availableFaces,
		returned_families: entries.length,
		filters: {
			role,
			layer,
			include_unavailable: includeUnavailable,
			...text === void 0 ? {} : { text }
		},
		recommendations,
		fonts: entries,
		warnings
	};
}
//#endregion
//#region src/security.ts
const BEARER = /\bBearer\s+[A-Za-z0-9._~+\/-]+=*/gi;
const QUERY_SECRET = /([?&](?:api[_-]?key|access[_-]?token|token|key|secret|password)=)[^&#\s]+/gi;
function redactText(value) {
	return value.replace(BEARER, "Bearer [REDACTED]").replace(QUERY_SECRET, "$1[REDACTED]");
}
function safeErrorMessage(error, maxChars = 2e3) {
	const redacted = redactText(error instanceof Error ? error.message : String(error));
	return redacted.length <= maxChars ? redacted : `${redacted.slice(0, maxChars)}…`;
}
//#endregion
//#region src/subprocess.ts
async function runCollected(subprocess, argv, options) {
	throwIfAborted(options.signal);
	const timeout = new AbortController();
	const timer = setTimeout(() => timeout.abort(new PptError("PPT_RESOURCE_LIMIT", `process timed out after ${options.timeoutMs}ms`)), options.timeoutMs);
	const combined = options.signal === void 0 ? timeout.signal : AbortSignal.any([options.signal, timeout.signal]);
	try {
		const handle = subprocess.spawn({
			argv,
			cwd: options.cwd,
			stdio: {
				stdin: options.stdin === void 0 ? "ignore" : { data: options.stdin },
				stdout: { maxBytes: options.maxOutputBytes },
				stderr: { maxBytes: options.maxOutputBytes }
			},
			graceMs: options.graceMs ?? 2e3,
			signal: combined,
			...options.env === void 0 ? {} : { env: options.env }
		});
		let outcome;
		try {
			outcome = await handle.done;
		} catch (error) {
			if (timeout.signal.aborted) throw timeout.signal.reason;
			if (options.signal?.aborted) throwIfAborted(options.signal);
			throw error;
		}
		if (timeout.signal.aborted) throw timeout.signal.reason;
		if (options.signal?.aborted) throwIfAborted(options.signal);
		const stdout = handle.collected.stdout?.readFrom(0);
		const stderr = handle.collected.stderr?.readFrom(0);
		return {
			exitCode: outcome.exitCode,
			signal: outcome.signal,
			stdout: stdout?.text ?? "",
			stderr: stderr?.text ?? "",
			stdoutTruncated: stdout?.lossy ?? false,
			stderrTruncated: stderr?.lossy ?? false
		};
	} finally {
		clearTimeout(timer);
	}
}
//#endregion
//#region src/browser-security.ts
function ipv4Parts(address) {
	const parts = address.split(".").map(Number);
	return parts.length === 4 && parts.every((part) => Number.isInteger(part) && part >= 0 && part <= 255) ? parts : void 0;
}
function isBlockedIp(address) {
	const normalized = address.toLowerCase().split("%")[0];
	const v4 = ipv4Parts(normalized);
	if (v4 !== void 0) {
		const a = v4[0];
		const b = v4[1];
		return a === 0 || a === 10 || a === 127 || a === 169 && b === 254 || a === 172 && b >= 16 && b <= 31 || a === 192 && b === 168 || a === 100 && b >= 64 && b <= 127 || a >= 224;
	}
	if (isIP(normalized) === 6) {
		const mapped = /^::ffff:([0-9a-f]{1,4}):([0-9a-f]{1,4})$/.exec(normalized);
		if (mapped !== null) {
			const high = Number.parseInt(mapped[1], 16);
			const low = Number.parseInt(mapped[2], 16);
			return isBlockedIp(`${high >> 8}.${high & 255}.${low >> 8}.${low & 255}`);
		}
		return normalized === "::" || normalized === "::1" || normalized.startsWith("fe8") || normalized.startsWith("fe9") || normalized.startsWith("fea") || normalized.startsWith("feb") || normalized.startsWith("fc") || normalized.startsWith("fd") || normalized.startsWith("ff") || normalized.startsWith("::ffff:127.") || normalized.startsWith("::ffff:10.") || normalized.startsWith("::ffff:192.168.") || normalized === "fd00:ec2::254";
	}
	return true;
}
async function validatePublicHttpUrl(input) {
	if (input.length > 2048) throw new PptError("BROWSER_URL_BLOCKED", "URL exceeds 2048 characters");
	let url;
	try {
		url = new URL(input);
	} catch (error) {
		throw new PptError("BROWSER_URL_BLOCKED", `invalid URL: ${input}`, { cause: error });
	}
	if (url.protocol !== "http:" && url.protocol !== "https:") throw new PptError("BROWSER_URL_BLOCKED", `only public HTTP(S) URLs are allowed: ${url.protocol}`);
	if (url.username !== "" || url.password !== "") throw new PptError("BROWSER_URL_BLOCKED", "URLs containing credentials are blocked");
	if (url.port !== "" && url.port !== "80" && url.port !== "443") throw new PptError("BROWSER_URL_BLOCKED", `non-standard port is blocked: ${url.port}`);
	const hostname = url.hostname.toLowerCase().replace(/\.$/, "");
	if (hostname === "localhost" || hostname.endsWith(".localhost") || hostname.endsWith(".local") || hostname.endsWith(".internal")) throw new PptError("BROWSER_URL_BLOCKED", `local hostname is blocked: ${hostname}`);
	const addresses = isIP(hostname) === 0 ? await lookup(hostname, {
		all: true,
		verbatim: true
	}) : [{ address: hostname }];
	if (addresses.length === 0 || addresses.some((entry) => isBlockedIp(entry.address))) throw new PptError("BROWSER_URL_BLOCKED", `URL resolves to a blocked or non-public address: ${hostname}`);
	return url;
}
//#endregion
//#region src/limits.ts
const DEFAULT_LIMITS = Object.freeze({
	maxSlides: 60,
	maxElementsPerSlide: 200,
	maxRedirects: 5,
	maxResponseBytes: 16777216,
	maxImageBytes: 20971520,
	maxImagePixels: 4e7,
	maxPythonMs: 12e4,
	maxPythonOutputChars: 2e4,
	maxBrowserTextChars: 2e4,
	maxGeneratedFiles: 100,
	maxGeneratedFileBytes: 52428800,
	maxToolResultChars: 3e4,
	maxImageSearchResults: 20
});
function boundedInteger(value, name, min, max) {
	if (!Number.isInteger(value) || value < min || value > max) throw new PptError("PPT_RESOURCE_LIMIT", `${name} must be an integer between ${min} and ${max}`);
	return value;
}
//#endregion
//#region src/paths.ts
const DRIVE_ROOTED_PATH = /^[A-Za-z]:[\\/]/;
const ROOTED_PATH = /^[\\/]/;
function isLocalFilesystemPath(input) {
	return DRIVE_ROOTED_PATH.test(input) || ROOTED_PATH.test(input);
}
function isPathInside(root, target) {
	const rel = relative(resolve(root), resolve(target));
	return rel === "" || !rel.startsWith(`..${sep}`) && rel !== ".." && !isAbsolute(rel);
}
function workspaceRelative(root, target) {
	if (!isPathInside(root, target)) throw new PptError("PPT_PATH_OUTSIDE_WORKSPACE", `path is outside workspace: ${target}`);
	return relative(resolve(root), resolve(target)).split(sep).join("/") || ".";
}
async function nearestExistingParent(path) {
	let cursor = resolve(path);
	for (;;) try {
		return await realpath(cursor);
	} catch {
		const parent = dirname(cursor);
		if (parent === cursor) throw new PptError("PPT_PATH_INVALID", `cannot resolve path parent: ${path}`);
		cursor = parent;
	}
}
async function resolveWorkspacePath(workspaceRoot, input, options = {}) {
	if (typeof input !== "string" || input.trim().length === 0 || input.includes("\0")) throw new PptError("PPT_PATH_INVALID", "path must be a non-empty string without NUL bytes");
	const root = await realpath(resolve(workspaceRoot));
	const candidate = resolve(root, input);
	if (!isPathInside(root, candidate)) throw new PptError("PPT_PATH_OUTSIDE_WORKSPACE", `path is outside workspace: ${input}`);
	if (!isPathInside(root, await nearestExistingParent(candidate))) throw new PptError("PPT_PATH_OUTSIDE_WORKSPACE", `path resolves outside workspace through a symlink: ${input}`);
	if (options.createParent) await mkdir(dirname(candidate), { recursive: true });
	if (options.mustExist) {
		let stat;
		try {
			stat = await lstat(candidate);
		} catch (error) {
			throw new PptError("PPT_PATH_INVALID", `path does not exist: ${input}`, { cause: error });
		}
		const kind = options.kind ?? "either";
		if (kind === "file" && !stat.isFile()) throw new PptError("PPT_PATH_INVALID", `path is not a regular file: ${input}`);
		if (kind === "directory" && !stat.isDirectory()) throw new PptError("PPT_PATH_INVALID", `path is not a directory: ${input}`);
		const resolvedExisting = await realpath(candidate);
		if (!isPathInside(root, resolvedExisting)) throw new PptError("PPT_PATH_OUTSIDE_WORKSPACE", `path resolves outside workspace: ${input}`);
		return resolvedExisting;
	}
	return candidate;
}
//#endregion
//#region src/atomic.ts
async function atomicWriteFile(target, data, options = {}) {
	throwIfAborted(options.signal);
	const directory = dirname(target);
	await mkdir(directory, { recursive: true });
	const temporary = join(directory, `.${basename(target)}.${process.pid}.${randomUUID()}.tmp`);
	let committed = false;
	try {
		const handle = await open(temporary, constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY, options.mode ?? 384);
		try {
			await handle.writeFile(data);
			await handle.sync();
		} finally {
			await handle.close();
		}
		throwIfAborted(options.signal);
		if (options.overwrite) await rename(temporary, target);
		else {
			try {
				await link(temporary, target);
			} catch (error) {
				if (error.code === "EEXIST") throw new PptError("PPT_OUTPUT_EXISTS", `output already exists: ${target}`);
				throw error;
			}
			await unlink(temporary);
		}
		committed = true;
	} finally {
		if (!committed) await unlink(temporary).catch(() => void 0);
	}
}
function atomicWriteText(target, text, options) {
	return atomicWriteFile(target, text, options);
}
function atomicWriteJson(target, value, options) {
	return atomicWriteText(target, `${JSON.stringify(value, null, 2)}\n`, options);
}
//#endregion
//#region src/artifacts.ts
function slugify(value) {
	return value.normalize("NFKC").toLowerCase().replace(/[^\p{Letter}\p{Number}]+/gu, "-").replace(/^-+|-+$/g, "").slice(0, 80) || "presentation";
}
async function allocateArtifactDirectory(workspace, title, outputRoot = "ppt-output") {
	const resolvedOutputRoot = await resolveWorkspacePath(workspace, outputRoot, { createParent: true });
	await mkdir(resolvedOutputRoot, { recursive: true });
	const base = slugify(title);
	let root;
	for (let suffix = 1; suffix <= 1e4; suffix += 1) {
		const candidate = join(resolvedOutputRoot, suffix === 1 ? base : `${base}-${suffix}`);
		try {
			await mkdir(candidate, { recursive: false });
			root = candidate;
			break;
		} catch (error) {
			if (error.code !== "EEXIST") throw error;
		}
	}
	if (root === void 0) throw new PptError("PPT_OUTPUT_EXISTS", `could not allocate artifact directory for ${base}`);
	const assets = join(root, "assets");
	const images = join(assets, "images");
	const preview = join(root, "preview");
	await Promise.all([mkdir(images, { recursive: true }), mkdir(preview, { recursive: true })]);
	const paths = {
		root,
		outline: join(root, "outline.json"),
		designPlan: join(root, "design-plan.json"),
		html: join(root, "deck.html"),
		pptx: join(root, "deck.pptx"),
		assets,
		images,
		sourceManifest: join(assets, "source-manifest.json"),
		preview,
		report: join(root, "report.json"),
		visualReview: join(root, "visual-review.json")
	};
	await atomicWriteJson(paths.sourceManifest, {
		version: 1,
		assets: []
	});
	return paths;
}
//#endregion
//#region src/image-search.ts
/** Actionable alternatives carried by a degraded result when the free providers deliver nothing. */
const IMAGE_SEARCH_FALLBACKS = [
	{
		id: "local-assets",
		tool: "read_image",
		summary: "reuse a raster already frozen under assets/images or listed in assets/source-manifest.json"
	},
	{
		id: "browser-capture",
		tool: "browser_visit",
		summary: "capture the visual anchor from a public page through the browser instead of the image providers"
	},
	{
		id: "custom-provider",
		tool: "image_search",
		summary: "inject an extra zero-configuration backend through the ImageSearchRuntime providers option"
	}
];
const FAILURE_KINDS = /* @__PURE__ */ new Set([
	"cancelled",
	"timeout",
	"rate_limited",
	"server_error",
	"invalid_response",
	"network_error"
]);
const DEGRADED_UNAVAILABLE = "degraded:unavailable";
const DEGRADED_EMPTY = "degraded:no_results";
const DEGRADED_PARTIAL = "degraded:partial";
const PROVIDER_NAME_PATTERN = /^[a-z0-9][a-z0-9-]{0,31}$/;
function fallbackWarnings() {
	return IMAGE_SEARCH_FALLBACKS.map((item) => `fallback:${item.id}`);
}
/**
* Rebuilds the structured story behind the machine-readable warnings of one search result.
* It takes the structural subset it actually reads, so a payload derived from the tool
* output schema (where `license_verified` widens to `boolean`) stays assignable.
*/
function describeImageSearchDegradation(result) {
	const failures = result.warnings.flatMap((warning) => {
		const separator = warning.indexOf(":");
		if (separator <= 0) return [];
		const provider = warning.slice(0, separator);
		const kind = warning.slice(separator + 1);
		if (!result.providers_used.includes(provider) || !FAILURE_KINDS.has(kind)) return [];
		return [{
			provider,
			kind
		}];
	});
	return {
		status: result.warnings.includes(DEGRADED_UNAVAILABLE) || result.results.length === 0 ? "unavailable" : failures.length > 0 || result.results.length < result.count ? "partial" : "ok",
		requested: result.count,
		returned: result.results.length,
		failures,
		fallbacks: IMAGE_SEARCH_FALLBACKS.filter((item) => result.warnings.includes(`fallback:${item.id}`))
	};
}
const OPENVERSE_ENDPOINT = "https://api.openverse.org/v1/images/";
const COMMONS_ENDPOINT = "https://commons.wikimedia.org/w/api.php";
const MAX_PROVIDER_BYTES = 2097152;
const ADULT_PATTERN = /(?:\bporn\b|\bnsfw\b|\bsexually explicit\b|\bnude\b|色情|成人内容|裸体)/iu;
function text$1(value, max = 500) {
	if (typeof value !== "string") return void 0;
	const normalized = value.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
	return normalized.length === 0 ? void 0 : normalized.slice(0, max);
}
function positiveInteger(value) {
	return typeof value === "number" && Number.isInteger(value) && value > 0 ? value : void 0;
}
function matchesOrientation(width, height, orientation) {
	if (orientation === "any" || width === void 0 || height === void 0) return true;
	const ratio = width / height;
	if (orientation === "square") return ratio >= .9 && ratio <= 1.1;
	return orientation === "landscape" ? ratio > 1.1 : ratio < .9;
}
async function readBoundedBytes(response, limit) {
	const declared = Number(response.headers.get("content-length") ?? 0);
	if (Number.isFinite(declared) && declared > limit) throw new PptError("PPT_RESOURCE_LIMIT", `response exceeds ${limit} bytes`);
	if (response.body === null) {
		const bytes = new Uint8Array(await response.arrayBuffer());
		if (bytes.byteLength > limit) throw new PptError("PPT_RESOURCE_LIMIT", `response exceeds ${limit} bytes`);
		return bytes;
	}
	const reader = response.body.getReader();
	const chunks = [];
	let total = 0;
	try {
		while (true) {
			const chunk = await reader.read();
			if (chunk.done) break;
			total += chunk.value.byteLength;
			if (total > limit) {
				await reader.cancel("response size limit exceeded").catch(() => void 0);
				throw new PptError("PPT_RESOURCE_LIMIT", `response exceeds ${limit} bytes`);
			}
			chunks.push(chunk.value);
		}
	} finally {
		reader.releaseLock();
	}
	const bytes = new Uint8Array(total);
	let offset = 0;
	for (const chunk of chunks) {
		bytes.set(chunk, offset);
		offset += chunk.byteLength;
	}
	return bytes;
}
async function readJson(response) {
	const bytes = await readBoundedBytes(response, MAX_PROVIDER_BYTES);
	return JSON.parse(new TextDecoder().decode(bytes));
}
async function safeCandidate(candidate, orientation, validateUrl) {
	if (ADULT_PATTERN.test(`${candidate.title} ${candidate.attribution ?? ""}`)) return void 0;
	if (!matchesOrientation(candidate.width, candidate.height, orientation)) return void 0;
	try {
		const [image, source] = await Promise.all([validateUrl(candidate.image_url), validateUrl(candidate.source_page)]);
		let thumbnail;
		if (candidate.thumbnail_url !== void 0) thumbnail = (await validateUrl(candidate.thumbnail_url)).href;
		return {
			...candidate,
			image_url: image.href,
			source_page: source.href,
			...thumbnail === void 0 ? {} : { thumbnail_url: thumbnail }
		};
	} catch {
		return;
	}
}
function upstreamCategory(error) {
	if (error instanceof PptError && error.code === "PPT_ABORTED") return "cancelled";
	if (error instanceof DOMException && ["AbortError", "TimeoutError"].includes(error.name)) return "timeout";
	const message = error instanceof Error ? error.message : String(error);
	if (/HTTP 429/.test(message)) return "rate_limited";
	if (/HTTP 5\d\d/.test(message)) return "server_error";
	if (/JSON|parse/i.test(message)) return "invalid_response";
	return "network_error";
}
var ImageSearchRuntime = class {
	fetcher;
	validateUrl;
	cache = /* @__PURE__ */ new Map();
	backends;
	strictOnUnavailable;
	constructor(fetcher = fetch, validateUrl = validatePublicHttpUrl, options = {}) {
		this.fetcher = fetcher;
		this.validateUrl = validateUrl;
		const injected = options.providers ?? [];
		for (const backend of injected) if (!PROVIDER_NAME_PATTERN.test(backend.name) || backend.name === "openverse" || backend.name === "wikimedia-commons") throw new PptError("IMAGE_SEARCH_FAILED", `invalid extra image provider name: ${backend.name}`);
		this.backends = [
			{
				name: "openverse",
				search: (query, amount, orientation, signal) => this.openverse(query, amount, orientation, signal)
			},
			{
				name: "wikimedia-commons",
				search: (query, amount, orientation, signal) => this.commons(query, amount, orientation, signal)
			},
			...injected
		];
		this.strictOnUnavailable = options.strictOnUnavailable === true;
	}
	async search(queryInput, countInput = 8, orientation = "any", signal) {
		throwIfAborted(signal);
		const query = queryInput.normalize("NFKC").replace(/\s+/g, " ").trim();
		if ([...query].length < 1 || [...query].length > 160) throw new PptError("IMAGE_SEARCH_FAILED", "query must contain 1..160 Unicode code points");
		const count = boundedInteger(countInput, "count", 1, 12);
		if (![
			"landscape",
			"portrait",
			"square",
			"any"
		].includes(orientation)) throw new PptError("IMAGE_SEARCH_FAILED", `unsupported orientation: ${orientation}`);
		if (ADULT_PATTERN.test(query)) throw new PptError("IMAGE_SEARCH_FAILED", "adult-content queries are blocked");
		const key = JSON.stringify([
			query.toLocaleLowerCase(),
			count,
			orientation
		]);
		const cached = this.cache.get(key);
		if (cached !== void 0 && cached.expires > Date.now()) return {
			...structuredClone(cached.result),
			cache_hit: true
		};
		const providersUsed = [];
		const warnings = [];
		const failures = [];
		const candidates = [];
		for (const backend of this.backends) {
			if (candidates.length >= count) break;
			providersUsed.push(backend.name);
			try {
				candidates.push(...await backend.search(query, Math.min(40, count * 3), orientation, signal));
			} catch (error) {
				if (signal?.aborted) throwIfAborted(signal);
				const kind = upstreamCategory(error);
				failures.push({
					provider: backend.name,
					kind
				});
				warnings.push(`${backend.name}:${kind}`);
			}
		}
		const seen = /* @__PURE__ */ new Set();
		const results = candidates.filter((item) => !seen.has(item.image_url) && seen.add(item.image_url)).slice(0, count);
		if (results.length === 0) {
			const unreachable = providersUsed.length > 0 && failures.length >= providersUsed.length;
			warnings.push(unreachable ? DEGRADED_UNAVAILABLE : DEGRADED_EMPTY);
			warnings.push(...fallbackWarnings());
			warnings.push(`insufficient_results:0/${count}`);
			if (this.strictOnUnavailable && unreachable) throw new PptError("IMAGE_SEARCH_FAILED", "no zero-configuration image provider is reachable", { details: {
				status: "unavailable",
				failures,
				fallbacks: IMAGE_SEARCH_FALLBACKS
			} });
			return {
				query,
				count,
				orientation,
				cache_hit: false,
				providers_used: providersUsed,
				warnings,
				results
			};
		}
		if (failures.length > 0) warnings.push(DEGRADED_PARTIAL);
		if (results.length < count) warnings.push(`insufficient_results:${results.length}/${count}`);
		const stored = {
			query,
			count,
			orientation,
			providers_used: providersUsed,
			warnings,
			results
		};
		this.cache.set(key, {
			expires: Date.now() + 6e5,
			result: structuredClone(stored)
		});
		return {
			...stored,
			cache_hit: false
		};
	}
	async request(url, signal) {
		const timeout = AbortSignal.timeout(15e3);
		const combined = signal === void 0 ? timeout : AbortSignal.any([signal, timeout]);
		const response = await this.fetcher(url, {
			signal: combined,
			redirect: "error",
			headers: { accept: "application/json" }
		});
		if (!response.ok) throw new Error(`HTTP ${response.status}`);
		return readJson(response);
	}
	async openverse(query, amount, orientation, signal) {
		const url = new URL(OPENVERSE_ENDPOINT);
		url.searchParams.set("q", query);
		url.searchParams.set("page_size", String(amount));
		url.searchParams.set("mature", "false");
		const body = await this.request(url, signal);
		const raw = Array.isArray(body.results) ? body.results : [];
		return (await Promise.all(raw.map(async (item) => {
			const row = item;
			if (row.mature === true) return void 0;
			const image = text$1(row.url, 2048);
			const source = text$1(row.foreign_landing_url, 2048);
			if (image === void 0 || source === void 0) return void 0;
			return safeCandidate({
				image_url: image,
				source_page: source,
				provider: "openverse",
				title: text$1(row.title) ?? "Untitled image",
				license: text$1(row.license) ?? "unknown",
				license_verified: false,
				...text$1(row.thumbnail, 2048) === void 0 ? {} : { thumbnail_url: text$1(row.thumbnail, 2048) },
				...positiveInteger(row.width) === void 0 ? {} : { width: positiveInteger(row.width) },
				...positiveInteger(row.height) === void 0 ? {} : { height: positiveInteger(row.height) },
				...text$1(row.mime_type, 100) === void 0 ? {} : { mime_type: text$1(row.mime_type, 100) },
				...text$1(row.creator) === void 0 ? {} : { author: text$1(row.creator) },
				...text$1(row.license_url, 2048) === void 0 ? {} : { license_url: text$1(row.license_url, 2048) },
				...text$1(row.attribution) === void 0 ? {} : { attribution: text$1(row.attribution) }
			}, orientation, this.validateUrl);
		}))).filter((item) => item !== void 0);
	}
	async commons(query, amount, orientation, signal) {
		const url = new URL(COMMONS_ENDPOINT);
		for (const [key, value] of Object.entries({
			action: "query",
			format: "json",
			origin: "*",
			generator: "search",
			gsrnamespace: "6",
			gsrsearch: `file:${query}`,
			gsrlimit: String(amount),
			prop: "imageinfo|info",
			inprop: "url",
			iiprop: "url|size|mime|extmetadata",
			iiurlwidth: "1280"
		})) url.searchParams.set(key, value);
		const body = await this.request(url, signal);
		const pages = Object.values(body.query?.pages ?? {});
		return (await Promise.all(pages.map(async (item) => {
			const page = item;
			const info = Array.isArray(page.imageinfo) ? page.imageinfo[0] : void 0;
			if (info === void 0) return void 0;
			const meta = info.extmetadata ?? {};
			const image = text$1(info.url, 2048);
			const source = text$1(page.fullurl, 2048);
			if (image === void 0 || source === void 0) return void 0;
			return safeCandidate({
				image_url: image,
				source_page: source,
				provider: "wikimedia-commons",
				title: text$1(meta.ObjectName?.value) ?? text$1(page.title) ?? "Wikimedia Commons image",
				license: text$1(meta.LicenseShortName?.value) ?? "unknown",
				license_verified: false,
				...text$1(info.thumburl, 2048) === void 0 ? {} : { thumbnail_url: text$1(info.thumburl, 2048) },
				...positiveInteger(info.width) === void 0 ? {} : { width: positiveInteger(info.width) },
				...positiveInteger(info.height) === void 0 ? {} : { height: positiveInteger(info.height) },
				...text$1(info.mime, 100) === void 0 ? {} : { mime_type: text$1(info.mime, 100) },
				...text$1(meta.Artist?.value) === void 0 ? {} : { author: text$1(meta.Artist?.value) },
				...text$1(meta.LicenseUrl?.value, 2048) === void 0 ? {} : { license_url: text$1(meta.LicenseUrl?.value, 2048) },
				...text$1(meta.Attribution?.value) === void 0 ? {} : { attribution: text$1(meta.Attribution?.value) }
			}, orientation, this.validateUrl);
		}))).filter((item) => item !== void 0);
	}
};
//#endregion
//#region src/art-direction.ts
const ART_COMPOSITIONS = [
	"hero",
	"editorial-split",
	"asymmetric-split",
	"process",
	"layered",
	"data-focus",
	"quote",
	"full-bleed",
	"closing"
];
const ART_DENSITIES = [
	"low",
	"medium",
	"high"
];
const ART_BACKGROUNDS = [
	"base",
	"inverse",
	"accent",
	"image"
];
const ART_TITLE_TREATMENTS = [
	"statement",
	"question",
	"label",
	"number-led"
];
const ART_ANCHOR_KINDS = [
	"none",
	"typography",
	"image",
	"data",
	"code",
	"diagram"
];
const ART_FRAME_POLICIES = [
	"none",
	"single",
	"grouped"
];
const ART_ROLES = [
	"title",
	"subtitle",
	"body",
	"metric",
	"code",
	"diagram",
	"visual-anchor",
	"supporting",
	"frame"
];
const text = (max) => z.string().transform((value) => value.normalize("NFC").trim()).pipe(z.string().refine((value) => [...value].length >= 1 && [...value].length <= max, `must contain 1..${max} Unicode code points`).refine((value) => !/[\r\n]/u.test(value) && !/<\/?[a-z][^>]*>/iu.test(value), "must not contain newlines or HTML"));
const color$1 = z.string().regex(/^#[0-9A-Fa-f]{6}$/u).transform((value) => value.toUpperCase());
const fontRole = z.strictObject({
	family: text(120),
	weight: z.number().int().min(100).max(900)
});
const visualAnchor = z.strictObject({
	kind: z.enum(ART_ANCHOR_KINDS),
	role: text(80),
	min_area_ratio: z.number().finite().min(.05).max(.9).optional()
}).superRefine((value, context) => {
	if (value.kind === "none" && value.min_area_ratio !== void 0) context.addIssue({
		code: "custom",
		path: ["min_area_ratio"],
		message: "none anchor cannot define min_area_ratio"
	});
	if (value.kind !== "none" && value.min_area_ratio === void 0) context.addIssue({
		code: "custom",
		path: ["min_area_ratio"],
		message: "visual anchor requires min_area_ratio"
	});
});
const artSlide = z.strictObject({
	page: z.number().int().min(1).max(60),
	job: text(160),
	takeaway: text(180),
	composition: z.enum(ART_COMPOSITIONS),
	density: z.enum(ART_DENSITIES),
	background_role: z.enum(ART_BACKGROUNDS),
	title_treatment: z.enum(ART_TITLE_TREATMENTS),
	visual_anchor: visualAnchor,
	frame_policy: z.enum(ART_FRAME_POLICIES),
	allow_intentional_repeat: z.boolean().default(false)
});
const ArtDirectionSchema = z.strictObject({
	version: z.literal(1).default(1),
	concept: text(100),
	audience_effect: text(180),
	palette: z.strictObject({
		background: z.array(color$1).min(1).max(4),
		surface: z.array(color$1).min(1).max(4),
		accent: color$1,
		text: z.array(color$1).min(1).max(4)
	}),
	typography: z.strictObject({
		display: fontRole,
		body: fontRole,
		latin: fontRole,
		code: fontRole
	}),
	rhythm: z.strictObject({
		background_sequence: z.array(z.enum(ART_BACKGROUNDS)).min(1).max(60),
		max_grouped_frame_slides: z.number().int().min(0).max(60).default(2),
		max_same_composition_run: z.number().int().min(1).max(6).default(1)
	}),
	slides: z.array(artSlide).min(1).max(60)
}).superRefine((plan, context) => {
	if (plan.rhythm.background_sequence.length !== plan.slides.length) context.addIssue({
		code: "custom",
		path: ["rhythm", "background_sequence"],
		message: "background_sequence must contain one entry per slide"
	});
	plan.slides.forEach((slide, index) => {
		if (slide.page !== index + 1) context.addIssue({
			code: "custom",
			path: [
				"slides",
				index,
				"page"
			],
			message: `page must be ${index + 1}`
		});
		if (plan.rhythm.background_sequence[index] !== slide.background_role) context.addIssue({
			code: "custom",
			path: [
				"slides",
				index,
				"background_role"
			],
			message: "background_role must match rhythm.background_sequence"
		});
	});
});
function validateArtDirection(value, expectedPages) {
	const result = ArtDirectionSchema.safeParse(value);
	if (!result.success) throw new PptError("PPT_ART_DIRECTION_INVALID", "PPT art direction validation failed", { details: { issues: result.error.issues.map((issue) => ({
		path: issue.path.join("."),
		message: issue.message
	})) } });
	if (expectedPages !== void 0 && result.data.slides.length !== expectedPages) throw new PptError("PPT_ART_DIRECTION_INVALID", "PPT art direction page count does not match outline", { details: { issues: [{
		path: "slides",
		message: `expected ${expectedPages} pages, received ${result.data.slides.length}`
	}] } });
	return result.data;
}
function artDirectionFindings(plan) {
	const findings = [];
	const grouped = plan.slides.filter((slide) => slide.frame_policy === "grouped");
	if (grouped.length > plan.rhythm.max_grouped_frame_slides) findings.push({
		code: "ART_GROUPED_FRAMES_OVER_BUDGET",
		severity: "warning",
		message: `${grouped.length} grouped-frame slides exceed the planned maximum of ${plan.rhythm.max_grouped_frame_slides}`,
		page: grouped[plan.rhythm.max_grouped_frame_slides]?.page
	});
	let run = 1;
	for (let index = 1; index < plan.slides.length; index += 1) {
		const current = plan.slides[index];
		const previous = plan.slides[index - 1];
		run = current.composition === previous.composition ? run + 1 : 1;
		if (run > plan.rhythm.max_same_composition_run && !current.allow_intentional_repeat) findings.push({
			code: "ART_COMPOSITION_REPEATED",
			severity: "warning",
			message: `composition ${current.composition} repeats beyond the planned run length`,
			page: current.page
		});
	}
	if (plan.slides.length >= 6 && new Set(plan.slides.map((slide) => slide.composition)).size < 3) findings.push({
		code: "ART_COMPOSITION_VARIETY_LOW",
		severity: "warning",
		message: "decks with six or more slides should use at least three composition families"
	});
	return findings;
}
function artDirectionReviewChecklist(plan) {
	const checklist = [`Confirm the deck expresses the art direction concept “${plan.concept}” and intended audience effect “${plan.audience_effect}”.`, "Compare adjacent slides for deliberate rhythm rather than accidental repetition of silhouette, density, background, or card structure."];
	for (const slide of plan.slides) {
		const anchor = slide.visual_anchor.kind === "none" ? "no visual anchor" : `${slide.visual_anchor.kind} visual anchor`;
		checklist.push(`Page ${slide.page}: verify ${slide.composition}, ${slide.density} density, ${slide.background_role} background, ${anchor}, and ${slide.frame_policy} frame policy deliver “${slide.takeaway}”.`);
	}
	return checklist.slice(0, 30);
}
//#endregion
//#region src/ir.ts
const SLIDE_WIDTH_IN = 13.333333;
const SLIDE_HEIGHT_IN = 7.5;
function roundFixed(value, digits = 6) {
	const factor = 10 ** digits;
	return Math.round((value + Number.EPSILON) * factor) / factor;
}
function pxToInches(value) {
	return roundFixed(value / 96);
}
function pxToPoints(value) {
	return roundFixed(value * .75, 2);
}
//#endregion
//#region src/themes.ts
const HEX = /^#[0-9A-F]{6}$/u;
/** WCAG 2.1 relative luminance of an #RRGGBB colour. */
function relativeLuminance(color) {
	const hex = color.replace("#", "");
	const channel = (offset) => {
		const value = Number.parseInt(hex.slice(offset, offset + 2), 16) / 255;
		return value <= .03928 ? value / 12.92 : ((value + .055) / 1.055) ** 2.4;
	};
	return .2126 * channel(0) + .7152 * channel(2) + .0722 * channel(4);
}
/** WCAG 2.1 contrast ratio between two #RRGGBB colours, from 1 to 21. */
function contrastRatio(a, b) {
	const first = relativeLuminance(a);
	const second = relativeLuminance(b);
	const lighter = Math.max(first, second);
	const darker = Math.min(first, second);
	return (lighter + .05) / (darker + .05);
}
/**
* Contrast is the one part of "looks good" that can be checked by machine, so
* every theme is held to it at load time rather than discovered in review.
*
* A palette is a set of tonal groups, not a cross product: `background[i]`,
* `surface[i]` and `text[i]` belong together, so text is paired by index. The
* accent is a single value that must stay readable on the primary group; on an
* inverted group it is expected to be swapped for the theme's inverted accent,
* which is reported as a warning rather than a failure.
*/
function themeFindings(theme) {
	const findings = [];
	const check = (label, foreground, background, minimum) => {
		const ratio = contrastRatio(foreground, background);
		if (ratio < minimum) findings.push({
			code: "THEME_CONTRAST_LOW",
			severity: "error",
			message: `${theme.id}: ${label} contrast ${ratio.toFixed(2)}:1 is below ${minimum}:1 (${foreground} on ${background})`
		});
	};
	const groups = Math.max(theme.palette.background.length, theme.palette.surface.length, theme.palette.text.length);
	for (let index = 0; index < groups; index += 1) {
		const background = theme.palette.background[index % theme.palette.background.length];
		const surface = theme.palette.surface[index % theme.palette.surface.length];
		const text = theme.palette.text[index % theme.palette.text.length];
		check(`body text ${text} on background ${background}`, text, background, 4.5);
		check(`body text ${text} on surface ${surface}`, text, surface, 4.5);
		if (index === 0) {
			check(`accent on background ${background}`, theme.palette.accent, background, 4.5);
			check(`accent on surface ${surface}`, theme.palette.accent, surface, 4.5);
			continue;
		}
		check(`inverted accent on background ${background}`, theme.palette.accent_inverted, background, 4.5);
		check(`inverted accent on surface ${surface}`, theme.palette.accent_inverted, surface, 4.5);
	}
	return findings;
}
function requireColor(theme, label, value) {
	if (!HEX.test(value)) throw new PptError("PPT_THEME_INVALID", `theme ${theme} ${label} must be #RRGGBB: ${value}`);
}
function requireFont(theme, role, family) {
	if (registeredFont(family) === void 0) throw new PptError("PPT_THEME_INVALID", `theme ${theme} ${role} font is not in the approved registry: ${family}`);
}
function validateTheme(theme) {
	if (!/^[a-z0-9][a-z0-9-]{1,40}$/u.test(theme.id)) throw new PptError("PPT_THEME_INVALID", `theme id must be lower-kebab-case: ${theme.id}`);
	if (theme.palette.background.length === 0 || theme.palette.background.length > 4) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} needs 1..4 background colours`);
	if (theme.palette.surface.length === 0 || theme.palette.surface.length > 4) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} needs 1..4 surface colours`);
	if (theme.palette.text.length === 0 || theme.palette.text.length > 4) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} needs 1..4 text colours`);
	theme.palette.background.forEach((value, index) => requireColor(theme.id, `background[${index}]`, value));
	theme.palette.surface.forEach((value, index) => requireColor(theme.id, `surface[${index}]`, value));
	theme.palette.text.forEach((value, index) => requireColor(theme.id, `text[${index}]`, value));
	requireColor(theme.id, "accent", theme.palette.accent);
	for (const role of [
		"display",
		"body",
		"latin",
		"code"
	]) {
		const entry = theme.typography[role];
		requireFont(theme.id, role, entry.family);
		if (!Number.isInteger(entry.weight) || entry.weight < 100 || entry.weight > 900 || entry.weight % 100 !== 0) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} ${role} weight must be a 100..900 multiple of 100: ${entry.weight}`);
	}
	if (theme.composition_cycle.length === 0) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} needs at least one composition`);
	if (theme.background_cycle.length === 0) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} needs at least one background role`);
	for (const composition of theme.composition_cycle) if (!ART_COMPOSITIONS.includes(composition)) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} has unknown composition: ${composition}`);
	for (const background of theme.background_cycle) if (!ART_BACKGROUNDS.includes(background)) throw new PptError("PPT_THEME_INVALID", `theme ${theme.id} has unknown background role: ${background}`);
	return theme;
}
const USAGE = [
	"ppt_themes returns reusable design recipes. Call it with no arguments to choose one, then pass the chosen theme_id to ppt_outline.",
	"A theme is a starting point for the Art Direction pass, not a substitute for it: keep concept and audience_effect specific to this deck, and write each page job and takeaway from the actual content.",
	"Call ppt_themes with theme_id and page_types to get the visual half of the plan for every page: composition, density, background role, title treatment, frame policy and the exact colours that page should use.",
	"layout_notes and decoration say how each composition is built and which ornament this theme permits, so the HTML follows the recipe instead of inventing one.",
	"Every palette traces to a published design system, named in palette_source. When the deck must follow a client brand instead, read the brand palette from its public style guide with browser_visit, then build the plan on those values and keep the same contrast discipline: body text and accent at 4.5:1 or better against every background and surface, and a separate accent for the inverted tonal group."
].join(" ");
function themeSummary(theme) {
	return {
		id: theme.id,
		name: theme.name,
		concept: theme.concept,
		scenes: [...theme.scenes],
		palette_source: theme.palette_source,
		accent: theme.palette.accent,
		accent_inverted: theme.palette.accent_inverted,
		display_font: theme.typography.display.family,
		body_font: theme.typography.body.family,
		signature: theme.decoration[0] ?? ""
	};
}
function listThemes(themes = PPT_THEMES, scene) {
	const selected = scene === void 0 || scene.trim() === "" ? themes : themes.filter((theme) => theme.scenes.some((entry) => entry.toLowerCase().includes(scene.trim().toLowerCase())));
	return {
		themes: selected.map(themeSummary),
		usage: USAGE,
		warnings: selected.length === 0 ? [`no theme matches scene "${scene ?? ""}"; call ppt_themes without a scene to list every theme`] : []
	};
}
function findTheme(id, themes = PPT_THEMES) {
	const theme = themes.find((candidate) => candidate.id === id);
	if (theme === void 0) throw new PptError("PPT_THEME_UNKNOWN", `unknown theme: ${id}`, { details: { available: themes.map((candidate) => candidate.id) } });
	return theme;
}
/**
* Expand the theme's cycles into one visual plan per page. The caller keeps the
* content decisions (job, takeaway, visual anchor), so this returns only the
* structural half of an Art Direction plan.
*/
function planThemePages(theme, types) {
	const covers = /* @__PURE__ */ new Set(["cover", "section"]);
	let groupedBudget = theme.rhythm_limits.max_grouped_frame_slides;
	let previous;
	return types.map((type, index) => {
		const composition = compositionFor(theme, type, index, previous);
		previous = composition;
		const background = theme.background_cycle[index % theme.background_cycle.length];
		const frame_policy = (composition === "layered" || composition === "data-focus") && groupedBudget > 0 ? "grouped" : "none";
		if (frame_policy === "grouped") groupedBudget -= 1;
		return {
			page: index + 1,
			type,
			composition,
			density: densityFor(composition),
			background_role: background,
			title_treatment: covers.has(type) ? "statement" : "label",
			frame_policy
		};
	});
}
/**
* Which compositions can carry which slide roles. Derived from the outline
* compatibility table in src/outline.ts, so a theme can never hand the outline
* a composition the page type cannot express.
*/
const COMPOSITION_TYPES = {
	hero: ["cover", "section"],
	"editorial-split": [
		"agenda",
		"content",
		"summary"
	],
	"asymmetric-split": [
		"content",
		"comparison",
		"data"
	],
	process: ["process", "timeline"],
	layered: [
		"agenda",
		"content",
		"summary",
		"data"
	],
	"data-focus": ["data"],
	quote: ["quote"],
	"full-bleed": [
		"cover",
		"section",
		"quote",
		"ending"
	],
	closing: ["ending"]
};
/** Used when a theme's cycle carries nothing the page type can express. */
const COMPOSITION_DEFAULTS = {
	cover: "hero",
	agenda: "layered",
	section: "hero",
	content: "editorial-split",
	comparison: "asymmetric-split",
	timeline: "process",
	process: "process",
	data: "data-focus",
	quote: "quote",
	summary: "layered",
	ending: "closing"
};
/**
* Walk the theme's own composition cycle from the current page onwards and take
* the first entry that fits this slide role *and* differs from the previous
* page's composition, so a deck keeps its theme's character instead of being
* forced onto one generic composition per role, and does not repeat a shape
* unless the role leaves the theme no alternative.
*/
function compositionFor(theme, type, index, previous) {
	const cycle = theme.composition_cycle;
	const candidates = [];
	for (let offset = 0; offset < cycle.length; offset += 1) {
		const candidate = cycle[(index + offset) % cycle.length];
		if (COMPOSITION_TYPES[candidate].includes(type) && !candidates.includes(candidate)) candidates.push(candidate);
	}
	if (candidates.length === 0) return COMPOSITION_DEFAULTS[type];
	return candidates.find((candidate) => candidate !== previous) ?? candidates[0];
}
function densityFor(composition) {
	if (composition === "hero" || composition === "quote" || composition === "full-bleed" || composition === "closing") return "low";
	if (composition === "layered" || composition === "data-focus") return "high";
	return "medium";
}
/**
* Confirm an authored plan still belongs to the theme it claims. Palette drift
* is the failure this catches: a deck that names a theme and then uses unrelated
* colours reads as neither.
*/
function themeFindingsForPlan(theme, plan) {
	const findings = [];
	const allowed = /* @__PURE__ */ new Set([
		...theme.palette.background,
		...theme.palette.surface,
		...theme.palette.text,
		theme.palette.accent,
		theme.palette.accent_inverted
	]);
	const declared = [
		plan.palette.accent,
		...plan.palette.background,
		...plan.palette.surface,
		...plan.palette.text
	];
	const foreign = [...new Set(declared.filter((value) => !allowed.has(value)))].sort();
	if (foreign.length > 0) findings.push({
		code: "THEME_PALETTE_DRIFT",
		severity: "warning",
		message: `theme ${theme.id} does not define ${foreign.join(", ")}; keep extra colours out or pick a theme that already carries them`
	});
	if (plan.palette.accent !== theme.palette.accent) findings.push({
		code: "THEME_ACCENT_REPLACED",
		severity: "warning",
		message: `theme ${theme.id} accent is ${theme.palette.accent} but the plan uses ${plan.palette.accent}`
	});
	for (const role of [
		"display",
		"body",
		"latin",
		"code"
	]) if (plan.typography[role].family !== theme.typography[role].family) findings.push({
		code: "THEME_TYPOGRAPHY_REPLACED",
		severity: "warning",
		message: `theme ${theme.id} ${role} font is ${theme.typography[role].family} but the plan uses ${plan.typography[role].family}`
	});
	return findings;
}
/**
* Twelve built-in themes. Every palette clears WCAG AA (4.5:1) for body text
* against every background and surface it defines, and for the accent as well,
* because the accent carries emphasis text and not only rules. The palettes come
* from Carbon, Tailwind, Radix, Fluent 2 and Open Color values; the typography is
* restricted to families this build can actually render (see FONT_REGISTRY).
*/
const PPT_THEMES = [
	{
		id: "carbon-blueprint",
		name: "碳素蓝图",
		concept: "Carbon blueprint clarity",
		audience_effect: "Engineering leadership reads a system diagram, not a slide",
		scenes: [
			"技术方案",
			"架构评审",
			"RFC 评审"
		],
		palette_source: "IBM Carbon gray10/90 + blue70/blue40",
		palette: {
			background: ["#FFFFFF", "#161616"],
			surface: ["#F4F4F4", "#262626"],
			accent: "#0043CE",
			accent_inverted: "#78A9FF",
			text: ["#161616", "#F4F4F4"]
		},
		typography: {
			display: {
				family: "Microsoft YaHei",
				weight: 700
			},
			body: {
				family: "Microsoft YaHei",
				weight: 400
			},
			latin: {
				family: "Segoe UI",
				weight: 600
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"editorial-split",
			"process",
			"data-focus",
			"asymmetric-split",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"inverse",
			"base",
			"base",
			"accent"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Accent rule: an 8px accent bar pinned to the left edge of the content area (position:absolute; left:0; top:0; bottom:0).",
			"Eyebrow label: 2px accent top border, 16px text, letter-spacing 0.12em, text-transform:uppercase.",
			"Cards: surface fill, 1px #E0E0E0 border, 8px radius. There is no shadow in this design system.",
			"Page number bottom right at 14px in #6F6F6F.",
			"Process pages use a 1px connector line with 32px numbered columns."
		],
		layout_notes: {
			hero: "Title at 72px over an accent bar; keep the subtitle to one line.",
			"editorial-split": "7fr/5fr split: claim on the left, evidence card on the right.",
			process: "1px spine across the content area, numbered nodes on a repeat(4,1fr) grid.",
			"data-focus": "Metric at 88px in the latin face with tabular figures, unit at 24px."
		}
	},
	{
		id: "slate-review",
		name: "板岩复盘",
		concept: "Quarterly ledger",
		audience_effect: "Every claim lands on a number the room already trusts",
		scenes: [
			"季度复盘",
			"经营分析",
			"OKR 回顾"
		],
		palette_source: "Tailwind slate50/900 + blue700",
		palette: {
			background: ["#F8FAFC", "#0F172A"],
			surface: ["#E2E8F0", "#1E293B"],
			accent: "#1D4ED8",
			accent_inverted: "#60A5FA",
			text: ["#0F172A", "#F1F5F9"]
		},
		typography: {
			display: {
				family: "Microsoft YaHei",
				weight: 700
			},
			body: {
				family: "Microsoft YaHei",
				weight: 400
			},
			latin: {
				family: "Arial",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"data-focus",
			"editorial-split",
			"process",
			"layered",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"inverse",
			"base",
			"base"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Data pages lead with an 88px metric in the latin face using tabular figures above a 1px baseline rule.",
			"Tables: left-aligned header at weight 400, 1px #E2E8F0 row rules, no vertical lines.",
			"Card groups: surface fill with a 1px border, offset by margin rather than shadow."
		],
		layout_notes: {
			hero: "One line of title and one number; nothing else competes.",
			"data-focus": "Metric block top left, supporting table or chart below with 24px gap.",
			"editorial-split": "Claim left, the quarter-over-quarter numbers right.",
			layered: "Three stacked surface cards, each offset 24px down and right of the last."
		}
	},
	{
		id: "midnight-raise",
		name: "午夜融资",
		concept: "Midnight conviction",
		audience_effect: "The room feels the market window closing and the ask as inevitable",
		scenes: [
			"融资路演",
			"战略汇报",
			"董事会"
		],
		palette_source: "IBM Carbon blue10/blue40 + coolGray90",
		palette: {
			background: ["#0B1220", "#FFFFFF"],
			surface: ["#21272A", "#F2F4F8"],
			accent: "#78A9FF",
			accent_inverted: "#0043CE",
			text: ["#EDF5FF", "#0B1220"]
		},
		typography: {
			display: {
				family: "Microsoft YaHei",
				weight: 700
			},
			body: {
				family: "Microsoft YaHei",
				weight: 400
			},
			latin: {
				family: "Segoe UI",
				weight: 600
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"quote",
			"editorial-split",
			"layered",
			"data-focus",
			"closing"
		],
		background_cycle: [
			"inverse",
			"inverse",
			"accent",
			"base",
			"inverse",
			"base"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"A single 1px cool-grey rule (1160x1) in the top-left of dark pages is the only geometry.",
			"Emphasis is accent text plus an 8px vertical bar; never a filled accent block behind body copy.",
			"Secondary text on dark pages uses #C6C6C6, which clears 10.96:1 against the base ink.",
			"The light inversion page switches the accent role to #0043CE for a 7.79:1 white-background read."
		],
		layout_notes: {
			hero: "Dark full-bleed field, 72px title on the left third, 1px rule above it.",
			quote: "A single line at 44px with the accent bar to its left, nothing else on the page.",
			"data-focus": "One 88px figure per page; the trend line sits under it as a 1px rule.",
			layered: "Two surface cards at 30% and 70% width, stacked with a 24px vertical gap."
		}
	},
	{
		id: "paper-ink",
		name: "纸墨",
		concept: "Paper and ink",
		audience_effect: "The deck reads like a considered essay rather than a sales pitch",
		scenes: [
			"学术报告",
			"白皮书",
			"深度长文"
		],
		palette_source: "dsh-ppt official fixture field #F7F5F0/#111318 + IBM Carbon magenta70",
		palette: {
			background: ["#F7F5F0", "#111318"],
			surface: ["#FFFFFF", "#1C1F26"],
			accent: "#9F1853",
			accent_inverted: "#FF7EB6",
			text: ["#111318", "#FFFFFF"]
		},
		typography: {
			display: {
				family: "SimSun",
				weight: 700
			},
			body: {
				family: "Microsoft YaHei",
				weight: 400
			},
			latin: {
				family: "Times New Roman",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"editorial-split",
			"quote",
			"process",
			"full-bleed",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"base",
			"inverse",
			"base"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"A 6px crimson vertical bar sits to the left of the title block.",
			"A 1px warm-grey divider (height:1px; background:#E5E0DF; margin:24px 0) separates blocks.",
			"Footer citations set small numbers in the latin serif face.",
			"Tint is spent only on the quote page rule; page fields stay the paper colour."
		],
		layout_notes: {
			hero: "Serif title at 72px with the crimson bar left of it; warm paper field behind.",
			"editorial-split": "8fr/4fr: the argument column left, a single pull quote right.",
			quote: "Centred serif paragraph at 32px between two 1px dividers.",
			"full-bleed": "One inverted paper field for the deck midpoint; title only."
		}
	},
	{
		id: "editorial-serif",
		name: "学刊",
		concept: "Scholarly editorial",
		audience_effect: "Reviewers read the argument as a paper with typed evidence",
		scenes: [
			"学术答辩",
			"研究报告",
			"政策解读"
		],
		palette_source: "IBM Carbon warmGray10/90 + orange70",
		palette: {
			background: ["#F7F3F2", "#171414"],
			surface: ["#E5E0DF", "#272525"],
			accent: "#8A3800",
			accent_inverted: "#FF832B",
			text: ["#171414", "#F7F3F2"]
		},
		typography: {
			display: {
				family: "SimSun",
				weight: 700
			},
			body: {
				family: "SimSun",
				weight: 400
			},
			latin: {
				family: "Times New Roman",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"editorial-split",
			"data-focus",
			"quote",
			"process",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"base",
			"inverse",
			"base"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Serif throughout: SimSun for headings and body, Times New Roman for latin and figures.",
			"Tables take a 3px current-colour header underline and 1px row rules.",
			"Section numbers are set as a 240px serif numeral watermark at z-index:0 in a light grey."
		],
		layout_notes: {
			hero: "Serif display at 72px, section numeral watermark behind it.",
			"editorial-split": "Equal halves, serif body at 24px with 1.6 line height.",
			"data-focus": "A table rather than a chart leads this theme: header rule plus row rules.",
			process: "Numbered steps in serif numerals, 1px rule between each."
		}
	},
	{
		id: "telemetry-teal",
		name: "遥测青",
		concept: "Instrument panel",
		audience_effect: "Operators see live signals and the one threshold that moved",
		scenes: [
			"运维复盘",
			"数据看板汇报",
			"性能评审"
		],
		palette_source: "IBM Carbon teal10/60/70/100 + gray100",
		palette: {
			background: ["#FFFFFF", "#081A1C"],
			surface: ["#D9FBFB", "#004144"],
			accent: "#005D5D",
			accent_inverted: "#3DDBD9",
			text: ["#081A1C", "#D9FBFB"]
		},
		typography: {
			display: {
				family: "DengXian",
				weight: 700
			},
			body: {
				family: "DengXian",
				weight: 400
			},
			latin: {
				family: "Segoe UI",
				weight: 600
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"data-focus",
			"process",
			"editorial-split",
			"layered",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"base",
			"inverse",
			"base"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Metric cards: teal-tint fill, 1px border, 24px padding.",
			"Threshold marker: a 2px solid accent rule across the content area.",
			"Step rail: an even grid of 32px numbers joined by a 1px line."
		],
		layout_notes: {
			hero: "Title left, the single threshold rule beneath it.",
			"data-focus": "Four metric cards on a repeat(4,1fr) grid, one highlighted with accent text.",
			process: "Numbered rail with the accent rule marking the current stage.",
			layered: "Two card rows on the dark page, 1px borders visible against the ink."
		}
	},
	{
		id: "graphite-minimal",
		name: "石墨极简",
		concept: "Graphite minimal",
		audience_effect: "Nothing decorative competes with the specification",
		scenes: [
			"产品规格",
			"内部对齐",
			"工程文档"
		],
		palette_source: "Radix gray1/gray3/gray12 + indigo11",
		palette: {
			background: ["#FCFCFC", "#202020"],
			surface: ["#F0F0F0", "#393939"],
			accent: "#3A5BC7",
			accent_inverted: "#8DA4F0",
			text: ["#202020", "#FCFCFC"]
		},
		typography: {
			display: {
				family: "Microsoft YaHei",
				weight: 700
			},
			body: {
				family: "Microsoft YaHei",
				weight: 400
			},
			latin: {
				family: "Arial",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"editorial-split",
			"data-focus",
			"layered",
			"process",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"base",
			"base",
			"inverse"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Zero ornament: the whole deck runs on 1px borders and type scale.",
			"Grid guides: 1px #F0F0F0 column rules, used only on hero and layered pages.",
			"Page numbers are set as a 240px grey numeral behind the top-right corner at z-index:0."
		],
		layout_notes: {
			hero: "One 72px line, a 1px rule, and a grey numeral watermark.",
			"editorial-split": "Strict 1fr/1fr halves divided by a 1px vertical rule.",
			"data-focus": "A single table or metric; no chart chrome at all.",
			layered: "Three cards on the same baseline, separated by 1px borders only."
		}
	},
	{
		id: "indigo-launch",
		name: "靛蓝发布",
		concept: "Launch night",
		audience_effect: "The feature set arrives as a staged reveal with a single colour of momentum",
		scenes: [
			"产品发布",
			"新版本宣讲",
			"大会 Keynote"
		],
		palette_source: "Tailwind indigo950/900/800/300/50",
		palette: {
			background: ["#1E1B4B", "#EEF2FF"],
			surface: ["#312E81", "#E0E7FF"],
			accent: "#A5B4FC",
			accent_inverted: "#3730A3",
			text: ["#EEF2FF", "#1E1B4B"]
		},
		typography: {
			display: {
				family: "Microsoft YaHei",
				weight: 700
			},
			body: {
				family: "DengXian",
				weight: 400
			},
			latin: {
				family: "Segoe UI",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"full-bleed",
			"quote",
			"layered",
			"editorial-split",
			"closing"
		],
		background_cycle: [
			"inverse",
			"inverse",
			"accent",
			"base",
			"inverse",
			"base"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Section pages are a full-bleed #312E81 field with a centred 44px title.",
			"Launch pages open with a 96px version number in the latin face.",
			"On light inversion pages the accent role moves to #3730A3 for an 8.88:1 read."
		],
		layout_notes: {
			hero: "Deep indigo field, 72px title left, version number as the only ornament.",
			"full-bleed": "Full-viewport indigo with the section name centred at 44px.",
			quote: "One line of the roadmap promise, centred, at 44px.",
			layered: "Feature cards in E0E7FF on the indigo field, 1px borders."
		}
	},
	{
		id: "amber-academy",
		name: "琥珀课堂",
		concept: "Workshop chalk",
		audience_effect: "Learners follow one numbered step at a time without losing the thread",
		scenes: [
			"培训",
			"工作坊",
			"新人上手"
		],
		palette_source: "Tailwind amber50/100/700/950 + IBM Carbon yellow20",
		palette: {
			background: ["#FFFBEB", "#451A03"],
			surface: ["#FEF3C7", "#78350F"],
			accent: "#92400E",
			accent_inverted: "#FBBF24",
			text: ["#451A03", "#FFFBEB"]
		},
		typography: {
			display: {
				family: "DengXian",
				weight: 700
			},
			body: {
				family: "DengXian",
				weight: 400
			},
			latin: {
				family: "Arial",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"process",
			"process",
			"editorial-split",
			"data-focus",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"base",
			"base",
			"inverse"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 3,
			max_same_composition_run: 2
		},
		decoration: [
			"Step cards: amber-tint fill with a 1px #FDDC69 border.",
			"Number chips: 32px squares, 8px radius, accent fill, light numeral.",
			"This theme deliberately allows two consecutive process pages; the rhythm limit says so."
		],
		layout_notes: {
			hero: "Warm field, 72px title, the workshop promise in one line.",
			process: "Two consecutive step pages are expected; keep three to four steps each.",
			"editorial-split": "Instruction left, worked example right.",
			"data-focus": "A small table of before/after values rather than a chart."
		}
	},
	{
		id: "crimson-brief",
		name: "绯红简报",
		concept: "Crimson brief",
		audience_effect: "One proposition, one number, one deadline, stated with heat",
		scenes: [
			"营销提案",
			"立项申请",
			"一页纸决策"
		],
		palette_source: "Tailwind rose50/700/900/950",
		palette: {
			background: ["#FFFFFF", "#4C0519"],
			surface: ["#FFF1F2", "#881337"],
			accent: "#BE123C",
			accent_inverted: "#FDA4AF",
			text: ["#4C0519", "#FFF1F2"]
		},
		typography: {
			display: {
				family: "Microsoft YaHei",
				weight: 700
			},
			body: {
				family: "Microsoft YaHei",
				weight: 400
			},
			latin: {
				family: "Arial",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"quote",
			"data-focus",
			"editorial-split",
			"asymmetric-split",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"inverse",
			"base",
			"base",
			"inverse"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"The single-proposition page is a quote composition with a 6px crimson bar on the left.",
			"The one decisive number is set at 88px in the accent colour with tabular figures.",
			"The closing page is a full #4C0519 field with #FFE4E6 text."
		],
		layout_notes: {
			hero: "White field, 72px title, the deadline as a single accent line beneath.",
			quote: "The core proposition at 44px beside the crimson bar.",
			"data-focus": "Exactly one figure, 88px, nothing else on the page.",
			closing: "Deep crimson field with the ask as the only text."
		}
	},
	{
		id: "fluent-azure",
		name: "流蓝企业",
		concept: "Fluent azure",
		audience_effect: "An enterprise audience recognizes a governed, reviewable proposal",
		scenes: [
			"企业汇报",
			"客户方案",
			"招标应答"
		],
		palette_source: "Microsoft Fluent 2 brandWeb 10/30/80/160",
		palette: {
			background: ["#FFFFFF", "#0A2E4A"],
			surface: ["#EBF3FC", "#0F548C"],
			accent: "#0F6CBD",
			accent_inverted: "#A9D3F7",
			text: ["#061724", "#EBF3FC"]
		},
		typography: {
			display: {
				family: "Microsoft YaHei",
				weight: 700
			},
			body: {
				family: "Microsoft YaHei",
				weight: 400
			},
			latin: {
				family: "Segoe UI",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"editorial-split",
			"layered",
			"process",
			"data-focus",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"inverse",
			"base",
			"accent"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Fluent 2 feel: 600-weight headings, radii limited to 4 and 8px, spacing on a 4px base.",
			"Section band: an 8px accent strip along the bottom edge of section pages.",
			"Cards: surface fill, 1px border, radius 8, padding 24."
		],
		layout_notes: {
			hero: "Title left with the section band along the bottom edge.",
			"editorial-split": "Proposal claim left, scope card right.",
			layered: "Three scope cards on a repeat(3,1fr) grid with 24px gaps.",
			"data-focus": "Milestone table with a 1px header rule."
		}
	},
	{
		id: "moss-annual",
		name: "苔绿年报",
		concept: "Annual moss",
		audience_effect: "Stakeholders read steady, compounding movement instead of a spike",
		scenes: [
			"年度报告",
			"ESG 披露",
			"长期规划"
		],
		palette_source: "Tailwind green50/100/700/900",
		palette: {
			background: ["#F0FDF4", "#14532D"],
			surface: ["#DCFCE7", "#166534"],
			accent: "#15803D",
			accent_inverted: "#86EFAC",
			text: ["#14532D", "#F0FDF4"]
		},
		typography: {
			display: {
				family: "Noto Sans SC",
				weight: 700
			},
			body: {
				family: "Noto Sans SC",
				weight: 400
			},
			latin: {
				family: "Arial",
				weight: 700
			},
			code: {
				family: "DejaVu Sans",
				weight: 400
			}
		},
		composition_cycle: [
			"hero",
			"data-focus",
			"editorial-split",
			"layered",
			"process",
			"closing"
		],
		background_cycle: [
			"base",
			"base",
			"base",
			"base",
			"inverse",
			"base"
		],
		rhythm_limits: {
			max_grouped_frame_slides: 2,
			max_same_composition_run: 1
		},
		decoration: [
			"Long-horizon charts sit on a green-tint band with 1px grid rules.",
			"Data cards: green-tint fill, 1px border, figures in the accent colour.",
			"Section pages are a full #14532D field with #F0FDF4 text."
		],
		layout_notes: {
			hero: "Soft green field, 72px title, the reporting period beneath.",
			"data-focus": "Multi-year comparison; prefer a table of years over a single spike chart.",
			"editorial-split": "Commitment left, measured result right.",
			layered: "Three period cards on the same baseline."
		}
	}
];
//#endregion
//#region src/slide-taxonomy.ts
/**
* Slide taxonomy shared by the outline schema, the art direction schema and the
* built-in theme library. It lives in its own module because all three need it:
* keeping it in `outline.ts` made `outline.ts -> themes.ts -> outline.ts` a
* cycle, which breaks the moment any of those modules is loaded first.
*/
const SLIDE_TYPES = [
	"cover",
	"agenda",
	"section",
	"content",
	"comparison",
	"timeline",
	"process",
	"data",
	"quote",
	"summary",
	"ending"
];
const SLIDE_LAYOUTS = [
	"cover",
	"center",
	"title-content",
	"split",
	"two-column",
	"three-column",
	"grid",
	"hero-image",
	"image-left",
	"image-right",
	"timeline-horizontal",
	"timeline-vertical",
	"process-horizontal",
	"process-vertical",
	"chart-focus",
	"quote-focus",
	"full-bleed",
	"closing"
];
/**
* Which layouts can carry which slide roles. A layout that appears here for the
* wrong role is rejected before anything is rendered, so this table is the
* single authority behind both the schema and the model-facing layout guide.
*/
const SLIDE_TYPE_LAYOUTS = {
	cover: ["cover"],
	center: [
		"cover",
		"section",
		"quote",
		"ending"
	],
	"title-content": [
		"agenda",
		"content",
		"summary"
	],
	split: [
		"content",
		"comparison",
		"data"
	],
	"two-column": [
		"agenda",
		"content",
		"comparison",
		"data",
		"summary"
	],
	"three-column": [
		"agenda",
		"content",
		"summary"
	],
	grid: [
		"agenda",
		"content",
		"data",
		"summary"
	],
	"hero-image": [
		"cover",
		"section",
		"content"
	],
	"image-left": ["content", "quote"],
	"image-right": ["content", "quote"],
	"timeline-horizontal": ["timeline"],
	"timeline-vertical": ["timeline"],
	"process-horizontal": ["process"],
	"process-vertical": ["process"],
	"chart-focus": ["data"],
	"quote-focus": ["quote"],
	"full-bleed": [
		"cover",
		"section",
		"quote",
		"ending"
	],
	closing: ["ending"]
};
/** Slide roles that may legitimately carry no visible content item. */
const CONTENT_OPTIONAL_TYPES = [
	"cover",
	"section",
	"ending"
];
/** Slide roles whose layout must be compatible with an ordered point sequence. */
const SEQUENCED_TYPES = ["timeline", "process"];
//#endregion
//#region src/outline.ts
const noMarkup = (value) => !/[\r\n]/u.test(value) && !/<\/?[a-z][^>]*>/iu.test(value);
function cleanString(max, multiline = false) {
	return z.string().transform((value) => value.normalize("NFC").trim()).pipe(z.string().refine((value) => [...value].length >= 1 && [...value].length <= max, `must contain 1..${max} Unicode code points`).refine((value) => multiline || noMarkup(value), "must not contain newlines or HTML"));
}
function safeReference(max) {
	return cleanString(max).refine((value) => {
		try {
			const url = new URL(value);
			return (url.protocol === "http:" || url.protocol === "https:") && url.username === "" && url.password === "";
		} catch {
			if (isAbsolute(value) || value.includes("\\")) return false;
			return value.split("/").every((segment) => segment !== ".." && segment !== "." && segment.length > 0);
		}
	}, "must be a public HTTP(S) URL or safe workspace-relative path");
}
const Point = z.strictObject({
	kind: z.literal("point"),
	text: cleanString(180),
	label: cleanString(40).optional(),
	group: cleanString(40).optional(),
	level: z.union([z.literal(1), z.literal(2)]).default(1),
	emphasis: z.boolean().default(false)
});
const Data = z.strictObject({
	kind: z.literal("data"),
	label: cleanString(60),
	value: z.union([z.number().finite(), cleanString(40)]),
	unit: cleanString(20).optional(),
	source: safeReference(500).optional(),
	note: cleanString(120).optional(),
	group: cleanString(40).optional(),
	emphasis: z.boolean().default(false)
});
const Image = z.strictObject({
	kind: z.literal("image"),
	role: z.enum([
		"hero",
		"supporting",
		"background",
		"portrait",
		"logo",
		"diagram"
	]),
	intent: cleanString(160),
	query: cleanString(160).optional(),
	asset: safeReference(240).optional(),
	caption: cleanString(120).optional(),
	group: cleanString(40).optional()
}).superRefine((item, context) => {
	if (item.query === void 0 === (item.asset === void 0)) context.addIssue({
		code: "custom",
		message: "image requires exactly one of query or asset"
	});
	if (item.role === "background" && item.caption !== void 0) context.addIssue({
		code: "custom",
		path: ["caption"],
		message: "background images cannot have captions"
	});
});
const Chart = z.strictObject({
	kind: z.literal("chart"),
	chart_type: z.enum([
		"bar",
		"line",
		"area",
		"pie",
		"donut",
		"scatter",
		"bubble",
		"radar",
		"waterfall",
		"funnel",
		"heatmap",
		"table"
	]),
	subject: cleanString(120),
	data_ref: safeReference(240).optional(),
	takeaway: cleanString(180),
	group: cleanString(40).optional()
});
const Note = z.strictObject({
	kind: z.literal("note"),
	purpose: z.enum(["speaker", "production"]),
	text: cleanString(500, true)
});
const OutlineContentItemSchema = z.discriminatedUnion("kind", [
	Point,
	Data,
	Image,
	Chart,
	Note
]);
const Style = z.strictObject({
	layout: z.enum(SLIDE_LAYOUTS),
	background: z.enum([
		"light",
		"dark",
		"accent",
		"image"
	]),
	accent: z.string().regex(/^#[0-9A-Fa-f]{6}$/u).transform((value) => value.toUpperCase()),
	title_font: cleanString(80),
	body_font: cleanString(80),
	visual_direction: cleanString(200)
});
const Slide = z.strictObject({
	page: z.number().int().min(1).max(60),
	type: z.enum(SLIDE_TYPES),
	title: cleanString(80),
	content: z.array(OutlineContentItemSchema).max(12),
	style: Style
}).superRefine((slide, context) => {
	const visible = slide.content.filter((item) => item.kind !== "note");
	const notes = slide.content.filter((item) => item.kind === "note");
	if (visible.length > 8) context.addIssue({
		code: "custom",
		path: ["content"],
		message: "a slide can contain at most 8 visible items"
	});
	if (notes.length > 2) context.addIssue({
		code: "custom",
		path: ["content"],
		message: "a slide can contain at most 2 notes"
	});
	if (![...CONTENT_OPTIONAL_TYPES].includes(slide.type) && visible.length === 0) context.addIssue({
		code: "custom",
		path: ["content"],
		message: `${slide.type} requires at least one visible item`
	});
	const titleLimit = slide.type === "cover" ? 80 : 60;
	if ([...slide.title].length > titleLimit) context.addIssue({
		code: "custom",
		path: ["title"],
		message: `${slide.type} title exceeds ${titleLimit} code points`
	});
	if (!SLIDE_TYPE_LAYOUTS[slide.style.layout].includes(slide.type)) context.addIssue({
		code: "custom",
		path: ["style", "layout"],
		message: `${slide.style.layout} is incompatible with ${slide.type}`
	});
	const images = slide.content.filter((item) => item.kind === "image");
	const backgrounds = images.filter((item) => item.role === "background");
	if (slide.style.background === "image" && backgrounds.length !== 1) context.addIssue({
		code: "custom",
		path: ["content"],
		message: "image background requires exactly one background image item"
	});
	if (slide.style.background !== "image" && backgrounds.length > 0) context.addIssue({
		code: "custom",
		path: ["content"],
		message: "background image item requires style.background=image"
	});
	if ([
		"hero-image",
		"image-left",
		"image-right"
	].includes(slide.style.layout) && images.every((item) => item.role === "background")) context.addIssue({
		code: "custom",
		path: ["content"],
		message: `${slide.style.layout} requires a non-background image`
	});
	const charts = slide.content.filter((item) => item.kind === "chart");
	if (slide.style.layout === "chart-focus" && charts.length !== 1) context.addIssue({
		code: "custom",
		path: ["content"],
		message: "chart-focus requires exactly one chart"
	});
	if (slide.type === "data" && slide.style.layout !== "chart-focus" && charts.length > 2) context.addIssue({
		code: "custom",
		path: ["content"],
		message: "data slides allow at most two charts"
	});
	for (const [index, chart] of slide.content.entries()) {
		if (chart.kind !== "chart" || chart.data_ref !== void 0) continue;
		const dataReady = slide.content.some((item) => item.kind === "data");
		const pending = slide.content.some((item) => item.kind === "note" && item.purpose === "production" && /(?:待补.*数据|data.*pending)/iu.test(item.text));
		if (!dataReady && !pending) context.addIssue({
			code: "custom",
			path: [
				"content",
				index,
				"data_ref"
			],
			message: "chart without data_ref requires a data item or explicit pending-data production note"
		});
	}
	if (slide.type === "comparison") {
		if (new Set(visible.flatMap((item) => "group" in item && item.group !== void 0 ? [item.group] : [])).size < 2) context.addIssue({
			code: "custom",
			path: ["content"],
			message: "comparison requires at least two explicit groups"
		});
	}
	if (SEQUENCED_TYPES.includes(slide.type)) {
		const points = slide.content.filter((item) => item.kind === "point").length;
		if (points < 3 || points > 8) context.addIssue({
			code: "custom",
			path: ["content"],
			message: `${slide.type} requires 3..8 point items`
		});
	}
});
const PptOutlineSchema = z.array(Slide).min(1).max(60).superRefine((slides, context) => {
	slides.forEach((slide, index) => {
		if (slide.page !== index + 1) context.addIssue({
			code: "custom",
			path: [index, "page"],
			message: `page must be ${index + 1}`
		});
	});
});
/**
* Check that an authored outline and plan still belong to the theme they name.
* Selecting a theme and then using unrelated colours or fonts is the failure
* this catches: the deck reads as neither the theme nor the brief.
*/
function themeConformanceFindings(theme, outline, designPlan) {
	const findings = [];
	if (designPlan === void 0) findings.push(`THEME_PLAN_MISSING: theme ${theme.id} was selected but no art_direction was supplied, so nothing enforces the theme`);
	else for (const finding of themeFindingsForPlan(theme, designPlan)) findings.push(`${finding.code}: ${finding.message}`);
	const palette = /* @__PURE__ */ new Set([
		...theme.palette.background,
		...theme.palette.surface,
		...theme.palette.text,
		theme.palette.accent,
		theme.palette.accent_inverted
	]);
	const families = new Set(Object.values(theme.typography).map((role) => role.family));
	for (const slide of outline) {
		if (!palette.has(slide.style.accent)) findings.push(`THEME_ACCENT_DRIFT (page ${slide.page}): style.accent ${slide.style.accent} is not defined by theme ${theme.id}`);
		for (const [role, family] of [["title_font", slide.style.title_font], ["body_font", slide.style.body_font]]) if (!families.has(family)) findings.push(`THEME_FONT_DRIFT (page ${slide.page}, ${role}): ${family} is not one of theme ${theme.id} typography families`);
	}
	return findings;
}
function resolveTheme(themeId) {
	return validateTheme(findTheme(themeId));
}
function bodyText(slide) {
	const values = [];
	for (const item of slide.content) if (item.kind === "point") values.push(item.label ?? "", item.text);
	else if (item.kind === "data") values.push(item.label, String(item.value), item.unit ?? "", item.note ?? "");
	else if (item.kind === "image") values.push(item.caption ?? "", item.intent);
	else if (item.kind === "chart") values.push(item.subject, item.takeaway);
	return values.filter(Boolean).join(" ") || slide.title;
}
function resolveFontPlan(outline, designPlan, options) {
	if (options === void 0) return {
		outline,
		...designPlan === void 0 ? {} : { designPlan },
		warnings: []
	};
	const resolvedOutline = structuredClone(outline);
	const resolvedDesign = designPlan === void 0 ? void 0 : structuredClone(designPlan);
	const warnings = /* @__PURE__ */ new Set();
	for (const slide of resolvedOutline) {
		const title = resolveRegisteredFont(slide.style.title_font, slide.title, options.discovered, options.platform);
		const body = resolveRegisteredFont(slide.style.body_font, bodyText(slide), options.discovered, options.platform);
		slide.style.title_font = title.resolved.name;
		slide.style.body_font = body.resolved.name;
		if (title.warning !== void 0) warnings.add(`FONT_FALLBACK (page ${slide.page}, title): ${title.warning}`);
		if (body.warning !== void 0) warnings.add(`FONT_FALLBACK (page ${slide.page}, body): ${body.warning}`);
	}
	if (resolvedDesign !== void 0) {
		const samples = {
			display: resolvedOutline.map((slide) => slide.title).join(" "),
			body: resolvedOutline.map(bodyText).join(" "),
			latin: "AaZz09",
			code: "AaZz09_{}[]();"
		};
		for (const role of Object.keys(samples)) {
			const requested = resolvedDesign.typography[role].family;
			const resolved = resolveRegisteredFont(requested, samples[role], options.discovered, options.platform);
			resolvedDesign.typography[role].family = resolved.resolved.name;
			if (resolved.warning !== void 0) warnings.add(`FONT_FALLBACK (art_direction.${role}): ${resolved.warning}`);
		}
	}
	return {
		outline: resolvedOutline,
		...resolvedDesign === void 0 ? {} : { designPlan: resolvedDesign },
		warnings: [...warnings]
	};
}
/**
* Compress zod issues into one line a caller can act on. The hosting surface
* renders `PptError.message` but not its `details`, so the offending JSON paths
* must reach the message itself.
*/
function describeOutlineIssues(issues) {
	const rendered = issues.map((issue) => {
		return `${issue.path.length === 0 ? "<root>" : issue.path.join(".")}: ${issue.message.length > 200 ? `${issue.message.slice(0, 197)}...` : issue.message}`;
	});
	const shown = rendered.slice(0, 8);
	return rendered.length > shown.length ? `${shown.join("; ")}; +${rendered.length - shown.length} more` : shown.join("; ");
}
function validatePptOutline(value) {
	const result = PptOutlineSchema.safeParse(value);
	if (result.success) return result.data;
	const issues = result.error.issues.map((issue) => ({
		path: issue.path,
		message: issue.message
	}));
	throw new PptError("PPT_OUTLINE_INVALID", `PPT outline validation failed: ${describeOutlineIssues(issues)}`, { details: { issues: issues.map((issue) => ({
		path: issue.path.join("."),
		message: issue.message
	})) } });
}
async function writePptOutline(workspace, artifactTitle, value, outputRoot = "ppt-output", signal, artDirection, fontResolution, themeId) {
	const validatedOutline = validatePptOutline(value);
	const validatedDesignPlan = artDirection === void 0 ? void 0 : validateArtDirection(artDirection, validatedOutline.length);
	const theme = themeId === void 0 ? void 0 : resolveTheme(themeId);
	const resolved = resolveFontPlan(validatedOutline, validatedDesignPlan, fontResolution);
	const outline = resolved.outline;
	const designPlan = resolved.designPlan;
	const paths = await allocateArtifactDirectory(workspace, artifactTitle, outputRoot);
	try {
		await atomicWriteJson(paths.outline, outline, { signal });
		if (designPlan !== void 0) await atomicWriteJson(paths.designPlan, designPlan, { signal });
	} catch (error) {
		await Promise.all([rm(paths.outline, { force: true }), rm(paths.designPlan, { force: true })]);
		throw error;
	}
	const typeCounts = {};
	const fonts = /* @__PURE__ */ new Set();
	const blocking = [];
	for (const slide of outline) {
		typeCounts[slide.type] = (typeCounts[slide.type] ?? 0) + 1;
		fonts.add(slide.style.title_font);
		fonts.add(slide.style.body_font);
		if (slide.content.some((item) => item.kind === "chart" && item.data_ref === void 0 && !slide.content.some((other) => other.kind === "data"))) blocking.push(`page ${slide.page}: chart data is explicitly pending`);
	}
	if (designPlan !== void 0) for (const role of Object.values(designPlan.typography)) fonts.add(role.family);
	const themeFindings = theme === void 0 ? [] : themeConformanceFindings(theme, outline, designPlan);
	return {
		artifact_dir: workspaceRelative(workspace, paths.root),
		outline_path: workspaceRelative(workspace, paths.outline),
		...designPlan === void 0 ? {} : { design_plan_path: workspaceRelative(workspace, paths.designPlan) },
		design_status: designPlan === void 0 ? "legacy" : "directed",
		page_count: outline.length,
		type_counts: typeCounts,
		fonts: [...fonts].sort(),
		warnings: [
			...resolved.warnings,
			...designPlan === void 0 ? ["ART_DIRECTION_MISSING: legacy outline created without design-plan.json"] : artDirectionFindings(designPlan).map((finding) => `${finding.code}${finding.page === void 0 ? "" : ` (page ${finding.page})`}: ${finding.message}`),
			...themeFindings
		],
		blocking_warnings: blocking,
		...theme === void 0 ? {} : { theme: {
			id: theme.id,
			name: theme.name,
			palette_source: theme.palette_source,
			accent: theme.palette.accent,
			accent_inverted: theme.palette.accent_inverted,
			findings: themeFindings
		} }
	};
}
//#endregion
//#region src/html.ts
const LEAF_KINDS = /* @__PURE__ */ new Set([
	"text",
	"image",
	"shape",
	"svg",
	"table"
]);
/** `NodeFilter.SHOW_TEXT`, inlined so the walker call does not depend on the DOM library's globals. */
const SHOW_TEXT = 4;
const BLOCKED_ELEMENTS = /* @__PURE__ */ new Set([
	"SCRIPT",
	"IFRAME",
	"OBJECT",
	"EMBED",
	"FORM",
	"INPUT",
	"TEXTAREA",
	"SELECT",
	"VIDEO",
	"AUDIO",
	"CANVAS",
	"FOREIGNOBJECT"
]);
const ALLOWED_CSS_PREFIXES = [
	"width",
	"height",
	"min-",
	"max-",
	"box-sizing",
	"position",
	"top",
	"right",
	"bottom",
	"left",
	"display",
	"flex",
	"grid",
	"gap",
	"row-gap",
	"column-gap",
	"align-",
	"justify-",
	"place-",
	"order",
	"margin",
	"padding",
	"font",
	"line-height",
	"letter-spacing",
	"text-",
	"white-space",
	"word-break",
	"overflow",
	"color",
	"background",
	"border",
	"border-radius",
	"opacity",
	"z-index",
	"object-fit",
	"object-position",
	"list-style",
	"vertical-align"
];
const BLOCKED_CSS = /(?:^|[;{])\s*(?:transform|filter|animation|transition|clip-path|mask|mix-blend-mode|perspective|backdrop-filter|box-shadow)\s*:/imu;
const REMOTE_RESOURCE = /(?:url\s*\(\s*['"]?\s*(?:https?:|data:|javascript:|file:)|@import\b)/iu;
function cssProperties(css) {
	const result = /* @__PURE__ */ new Set();
	for (const match of css.matchAll(/(?:^|[;{])\s*([a-z-]+)\s*:/gimu)) result.add(match[1].toLowerCase());
	return [...result];
}
function localReference(value) {
	if (value.trim().length === 0 || value.includes("\0") || value.includes("\\") || isAbsolute(value)) return false;
	try {
		new URL(value);
		return false;
	} catch {}
	return value.split("/").every((segment) => segment !== ".." && segment !== ".");
}
async function validateDeckHtmlSource(workspace, artifactRoot, html, outlineLength, designPlan, strictDesign = false) {
	if (Buffer.byteLength(html) > 5242880) throw new PptError("PPT_RESOURCE_LIMIT", "HTML source exceeds 5 MiB");
	const { Window } = await import("happy-dom");
	const domWindow = new Window();
	domWindow.document.write(html);
	const document = domWindow.document;
	const issues = [];
	const unsupported = /* @__PURE__ */ new Set();
	const allowedFonts = new Set(FONT_REGISTRY.map((font) => font.name));
	const usedFonts = /* @__PURE__ */ new Set();
	const primaryFonts = /* @__PURE__ */ new Set();
	const designFindings = designPlan === void 0 ? [] : artDirectionFindings(designPlan);
	for (const element of document.querySelectorAll("*")) {
		if (BLOCKED_ELEMENTS.has(element.tagName)) issues.push(`blocked element: ${element.tagName.toLowerCase()}`);
		for (const attribute of [...element.attributes]) {
			if (/^on/iu.test(attribute.name)) issues.push(`event handler attribute is blocked: ${attribute.name}`);
			if ([
				"src",
				"href",
				"xlink:href"
			].includes(attribute.name) && attribute.value.trim().length > 0) {
				if (!localReference(attribute.value)) issues.push(`non-local resource is blocked: ${attribute.value.slice(0, 120)}`);
			}
		}
	}
	const css = [...document.querySelectorAll("style")].map((style) => style.textContent ?? "").join("\n") + "\n" + [...document.querySelectorAll("[style]")].map((element) => element.getAttribute("style") ?? "").join("\n");
	if (REMOTE_RESOURCE.test(css)) issues.push("CSS contains a remote, data, script, or file resource");
	if (BLOCKED_CSS.test(css)) issues.push("CSS contains an unsupported effects or animation property");
	for (const property of cssProperties(css)) if (!ALLOWED_CSS_PREFIXES.some((prefix) => prefix.endsWith("-") ? property.startsWith(prefix) : property === prefix || property.startsWith(`${prefix}-`))) unsupported.add(property);
	if (unsupported.size > 0) issues.push(`unsupported CSS properties: ${[...unsupported].sort().join(", ")}`);
	for (const match of css.matchAll(/font-family\s*:\s*([^;}{]+)/gimu)) {
		const families = match[1].split(",").map((value) => value.trim().replace(/^['"]|['"]$/g, "")).filter(Boolean);
		if (families[0] !== void 0) primaryFonts.add(families[0]);
		for (const family of families) {
			usedFonts.add(family);
			if (!allowedFonts.has(family)) issues.push(`unauthorized font: ${family}`);
		}
	}
	const slides = [...document.querySelectorAll(".ppt-slide[data-page]")];
	if (slides.length !== outlineLength) issues.push(`expected ${outlineLength} .ppt-slide elements, found ${slides.length}`);
	const ids = /* @__PURE__ */ new Set();
	slides.forEach((slide, index) => {
		if (slide.dataset.page !== String(index + 1)) issues.push(`slide ${index + 1} has non-contiguous data-page`);
		const planned = designPlan?.slides[index];
		if (planned !== void 0) {
			if (slide.dataset.artComposition !== planned.composition) issues.push(`page ${index + 1} data-art-composition must be ${planned.composition}`);
			if (slide.dataset.artDensity !== planned.density) issues.push(`page ${index + 1} data-art-density must be ${planned.density}`);
			if (slide.dataset.artBackground !== planned.background_role) issues.push(`page ${index + 1} data-art-background must be ${planned.background_role}`);
			const roleElements = [...slide.querySelectorAll("[data-art-role]")];
			for (const element of roleElements) if (!ART_ROLES.includes(element.dataset.artRole ?? "")) issues.push(`page ${index + 1} has invalid data-art-role: ${element.dataset.artRole ?? ""}`);
			const anchors = roleElements.filter((element) => element.dataset.artRole === "visual-anchor");
			if (planned.visual_anchor.kind === "none" && anchors.length !== 0) issues.push(`page ${index + 1} declares no visual anchor but HTML contains ${anchors.length}`);
			if (planned.visual_anchor.kind !== "none" && anchors.length !== 1) issues.push(`page ${index + 1} requires exactly one visual-anchor, found ${anchors.length}`);
			const frames = roleElements.filter((element) => element.dataset.artRole === "frame").length;
			if (planned.frame_policy === "none" && frames > 0) issues.push(`page ${index + 1} frame_policy=none but HTML contains frames`);
			if (planned.frame_policy === "single" && frames > 1) issues.push(`page ${index + 1} frame_policy=single but HTML contains ${frames} frames`);
			if (planned.frame_policy === "grouped" && frames < 2) designFindings.push({
				code: "ART_GROUPED_FRAMES_NOT_REALIZED",
				severity: "warning",
				message: "grouped frame policy is not visibly realized",
				page: index + 1
			});
		}
		const walker = document.createTreeWalker(slide, SHOW_TEXT);
		for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) if ((node.textContent ?? "").trim().length > 0 && node.parentElement?.closest("[data-ppt-id][data-ppt-kind]") === null) {
			issues.push(`page ${index + 1} contains visible text outside a convertible leaf`);
			break;
		}
		for (const leaf of slide.querySelectorAll("[data-ppt-id], [data-ppt-kind]")) {
			const id = leaf.dataset.pptId;
			const kind = leaf.dataset.pptKind;
			if (id === void 0 || !/^[A-Za-z][A-Za-z0-9_.-]{0,79}$/u.test(id)) issues.push(`page ${index + 1} has invalid or missing data-ppt-id`);
			else if (ids.has(id)) issues.push(`duplicate data-ppt-id: ${id}`);
			else ids.add(id);
			if (kind === void 0 || !LEAF_KINDS.has(kind)) issues.push(`page ${index + 1} has invalid or missing data-ppt-kind`);
			if (leaf.querySelector("[data-ppt-id][data-ppt-kind]") !== null) issues.push(`${id ?? "unknown"} is not a leaf`);
			if (kind === "image" && leaf.tagName !== "IMG") issues.push(`${id ?? "unknown"} kind=image must be an img`);
			if (kind === "svg" && leaf.tagName !== "svg") issues.push(`${id ?? "unknown"} kind=svg must be an svg`);
			if (kind === "table" && leaf.tagName !== "TABLE") issues.push(`${id ?? "unknown"} kind=table must be a table`);
			const z = leaf.dataset.pptZ;
			if (z !== void 0 && (!/^-?\d+$/u.test(z) || Math.abs(Number(z)) > 1e4)) issues.push(`${id ?? "unknown"} has invalid data-ppt-z`);
		}
	});
	if (designPlan !== void 0) {
		const plannedFonts = new Set(Object.values(designPlan.typography).map((role) => role.family));
		for (const family of plannedFonts) if (!usedFonts.has(family)) designFindings.push({
			code: "ART_FONT_ROLE_UNUSED",
			severity: "warning",
			message: `planned font family ${family} is not declared in HTML CSS`
		});
		if (strictDesign) {
			for (const finding of designFindings) if (finding.severity === "warning") issues.push(`${finding.code}${finding.page === void 0 ? "" : ` page ${finding.page}`}: ${finding.message}`);
		}
	}
	const assetReferences = /* @__PURE__ */ new Set();
	for (const element of document.querySelectorAll("img[src], svg image[href], svg image[xlink\\:href]")) assetReferences.add(element.getAttribute("src") ?? element.getAttribute("href") ?? element.getAttribute("xlink:href") ?? "");
	for (const match of css.matchAll(/url\s*\(\s*['"]?([^'")]+)['"]?\s*\)/gimu)) assetReferences.add(match[1].trim());
	for (const ref of assetReferences) {
		if (!localReference(ref)) continue;
		try {
			if (!isPathInside(artifactRoot, await resolveWorkspacePath(workspace, join(artifactRoot, ref), {
				mustExist: true,
				kind: "file"
			}))) issues.push(`asset leaves artifact directory: ${ref}`);
		} catch {
			issues.push(`missing or invalid local asset: ${ref}`);
		}
	}
	domWindow.close();
	if (issues.length > 0) throw new PptError("HTML_CREATE_VALIDATION_FAILED", "HTML static validation failed", { details: { issues } });
	return {
		fonts: [...usedFonts].sort(),
		primaryFonts: [...primaryFonts].sort(),
		unsupported: [...unsupported].sort(),
		designFindings
	};
}
async function createHtmlDeck(browser, owner, workspace, outlinePathInput, html, signal, designPlanPathInput, strictDesign = false, fontDirs) {
	throwIfAborted(signal);
	const outlinePath = await resolveWorkspacePath(workspace, outlinePathInput, {
		mustExist: true,
		kind: "file"
	});
	const artifactRoot = dirname(outlinePath);
	const outline = validatePptOutline(JSON.parse(await readFile(outlinePath, "utf8")));
	let designPlan;
	if (designPlanPathInput !== void 0) {
		const designPlanPath = await resolveWorkspacePath(workspace, designPlanPathInput, {
			mustExist: true,
			kind: "file"
		});
		if (dirname(designPlanPath) !== artifactRoot) throw new PptError("HTML_CREATE_INPUT_INVALID", "design plan must be in the same artifact directory as outline.json");
		designPlan = validateArtDirection(JSON.parse(await readFile(designPlanPath, "utf8")), outline.length);
	}
	if (outline.some((slide) => slide.content.some((item) => item.kind === "chart" && item.data_ref === void 0 && !slide.content.some((other) => other.kind === "data")))) throw new PptError("HTML_CREATE_INPUT_INVALID", "outline still contains a chart with explicitly pending data");
	const output = join(artifactRoot, "deck.html");
	try {
		await access(output);
		throw new PptError("PPT_OUTPUT_EXISTS", `output already exists: ${workspaceRelative(workspace, output)}`);
	} catch (error) {
		if (error instanceof PptError) throw error;
	}
	const validation = await validateDeckHtmlSource(workspace, artifactRoot, html, outline.length, designPlan, strictDesign);
	if (fontDirs !== void 0) {
		const discovered = await discoverRegisteredFonts(fontDirs);
		const available = new Set(discovered.map((font) => font.name));
		const unavailable = validation.primaryFonts.filter((font) => !available.has(font));
		if (unavailable.length > 0) throw new PptError("PPT_DEPENDENCY_MISSING", "HTML declares a primary font that is not installed in the approved registry", { details: {
			unavailable,
			available: [...available].sort(),
			scope: "approved_registry"
		} });
	}
	const temporary = join(artifactRoot, `.deck.${process.pid}.${randomUUID()}.html`);
	const previewDirectory = join(artifactRoot, "preview");
	try {
		await atomicWriteText(temporary, html, { signal });
		const rendered = await browser.renderHtmlPreview(owner, workspace, temporary, previewDirectory, outline.length, FONT_REGISTRY.map((font) => font.name), signal);
		if (designPlan !== void 0) {
			for (const page of rendered.designPages) {
				const planned = designPlan.slides[page.page - 1];
				const typographyRole = (role) => {
					if (role === "title" || role === "subtitle") return "display";
					if (role === "body" || role === "supporting" || role === "frame" || role === "diagram") return "body";
					if (role === "metric") return "latin";
					if (role === "code") return "code";
					if (role === "visual-anchor") {
						if (planned.visual_anchor.kind === "typography") return "display";
						if (planned.visual_anchor.kind === "code") return "code";
						if (planned.visual_anchor.kind === "data") return "latin";
					}
				};
				for (const roleStyle of page.roleStyles) {
					const role = typographyRole(roleStyle.role);
					if (role === void 0) continue;
					const expected = designPlan.typography[role];
					if (roleStyle.fontFamily !== expected.family || roleStyle.fontWeight !== expected.weight) validation.designFindings.push({
						code: "ART_TYPOGRAPHY_ROLE_MISMATCH",
						severity: "warning",
						page: page.page,
						message: `${roleStyle.role} uses ${roleStyle.fontFamily} ${roleStyle.fontWeight}, expected ${expected.family} ${expected.weight}`
					});
				}
				const minimum = planned.visual_anchor.min_area_ratio;
				if (minimum !== void 0 && (page.anchorAreaRatio ?? 0) < minimum) validation.designFindings.push({
					code: "ART_VISUAL_ANCHOR_TOO_SMALL",
					severity: "warning",
					page: page.page,
					message: `visual anchor covers ${((page.anchorAreaRatio ?? 0) * 100).toFixed(1)}% of the slide, below the planned ${(minimum * 100).toFixed(1)}%`
				});
				if (page.page > 1 && !planned.allow_intentional_repeat) {
					const previous = rendered.designPages[page.page - 2];
					const intersection = page.occupancy.filter((value, index) => value === 1 && previous.occupancy[index] === 1).length;
					const union = page.occupancy.filter((value, index) => value === 1 || previous.occupancy[index] === 1).length;
					if (union > 0 && intersection / union > .88) validation.designFindings.push({
						code: "ART_SILHOUETTE_REPEATED",
						severity: "warning",
						page: page.page,
						message: "adjacent slide occupancy silhouettes are highly similar"
					});
				}
			}
			if (strictDesign && validation.designFindings.some((finding) => finding.severity === "warning")) throw new PptError("HTML_CREATE_VALIDATION_FAILED", "HTML design fidelity validation failed in strict mode", { details: { issues: validation.designFindings.map((finding) => ({
				code: finding.code,
				page: finding.page,
				message: finding.message
			})) } });
		}
		await atomicWriteText(output, html, { signal });
		const designValidation = join(artifactRoot, "design-validation.json");
		await atomicWriteJson(designValidation, {
			version: 1,
			mode: designPlan === void 0 ? "legacy" : "directed",
			checks: designPlan === void 0 ? [] : [
				"html-art-attributes",
				"art-roles",
				"visual-anchor-area",
				"frame-policy",
				"occupancy-silhouette",
				"font-declarations",
				"computed-typography-roles"
			],
			pages: rendered.designPages,
			findings: validation.designFindings
		}, { signal });
		return {
			html_path: workspaceRelative(workspace, output),
			page_count: outline.length,
			preview_paths: rendered.previews,
			fonts: [.../* @__PURE__ */ new Set([...validation.fonts, ...rendered.fonts])].sort(),
			external_resources: "none",
			warnings: [
				...designPlan === void 0 ? ["ART_DIRECTION_MISSING: HTML validated in legacy design mode"] : [],
				...validation.designFindings.map((finding) => `${finding.code}${finding.page === void 0 ? "" : ` (page ${finding.page})`}: ${finding.message}`),
				...rendered.warnings
			],
			unsupported_css: validation.unsupported,
			design_status: designPlan === void 0 ? "legacy" : "directed",
			design_findings: validation.designFindings,
			design_validation_path: workspaceRelative(workspace, designValidation)
		};
	} finally {
		await rm(temporary, { force: true });
	}
}
//#endregion
//#region src/zip.ts
const ZIP_LOCAL_SIGNATURE = 67324752;
const ZIP_CENTRAL_SIGNATURE = 33639248;
const ZIP_END_SIGNATURE = 101010256;
const ZIP_FLAG_ENCRYPTED = 1;
const ZIP_FLAG_DESCRIPTOR = 8;
const ZIP64_SENTINEL = 4294967295;
/** DOS epoch: a fixed timestamp keeps a rewritten package reproducible. */
const DOS_TIME = 0;
const DOS_DATE = 33;
const CRC32_TABLE = (() => {
	const table = /* @__PURE__ */ new Uint32Array(256);
	for (let index = 0; index < 256; index += 1) {
		let value = index;
		for (let bit = 0; bit < 8; bit += 1) value = (value & 1) === 1 ? 3988292384 ^ value >>> 1 : value >>> 1;
		table[index] = value >>> 0;
	}
	return table;
})();
function crc32(data) {
	let value = 4294967295;
	for (let index = 0; index < data.length; index += 1) value = value >>> 8 ^ CRC32_TABLE[(value ^ data[index]) & 255];
	return (value ^ 4294967295) >>> 0;
}
function findEndOfCentralDirectory(view, length) {
	if (length < 22) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX is not a readable ZIP package");
	for (let offset = length - 22; offset >= 0; offset -= 1) {
		if (view.getUint32(offset, true) !== ZIP_END_SIGNATURE) continue;
		if (offset + 22 + view.getUint16(offset + 20, true) === length) return offset;
	}
	throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX is not a readable ZIP package: end of central directory is missing");
}
function parseZipDirectory(data) {
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	const end = findEndOfCentralDirectory(view, data.byteLength);
	const count = view.getUint16(end + 10, true);
	const comment = data.slice(end + 22, end + 22 + view.getUint16(end + 20, true));
	const entries = [];
	let cursor = view.getUint32(end + 16, true);
	for (let index = 0; index < count; index += 1) {
		if (view.getUint32(cursor, true) !== ZIP_CENTRAL_SIGNATURE) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX zip central directory is malformed");
		const flags = view.getUint16(cursor + 8, true);
		const method = view.getUint16(cursor + 10, true);
		const crc = view.getUint32(cursor + 16, true);
		const compressedSize = view.getUint32(cursor + 20, true);
		const uncompressedSize = view.getUint32(cursor + 24, true);
		const nameLength = view.getUint16(cursor + 28, true);
		const extraLength = view.getUint16(cursor + 30, true);
		const entryCommentLength = view.getUint16(cursor + 32, true);
		const localOffset = view.getUint32(cursor + 42, true);
		const centralSize = 46 + nameLength + extraLength + entryCommentLength;
		if ((flags & ZIP_FLAG_ENCRYPTED) !== 0 || (flags & ZIP_FLAG_DESCRIPTOR) !== 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX zip entries with encryption or data descriptors are unsupported");
		if (compressedSize === ZIP64_SENTINEL || uncompressedSize === ZIP64_SENTINEL || localOffset === ZIP64_SENTINEL) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX zip64 archives are unsupported");
		if (view.getUint32(localOffset, true) !== ZIP_LOCAL_SIGNATURE) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX zip local header is malformed");
		const localNameLength = view.getUint16(localOffset + 26, true);
		const localExtraLength = view.getUint16(localOffset + 28, true);
		const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
		entries.push({
			name: strFromU8(data.subarray(cursor + 46, cursor + 46 + nameLength)),
			method,
			compressed_size: compressedSize,
			uncompressed_size: uncompressedSize,
			local_header_offset: localOffset,
			flags,
			version_needed: view.getUint16(cursor + 6, true),
			time: view.getUint16(cursor + 12, true),
			date: view.getUint16(cursor + 14, true),
			crc,
			local_extra: data.slice(localOffset + 30 + localNameLength, dataOffset),
			data_offset: dataOffset,
			header: data.slice(localOffset, dataOffset + compressedSize),
			central: data.slice(cursor, cursor + centralSize)
		});
		cursor += centralSize;
	}
	return {
		entries,
		comment
	};
}
function readZipDirectory(data) {
	try {
		return parseZipDirectory(data);
	} catch (error) {
		throw asPptError(error, "PPT_CREATE_INVALID_PACKAGE", "PPTX zip directory is not readable");
	}
}
function readEntryBytes(entry, data) {
	const raw = data.subarray(entry.data_offset, entry.data_offset + entry.compressed_size);
	if (entry.method === 0) return raw.slice();
	if (entry.method !== 8) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `unsupported zip compression method ${entry.method} for ${entry.name}`);
	try {
		return decompressSync(raw);
	} catch (error) {
		throw new PptError("PPT_CREATE_INVALID_PACKAGE", `zip entry is not readable: ${entry.name}`, { cause: error });
	}
}
function rebuildLocalEntry(entry, contents) {
	const name = strToU8(entry.name);
	const payload = entry.method === 8 ? deflateSync(contents, { level: 6 }) : contents.slice();
	const crc = crc32(contents);
	const header = new Uint8Array(30 + name.byteLength + entry.local_extra.byteLength + payload.byteLength);
	const view = new DataView(header.buffer);
	view.setUint32(0, ZIP_LOCAL_SIGNATURE, true);
	view.setUint16(4, Math.max(entry.version_needed, 20), true);
	view.setUint16(6, entry.flags, true);
	view.setUint16(8, entry.method, true);
	view.setUint16(10, entry.time, true);
	view.setUint16(12, entry.date, true);
	view.setUint32(14, crc, true);
	view.setUint32(18, payload.byteLength, true);
	view.setUint32(22, contents.byteLength, true);
	view.setUint16(26, name.byteLength, true);
	view.setUint16(28, entry.local_extra.byteLength, true);
	header.set(name, 30);
	header.set(entry.local_extra, 30 + name.byteLength);
	header.set(payload, 30 + name.byteLength + entry.local_extra.byteLength);
	return {
		header,
		crc,
		compressed_size: payload.byteLength,
		uncompressed_size: contents.byteLength
	};
}
/** Builds the local header and central directory record for an appended part. */
function buildAddedEntry(name, contents) {
	const nameBytes = strToU8(name);
	const payload = deflateSync(contents, { level: 6 });
	const crc = crc32(contents);
	const header = new Uint8Array(30 + nameBytes.byteLength + payload.byteLength);
	const view = new DataView(header.buffer);
	view.setUint32(0, ZIP_LOCAL_SIGNATURE, true);
	view.setUint16(4, 20, true);
	view.setUint16(6, 0, true);
	view.setUint16(8, 8, true);
	view.setUint16(10, DOS_TIME, true);
	view.setUint16(12, DOS_DATE, true);
	view.setUint32(14, crc, true);
	view.setUint32(18, payload.byteLength, true);
	view.setUint32(22, contents.byteLength, true);
	view.setUint16(26, nameBytes.byteLength, true);
	view.setUint16(28, 0, true);
	header.set(nameBytes, 30);
	header.set(payload, 30 + nameBytes.byteLength);
	const central = new Uint8Array(46 + nameBytes.byteLength);
	const centralView = new DataView(central.buffer);
	centralView.setUint32(0, ZIP_CENTRAL_SIGNATURE, true);
	centralView.setUint16(4, 20, true);
	centralView.setUint16(6, 20, true);
	centralView.setUint16(8, 0, true);
	centralView.setUint16(10, 8, true);
	centralView.setUint16(12, DOS_TIME, true);
	centralView.setUint16(14, DOS_DATE, true);
	centralView.setUint32(16, crc, true);
	centralView.setUint32(20, payload.byteLength, true);
	centralView.setUint32(24, contents.byteLength, true);
	centralView.setUint16(28, nameBytes.byteLength, true);
	centralView.setUint32(42, 0, true);
	central.set(nameBytes, 46);
	return {
		header,
		central,
		crc,
		compressed_size: payload.byteLength,
		uncompressed_size: contents.byteLength
	};
}
function patchCentralEntry(entry, offset, crc, compressedSize, uncompressedSize) {
	const central = entry.central.slice();
	const view = new DataView(central.buffer, central.byteOffset, central.byteLength);
	view.setUint32(16, crc, true);
	view.setUint32(20, compressedSize, true);
	view.setUint32(24, uncompressedSize, true);
	view.setUint32(42, offset, true);
	return central;
}
function rewriteZip(data, plan) {
	const directory = readZipDirectory(data);
	const present = new Set(directory.entries.map((entry) => entry.name));
	for (const name of plan.additions?.keys() ?? []) if (present.has(name)) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `zip entry already exists: ${name}`);
	const locals = [];
	const centrals = [];
	let offset = 0;
	let count = 0;
	for (const entry of directory.entries) {
		const replacement = plan.transform?.(entry.name, () => readEntryBytes(entry, data));
		if (replacement === null) continue;
		let header = entry.header;
		let crc = entry.crc;
		let compressedSize = entry.compressed_size;
		let uncompressedSize = entry.uncompressed_size;
		if (replacement !== void 0) {
			const rebuilt = rebuildLocalEntry(entry, replacement);
			header = rebuilt.header;
			crc = rebuilt.crc;
			compressedSize = rebuilt.compressed_size;
			uncompressedSize = rebuilt.uncompressed_size;
		}
		locals.push(header);
		centrals.push(patchCentralEntry(entry, offset, crc, compressedSize, uncompressedSize));
		offset += header.byteLength;
		count += 1;
	}
	for (const [name, contents] of plan.additions ?? []) {
		const added = buildAddedEntry(name, contents);
		new DataView(added.central.buffer).setUint32(42, offset, true);
		locals.push(added.header);
		centrals.push(added.central);
		offset += added.header.byteLength;
		count += 1;
	}
	const centralSize = centrals.reduce((total, entry) => total + entry.byteLength, 0);
	const output = new Uint8Array(offset + centralSize + 22 + directory.comment.byteLength);
	let cursor = 0;
	for (const part of [...locals, ...centrals]) {
		output.set(part, cursor);
		cursor += part.byteLength;
	}
	const view = new DataView(output.buffer);
	view.setUint32(cursor, ZIP_END_SIGNATURE, true);
	view.setUint16(cursor + 8, count, true);
	view.setUint16(cursor + 10, count, true);
	view.setUint32(cursor + 12, centralSize, true);
	view.setUint32(cursor + 16, offset, true);
	view.setUint16(cursor + 20, directory.comment.byteLength, true);
	output.set(directory.comment, cursor + 22);
	return output;
}
//#endregion
//#region src/transitions.ts
/**
* Slide transitions live outside the pptxgenjs feature set, so a deck opts in by
* declaring a per-page plan and letting the PPTX writer rewrite only the slide XML
* parts after pptxgenjs produced the package. Every other zip entry is copied
* verbatim so the result stays byte-comparable with the untouched archive.
*/
const SLIDE_TRANSITION_TYPES = [
	"cut",
	"fade",
	"dissolve",
	"push",
	"wipe",
	"cover",
	"pull"
];
const SLIDE_TRANSITION_DIRECTIONS = [
	"left",
	"right",
	"up",
	"down"
];
const SLIDE_TRANSITION_SPEEDS = [
	"slow",
	"med",
	"fast"
];
const PRESENTATION_NAMESPACE$1 = "http://schemas.openxmlformats.org/presentationml/2006/main";
const SLIDE_PART$1 = /^ppt\/slides\/slide(\d+)\.xml$/u;
const DIRECTIONAL_TYPES = [
	"push",
	"wipe",
	"cover",
	"pull"
];
const DIRECTION_ATTRIBUTES = {
	left: "l",
	right: "r",
	up: "u",
	down: "d"
};
function isSlideTransitionType(value) {
	return SLIDE_TRANSITION_TYPES.includes(value);
}
function normalizeSlideTransition(transition) {
	const type = transition.type;
	if (!isSlideTransitionType(type)) throw new PptError("PPT_CREATE_INPUT_INVALID", `unsupported slide transition type: ${type}`, { details: { supported: [...SLIDE_TRANSITION_TYPES] } });
	if (transition.direction !== void 0) {
		if (!DIRECTIONAL_TYPES.includes(type)) throw new PptError("PPT_CREATE_INPUT_INVALID", `slide transition ${type} does not accept a direction`);
		if (!SLIDE_TRANSITION_DIRECTIONS.includes(transition.direction)) throw new PptError("PPT_CREATE_INPUT_INVALID", `unsupported slide transition direction: ${String(transition.direction)}`);
	}
	if (transition.speed !== void 0 && !SLIDE_TRANSITION_SPEEDS.includes(transition.speed)) throw new PptError("PPT_CREATE_INPUT_INVALID", `unsupported slide transition speed: ${String(transition.speed)}`);
	if (transition.advanceAfterMs !== void 0 && (!Number.isInteger(transition.advanceAfterMs) || transition.advanceAfterMs <= 0)) throw new PptError("PPT_CREATE_INPUT_INVALID", `slide transition advanceAfterMs must be a positive integer: ${String(transition.advanceAfterMs)}`);
	if (transition.advanceOnClick !== void 0 && transition.advanceOnClick !== false) throw new PptError("PPT_CREATE_INPUT_INVALID", "slide transition advanceOnClick may only be set to false");
	const normalized = { type };
	if (transition.direction !== void 0) normalized.direction = transition.direction;
	if (transition.speed !== void 0) normalized.speed = transition.speed;
	if (transition.advanceAfterMs !== void 0) normalized.advanceAfterMs = transition.advanceAfterMs;
	if (transition.advanceOnClick === false) normalized.advanceOnClick = false;
	return normalized;
}
function transitionElementXml(transition) {
	const normalized = normalizeSlideTransition(transition);
	const attributes = [];
	if (normalized.advanceOnClick === false) attributes.push("advClick=\"0\"");
	if (normalized.advanceAfterMs !== void 0) attributes.push(`advTm="${normalized.advanceAfterMs}"`);
	if (normalized.speed !== void 0 && normalized.type !== "cut") attributes.push(`spd="${normalized.speed}"`);
	const direction = DIRECTIONAL_TYPES.includes(normalized.type) && normalized.direction !== void 0 ? ` dir="${DIRECTION_ATTRIBUTES[normalized.direction]}"` : "";
	return `${attributes.length === 0 ? "<p:transition>" : `<p:transition ${attributes.join(" ")}>`}<p:${normalized.type}${direction}/></p:transition>`;
}
function elementEndOffset$1(body, name) {
	const match = new RegExp(`<${name}\\b[^>]*?(/?)>`, "u").exec(body);
	if (match === null) return -1;
	if (match[1] === "/") return match.index + match[0].length;
	const close = body.indexOf(`</${name}>`, match.index + match[0].length);
	if (close < 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `slide XML has an unterminated ${name} element`);
	return close + name.length + 3;
}
/** CT_Slide order is cSld, clrMapOvr, transition, timing, extLst. */
function transitionInsertionOffset(body) {
	const csldEnd = elementEndOffset$1(body, "p:cSld");
	if (csldEnd < 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML has no p:cSld element");
	const clrMapEnd = elementEndOffset$1(body, "p:clrMapOvr");
	return clrMapEnd >= csldEnd ? clrMapEnd : csldEnd;
}
function injectSlideTransition(xml, transition) {
	const element = transitionElementXml(transition);
	const root = /<p:sld\b[^>]*>/u.exec(xml);
	if (root === null) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML does not contain a p:sld root element");
	const bodyStart = root.index + root[0].length;
	const bodyEnd = xml.lastIndexOf("</p:sld>");
	if (bodyEnd < bodyStart) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML has an unterminated p:sld root element");
	const body = xml.slice(bodyStart, bodyEnd);
	if (/<p:transition\b/u.test(body)) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML already declares a p:transition element");
	const payload = /xmlns:p=/u.test(root[0]) ? element : element.replace("<p:transition", `<p:transition xmlns:p="${PRESENTATION_NAMESPACE$1}"`);
	const at = bodyStart + transitionInsertionOffset(body);
	return `${xml.slice(0, at)}${payload}${xml.slice(at)}`;
}
function rewritePptxTransitions(data, plan) {
	if (plan.size === 0) return data;
	const targets = /* @__PURE__ */ new Map();
	for (const [page, transition] of plan) {
		if (!Number.isInteger(page) || page < 1) throw new PptError("PPT_CREATE_INPUT_INVALID", `slide transition page must be a positive integer: ${String(page)}`);
		targets.set(page, transition);
	}
	const injected = /* @__PURE__ */ new Set();
	const output = rewriteZip(data, { transform: (name, read) => {
		const match = SLIDE_PART$1.exec(name);
		const page = match === null ? void 0 : Number(match[1]);
		const transition = page === void 0 ? void 0 : targets.get(page);
		if (page === void 0 || transition === void 0) return void 0;
		injected.add(page);
		return strToU8(injectSlideTransition(strFromU8(read()), transition));
	} });
	for (const page of targets.keys()) if (!injected.has(page)) throw new PptError("PPT_CREATE_INPUT_INVALID", `no slide part matches transition page ${page}`);
	return output;
}
//#endregion
//#region src/animation.ts
/**
* Text and shape animation lives in the slide's `p:timing` timeline, which
* pptxgenjs does not emit at all. A deck opts in by declaring, per page, which
* element plays which entrance effect; the finished package is then rewritten so
* only `ppt/slides/slideN.xml` changes.
*
* The effect catalogue and every literal in the emitted XML below were taken from
* timelines that PowerPoint 16.0 wrote itself (probed through COM), not from the
* prose of the specification, because the specification deliberately leaves the
* `presetID` numbering unspecified:
*
* - `presetClass` partitions the `presetID` numbering space. `presetID="10"` is
*   Fade only under `presetClass="entr"`; the same number means something else
*   under `emph`, `exit`, or `path`. This module therefore always writes both.
* - `presetSubtype` is a direction bitmask, not a per-effect enum:
*   1 = from top, 2 = from right, 4 = from bottom, 8 = from left, 16 = in,
*   32 = out. Derived values are the bitwise or: 10 = 2|8 = horizontal,
*   5 = 1|4 = vertical, 9 = top-left, 3 = top-right, 6 = bottom-right,
*   12 = bottom-left. Confirmed against the defaults PowerPoint emits
*   (Fade 0, Blinds 10, Box 16, Strips 12, Wheel 1, Zoom 16, Split 21).
* - The renderer obeys the effect *nodes*, not the preset triple: `p:animEffect`
*   carries a filter string and `p:anim` carries explicit from/to values, while
*   the preset triple only drives the UI's effect list. That is why a direction
*   change is expressed by rewriting the node, and why `peek` can legitimately
*   pair `presetSubtype="4"` with `filter="wipe(up)"` (PowerPoint emits exactly
*   that: peek's filter runs opposite to its subtype).
* - Every entrance effect PowerPoint emitted carried a `p:set` that forces
*   `style.visibility` to `visible`; without it the shape can stay hidden.
*/
const TEXT_ANIMATION_EFFECTS = [
	"appear",
	"flash-once",
	"fade",
	"dissolve",
	"wedge",
	"wipe",
	"blinds",
	"checkerboard",
	"random-bars",
	"box",
	"circle",
	"diamond",
	"plus",
	"split",
	"strips",
	"wheel",
	"zoom",
	"fly-in",
	"crawl",
	"peek",
	"stretch",
	"swivel",
	"spiral",
	"bounce",
	"credits",
	"float-in",
	"grow-turn",
	"rise-up",
	"unfold"
];
const TEXT_ANIMATION_DIRECTIONS = [
	"left",
	"right",
	"up",
	"down"
];
const TEXT_ANIMATION_STARTS = [
	"on-click",
	"with-previous",
	"after-previous"
];
const PRESENTATION_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";
const SLIDE_PART = /^ppt\/slides\/slide(\d+)\.xml$/u;
const SHAPE_ID = /<p:cNvPr id="(\d+)" name="([^"]*)"/gu;
/** Bit values of the `presetSubtype` direction mask. */
const SUBTYPE_TOP = 1;
const SUBTYPE_RIGHT = 2;
const SUBTYPE_BOTTOM = 4;
const SUBTYPE_LEFT = 8;
const SUBTYPE_IN = 16;
const DIRECTION_BITS = {
	up: SUBTYPE_TOP,
	right: SUBTYPE_RIGHT,
	down: SUBTYPE_BOTTOM,
	left: SUBTYPE_LEFT
};
const directionBit = (direction) => DIRECTION_BITS[direction];
/** Effects whose direction only collapses to a horizontal or vertical axis. */
const axesSubtype = (horizontal, vertical) => (direction) => direction === "up" || direction === "down" ? vertical : horizontal;
const HORIZONTAL = 10;
const VERTICAL = 5;
/**
* A motion that slides the shape in from an off-slide edge. The from-bottom
* pair is PowerPoint's own output; the other edges are its mirror images about
* the slide centre, using the same `1 + size/2` offset convention.
*/
function edgeMotion(direction) {
	if (direction === "down") return {
		attr: "ppt_y",
		from: "1+#ppt_h/2",
		to: "#ppt_y",
		additive: true
	};
	if (direction === "up") return {
		attr: "ppt_y",
		from: "-#ppt_h/2",
		to: "#ppt_y",
		additive: true
	};
	if (direction === "right") return {
		attr: "ppt_x",
		from: "1+#ppt_w/2",
		to: "#ppt_x",
		additive: true
	};
	return {
		attr: "ppt_x",
		from: "-#ppt_w/2",
		to: "#ppt_x",
		additive: true
	};
}
function horizontalVertical(horizontal, vertical) {
	return {
		left: horizontal,
		right: horizontal,
		up: vertical,
		down: vertical
	};
}
const EFFECT_DEFINITIONS = {
	appear: {
		presetID: 1,
		subtype: 0,
		nodes: { kind: "visibility" }
	},
	"flash-once": {
		presetID: 11,
		subtype: 0,
		nodes: { kind: "visibility" }
	},
	fade: {
		presetID: 10,
		subtype: 0,
		nodes: {
			kind: "filter",
			filter: "fade"
		}
	},
	dissolve: {
		presetID: 9,
		subtype: 0,
		nodes: {
			kind: "filter",
			filter: "dissolve"
		}
	},
	wedge: {
		presetID: 20,
		subtype: 0,
		nodes: {
			kind: "filter",
			filter: "wedge"
		}
	},
	wipe: {
		presetID: 22,
		subtype: SUBTYPE_BOTTOM,
		subtypeFor: directionBit,
		nodes: {
			kind: "filter",
			filter: "wipe(down)",
			directional: {
				up: "wipe(up)",
				down: "wipe(down)",
				left: "wipe(left)",
				right: "wipe(right)"
			}
		}
	},
	blinds: {
		presetID: 3,
		subtype: HORIZONTAL,
		subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
		nodes: {
			kind: "filter",
			filter: "blinds(horizontal)",
			directional: horizontalVertical("blinds(horizontal)", "blinds(vertical)")
		}
	},
	checkerboard: {
		presetID: 5,
		subtype: HORIZONTAL,
		subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
		nodes: {
			kind: "filter",
			filter: "checkerboard(across)",
			directional: horizontalVertical("checkerboard(across)", "checkerboard(down)")
		}
	},
	"random-bars": {
		presetID: 14,
		subtype: HORIZONTAL,
		subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
		nodes: {
			kind: "filter",
			filter: "randombar(horizontal)",
			directional: horizontalVertical("randombar(horizontal)", "randombar(vertical)")
		}
	},
	box: {
		presetID: 4,
		subtype: SUBTYPE_IN,
		subtypeFor: directionBit,
		nodes: {
			kind: "filter",
			filter: "box(in)"
		}
	},
	circle: {
		presetID: 6,
		subtype: SUBTYPE_IN,
		subtypeFor: directionBit,
		nodes: {
			kind: "filter",
			filter: "circle(in)"
		}
	},
	diamond: {
		presetID: 8,
		subtype: SUBTYPE_IN,
		subtypeFor: directionBit,
		nodes: {
			kind: "filter",
			filter: "diamond(in)"
		}
	},
	plus: {
		presetID: 13,
		subtype: SUBTYPE_IN,
		subtypeFor: directionBit,
		nodes: {
			kind: "filter",
			filter: "plus(in)"
		}
	},
	split: {
		presetID: 16,
		subtype: 21,
		nodes: {
			kind: "filter",
			filter: "barn(inVertical)"
		}
	},
	strips: {
		presetID: 18,
		subtype: 12,
		nodes: {
			kind: "filter",
			filter: "strips(downLeft)"
		}
	},
	wheel: {
		presetID: 21,
		subtype: 1,
		nodes: {
			kind: "filter",
			filter: "wheel(1)"
		}
	},
	zoom: {
		presetID: 23,
		subtype: SUBTYPE_IN,
		nodes: {
			kind: "motion",
			motions: () => [{
				attr: "ppt_w",
				from: 0,
				to: "#ppt_w"
			}, {
				attr: "ppt_h",
				from: 0,
				to: "#ppt_h"
			}]
		}
	},
	"fly-in": {
		presetID: 2,
		subtype: SUBTYPE_BOTTOM,
		subtypeFor: directionBit,
		nodes: {
			kind: "motion",
			motions: (direction) => [edgeMotion(direction ?? "down")]
		}
	},
	crawl: {
		presetID: 7,
		subtype: SUBTYPE_BOTTOM,
		subtypeFor: directionBit,
		nodes: {
			kind: "motion",
			motions: (direction) => [edgeMotion(direction ?? "down")]
		}
	},
	peek: {
		presetID: 12,
		subtype: SUBTYPE_BOTTOM,
		subtypeFor: directionBit,
		nodes: {
			kind: "filter",
			filter: "wipe(up)",
			directional: {
				up: "wipe(down)",
				down: "wipe(up)",
				left: "wipe(right)",
				right: "wipe(left)"
			}
		}
	},
	stretch: {
		presetID: 17,
		subtype: HORIZONTAL,
		subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
		nodes: {
			kind: "motion",
			motions: (direction) => direction === "up" || direction === "down" ? [{
				attr: "ppt_h",
				from: 0,
				to: "#ppt_h"
			}] : [{
				attr: "ppt_w",
				from: 0,
				to: "#ppt_w"
			}]
		}
	},
	swivel: {
		presetID: 19,
		subtype: HORIZONTAL,
		subtypeFor: axesSubtype(HORIZONTAL, VERTICAL),
		nodes: {
			kind: "motion",
			motions: (direction) => direction === "up" || direction === "down" ? [{
				attr: "ppt_h",
				from: 0,
				to: "#ppt_h"
			}] : [{
				attr: "ppt_w",
				from: 0,
				to: "#ppt_w"
			}]
		}
	},
	spiral: {
		presetID: 15,
		subtype: 0,
		nodes: {
			kind: "motion",
			motions: () => [
				{
					attr: "ppt_w",
					from: 0,
					to: "#ppt_w"
				},
				{
					attr: "ppt_h",
					from: 0,
					to: "#ppt_h"
				},
				{
					attr: "ppt_x",
					from: "#ppt_x",
					to: "#ppt_x"
				},
				{
					attr: "ppt_y",
					from: "#ppt_y",
					to: "#ppt_y"
				}
			]
		}
	},
	bounce: {
		presetID: 26,
		subtype: 0,
		nodes: {
			kind: "motion",
			motions: () => [{
				attr: "ppt_y",
				from: "#ppt_y-0.25",
				to: "#ppt_y",
				additive: true
			}]
		}
	},
	credits: {
		presetID: 28,
		subtype: 0,
		nodes: {
			kind: "motion",
			motions: () => [{
				attr: "ppt_y",
				from: "#ppt_y+1",
				to: "#ppt_y-1",
				additive: true
			}]
		}
	},
	"float-in": {
		presetID: 30,
		subtype: 0,
		nodes: {
			kind: "motion",
			motions: () => [
				{
					attr: "style.rotation",
					from: -90,
					to: 0
				},
				{
					attr: "ppt_x",
					from: "#ppt_x+0.4",
					to: "#ppt_x-0.05",
					additive: true
				},
				{
					attr: "ppt_y",
					from: "#ppt_y-0.4",
					to: "#ppt_y+0.1",
					additive: true
				}
			]
		}
	},
	"grow-turn": {
		presetID: 31,
		subtype: 0,
		nodes: {
			kind: "motion",
			motions: () => [
				{
					attr: "ppt_w",
					from: 0,
					to: "#ppt_w"
				},
				{
					attr: "ppt_h",
					from: 0,
					to: "#ppt_h"
				},
				{
					attr: "style.rotation",
					from: 90,
					to: 0
				}
			]
		}
	},
	"rise-up": {
		presetID: 37,
		subtype: 0,
		nodes: {
			kind: "motion",
			motions: () => [{
				attr: "ppt_y",
				from: "#ppt_y+1",
				to: "#ppt_y-.03",
				additive: true
			}]
		}
	},
	unfold: {
		presetID: 40,
		subtype: 0,
		nodes: {
			kind: "motion",
			motions: () => [{
				attr: "ppt_x",
				from: "#ppt_x-.1",
				to: "#ppt_x",
				additive: true
			}]
		}
	}
};
function isTextAnimationEffect(value) {
	return TEXT_ANIMATION_EFFECTS.includes(value);
}
function isTextAnimationDirection(value) {
	return TEXT_ANIMATION_DIRECTIONS.includes(value);
}
function fail(message) {
	throw new PptError("PPT_CREATE_INPUT_INVALID", message);
}
function normalizeTextAnimation(animation) {
	if (typeof animation.target !== "string" || animation.target.trim() === "") fail("text animation target must be a non-empty element id");
	if (!isTextAnimationEffect(animation.effect)) fail(`unsupported text animation effect: ${String(animation.effect)}`);
	const definition = EFFECT_DEFINITIONS[animation.effect];
	if (animation.direction !== void 0) {
		if (!isTextAnimationDirection(animation.direction)) fail(`unsupported text animation direction: ${String(animation.direction)}`);
		if (definition.subtypeFor === void 0) fail(`text animation effect ${animation.effect} does not accept a direction`);
	}
	if (animation.start !== void 0 && !TEXT_ANIMATION_STARTS.includes(animation.start)) fail(`unsupported text animation start: ${String(animation.start)}`);
	if (animation.durationMs !== void 0 && (!Number.isInteger(animation.durationMs) || animation.durationMs < 1 || animation.durationMs > 6e4)) fail(`text animation durationMs must be an integer between 1 and 60000: ${String(animation.durationMs)}`);
	const normalized = {
		target: animation.target.trim(),
		effect: animation.effect
	};
	if (animation.direction !== void 0) normalized.direction = animation.direction;
	if (animation.byParagraph === true) normalized.byParagraph = true;
	if (animation.start !== void 0) normalized.start = animation.start;
	if (animation.durationMs !== void 0) normalized.durationMs = animation.durationMs;
	return normalized;
}
function planSlideAnimations(entries) {
	const plan = /* @__PURE__ */ new Map();
	for (const entry of entries) {
		if (!Number.isInteger(entry.page) || entry.page < 1) fail(`text animation page must be a positive integer: ${String(entry.page)}`);
		if (entry.animations.length > 24) fail(`page ${entry.page} declares ${entry.animations.length} text animations; the limit is 24`);
		const seen = /* @__PURE__ */ new Set();
		const animations = entry.animations.map((animation) => {
			const normalized = normalizeTextAnimation(animation);
			const key = `${normalized.target}\u0000${normalized.byParagraph === true ? "paragraph" : "shape"}`;
			if (seen.has(key)) fail(`page ${entry.page} declares two animations for ${normalized.target}; a target may be animated once per build unit`);
			seen.add(key);
			return normalized;
		});
		plan.set(entry.page, animations);
	}
	return plan;
}
/** `p:spTgt/@spid` addresses `p:cNvPr/@id`, so the id has to come from the slide part. */
function indexShapes(xml) {
	const idByName = /* @__PURE__ */ new Map();
	for (const match of xml.matchAll(SHAPE_ID)) idByName.set(match[2], match[1]);
	const paragraphCountByName = /* @__PURE__ */ new Map();
	for (const [name, id] of idByName) {
		const start = xml.indexOf(`<p:cNvPr id="${id}"`);
		if (start < 0) continue;
		const end = xml.indexOf("</p:sp>", start);
		const block = xml.slice(start, end < 0 ? void 0 : end);
		const paragraphs = [...block.matchAll(/<a:p[ >]/gu)].length;
		const paragraphProperties = [...block.matchAll(/<a:pPr[ >]/gu)].length;
		paragraphCountByName.set(name, Math.max(paragraphs, paragraphProperties));
	}
	return {
		idByName,
		paragraphCountByName
	};
}
function escapeAttribute(value) {
	return value.replace(/&/gu, "&amp;").replace(/</gu, "&lt;").replace(/>/gu, "&gt;").replace(/"/gu, "&quot;");
}
var TimelineBuilder = class {
	nextId = 1;
	id() {
		const value = this.nextId;
		this.nextId += 1;
		return value;
	}
};
function targetXml(spid, paragraph) {
	if (paragraph === void 0) return `<p:spTgt spid="${spid}"/>`;
	return `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="${paragraph}" end="${paragraph}"/></p:txEl></p:spTgt>`;
}
function visibilitySet(builder, target) {
	return `<p:set><p:cBhvr><p:cTn id="${builder.id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>${target}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
}
function animEffectNode(builder, target, filter, durationMs) {
	return `<p:animEffect transition="in" filter="${escapeAttribute(filter)}"><p:cBhvr><p:cTn id="${builder.id()}" dur="${durationMs}"/><p:tgtEl>${target}</p:tgtEl></p:cBhvr></p:animEffect>`;
}
/**
* A numeric from/to is written as `p:fltVal` and a formula as `p:strVal`, because
* PowerPoint rejects a formula it cannot parse but happily reads a literal — this
* is exactly how it writes Zoom's zero scale itself.
*/
function animNode(builder, target, motion, durationMs) {
	const value = (input) => typeof input === "number" ? `<p:fltVal val="${input}"/>` : `<p:strVal val="${escapeAttribute(input)}"/>`;
	return `<p:anim calcmode="lin" valueType="num">${motion.additive === true ? "<p:cBhvr additive=\"base\">" : "<p:cBhvr>"}<p:cTn id="${builder.id()}" dur="${durationMs}" fill="hold"/><p:tgtEl>${target}</p:tgtEl><p:attrNameLst><p:attrName>${motion.attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val>${value(motion.from)}</p:val></p:tav><p:tav tm="100000"><p:val>${value(motion.to)}</p:val></p:tav></p:tavLst></p:anim>`;
}
function effectXml(builder, animation, shape, start) {
	const spid = shape.idByName.get(animation.target);
	if (spid === void 0) throw new PptError("PPT_CREATE_INPUT_INVALID", `text animation target is not present on the slide: ${animation.target}`);
	const definition = EFFECT_DEFINITIONS[animation.effect];
	const durationMs = animation.durationMs ?? 500;
	const paragraphs = animation.byParagraph === true ? shape.paragraphCountByName.get(animation.target) ?? 0 : 0;
	if (animation.byParagraph === true && paragraphs < 1) throw new PptError("PPT_CREATE_INPUT_INVALID", `byParagraph animation target has no text paragraphs: ${animation.target}`);
	const units = animation.byParagraph === true ? Array.from({ length: paragraphs }, (_, index) => index) : [void 0];
	const subtype = definition.subtypeFor !== void 0 && animation.direction !== void 0 ? definition.subtypeFor(animation.direction) : definition.subtype;
	const buildOne = (paragraph) => {
		const effectId = builder.id();
		const target = targetXml(spid, paragraph);
		const parts = [visibilitySet(builder, target)];
		if (definition.nodes.kind === "visibility") {} else if (definition.nodes.kind === "filter") {
			const filter = animation.direction === void 0 ? definition.nodes.filter : definition.nodes.directional?.[animation.direction] ?? definition.nodes.filter;
			parts.push(animEffectNode(builder, target, filter, durationMs));
		} else for (const motion of definition.nodes.motions(animation.direction)) parts.push(animNode(builder, target, motion, durationMs));
		return `<p:par><p:cTn id="${effectId}" presetID="${definition.presetID}" presetClass="entr" presetSubtype="${subtype}" fill="hold" grpId="0" nodeType="__NODE__"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${parts.join("")}</p:childTnLst></p:cTn></p:par>`;
	};
	const nodeType = start === "with-previous" ? "withEffect" : start === "after-previous" ? "afterEffect" : "clickEffect";
	if (units.length === 1) return buildOne(units[0]).replace("__NODE__", nodeType);
	return units.map((paragraph) => buildOne(paragraph).replace("__NODE__", "clickEffect")).join("");
}
/** Every `on-click` opens a new timeline group; the first animation always opens one. */
function planAnimationGroups(animations) {
	const groups = [];
	for (const animation of animations) {
		const requested = animation.start ?? "on-click";
		if (groups.length === 0 || requested === "on-click") {
			groups.push([{
				animation,
				start: "on-click"
			}]);
			continue;
		}
		groups[groups.length - 1].push({
			animation,
			start: requested
		});
	}
	return groups;
}
function animationTimingXml(animations, shape) {
	if (animations.length === 0) throw new PptError("PPT_CREATE_INPUT_INVALID", "a slide animation timeline needs at least one animation");
	const builder = new TimelineBuilder();
	const rootId = builder.id();
	const seqId = builder.id();
	const groups = [];
	for (const group of planAnimationGroups(animations)) {
		const groupId = builder.id();
		const innerId = builder.id();
		let effects = "";
		for (const planned of group) effects += effectXml(builder, planned.animation, shape, planned.start);
		groups.push(`<p:par><p:cTn id="${groupId}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${innerId}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`);
	}
	return `<p:timing><p:tnLst><p:par><p:cTn id="${rootId}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="${seqId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groups.join("")}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}
function elementEndOffset(body, name) {
	const match = new RegExp(`<${name}\\b[^>]*?(/?)>`, "u").exec(body);
	if (match === null) return -1;
	if (match[1] === "/") return match.index + match[0].length;
	const close = body.indexOf(`</${name}>`, match.index + match[0].length);
	if (close < 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `slide XML has an unterminated ${name} element`);
	return close + name.length + 3;
}
/** CT_Slide order is cSld, clrMapOvr, transition, timing, extLst. */
function timingInsertionOffset(body) {
	const csldEnd = elementEndOffset(body, "p:cSld");
	if (csldEnd < 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML has no p:cSld element");
	const clrMapEnd = elementEndOffset(body, "p:clrMapOvr");
	const transitionEnd = elementEndOffset(body, "p:transition");
	return Math.max(csldEnd, clrMapEnd, transitionEnd);
}
function injectSlideAnimation(xml, animations) {
	const root = /<p:sld\b[^>]*>/u.exec(xml);
	if (root === null) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML does not contain a p:sld root element");
	const bodyStart = root.index + root[0].length;
	const bodyEnd = xml.lastIndexOf("</p:sld>");
	if (bodyEnd < bodyStart) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML has an unterminated p:sld root element");
	const withoutTiming = xml.slice(bodyStart, bodyEnd).replace(/<p:timing\b[\s\S]*?<\/p:timing>/u, "");
	if (withoutTiming.includes("<p:timing")) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "slide XML has an unterminated p:timing element");
	const payload = animationTimingXml(animations, indexShapes(xml));
	const namespaced = /xmlns:p=/u.test(root[0]) ? payload : payload.replace("<p:timing", `<p:timing xmlns:p="${PRESENTATION_NAMESPACE}"`);
	const at = bodyStart + timingInsertionOffset(withoutTiming);
	const rebuilt = `${withoutTiming.slice(0, at)}${namespaced}${withoutTiming.slice(at)}`;
	return `${xml.slice(0, bodyStart)}${rebuilt}${xml.slice(bodyEnd)}`;
}
function rewritePptxAnimations(data, plan) {
	if (plan.size === 0) return data;
	const targets = /* @__PURE__ */ new Map();
	for (const [page, animations] of plan) {
		if (!Number.isInteger(page) || page < 1) fail(`text animation page must be a positive integer: ${String(page)}`);
		targets.set(page, animations);
	}
	const injected = /* @__PURE__ */ new Set();
	const output = rewriteZip(data, { transform: (name, read) => {
		const match = SLIDE_PART.exec(name);
		const page = match === null ? void 0 : Number(match[1]);
		const animations = page === void 0 ? void 0 : targets.get(page);
		if (page === void 0 || animations === void 0) return void 0;
		injected.add(page);
		return strToU8(injectSlideAnimation(strFromU8(read()), animations));
	} });
	for (const page of targets.keys()) if (!injected.has(page)) fail(`no slide part matches text animation page ${page}`);
	return output;
}
//#endregion
//#region src/pptx.ts
function color(value) {
	if (value === "transparent") return void 0;
	const match = /rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)(?:\s*[,/]\s*(\d+(?:\.\d+)?))?\s*\)/iu.exec(value);
	if (match === null) return void 0;
	const hex = [
		match[1],
		match[2],
		match[3]
	].map((channel) => Math.max(0, Math.min(255, Math.round(Number(channel)))).toString(16).padStart(2, "0")).join("").toUpperCase();
	const alpha = match[4] === void 0 ? 1 : Math.max(0, Math.min(1, Number(match[4])));
	return {
		hex,
		transparency: Math.round((1 - alpha) * 100)
	};
}
function baseOptions(element) {
	return {
		x: pxToInches(element.box.x),
		y: pxToInches(element.box.y),
		w: pxToInches(element.box.w),
		h: pxToInches(element.box.h),
		objectName: element.id
	};
}
function textOptions(style) {
	const foreground = color(style.color);
	const fill = color(style.backgroundColor);
	const line = color(style.borderColor);
	const align = [
		"left",
		"center",
		"right",
		"justify"
	].includes(style.textAlign) ? style.textAlign : "left";
	return {
		fontFace: style.fontFamily,
		fontSize: pxToPoints(style.fontSizePx),
		bold: style.fontWeight >= 600,
		italic: style.fontStyle === "italic",
		color: foreground?.hex ?? "000000",
		margin: 0,
		breakLine: false,
		fit: "none",
		align,
		valign: style.verticalAlign === "bottom" ? "bottom" : style.verticalAlign === "middle" ? "mid" : "top",
		lineSpacing: pxToPoints(style.lineHeightPx),
		...fill === void 0 ? {} : { fill: {
			color: fill.hex,
			transparency: fill.transparency
		} },
		...style.borderStyle === "none" || style.borderWidthPx <= 0 || line === void 0 ? { line: { type: "none" } } : { line: {
			color: line.hex,
			transparency: line.transparency,
			width: pxToPoints(style.borderWidthPx)
		} }
	};
}
function imageSource(pathOrUrl, workspace, artifactRoot) {
	let path;
	if (isLocalFilesystemPath(pathOrUrl)) path = pathOrUrl;
	else {
		let parsed;
		try {
			parsed = new URL(pathOrUrl);
		} catch {
			parsed = void 0;
		}
		if (parsed === void 0) path = pathOrUrl;
		else if (parsed.protocol === "file:") path = fileURLToPath(parsed);
		else throw new PptError("PPT_CREATE_ASSET_MISSING", "PPTX images must be local frozen files");
	}
	if (!isPathInside(workspace, path) || !isPathInside(artifactRoot, path)) throw new PptError("PPT_CREATE_ASSET_MISSING", `image is outside the artifact directory: ${path}`);
	return path;
}
function addNativeElement(pptx, slide, element, workspace, artifactRoot) {
	const base = baseOptions(element);
	if (element.kind === "text") {
		const runs = (element.runs ?? []).map((run) => ({
			text: run.text,
			options: {
				fontFace: run.fontFamily,
				fontSize: pxToPoints(run.fontSizePx),
				bold: run.fontWeight >= 600,
				italic: run.fontStyle === "italic",
				color: color(run.color)?.hex ?? "000000",
				...run.textDecoration.includes("underline") ? { underline: { style: "sng" } } : {}
			}
		}));
		slide.addText(runs.length > 0 ? runs : element.text ?? "", {
			...base,
			...textOptions(element.style)
		});
		return;
	}
	if (element.kind === "image") {
		if (element.imagePath === void 0) throw new PptError("PPT_CREATE_ASSET_MISSING", `image path missing for ${element.id}`);
		const path = imageSource(element.imagePath, workspace, artifactRoot);
		const fit = element.style.objectFit === "contain" ? "contain" : "cover";
		slide.addImage({
			path,
			...base,
			sizing: {
				type: fit,
				w: base.w,
				h: base.h
			},
			transparency: Math.round((1 - element.style.opacity) * 100)
		});
		return;
	}
	if (element.kind === "svg") {
		if (element.svg === void 0) throw new PptError("PPT_CREATE_UNSUPPORTED_ELEMENT", `SVG markup missing for ${element.id}`);
		const data = `data:image/svg+xml;base64,${Buffer.from(element.svg).toString("base64")}`;
		slide.addImage({
			data,
			...base,
			sizing: {
				type: "contain",
				w: base.w,
				h: base.h
			}
		});
		return;
	}
	if (element.kind === "table") {
		slide.addTable((element.table ?? []).map((row) => row.map((cell) => ({ text: cell }))), {
			...base,
			border: {
				type: "solid",
				color: color(element.style.borderColor)?.hex ?? "999999",
				pt: pxToPoints(Math.max(1, element.style.borderWidthPx))
			},
			color: color(element.style.color)?.hex ?? "000000",
			fill: color(element.style.backgroundColor)?.hex ?? "FFFFFF",
			fontFace: element.style.fontFamily,
			fontSize: pxToPoints(element.style.fontSizePx),
			margin: 0
		});
		return;
	}
	const fill = color(element.style.backgroundColor);
	const line = color(element.style.borderColor);
	const radius = Number.parseFloat(element.style.borderRadius);
	const shape = element.box.w <= 2 || element.box.h <= 2 ? pptx.ShapeType.line : /%/u.test(element.style.borderRadius) && radius >= 50 ? pptx.ShapeType.ellipse : radius > 0 ? pptx.ShapeType.roundRect : pptx.ShapeType.rect;
	slide.addShape(shape, {
		...base,
		fill: fill === void 0 ? { type: "none" } : {
			color: fill.hex,
			transparency: fill.transparency
		},
		line: element.style.borderStyle === "none" || element.style.borderWidthPx <= 0 || line === void 0 ? { type: "none" } : {
			color: line.hex,
			transparency: line.transparency,
			width: pxToPoints(element.style.borderWidthPx)
		}
	});
}
function relationshipSource(path) {
	if (path === "_rels/.rels") return "";
	const index = path.indexOf("/_rels/");
	if (index < 0 || !path.endsWith(".rels")) return "";
	return `${path.slice(0, index)}/${path.slice(index + 7, -5)}`;
}
function inspectPptxPackage(data, expectedPages) {
	let files;
	try {
		files = unzipSync(data);
	} catch (error) {
		throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX is not a readable ZIP package", { cause: error });
	}
	const names = Object.keys(files).sort();
	for (const required of [
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels"
	]) if (files[required] === void 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `PPTX entry is missing: ${required}`);
	const slideNames = names.filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).sort((a, b) => Number(a.match(/\d+/u)[0]) - Number(b.match(/\d+/u)[0]));
	if (slideNames.length !== expectedPages) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `expected ${expectedPages} slides, found ${slideNames.length}`);
	slideNames.forEach((name, index) => {
		if (name !== `ppt/slides/slide${index + 1}.xml`) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `slide sequence is not contiguous at ${name}`);
	});
	const presentation = strFromU8(files["ppt/presentation.xml"]);
	const size = /<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/u.exec(presentation);
	if (size === null) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "presentation slide size is missing");
	const widthEmu = Number(size[1]);
	const heightEmu = Number(size[2]);
	if (widthEmu !== 12192e3 || heightEmu !== 6858e3) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `unexpected slide size: ${widthEmu}x${heightEmu}`);
	const presentationRels = strFromU8(files["ppt/_rels/presentation.xml.rels"]);
	const relationTargets = /* @__PURE__ */ new Map();
	for (const match of presentationRels.matchAll(/<Relationship\b([^>]*)>/giu)) {
		const id = /\bId="([^"]+)"/iu.exec(match[1])?.[1];
		const target = /\bTarget="([^"]+)"/iu.exec(match[1])?.[1];
		if (id !== void 0 && target !== void 0) relationTargets.set(id, posix.normalize(posix.join("ppt", target)));
	}
	const orderedSlideIds = [...presentation.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"/giu)].map((match) => match[1]);
	if (orderedSlideIds.length !== expectedPages) throw new PptError("PPT_CREATE_INVALID_PACKAGE", "presentation slide order list does not match page count");
	orderedSlideIds.forEach((id, index) => {
		if (relationTargets.get(id) !== `ppt/slides/slide${index + 1}.xml`) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `presentation slide order is invalid at page ${index + 1}`);
	});
	for (const name of names.filter((entry) => entry.endsWith(".rels"))) {
		const xml = strFromU8(files[name]);
		if (/TargetMode="External"/iu.test(xml)) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `external relationship is forbidden: ${name}`);
		const source = relationshipSource(name);
		const base = source === "" ? "" : posix.dirname(source);
		for (const match of xml.matchAll(/<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*>/giu)) {
			const target = match[1];
			if (/^[a-z]+:/iu.test(target)) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `non-package relationship target: ${target}`);
			const resolved = target.startsWith("/") ? target.slice(1) : posix.normalize(posix.join(base, target));
			if (files[resolved] === void 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `relationship target is missing: ${resolved}`);
		}
	}
	for (const name of names.filter((entry) => entry.startsWith("ppt/media/") && !entry.endsWith("/"))) if (files[name].byteLength === 0) throw new PptError("PPT_CREATE_INVALID_PACKAGE", `empty media part: ${name}`);
	return {
		pageCount: slideNames.length,
		widthEmu,
		heightEmu,
		entries: names
	};
}
async function createPptx(browser, owner, workspace, htmlPathInput, outlinePathInput, outputPathInput, fallbackMode = "reject", signal, transitions, animations) {
	throwIfAborted(signal, "PPT_CREATE_ABORTED");
	const [htmlPath, outlinePath, outputPath] = await Promise.all([
		resolveWorkspacePath(workspace, htmlPathInput, {
			mustExist: true,
			kind: "file"
		}),
		resolveWorkspacePath(workspace, outlinePathInput, {
			mustExist: true,
			kind: "file"
		}),
		resolveWorkspacePath(workspace, outputPathInput)
	]);
	const artifactRoot = dirname(outlinePath);
	if (dirname(htmlPath) !== artifactRoot || dirname(outputPath) !== artifactRoot) throw new PptError("PPT_CREATE_INPUT_INVALID", "HTML, outline, and PPTX output must share one artifact directory");
	const outline = validatePptOutline(JSON.parse(await readFile(outlinePath, "utf8")));
	try {
		await validateDeckHtmlSource(workspace, artifactRoot, await readFile(htmlPath, "utf8"), outline.length);
	} catch (error) {
		const issues = error instanceof PptError && Array.isArray(error.details?.issues) ? error.details.issues.map(String) : [];
		if (issues.some((issue) => /missing or invalid local asset/iu.test(issue))) throw new PptError("PPT_CREATE_ASSET_MISSING", "HTML references a missing or invalid local asset", {
			cause: error,
			details: { issues }
		});
		throw error;
	}
	let ir;
	try {
		ir = await browser.extractDeckIr(owner, workspace, htmlPath, outline.length, signal);
	} catch (error) {
		if (signal?.aborted) throw new PptError("PPT_CREATE_ABORTED", "PPTX creation was cancelled", { cause: error });
		throw error;
	}
	outline.forEach((slide, index) => {
		ir.slides[index].speakerNotes = slide.content.flatMap((item) => item.kind === "note" && item.purpose === "speaker" ? [item.text] : []);
	});
	const rasterized = [];
	for (const slide of ir.slides) for (const element of slide.elements) {
		if (element.unsupportedReason === void 0) continue;
		if (fallbackMode !== "rasterize-element") throw new PptError("PPT_CREATE_UNSUPPORTED_ELEMENT", `page ${slide.page} element ${element.id}: ${element.unsupportedReason}`);
		const target = join(artifactRoot, "assets", "rasterized", `page-${slide.page}-${element.id}.png`);
		const imagePath = await browser.rasterizeElement(owner, workspace, htmlPath, element.id, target, signal);
		rasterized.push({
			page: slide.page,
			element_id: element.id,
			reason: element.unsupportedReason,
			image_path: imagePath
		});
		element.kind = "image";
		element.imagePath = join(workspace, imagePath);
		delete element.unsupportedReason;
	}
	const pptx = new PptxGenJS();
	pptx.defineLayout({
		name: "DSH_PPT_16_9",
		width: SLIDE_WIDTH_IN,
		height: SLIDE_HEIGHT_IN
	});
	pptx.layout = "DSH_PPT_16_9";
	pptx.author = "DSH PPT";
	pptx.company = "DSH";
	pptx.subject = "Editable PPTX generated from constrained HTML";
	pptx.title = outline[0]?.title ?? "Presentation";
	let nativeElementCount = 0;
	for (const slideIr of ir.slides) {
		const slide = pptx.addSlide();
		for (const element of slideIr.elements) {
			addNativeElement(pptx, slide, element, workspace, artifactRoot);
			if (!rasterized.some((item) => item.page === slideIr.page && item.element_id === element.id)) nativeElementCount += 1;
		}
		if (slideIr.speakerNotes.length > 0) slide.addNotes(slideIr.speakerNotes.join("\n\n"));
	}
	const temporary = join(artifactRoot, `.deck.${process.pid}.${Date.now()}.pptx`);
	try {
		await pptx.writeFile({
			fileName: temporary,
			compression: true
		});
		throwIfAborted(signal, "PPT_CREATE_ABORTED");
		let bytes = new Uint8Array(await readFile(temporary));
		if (transitions !== void 0) bytes = rewritePptxTransitions(bytes, transitions);
		if (animations !== void 0) bytes = rewritePptxAnimations(bytes, animations);
		inspectPptxPackage(bytes, outline.length);
		await atomicWriteFile(outputPath, bytes, { signal });
		return {
			pptx_path: workspaceRelative(workspace, outputPath),
			page_count: outline.length,
			native_element_count: nativeElementCount,
			rasterized_elements: rasterized,
			structural_status: "passed"
		};
	} catch (error) {
		if (signal?.aborted) throw new PptError("PPT_CREATE_ABORTED", "PPTX creation was cancelled", { cause: error });
		throw error;
	} finally {
		await rm(temporary, { force: true });
	}
}
//#endregion
//#region src/ppt-image.ts
const KEYNOTE_SCRIPT = `on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set inputFile to POSIX file inputPath as alias
  set outputFile to POSIX file outputPath
  tell application "Keynote"
    set sourceDocument to open inputFile
    export sourceDocument to outputFile as slide images with properties {image format:PNG}
    close sourceDocument saving no
  end tell
end run
`;
const POWERPOINT_SCRIPT = `param(
  [Parameter(Mandatory=$true)][string]$InputPptx,
  [Parameter(Mandatory=$true)][string]$OutputDir
)
$ErrorActionPreference = "Stop"
$application = $null
$presentation = $null
try {
  $application = New-Object -ComObject PowerPoint.Application
  $presentation = $application.Presentations.Open($InputPptx, $true, $true, $false)
  $presentation.Export($OutputDir, "PNG", 1280, 720)
  Write-Output ("PowerPoint " + $application.Version)
} finally {
  if ($presentation -ne $null) { $presentation.Close() }
  if ($application -ne $null) { $application.Quit() }
  if ($presentation -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
  if ($application -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application) }
}
`;
const POWERPOINT_MAC_SCREEN_SCRIPT = `on run argv
  set inputPath to item 1 of argv
  set outputDirectory to item 2 of argv
  set pageCount to (item 3 of argv) as integer
  set captureBinary to item 4 of argv
  set screenIndex to (item 5 of argv) as integer
  set inputFile to POSIX file inputPath as alias
  set sourcePresentation to missing value
  set showView to missing value
  tell application "Microsoft PowerPoint"
    try
      activate
      open inputFile
      set sourcePresentation to active presentation
      set settings to slide show settings of sourcePresentation
      try
        set advance mode of settings to slide show advance manual advance
      end try
      try
        set range type of settings to slide show range show all
      end try
      try
        set loop until stopped of settings to false
      end try
      try
        set show with presenter of settings to false
      end try
      set showWindow to run slide show settings
      set showView to slideshow view of showWindow
      delay 1
      repeat with pageNumber from 1 to pageCount
        if pageNumber > 1 then
          go to next slide showView
          delay 1
        end if
        set targetPath to outputDirectory & "/page-" & my zeroPad(pageNumber) & ".png"
        do shell script quoted form of captureBinary & " -x -D " & screenIndex & " " & quoted form of targetPath
      end repeat
      exit slide show showView
      close sourcePresentation saving no
    on error errorMessage number errorNumber
      try
        if showView is not missing value then exit slide show showView
      end try
      try
        if sourcePresentation is not missing value then close sourcePresentation saving no
      end try
      error errorMessage number errorNumber
    end try
  end tell
end run

on zeroPad(pageNumber)
  if pageNumber < 10 then return "00" & pageNumber
  if pageNumber < 100 then return "0" & pageNumber
  return pageNumber as text
end zeroPad
`;
function xmlEscape(value) {
	return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll("\"", "&quot;").replaceAll("'", "&apos;");
}
function fontconfigDocument(fontDirs, cacheDir) {
	return `<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n${[...new Set(fontDirs.map((path) => path.replaceAll("\\", "/")))].map((path) => `  <dir>${xmlEscape(path)}</dir>`).join("\n")}\n  <cachedir>${xmlEscape(cacheDir.replaceAll("\\", "/"))}</cachedir>\n  <config><rescan><int>30</int></rescan></config>\n</fontconfig>\n`;
}
function pptxPageCount(data) {
	let files;
	try {
		files = unzipSync(data);
	} catch (error) {
		throw new PptError("PPT_CREATE_INVALID_PACKAGE", "PPTX is not a readable ZIP package", { cause: error });
	}
	const pages = Object.keys(files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).length;
	boundedInteger(pages, "PPTX page count", 1, DEFAULT_LIMITS.maxSlides);
	inspectPptxPackage(data, pages);
	return pages;
}
async function firstExecutable(subprocess, candidates, signal) {
	for (const candidate of candidates) try {
		return await subprocess.resolveExecutable(candidate, void 0, signal);
	} catch {}
}
async function firstExisting(candidates) {
	for (const candidate of candidates) try {
		await access(candidate);
		return candidate;
	} catch {}
}
async function fingerprint(path, fallback) {
	if (path === void 0) return fallback;
	try {
		const info = await stat(path);
		return `${basename(path)}:${Math.round(info.mtimeMs)}:${info.size}`;
	} catch {
		return fallback;
	}
}
async function filesRecursively(root) {
	const output = [];
	for (const item of await readdir(root, { withFileTypes: true })) {
		const path = join(root, item.name);
		if (item.isDirectory()) output.push(...await filesRecursively(path));
		else if (item.isFile()) output.push(path);
	}
	return output;
}
function naturalPageNumber(path) {
	const numbers = basename(path).match(/\d+/gu);
	return numbers === null ? Number.MAX_SAFE_INTEGER : Number(numbers.at(-1));
}
function processFailure(name, exitCode, stdout, stderr) {
	const diagnostic = safeErrorMessage(stderr.trim() || stdout.trim(), 1e3);
	return `${name} exited with ${exitCode ?? "no exit code"}${diagnostic.length === 0 ? "" : `: ${diagnostic}`}`;
}
async function createContactSheets(workspace, previews, targetDir, signal) {
	const results = [];
	await mkdir(targetDir, { recursive: true });
	for (let start = 0; start < previews.length; start += 4) {
		throwIfAborted(signal);
		const paths = previews.slice(start, start + 4);
		const composites = await Promise.all(paths.map(async (path, index) => ({
			input: await sharp(join(workspace, path)).resize(620, 349, {
				fit: "contain",
				background: "#FFFFFF"
			}).png().toBuffer(),
			left: index % 2 * 640 + 10,
			top: Math.floor(index / 2) * 360 + 5
		})));
		const target = join(targetDir, `contact-sheet-${String(start / 4 + 1).padStart(3, "0")}.png`);
		await atomicWriteFile(target, await sharp({ create: {
			width: 1280,
			height: 720,
			channels: 4,
			background: "#E5E7EB"
		} }).composite(composites).png().toBuffer(), {
			overwrite: true,
			signal
		});
		results.push(workspaceRelative(workspace, target));
	}
	return results;
}
var PptImageRuntime = class {
	subprocess;
	sandbox;
	resources;
	executables;
	fontDirs;
	platform;
	constructor(subprocess, sandbox, resources, executables = {}, fontDirs = [], platform = process.platform) {
		this.subprocess = subprocess;
		this.sandbox = sandbox;
		this.resources = resources;
		this.executables = executables;
		this.fontDirs = fontDirs;
		this.platform = platform;
	}
	async render(owner, workspace, pptxPathInput, options = {}, signal) {
		throwIfAborted(signal);
		const pptxPath = await resolveWorkspacePath(workspace, pptxPathInput, {
			mustExist: true,
			kind: "file"
		});
		if (!pptxPath.toLowerCase().endsWith(".pptx")) throw new PptError("PPT_PATH_INVALID", "ppt_image requires a .pptx file");
		const bytes = new Uint8Array(await readFile(pptxPath));
		const pageCount = pptxPageCount(bytes);
		const hash = createHash("sha256").update(bytes).digest("hex");
		const artifactRoot = dirname(pptxPath);
		const outputDirectory = await resolveWorkspacePath(workspace, options.outputDirectory ?? join(workspaceRelative(workspace, artifactRoot), "preview", "pptx"));
		if (dirname(dirname(outputDirectory)) !== artifactRoot && dirname(outputDirectory) !== artifactRoot) throw new PptError("PPT_PATH_INVALID", "ppt_image output directory must remain inside the PPTX artifact directory");
		await mkdir(outputDirectory, { recursive: true });
		const manifestPath = join(outputDirectory, "render-manifest.json");
		const attempts = [];
		if (this.subprocess === void 0 || this.sandbox === void 0) return {
			status: "not_available",
			page_count: pageCount,
			image_paths: [],
			contact_sheet_paths: [],
			cached: false,
			attempts: [{
				backend: "libreoffice",
				status: "not_available",
				message: "DSH subprocess or sandbox service is unavailable"
			}],
			warnings: []
		};
		const requested = options.backend ?? "auto";
		for (const backend of pptImageBackendOrder(this.platform, requested)) {
			throwIfAborted(signal);
			const available = await this.discover(backend, signal);
			if (available === void 0) {
				attempts.push({
					backend,
					status: "not_available",
					message: `${backend} rendering dependencies were not found`
				});
				continue;
			}
			if (options.force !== true) {
				const cached = await this.cachedResult(workspace, manifestPath, hash, available, attempts);
				if (cached !== void 0) return cached;
			}
			const temporary = join(artifactRoot, `.ppt-image-${randomUUID()}`);
			const rawDirectory = join(temporary, "raw");
			const normalizedDirectory = join(temporary, "normalized");
			await Promise.all([mkdir(rawDirectory, { recursive: true }), mkdir(normalizedDirectory, { recursive: true })]);
			this.resources.open(owner, workspace);
			const untrack = this.resources.trackTemporaryPath(owner, temporary);
			try {
				await this.renderBackend(available, workspace, pptxPath, temporary, rawDirectory, options.nativeAutomationApproved === true, options.screenIndex, signal);
				const rawImages = (await filesRecursively(rawDirectory)).filter((path) => /\.png$/iu.test(path)).sort((left, right) => naturalPageNumber(left) - naturalPageNumber(right) || left.localeCompare(right));
				if (rawImages.length !== pageCount) throw new PptError("PPT_RENDER_FAILED", `${backend} rendered ${rawImages.length} pages, expected ${pageCount}`);
				const normalized = [];
				for (let index = 0; index < rawImages.length; index += 1) {
					throwIfAborted(signal);
					const source = rawImages[index];
					const metadata = await sharp(source, { limitInputPixels: DEFAULT_LIMITS.maxImagePixels }).metadata();
					if (metadata.width === void 0 || metadata.height === void 0) throw new PptError("PPT_RENDER_FAILED", `${basename(source)} has no decodable dimensions`);
					const ratio = metadata.width / metadata.height;
					const target = join(normalizedDirectory, `page-${String(index + 1).padStart(3, "0")}.png`);
					let pipeline = sharp(source, { limitInputPixels: DEFAULT_LIMITS.maxImagePixels }).flatten({ background: "#FFFFFF" });
					if (available.captureMethod === "screen-capture") {
						const targetRatio = 16 / 9;
						if (ratio > targetRatio) {
							const width = Math.max(1, Math.floor(metadata.height * targetRatio));
							pipeline = pipeline.extract({
								left: Math.floor((metadata.width - width) / 2),
								top: 0,
								width,
								height: metadata.height
							});
						} else {
							const height = Math.max(1, Math.floor(metadata.width / targetRatio));
							pipeline = pipeline.extract({
								left: 0,
								top: Math.floor((metadata.height - height) / 2),
								width: metadata.width,
								height
							});
						}
					} else if (Math.abs(ratio / (16 / 9) - 1) > .03) throw new PptError("PPT_RENDER_FAILED", `${basename(source)} is not a supported 16:9 slide image`);
					await pipeline.resize(1280, 720, { fit: "fill" }).toColourspace("srgb").png().toFile(target);
					normalized.push(target);
				}
				for (const name of await readdir(outputDirectory)) if (/^page-\d+\.png$/u.test(name)) await rm(join(outputDirectory, name), { force: true });
				const imagePaths = [];
				for (let index = 0; index < normalized.length; index += 1) {
					const target = join(outputDirectory, `page-${String(index + 1).padStart(3, "0")}.png`);
					await atomicWriteFile(target, await readFile(normalized[index]), {
						overwrite: true,
						signal
					});
					imagePaths.push(workspaceRelative(workspace, target));
				}
				const contactRoot = dirname(outputDirectory);
				for (const name of await readdir(contactRoot)) if (/^contact-sheet-\d+\.png$/u.test(name)) await rm(join(contactRoot, name), { force: true });
				const contactSheets = await createContactSheets(workspace, imagePaths, contactRoot, signal);
				const warnings = available.captureMethod === "screen-capture" ? ["PowerPoint screen capture is a last-resort renderer; slide animations and multi-display configuration can affect the captured frame."] : [];
				await atomicWriteJson(manifestPath, {
					version: 2,
					pptx_path: workspaceRelative(workspace, pptxPath),
					pptx_sha256: hash,
					backend,
					backend_version: available.version,
					capture_method: available.captureMethod,
					page_count: pageCount,
					image_paths: imagePaths,
					contact_sheet_paths: contactSheets,
					warnings,
					generated_at: (/* @__PURE__ */ new Date()).toISOString()
				}, {
					overwrite: true,
					signal
				});
				attempts.push({
					backend,
					capture_method: available.captureMethod,
					status: "passed",
					message: `${backend} rendered ${pageCount} pages`
				});
				return {
					status: "passed",
					backend,
					backend_version: available.version,
					capture_method: available.captureMethod,
					page_count: pageCount,
					image_paths: imagePaths,
					contact_sheet_paths: contactSheets,
					manifest_path: workspaceRelative(workspace, manifestPath),
					cached: false,
					attempts,
					warnings
				};
			} catch (error) {
				if (signal?.aborted) throwIfAborted(signal);
				attempts.push({
					backend,
					capture_method: available.captureMethod,
					status: "failed",
					message: safeErrorMessage(error)
				});
			} finally {
				untrack();
				await rm(temporary, {
					recursive: true,
					force: true
				});
			}
		}
		return {
			status: attempts.some((attempt) => attempt.status === "failed") ? "failed" : "not_available",
			page_count: pageCount,
			image_paths: [],
			contact_sheet_paths: [],
			cached: false,
			attempts,
			warnings: []
		};
	}
	async discover(backend, signal) {
		if (backend === "keynote") {
			if (this.platform !== "darwin") return void 0;
			const [osascript, keynote] = await Promise.all([firstExecutable(this.subprocess, this.executables.osascript ?? appleScriptCandidates(this.platform), signal), firstExisting(this.executables.keynote ?? keynoteCandidates(this.platform))]);
			if (osascript === void 0 || keynote === void 0) return void 0;
			return {
				backend,
				captureMethod: "native-export",
				version: await fingerprint(keynote, "Keynote"),
				executables: {
					osascript,
					keynote
				}
			};
		}
		if (backend === "powerpoint") {
			if (this.platform === "darwin") {
				const [osascript, powerpoint, screencapture] = await Promise.all([
					firstExecutable(this.subprocess, this.executables.osascript ?? appleScriptCandidates(this.platform), signal),
					firstExisting(this.executables.powerpoint ?? powerPointCandidates(this.platform)),
					firstExecutable(this.subprocess, this.executables.screencapture ?? screenCaptureCandidates(this.platform), signal)
				]);
				if (osascript === void 0 || powerpoint === void 0 || screencapture === void 0) return void 0;
				return {
					backend,
					captureMethod: "screen-capture",
					version: await fingerprint(powerpoint, "PowerPoint-screen-capture"),
					executables: {
						osascript,
						powerpoint,
						screencapture
					}
				};
			}
			if (this.platform !== "win32") return void 0;
			const powershell = await firstExecutable(this.subprocess, this.executables.powershell ?? powerShellCandidates(this.platform), signal);
			if (powershell === void 0) return void 0;
			const powerpoint = await firstExisting(this.executables.powerpoint ?? powerPointCandidates(this.platform));
			return {
				backend,
				captureMethod: "native-export",
				version: await fingerprint(powerpoint, "PowerPoint-COM"),
				executables: {
					powershell,
					...powerpoint === void 0 ? {} : { powerpoint }
				}
			};
		}
		const [soffice, pdftoppm] = await Promise.all([firstExecutable(this.subprocess, this.executables.soffice ?? libreOfficeCandidates(this.platform), signal), firstExecutable(this.subprocess, this.executables.pdftoppm ?? pdfToPpmCandidates(this.platform), signal)]);
		if (soffice === void 0 || pdftoppm === void 0) return void 0;
		return {
			backend,
			captureMethod: "pdf-raster",
			version: await fingerprint(soffice, "LibreOffice"),
			executables: {
				soffice,
				pdftoppm
			}
		};
	}
	async cachedResult(workspace, manifestPath, hash, available, attempts) {
		try {
			const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
			if (manifest.version !== 2 || manifest.pptx_sha256 !== hash || manifest.backend !== available.backend || manifest.backend_version !== available.version || manifest.capture_method !== available.captureMethod) return void 0;
			for (const path of [...manifest.image_paths, ...manifest.contact_sheet_paths]) await resolveWorkspacePath(workspace, path, {
				mustExist: true,
				kind: "file"
			});
			attempts.push({
				backend: available.backend,
				capture_method: available.captureMethod,
				status: "passed",
				message: "reused complete render cache"
			});
			return {
				status: "passed",
				backend: manifest.backend,
				backend_version: manifest.backend_version,
				capture_method: manifest.capture_method,
				page_count: manifest.page_count,
				image_paths: manifest.image_paths,
				contact_sheet_paths: manifest.contact_sheet_paths,
				manifest_path: workspaceRelative(workspace, manifestPath),
				cached: true,
				attempts,
				warnings: manifest.warnings
			};
		} catch {
			return;
		}
	}
	async renderBackend(available, workspace, pptxPath, temporary, rawDirectory, nativeAutomationApproved, screenIndexInput, signal) {
		if (available.backend === "keynote") {
			const script = join(temporary, "render.applescript");
			await atomicWriteText(script, KEYNOTE_SCRIPT, { signal });
			const output = join(rawDirectory, "export");
			const argv = [
				available.executables.osascript,
				script,
				pptxPath,
				output
			];
			const command = nativeAutomationApproved ? argv : (await this.sandbox.confine(argv, {
				mode: "workspace-write",
				workspaceRoot: workspace
			})).argv;
			const result = await runCollected(this.subprocess, command, {
				cwd: dirname(pptxPath),
				signal,
				timeoutMs: 9e4,
				maxOutputBytes: 32768
			});
			if (result.exitCode !== 0) throw new PptError("PPT_RENDER_FAILED", processFailure("Keynote export", result.exitCode, result.stdout, result.stderr));
			return;
		}
		if (available.backend === "powerpoint") {
			if (this.platform === "darwin") {
				const screenIndex = screenIndexInput ?? 1;
				boundedInteger(screenIndex, "screen_index", 1, 16);
				const script = join(temporary, "render-powerpoint-screen.applescript");
				const sourceCopy = join(temporary, "source.pptx");
				await copyFile(pptxPath, sourceCopy);
				await atomicWriteText(script, POWERPOINT_MAC_SCREEN_SCRIPT, { signal });
				const argv = [
					available.executables.osascript,
					script,
					sourceCopy,
					rawDirectory,
					String(pptxPageCount(new Uint8Array(await readFile(sourceCopy)))),
					available.executables.screencapture,
					String(screenIndex)
				];
				const command = nativeAutomationApproved ? argv : (await this.sandbox.confine(argv, {
					mode: "workspace-write",
					workspaceRoot: workspace
				})).argv;
				const result = await runCollected(this.subprocess, command, {
					cwd: dirname(pptxPath),
					signal,
					timeoutMs: 18e4,
					maxOutputBytes: 32768
				});
				if (result.exitCode !== 0) throw new PptError("PPT_RENDER_FAILED", processFailure("PowerPoint screen capture", result.exitCode, result.stdout, result.stderr));
				return;
			}
			const script = join(temporary, "render.ps1");
			const output = join(rawDirectory, "export");
			await mkdir(output, { recursive: true });
			await atomicWriteText(script, POWERPOINT_SCRIPT, { signal });
			const argv = [
				available.executables.powershell,
				"-NoProfile",
				"-NonInteractive",
				"-ExecutionPolicy",
				"Bypass",
				"-File",
				script,
				"-InputPptx",
				pptxPath,
				"-OutputDir",
				output
			];
			const command = nativeAutomationApproved ? argv : (await this.sandbox.confine(argv, {
				mode: "workspace-write",
				workspaceRoot: workspace
			})).argv;
			const result = await runCollected(this.subprocess, command, {
				cwd: dirname(pptxPath),
				signal,
				timeoutMs: 9e4,
				maxOutputBytes: 32768
			});
			if (result.exitCode !== 0) throw new PptError("PPT_RENDER_FAILED", processFailure("PowerPoint export", result.exitCode, result.stdout, result.stderr));
			return;
		}
		const profile = join(temporary, "profile");
		const fontCache = join(temporary, "font-cache");
		const fontconfig = join(temporary, "fonts.conf");
		const pdfDirectory = join(temporary, "pdf");
		await Promise.all([
			mkdir(profile, { recursive: true }),
			mkdir(fontCache, { recursive: true }),
			mkdir(pdfDirectory, { recursive: true })
		]);
		const renderFontDirs = [...systemFontDirectories(this.platform), ...this.fontDirs];
		await atomicWriteText(fontconfig, fontconfigDocument(renderFontDirs, fontCache), { signal });
		const renderEnv = {
			FONTCONFIG_FILE: fontconfig,
			FONTCONFIG_PATH: temporary,
			XDG_CACHE_HOME: fontCache,
			SAL_PRIVATE_FONTPATH: renderFontDirs.join(delimiter)
		};
		const convert = await this.sandbox.confine([
			available.executables.soffice,
			"--headless",
			"--nologo",
			"--nodefault",
			"--nolockcheck",
			"--norestore",
			`-env:UserInstallation=${pathToFileURL(profile).href}`,
			"--convert-to",
			"pdf",
			"--outdir",
			pdfDirectory,
			pptxPath
		], {
			mode: "workspace-write",
			workspaceRoot: workspace
		});
		if (convert.enforcement !== "full") throw new PptError("PPT_RENDER_FAILED", "LibreOffice sandbox enforcement is partial");
		const converted = await runCollected(this.subprocess, convert.argv, {
			cwd: dirname(pptxPath),
			signal,
			timeoutMs: 6e4,
			maxOutputBytes: 32768,
			env: renderEnv
		});
		if (converted.exitCode !== 0) throw new PptError("PPT_RENDER_FAILED", processFailure("LibreOffice", converted.exitCode, converted.stdout, converted.stderr));
		const pdf = join(pdfDirectory, `${basename(pptxPath, ".pptx")}.pdf`);
		await access(pdf);
		const raster = await this.sandbox.confine([
			available.executables.pdftoppm,
			"-png",
			"-r",
			"96",
			"-scale-to-x",
			"1280",
			"-scale-to-y",
			"720",
			"-cropbox",
			pdf,
			join(rawDirectory, "page")
		], {
			mode: "workspace-write",
			workspaceRoot: workspace
		});
		const rasterized = await runCollected(this.subprocess, raster.argv, {
			cwd: dirname(pptxPath),
			signal,
			timeoutMs: 6e4,
			maxOutputBytes: 32768
		});
		if (rasterized.exitCode !== 0) throw new PptError("PPT_RENDER_FAILED", processFailure("pdftoppm", rasterized.exitCode, rasterized.stdout, rasterized.stderr));
	}
};
//#endregion
//#region src/quality.ts
const visualReviewSchema = z.object({
	version: z.literal(1),
	status: z.enum([
		"not_performed",
		"passed",
		"failed",
		"not_available"
	]),
	checklist: z.array(z.string().min(1).max(500)).max(30),
	reviewed_assets: z.array(z.string().min(1).max(500)).max(200),
	findings: z.array(z.object({
		page: z.number().int().positive().optional(),
		severity: z.enum(["warning", "error"]),
		message: z.string().min(1).max(1e3)
	}).strict()).max(500),
	completed_at: z.iso.datetime({ offset: true }).optional()
}).strict();
function computeOverallStatus(report) {
	const statuses = [
		report.structural_status,
		report.render_status,
		report.automatic_visual_status,
		report.model_visual_status
	];
	if (statuses.includes("failed")) return "failed";
	return statuses.every((status) => status === "passed") ? "verified" : "unverified";
}
function synchronizeStatuses(report) {
	report.structural_status = report.layers.structural.status;
	report.render_status = report.layers.render.status;
	report.automatic_visual_status = report.layers.automatic_visual.status;
	report.model_visual_status = report.layers.model_visual.status;
	report.overall_status = computeOverallStatus(report);
}
function structuralFindings(data, expectedPages) {
	inspectPptxPackage(data, expectedPages);
	const files = unzipSync(data);
	const findings = [];
	for (let page = 1; page <= expectedPages; page += 1) {
		const xml = strFromU8(files[`ppt/slides/slide${page}.xml`]);
		for (const block of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/gu)) {
			if (!/<p:txBody\b/u.test(block[0]) || !/\btxBox="1"/u.test(block[0])) continue;
			if ([...block[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/gu)].map((match) => match[1].replace(/<[^>]+>/gu, "")).join("").trim().length === 0) findings.push({
				code: "EMPTY_TEXT_BOX",
				severity: "error",
				message: "slide contains an empty text box",
				page
			});
		}
		const pictures = [...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/gu)];
		const textCount = [...xml.matchAll(/<a:t>\s*[^<\s][\s\S]*?<\/a:t>/gu)].length;
		if (pictures.length === 1 && textCount === 0 && /<a:off x="0" y="0"\/>[\s\S]*?<a:ext cx="12192000" cy="6858000"\/>/u.test(pictures[0][0])) findings.push({
			code: "FULL_PAGE_RASTER",
			severity: "error",
			message: "slide degraded to a single full-page image",
			page
		});
	}
	const allXml = Object.entries(files).filter(([name]) => name.endsWith(".xml")).map(([, bytes]) => strFromU8(bytes)).join("\n");
	const visibleText = [...allXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/gu)].map((match) => match[1]).join("\n");
	if (/(?:�|\bTODO\b|\bTBD\b|\bPLACEHOLDER\b|待补(?:充|数据)?)/iu.test(visibleText)) findings.push({
		code: "PLACEHOLDER_TEXT",
		severity: "error",
		message: "placeholder or replacement text remains in the PPTX"
	});
	const sizes = [...allXml.matchAll(/<(?:a:rPr|a:defRPr)\b[^>]*\bsz="(\d+)"/gu)].map((match) => Number(match[1]) / 100);
	if (sizes.some((size) => size > 0 && size < 10)) findings.push({
		code: "SMALL_FONT",
		severity: "warning",
		message: `minimum detected font size is ${Math.min(...sizes).toFixed(1)}pt`
	});
	return findings;
}
async function imageQualityFindings(data, expectedPages, signal) {
	const files = unzipSync(data);
	const findings = [];
	for (let page = 1; page <= expectedPages; page += 1) {
		throwIfAborted(signal);
		const slide = strFromU8(files[`ppt/slides/slide${page}.xml`]);
		const relName = `ppt/slides/_rels/slide${page}.xml.rels`;
		const rels = files[relName] === void 0 ? "" : strFromU8(files[relName]);
		const targets = /* @__PURE__ */ new Map();
		for (const match of rels.matchAll(/<Relationship\b([^>]*)>/giu)) {
			const id = /\bId="([^"]+)"/iu.exec(match[1])?.[1];
			const target = /\bTarget="([^"]+)"/iu.exec(match[1])?.[1];
			if (id !== void 0 && target !== void 0) targets.set(id, posix.normalize(posix.join("ppt/slides", target)));
		}
		for (const match of slide.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/gu)) {
			const block = match[0];
			const relationId = /<a:blip\b[^>]*\br:embed="([^"]+)"/iu.exec(block)?.[1];
			const extent = /<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/iu.exec(block);
			const target = relationId === void 0 ? void 0 : targets.get(relationId);
			if (extent === null || target === void 0 || files[target] === void 0 || target.endsWith(".svg")) continue;
			let metadata;
			try {
				metadata = await sharp(files[target]).metadata();
			} catch {
				continue;
			}
			if (metadata.width === void 0 || metadata.height === void 0) continue;
			const displayWidth = Number(extent[1]) / 914400 * 96;
			const displayHeight = Number(extent[2]) / 914400 * 96;
			if (metadata.width + 1 < displayWidth || metadata.height + 1 < displayHeight) findings.push({
				code: "LOW_RES_IMAGE",
				severity: "warning",
				message: `${basename(target)} is ${metadata.width}x${metadata.height} but displayed near ${Math.round(displayWidth)}x${Math.round(displayHeight)}`,
				page
			});
			const sourceRatio = metadata.width / metadata.height;
			const displayRatio = displayWidth / displayHeight;
			if (!/<a:srcRect\b/u.test(block) && Math.abs(sourceRatio / displayRatio - 1) > .15) findings.push({
				code: "STRETCHED_IMAGE",
				severity: "error",
				message: `${basename(target)} aspect ratio does not match its uncropped display box`,
				page
			});
		}
	}
	return findings;
}
async function analyzePage(path, page, signal) {
	throwIfAborted(signal);
	const { data, info } = await sharp(path).removeAlpha().raw().toBuffer({ resolveWithObject: true });
	throwIfAborted(signal);
	if (info.width !== 1280 || info.height !== 720) return {
		metrics: {
			page,
			width: info.width,
			height: info.height
		},
		findings: [{
			code: "RENDER_SIZE",
			severity: "error",
			message: `rendered page is ${info.width}x${info.height}, expected 1280x720`,
			page
		}]
	};
	const channels = info.channels;
	const background = [
		data[0],
		data[1],
		data[2]
	];
	let sum = 0;
	let sumSquares = 0;
	let changed = 0;
	let black = 0;
	let edgeChanged = 0;
	let edgePixels = 0;
	for (let y = 0; y < info.height; y += 1) for (let x = 0; x < info.width; x += 1) {
		const offset = (y * info.width + x) * channels;
		const r = data[offset];
		const g = data[offset + 1];
		const b = data[offset + 2];
		const luminance = (r + g + b) / 3;
		sum += luminance;
		sumSquares += luminance * luminance;
		const differs = Math.abs(r - background[0]) + Math.abs(g - background[1]) + Math.abs(b - background[2]) > 45;
		if (differs) changed += 1;
		if (luminance < 5) black += 1;
		if (x < 4 || y < 4 || x >= info.width - 4 || y >= info.height - 4) {
			edgePixels += 1;
			if (differs) edgeChanged += 1;
		}
	}
	const pixels = info.width * info.height;
	const mean = sum / pixels;
	const stdev = Math.sqrt(Math.max(0, sumSquares / pixels - mean * mean));
	const coverage = changed / pixels;
	const blackRatio = black / pixels;
	const edgeRatio = edgeChanged / edgePixels;
	const findings = [];
	if (coverage < .005 || stdev < 1) findings.push({
		code: "BLANK_PAGE",
		severity: "error",
		message: "rendered page is blank or nearly uniform",
		page
	});
	if (blackRatio > .98) findings.push({
		code: "BLACK_PAGE",
		severity: "error",
		message: "rendered page is almost entirely black",
		page
	});
	if (edgeRatio > .12) findings.push({
		code: "EDGE_CONTENT",
		severity: "warning",
		message: "significant content touches the slide boundary",
		page
	});
	return {
		metrics: {
			page,
			width: info.width,
			height: info.height,
			mean,
			stdev,
			content_coverage: coverage,
			black_ratio: blackRatio,
			edge_ratio: edgeRatio
		},
		findings
	};
}
async function compareImages(htmlPath, pptxPath, page, signal) {
	throwIfAborted(signal);
	const left = await sharp(htmlPath).removeAlpha().raw().toBuffer({ resolveWithObject: true });
	const right = await sharp(pptxPath).removeAlpha().raw().toBuffer({ resolveWithObject: true });
	throwIfAborted(signal);
	if (left.info.width !== right.info.width || left.info.height !== right.info.height || left.info.channels !== right.info.channels) return {
		metrics: { page },
		findings: [{
			code: "PREVIEW_DIMENSION_MISMATCH",
			severity: "error",
			message: "HTML and PPTX preview dimensions differ",
			page
		}]
	};
	let total = 0;
	let gross = 0;
	for (let index = 0; index < left.data.length; index += 1) {
		const delta = Math.abs(left.data[index] - right.data[index]);
		total += delta;
		if (delta > 64) gross += 1;
	}
	const meanAbsoluteDifference = total / left.data.length;
	const grossDifferenceRatio = gross / left.data.length;
	const findings = [];
	if (meanAbsoluteDifference > 45 || grossDifferenceRatio > .35) findings.push({
		code: "LAYOUT_DRIFT",
		severity: "error",
		message: `HTML/PPTX render drift is too large (MAD ${meanAbsoluteDifference.toFixed(1)}, gross ${(grossDifferenceRatio * 100).toFixed(1)}%)`,
		page
	});
	else if (meanAbsoluteDifference > 25 || grossDifferenceRatio > .18) findings.push({
		code: "LAYOUT_DRIFT_WARNING",
		severity: "warning",
		message: "HTML/PPTX render difference is elevated",
		page
	});
	return {
		metrics: {
			page,
			mean_absolute_difference: meanAbsoluteDifference,
			gross_difference_ratio: grossDifferenceRatio
		},
		findings
	};
}
function cjkTextRegions(data, expectedPages) {
	const files = unzipSync(data);
	const regions = [];
	for (let page = 1; page <= expectedPages; page += 1) {
		const xml = strFromU8(files[`ppt/slides/slide${page}.xml`]);
		for (const match of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/gu)) {
			const block = match[0];
			const text = [...block.matchAll(/<a:t>([\s\S]*?)<\/a:t>/gu)].map((item) => item[1]).join("");
			if (!/[\u3400-\u9FFF\uF900-\uFAFF]/u.test(text)) continue;
			const transform = /<a:xfrm\b[^>]*>[\s\S]*?<a:off\b[^>]*\bx="(\d+)"[^>]*\by="(\d+)"[^>]*\/>[\s\S]*?<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"[^>]*\/>/u.exec(block);
			if (transform === null) continue;
			regions.push({
				page,
				x: Number(transform[1]) / 9525,
				y: Number(transform[2]) / 9525,
				width: Number(transform[3]) / 9525,
				height: Number(transform[4]) / 9525
			});
		}
	}
	return regions;
}
async function regionEdgeCount(path, region) {
	const left = Math.max(0, Math.min(1279, Math.floor(region.x)));
	const top = Math.max(0, Math.min(719, Math.floor(region.y)));
	const width = Math.max(1, Math.min(1280 - left, Math.ceil(region.width)));
	const height = Math.max(1, Math.min(720 - top, Math.ceil(region.height)));
	const { data, info } = await sharp(path).extract({
		left,
		top,
		width,
		height
	}).removeAlpha().raw().toBuffer({ resolveWithObject: true });
	let edges = 0;
	for (let y = 0; y < height - 1; y += 1) for (let x = 0; x < width - 1; x += 1) {
		const current = (y * width + x) * info.channels;
		const right = current + info.channels;
		const down = current + width * info.channels;
		let delta = 0;
		for (let channel = 0; channel < Math.min(3, info.channels); channel += 1) {
			delta += Math.abs(data[current + channel] - data[right + channel]);
			delta += Math.abs(data[current + channel] - data[down + channel]);
		}
		if (delta > 80) edges += 1;
	}
	return edges;
}
async function cjkRenderFindings(pptx, htmlPreviews, pptxPreviews, expectedPages, signal) {
	const byPage = /* @__PURE__ */ new Map();
	for (const region of cjkTextRegions(pptx, expectedPages)) byPage.set(region.page, [...byPage.get(region.page) ?? [], region]);
	const findings = [];
	for (const [page, regions] of byPage) {
		throwIfAborted(signal);
		let htmlEdges = 0;
		let renderedEdges = 0;
		for (const region of regions) {
			htmlEdges += await regionEdgeCount(htmlPreviews[page - 1], region);
			renderedEdges += await regionEdgeCount(pptxPreviews[page - 1], region);
		}
		if (htmlEdges < 100) continue;
		const ratio = renderedEdges / htmlEdges;
		if (ratio < .45) findings.push({
			code: "CJK_GLYPHS_MISSING",
			severity: "error",
			message: `rendered CJK text edge coverage is only ${(ratio * 100).toFixed(1)}% of the HTML preview`,
			page
		});
		else if (ratio < .65) findings.push({
			code: "CJK_FONT_SUBSTITUTION",
			severity: "warning",
			message: `rendered CJK text edge coverage is reduced to ${(ratio * 100).toFixed(1)}% of the HTML preview`,
			page
		});
	}
	return findings;
}
var QualityRuntime = class {
	pptImage;
	constructor(subprocess, sandbox, resources, executables = {}, fontDirs = [], pptImage) {
		this.pptImage = pptImage ?? new PptImageRuntime(subprocess, sandbox, resources, executables, fontDirs);
	}
	async refresh(owner, workspace, pptxPathInput, modelReviewAvailable, signal, nativeAutomationApproved = false) {
		throwIfAborted(signal);
		const pptxPath = await resolveWorkspacePath(workspace, pptxPathInput, {
			mustExist: true,
			kind: "file"
		});
		const artifactRoot = dirname(pptxPath);
		const reportPath = join(artifactRoot, "report.json");
		const visualReviewPath = join(artifactRoot, "visual-review.json");
		let existing;
		try {
			existing = JSON.parse(await readFile(reportPath, "utf8"));
		} catch (error) {
			if (error.code === "ENOENT") return void 0;
			throw error;
		}
		if (existing.version !== 1 || existing.machine_owned !== true || existing.pptx_path !== workspaceRelative(workspace, pptxPath)) throw new PptError("PPT_QUALITY_FAILED", "existing report.json is not the machine report for this PPTX");
		const expectedPages = pptxPageCount(new Uint8Array(await readFile(pptxPath)));
		const htmlPreviews = Array.from({ length: expectedPages }, (_, index) => workspaceRelative(workspace, join(artifactRoot, "preview", `page-${String(index + 1).padStart(3, "0")}.png`)));
		await Promise.all(htmlPreviews.map((path) => resolveWorkspacePath(workspace, path, {
			mustExist: true,
			kind: "file"
		})));
		return this.evaluate(owner, workspace, workspaceRelative(workspace, pptxPath), htmlPreviews, workspaceRelative(workspace, reportPath), workspaceRelative(workspace, visualReviewPath), expectedPages, modelReviewAvailable, existing.conversion, signal, nativeAutomationApproved);
	}
	async evaluate(owner, workspace, pptxPathInput, htmlPreviewInputs, reportPathInput, visualReviewPathInput, expectedPages, modelReviewAvailable, conversion, signal, nativeAutomationApproved = false) {
		throwIfAborted(signal);
		expectedPages = boundedInteger(expectedPages, "expectedPages", 1, DEFAULT_LIMITS.maxSlides);
		const pptxPath = await resolveWorkspacePath(workspace, pptxPathInput, {
			mustExist: true,
			kind: "file"
		});
		const reportPath = await resolveWorkspacePath(workspace, reportPathInput);
		const visualReviewPath = await resolveWorkspacePath(workspace, visualReviewPathInput);
		const artifactRoot = dirname(pptxPath);
		let designPlan;
		try {
			designPlan = validateArtDirection(JSON.parse(await readFile(join(artifactRoot, "design-plan.json"), "utf8")), expectedPages);
		} catch (error) {
			if (error.code !== "ENOENT") throw error;
		}
		let htmlDesignValidation = {};
		try {
			htmlDesignValidation = JSON.parse(await readFile(join(artifactRoot, "design-validation.json"), "utf8"));
		} catch (error) {
			if (error.code !== "ENOENT") throw error;
		}
		const designFindings = [...designPlan === void 0 ? [] : artDirectionFindings(designPlan), ...htmlDesignValidation.findings ?? []].filter((finding, index, all) => all.findIndex((other) => other.code === finding.code && other.page === finding.page) === index);
		const htmlPreviews = await Promise.all(htmlPreviewInputs.map(async (path) => workspaceRelative(workspace, await resolveWorkspacePath(workspace, path, {
			mustExist: true,
			kind: "file"
		}))));
		const report = {
			version: 1,
			machine_owned: true,
			generated_at: (/* @__PURE__ */ new Date()).toISOString(),
			pptx_path: workspaceRelative(workspace, pptxPath),
			structural_status: "not_performed",
			render_status: "not_performed",
			automatic_visual_status: "not_performed",
			model_visual_status: modelReviewAvailable ? "not_performed" : "not_available",
			overall_status: "unverified",
			layers: {
				structural: {
					status: "not_performed",
					findings: []
				},
				render: {
					status: "not_performed",
					findings: []
				},
				automatic_visual: {
					status: "not_performed",
					findings: [],
					pages: [],
					html_comparison_pages: [],
					design_fidelity: {
						mode: designPlan === void 0 ? "legacy" : "directed",
						checks: designPlan === void 0 ? [] : [
							"art-direction-schema",
							"page-sequence",
							"composition-rhythm",
							"frame-budget",
							"review-checklist",
							...htmlDesignValidation.checks ?? []
						],
						pages: htmlDesignValidation.pages ?? designPlan?.slides.map((slide) => ({
							page: slide.page,
							composition: slide.composition,
							density: slide.density,
							background_role: slide.background_role,
							visual_anchor: slide.visual_anchor.kind,
							frame_policy: slide.frame_policy
						})) ?? [],
						findings: designFindings
					}
				},
				model_visual: {
					status: modelReviewAvailable ? "not_performed" : "not_available",
					findings: []
				}
			},
			artifacts: {
				html_previews: htmlPreviews,
				pptx_previews: [],
				contact_sheets: [],
				high_risk_previews: [],
				visual_review: workspaceRelative(workspace, visualReviewPath)
			},
			...conversion === void 0 ? {} : { conversion }
		};
		try {
			const bytes = new Uint8Array(await readFile(pptxPath));
			report.layers.structural.findings = [...structuralFindings(bytes, expectedPages), ...await imageQualityFindings(bytes, expectedPages, signal)];
			report.layers.structural.status = report.layers.structural.findings.some((item) => item.severity === "error") ? "failed" : "passed";
		} catch (error) {
			report.layers.structural.status = "failed";
			report.layers.structural.findings.push({
				code: "STRUCTURE_INVALID",
				severity: "error",
				message: error instanceof Error ? error.message : String(error)
			});
		}
		const review = {
			version: 1,
			status: modelReviewAvailable ? "not_performed" : "not_available",
			checklist: [
				"Inspect every contact sheet for consistency and blank, black, duplicated, or missing pages.",
				"Inspect cover, agenda, data/chart, comparison, ending, and every machine-flagged page at full resolution.",
				"Check text clipping, overlap, hierarchy, alignment, contrast, image quality, stretched images, and font substitution.",
				"Record reviewed asset paths and page-specific findings; do not edit report.json directly.",
				...designPlan === void 0 ? ["This deck has no design-plan.json; record that Art Direction fidelity could not be verified."] : artDirectionReviewChecklist(designPlan).slice(0, 26)
			],
			reviewed_assets: [],
			findings: []
		};
		const rendered = await this.pptImage.render(owner, workspace, workspaceRelative(workspace, pptxPath), {
			backend: "auto",
			nativeAutomationApproved
		}, signal);
		if (rendered.status === "not_available") {
			report.layers.render.status = "not_available";
			report.layers.render.findings.push({
				code: "RENDERER_NOT_AVAILABLE",
				severity: "warning",
				message: rendered.attempts.map((attempt) => `${attempt.backend}: ${attempt.message}`).join("; ") || "no supported PPTX renderer was found"
			});
			report.layers.automatic_visual.status = "not_available";
		} else if (rendered.status === "failed") {
			report.layers.render.status = "failed";
			report.layers.render.findings.push({
				code: "RENDER_FAILED",
				severity: "error",
				message: rendered.attempts.map((attempt) => `${attempt.backend}: ${attempt.message}`).join("; ")
			});
			report.layers.automatic_visual.status = "not_performed";
		} else {
			report.layers.render.status = "passed";
			report.layers.render.name = rendered.backend;
			report.layers.render.version = rendered.backend_version;
			report.artifacts.pptx_previews = rendered.image_paths;
			report.artifacts.contact_sheets = rendered.contact_sheet_paths;
			try {
				const analyses = await Promise.all(report.artifacts.pptx_previews.map((path, index) => analyzePage(join(workspace, path), index + 1, signal)));
				report.layers.automatic_visual.pages = analyses.map((item) => item.metrics);
				const comparisons = await Promise.all(report.artifacts.pptx_previews.map((path, index) => compareImages(join(workspace, htmlPreviews[index]), join(workspace, path), index + 1, signal)));
				const cjkFindings = await cjkRenderFindings(new Uint8Array(await readFile(pptxPath)), htmlPreviews.map((path) => join(workspace, path)), report.artifacts.pptx_previews.map((path) => join(workspace, path)), expectedPages, signal);
				report.layers.automatic_visual.html_comparison_pages = comparisons.map((item) => item.metrics);
				report.layers.automatic_visual.findings = [
					...designFindings,
					...analyses.flatMap((item) => item.findings),
					...comparisons.flatMap((item) => item.findings),
					...cjkFindings
				];
				report.layers.automatic_visual.status = report.layers.automatic_visual.findings.some((item) => item.severity === "error") ? "failed" : "passed";
				const risky = /* @__PURE__ */ new Set([...report.layers.structural.findings.flatMap((item) => item.page === void 0 ? [] : [item.page]), ...report.layers.automatic_visual.findings.flatMap((item) => item.page === void 0 ? [] : [item.page])]);
				for (const page of [...risky].sort((a, b) => a - b)) {
					throwIfAborted(signal);
					const source = join(workspace, report.artifacts.pptx_previews[page - 1]);
					const target = join(artifactRoot, "preview", `high-risk-page-${String(page).padStart(3, "0")}.png`);
					await copyFile(source, target);
					report.artifacts.high_risk_previews.push(workspaceRelative(workspace, target));
				}
			} catch (error) {
				if (signal?.aborted) throwIfAborted(signal);
				report.layers.automatic_visual.status = "failed";
				report.layers.automatic_visual.findings.push({
					code: "AUTOMATIC_VISUAL_FAILED",
					severity: "error",
					message: error instanceof Error ? error.message : String(error)
				});
			}
		}
		synchronizeStatuses(report);
		await atomicWriteJson(visualReviewPath, review, { overwrite: true });
		await atomicWriteJson(reportPath, report, { overwrite: true });
		return report;
	}
};
async function applyVisualReview(workspace, reportPathInput, reviewPathInput) {
	const reportPath = await resolveWorkspacePath(workspace, reportPathInput, {
		mustExist: true,
		kind: "file"
	});
	const reviewPath = await resolveWorkspacePath(workspace, reviewPathInput, {
		mustExist: true,
		kind: "file"
	});
	const report = JSON.parse(await readFile(reportPath, "utf8"));
	const parsed = visualReviewSchema.safeParse(JSON.parse(await readFile(reviewPath, "utf8")));
	if (!parsed.success) throw new PptError("PPT_QUALITY_FAILED", "visual-review.json does not match the required schema", { details: parsed.error.flatten() });
	const review = parsed.data;
	if (review.version !== 1 || ![
		"passed",
		"failed",
		"not_available"
	].includes(review.status)) throw new PptError("PPT_QUALITY_FAILED", "visual-review.json has not been completed with a valid status");
	if (review.status === "passed" && review.reviewed_assets.length === 0) throw new PptError("PPT_QUALITY_FAILED", "a passed visual review must record reviewed asset paths");
	if (review.status === "passed" && review.findings.some((item) => item.severity === "error")) throw new PptError("PPT_QUALITY_FAILED", "a passed visual review cannot contain error findings");
	const reviewedAssets = /* @__PURE__ */ new Set();
	for (const asset of review.reviewed_assets) reviewedAssets.add(workspaceRelative(workspace, await resolveWorkspacePath(workspace, asset, {
		mustExist: true,
		kind: "file"
	})));
	if (review.status === "passed") {
		const missing = [...report.artifacts.contact_sheets, ...report.artifacts.high_risk_previews].filter((asset) => !reviewedAssets.has(asset));
		if (missing.length > 0) throw new PptError("PPT_QUALITY_FAILED", "a passed visual review must cover every contact sheet and high-risk preview", { details: { missing } });
	}
	report.layers.model_visual = {
		status: review.status === "passed" ? "passed" : review.status === "failed" ? "failed" : "not_available",
		findings: review.findings.map((item) => ({
			code: "MODEL_VISUAL_REVIEW",
			severity: item.severity,
			message: item.message,
			...item.page === void 0 ? {} : { page: item.page }
		}))
	};
	synchronizeStatuses(report);
	await atomicWriteJson(reportPath, report, { overwrite: true });
	return report;
}
//#endregion
export { DEFAULT_LIMITS as A, appleScriptCandidates as B, validateTheme as C, isPathInside as D, isLocalFilesystemPath as E, buildFontCatalog as F, pdfToPpmCandidates as G, isSupportedPlatform as H, discoverRegisteredFonts as I, screenCaptureCandidates as J, powerPointCandidates as K, installedFontsAsDiscovered as L, validatePublicHttpUrl as M, runCollected as N, resolveWorkspacePath as O, safeErrorMessage as P, registeredFont as R, themeSummary as S, describeImageSearchDegradation as T, keynoteCandidates as U, browserSystemCandidates as V, libreOfficeCandidates as W, systemFontDirectories as Y, contrastRatio as _, TEXT_ANIMATION_DIRECTIONS as a, planThemePages as b, planSlideAnimations as c, SLIDE_TRANSITION_TYPES as d, createHtmlDeck as f, PPT_THEMES as g, SLIDE_TYPES as h, createPptx as i, boundedInteger as j, workspaceRelative as k, SLIDE_TRANSITION_DIRECTIONS as l, SLIDE_LAYOUTS as m, applyVisualReview as n, TEXT_ANIMATION_EFFECTS as o, writePptOutline as p, powerShellCandidates as q, PptImageRuntime as r, TEXT_ANIMATION_STARTS as s, QualityRuntime as t, SLIDE_TRANSITION_SPEEDS as u, findTheme as v, ImageSearchRuntime as w, themeFindings as x, listThemes as y, summarizeFontAvailability as z };

//# sourceMappingURL=quality-BOrmVVNM.mjs.map