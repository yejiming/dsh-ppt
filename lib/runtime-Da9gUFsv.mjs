import { i as throwIfAborted, n as PptError, r as asPptError } from "./errors-BdtilPdq.mjs";
import { A as DEFAULT_LIMITS, B as appleScriptCandidates, D as isPathInside, E as isLocalFilesystemPath, G as pdfToPpmCandidates, H as isSupportedPlatform, I as discoverRegisteredFonts, J as screenCaptureCandidates, K as powerPointCandidates, M as validatePublicHttpUrl, N as runCollected, O as resolveWorkspacePath, P as safeErrorMessage, R as registeredFont, U as keynoteCandidates, V as browserSystemCandidates, W as libreOfficeCandidates, j as boundedInteger, k as workspaceRelative, q as powerShellCandidates, r as PptImageRuntime, t as QualityRuntime, w as ImageSearchRuntime, z as summarizeFontAvailability } from "./quality-BOrmVVNM.mjs";
import { PPT_MODE_TOOL_NAMES } from "./schemas.mjs";
import { extname, join } from "node:path";
import { access, mkdir, readdir, rm, stat } from "node:fs/promises";
import { constants } from "node:fs";
import { chromium } from "playwright-core";
import { fileURLToPath, pathToFileURL } from "node:url";
import sharp from "sharp";
//#region src/browser-discovery.ts
async function executable(path) {
	try {
		await access(path, constants.X_OK);
		return true;
	} catch {
		return false;
	}
}
async function discoverBrowserExecutable(configured) {
	const checked = [];
	const candidates = [];
	if (configured !== void 0 && configured.trim().length > 0) candidates.push({
		path: configured,
		source: "configured"
	});
	try {
		candidates.push({
			path: chromium.executablePath(),
			source: "playwright"
		});
	} catch {}
	candidates.push(...browserSystemCandidates().map((path) => ({
		path,
		source: "system"
	})));
	for (const candidate of candidates) {
		if (checked.includes(candidate.path)) continue;
		checked.push(candidate.path);
		if (await executable(candidate.path)) return {
			executable: candidate.path,
			source: candidate.source,
			checked
		};
	}
	return { checked };
}
//#endregion
//#region src/diagnostics.ts
async function pythonCheck(runtime) {
	const subprocess = runtime.options.context?.get("subprocess");
	if (subprocess === void 0) return {
		id: "python",
		status: "failed",
		message: "DSH subprocess service is unavailable"
	};
	try {
		const executable = await subprocess.resolveExecutable(runtime.options.pythonExecutable ?? (process.platform === "win32" ? "python" : "python3"));
		const result = await runCollected(subprocess, [
			executable,
			"-c",
			"import matplotlib, PIL, cv2; print(\"ok\")"
		], {
			cwd: process.cwd(),
			timeoutMs: 1e4,
			maxOutputBytes: 8192,
			env: {
				MPLBACKEND: "Agg",
				PYTHONNOUSERSITE: "1"
			}
		});
		if (result.exitCode !== 0) return {
			id: "python",
			status: "failed",
			message: "Python dependencies are missing",
			details: {
				executable,
				stderr: result.stderr
			}
		};
		return {
			id: "python",
			status: "ready",
			message: "Python analysis runtime is ready",
			details: { executable }
		};
	} catch (error) {
		return {
			id: "python",
			status: "failed",
			message: safeErrorMessage(error)
		};
	}
}
async function rendererCheck(runtime) {
	const subprocess = runtime.options.context?.get("subprocess");
	if (subprocess === void 0) return {
		id: "renderer",
		status: "not_available",
		message: "DSH subprocess service is unavailable"
	};
	const resolveFirst = async (candidates) => {
		for (const candidate of candidates) try {
			return await subprocess.resolveExecutable(candidate);
		} catch {}
	};
	const existingFirst = async (candidates) => {
		for (const candidate of candidates) try {
			await access(candidate);
			return candidate;
		} catch {}
	};
	const nativeBackends = [];
	if (process.platform === "darwin") {
		const [launcher, keynote, powerpoint, capture] = await Promise.all([
			resolveFirst(appleScriptCandidates()),
			existingFirst(keynoteCandidates()),
			existingFirst(powerPointCandidates()),
			resolveFirst(screenCaptureCandidates())
		]);
		if (launcher !== void 0 && keynote !== void 0) nativeBackends.push({
			backend: "keynote",
			launcher,
			application: keynote
		});
		if (launcher !== void 0 && powerpoint !== void 0 && capture !== void 0) nativeBackends.push({
			backend: "powerpoint",
			launcher,
			application: powerpoint,
			capture
		});
	} else if (process.platform === "win32") {
		const [launcher, powerpoint] = await Promise.all([resolveFirst(powerShellCandidates()), existingFirst(powerPointCandidates())]);
		if (launcher !== void 0) nativeBackends.push({
			backend: "powerpoint",
			launcher,
			application: powerpoint
		});
	}
	let office;
	for (const candidate of libreOfficeCandidates()) try {
		office = await subprocess.resolveExecutable(candidate);
		break;
	} catch {}
	let rasterizer;
	for (const candidate of pdfToPpmCandidates()) try {
		rasterizer = await subprocess.resolveExecutable(candidate);
		break;
	} catch {}
	if (nativeBackends.length > 0 || office !== void 0 && rasterizer !== void 0) return {
		id: "renderer",
		status: "ready",
		message: "PPTX image renderer is available",
		details: {
			native_backends: nativeBackends,
			libreoffice: office,
			rasterizer
		}
	};
	return {
		id: "renderer",
		status: "not_available",
		message: "No supported PPTX renderer was found; install Keynote or PowerPoint (macOS), PowerPoint (Windows), or LibreOffice with Poppler",
		details: {
			native_backends: nativeBackends,
			libreoffice: office,
			rasterizer
		}
	};
}
async function fontCheck(runtime) {
	try {
		const fonts = await discoverRegisteredFonts(runtime.options.fontDirs);
		const availability = summarizeFontAvailability(fonts);
		const english = availability.roles["latin-sans"].available || availability.roles["latin-serif"].available;
		const chinese = availability.roles["cjk-sans"].available || availability.roles["cjk-serif"].available;
		const summary = fonts.map((font) => ({
			name: font.name,
			file: font.file,
			sha256: font.sha256,
			weight: font.weight,
			glyphCount: font.glyphCount,
			supportsLatin: font.supportsLatin,
			supportsCjk: font.supportsCjk,
			layer: registeredFont(font.name)?.layer,
			platforms: registeredFont(font.name)?.platforms
		}));
		if (!english || !chinese) return {
			id: "fonts",
			status: "failed",
			message: "The plugin approved font registry lacks an installed Latin or CJK family",
			details: {
				scope_note: "This is the plugin approved registry subset, not the host-wide font inventory.",
				english,
				chinese,
				availability,
				fonts: summary
			}
		};
		return {
			id: "fonts",
			status: "ready",
			message: `${availability.availableFamilies} approved font families are installed; Latin and CJK roles are covered`,
			details: {
				scope_note: "This is the plugin approved registry subset, not the host-wide font inventory.",
				availability,
				fonts: summary
			}
		};
	} catch (error) {
		return {
			id: "fonts",
			status: "failed",
			message: safeErrorMessage(error)
		};
	}
}
async function visionCheck(runtime, agent) {
	const context = runtime.options.context;
	if (context?.get("attachments") === void 0) return {
		id: "attachments",
		status: "failed",
		message: "Durable attachment service is unavailable; read_image will not register"
	};
	if (agent === void 0 || agent.options.provider === void 0 || agent.options.model === void 0) return {
		id: "vision_model",
		status: "not_available",
		message: "No active agent route was supplied for image-capability diagnosis"
	};
	const llm = context.get("llm");
	if (llm === void 0) return {
		id: "vision_model",
		status: "failed",
		message: "LLM service is unavailable"
	};
	try {
		const info = await llm.resolveModelInfo(agent.options.provider, agent.options.model);
		return info.inputModalities?.includes("image") ? {
			id: "vision_model",
			status: "ready",
			message: "The active model accepts image input",
			details: {
				provider: info.provider,
				model: info.id
			}
		} : {
			id: "vision_model",
			status: "failed",
			message: "The active model does not declare image input",
			details: {
				provider: info.provider,
				model: info.id
			}
		};
	} catch (error) {
		return {
			id: "vision_model",
			status: "failed",
			message: safeErrorMessage(error)
		};
	}
}
async function diagnosePptRuntime(runtime, agent) {
	const checks = [];
	checks.push(isSupportedPlatform() ? {
		id: "platform",
		status: "ready",
		message: `Supported platform: ${process.platform}`
	} : {
		id: "platform",
		status: "failed",
		message: `Unsupported platform: ${process.platform}`
	});
	const surface = runtime.toolSurface;
	checks.push(surface !== void 0 && surface.missing.length === 0 && surface.unexpected.length === 0 ? {
		id: "tools",
		status: "ready",
		message: `All ${PPT_MODE_TOOL_NAMES.length} PPT mode tools are visible`
	} : {
		id: "tools",
		status: "failed",
		message: "PPT mode tool surface is incomplete",
		details: surface ?? { missing: ["surface not mounted"] }
	});
	const browser = await discoverBrowserExecutable(runtime.options.browserExecutable);
	checks.push(browser.executable === void 0 ? {
		id: "browser",
		status: "failed",
		message: "No compatible Chromium/Chrome executable was found",
		details: { checked: browser.checked }
	} : {
		id: "browser",
		status: "ready",
		message: "Browser runtime is ready",
		details: {
			executable: browser.executable,
			source: browser.source
		}
	});
	checks.push(await pythonCheck(runtime));
	checks.push(await fontCheck(runtime));
	checks.push(await rendererCheck(runtime));
	const attachmentReady = runtime.options.context?.get("attachments") !== void 0;
	checks.push(attachmentReady ? {
		id: "attachments",
		status: "ready",
		message: "Durable attachment service is available"
	} : {
		id: "attachments",
		status: "failed",
		message: "Durable attachment service is unavailable; read_image will not register"
	});
	if (attachmentReady) checks.push(await visionCheck(runtime, agent));
	return {
		status: checks.some((check) => check.status === "failed") ? "failed" : checks.some((check) => check.status !== "ready") ? "degraded" : "ready",
		checks
	};
}
//#endregion
//#region src/session-resources.ts
function ownerKey(owner) {
	if (owner.agentId.trim().length === 0 || owner.sessionId.trim().length === 0) throw new Error("agentId and sessionId are required");
	return `${owner.agentId}\0${owner.sessionId}`;
}
var SessionResourceRegistry = class {
	sessions = /* @__PURE__ */ new Map();
	disposed = false;
	open(owner, workspace) {
		if (this.disposed) throw new Error("session resource registry is disposed");
		const key = ownerKey(owner);
		const current = this.sessions.get(key);
		if (current !== void 0) {
			if (current.workspace !== workspace) throw new Error("session workspace cannot change while resources are active");
			return this.snapshot(current);
		}
		const state = {
			owner: { ...owner },
			workspace,
			temporaryPaths: /* @__PURE__ */ new Set(),
			resources: /* @__PURE__ */ new Set()
		};
		this.sessions.set(key, state);
		return this.snapshot(state);
	}
	track(owner, resource) {
		const state = this.require(owner);
		state.resources.add(resource);
		return () => state.resources.delete(resource);
	}
	trackTemporaryPath(owner, path) {
		const state = this.require(owner);
		state.temporaryPaths.add(path);
		return () => state.temporaryPaths.delete(path);
	}
	state(owner) {
		const state = this.sessions.get(ownerKey(owner));
		return state === void 0 ? void 0 : this.snapshot(state);
	}
	async release(owner) {
		const key = ownerKey(owner);
		const state = this.sessions.get(key);
		if (state === void 0) return;
		this.sessions.delete(key);
		const errors = [];
		for (const resource of [...state.resources].reverse()) try {
			await resource.dispose();
		} catch (error) {
			errors.push(error);
		}
		for (const path of state.temporaryPaths) try {
			await rm(path, {
				recursive: true,
				force: true
			});
		} catch (error) {
			errors.push(error);
		}
		if (errors.length > 0) throw new AggregateError(errors, `failed to release ${errors.length} PPT session resources`);
	}
	async dispose() {
		if (this.disposed) return;
		this.disposed = true;
		const owners = [...this.sessions.values()].map((state) => state.owner);
		const errors = (await Promise.allSettled(owners.map((owner) => this.release(owner)))).filter((result) => result.status === "rejected").map((result) => result.reason);
		if (errors.length > 0) throw new AggregateError(errors, `failed to dispose ${errors.length} PPT sessions`);
	}
	require(owner) {
		const state = this.sessions.get(ownerKey(owner));
		if (state === void 0) throw new Error("session resources must be opened before tracking");
		return state;
	}
	snapshot(state) {
		return Object.freeze({
			owner: Object.freeze({ ...state.owner }),
			workspace: state.workspace,
			temporaryPaths: new Set(state.temporaryPaths),
			resourceCount: state.resources.size
		});
	}
};
//#endregion
//#region src/browser.ts
function keyOf(owner) {
	return `${owner.agentId}\0${owner.sessionId}`;
}
const SLIDE_SIZE = {
	width: 1280,
	height: 720
};
/** Only a leaf covering the slide on all four sides counts as a full-bleed layer. */
const FULL_BLEED_TOLERANCE_PX = 1;
function coversWholeSlide(box, slide) {
	return box.left <= FULL_BLEED_TOLERANCE_PX && box.top <= FULL_BLEED_TOLERANCE_PX && box.right >= slide.width - FULL_BLEED_TOLERANCE_PX && box.bottom >= slide.height - FULL_BLEED_TOLERANCE_PX;
}
/**
* 4x4 occupancy silhouette of one slide, used to detect adjacent pages with the same layout.
* Full-bleed layers are skipped: a page background leaf is identical on every page and would
* light all sixteen cells everywhere, so every adjacent pair would look repeated.
*/
function occupancySilhouette(leafBoxes, slide = SLIDE_SIZE) {
	const occupancy = Array.from({ length: 16 }, () => 0);
	const cellWidth = slide.width / 4;
	const cellHeight = slide.height / 4;
	for (const box of leafBoxes) {
		if (coversWholeSlide(box, slide)) continue;
		for (let row = 0; row < 4; row += 1) for (let column = 0; column < 4; column += 1) {
			const left = column * cellWidth;
			const top = row * cellHeight;
			if (box.right > left && box.left < left + cellWidth && box.bottom > top && box.top < top + cellHeight) occupancy[row * 4 + column] = 1;
		}
	}
	return occupancy;
}
var BrowserRuntime = class {
	resources;
	configuredExecutable;
	outputRoot;
	browser;
	launching;
	states = /* @__PURE__ */ new Map();
	constructor(resources, configuredExecutable, outputRoot = "ppt-output") {
		this.resources = resources;
		this.configuredExecutable = configuredExecutable;
		this.outputRoot = outputRoot;
	}
	async visit(owner, workspace, input, signal) {
		await this.abortOwnerIfRequested(owner, signal);
		const state = await this.requireState(owner, workspace, signal);
		const url = await this.resolveVisitUrl(workspace, input);
		const response = await this.cancellable(owner, state.page.goto(url.href, {
			waitUntil: "domcontentloaded",
			timeout: 3e4
		}), signal);
		throwIfAborted(signal);
		if (response !== null) {
			let redirects = 0;
			let request = response.request().redirectedFrom();
			while (request !== null) {
				redirects += 1;
				request = request.redirectedFrom();
			}
			if (redirects > DEFAULT_LIMITS.maxRedirects) throw new PptError("BROWSER_LIMIT_EXCEEDED", "redirect limit exceeded");
			const length = Number(response.headers()["content-length"] ?? 0);
			if (Number.isFinite(length) && length > DEFAULT_LIMITS.maxResponseBytes) throw new PptError("BROWSER_LIMIT_EXCEEDED", `response exceeds ${DEFAULT_LIMITS.maxResponseBytes} bytes`);
		}
		await this.validateCurrentUrl(state);
		await this.bumpVersion(state);
		return this.pageResult(state);
	}
	async find(owner, query, signal) {
		await this.abortOwnerIfRequested(owner, signal);
		const state = this.requireExisting(owner);
		if (query.trim().length === 0 || query.length > 200) throw new PptError("BROWSER_LIMIT_EXCEEDED", "find query must contain 1..200 characters");
		await this.refreshMutationVersion(state);
		const refs = await state.page.evaluate(({ query, version, start }) => {
			const normalized = query.toLocaleLowerCase();
			const candidates = [...document.querySelectorAll("a,button,[role=\"button\"],[role=\"link\"],summary,h1,h2,h3,p,li")];
			const visible = (element) => {
				const style = getComputedStyle(element);
				const box = element.getBoundingClientRect();
				return style.visibility !== "hidden" && style.display !== "none" && box.width > 0 && box.height > 0;
			};
			const output = [];
			let sequence = start;
			for (const element of candidates) {
				const text = (element.textContent ?? "").replace(/\s+/g, " ").trim();
				if (!visible(element) || !text.toLocaleLowerCase().includes(normalized)) continue;
				const ref = `v${version}-e${sequence++}`;
				element.setAttribute("data-dsh-ppt-ref", ref);
				const anchor = element instanceof HTMLAnchorElement ? element : void 0;
				output.push({
					ref,
					tag: element.tagName.toLocaleLowerCase(),
					text: text.slice(0, 240),
					...anchor?.href === void 0 ? {} : { href: anchor.href },
					clickable: anchor !== void 0 || element instanceof HTMLButtonElement || element.getAttribute("role") === "button" || element.getAttribute("role") === "link" || element.tagName === "SUMMARY"
				});
				if (output.length >= 20) break;
			}
			return {
				output,
				next: sequence,
				mutation: Number(window.__dshPptMutationVersion ?? 0)
			};
		}, {
			query: query.trim(),
			version: state.version,
			start: state.refSequence
		});
		state.refSequence = refs.next;
		state.observedMutation = refs.mutation;
		state.refs.clear();
		for (const item of refs.output) state.refs.set(item.ref, {
			version: state.version,
			selector: `[data-dsh-ppt-ref="${item.ref}"]`,
			clickable: item.clickable
		});
		return this.pageResult(state, refs.output);
	}
	async click(owner, ref, signal) {
		await this.abortOwnerIfRequested(owner, signal);
		const state = this.requireExisting(owner);
		await this.refreshMutationVersion(state);
		const target = state.refs.get(ref);
		if (target === void 0 || target.version !== state.version) throw new PptError("BROWSER_REF_STALE", `element reference is stale: ${ref}`);
		if (!target.clickable) throw new PptError("BROWSER_REF_STALE", `element is not clickable: ${ref}`);
		const locator = state.page.locator(target.selector);
		if (await locator.count() !== 1 || !await locator.isVisible()) throw new PptError("BROWSER_REF_STALE", `element reference no longer resolves: ${ref}`);
		await this.cancellable(owner, locator.click({ timeout: 1e4 }), signal);
		await state.page.waitForTimeout(250);
		throwIfAborted(signal);
		await this.validateCurrentUrl(state);
		await this.bumpVersion(state);
		return this.pageResult(state);
	}
	async scroll(owner, direction, amount = 640, signal) {
		await this.abortOwnerIfRequested(owner, signal);
		const state = this.requireExisting(owner);
		if (!Number.isInteger(amount) || amount < 100 || amount > 2e3) throw new PptError("BROWSER_LIMIT_EXCEEDED", "scroll amount must be an integer between 100 and 2000");
		state.scrollCount += 1;
		if (state.scrollCount > 100) throw new PptError("BROWSER_LIMIT_EXCEEDED", "scroll call limit exceeded");
		await state.page.evaluate((delta) => window.scrollBy({
			top: delta,
			behavior: "instant"
		}), direction === "down" ? amount : -amount);
		await state.page.waitForTimeout(150);
		await this.bumpVersion(state);
		return this.pageResult(state);
	}
	async renderHtmlPreview(owner, workspace, htmlPath, previewDirectory, pageCount, allowedFonts, signal) {
		await this.visit(owner, workspace, htmlPath, signal);
		const state = this.requireExisting(owner);
		const preview = await resolveWorkspacePath(workspace, previewDirectory);
		await mkdir(preview, { recursive: true });
		await state.page.addStyleTag({ content: "*,*::before,*::after{animation:none!important;transition:none!important;caret-color:transparent!important}" });
		await state.page.evaluate(async () => {
			await document.fonts.ready;
		});
		const inspection = await state.page.evaluate(({ count, fonts }) => {
			const slides = [...document.querySelectorAll(".ppt-slide[data-page]")];
			const errors = [];
			const warnings = [];
			const usedFonts = /* @__PURE__ */ new Set();
			if (slides.length !== count) errors.push(`expected ${count} slides but found ${slides.length}`);
			const designPages = [];
			slides.forEach((slide, slideIndex) => {
				const slideBox = slide.getBoundingClientRect();
				const leafBoxes = [];
				if (Math.abs(slideBox.width - 1280) > .5 || Math.abs(slideBox.height - 720) > .5) errors.push(`page ${slideIndex + 1}: slide box is ${slideBox.width}x${slideBox.height}, expected 1280x720`);
				for (const leaf of slide.querySelectorAll("[data-ppt-id][data-ppt-kind]")) {
					const box = leaf.getBoundingClientRect();
					leafBoxes.push({
						left: box.left - slideBox.left,
						top: box.top - slideBox.top,
						right: box.right - slideBox.left,
						bottom: box.bottom - slideBox.top
					});
					if (box.left < slideBox.left - 1 || box.top < slideBox.top - 1 || box.right > slideBox.right + 1 || box.bottom > slideBox.bottom + 1) errors.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} exceeds slide bounds`);
					const family = getComputedStyle(leaf).fontFamily.split(",")[0].trim().replace(/^['"]|['"]$/g, "");
					if (family.length > 0) {
						usedFonts.add(family);
						if (!fonts.includes(family)) errors.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} uses unauthorized font ${family}`);
					}
					if (leaf.dataset.pptKind === "text" && (leaf.scrollWidth > leaf.clientWidth + 1 || leaf.scrollHeight > leaf.clientHeight + 1)) errors.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} has text overflow`);
					if (box.width < 1 || box.height < 1) warnings.push(`page ${slideIndex + 1}: ${leaf.dataset.pptId} has near-zero bounds`);
				}
				const anchor = slide.querySelector("[data-art-role=\"visual-anchor\"]")?.getBoundingClientRect();
				designPages.push({
					page: slideIndex + 1,
					...anchor === void 0 ? {} : { anchorAreaRatio: Math.max(0, anchor.width * anchor.height) / Math.max(1, slideBox.width * slideBox.height) },
					frameCount: slide.querySelectorAll("[data-art-role=\"frame\"]").length,
					leafBoxes,
					roleStyles: [...slide.querySelectorAll("[data-art-role]")].map((element) => {
						const style = getComputedStyle(element);
						return {
							role: element.dataset.artRole ?? "",
							fontFamily: style.fontFamily.split(",")[0].trim().replace(/^['"]|['"]$/g, ""),
							fontWeight: Number.parseInt(style.fontWeight, 10) || 400
						};
					})
				});
			});
			return {
				errors,
				warnings,
				fonts: [...usedFonts].sort(),
				designPages
			};
		}, {
			count: pageCount,
			fonts: [...allowedFonts]
		});
		if (inspection.errors.length > 0) throw new PptError("HTML_CREATE_VALIDATION_FAILED", "HTML browser validation failed", { details: { issues: inspection.errors } });
		const designPages = inspection.designPages.map((page) => ({
			page: page.page,
			...page.anchorAreaRatio === void 0 ? {} : { anchorAreaRatio: page.anchorAreaRatio },
			frameCount: page.frameCount,
			occupancy: occupancySilhouette(page.leafBoxes),
			roleStyles: page.roleStyles
		}));
		const previews = [];
		for (let index = 0; index < pageCount; index += 1) {
			throwIfAborted(signal);
			const target = join(preview, `page-${String(index + 1).padStart(3, "0")}.png`);
			await this.cancellable(owner, state.page.locator(".ppt-slide[data-page]").nth(index).screenshot({
				path: target,
				type: "png",
				animations: "disabled"
			}), signal);
			previews.push(workspaceRelative(workspace, target));
		}
		return {
			previews,
			fonts: inspection.fonts,
			warnings: inspection.warnings,
			designPages
		};
	}
	async extractDeckIr(owner, workspace, htmlPath, pageCount, signal) {
		await this.visit(owner, workspace, htmlPath, signal);
		return {
			widthPx: 1280,
			heightPx: 720,
			slides: await this.requireExisting(owner).page.evaluate((expected) => {
				const px = (value, fallback = 0) => {
					const parsed = Number.parseFloat(value);
					return Number.isFinite(parsed) ? parsed : fallback;
				};
				const textRuns = (element) => {
					const runs = [];
					const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
					for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
						const value = node.textContent ?? "";
						if (value.length === 0) continue;
						const parent = node.parentElement ?? element;
						const style = getComputedStyle(parent);
						runs.push({
							text: value,
							fontFamily: style.fontFamily.split(",")[0].trim().replace(/^['"]|['"]$/g, ""),
							fontSizePx: px(style.fontSize),
							fontWeight: Number.parseInt(style.fontWeight, 10) || 400,
							fontStyle: style.fontStyle,
							color: style.color,
							textDecoration: style.textDecorationLine
						});
					}
					return runs;
				};
				const pages = [...document.querySelectorAll(".ppt-slide[data-page]")];
				if (pages.length !== expected) throw new Error(`expected ${expected} slides, found ${pages.length}`);
				return pages.map((slide, slideIndex) => {
					const slideBox = slide.getBoundingClientRect();
					const elements = [...slide.querySelectorAll("[data-ppt-id][data-ppt-kind]")].map((element, domOrder) => {
						const box = element.getBoundingClientRect();
						const style = getComputedStyle(element);
						const kind = element.dataset.pptKind;
						const item = {
							id: element.dataset.pptId,
							kind,
							z: Number.parseInt(element.dataset.pptZ ?? "0", 10),
							domOrder,
							box: {
								x: box.left - slideBox.left,
								y: box.top - slideBox.top,
								w: box.width,
								h: box.height
							},
							style: {
								fontFamily: style.fontFamily.split(",")[0].trim().replace(/^['"]|['"]$/g, ""),
								fontSizePx: px(style.fontSize),
								fontWeight: Number.parseInt(style.fontWeight, 10) || 400,
								fontStyle: style.fontStyle,
								color: style.color,
								backgroundColor: style.backgroundColor,
								backgroundImage: style.backgroundImage,
								borderColor: style.borderTopColor,
								borderWidthPx: px(style.borderTopWidth),
								borderStyle: style.borderTopStyle,
								borderRadius: style.borderRadius,
								textAlign: style.textAlign,
								verticalAlign: style.verticalAlign,
								lineHeightPx: style.lineHeight === "normal" ? px(style.fontSize) * 1.2 : px(style.lineHeight),
								opacity: px(style.opacity, 1),
								objectFit: style.objectFit,
								objectPosition: style.objectPosition
							}
						};
						if (kind === "text") {
							item.text = element.innerText;
							item.runs = textRuns(element);
							if (style.backgroundImage !== "none") item.unsupportedReason = "gradient or background image on text";
						} else if (kind === "image") {
							const image = element;
							item.imagePath = image.currentSrc || image.src;
							if (!image.complete || image.naturalWidth === 0 || image.naturalHeight === 0) item.unsupportedReason = "image failed to load";
						} else if (kind === "svg") {
							item.svg = element.outerHTML;
							if (element.querySelector("filter,mask,pattern,foreignObject,script") !== null) item.unsupportedReason = "complex SVG feature";
						} else if (kind === "table") {
							const table = element;
							item.table = [...table.rows].map((row) => [...row.cells].map((cell) => cell.innerText));
							if ([...table.rows].some((row) => [...row.cells].some((cell) => cell.rowSpan !== 1 || cell.colSpan !== 1))) item.unsupportedReason = "merged table cells";
						} else if (kind === "shape" && style.backgroundImage !== "none") item.unsupportedReason = "gradient or background image on shape";
						return item;
					}).sort((a, b) => a.z - b.z || a.domOrder - b.domOrder);
					return {
						page: slideIndex + 1,
						elements,
						speakerNotes: []
					};
				});
			}, pageCount)
		};
	}
	async rasterizeElement(owner, workspace, htmlPath, elementId, targetInput, signal) {
		await this.visit(owner, workspace, htmlPath, signal);
		const state = this.requireExisting(owner);
		const target = await resolveWorkspacePath(workspace, targetInput, { createParent: true });
		const locator = state.page.locator(`[data-ppt-id="${elementId}"]`);
		if (await locator.count() !== 1) throw new PptError("PPT_CREATE_UNSUPPORTED_ELEMENT", `cannot rasterize missing or duplicate element: ${elementId}`);
		await this.cancellable(owner, locator.screenshot({
			path: target,
			type: "png",
			animations: "disabled"
		}), signal);
		return workspaceRelative(workspace, target);
	}
	async dispose() {
		const states = [...this.states.values()];
		this.states.clear();
		await Promise.allSettled(states.map((state) => state.context.close()));
		const browser = this.browser;
		this.browser = void 0;
		if (browser !== void 0) await browser.close();
	}
	async launch(signal) {
		if (this.browser !== void 0) return this.browser;
		if (this.launching !== void 0) return this.launching;
		this.launching = (async () => {
			const found = await discoverBrowserExecutable(this.configuredExecutable);
			if (found.executable === void 0) throw new PptError("BROWSER_NOT_READY", "no compatible Chromium/Chrome executable found");
			throwIfAborted(signal);
			const browser = await chromium.launch({
				executablePath: found.executable,
				headless: true,
				args: ["--disable-dev-shm-usage"]
			});
			this.browser = browser;
			return browser;
		})().finally(() => {
			this.launching = void 0;
		});
		return this.launching;
	}
	async requireState(owner, workspace, signal) {
		const current = this.states.get(keyOf(owner));
		if (current !== void 0) {
			if (current.workspace !== workspace) throw new PptError("BROWSER_URL_BLOCKED", "browser workspace cannot change within a session");
			return current;
		}
		this.resources.open(owner, workspace);
		const context = await (await this.launch(signal)).newContext({
			acceptDownloads: false,
			serviceWorkers: "block",
			viewport: {
				width: 1280,
				height: 720
			}
		});
		await context.addInitScript(() => {
			const target = window;
			target.__dshPptMutationVersion = 0;
			new MutationObserver((records) => {
				if (records.some((record) => record.type !== "attributes" || record.attributeName !== "data-dsh-ppt-ref")) target.__dshPptMutationVersion = (target.__dshPptMutationVersion ?? 0) + 1;
			}).observe(document, {
				childList: true,
				subtree: true,
				attributes: true,
				characterData: true
			});
		});
		await context.route("**/*", async (route) => {
			const requestUrl = new URL(route.request().url());
			try {
				if (requestUrl.protocol === "file:") {
					const path = fileURLToPath(requestUrl);
					const outputRoot = await resolveWorkspacePath(workspace, this.outputRoot);
					if (!isPathInside(outputRoot, path)) throw new PptError("BROWSER_URL_BLOCKED", "local browser path is outside plugin output");
				} else await validatePublicHttpUrl(requestUrl.href);
				await route.continue();
			} catch {
				await route.abort("blockedbyclient");
			}
		});
		const page = await context.newPage();
		page.on("dialog", (dialog) => {
			dialog.dismiss();
		});
		page.on("download", (download) => {
			download.cancel();
		});
		page.on("popup", (popup) => {
			popup.close();
		});
		const state = {
			owner: { ...owner },
			workspace,
			context,
			page,
			version: 0,
			observedMutation: 0,
			refs: /* @__PURE__ */ new Map(),
			refSequence: 1,
			scrollCount: 0
		};
		this.states.set(keyOf(owner), state);
		this.resources.track(owner, {
			label: "browser-context",
			dispose: async () => {
				this.states.delete(keyOf(owner));
				await context.close().catch(() => void 0);
			}
		});
		return state;
	}
	requireExisting(owner) {
		const state = this.states.get(keyOf(owner));
		if (state === void 0) throw new PptError("BROWSER_NOT_READY", "call browser_visit before using this browser tool");
		return state;
	}
	async resolveVisitUrl(workspace, input) {
		if (isLocalFilesystemPath(input)) return this.resolveLocalFileUrl(workspace, input);
		try {
			const parsed = new URL(input);
			if (parsed.protocol === "file:") return await this.resolveLocalFileUrl(workspace, fileURLToPath(parsed));
			return await validatePublicHttpUrl(parsed.href);
		} catch (error) {
			if (error instanceof PptError) throw error;
			return this.resolveLocalFileUrl(workspace, input);
		}
	}
	async resolveLocalFileUrl(workspace, input) {
		const path = await resolveWorkspacePath(workspace, input, {
			mustExist: true,
			kind: "file"
		});
		const outputRoot = await resolveWorkspacePath(workspace, this.outputRoot);
		if (!isPathInside(outputRoot, path)) throw new PptError("BROWSER_URL_BLOCKED", "local HTML is outside plugin output");
		return pathToFileURL(path);
	}
	async validateCurrentUrl(state) {
		const current = new URL(state.page.url());
		if (current.protocol === "file:") {
			const outputRoot = await resolveWorkspacePath(state.workspace, this.outputRoot);
			if (!isPathInside(outputRoot, fileURLToPath(current))) throw new PptError("BROWSER_URL_BLOCKED", "navigation left plugin output");
		} else await validatePublicHttpUrl(current.href);
	}
	async bumpVersion(state) {
		state.version += 1;
		state.refs.clear();
		state.observedMutation = await state.page.evaluate(() => Number(window.__dshPptMutationVersion ?? 0));
	}
	async refreshMutationVersion(state) {
		const mutation = await state.page.evaluate(() => Number(window.__dshPptMutationVersion ?? 0));
		if (mutation !== state.observedMutation) {
			state.version += 1;
			state.refs.clear();
			state.observedMutation = mutation;
		}
	}
	async pageResult(state, elements) {
		const [title, text] = await Promise.all([state.page.title(), state.page.locator("body").innerText({ timeout: 5e3 }).catch(() => "")]);
		return {
			url: state.page.url(),
			title: title.slice(0, 500),
			text: text.replace(/\u0000/g, "").slice(0, DEFAULT_LIMITS.maxBrowserTextChars),
			page_version: state.version,
			content_is_untrusted: true,
			...elements === void 0 ? {} : { elements }
		};
	}
	async cancellable(owner, operation, signal) {
		throwIfAborted(signal);
		if (signal === void 0) return operation;
		return new Promise((resolve, reject) => {
			let settled = false;
			const finish = (callback) => {
				if (settled) return;
				settled = true;
				signal.removeEventListener("abort", abort);
				callback();
			};
			const abort = () => {
				if (settled) return;
				settled = true;
				signal.removeEventListener("abort", abort);
				this.resources.release(owner).then(() => reject(new PptError("PPT_ABORTED", "browser operation aborted and its isolated page was released")), (error) => reject(new PptError("PPT_ABORTED", "browser operation aborted but page cleanup failed", { cause: error })));
			};
			signal.addEventListener("abort", abort, { once: true });
			operation.then((value) => finish(() => resolve(value)), (error) => finish(() => reject(error)));
		});
	}
	async abortOwnerIfRequested(owner, signal) {
		if (!signal?.aborted) return;
		await this.resources.release(owner);
		throwIfAborted(signal);
	}
};
//#endregion
//#region src/python.ts
const MIME_BY_EXTENSION = Object.freeze({
	".png": "image/png",
	".jpg": "image/jpeg",
	".jpeg": "image/jpeg",
	".webp": "image/webp",
	".gif": "image/gif",
	".svg": "image/svg+xml",
	".csv": "text/csv",
	".json": "application/json",
	".txt": "text/plain",
	".pdf": "application/pdf"
});
async function snapshotFiles(root, signal) {
	const files = /* @__PURE__ */ new Map();
	const queue = [root];
	let seen = 0;
	while (queue.length > 0) {
		throwIfAborted(signal);
		const directory = queue.pop();
		for (const entry of await readdir(directory, { withFileTypes: true })) {
			if (entry.name === ".git" || entry.isSymbolicLink()) continue;
			const path = join(directory, entry.name);
			if (entry.isDirectory()) queue.push(path);
			else if (entry.isFile()) {
				seen += 1;
				if (seen > 1e4) throw new PptError("PPT_RESOURCE_LIMIT", "Python workspace scan exceeds 10000 files");
				const info = await stat(path);
				files.set(path, {
					size: info.size,
					mtimeMs: info.mtimeMs
				});
			}
		}
	}
	return files;
}
async function artifactMetadata(workspace, path) {
	const info = await stat(path);
	if (info.size > DEFAULT_LIMITS.maxGeneratedFileBytes) throw new PptError("PPT_RESOURCE_LIMIT", `${workspaceRelative(workspace, path)} exceeds the generated-file size limit`);
	const mime = MIME_BY_EXTENSION[extname(path).toLowerCase()] ?? "application/octet-stream";
	const base = {
		path: workspaceRelative(workspace, path),
		size: info.size,
		mime_type: mime
	};
	if (!mime.startsWith("image/")) return base;
	try {
		const metadata = await sharp(path, { limitInputPixels: DEFAULT_LIMITS.maxImagePixels }).metadata();
		return {
			...base,
			...metadata.width === void 0 ? {} : { width: metadata.width },
			...metadata.height === void 0 ? {} : { height: metadata.height }
		};
	} catch (error) {
		throw new PptError("PYTHON_EXECUTION_FAILED", `generated image is invalid: ${base.path}`, { cause: error });
	}
}
var PythonRuntime = class {
	subprocess;
	sandbox;
	resources;
	executable;
	constructor(subprocess, sandbox, resources, executable = process.platform === "win32" ? "python" : "python3") {
		this.subprocess = subprocess;
		this.sandbox = sandbox;
		this.resources = resources;
		this.executable = executable;
	}
	async execute(owner, workspace, input, signal) {
		throwIfAborted(signal);
		if (this.subprocess === void 0 || this.sandbox === void 0) throw new PptError("PPT_CAPABILITY_UNAVAILABLE", "Python requires DSH subprocess and sandbox services");
		if (input.code.trim().length === 0 || [...input.code].length > 2e5) throw new PptError("PYTHON_EXECUTION_FAILED", "Python code must contain 1..200000 Unicode code points");
		const cwd = await resolveWorkspacePath(workspace, input.cwd ?? ".", {
			mustExist: true,
			kind: "directory"
		});
		const expected = await Promise.all((input.expected_outputs ?? []).map((path) => resolveWorkspacePath(workspace, path)));
		if (expected.length > DEFAULT_LIMITS.maxGeneratedFiles) throw new PptError("PPT_RESOURCE_LIMIT", "too many expected Python outputs");
		const timeoutMs = boundedInteger(input.timeout_ms ?? DEFAULT_LIMITS.maxPythonMs, "timeout_ms", 1e3, DEFAULT_LIMITS.maxPythonMs);
		this.resources.open(owner, workspace);
		const resolvedExecutable = await this.subprocess.resolveExecutable(this.executable, void 0, signal);
		const environment = {
			MPLBACKEND: "Agg",
			PYTHONNOUSERSITE: "1",
			PYTHONDONTWRITEBYTECODE: "1"
		};
		const preflight = await this.sandbox.confine([
			resolvedExecutable,
			"-c",
			"import matplotlib, PIL, cv2; matplotlib.use(\"Agg\"); print(\"ready\")"
		], {
			mode: "read-only",
			workspaceRoot: workspace
		});
		if (preflight.enforcement !== "full") throw new PptError("PPT_CAPABILITY_UNAVAILABLE", "Python sandbox enforcement is partial");
		let dependency;
		try {
			dependency = await runCollected(this.subprocess, preflight.argv, {
				cwd,
				env: environment,
				signal,
				timeoutMs: Math.min(timeoutMs, 15e3),
				maxOutputBytes: 8192
			});
		} catch (error) {
			if (signal?.aborted) throwIfAborted(signal);
			throw asPptError(error, "PYTHON_DEPENDENCY_MISSING", "Python dependency preflight failed");
		}
		if (dependency.exitCode !== 0) throw new PptError("PYTHON_DEPENDENCY_MISSING", "Python requires matplotlib, Pillow, and opencv-python", { details: {
			stderr: dependency.stderr.slice(0, 2e3),
			executable: resolvedExecutable
		} });
		const before = await snapshotFiles(cwd, signal);
		const source = [
			"import os",
			"os.environ[\"MPLBACKEND\"] = \"Agg\"",
			"import matplotlib",
			"matplotlib.use(\"Agg\")",
			"import PIL, cv2",
			input.code
		].join("\n");
		const confined = await this.sandbox.confine([resolvedExecutable, "-"], {
			mode: "workspace-write",
			workspaceRoot: workspace
		});
		if (confined.enforcement !== "full") throw new PptError("PPT_CAPABILITY_UNAVAILABLE", "Python sandbox enforcement is partial");
		const started = Date.now();
		let result;
		try {
			result = await runCollected(this.subprocess, confined.argv, {
				cwd,
				env: environment,
				signal,
				stdin: source,
				timeoutMs,
				maxOutputBytes: DEFAULT_LIMITS.maxPythonOutputChars * 4
			});
		} catch (error) {
			if (signal?.aborted) throwIfAborted(signal);
			throw asPptError(error, "PYTHON_EXECUTION_FAILED", "Python execution failed");
		}
		const after = await snapshotFiles(cwd, signal);
		const changed = [...after.entries()].filter(([path, stamp]) => {
			const old = before.get(path);
			return old === void 0 || old.size !== stamp.size || old.mtimeMs !== stamp.mtimeMs;
		}).map(([path]) => path);
		if (changed.length > DEFAULT_LIMITS.maxGeneratedFiles) throw new PptError("PPT_RESOURCE_LIMIT", `Python generated more than ${DEFAULT_LIMITS.maxGeneratedFiles} files`);
		for (const path of expected) if (!after.has(path)) throw new PptError("PYTHON_EXECUTION_FAILED", `expected output was not created: ${workspaceRelative(workspace, path)}`);
		const artifacts = await Promise.all(changed.sort().map((path) => artifactMetadata(workspace, path)));
		const exitCode = result.exitCode ?? -1;
		if (exitCode !== 0) throw new PptError("PYTHON_EXECUTION_FAILED", `Python exited with code ${exitCode}`, { details: {
			exit_code: exitCode,
			stdout: result.stdout,
			stderr: result.stderr
		} });
		return {
			exit_code: exitCode,
			stdout: result.stdout.slice(0, DEFAULT_LIMITS.maxPythonOutputChars),
			stderr: result.stderr.slice(0, DEFAULT_LIMITS.maxPythonOutputChars),
			stdout_truncated: result.stdoutTruncated || result.stdout.length > DEFAULT_LIMITS.maxPythonOutputChars,
			stderr_truncated: result.stderrTruncated || result.stderr.length > DEFAULT_LIMITS.maxPythonOutputChars,
			duration_ms: Date.now() - started,
			artifacts
		};
	}
};
//#endregion
//#region src/runtime.ts
var DefaultPptRuntime = class {
	options;
	toolSurface;
	disposed = false;
	resources = new SessionResourceRegistry();
	browser;
	python;
	imageSearch = new ImageSearchRuntime();
	pptImage;
	quality;
	constructor(options) {
		this.options = Object.freeze({
			...options,
			fontDirs: Object.freeze([...options.fontDirs ?? []])
		});
		this.browser = new BrowserRuntime(this.resources, options.browserExecutable, options.outputRoot);
		this.python = new PythonRuntime(options.context?.get("subprocess"), options.context?.get("sandbox"), this.resources, options.pythonExecutable);
		this.pptImage = new PptImageRuntime(options.context?.get("subprocess"), options.context?.get("sandbox"), this.resources, {}, options.fontDirs);
		this.quality = new QualityRuntime(options.context?.get("subprocess"), options.context?.get("sandbox"), this.resources, {}, options.fontDirs, this.pptImage);
	}
	recordToolSurface(status) {
		if (this.disposed) return;
		this.toolSurface = Object.freeze({
			visible: Object.freeze([...status.visible]),
			missing: Object.freeze([...status.missing]),
			unexpected: Object.freeze([...status.unexpected])
		});
	}
	async canReviewImages(agent) {
		const context = this.options.context;
		if (context?.get("attachments") === void 0 || agent?.options.provider === void 0 || agent.options.model === void 0) return false;
		const llm = context.get("llm");
		if (llm === void 0) return false;
		try {
			return (await llm.resolveModelInfo(agent.options.provider, agent.options.model)).inputModalities?.includes("image") ?? false;
		} catch {
			return false;
		}
	}
	diagnose(agent) {
		return diagnosePptRuntime(this, agent);
	}
	async dispose() {
		if (this.disposed) return;
		this.disposed = true;
		await this.browser.dispose();
		await this.resources.dispose();
	}
};
function createPptRuntime(options = {}) {
	return new DefaultPptRuntime(options);
}
//#endregion
export { createPptRuntime as t };

//# sourceMappingURL=runtime-Da9gUFsv.mjs.map