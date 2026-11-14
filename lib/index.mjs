import { n as PptError, t as PPT_ERROR_CODES } from "./errors-BdtilPdq.mjs";
import { H as isSupportedPlatform } from "./quality-BOrmVVNM.mjs";
import { t as createPptRuntime } from "./runtime-Da9gUFsv.mjs";
import { PPT_MODE_TOOL_NAMES, PPT_TOOL_NAMES } from "./schemas.mjs";
import z from "schemastery";
//#region src/index.ts
const name = "dsh-ppt";
const inject = ["sandbox", "subprocess"];
const Config = z.object({
	presetId: z.string().pattern(/^[a-z0-9][a-z0-9-]*$/).default("ppt"),
	pythonExecutable: z.string().default(process.platform === "win32" ? "python" : "python3"),
	browserExecutable: z.string().default(""),
	fontDirs: z.array(z.string()).default([]),
	outputRoot: z.string().default("ppt-output")
});
function assertSupportedPlatform(platform = process.platform) {
	if (!isSupportedPlatform(platform)) throw new PptError("PPT_PLATFORM_UNSUPPORTED", `PPT mode supports macOS, Linux, and Windows; unsupported platform: ${platform}`);
}
/**
* Provide the host-plane PPT runtime. The model-facing PPT preset is declared by
* this package's `cordis.patch.yml`, because the DSH agent-preset registry
* composes presets from ordinary plugin rows and no longer scans
* `<dshHome>/.agent-presets`.
*/
async function apply(ctx, config) {
	assertSupportedPlatform();
	const runtime = createPptRuntime({
		context: ctx,
		pythonExecutable: config.pythonExecutable,
		browserExecutable: config.browserExecutable || void 0,
		fontDirs: config.fontDirs,
		outputRoot: config.outputRoot
	});
	ctx.provide("pptRuntime", runtime);
	ctx.effect(() => () => runtime.dispose(), "dsh-ppt: dispose runtime resources");
}
//#endregion
export { Config, PPT_ERROR_CODES, PPT_MODE_TOOL_NAMES, PPT_TOOL_NAMES, PptError, apply, assertSupportedPlatform, createPptRuntime, inject, name };

//# sourceMappingURL=index.mjs.map