//#region src/schemas.ts
/**
* The shell tool this host exposes. DSH's stock presets mount `tool-bash` on
* darwin and linux and `tool-pwsh` on Windows, so the PPT preset follows the same
* split and its audit expects whichever one is actually available.
*/
const PPT_SHELL_TOOL_NAME = process.platform === "win32" ? "pwsh" : "bash";
const PPT_NATIVE_TOOL_NAMES = [
	PPT_SHELL_TOOL_NAME,
	"ask_user_question",
	"edit",
	"read",
	"read_image",
	"todo_write",
	"web_search",
	"write"
];
const PPT_TOOL_NAMES = [
	"browser_click",
	"browser_find",
	"browser_scroll_down",
	"browser_scroll_up",
	"browser_visit",
	"html_create",
	"image_search",
	"ppt_create",
	"ppt_fonts",
	"ppt_image",
	"ppt_outline",
	"ppt_themes",
	"python"
];
const PPT_MODE_TOOL_NAMES = [...PPT_NATIVE_TOOL_NAMES, ...PPT_TOOL_NAMES].sort();
//#endregion
export { PPT_MODE_TOOL_NAMES, PPT_NATIVE_TOOL_NAMES, PPT_SHELL_TOOL_NAME, PPT_TOOL_NAMES };

//# sourceMappingURL=schemas.mjs.map