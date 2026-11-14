/**
 * The shell tool this host exposes. DSH's stock presets mount `tool-bash` on
 * darwin and linux and `tool-pwsh` on Windows, so the PPT preset follows the same
 * split and its audit expects whichever one is actually available.
 */
export const PPT_SHELL_TOOL_NAME = process.platform === 'win32' ? 'pwsh' : 'bash'

export const PPT_NATIVE_TOOL_NAMES = [
  PPT_SHELL_TOOL_NAME,
  'ask_user_question',
  'edit',
  'read',
  'read_image',
  'todo_write',
  'web_search',
  'write',
] as const

export const PPT_TOOL_NAMES = [
  'browser_click',
  'browser_find',
  'browser_scroll_down',
  'browser_scroll_up',
  'browser_visit',
  'html_create',
  'image_search',
  'ppt_create',
  'ppt_fonts',
  'ppt_image',
  'ppt_outline',
  'ppt_themes',
  'python',
] as const

export const PPT_MODE_TOOL_NAMES = [...PPT_NATIVE_TOOL_NAMES, ...PPT_TOOL_NAMES].sort()

export type PptToolName = typeof PPT_TOOL_NAMES[number]
