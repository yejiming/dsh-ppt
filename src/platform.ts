import { homedir } from 'node:os'
import { posix, win32 } from 'node:path'

export function isSupportedPlatform(platform: NodeJS.Platform = process.platform): boolean {
  return platform === 'darwin' || platform === 'linux' || platform === 'win32'
}

export function systemFontDirectories(
  platform: NodeJS.Platform = process.platform,
  home: string = homedir(),
  env: NodeJS.ProcessEnv = process.env,
): string[] {
  if (platform === 'darwin') {
    return [
      '/System/Library/Fonts', '/System/Library/Fonts/Supplemental',
      '/System/Library/AssetsV2/com_apple_MobileAsset_Font7', '/System/Library/AssetsV2/com_apple_MobileAsset_Font8',
      '/Library/Fonts', posix.join(home, 'Library/Fonts'),
    ]
  }
  if (platform === 'win32') {
    const windows = env.SystemRoot ?? env.WINDIR ?? 'C:\\Windows'
    const local = env.LOCALAPPDATA
    return [win32.join(windows, 'Fonts'), ...(local === undefined ? [] : [win32.join(local, 'Microsoft', 'Windows', 'Fonts')])]
  }
  return ['/usr/share/fonts', '/usr/local/share/fonts', posix.join(home, '.local/share/fonts'), posix.join(home, '.fonts')]
}

export function libreOfficeCandidates(
  platform: NodeJS.Platform = process.platform,
  env: NodeJS.ProcessEnv = process.env,
): string[] {
  if (platform === 'darwin') {
    return ['soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice', '/Applications/LibreOfficeDev.app/Contents/MacOS/soffice']
  }
  if (platform === 'win32') {
    const roots = [env.ProgramFiles, env['ProgramFiles(x86)'], env.LOCALAPPDATA].filter((value): value is string => typeof value === 'string' && value.length > 0)
    return ['soffice.exe', 'soffice', ...roots.map(root => win32.join(root, 'LibreOffice', 'program', 'soffice.exe'))]
  }
  return ['soffice', 'libreoffice', '/usr/bin/soffice', '/usr/bin/libreoffice']
}

export function pdfToPpmCandidates(platform: NodeJS.Platform = process.platform): string[] {
  return platform === 'win32' ? ['pdftoppm.exe', 'pdftoppm'] : ['pdftoppm']
}

export function appleScriptCandidates(platform: NodeJS.Platform = process.platform): string[] {
  return platform === 'darwin' ? ['/usr/bin/osascript', 'osascript'] : []
}

export function screenCaptureCandidates(platform: NodeJS.Platform = process.platform): string[] {
  return platform === 'darwin' ? ['/usr/sbin/screencapture', 'screencapture'] : []
}

export function keynoteCandidates(
  platform: NodeJS.Platform = process.platform,
  home: string = homedir(),
): string[] {
  return platform === 'darwin'
    ? ['/Applications/Keynote.app', posix.join(home, 'Applications/Keynote.app')]
    : []
}

export function powerShellCandidates(platform: NodeJS.Platform = process.platform): string[] {
  return platform === 'win32'
    ? ['powershell.exe', 'powershell', 'pwsh.exe', 'pwsh']
    : []
}

export function powerPointCandidates(
  platform: NodeJS.Platform = process.platform,
  env: NodeJS.ProcessEnv = process.env,
  home: string = homedir(),
): string[] {
  if (platform === 'darwin') {
    return ['/Applications/Microsoft PowerPoint.app', posix.join(home, 'Applications/Microsoft PowerPoint.app')]
  }
  if (platform !== 'win32') return []
  const roots = [env.ProgramFiles, env['ProgramFiles(x86)'], env.LOCALAPPDATA]
    .filter((value): value is string => typeof value === 'string' && value.length > 0)
  return roots.flatMap(root => [
    win32.join(root, 'Microsoft Office', 'root', 'Office16', 'POWERPNT.EXE'),
    win32.join(root, 'Microsoft Office', 'Office16', 'POWERPNT.EXE'),
  ])
}

export type PptImageBackend = 'keynote' | 'powerpoint' | 'libreoffice'

export function pptImageBackendOrder(
  platform: NodeJS.Platform = process.platform,
  requested: 'auto' | PptImageBackend = 'auto',
): PptImageBackend[] {
  if (requested !== 'auto') return [requested]
  if (platform === 'darwin') return ['keynote', 'libreoffice', 'powerpoint']
  if (platform === 'win32') return ['powerpoint', 'libreoffice']
  return platform === 'linux' ? ['libreoffice'] : []
}

export function browserSystemCandidates(
  platform: NodeJS.Platform = process.platform,
  home: string = homedir(),
  env: NodeJS.ProcessEnv = process.env,
): string[] {
  if (platform === 'darwin') {
    return [
      '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      '/Applications/Chromium.app/Contents/MacOS/Chromium',
      posix.join(home, 'Applications/Google Chrome.app/Contents/MacOS/Google Chrome'),
    ]
  }
  if (platform === 'win32') {
    const roots = [env.ProgramFiles, env['ProgramFiles(x86)'], env.LOCALAPPDATA].filter((value): value is string => typeof value === 'string' && value.length > 0)
    return roots.flatMap(root => [
      win32.join(root, 'Google', 'Chrome', 'Application', 'chrome.exe'),
      win32.join(root, 'Chromium', 'Application', 'chrome.exe'),
      win32.join(root, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    ])
  }
  return ['/usr/bin/google-chrome', '/usr/bin/google-chrome-stable', '/usr/bin/chromium', '/usr/bin/chromium-browser', '/snap/bin/chromium']
}
