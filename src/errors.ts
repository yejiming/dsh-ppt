export const PPT_ERROR_CODES = [
  'PPT_ABORTED',
  'PPT_PLATFORM_UNSUPPORTED',
  'PPT_RESOURCE_LIMIT',
  'PPT_PATH_OUTSIDE_WORKSPACE',
  'PPT_PATH_INVALID',
  'PPT_OUTPUT_EXISTS',
  'PPT_DEPENDENCY_MISSING',
  'PPT_CAPABILITY_UNAVAILABLE',
  'BROWSER_URL_BLOCKED',
  'BROWSER_REF_STALE',
  'BROWSER_NOT_READY',
  'BROWSER_LIMIT_EXCEEDED',
  'PYTHON_DEPENDENCY_MISSING',
  'PYTHON_EXECUTION_FAILED',
  'IMAGE_SEARCH_FAILED',
  'IMAGE_ASSET_INVALID',
  'PPT_OUTLINE_INVALID',
  'PPT_ART_DIRECTION_INVALID',
  'PPT_THEME_INVALID',
  'PPT_THEME_UNKNOWN',
  'HTML_CREATE_INPUT_INVALID',
  'HTML_CREATE_UNSUPPORTED_CSS',
  'HTML_CREATE_VALIDATION_FAILED',
  'PPT_CREATE_INPUT_INVALID',
  'PPT_CREATE_UNSUPPORTED_ELEMENT',
  'PPT_CREATE_ASSET_MISSING',
  'PPT_CREATE_WRITE_FAILED',
  'PPT_CREATE_INVALID_PACKAGE',
  'PPT_CREATE_ABORTED',
  'PPT_RENDER_NOT_AVAILABLE',
  'PPT_RENDER_FAILED',
  'PPT_QUALITY_FAILED',
] as const

export type PptErrorCode = typeof PPT_ERROR_CODES[number]

export interface PptErrorOptions {
  cause?: unknown
  details?: Readonly<Record<string, unknown>>
}

export class PptError extends Error {
  readonly code: PptErrorCode
  readonly details?: Readonly<Record<string, unknown>>

  constructor(code: PptErrorCode, message: string, options: PptErrorOptions = {}) {
    super(message, { cause: options.cause })
    this.name = 'PptError'
    this.code = code
    this.details = options.details
  }

  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      code: this.code,
      message: this.message,
      ...(this.details === undefined ? {} : { details: this.details }),
    }
  }
}

export function asPptError(error: unknown, code: PptErrorCode, prefix?: string): PptError {
  if (error instanceof PptError) return error
  const message = error instanceof Error ? error.message : String(error)
  return new PptError(code, prefix === undefined ? message : `${prefix}: ${message}`, { cause: error })
}

export function throwIfAborted(signal?: AbortSignal, code: PptErrorCode = 'PPT_ABORTED'): void {
  if (!signal?.aborted) return
  const reason = signal.reason instanceof Error ? signal.reason.message : String(signal.reason ?? 'aborted')
  throw new PptError(code, `operation aborted: ${reason}`)
}
