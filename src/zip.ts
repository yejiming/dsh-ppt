import { decompressSync, deflateSync, strFromU8, strToU8 } from 'fflate'
import { PptError, asPptError } from './errors.ts'

/**
 * OOXML parts that pptxgenjs cannot produce — slide transitions, animation
 * timelines, embedded font parts — are injected by rewriting the finished
 * package. Untouched entries keep their original local header bytes, so a
 * package that is not targeted stays byte-comparable with what pptxgenjs wrote.
 */

export interface ZipEntryInfo {
  name: string
  method: number
  compressed_size: number
  uncompressed_size: number
  local_header_offset: number
}

export interface ZipEntry extends ZipEntryInfo {
  flags: number
  version_needed: number
  time: number
  date: number
  crc: number
  local_extra: Uint8Array
  data_offset: number
  header: Uint8Array
  central: Uint8Array
}

export interface ZipDirectory {
  entries: ZipEntry[]
  comment: Uint8Array
}

const ZIP_LOCAL_SIGNATURE = 0x04034b50
const ZIP_CENTRAL_SIGNATURE = 0x02014b50
const ZIP_END_SIGNATURE = 0x06054b50
const ZIP_FLAG_ENCRYPTED = 0x1
const ZIP_FLAG_DESCRIPTOR = 0x8
export const ZIP_METHOD_STORE = 0
export const ZIP_METHOD_DEFLATE = 8
const ZIP64_SENTINEL = 0xffffffff
/** DOS epoch: a fixed timestamp keeps a rewritten package reproducible. */
const DOS_TIME = 0
const DOS_DATE = 0x0021

const CRC32_TABLE = ((): Uint32Array => {
  const table = new Uint32Array(256)
  for (let index = 0; index < 256; index += 1) {
    let value = index
    for (let bit = 0; bit < 8; bit += 1) value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
    table[index] = value >>> 0
  }
  return table
})()

export function crc32(data: Uint8Array): number {
  let value = 0xffffffff
  for (let index = 0; index < data.length; index += 1) {
    value = (value >>> 8) ^ CRC32_TABLE[(value ^ data[index]!) & 0xff]!
  }
  return (value ^ 0xffffffff) >>> 0
}

function findEndOfCentralDirectory(view: DataView, length: number): number {
  if (length < 22) throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX is not a readable ZIP package')
  for (let offset = length - 22; offset >= 0; offset -= 1) {
    if (view.getUint32(offset, true) !== ZIP_END_SIGNATURE) continue
    if (offset + 22 + view.getUint16(offset + 20, true) === length) return offset
  }
  throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX is not a readable ZIP package: end of central directory is missing')
}

function parseZipDirectory(data: Uint8Array): ZipDirectory {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  const end = findEndOfCentralDirectory(view, data.byteLength)
  const count = view.getUint16(end + 10, true)
  const comment = data.slice(end + 22, end + 22 + view.getUint16(end + 20, true))
  const entries: ZipEntry[] = []
  let cursor = view.getUint32(end + 16, true)
  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(cursor, true) !== ZIP_CENTRAL_SIGNATURE) {
      throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX zip central directory is malformed')
    }
    const flags = view.getUint16(cursor + 8, true)
    const method = view.getUint16(cursor + 10, true)
    const crc = view.getUint32(cursor + 16, true)
    const compressedSize = view.getUint32(cursor + 20, true)
    const uncompressedSize = view.getUint32(cursor + 24, true)
    const nameLength = view.getUint16(cursor + 28, true)
    const extraLength = view.getUint16(cursor + 30, true)
    const entryCommentLength = view.getUint16(cursor + 32, true)
    const localOffset = view.getUint32(cursor + 42, true)
    const centralSize = 46 + nameLength + extraLength + entryCommentLength
    if ((flags & ZIP_FLAG_ENCRYPTED) !== 0 || (flags & ZIP_FLAG_DESCRIPTOR) !== 0) {
      throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX zip entries with encryption or data descriptors are unsupported')
    }
    if (compressedSize === ZIP64_SENTINEL || uncompressedSize === ZIP64_SENTINEL || localOffset === ZIP64_SENTINEL) {
      throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX zip64 archives are unsupported')
    }
    if (view.getUint32(localOffset, true) !== ZIP_LOCAL_SIGNATURE) {
      throw new PptError('PPT_CREATE_INVALID_PACKAGE', 'PPTX zip local header is malformed')
    }
    const localNameLength = view.getUint16(localOffset + 26, true)
    const localExtraLength = view.getUint16(localOffset + 28, true)
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength
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
      central: data.slice(cursor, cursor + centralSize),
    })
    cursor += centralSize
  }
  return { entries, comment }
}

export function readZipDirectory(data: Uint8Array): ZipDirectory {
  try {
    return parseZipDirectory(data)
  } catch (error) {
    throw asPptError(error, 'PPT_CREATE_INVALID_PACKAGE', 'PPTX zip directory is not readable')
  }
}

export function readEntryBytes(entry: ZipEntry, data: Uint8Array): Uint8Array {
  const raw = data.subarray(entry.data_offset, entry.data_offset + entry.compressed_size)
  if (entry.method === ZIP_METHOD_STORE) return raw.slice()
  if (entry.method !== ZIP_METHOD_DEFLATE) {
    throw new PptError('PPT_CREATE_INVALID_PACKAGE', `unsupported zip compression method ${entry.method} for ${entry.name}`)
  }
  try {
    // fflate auto-detects gzip, zlib, and raw DEFLATE, so a deflated zip entry needs no format option.
    return decompressSync(raw)
  } catch (error) {
    throw new PptError('PPT_CREATE_INVALID_PACKAGE', `zip entry is not readable: ${entry.name}`, { cause: error })
  }
}

export function inspectZipEntries(data: Uint8Array): ZipEntryInfo[] {
  return readZipDirectory(data).entries.map(entry => ({
    name: entry.name,
    method: entry.method,
    compressed_size: entry.compressed_size,
    uncompressed_size: entry.uncompressed_size,
    local_header_offset: entry.local_header_offset,
  }))
}

interface RebuiltEntry {
  header: Uint8Array
  crc: number
  compressed_size: number
  uncompressed_size: number
}

function rebuildLocalEntry(entry: ZipEntry, contents: Uint8Array): RebuiltEntry {
  const name = strToU8(entry.name)
  const payload = entry.method === ZIP_METHOD_DEFLATE ? deflateSync(contents, { level: 6 }) : contents.slice()
  const crc = crc32(contents)
  const header = new Uint8Array(30 + name.byteLength + entry.local_extra.byteLength + payload.byteLength)
  const view = new DataView(header.buffer)
  view.setUint32(0, ZIP_LOCAL_SIGNATURE, true)
  view.setUint16(4, Math.max(entry.version_needed, 20), true)
  view.setUint16(6, entry.flags, true)
  view.setUint16(8, entry.method, true)
  view.setUint16(10, entry.time, true)
  view.setUint16(12, entry.date, true)
  view.setUint32(14, crc, true)
  view.setUint32(18, payload.byteLength, true)
  view.setUint32(22, contents.byteLength, true)
  view.setUint16(26, name.byteLength, true)
  view.setUint16(28, entry.local_extra.byteLength, true)
  header.set(name, 30)
  header.set(entry.local_extra, 30 + name.byteLength)
  header.set(payload, 30 + name.byteLength + entry.local_extra.byteLength)
  return { header, crc, compressed_size: payload.byteLength, uncompressed_size: contents.byteLength }
}

/** Builds the local header and central directory record for an appended part. */
function buildAddedEntry(name: string, contents: Uint8Array): { header: Uint8Array; central: Uint8Array } & RebuiltEntry {
  const nameBytes = strToU8(name)
  const payload = deflateSync(contents, { level: 6 })
  const crc = crc32(contents)
  const header = new Uint8Array(30 + nameBytes.byteLength + payload.byteLength)
  const view = new DataView(header.buffer)
  view.setUint32(0, ZIP_LOCAL_SIGNATURE, true)
  view.setUint16(4, 20, true)
  view.setUint16(6, 0, true)
  view.setUint16(8, ZIP_METHOD_DEFLATE, true)
  view.setUint16(10, DOS_TIME, true)
  view.setUint16(12, DOS_DATE, true)
  view.setUint32(14, crc, true)
  view.setUint32(18, payload.byteLength, true)
  view.setUint32(22, contents.byteLength, true)
  view.setUint16(26, nameBytes.byteLength, true)
  view.setUint16(28, 0, true)
  header.set(nameBytes, 30)
  header.set(payload, 30 + nameBytes.byteLength)
  const central = new Uint8Array(46 + nameBytes.byteLength)
  const centralView = new DataView(central.buffer)
  centralView.setUint32(0, ZIP_CENTRAL_SIGNATURE, true)
  centralView.setUint16(4, 20, true)
  centralView.setUint16(6, 20, true)
  centralView.setUint16(8, 0, true)
  centralView.setUint16(10, ZIP_METHOD_DEFLATE, true)
  centralView.setUint16(12, DOS_TIME, true)
  centralView.setUint16(14, DOS_DATE, true)
  centralView.setUint32(16, crc, true)
  centralView.setUint32(20, payload.byteLength, true)
  centralView.setUint32(24, contents.byteLength, true)
  centralView.setUint16(28, nameBytes.byteLength, true)
  centralView.setUint32(42, 0, true)
  central.set(nameBytes, 46)
  return { header, central, crc, compressed_size: payload.byteLength, uncompressed_size: contents.byteLength }
}

function patchCentralEntry(entry: ZipEntry, offset: number, crc: number, compressedSize: number, uncompressedSize: number): Uint8Array {
  const central = entry.central.slice()
  const view = new DataView(central.buffer, central.byteOffset, central.byteLength)
  view.setUint32(16, crc, true)
  view.setUint32(20, compressedSize, true)
  view.setUint32(24, uncompressedSize, true)
  view.setUint32(42, offset, true)
  return central
}

export interface ZipRewritePlan {
  /**
   * Called once per entry with a lazy reader. Returning replacement bytes swaps
   * the part, `undefined` keeps it byte-identical, and `null` drops it. Entries
   * whose content is never read stay compressed on disk exactly as they were.
   */
  transform?: (name: string, read: () => Uint8Array) => Uint8Array | null | undefined
  /** Parts to append; an existing entry with the same name is an error. */
  additions?: ReadonlyMap<string, Uint8Array>
}

export function rewriteZip(data: Uint8Array, plan: ZipRewritePlan): Uint8Array {
  const directory = readZipDirectory(data)
  const present = new Set(directory.entries.map(entry => entry.name))
  for (const name of plan.additions?.keys() ?? []) {
    if (present.has(name)) throw new PptError('PPT_CREATE_INVALID_PACKAGE', `zip entry already exists: ${name}`)
  }
  const locals: Uint8Array[] = []
  const centrals: Uint8Array[] = []
  let offset = 0
  let count = 0
  for (const entry of directory.entries) {
    const replacement = plan.transform?.(entry.name, () => readEntryBytes(entry, data))
    if (replacement === null) continue
    let header = entry.header
    let crc = entry.crc
    let compressedSize = entry.compressed_size
    let uncompressedSize = entry.uncompressed_size
    if (replacement !== undefined) {
      const rebuilt = rebuildLocalEntry(entry, replacement)
      header = rebuilt.header
      crc = rebuilt.crc
      compressedSize = rebuilt.compressed_size
      uncompressedSize = rebuilt.uncompressed_size
    }
    locals.push(header)
    centrals.push(patchCentralEntry(entry, offset, crc, compressedSize, uncompressedSize))
    offset += header.byteLength
    count += 1
  }
  for (const [name, contents] of plan.additions ?? []) {
    const added = buildAddedEntry(name, contents)
    const central = new DataView(added.central.buffer)
    central.setUint32(42, offset, true)
    locals.push(added.header)
    centrals.push(added.central)
    offset += added.header.byteLength
    count += 1
  }
  const centralSize = centrals.reduce((total, entry) => total + entry.byteLength, 0)
  const output = new Uint8Array(offset + centralSize + 22 + directory.comment.byteLength)
  let cursor = 0
  for (const part of [...locals, ...centrals]) {
    output.set(part, cursor)
    cursor += part.byteLength
  }
  const view = new DataView(output.buffer)
  view.setUint32(cursor, ZIP_END_SIGNATURE, true)
  view.setUint16(cursor + 8, count, true)
  view.setUint16(cursor + 10, count, true)
  view.setUint32(cursor + 12, centralSize, true)
  view.setUint32(cursor + 16, offset, true)
  view.setUint16(cursor + 20, directory.comment.byteLength, true)
  output.set(directory.comment, cursor + 22)
  return output
}
