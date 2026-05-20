// ── BIFF Formula Token Decoder ───────────────────────────────────────
// Converts the RPN token stream used by BIFF8/XLS formulas into readable
// Excel A1 formulas. The decoder intentionally favours resilient output:
// unknown tokens are skipped or represented as placeholders instead of
// aborting the entire worksheet parse.

export interface FormulaExternSheetRef {
  firstSheet: number
  lastSheet: number
}

export interface FormulaNameRef {
  name: string
}

export interface FormulaDecodeContext {
  currentRow: number
  currentCol: number
  sheetNames?: string[]
  externSheets?: FormulaExternSheetRef[]
  names?: FormulaNameRef[]
}

const ERROR_TEXT: Record<number, string> = {
  0x00: "#NULL!",
  0x07: "#DIV/0!",
  0x0f: "#VALUE!",
  0x17: "#REF!",
  0x1d: "#NAME?",
  0x24: "#NUM!",
  0x2a: "#N/A",
}

const FIXED_ARG_COUNTS: Record<number, number> = {
  0: 1,
  2: 1,
  3: 1,
  4: 1,
  5: 1,
  6: 1,
  7: 1,
  8: 0,
  9: 0,
  10: 0,
  12: 1,
  15: 1,
  16: 1,
  17: 1,
  18: 1,
  19: 0,
  20: 1,
  21: 1,
  22: 1,
  23: 1,
  24: 1,
  25: 1,
  26: 1,
  27: 2,
  30: 2,
  31: 3,
  32: 1,
  33: 1,
  34: 0,
  35: 0,
  38: 1,
  39: 2,
  48: 2,
  56: 3,
  57: 3,
  58: 3,
  59: 3,
  60: 3,
  63: 0,
  64: 2,
  65: 3,
  66: 3,
  67: 1,
  68: 1,
  69: 1,
  70: 1,
  71: 1,
  72: 1,
  73: 1,
  74: 0,
  75: 1,
  76: 1,
  77: 1,
  86: 1,
  97: 2,
  98: 1,
  99: 1,
  105: 1,
  109: 2,
  111: 1,
  112: 1,
  113: 1,
  114: 1,
  115: 2,
  116: 2,
  117: 2,
  118: 1,
  119: 4,
  120: 4,
  121: 1,
  124: 2,
  125: 1,
  126: 1,
  127: 1,
  128: 1,
  129: 1,
  130: 1,
  131: 1,
  140: 1,
  141: 1,
  142: 3,
  162: 1,
  163: 1,
  167: 4,
  169: 1,
  183: 1,
  184: 1,
  190: 1,
  193: 1,
  194: 1,
  197: 1,
  198: 1,
  199: 3,
  212: 2,
  213: 2,
  216: 2,
  220: 2,
  221: 0,
  227: 1,
  228: 1,
  229: 1,
  230: 1,
  231: 1,
  232: 1,
  233: 1,
  234: 1,
  235: 3,
  244: 1,
  247: 5,
  252: 2,
  255: 1,
  261: 1,
  269: 1,
  270: 1,
  271: 1,
  272: 1,
  273: 1,
  274: 1,
  275: 1,
  279: 2,
  285: 1,
  300: 2,
  301: 2,
  303: 1,
  304: 1,
  305: 2,
  306: 2,
  307: 2,
}

const FUNCTION_NAMES: Record<number, string> = {
  0: "COUNT",
  1: "IF",
  2: "ISNA",
  3: "ISERROR",
  4: "SUM",
  5: "AVERAGE",
  6: "MIN",
  7: "MAX",
  8: "ROW",
  9: "COLUMN",
  10: "NA",
  11: "NPV",
  12: "STDEV",
  13: "DOLLAR",
  14: "FIXED",
  15: "SIN",
  16: "COS",
  17: "TAN",
  18: "ATAN",
  19: "PI",
  20: "SQRT",
  21: "EXP",
  22: "LN",
  23: "LOG10",
  24: "ABS",
  25: "INT",
  26: "SIGN",
  27: "ROUND",
  28: "LOOKUP",
  29: "INDEX",
  30: "REPT",
  31: "MID",
  32: "LEN",
  33: "VALUE",
  34: "TRUE",
  35: "FALSE",
  36: "AND",
  37: "OR",
  38: "NOT",
  39: "MOD",
  40: "DCOUNT",
  41: "DSUM",
  42: "DAVERAGE",
  43: "DMIN",
  44: "DMAX",
  45: "DSTDEV",
  46: "VAR",
  47: "DVAR",
  48: "TEXT",
  49: "LINEST",
  50: "TREND",
  51: "LOGEST",
  52: "GROWTH",
  56: "PV",
  57: "FV",
  58: "NPER",
  59: "PMT",
  60: "RATE",
  61: "MIRR",
  62: "IRR",
  63: "RAND",
  64: "MATCH",
  65: "DATE",
  66: "TIME",
  67: "DAY",
  68: "MONTH",
  69: "YEAR",
  70: "WEEKDAY",
  71: "HOUR",
  72: "MINUTE",
  73: "SECOND",
  74: "NOW",
  75: "AREAS",
  76: "ROWS",
  77: "COLUMNS",
  78: "OFFSET",
  82: "SEARCH",
  83: "TRANSPOSE",
  86: "TYPE",
  97: "ATAN2",
  98: "ASIN",
  99: "ACOS",
  100: "CHOOSE",
  101: "HLOOKUP",
  102: "VLOOKUP",
  105: "ISREF",
  109: "LOG",
  111: "CHAR",
  112: "LOWER",
  113: "UPPER",
  114: "PROPER",
  115: "LEFT",
  116: "RIGHT",
  117: "EXACT",
  118: "TRIM",
  119: "REPLACE",
  120: "SUBSTITUTE",
  121: "CODE",
  124: "FIND",
  125: "CELL",
  126: "ISERR",
  127: "ISTEXT",
  128: "ISNUMBER",
  129: "ISBLANK",
  130: "T",
  131: "N",
  140: "DATEVALUE",
  141: "TIMEVALUE",
  142: "SLN",
  162: "CLEAN",
  163: "MDETERM",
  167: "IPMT",
  169: "COUNTA",
  183: "PRODUCT",
  184: "FACT",
  190: "ISNONTEXT",
  193: "STDEVP",
  194: "VARP",
  197: "TRUNC",
  198: "ISLOGICAL",
  199: "DCOUNTA",
  204: "USDOLLAR",
  205: "FINDB",
  206: "SEARCHB",
  207: "REPLACEB",
  208: "LEFTB",
  209: "RIGHTB",
  210: "MIDB",
  212: "ROUNDUP",
  213: "ROUNDDOWN",
  216: "RANK",
  220: "DAYS360",
  221: "TODAY",
  222: "VDB",
  227: "MEDIAN",
  228: "SUMPRODUCT",
  229: "SINH",
  230: "COSH",
  231: "TANH",
  232: "ASINH",
  233: "ACOSH",
  234: "ATANH",
  235: "DGET",
  244: "INFO",
  247: "DB",
  252: "FREQUENCY",
  255: "ERROR.TYPE",
  261: "AVEDEV",
  269: "AVERAGEA",
  270: "MAXA",
  271: "MINA",
  272: "STDEVPA",
  273: "VARPA",
  274: "STDEVA",
  275: "VARA",
  276: "BAHTTEXT",
  279: "COMBIN",
  285: "MULTINOMIAL",
  300: "FLOOR",
  301: "CEILING",
  303: "ROMAN",
  304: "SUMSQ",
  305: "SUMX2MY2",
  306: "SUMX2PY2",
  307: "SUMXMY2",
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function u16(bytes: Uint8Array, offset: number): number {
  return view(bytes).getUint16(offset, true)
}

function f64(bytes: Uint8Array, offset: number): number {
  return view(bytes).getFloat64(offset, true)
}

function decodeUtf16Le(bytes: Uint8Array): string {
  try {
    return new TextDecoder("utf-16le").decode(bytes)
  } catch {
    let s = ""
    for (let i = 0; i + 1 < bytes.length; i += 2) {
      s += String.fromCharCode(bytes[i]! | (bytes[i + 1]! << 8))
    }
    return s
  }
}

function decodeCompressedUnicode(bytes: Uint8Array): string {
  let s = ""
  for (const b of bytes) s += String.fromCharCode(b)
  return s
}

function colToLetter(col: number): string {
  let result = ""
  let n = Math.max(0, col)
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result
    n = Math.floor(n / 26) - 1
  }
  return result
}

function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}

function escapeString(value: string): string {
  return `"${value.replace(/"/g, '""')}"`
}

function signed16(value: number): number {
  return value & 0x8000 ? value - 0x10000 : value
}

function formatCell(row: number, col: number, colFlags = 0): string {
  const rowRelative = (colFlags & 0x8000) !== 0
  const colRelative = (colFlags & 0x4000) !== 0
  const colText = `${colRelative ? "" : "$"}${colToLetter(col)}`
  const rowText = `${rowRelative ? "" : "$"}${row + 1}`
  return `${colText}${rowText}`
}

function readRef(data: Uint8Array, pos: number): { text: string; offset: number } {
  if (pos + 4 > data.length) return { text: "#REF!", offset: data.length }
  const row = u16(data, pos)
  const colFlags = u16(data, pos + 2)
  return { text: formatCell(row, colFlags & 0xff, colFlags), offset: pos + 4 }
}

function readRefN(
  data: Uint8Array,
  pos: number,
  ctx: FormulaDecodeContext,
): { text: string; offset: number } {
  if (pos + 4 > data.length) return { text: "#REF!", offset: data.length }
  const rowRaw = u16(data, pos)
  const colFlags = u16(data, pos + 2)
  const row = (colFlags & 0x8000) !== 0 ? ctx.currentRow + signed16(rowRaw) : rowRaw
  const colRaw = colFlags & 0xff
  const col = (colFlags & 0x4000) !== 0 ? ctx.currentCol + signed16(colRaw) : colRaw
  return { text: formatCell(row, col, colFlags), offset: pos + 4 }
}

function readArea(data: Uint8Array, pos: number): { text: string; offset: number } {
  if (pos + 8 > data.length) return { text: "#REF!", offset: data.length }
  const row1 = u16(data, pos)
  const row2 = u16(data, pos + 2)
  const col1Flags = u16(data, pos + 4)
  const col2Flags = u16(data, pos + 6)
  const start = formatCell(row1, col1Flags & 0xff, col1Flags)
  const end = formatCell(row2, col2Flags & 0xff, col2Flags)
  return { text: `${start}:${end}`, offset: pos + 8 }
}

function readAreaN(
  data: Uint8Array,
  pos: number,
  ctx: FormulaDecodeContext,
): { text: string; offset: number } {
  if (pos + 8 > data.length) return { text: "#REF!", offset: data.length }
  const row1Raw = u16(data, pos)
  const row2Raw = u16(data, pos + 2)
  const col1Flags = u16(data, pos + 4)
  const col2Flags = u16(data, pos + 6)
  const row1 = (col1Flags & 0x8000) !== 0 ? ctx.currentRow + signed16(row1Raw) : row1Raw
  const row2 = (col2Flags & 0x8000) !== 0 ? ctx.currentRow + signed16(row2Raw) : row2Raw
  const col1Raw = col1Flags & 0xff
  const col2Raw = col2Flags & 0xff
  const col1 = (col1Flags & 0x4000) !== 0 ? ctx.currentCol + signed16(col1Raw) : col1Raw
  const col2 = (col2Flags & 0x4000) !== 0 ? ctx.currentCol + signed16(col2Raw) : col2Raw
  return {
    text: `${formatCell(row1, col1, col1Flags)}:${formatCell(row2, col2, col2Flags)}`,
    offset: pos + 8,
  }
}

function sheetPrefix(ixti: number, ctx: FormulaDecodeContext): string {
  const ref = ctx.externSheets?.[ixti]
  if (!ref) return ""
  const first = ctx.sheetNames?.[ref.firstSheet]
  const last = ctx.sheetNames?.[ref.lastSheet]
  if (!first) return ""
  if (last && last !== first) return `${quoteSheetName(first)}:${quoteSheetName(last)}!`
  return `${quoteSheetName(first)}!`
}

function readArgs(stack: string[], count: number): string[] {
  const args: string[] = []
  for (let i = 0; i < count; i++) args.unshift(stack.pop() ?? "")
  return args
}

function callFunction(stack: string[], fnId: number, argCount?: number): void {
  const name = FUNCTION_NAMES[fnId] ?? `FUNC${fnId}`
  const count = argCount ?? FIXED_ARG_COUNTS[fnId] ?? 0
  const args = readArgs(stack, count)
  stack.push(`${name}(${args.join(",")})`)
}

function pushBinary(stack: string[], op: string): void {
  const right = stack.pop() ?? ""
  const left = stack.pop() ?? ""
  stack.push(`${left}${op}${right}`)
}

function pushUnary(stack: string[], prefix: string, suffix = ""): void {
  const value = stack.pop() ?? ""
  stack.push(`${prefix}${value}${suffix}`)
}

function skipBytes(pos: number, bytes: number, len: number): number {
  return Math.min(pos + bytes, len)
}

/** Decode a BIFF8 formula RPN token stream into an Excel formula body. */
export function decodeBiffFormula(tokens: Uint8Array, ctx: FormulaDecodeContext): string {
  const stack: string[] = []
  let pos = 0

  while (pos < tokens.length) {
    const ptg = tokens[pos++]!

    switch (ptg) {
      case 0x03:
        pushBinary(stack, "+")
        break
      case 0x04:
        pushBinary(stack, "-")
        break
      case 0x05:
        pushBinary(stack, "*")
        break
      case 0x06:
        pushBinary(stack, "/")
        break
      case 0x07:
        pushBinary(stack, "^")
        break
      case 0x08:
        pushBinary(stack, "&")
        break
      case 0x09:
        pushBinary(stack, "<")
        break
      case 0x0a:
        pushBinary(stack, "<=")
        break
      case 0x0b:
        pushBinary(stack, "=")
        break
      case 0x0c:
        pushBinary(stack, ">=")
        break
      case 0x0d:
        pushBinary(stack, ">")
        break
      case 0x0e:
        pushBinary(stack, "<>")
        break
      case 0x0f:
        pushBinary(stack, " ")
        break
      case 0x10:
        pushBinary(stack, ",")
        break
      case 0x11:
        pushBinary(stack, ":")
        break
      case 0x12:
        break
      case 0x13:
        pushUnary(stack, "-")
        break
      case 0x14:
        pushUnary(stack, "", "%")
        break
      case 0x15:
        pushUnary(stack, "(", ")")
        break
      case 0x16:
        stack.push("")
        break
      case 0x17: {
        if (pos >= tokens.length) break
        const cch = tokens[pos++] ?? 0
        const flags = tokens[pos++] ?? 0
        const is16 = (flags & 0x01) !== 0
        const bytes = cch * (is16 ? 2 : 1)
        const raw = tokens.subarray(pos, Math.min(pos + bytes, tokens.length))
        pos = skipBytes(pos, bytes, tokens.length)
        stack.push(escapeString(is16 ? decodeUtf16Le(raw) : decodeCompressedUnicode(raw)))
        break
      }
      case 0x19: {
        if (pos + 3 > tokens.length) {
          pos = tokens.length
          break
        }
        const grbit = tokens[pos] ?? 0
        const data = u16(tokens, pos + 1)
        pos += 3
        if ((grbit & 0x10) !== 0) callFunction(stack, 4, 1)
        else if ((grbit & 0x04) !== 0) pos = skipBytes(pos, data * 2, tokens.length)
        break
      }
      case 0x1d:
        stack.push((tokens[pos++] ?? 0) ? "TRUE" : "FALSE")
        break
      case 0x1e:
        if (pos + 2 <= tokens.length) {
          stack.push(String(u16(tokens, pos)))
          pos += 2
        } else pos = tokens.length
        break
      case 0x1f:
        if (pos + 8 <= tokens.length) {
          stack.push(String(f64(tokens, pos)))
          pos += 8
        } else pos = tokens.length
        break
      case 0x1c:
        stack.push(ERROR_TEXT[tokens[pos++] ?? 0] ?? "#VALUE!")
        break
      case 0x01:
      case 0x02:
        pos = skipBytes(pos, 4, tokens.length)
        stack.push("#FORMULA!")
        break
      default:
        if (ptg === 0x21 || ptg === 0x41 || ptg === 0x61) {
          if (pos + 2 <= tokens.length) {
            const fnId = u16(tokens, pos)
            pos += 2
            callFunction(stack, fnId)
          } else pos = tokens.length
        } else if (ptg === 0x22 || ptg === 0x42 || ptg === 0x62) {
          if (pos + 3 <= tokens.length) {
            const argc = tokens[pos++] ?? 0
            const fnId = u16(tokens, pos)
            pos += 2
            callFunction(stack, fnId, argc & 0x7f)
          } else pos = tokens.length
        } else if (ptg === 0x23 || ptg === 0x43 || ptg === 0x63) {
          if (pos + 4 <= tokens.length) {
            const nameIndex = u16(tokens, pos) - 1
            pos += 4
            stack.push(ctx.names?.[nameIndex]?.name ?? `Name${nameIndex + 1}`)
          } else pos = tokens.length
        } else if (ptg === 0x24 || ptg === 0x44 || ptg === 0x64) {
          const ref = readRef(tokens, pos)
          stack.push(ref.text)
          pos = ref.offset
        } else if (ptg === 0x25 || ptg === 0x45 || ptg === 0x65) {
          const area = readArea(tokens, pos)
          stack.push(area.text)
          pos = area.offset
        } else if (ptg === 0x2a || ptg === 0x4a || ptg === 0x6a) {
          pos = skipBytes(pos, 4, tokens.length)
          stack.push("#REF!")
        } else if (ptg === 0x2b || ptg === 0x4b || ptg === 0x6b) {
          pos = skipBytes(pos, 8, tokens.length)
          stack.push("#REF!")
        } else if (ptg === 0x2c || ptg === 0x4c || ptg === 0x6c) {
          const ref = readRefN(tokens, pos, ctx)
          stack.push(ref.text)
          pos = ref.offset
        } else if (ptg === 0x2d || ptg === 0x4d || ptg === 0x6d) {
          const area = readAreaN(tokens, pos, ctx)
          stack.push(area.text)
          pos = area.offset
        } else if (ptg === 0x39 || ptg === 0x59 || ptg === 0x79) {
          pos = skipBytes(pos, 6, tokens.length)
          stack.push("#NAME?")
        } else if (ptg === 0x3a || ptg === 0x5a || ptg === 0x7a) {
          if (pos + 6 <= tokens.length) {
            const prefix = sheetPrefix(u16(tokens, pos), ctx)
            const ref = readRef(tokens, pos + 2)
            stack.push(`${prefix}${ref.text}`)
            pos = ref.offset
          } else pos = tokens.length
        } else if (ptg === 0x3b || ptg === 0x5b || ptg === 0x7b) {
          if (pos + 10 <= tokens.length) {
            const prefix = sheetPrefix(u16(tokens, pos), ctx)
            const area = readArea(tokens, pos + 2)
            stack.push(`${prefix}${area.text}`)
            pos = area.offset
          } else pos = tokens.length
        } else if (ptg === 0x3c || ptg === 0x5c || ptg === 0x7c) {
          pos = skipBytes(pos, 6, tokens.length)
          stack.push("#REF!")
        } else if (ptg === 0x3d || ptg === 0x5d || ptg === 0x7d) {
          pos = skipBytes(pos, 10, tokens.length)
          stack.push("#REF!")
        } else if (ptg === 0x20 || ptg === 0x40 || ptg === 0x60) {
          pos = skipBytes(pos, 7, tokens.length)
          stack.push("{...}")
        } else if (ptg === 0x26 || ptg === 0x46 || ptg === 0x66) {
          pos = skipBytes(pos, 6, tokens.length)
        } else if (ptg === 0x27 || ptg === 0x47 || ptg === 0x67) {
          pos = skipBytes(pos, 6, tokens.length)
        } else if (ptg === 0x29 || ptg === 0x49 || ptg === 0x69) {
          pos = skipBytes(pos, 2, tokens.length)
        } else {
          // Unknown extension or future token. Stop to avoid desynchronising
          // all remaining tokens; the cached cell value remains available.
          pos = tokens.length
        }
        break
    }
  }

  return stack.length > 0 ? stack[stack.length - 1]! : ""
}
