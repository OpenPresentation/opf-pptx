// A single left-to-right PowerShell lexer for offline source-policy scans.
// It blanks comments and literal string text to spaces (keeping CR/LF and every code-unit offset), so a regex over the
// result sees only code. Constructs this lexer does not model are reported as issues so callers can fail closed.

// Windows PowerShell 5.1 also accepts the typographic quote characters as string delimiters.
const SINGLE_QUOTES = new Set(["'", '\u2018', '\u2019', '\u201A', '\u201B']);
const DOUBLE_QUOTES = new Set(['"', '\u201C', '\u201D', '\u201E']);
// A '#' after one of these starts a comment unambiguously; after anything else PowerShell may read it as part of a
// bare-word argument (for example `Write-Output a#b`), so it is reported rather than guessed.
const COMMENT_BOUNDARY = /[\s;,(){}|&]/;

export function scanPowerShellSource(sourceText) {
  const text = String(sourceText);
  const out = Array.from({length: text.length}, (_, index) => text[index]);
  const issues = [];
  const issue = (kind, index) => issues.push({kind, index});
  const blank = (from, to) => { for (let index = from; index < to && index < text.length; index++) if (out[index] !== '\r' && out[index] !== '\n') out[index] = ' '; };
  const atLineStart = index => index === 0 || text[index - 1] === '\n';

  // Scans code from `start`. Nested (a `$(...)` inside an expandable string) stops at the unmatched ')' and returns its index.
  function scanCode(start, nested) {
    let index = start, depth = 0;
    while (index < text.length) {
      const char = text[index], next = text[index + 1];
      if (char === '`') { index += 2; continue; }
      if (char === '<' && next === '#') { index = scanBlockComment(index); continue; }
      if (char === '#') {
        if (index > 0 && !COMMENT_BOUNDARY.test(text[index - 1])) issue('ambiguous-comment-start', index);
        index = scanLineComment(index); continue;
      }
      if (char === '@' && (SINGLE_QUOTES.has(next) || DOUBLE_QUOTES.has(next))) {
        const end = scanHereString(index, DOUBLE_QUOTES.has(next));
        if (end !== null) { index = end; continue; }
        index += 1; continue;
      }
      if (SINGLE_QUOTES.has(char)) { index = scanSingleQuoted(index); continue; }
      if (DOUBLE_QUOTES.has(char)) { index = scanDoubleQuoted(index); continue; }
      if (char === '$' && next === '{') { index = scanBracedVariable(index); continue; }
      if (nested && char === '(') depth++;
      if (nested && char === ')') { if (depth === 0) return index; depth--; }
      index++;
    }
    if (nested) issue('unterminated-subexpression', start);
    return index;
  }

  function scanLineComment(start) {
    let index = start;
    while (index < text.length && text[index] !== '\r' && text[index] !== '\n') index++;
    blank(start, index);
    return index;
  }

  function scanBlockComment(start) {
    const close = text.indexOf('#>', start + 2);
    const end = close === -1 ? text.length : close + 2;
    if (close === -1) issue('unterminated-block-comment', start);
    blank(start, end);
    return end;
  }

  // '...' with '' as the only escape; delimiters are kept, the text between them is blanked.
  function scanSingleQuoted(start) {
    let index = start + 1;
    while (index < text.length) {
      if (SINGLE_QUOTES.has(text[index])) {
        if (SINGLE_QUOTES.has(text[index + 1])) { index += 2; continue; }
        blank(start + 1, index);
        return index + 1;
      }
      index++;
    }
    issue('unterminated-string', start);
    blank(start + 1, text.length);
    return text.length;
  }

  // "..." with backtick escapes and "" doubling; $(...) subexpressions are code and stay visible.
  function scanDoubleQuoted(start) {
    let index = start + 1, literalStart = index;
    while (index < text.length) {
      const char = text[index];
      if (char === '`') { index += 2; continue; }
      if (DOUBLE_QUOTES.has(char)) {
        if (DOUBLE_QUOTES.has(text[index + 1])) { index += 2; continue; }
        blank(literalStart, index);
        return index + 1;
      }
      if (char === '$' && text[index + 1] === '(') {
        blank(literalStart, index);
        index = scanCode(index + 2, true) + 1;
        literalStart = index;
        continue;
      }
      index++;
    }
    issue('unterminated-string', start);
    blank(literalStart, text.length);
    return text.length;
  }

  // @'...'@ and @"..."@: the header must end its line and the terminator must start a line. Returns null when the
  // opener is not a valid here-string header so the caller rescans it as ordinary code.
  function scanHereString(start, expandable) {
    let index = start + 2;
    while (text[index] === ' ' || text[index] === '\t') index++;
    if (text[index] === '\r' && text[index + 1] === '\n') index += 2;
    else if (text[index] === '\n') index += 1;
    else { issue('here-string-header', start); return null; }
    const quotes = expandable ? DOUBLE_QUOTES : SINGLE_QUOTES;
    let literalStart = start + 2;
    while (index < text.length) {
      if (atLineStart(index) && quotes.has(text[index]) && text[index + 1] === '@') {
        blank(literalStart, index);
        return index + 2;
      }
      if (expandable && text[index] === '`') { index += 2; continue; }
      if (expandable && text[index] === '$' && text[index + 1] === '(') {
        blank(literalStart, index);
        index = scanCode(index + 2, true) + 1;
        literalStart = index;
        continue;
      }
      index++;
    }
    issue('unterminated-here-string', start);
    blank(literalStart, text.length);
    return text.length;
  }

  // ${name} may contain quotes or '#'; it is a variable reference, so it stays visible but is never tokenized.
  function scanBracedVariable(start) {
    let index = start + 2;
    while (index < text.length && text[index] !== '}') index += text[index] === '`' ? 2 : 1;
    if (index >= text.length) { issue('unterminated-braced-variable', start); return text.length; }
    return index + 1;
  }

  scanCode(0, false);
  return {code: out.join(''), issues};
}

export function stripPowerShellLiteralsForScan(sourceText) {
  return scanPowerShellSource(sourceText).code;
}

const OPENERS = {')': '(', ']': '[', '}': '{'};
const ASSIGNMENT_AHEAD = String.raw`(?=\s*(?:[-+*/%]?=(?!=)|\+\+|--))`;
const QUOTE_CLASS = `['"\u2018-\u201E]`;

// Walks back from `end` (exclusive) over a member/index/invocation chain to its root variable.
function chainRoot(code, end) {
  let pos = end;
  while (pos > 0) {
    const char = code[pos - 1];
    if (OPENERS[char]) {
      let depth = 0, index = pos - 1;
      for (; index >= 0; index--) {
        if (OPENERS[code[index]]) depth++;
        else if (Object.values(OPENERS).includes(code[index]) && --depth === 0) break;
      }
      if (index < 0) return null;
      if (char === '}' && code[index - 1] === '$') return {root: code.slice(index + 1, pos - 1).replace(/^(?:script|global|local|private|using):/i, ''), start: index - 1};
      pos = index;
      continue;
    }
    if (char === '.') { pos--; continue; }
    if (/\w/.test(char)) {
      let start = pos;
      while (start > 0 && /[\w:]/.test(code[start - 1])) start--;
      const token = code.slice(start, pos);
      if (code[start - 1] === '$' && !token.startsWith(':')) return {root: token.replace(/^(?:script|global|local|private|using):/i, ''), start: start - 1};
      if (code[start - 1] === '.' && !token.includes(':')) { pos = start - 1; continue; }
      return null;
    }
    return null;
  }
  return null;
}

// Every assignment (=, op=, ++, --) whose target is a member or index expression, in code only.
// `root` is the variable the chain starts from (scope prefix removed) or null when it is not a plain variable.
export function memberAssignments(sourceText) {
  const {code} = scanPowerShellSource(sourceText);
  const found = [];
  const targets = new RegExp(String.raw`(?:\.|::)\s*([A-Za-z_]\w*)${ASSIGNMENT_AHEAD}|\]${ASSIGNMENT_AHEAD}`, 'g');
  for (const match of code.matchAll(targets)) {
    const end = match.index + match[0].length;
    const isIndex = match[0] === ']';
    const origin = match[0].startsWith('::') ? null : chainRoot(code, isIndex ? end : match.index);
    found.push({index: match.index, root: origin?.root ?? null, path: origin ? code.slice(origin.start, end).replace(/\s+/g, '') : code.slice(match.index, end).replace(/\s+/g, ''), member: isIndex ? null : match[1]});
  }
  for (const match of code.matchAll(/(?:\+\+|--)\s*(\$\{[^}]*\}|\$[\w:]+)((?:\.\w+|\[[^\]]*\])+)/g)) {
    const root = match[1].startsWith('${') ? match[1].slice(2, -1) : match[1].slice(1);
    found.push({index: match.index, root: root.replace(/^(?:script|global|local|private|using):/i, ''), path: (match[1] + match[2]).replace(/\s+/g, ''), member: /\.(\w+)$/.exec(match[2])?.[1] ?? null});
  }
  return found.sort((left, right) => left.index - right.index);
}

// Member assignments outside the allowed local roots and exact single-level COM setters (for example 'wholeFont.Name'),
// plus every set_X(...) setter-method invocation. Comparisons are case-insensitive, as PowerShell's are.
export function disallowedMemberAssignments(sourceText, {localRoots = [], comSetters = []} = {}) {
  const locals = new Set(localRoots.map(name => name.toLowerCase()));
  const setters = new Set(comSetters.map(name => `$${name}`.toLowerCase()));
  const rejected = memberAssignments(sourceText).filter(item => {
    if (item.root !== null && locals.has(item.root.toLowerCase())) return false;
    return !setters.has(item.path.replace(/^\$(?:script|global|local|private|using):/i, '$').toLowerCase());
  }).map(item => item.path);
  const code = stripPowerShellLiteralsForScan(sourceText);
  for (const match of code.matchAll(/\.\s*(set_\w+)\s*\(/gi)) rejected.push(`.${match[1]}()`);
  return rejected;
}

// Extent of `function NAME { ... }` in lexed code, or null.
export function functionExtent(code, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const matches = [...code.matchAll(new RegExp(String.raw`(?<![\w-])function\s+${escaped}(?![\w-])[^{]*\{`, 'gi'))];
  if (matches.length !== 1) return null;
  let depth = 0;
  for (let index = matches[0].index + matches[0][0].length - 1; index < code.length; index++) {
    if (code[index] === '{') depth++;
    else if (code[index] === '}' && --depth === 0) return {start: matches[0].index, end: index + 1};
  }
  return null;
}

// Every `function NAME {...}` extent in lexed code.
function functionExtents(code) {
  const extents = [];
  for (const match of code.matchAll(/(?<![\w-])function\s+([\w-]+)[^{]*\{/gi)) {
    let depth = 0;
    for (let index = match.index + match[0].length - 1; index < code.length; index++) {
      if (code[index] === '{') depth++;
      else if (code[index] === '}' && --depth === 0) { extents.push({name: match[1], start: match.index, end: index + 1}); break; }
    }
  }
  return extents;
}

// Name of the innermost function containing `index`, or null at script level.
function enclosingFunction(extents, index) {
  let owner = null;
  for (const extent of extents) if (index > extent.start && index < extent.end && (owner === null || extent.start > owner.start)) owner = extent;
  return owner?.name ?? null;
}

const SCOPE_PREFIX = String.raw`(?:(?:script|global|local|private):)?`;

// Dynamic code, all forbidden except the listed exemptions:
// - Invoke-Expression, iex, Add-Type, Invoke-Command, icm (optionally module-qualified). Only exact `exemptInvocations`
//   inside the one `exemptFunction` body are allowed, so a pure-regression helper may re-evaluate its own definitions.
// - `& 'name'` string-named commands and `$x.'Member'()` string-named member calls.
// - [scriptblock]::Create, .InvokeScript(, .NewScriptBlock( and any $ExecutionContext reference.
// - `& <expression>` and `. <expression>` whose target is not a literal: allowed only for an exact `invocationSites`
//   entry {function (null = script level), operator ('&' or '.'), variable}, for example the COM wrapper's `& $Operation`.
export function dynamicCodeInvocations(sourceText, {exemptFunction = null, exemptInvocations = [], invocationSites = []} = {}) {
  const code = stripPowerShellLiteralsForScan(sourceText);
  const extent = exemptFunction ? functionExtent(code, exemptFunction) : null;
  const extents = functionExtents(code);
  const found = [];
  for (const match of code.matchAll(/(?<![\w$.\\-])(?:[\w.]+\\)?(Invoke-Expression|iex|Add-Type|Invoke-Command|icm)(?![\w-])/gi)) {
    const line = code.slice(match.index, code.indexOf('\n', match.index) === -1 ? code.length : code.indexOf('\n', match.index)).trimEnd();
    const exempt = extent !== null && match.index > extent.start && match.index < extent.end && match[0] === 'Invoke-Expression' && exemptInvocations.some(text => line === text);
    if (!exempt) found.push(match[0]);
  }
  for (const match of code.matchAll(new RegExp(String.raw`(?:^|[^\w$\])])&\s*${QUOTE_CLASS}`, 'gm'))) found.push(`string-named command at ${match.index}`);
  for (const match of code.matchAll(new RegExp(String.raw`\.\s*${QUOTE_CLASS}[^\r\n]*?${QUOTE_CLASS}\s*\(`, 'g'))) found.push(`string-named member at ${match.index}`);
  for (const match of code.matchAll(/\[\s*(?:System\.Management\.Automation\.)?ScriptBlock\s*\]\s*::\s*Create\s*\(/gi)) found.push(`[scriptblock]::Create at ${match.index}`);
  for (const match of code.matchAll(/\.\s*(InvokeScript|NewScriptBlock)\s*\(/gi)) found.push(`.${match[1]}() at ${match.index}`);
  for (const match of code.matchAll(/\.\$(?:\{[^}]*\}|\w+|\()[^\s=]*?\(/g)) found.push(`dynamic member name at ${match.index}`);
  for (const match of code.matchAll(new RegExp(String.raw`\$\{?${SCOPE_PREFIX}ExecutionContext(?![\w])`, 'gi'))) found.push(`$ExecutionContext at ${match.index}`);
  // `&` not part of a word, `&&` or a redirection such as 2>&1; `.` dot-source at a statement boundary followed by space.
  const invocations = [
    ...[...code.matchAll(/(?<![\w$\])}>&])&(?!&)\s*(?=[$(])/g)].map(match => ({operator: '&', match})),
    ...[...code.matchAll(/(?<=^|[\s;{(|])\.\s+(?=[$(])/gm)].map(match => ({operator: '.', match})),
  ];
  for (const {operator, match} of invocations) {
    const rest = code.slice(match.index + match[0].length);
    const variable = new RegExp(String.raw`^\$${SCOPE_PREFIX}(\w+)(?![\w.\[(:])`, 'i').exec(rest)?.[1] ?? null;
    const owner = enclosingFunction(extents, match.index);
    const allowed = variable !== null && invocationSites.some(site => (site.function ?? null) === owner && site.operator === operator && site.variable === variable);
    if (!allowed) found.push(`dynamic ${operator} invocation${variable ? ` of $${variable}` : ''} in ${owner ?? 'script scope'} at ${match.index}`);
  }
  return found;
}

// Shared static harness policy: unmodeled syntax fails closed, dynamic code is forbidden outside the exempt
// pure-regression helper and listed invocation sites, and member assignments are limited to local report roots and the
// documented COM setters.
export function auditHarnessSourcePolicy(sourceText, {label, localRoots = [], comSetters = [], exemptFunction = null, exemptInvocations = [], invocationSites = []} = {}) {
  const failures = [];
  const {issues} = scanPowerShellSource(sourceText);
  if (issues.length) failures.push({code: 'source-scan-unsupported', message: `${label} has PowerShell syntax the source-policy lexer does not model: ${issues.map(item => `${item.kind}@${item.index}`).join(', ')}`});
  const dynamic = dynamicCodeInvocations(sourceText, {exemptFunction, exemptInvocations, invocationSites});
  if (dynamic.length) failures.push({code: 'dynamic-code', message: `${label} must not use dynamic code (Invoke-Expression, iex, Add-Type, Invoke-Command, [scriptblock]::Create, InvokeScript, $ExecutionContext, string-named or non-literal & / . invocations) outside its listed exemptions: ${dynamic.join(', ')}`});
  const assignments = disallowedMemberAssignments(sourceText, {localRoots, comSetters});
  if (assignments.length) failures.push({code: 'com-property-assignment', message: `${label} assigns members outside its local report roots and documented COM setters: ${assignments.join(', ')}`});
  return failures;
}
