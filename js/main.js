// js/main.js

// ===== 多言語辞書 =====
const I18N = {
  ja: {
    title: 'エクセルクソ長数式が"少しだけ"読みやすくなる"かもしれない"ツール。',
    headerTitle: 'エクセルクソ長数式が"少しだけ"読みやすくなる"かもしれない"ツール。',
    inputLabel: '元の数式（Excel 風）',
    outputLabel: '整形結果',
    runButton: '整形する',
    hint: 'IF や FILTER などの関数を、引数ラベル付きのツリー構造で表示する簡易版デモ',
    errorPrefix: 'エラー: ',
    labelModeCaption: '引数ラベル表示:'
  },
  en: {
    title: 'A tool that *might* make disgusting long Excel formulas *slightly* more readable.',
    headerTitle: 'A tool that *might* make disgusting long Excel formulas *slightly* more readable.',
    inputLabel: 'Original formula (Excel-like)',
    outputLabel: 'Formatted result',
    runButton: 'Format',
    hint: 'Simple demo: show Excel-like formulas as labeled tree structures (IF, FILTER, etc.)',
    errorPrefix: 'Error: ',
    labelModeCaption: 'Argument labels:'
  }
};

let currentLang = 'ja';
let labelMode = 'auto'; // 'auto' | 'ja1' | 'ja2' | 'en' | 'off'
let ARG_LABELS = {};
let labelsLoaded = false;

// ===== 言語まわり =====

function detectInitialLang() {
  const stored = localStorage.getItem('formula_viewer_lang');
  if (stored && (stored === 'ja' || stored === 'en')) {
    return stored;
  }
  const nav = navigator.language || navigator.userLanguage || 'en';
  return nav.startsWith('ja') ? 'ja' : 'en';
}

function applyLang(lang) {
  const dict = I18N[lang] || I18N.ja;
  currentLang = lang;
  localStorage.setItem('formula_viewer_lang', lang);

  document.title = dict.title;
  document.getElementById('headerTitle').textContent = dict.headerTitle;
  document.getElementById('labelInput').textContent  = dict.inputLabel;
  document.getElementById('labelOutput').textContent = dict.outputLabel;
  document.getElementById('runButton').textContent   = dict.runButton;
  document.getElementById('hintText').textContent    = dict.hint;
  document.getElementById('labelModeCaption').textContent = dict.labelModeCaption;

  document.getElementById('btnJa').classList.toggle('active', lang === 'ja');
  document.getElementById('btnEn').classList.toggle('active', lang === 'en');
}

// ===== ラベルJSON読み込み =====

async function loadArgLabels() {
  const errorEl = document.getElementById('error');
  try {
    const res = await fetch('./labels/excel_formula_arg_labels_full_ja.json', {
      cache: 'no-store'
    });
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }
    const json = await res.json();
    ARG_LABELS = json || {};
    labelsLoaded = true;
  } catch (e) {
    labelsLoaded = false;
    console.error('ラベルJSONの読み込みに失敗:', e);
    const prefix = (I18N[currentLang] || I18N.ja).errorPrefix;
    errorEl.textContent = prefix + '引数ラベル辞書の読み込みに失敗しました（汎用ラベルで続行します）。';
  }
}

// ===== 関数引数ラベル取得 =====

function getArgLabel(funcName, index) {
  funcName = String(funcName || '').toUpperCase();

  if (labelMode === 'off') return '';

  const meta = ARG_LABELS[funcName];

  // 実際に使うキーを決める
  let key;
  if (labelMode === 'auto') {
    key = (currentLang === 'ja') ? 'ja1' : 'en';
  } else if (labelMode === 'ja1' || labelMode === 'ja2' || labelMode === 'en') {
    key = labelMode;
  } else {
    key = 'ja1';
  }

  const arr = meta && meta[key];
  if (arr && arr[index]) {
    return arr[index];
  }

  // IFだけは保険で手動定義
  if (funcName === 'IF') {
    if (key === 'en') {
      return ['logical_test', 'value_if_true', 'value_if_false'][index] || `arg${index+1}`;
    } else {
      return ['条件式', '真の場合', '偽の場合'][index]
          || (currentLang === 'ja' ? `引数${index+1}` : `arg${index+1}`);
    }
  }

  // その他は汎用ラベル
  if (currentLang === 'ja') {
    return `引数${index + 1}`;
  } else {
    return `arg${index + 1}`;
  }
}

// ===== トークナイザ =====

function tokenize(input) {
  let s = input.trim();
  if (s[0] === '=') s = s.slice(1);
  const tokens = [];
  let i = 0;

  const isLetter = c => /[A-Za-z_]/.test(c);
  const isDigit  = c => /[0-9]/.test(c);
  const isIdentChar = c => /[A-Za-z0-9_\$]/.test(c);

  while (i < s.length) {
    const c = s[i];

    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') {
      i++;
      continue;
    }

    if (isLetter(c) || c === '$') {
      let start = i;
      i++;
      while (i < s.length && isIdentChar(s[i])) i++;
      tokens.push({ type: 'ident', value: s.slice(start, i) });
      continue;
    }

    if (isDigit(c) || (c === '.' && isDigit(s[i+1] || ''))) {
      let start = i;
      i++;
      while (i < s.length && (isDigit(s[i]) || s[i] === '.')) i++;
      tokens.push({ type: 'number', value: s.slice(start, i) });
      continue;
    }

    if (c === '"') {
      i++;
      let str = '';
      while (i < s.length && s[i] !== '"') {
        str += s[i++];
      }
      if (s[i] === '"') i++;
      tokens.push({ type: 'string', value: str });
      continue;
    }

    if (c === ',' || c === '(' || c === ')') {
      tokens.push({ type: c, value: c });
      i++;
      continue;
    }

    const two = s.slice(i, i+2);
    if (['>=', '<=', '<>', '=='].includes(two)) {
      tokens.push({ type: 'op', value: two });
      i += 2;
      continue;
    }

    if ('=+-*/^<>'.includes(c)) {
      tokens.push({ type: 'op', value: c });
      i++;
      continue;
    }

    throw new Error('未知の文字: ' + c);
  }

  tokens.push({ type: 'eof', value: '' });
  return tokens;
}

// ===== パーサ =====

function parseFormula(tokens) {
  let pos = 0;
  const peek = () => tokens[pos];
  const consume = (type, value) => {
    const t = tokens[pos];
    if (type && t.type !== type) {
      throw new Error(`トークンタイプ不一致: 期待=${type} 実際=${t.type}`);
    }
    if (value && t.value !== value) {
      throw new Error(`トークン値不一致: 期待=${value} 実際=${t.value}`);
    }
    pos++;
    return t;
  };

  function parseExpression() {
    return parseComparison();
  }

  function parseComparison() {
    let node = parseAddSub();
    while (peek().type === 'op' && ['=','<>','<','>','<=','>='].includes(peek().value)) {
      const op = consume('op').value;
      const right = parseAddSub();
      node = { type: 'Binary', op, left: node, right };
    }
    return node;
  }

  function parseAddSub() {
    let node = parseMulDiv();
    while (peek().type === 'op' && ['+','-'].includes(peek().value)) {
      const op = consume('op').value;
      const right = parseMulDiv();
      node = { type: 'Binary', op, left: node, right };
    }
    return node;
  }

  function parseMulDiv() {
    let node = parsePrimary();
    while (peek().type === 'op' && ['*','/','^'].includes(peek().value)) {
      const op = consume('op').value;
      const right = parsePrimary();
      node = { type: 'Binary', op, left: node, right };
    }
    return node;
  }

  function parsePrimary() {
    const t = peek();

    if (t.type === 'number') {
      consume('number');
      return { type: 'Literal', kind: 'number', value: t.value };
    }
    if (t.type === 'string') {
      consume('string');
      return { type: 'Literal', kind: 'string', value: `"${t.value}"` };
    }
    if (t.type === 'ident') {
      const name = consume('ident').value;
      if (peek().type === '(') {
        consume('(');
        const args = [];
        if (peek().type !== ')') {
          while (true) {
            args.push(parseExpression());
            if (peek().type === ',') {
              consume(',');
              continue;
            }
            break;
          }
        }
        consume(')');
        return { type: 'Func', name, args };
      } else {
        return { type: 'Identifier', name };
      }
    }
    if (t.type === '(') {
      consume('(');
      const node = parseExpression();
      consume(')');
      return { type: 'Paren', inner: node };
    }

    throw new Error('予期せぬトークン: ' + t.type + ' ' + t.value);
  }

  const ast = parseExpression();
  if (peek().type !== 'eof') {
    throw new Error('式の後ろに余分なトークンがあります');
  }
  return ast;
}

// ===== AST → 文字列 =====

function exprToInlineString(node) {
  switch (node.type) {
    case 'Literal':
      return node.value;
    case 'Identifier':
      return node.name;
    case 'Paren':
      return '(' + exprToInlineString(node.inner) + ')';
    case 'Binary':
      return exprToInlineString(node.left) + ' ' + node.op + ' ' + exprToInlineString(node.right);
    case 'Func': {
      const args = node.args.map(exprToInlineString).join(', ');
      return node.name + '(' + args + ')';
    }
    default:
      return '?';
  }
}

function formatBinary(node, depth) {
  const indent = '  '.repeat(depth);
  const indentChild = '  '.repeat(depth + 1);
  const lines = [];
  lines.push(indent + '(');
  lines.push(formatNode(node.left, depth + 1));
  lines.push(indentChild + node.op);
  lines.push(formatNode(node.right, depth + 1));
  lines.push(indent + ')');
  return lines.join('\n');
}

function formatFuncNode(node, depth) {
  const funcName = node.name.toUpperCase();
  return formatLabeledFunc(funcName, node.args, depth);
}

function formatLabeledFunc(funcName, args, depth) {
  const indent = '  '.repeat(depth);
  const barPrefix = '| '.repeat(depth + 1);
  const labelWidth = 14;
  const lines = [];

  lines.push(indent + funcName + '(');

  args.forEach((arg, i) => {
    const labelRaw = getArgLabel(funcName, i);
    const label = labelRaw.padEnd(labelWidth, ' ');
    const formatted = formatNode(arg, depth + 1);
    const argLines = formatted.split('\n');

    // 1行目
    lines.push(
      barPrefix +
      label + '  ' +
      argLines[0].trimStart()
    );

    // 2行目以降
    for (let j = 1; j < argLines.length; j++) {
      lines.push(
        barPrefix +
        ' '.repeat(labelWidth) + '  ' +
        argLines[j]
      );
    }
  });

  lines.push(indent + ')');
  return lines.join('\n');
}

function formatNode(node, depth = 0) {
  switch (node.type) {
    case 'Func':
      return formatFuncNode(node, depth);
    case 'Binary':
      return formatBinary(node, depth);
    case 'Literal':
    case 'Identifier':
      return '  '.repeat(depth) + exprToInlineString(node);
    case 'Paren': {
      const inner = formatNode(node.inner, depth + 1);
      const indent = '  '.repeat(depth);
      return indent + '(\n' + inner + '\n' + indent + ')';
    }
    default:
      return '  '.repeat(depth) + exprToInlineString(node);
  }
}
.formula-wrapper {
  font-family: Consolas, "Courier New", monospace;
}

.formula-header {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  margin-bottom: 4px;
}

.formula-header button.toggle-btn {
  padding: 2px 6px;
  font-size: 11px;
  border-radius: 4px;
  border: 1px solid #555;
  background: #333;
  color: #eee;
  cursor: pointer;
}

.formula-header button.toggle-btn:hover {
  background: #444;
}

.formula-inline {
  white-space: nowrap;
  overflow-x: auto;
}

.formula-body {
  margin-left: 20px;
}

// ===== UIひもづけ =====

const inputEl  = document.getElementById('inputFormula');
const outputEl = document.getElementById('output');
const errorEl  = document.getElementById('error');
const runBtn   = document.getElementById('runButton');
const btnJa    = document.getElementById('btnJa');
const btnEn    = document.getElementById('btnEn');
const labelModeSelect = document.getElementById('labelModeSelect');

function runFormatter() {
  const src = inputEl.value;
  errorEl.textContent = '';
  try {
    const tokens = tokenize(src);
    const ast = parseFormula(tokens);
    const formatted = formatNode(ast, 0);
    outputEl.textContent = formatted;
  } catch (e) {
    outputEl.textContent = '';
    const prefix = (I18N[currentLang] || I18N.ja).errorPrefix;
    errorEl.textContent = prefix + e.message;
    console.error(e);
  }
}

runBtn.addEventListener('click', runFormatter);

btnJa.addEventListener('click', () => {
  applyLang('ja');
  runFormatter();
});
btnEn.addEventListener('click', () => {
  applyLang('en');
  runFormatter();
});

labelModeSelect.addEventListener('change', () => {
  labelMode = labelModeSelect.value;
  runFormatter();
});

window.addEventListener('DOMContentLoaded', async () => {
  const lang = detectInitialLang();
  applyLang(lang);
  await loadArgLabels(); // 失敗しても内部でエラー表示＋汎用ラベルで継続
  runFormatter();
});
