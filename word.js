// ═══════════════════════════════════════════════════════════
//  word.js  –  TheCFG Office / Word module
//
//  Phase 1 (코어 엔진 안정화):
//   - 커스텀 실행취소/다시실행 스택 (execCommand undo/redo 미사용)
//   - 붙여넣기 HTML sanitize (외부 서식 정리, 위험 태그/속성 제거)
//   - 키보드 단축키 (Ctrl/Cmd + B/I/U/S/Z/Shift+Z/Y)
//   - 툴바 버튼이 현재 커서 위치의 서식 상태를 반영 (active 표시)
// ═══════════════════════════════════════════════════════════

const Word = (() => {
  // ── DOM ────────────────────────────────────────────────────
  let editor, charCountEl, wordCountEl, toolbarEl;
  let undoBtn, redoBtn, fontNameSelect, fontSizeSelect;

  // ── State ──────────────────────────────────────────────────
  let currentFilename = '';

  // 툴바 토글 버튼과 동기화할 execCommand 상태 목록
  const TOGGLE_COMMANDS = [
    'bold', 'italic', 'underline', 'strikeThrough',
    'superscript', 'subscript',
    'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull',
    'insertUnorderedList', 'insertOrderedList',
  ];

  // ── 실행취소 / 다시실행 히스토리 ────────────────────────────
  const HISTORY_LIMIT = 100;
  const TYPING_GROUP_MS = 600; // 이 시간 안의 연속 타이핑은 하나의 undo 단계로 묶음
  let undoStack = [];
  let redoStack = [];
  let lastInputType = null;
  let lastSnapshotTime = 0;

  function resetHistory() {
    undoStack = [];
    redoStack = [];
    lastInputType = null;
    lastSnapshotTime = 0;
    updateHistoryButtons();
  }

  function captureSnapshot() {
    return { html: editor.innerHTML, caret: getCaretOffset(editor) };
  }

  function pushUndo() {
    undoStack.push(captureSnapshot());
    if (undoStack.length > HISTORY_LIMIT) undoStack.shift();
    redoStack = []; // 새 변경이 생기면 redo 스택은 무효
    updateHistoryButtons();
  }

  function undo() {
    if (undoStack.length === 0) return;
    redoStack.push(captureSnapshot());
    const prev = undoStack.pop();
    editor.innerHTML = prev.html;
    setCaretOffset(editor, prev.caret);
    afterContentChange();
  }

  function redo() {
    if (redoStack.length === 0) return;
    undoStack.push(captureSnapshot());
    const next = redoStack.pop();
    editor.innerHTML = next.html;
    setCaretOffset(editor, next.caret);
    afterContentChange();
  }

  function updateHistoryButtons() {
    if (undoBtn) undoBtn.disabled = undoStack.length === 0;
    if (redoBtn) redoBtn.disabled = redoStack.length === 0;
  }

  function afterContentChange() {
    editor.focus();
    updateStats();
    updateHistoryButtons();
    syncToolbarState();
  }

  // ── 캐럿(커서) 오프셋 — contenteditable 내 텍스트 기준 문자 위치 ──
  function getCaretOffset(root) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return 0;
    const range = sel.getRangeAt(0);
    if (!root.contains(range.startContainer)) return 0;
    const preRange = range.cloneRange();
    preRange.selectNodeContents(root);
    preRange.setEnd(range.startContainer, range.startOffset);
    return preRange.toString().length;
  }

  function setCaretOffset(root, offset) {
    const sel = window.getSelection();
    const range = document.createRange();
    let remaining = offset;
    let found = false;

    (function walk(node) {
      if (found) return;
      if (node.nodeType === Node.TEXT_NODE) {
        const len = node.textContent.length;
        if (remaining <= len) {
          range.setStart(node, remaining);
          range.collapse(true);
          found = true;
        } else {
          remaining -= len;
        }
      } else {
        for (let i = 0; i < node.childNodes.length && !found; i++) walk(node.childNodes[i]);
      }
    })(root);

    if (!found) {
      range.selectNodeContents(root);
      range.collapse(false);
    }
    sel.removeAllRanges();
    sel.addRange(range);
  }

  // ── Init ───────────────────────────────────────────────────
  function init() {
    editor         = document.getElementById('word-editor');
    charCountEl    = document.getElementById('word-char-count');
    wordCountEl    = document.getElementById('word-word-count');
    toolbarEl      = document.getElementById('word-toolbar');
    undoBtn        = document.getElementById('word-undo-btn');
    redoBtn        = document.getElementById('word-redo-btn');
    fontNameSelect = document.getElementById('word-font-name');
    fontSizeSelect = document.getElementById('word-font-size');

    editor.addEventListener('input', updateStats);
    editor.addEventListener('keyup', updateStats);

    // 타이핑/삭제를 적당한 단위로 묶어 undo 스냅샷 기록
    editor.addEventListener('beforeinput', e => {
      const type = e.inputType;
      const now  = Date.now();
      const isTyping = type === 'insertText' || type === 'deleteContentBackward' || type === 'deleteContentForward';
      const shouldSnapshot = !isTyping || type !== lastInputType || (now - lastSnapshotTime) > TYPING_GROUP_MS;
      if (shouldSnapshot) pushUndo();
      lastInputType    = type;
      lastSnapshotTime = now;
    });

    // 붙여넣기: 외부 서식을 정리한 뒤 안전하게 삽입
    editor.addEventListener('paste', handlePaste);

    // 단축키
    editor.addEventListener('keydown', handleKeydown);

    // 서식 상태 동기화 (커서 이동, 클릭, 키 입력 등)
    editor.addEventListener('keyup', syncToolbarState);
    editor.addEventListener('mouseup', syncToolbarState);
    document.addEventListener('selectionchange', () => {
      if (document.activeElement === editor) syncToolbarState();
    });

    resetHistory();
    updateStats();
    syncToolbarState();
  }

  // ── Stats ──────────────────────────────────────────────────
  function updateStats() {
    const text  = editor.innerText || '';
    const chars = text.length;
    const words = text.trim() === '' ? 0 : text.trim().split(/\s+/).length;
    if (charCountEl) charCountEl.textContent = chars.toLocaleString();
    if (wordCountEl) wordCountEl.textContent = words.toLocaleString();
  }

  // ── 툴바 상태 동기화 ──────────────────────────────────────
  function syncToolbarState() {
    if (!toolbarEl) return;

    TOGGLE_COMMANDS.forEach(cmd => {
      const btn = toolbarEl.querySelector(`[data-cmd="${cmd}"]`);
      if (!btn) return;
      let isActive = false;
      try { isActive = document.queryCommandState(cmd); } catch { /* 일부 브라우저 미지원 명령 무시 */ }
      btn.classList.toggle('active', isActive);
    });

    if (fontNameSelect) {
      try {
        const raw = (document.queryCommandValue('fontName') || '').replace(/^["']|["']$/g, '');
        const match = [...fontNameSelect.options].find(o => o.value.toLowerCase() === raw.toLowerCase());
        if (match) fontNameSelect.value = match.value;
      } catch { /* ignore */ }
    }
    if (fontSizeSelect) {
      try {
        const size = document.queryCommandValue('fontSize');
        if (size) fontSizeSelect.value = size;
      } catch { /* ignore */ }
    }
  }

  // ── 단축키 ─────────────────────────────────────────────────
  function handleKeydown(e) {
    const ctrl = e.ctrlKey || e.metaKey;
    if (!ctrl) return;

    switch (e.key.toLowerCase()) {
      case 'b': e.preventDefault(); formatDoc('bold'); break;
      case 'i': e.preventDefault(); formatDoc('italic'); break;
      case 'u': e.preventDefault(); formatDoc('underline'); break;
      case 's': e.preventDefault(); showSaveDialog(); break;
      case 'z':
        e.preventDefault();
        if (e.shiftKey) redo(); else undo();
        break;
      case 'y': e.preventDefault(); redo(); break;
    }
  }

  // ── 붙여넣기 정리(sanitize) ───────────────────────────────
  const ALLOWED_TAGS = [
    'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'P', 'BR',
    'UL', 'OL', 'LI', 'A', 'SPAN', 'DIV',
    'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
    'TABLE', 'THEAD', 'TBODY', 'TR', 'TD', 'TH',
    'IMG', 'SUB', 'SUP',
  ];
  const ALLOWED_ATTRS = {
    A: ['href'], IMG: ['src', 'alt'],
    SPAN: ['style'], DIV: ['style'], P: ['style'],
    TD: ['colspan', 'rowspan'], TH: ['colspan', 'rowspan'],
  };
  const ALLOWED_STYLE_PROPS = [
    'color', 'background-color', 'font-weight', 'font-style',
    'text-decoration', 'font-size', 'text-align',
  ];

  function sanitizeStyle(styleStr) {
    return styleStr.split(';')
      .map(s => s.trim())
      .filter(Boolean)
      .filter(rule => ALLOWED_STYLE_PROPS.includes(rule.split(':')[0].trim().toLowerCase()))
      .join('; ');
  }

  function sanitizeAttributes(el) {
    const allowed = ALLOWED_ATTRS[el.tagName] || [];
    Array.from(el.attributes).forEach(attr => {
      const name = attr.name.toLowerCase();
      if (!allowed.includes(name)) { el.removeAttribute(attr.name); return; }
      if ((name === 'href' || name === 'src') && /^\s*javascript:/i.test(attr.value)) {
        el.removeAttribute(attr.name); return;
      }
      if (name === 'style') el.setAttribute('style', sanitizeStyle(attr.value));
    });
    if (el.tagName === 'A') {
      el.setAttribute('target', '_blank');
      el.setAttribute('rel', 'noopener noreferrer');
    }
  }

  // 허용되지 않은 태그는 벗겨내고(unwrap) 내용은 보존, 허용된 태그는 속성만 정리
  function cleanNode(parent) {
    Array.from(parent.childNodes).forEach(child => {
      if (child.nodeType === Node.TEXT_NODE) return;
      if (child.nodeType !== Node.ELEMENT_NODE) { parent.removeChild(child); return; }

      if (!ALLOWED_TAGS.includes(child.tagName)) {
        while (child.firstChild) parent.insertBefore(child.firstChild, child);
        parent.removeChild(child);
        cleanNode(parent); // unwrap으로 새로 올라온 자식들도 마저 정리
        return;
      }

      sanitizeAttributes(child);
      cleanNode(child);
    });
  }

  function sanitizeHtml(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    cleanNode(doc.body);
    return doc.body.innerHTML;
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // execCommand('insertHTML') 대신 Range API로 직접 삽입
  function insertHtmlAtCaret(html) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const fragment = range.createContextualFragment(html);
    const lastNode = fragment.lastChild;
    range.insertNode(fragment);
    if (lastNode) {
      range.setStartAfter(lastNode);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }

  function handlePaste(e) {
    e.preventDefault();
    pushUndo();

    const cd = e.clipboardData || window.clipboardData;
    const html = cd.getData('text/html');
    const plain = cd.getData('text/plain');

    const cleaned = html
      ? sanitizeHtml(html)
      : escapeHtml(plain).replace(/\r\n|\r|\n/g, '<br>');

    insertHtmlAtCaret(cleaned);
    afterContentChange();
  }

  // ── Format ─────────────────────────────────────────────────
  function formatDoc(command, value = null) {
    if (command === 'undo') { undo(); return; }
    if (command === 'redo') { redo(); return; }

    pushUndo();
    document.execCommand(command, false, value);
    afterContentChange();
  }

  // ── Document actions ───────────────────────────────────────
  function clearDocument() {
    if (confirm('작성 중인 모든 내용이 삭제됩니다. 새로 시작하시겠습니까?')) {
      editor.innerHTML = '';
      resetHistory();
      editor.focus();
      updateStats();
      syncToolbarState();
    }
  }

  // ── Save dialog ────────────────────────────────────────────
  function showSaveDialog() {
    const base = currentFilename.replace(/\.(html|txt)$/i, '') || '새 문서';
    const inp  = document.getElementById('word-filename-input');
    inp.value  = base;

    document.getElementById('word-save-modal').classList.remove('hidden');
    document.getElementById('word-filename-section').classList.remove('hidden');
    document.getElementById('word-txt-warning').classList.add('hidden');
    inp.focus(); inp.select();
  }

  function hideSaveDialog() {
    document.getElementById('word-save-modal').classList.add('hidden');
    document.getElementById('word-dont-show-txt').checked = false;
  }

  function handleSaveChoice(type) {
    const filename = document.getElementById('word-filename-input').value.trim();
    if (!filename) { showMsg('파일 이름을 입력해주세요.'); return; }
    currentFilename = filename;

    if (type === 'html') {
      saveFile(currentFilename, 'html');
      hideSaveDialog();
    } else {
      if (localStorage.getItem('dontShowTxtWarning') === 'true') {
        saveFile(currentFilename, 'txt');
        hideSaveDialog();
      } else {
        document.getElementById('word-filename-section').classList.add('hidden');
        document.getElementById('word-txt-warning').classList.remove('hidden');
      }
    }
  }

  function confirmTxtSave() {
    if (document.getElementById('word-dont-show-txt').checked) {
      localStorage.setItem('dontShowTxtWarning', 'true');
    }
    saveFile(currentFilename, 'txt');
    hideSaveDialog();
  }

  function saveFile(filename, type) {
    let content, mimeType, finalName = filename;
    if (type === 'html') {
      content  = editor.innerHTML;
      mimeType = 'text/html;charset=utf-8';
      if (!finalName.toLowerCase().endsWith('.html')) finalName += '.html';
    } else {
      content  = editor.innerText;
      mimeType = 'text/plain;charset=utf-8';
      if (!finalName.toLowerCase().endsWith('.txt')) finalName += '.txt';
    }
    const blob = new Blob([content], { type: mimeType });
    const a    = document.createElement('a');
    a.href     = URL.createObjectURL(blob);
    a.download = finalName;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    currentFilename = finalName;
    showMsg(`"${finalName}"(으)로 저장되었습니다.`);
  }

  // ── Open file ──────────────────────────────────────────────
  function openFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith('.html') && !lower.endsWith('.txt')) {
      showMsg('텍스트(.txt) 또는 HTML(.html) 파일만 열 수 있습니다.');
      event.target.value = ''; return;
    }
    const reader = new FileReader();
    reader.onload = e => {
      if (lower.endsWith('.html')) editor.innerHTML = e.target.result;
      else editor.innerText = e.target.result;
      currentFilename = file.name;
      resetHistory();
      updateStats();
      syncToolbarState();
      showMsg(`"${file.name}" 파일이 로드되었습니다.`);
    };
    reader.onerror = () => showMsg('파일을 읽는 중 오류가 발생했습니다.');
    reader.readAsText(file);
    event.target.value = '';
  }

  // ── Message ────────────────────────────────────────────────
  function showMsg(msg) {
    document.getElementById('word-msg-text').textContent = msg;
    document.getElementById('word-msg-modal').classList.remove('hidden');
  }

  function hideMsg() {
    document.getElementById('word-msg-modal').classList.add('hidden');
  }

  // ── Public API ─────────────────────────────────────────────
  return { init, formatDoc, clearDocument, showSaveDialog, hideSaveDialog,
           handleSaveChoice, confirmTxtSave, openFile, showMsg, hideMsg };
})();