// ═══════════════════════════════════════════════════════════
//  word.js  –  TheCFG Office / Word module
// ═══════════════════════════════════════════════════════════

const Word = (() => {
  // ── DOM ────────────────────────────────────────────────────
  let editor, charCountEl, wordCountEl;

  // ── State ──────────────────────────────────────────────────
  let currentFilename = '';

  // ── Init ───────────────────────────────────────────────────
  function init() {
    editor      = document.getElementById('word-editor');
    charCountEl = document.getElementById('word-char-count');
    wordCountEl = document.getElementById('word-word-count');

    editor.addEventListener('input', updateStats);
    editor.addEventListener('keyup', updateStats);
    updateStats();
  }

  // ── Stats ──────────────────────────────────────────────────
  function updateStats() {
    const text  = editor.innerText || '';
    const chars = text.length;
    const words = text.trim() === '' ? 0 : text.trim().split(/\s+/).length;
    if (charCountEl) charCountEl.textContent = chars.toLocaleString();
    if (wordCountEl) wordCountEl.textContent = words.toLocaleString();
  }

  // ── Format ─────────────────────────────────────────────────
  function formatDoc(command, value = null) {
    document.execCommand(command, false, value);
    editor.focus();
    updateStats();
  }

  // ── Document actions ───────────────────────────────────────
  function clearDocument() {
    if (confirm('작성 중인 모든 내용이 삭제됩니다. 새로 시작하시겠습니까?')) {
      editor.innerHTML = '';
      editor.focus();
      updateStats();
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
      updateStats();
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
