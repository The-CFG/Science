// ═══════════════════════════════════════════════════════════
//  app.js  –  TheCFG Office / main
// ═══════════════════════════════════════════════════════════

// ── Tab switching ───────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll('.app-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.app-panel').forEach(p => p.classList.remove('active'));

  document.querySelector(`.app-tab[data-app="${name}"]`).classList.add('active');
  document.getElementById(`${name}-panel`).classList.add('active');

  // Update status bar
  document.getElementById('word-statusbar').classList.toggle('hidden', name !== 'word');
  document.getElementById('sheet-statusbar').classList.toggle('hidden', name !== 'sheet');
}

// ── Expose to inline onclick handlers ──────────────────────
// Word
function wFormat(cmd, val)   { Word.formatDoc(cmd, val); }
function wClear()            { Word.clearDocument(); }
function wSave()             { Word.showSaveDialog(); }
function wHideSave()         { Word.hideSaveDialog(); }
function wSaveChoice(t)      { Word.handleSaveChoice(t); }
function wConfirmTxt()       { Word.confirmTxtSave(); }
function wOpenFile(e)        { Word.openFile(e); }
function wHideMsg()          { Word.hideMsg(); }

// Sheet
function sAddRow()           { Sheet.addRow(); }
function sAddCol()           { Sheet.addColumn(); }
function sSave()             { Sheet.saveSheet(); }
function sLoad()             { Sheet.loadSheet(); }
function sFileLoad(e)        { Sheet.handleFileLoad(e); }

// ── Boot ────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  Word.init();
  Sheet.init();
  Sheet.updateSheetStatus();
  switchTab('word'); // default
});
