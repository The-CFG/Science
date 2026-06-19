// ═══════════════════════════════════════════════════════════
//  app.js  –  TheCFG Office / main
// ═══════════════════════════════════════════════════════════

// ── Tab switching ───────────────────────────────────────────
// New apps don't need any changes here: as long as the panel id is
// "{name}-panel", the tab button has data-app="{name}", and (optionally)
// a top action group has class="top-action-group" data-app="{name}",
// switchTab() will wire them up automatically.
function switchTab(name) {
  document.querySelectorAll('.app-tab').forEach(t => t.classList.toggle('active', t.dataset.app === name));
  document.querySelectorAll('.app-panel').forEach(p => p.classList.toggle('active', p.id === `${name}-panel`));
  document.querySelectorAll('.top-action-group').forEach(g => {
    g.style.display = (g.dataset.app === name) ? '' : 'none';
  });

  // Update status bar
  document.getElementById('word-statusbar').classList.toggle('hidden', name !== 'word');
  document.getElementById('sheet-statusbar').classList.toggle('hidden', name !== 'sheet');
  document.getElementById('htmlblast-statusbar').classList.toggle('hidden', name !== 'htmlblast');
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

// HTMLBlast
function hbRun()             { HtmlBlast.run(); }
function hbSave()            { HtmlBlast.save(); }
function hbLoad()            { HtmlBlast.load(); }

// ── Boot ────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  Word.init();
  Sheet.init();
  Sheet.updateSheetStatus();
  HtmlBlast.init();
  switchTab('home'); // default: app selection home screen
});