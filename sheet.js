// ═══════════════════════════════════════════════════════════
//  sheet.js  –  TheCFG Office / Sheet module
// ═══════════════════════════════════════════════════════════

const Sheet = (() => {
  // ── State ──────────────────────────────────────────────────
  let gridData    = [];   // Array<Array<{ raw: string, display: string }>>
  let focusedCell = null; // { row, col } | null

  const INIT_ROWS = 20;
  const INIT_COLS = 8;

  // ── DOM refs ───────────────────────────────────────────────
  let tbody, theadRow, cellRefEl, formulaInputEl;

  // ── Init ───────────────────────────────────────────────────
  function init() {
    tbody        = document.getElementById('sheet-tbody');
    theadRow     = document.getElementById('sheet-thead-row');
    cellRefEl    = document.getElementById('sheet-cell-ref');
    formulaInputEl = document.getElementById('sheet-formula-input');

    // Formula bar: commit on Enter / blur
    formulaInputEl.addEventListener('keydown', e => {
      if (e.key === 'Enter' && focusedCell) {
        const { row, col } = focusedCell;
        gridData[row][col].raw = formulaInputEl.value;
        gridData = recomputeDisplays(gridData);
        renderBody();
        // re-focus same cell
        const inp = getCellInput(row, col);
        if (inp) inp.focus();
      }
    });

    gridData = buildEmptyGrid(INIT_ROWS, INIT_COLS);
    renderAll();
  }

  // ── Grid helpers ───────────────────────────────────────────
  function buildEmptyGrid(rows, cols) {
    return Array.from({ length: rows }, () =>
      Array.from({ length: cols }, () => ({ raw: '', display: '' }))
    );
  }

  function numRows() { return gridData.length; }
  function numCols() { return gridData[0]?.length ?? 0; }

  // ── Formula engine ─────────────────────────────────────────
  function parseCellRef(ref) {
    const cm = ref.match(/[A-Za-z]+/);
    const rm = ref.match(/\d+/);
    if (!cm || !rm) throw new Error('Invalid ref: ' + ref);
    let col = 0;
    for (const ch of cm[0].toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64;
    col -= 1;
    return { col, row: parseInt(rm[0], 10) - 1 };
  }

  function parseRange(s) {
    const parts = s.split(':');
    const a = parseCellRef(parts[0]);
    const b = parts.length > 1 ? parseCellRef(parts[1]) : a;
    return {
      startRow: Math.min(a.row, b.row), endRow: Math.max(a.row, b.row),
      startCol: Math.min(a.col, b.col), endCol: Math.max(a.col, b.col)
    };
  }

  function getNumeric(arg, grid) {
    try {
      const { row, col } = parseCellRef(arg);
      const nr = grid.length, nc = grid[0]?.length ?? 0;
      if (row >= 0 && row < nr && col >= 0 && col < nc) {
        const v = parseFloat(evaluateCell(row, col, grid).display);
        return isNaN(v) ? 0 : v;
      }
      return 0;
    } catch {
      const n = parseFloat(arg);
      return isNaN(n) ? 0 : n;
    }
  }

  function evaluateCell(r, c, grid) {
    const nr = grid.length, nc = grid[0]?.length ?? 0;
    if (r < 0 || r >= nr || c < 0 || c >= nc) return { display: '#REF!' };
    const cell = grid[r]?.[c];
    if (!cell) return { display: '#REF!' };
    const raw = cell.raw;
    if (typeof raw !== 'string' || !raw.startsWith('=')) return { display: raw };

    try {
      const formula = raw.slice(1).trim();
      const upper   = formula.toUpperCase();

      if (upper.startsWith('SUM(') && upper.endsWith(')')) {
        const { startRow, endRow, startCol, endCol } = parseRange(formula.slice(4, -1));
        let sum = 0;
        for (let rr = startRow; rr <= endRow; rr++)
          for (let cc = startCol; cc <= endCol; cc++) {
            if (rr < nr && cc < nc) {
              const v = parseFloat(evaluateCell(rr, cc, grid).display);
              if (!isNaN(v)) sum += v;
            }
          }
        return { display: String(sum) };
      }

      if (upper.startsWith('AVERAGE(') && upper.endsWith(')')) {
        const { startRow, endRow, startCol, endCol } = parseRange(formula.slice(8, -1));
        let sum = 0, count = 0;
        for (let rr = startRow; rr <= endRow; rr++)
          for (let cc = startCol; cc <= endCol; cc++) {
            if (rr < nr && cc < nc) {
              const v = parseFloat(evaluateCell(rr, cc, grid).display);
              if (!isNaN(v)) { sum += v; count++; }
            }
          }
        return { display: count === 0 ? '#DIV/0!' : String(sum / count) };
      }

      if (upper.startsWith('COUNT(') && upper.endsWith(')')) {
        const { startRow, endRow, startCol, endCol } = parseRange(formula.slice(6, -1));
        let count = 0;
        for (let rr = startRow; rr <= endRow; rr++)
          for (let cc = startCol; cc <= endCol; cc++) {
            if (rr < nr && cc < nc) {
              const v = parseFloat(evaluateCell(rr, cc, grid).display);
              if (!isNaN(v)) count++;
            }
          }
        return { display: String(count) };
      }

      if (upper.startsWith('MAX(') && upper.endsWith(')')) {
        const { startRow, endRow, startCol, endCol } = parseRange(formula.slice(4, -1));
        let max = -Infinity, found = false;
        for (let rr = startRow; rr <= endRow; rr++)
          for (let cc = startCol; cc <= endCol; cc++) {
            if (rr < nr && cc < nc) {
              const v = parseFloat(evaluateCell(rr, cc, grid).display);
              if (!isNaN(v)) { max = Math.max(max, v); found = true; }
            }
          }
        return { display: found ? String(max) : '' };
      }

      if (upper.startsWith('MIN(') && upper.endsWith(')')) {
        const { startRow, endRow, startCol, endCol } = parseRange(formula.slice(4, -1));
        let min = Infinity, found = false;
        for (let rr = startRow; rr <= endRow; rr++)
          for (let cc = startCol; cc <= endCol; cc++) {
            if (rr < nr && cc < nc) {
              const v = parseFloat(evaluateCell(rr, cc, grid).display);
              if (!isNaN(v)) { min = Math.min(min, v); found = true; }
            }
          }
        return { display: found ? String(min) : '' };
      }

      const twoArg = upper.match(/^(MINUS|TIME|DIVISION)\(/);
      if (twoArg && upper.endsWith(')')) {
        const fn      = twoArg[1];
        const argsStr = formula.slice(fn.length + 1, -1);
        const args    = argsStr.split(',').map(a => a.trim());
        if (args.length !== 2) return { display: '#VALUE!' };
        const v1 = getNumeric(args[0], grid);
        const v2 = getNumeric(args[1], grid);
        if (fn === 'MINUS')    return { display: String(v1 - v2) };
        if (fn === 'TIME')     return { display: String(v1 * v2) };
        if (fn === 'DIVISION') return v2 === 0 ? { display: '#DIV/0!' } : { display: String(v1 / v2) };
      }

      // Simple arithmetic expression (no cell refs)
      if (/^[\d\s\+\-\*\/\(\)\.]+$/.test(formula)) {
        // eslint-disable-next-line no-eval
        const result = Function('"use strict"; return (' + formula + ')')();
        return { display: String(result) };
      }

      return { display: '#NAME?' };
    } catch {
      return { display: '#ERROR!' };
    }
  }

  function recomputeDisplays(grid) {
    return grid.map((row, r) =>
      row.map((cell, c) => ({ raw: cell.raw, display: evaluateCell(r, c, grid).display }))
    );
  }

  // ── Render ─────────────────────────────────────────────────
  function renderAll() { renderHeader(); renderBody(); }

  function renderHeader() {
    while (theadRow.children.length > 1) theadRow.removeChild(theadRow.lastChild);
    for (let c = 0; c < numCols(); c++) {
      const th = document.createElement('th');
      th.className = 'col-header';
      th.textContent = colLabel(c);
      theadRow.appendChild(th);
    }
  }

  function renderBody() {
    tbody.innerHTML = '';
    gridData.forEach((row, r) => {
      const tr = document.createElement('tr');
      const tdNum = document.createElement('td');
      tdNum.className = 'row-num';
      tdNum.textContent = r + 1;
      tr.appendChild(tdNum);

      row.forEach((cell, c) => {
        const td  = document.createElement('td');
        const inp = document.createElement('input');
        inp.type = 'text';
        inp.spellcheck = false;
        inp.value = (focusedCell?.row === r && focusedCell?.col === c) ? cell.raw : cell.display;

        inp.addEventListener('focus', () => {
          focusedCell = { row: r, col: c };
          inp.value   = gridData[r][c].raw;
          cellRefEl.value    = colLabel(c) + (r + 1);
          formulaInputEl.value = gridData[r][c].raw;
        });

        inp.addEventListener('blur', () => {
          gridData[r][c].raw = inp.value;
          gridData = recomputeDisplays(gridData);
          focusedCell = null;
          inp.value = gridData[r][c].display;
          refreshOtherDisplays(r, c);
        });

        inp.addEventListener('input', () => {
          gridData[r][c].raw = inp.value;
          formulaInputEl.value = inp.value;
          gridData = recomputeDisplays(gridData);
          refreshOtherDisplays(r, c);
        });

        td.appendChild(inp);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  function refreshOtherDisplays(editR, editC) {
    const rows = tbody.querySelectorAll('tr');
    rows.forEach((tr, r) => {
      const inputs = tr.querySelectorAll('input');
      inputs.forEach((inp, c) => {
        if (r === editR && c === editC) return;
        if (focusedCell?.row === r && focusedCell?.col === c) return;
        inp.value = gridData[r][c].display;
      });
    });
  }

  function getCellInput(r, c) {
    const tr = tbody.querySelectorAll('tr')[r];
    return tr?.querySelectorAll('input')[c] ?? null;
  }

  function colLabel(c) {
    let label = '';
    let n = c + 1;
    while (n > 0) {
      label = String.fromCharCode(64 + (n % 26 || 26)) + label;
      n = Math.floor((n - 1) / 26);
    }
    return label;
  }

  // ── Grid mutations ─────────────────────────────────────────
  function addRow() {
    gridData.push(Array.from({ length: numCols() }, () => ({ raw: '', display: '' })));
    gridData = recomputeDisplays(gridData);
    renderBody();
    updateSheetStatus();
  }

  function addColumn() {
    gridData = gridData.map(row => [...row, { raw: '', display: '' }]);
    gridData = recomputeDisplays(gridData);
    renderAll();
    updateSheetStatus();
  }

  function updateSheetStatus() {
    const el = document.getElementById('sheet-status');
    if (el) el.textContent = `${numRows()} 행 × ${numCols()} 열`;
  }

  // ── Save / Load ────────────────────────────────────────────
  function saveSheet() {
    const payload = gridData.map(row => row.map(cell => ({ raw: cell.raw })));
    const blob    = new Blob([JSON.stringify(payload)], { type: 'application/json' });
    const a       = document.createElement('a');
    a.href        = URL.createObjectURL(blob);
    a.download    = 'spreadsheet.json';
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    showSheetMsg('시트가 저장되었습니다.');
  }

  function loadSheet() {
    document.getElementById('sheet-file-input').click();
  }

  function handleFileLoad(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const loaded = JSON.parse(e.target.result);
        gridData = recomputeDisplays(
          loaded.map(row => row.map(cell => ({ raw: cell.raw, display: '' })))
        );
        renderAll();
        updateSheetStatus();
        showSheetMsg('시트가 불러와졌습니다.');
      } catch {
        showSheetMsg('파일 형식 오류: 올바른 JSON 파일인지 확인하세요.');
      }
    };
    reader.readAsText(file);
    event.target.value = '';
  }

  function showSheetMsg(msg) {
    Word.showMsg(msg); // reuse the shared message modal
  }

  // ── Public API ─────────────────────────────────────────────
  return { init, addRow, addColumn, saveSheet, loadSheet, handleFileLoad, updateSheetStatus };
})();
