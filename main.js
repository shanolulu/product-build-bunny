/* global XLSX */

// ─── State ──────────────────────────────────────────────────
const state = {
  left:  { workbook: null, sheetName: null, data: [], fileName: '' },
  right: { workbook: null, sheetName: null, data: [], fileName: '' },
  undoStack: [],
};

// ─── Setup ──────────────────────────────────────────────────
function setupDropZone(side) {
  const dropZone  = document.getElementById(`drop-${side}`);
  const fileInput = document.getElementById(`file-${side}`);

  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) loadFile(side, file);
  });
  // 클릭으로 drop-zone 전체에서 파일 선택 가능
  dropZone.addEventListener('click', e => {
    if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'LABEL') {
      fileInput.click();
    }
  });
  fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) loadFile(side, file);
    fileInput.value = '';
  });
}

// ─── File Loading ────────────────────────────────────────────
function loadFile(side, file) {
  const reader = new FileReader();
  reader.onload = e => parseFile(side, e.target.result, file.name);
  reader.readAsArrayBuffer(file);
}

function parseFile(side, buffer, fileName) {
  const workbook  = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const data      = sheetToData(workbook, sheetName);

  state[side].workbook  = workbook;
  state[side].sheetName = sheetName;
  state[side].fileName  = fileName;
  state[side].data      = data;

  // 파일 이름 표시
  document.getElementById(`title-${side}`).textContent = fileName;

  // 드롭존 숨기고 테이블/푸터 표시
  document.getElementById(`drop-${side}`).classList.add('hidden');
  document.getElementById(`table-${side}`).classList.remove('hidden');
  document.getElementById(`footer-${side}`).classList.remove('hidden');

  renderSheetTabs(side, workbook.SheetNames);
  refreshView();
}

function sheetToData(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
}

// ─── Sheet Tabs ──────────────────────────────────────────────
function renderSheetTabs(side, sheetNames) {
  const tabsEl = document.getElementById(`tabs-${side}`);
  tabsEl.innerHTML = '';

  if (sheetNames.length <= 1) {
    tabsEl.classList.add('hidden');
    return;
  }
  tabsEl.classList.remove('hidden');

  sheetNames.forEach(name => {
    const tab = document.createElement('button');
    tab.className = 'sheet-tab' + (name === state[side].sheetName ? ' active' : '');
    tab.textContent = name;
    tab.addEventListener('click', () => {
      state[side].sheetName = name;
      state[side].data      = sheetToData(state[side].workbook, name);
      renderSheetTabs(side, sheetNames);
      refreshView();
    });
    tabsEl.appendChild(tab);
  });
}

// ─── Diff ────────────────────────────────────────────────────
function buildDiffMap() {
  const L = state.left.data;
  const R = state.right.data;
  const maxRows = Math.max(L.length, R.length);

  const leftMap  = [];
  const rightMap = [];

  for (let r = 0; r < maxRows; r++) {
    const lRow = L[r] ?? null;
    const rRow = R[r] ?? null;

    if (!lRow && rRow) {
      const cols = rRow.length;
      leftMap.push(new Array(cols).fill('placeholder'));
      rightMap.push(new Array(cols).fill('added'));
    } else if (lRow && !rRow) {
      const cols = lRow.length;
      leftMap.push(new Array(cols).fill('deleted'));
      rightMap.push(new Array(cols).fill('placeholder'));
    } else {
      const maxCols = Math.max(lRow.length, rRow.length);
      const lDiff = [];
      const rDiff = [];
      for (let c = 0; c < maxCols; c++) {
        const lVal = lRow[c] !== undefined ? String(lRow[c]) : '';
        const rVal = rRow[c] !== undefined ? String(rRow[c]) : '';
        const diff = lVal !== rVal ? 'modified' : 'same';
        lDiff.push(diff);
        rDiff.push(diff);
      }
      leftMap.push(lDiff);
      rightMap.push(rDiff);
    }
  }
  return { left: leftMap, right: rightMap };
}

// ─── Render ──────────────────────────────────────────────────
function refreshView() {
  const bothLoaded = state.left.data.length > 0 && state.right.data.length > 0;

  if (bothLoaded) {
    const diffMap = buildDiffMap();
    renderTableWithDiff('left',  diffMap);
    renderTableWithDiff('right', diffMap);
    updateSummary(diffMap);
    document.getElementById('header-hint').classList.remove('hidden');
  } else {
    if (state.left.data.length  > 0) renderTablePlain('left');
    if (state.right.data.length > 0) renderTablePlain('right');
    document.getElementById('summary').classList.add('hidden');
    document.getElementById('header-hint').classList.add('hidden');
  }
}

function renderTablePlain(side) {
  const container = document.getElementById(`table-${side}`);
  const data = state[side].data;
  if (!data || data.length === 0) { container.innerHTML = '<p style="padding:20px;color:#94a3b8">데이터 없음</p>'; return; }
  const maxCols = Math.max(...data.map(r => r ? r.length : 0), 1);
  let html = '<table><tbody>';
  data.forEach(row => {
    html += '<tr>';
    for (let c = 0; c < maxCols; c++) {
      const v = row && row[c] !== undefined ? row[c] : '';
      html += `<td>${esc(String(v))}</td>`;
    }
    html += '</tr>';
  });
  html += '</tbody></table>';
  container.innerHTML = html;
}

function renderTableWithDiff(side, diffMap) {
  const container = document.getElementById(`table-${side}`);
  const data      = state[side].data;
  const diffs     = diffMap[side];
  const maxRows   = diffs.length;
  const maxCols   = Math.max(
    ...state.left.data.map(r => r ? r.length : 0),
    ...state.right.data.map(r => r ? r.length : 0),
    1
  );

  let html = '<table><tbody>';

  for (let r = 0; r < maxRows; r++) {
    const rowData = data[r] || [];
    const rowDiff = diffs[r] || [];
    const firstDiff = rowDiff[0];

    if (firstDiff === 'placeholder') {
      html += '<tr class="row-placeholder">';
      for (let c = 0; c < maxCols; c++) html += '<td class="cell-placeholder"></td>';
      html += '</tr>';
      continue;
    }

    const rowClass = firstDiff === 'added'   ? 'row-added'
                   : firstDiff === 'deleted' ? 'row-deleted'
                   : '';
    html += `<tr class="${rowClass}">`;

    for (let c = 0; c < maxCols; c++) {
      const v        = rowData[c] !== undefined ? rowData[c] : '';
      const cellDiff = rowDiff[c] || 'same';

      let cls   = '';
      let attrs = '';

      if (firstDiff === 'added')   cls = 'cell-added';
      else if (firstDiff === 'deleted') cls = 'cell-deleted';
      else if (cellDiff === 'modified') {
        cls   = 'cell-modified';
        attrs = ` data-row="${r}" data-col="${c}" data-side="${side}"`;
      }

      html += `<td class="${cls}"${attrs}>${esc(String(v))}</td>`;
    }
    html += '</tr>';
  }

  html += '</tbody></table>';
  container.innerHTML = html;

  container.querySelectorAll('[data-side]').forEach(cell => {
    cell.addEventListener('click', handleCellClick);
  });
}

function updateSummary(diffMap) {
  let modified = 0, added = 0, deleted = 0;
  diffMap.left.forEach((row, r) => {
    const f = row[0];
    if (f === 'deleted')     deleted++;
    else if (diffMap.right[r]?.[0] === 'added') added++;
    else row.forEach(c => { if (c === 'modified') modified++; });
  });

  document.getElementById('summary').classList.remove('hidden');
  document.getElementById('sum-modified').textContent = `수정 ${modified}셀`;
  document.getElementById('sum-added').textContent    = `추가 ${added}행`;
  document.getElementById('sum-deleted').textContent  = `삭제 ${deleted}행`;
}

// ─── Cell Click → Apply ──────────────────────────────────────
let ctxTarget = null;

function handleCellClick(e) {
  e.stopPropagation();
  const cell = e.currentTarget;
  ctxTarget  = {
    side: cell.dataset.side,
    row:  parseInt(cell.dataset.row),
    col:  parseInt(cell.dataset.col),
  };
  showContextMenu(e.clientX, e.clientY, cell.dataset.side);
}

function showContextMenu(x, y, side) {
  const menu  = document.getElementById('context-menu');
  const label = document.getElementById('ctx-apply');
  label.textContent = side === 'left' ? '→ 우측에 적용' : '← 좌측에 적용';

  // 화면 밖으로 나가지 않도록 조정
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
  menu.classList.remove('hidden');

  requestAnimationFrame(() => {
    const rect = menu.getBoundingClientRect();
    if (rect.right  > window.innerWidth)  menu.style.left = (x - rect.width)  + 'px';
    if (rect.bottom > window.innerHeight) menu.style.top  = (y - rect.height) + 'px';
  });
}

function hideContextMenu() {
  document.getElementById('context-menu').classList.add('hidden');
}

function applyToOpposite() {
  if (!ctxTarget) return;
  const { side, row, col } = ctxTarget;
  const opp = side === 'left' ? 'right' : 'left';

  pushUndo();

  const srcVal = state[side].data[row]?.[col] ?? '';

  // 대상 배열이 부족하면 늘림
  while (state[opp].data.length <= row) state[opp].data.push([]);
  while (state[opp].data[row].length <= col) state[opp].data[row].push('');
  state[opp].data[row][col] = srcVal;

  hideContextMenu();
  ctxTarget = null;
  refreshView();
}

// ─── Undo ────────────────────────────────────────────────────
function pushUndo() {
  state.undoStack.push({
    left:  JSON.parse(JSON.stringify(state.left.data)),
    right: JSON.parse(JSON.stringify(state.right.data)),
  });
  if (state.undoStack.length > 50) state.undoStack.shift();
}

function undo() {
  if (state.undoStack.length === 0) return;
  const prev = state.undoStack.pop();
  state.left.data  = prev.left;
  state.right.data = prev.right;
  refreshView();
}

// ─── Download ────────────────────────────────────────────────
function downloadFile(side) {
  const s    = state[side];
  if (!s.workbook) return;

  const wb   = XLSX.utils.book_new();
  // 변경된 데이터를 각 시트에 반영 (현재 선택 시트만 수정)
  s.workbook.SheetNames.forEach(name => {
    const sheet = name === s.sheetName
      ? XLSX.utils.aoa_to_sheet(s.data)
      : s.workbook.Sheets[name];
    XLSX.utils.book_append_sheet(wb, sheet, name);
  });

  const base = s.fileName.replace(/\.[^.]+$/, '');
  XLSX.writeFile(wb, `${base}_수정본.xlsx`);
}

// ─── Helpers ─────────────────────────────────────────────────
function esc(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ─── Init ────────────────────────────────────────────────────
setupDropZone('left');
setupDropZone('right');

document.getElementById('download-left-btn').addEventListener('click',  () => downloadFile('left'));
document.getElementById('download-right-btn').addEventListener('click', () => downloadFile('right'));

document.getElementById('ctx-apply').addEventListener('click', applyToOpposite);

document.addEventListener('click', e => {
  if (!e.target.closest('#context-menu')) hideContextMenu();
});

document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
    e.preventDefault();
    undo();
  }
});
