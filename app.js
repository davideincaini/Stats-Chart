/* ===== Globals ===== */
let parsedColumns = {};
let chartInstances = []; // array of { id, type, instance }
let gridCols = 2;
let gridRows = 10;

/* ===== Multiple savedDatasets ===== */
let savedDatasets = {}; // { name: { headers, rows, gridCols, gridRows } }
let activeDataset = 'Dataset 1';

/* ===== Undo/Redo ===== */
let undoStack = [];
let redoStack = [];
let undoTimer = null;

function snapshotGrid() {
  const headers = Array.from(gridHead.querySelectorAll('th:not(.row-num-header) input')).map(i => i.value);
  const rows = Array.from(gridBody.querySelectorAll('tr')).map(tr =>
    Array.from(tr.querySelectorAll('td:not(.row-num) input')).map(i => i.value)
  );
  return { headers, rows, gridCols, gridRows };
}

function restoreSnapshot(snap) {
  initGrid(snap.rows.length, snap.headers.length, true);
  const hInputs = gridHead.querySelectorAll('th:not(.row-num-header) input');
  snap.headers.forEach((h, i) => { if (hInputs[i]) hInputs[i].value = h; });
  const bodyRows = gridBody.querySelectorAll('tr');
  snap.rows.forEach((row, ri) => {
    if (!bodyRows[ri]) return;
    const cells = bodyRows[ri].querySelectorAll('td:not(.row-num) input');
    row.forEach((v, ci) => { if (cells[ci]) cells[ci].value = v; });
  });
  validateAllCells();
}

function pushUndo() {
  clearTimeout(undoTimer);
  undoTimer = setTimeout(() => {
    undoStack.push(snapshotGrid());
    if (undoStack.length > 50) undoStack.shift();
    redoStack = [];
    saveToLocalStorage();
  }, 300);
}

/* ===== localStorage persistence ===== */
function saveToLocalStorage() {
  const snap = snapshotGrid();
  savedDatasets[activeDataset] = snap;
  try {
    localStorage.setItem('statsapp_datasets', JSON.stringify(savedDatasets));
    localStorage.setItem('statsapp_active', activeDataset);
  } catch(e) {}
}

function loadFromLocalStorage() {
  try {
    const raw = localStorage.getItem('statsapp_datasets');
    const act = localStorage.getItem('statsapp_active');
    if (raw) {
      savedDatasets = JSON.parse(raw);
      if (act && savedDatasets[act]) activeDataset = act;
      const snap = savedDatasets[activeDataset];
      if (snap) {
        restoreSnapshot(snap);
        return true;
      }
    }
  } catch(e) {}
  return false;
}

function undo() {
  if (!undoStack.length) return;
  redoStack.push(snapshotGrid());
  restoreSnapshot(undoStack.pop());
}

function redo() {
  if (!redoStack.length) return;
  undoStack.push(snapshotGrid());
  restoreSnapshot(redoStack.pop());
}

/* ===== DOM refs ===== */
const csvUpload = document.getElementById('csv-upload');
const analyzeBtn = document.getElementById('analyze-btn');
const statsSection = document.getElementById('stats-section');
const statsGrid = document.getElementById('stats-grid');
const chartsSection = document.getElementById('charts-section');
const chartsContainer = document.getElementById('charts-container');
const insightsSection = document.getElementById('insights-section');
const insightsContent = document.getElementById('insights-content');
const correlationSection = document.getElementById('correlation-section');
const correlationContent = document.getElementById('correlation-content');
const invalidBadge = document.getElementById('invalid-badge');
const hypothesisSection = document.getElementById('hypothesis-section');
const hypothesisContent = document.getElementById('hypothesis-content');
const effectSizeSection = document.getElementById('effect-size-section');
const effectSizeContent = document.getElementById('effect-size-content');
const groupStatsSection = document.getElementById('group-stats-section');
const groupStatsContent = document.getElementById('group-stats-content');
const datasetTabsEl = document.getElementById('dataset-tabs');

const CHART_TYPES = [
  { value: 'auto', label: 'Auto (best fit)' },
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'scatter', label: 'Scatter' },
  { value: 'histogram', label: 'Histogram' },
  { value: 'pie', label: 'Pie / Doughnut' },
  { value: 'boxplot', label: 'Box Plot' },
  { value: 'area', label: 'Area' },
  { value: 'radar', label: 'Radar' },
  { value: 'polarArea', label: 'Polar Area' },
  { value: 'bubble', label: 'Bubble' },
  { value: 'stem', label: 'Stem-and-Leaf (Bar)' },
  { value: 'cumulative', label: 'Cumulative Frequency' },
  { value: 'pareto', label: 'Pareto' },
  { value: 'kde', label: 'KDE (Density)' },
];
let chartIdCounter = 0;
const gridTable = document.getElementById('data-grid');
const gridHead = gridTable.querySelector('thead tr');
const gridBody = gridTable.querySelector('tbody');

/* ===== Column sorting ===== */
let sortCol = -1;
let sortAsc = true;

function handleHeaderClick(e) {
  const th = e.target.closest('th');
  if (!th || th.classList.contains('row-num-header')) return;
  if (e.target.tagName === 'INPUT') return; // don't sort when editing header
  const ths = Array.from(gridHead.querySelectorAll('th:not(.row-num-header)'));
  const colIdx = ths.indexOf(th);
  if (colIdx === -1) return;

  if (sortCol === colIdx) {
    sortAsc = !sortAsc;
  } else {
    sortCol = colIdx;
    sortAsc = true;
  }
  sortByColumn(colIdx, sortAsc);
  updateSortIndicators();
}

function sortByColumn(colIdx, asc) {
  const rows = Array.from(gridBody.querySelectorAll('tr'));
  rows.sort((a, b) => {
    const aVal = a.querySelectorAll('td:not(.row-num) input')[colIdx]?.value.trim() || '';
    const bVal = b.querySelectorAll('td:not(.row-num) input')[colIdx]?.value.trim() || '';
    const aNum = Number(aVal);
    const bNum = Number(bVal);
    if (!isNaN(aNum) && !isNaN(bNum)) return asc ? aNum - bNum : bNum - aNum;
    return asc ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
  });
  rows.forEach((tr, i) => {
    gridBody.appendChild(tr);
    tr.querySelector('.row-num').textContent = i + 1;
  });
}

function updateSortIndicators() {
  gridHead.querySelectorAll('.sort-indicator').forEach(el => el.remove());
  if (sortCol < 0) return;
  const ths = gridHead.querySelectorAll('th:not(.row-num-header)');
  if (!ths[sortCol]) return;
  const span = document.createElement('span');
  span.className = 'sort-indicator';
  span.textContent = sortAsc ? ' ▲' : ' ▼';
  ths[sortCol].appendChild(span);
}

/* ===== Data validation ===== */
function validateAllCells() {
  let count = 0;
  gridBody.querySelectorAll('td:not(.row-num) input').forEach(inp => {
    const v = inp.value.trim();
    if (v !== '' && isNaN(Number(v))) {
      inp.classList.add('invalid');
      count++;
    } else {
      inp.classList.remove('invalid');
    }
  });
  if (count > 0) {
    invalidBadge.textContent = count;
    invalidBadge.classList.remove('hidden');
  } else {
    invalidBadge.classList.add('hidden');
  }
}

/* ===== Grid management ===== */
function initGrid(rows, cols, skipUndo) {
  if (!skipUndo) pushUndo();
  gridRows = rows;
  gridCols = cols;
  gridHead.innerHTML = '<th class="row-num-header"></th>';
  gridBody.innerHTML = '';
  for (let c = 0; c < cols; c++) {
    const th = document.createElement('th');
    th.addEventListener('click', handleHeaderClick);
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.value = `Col ${c + 1}`;
    inp.setAttribute('autocomplete', 'off');
    inp.setAttribute('autocorrect', 'off');
    inp.setAttribute('spellcheck', 'false');
    th.appendChild(inp);
    const delBtn = document.createElement('span');
    delBtn.className = 'col-delete-btn';
    delBtn.textContent = '×';
    delBtn.addEventListener('click', e => { e.stopPropagation(); removeColumn(th); });
    th.appendChild(delBtn);
    gridHead.appendChild(th);
  }
  for (let r = 0; r < rows; r++) {
    addRowToDOM(r + 1);
  }
  sortCol = -1;
  sortAsc = true;
}

function addRowToDOM(rowNum) {
  const tr = document.createElement('tr');
  const numTd = document.createElement('td');
  numTd.className = 'row-num';
  numTd.textContent = rowNum;
  tr.appendChild(numTd);
  for (let c = 0; c < gridCols; c++) {
    const td = document.createElement('td');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.inputMode = 'decimal';
    inp.setAttribute('autocomplete', 'off');
    inp.setAttribute('autocorrect', 'off');
    inp.setAttribute('spellcheck', 'false');
    inp.addEventListener('keydown', handleCellKey);
    inp.addEventListener('input', () => { pushUndo(); validateAllCells(); });
    td.appendChild(inp);
    tr.appendChild(td);
  }
  gridBody.appendChild(tr);
}

function removeColumn(th) {
  const ths = Array.from(gridHead.querySelectorAll('th:not(.row-num-header)'));
  if (ths.length <= 1) return; // keep at least 1 column
  pushUndo();
  const colIdx = ths.indexOf(th);
  if (colIdx === -1) return;
  th.remove();
  gridCols--;
  gridBody.querySelectorAll('tr').forEach(tr => {
    const tds = tr.querySelectorAll('td:not(.row-num)');
    if (tds[colIdx]) tds[colIdx].remove();
  });
  if (sortCol === colIdx) { sortCol = -1; }
  else if (sortCol > colIdx) { sortCol--; }
  updateSortIndicators();
  validateAllCells();
  saveToLocalStorage();
}

function addColumn() {
  pushUndo();
  gridCols++;
  const th = document.createElement('th');
  th.addEventListener('click', handleHeaderClick);
  const inp = document.createElement('input');
  inp.type = 'text';
  inp.value = `Col ${gridCols}`;
  inp.setAttribute('autocomplete', 'off');
  inp.setAttribute('autocorrect', 'off');
  inp.setAttribute('spellcheck', 'false');
  th.appendChild(inp);
  const delBtn = document.createElement('span');
  delBtn.className = 'col-delete-btn';
  delBtn.textContent = '×';
  delBtn.addEventListener('click', e => { e.stopPropagation(); removeColumn(th); });
  th.appendChild(delBtn);
  gridHead.appendChild(th);
  const rows = gridBody.querySelectorAll('tr');
  rows.forEach(tr => {
    const td = document.createElement('td');
    const ci = document.createElement('input');
    ci.type = 'text';
    ci.inputMode = 'decimal';
    ci.setAttribute('autocomplete', 'off');
    ci.setAttribute('autocorrect', 'off');
    ci.setAttribute('spellcheck', 'false');
    ci.addEventListener('keydown', handleCellKey);
    ci.addEventListener('input', () => { pushUndo(); validateAllCells(); });
    td.appendChild(ci);
    tr.appendChild(td);
  });
}

function addRow() {
  pushUndo();
  gridRows++;
  addRowToDOM(gridRows);
}

function clearGrid() {
  pushUndo();
  initGrid(10, 2, true);
}

function handleCellKey(e) {
  if (e.key === 'Tab' && !e.shiftKey) {
    const td = e.target.closest('td');
    const tr = td.closest('tr');
    const isLastCol = !td.nextElementSibling;
    const isLastRow = !tr.nextElementSibling;
    if (isLastCol && isLastRow) {
      e.preventDefault();
      addRow();
      const newRow = gridBody.lastElementChild;
      const firstInput = newRow.querySelector('td:nth-child(2) input');
      if (firstInput) firstInput.focus();
    }
  }
  if (e.key === 'Enter') {
    e.preventDefault();
    const td = e.target.closest('td');
    const tr = td.closest('tr');
    const colIdx = Array.from(tr.children).indexOf(td);
    const nextRow = tr.nextElementSibling;
    if (nextRow) {
      const cell = nextRow.children[colIdx];
      if (cell) cell.querySelector('input')?.focus();
    } else {
      addRow();
      const newRow = gridBody.lastElementChild;
      const cell = newRow.children[colIdx];
      if (cell) cell.querySelector('input')?.focus();
    }
  }
}

function getGridData() {
  const headers = Array.from(gridHead.querySelectorAll('th:not(.row-num-header) input'))
    .map((inp, i) => inp.value.trim() || `Col ${i + 1}`);
  const cols = {};
  headers.forEach(h => { cols[h] = []; });
  const rows = gridBody.querySelectorAll('tr');
  rows.forEach(tr => {
    const cells = tr.querySelectorAll('td:not(.row-num) input');
    cells.forEach((inp, i) => {
      const v = inp.value.trim();
      if (v !== '') cols[headers[i]].push(v);
    });
  });
  Object.keys(cols).forEach(k => { if (!cols[k].length) delete cols[k]; });
  return cols;
}

/* ===== Paste from Excel ===== */
document.getElementById('grid-wrapper').addEventListener('paste', e => {
  const text = (e.clipboardData || window.clipboardData).getData('text');
  if (!text || !text.includes('\t')) return;
  e.preventDefault();
  pushUndo();
  const lines = text.trim().split(/\r?\n/);
  const rows = lines.map(l => l.split('\t'));
  const numCols = Math.max(...rows.map(r => r.length));
  const firstRow = rows[0];
  const hasHeader = firstRow.some(v => isNaN(Number(v)));
  const headers = hasHeader ? firstRow : null;
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const numRows = dataRows.length;

  initGrid(Math.max(numRows, 10), Math.max(numCols, 2), true);
  if (headers) {
    const hInputs = gridHead.querySelectorAll('th:not(.row-num-header) input');
    headers.forEach((h, i) => { if (hInputs[i]) hInputs[i].value = h; });
  }
  const bodyRows = gridBody.querySelectorAll('tr');
  dataRows.forEach((dr, ri) => {
    if (!bodyRows[ri]) return;
    const cells = bodyRows[ri].querySelectorAll('td:not(.row-num) input');
    dr.forEach((v, ci) => { if (cells[ci]) cells[ci].value = v.trim(); });
  });
  validateAllCells();
});

/* ===== CSV parsing (for file upload) ===== */
function parseCSVText(text) {
  const lines = text.trim().split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return {};
  const sep = lines[0].includes('\t') ? '\t' : ',';
  const rows = lines.map(l => l.split(sep).map(v => v.trim()));
  const firstRow = rows[0];
  const hasHeader = firstRow.some(v => isNaN(Number(v)));
  const headers = hasHeader
    ? firstRow.map((h, i) => h || `Col ${i + 1}`)
    : firstRow.map((_, i) => `Col ${i + 1}`);
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const cols = {};
  headers.forEach((h, i) => {
    cols[h] = dataRows.map(r => (r[i] !== undefined ? r[i] : '')).filter(v => v !== '');
  });
  return cols;
}

function populateGridFromCSV(text) {
  pushUndo();
  const lines = text.trim().split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return;
  const sep = lines[0].includes('\t') ? '\t' : ',';
  const rows = lines.map(l => l.split(sep).map(v => v.trim()));
  const firstRow = rows[0];
  const hasHeader = firstRow.some(v => isNaN(Number(v)));
  const headers = hasHeader ? firstRow : null;
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const numCols = Math.max(...rows.map(r => r.length));
  const numRows = dataRows.length;

  initGrid(Math.max(numRows, 10), numCols, true);

  if (headers) {
    const hInputs = gridHead.querySelectorAll('th:not(.row-num-header) input');
    headers.forEach((h, i) => { if (hInputs[i]) hInputs[i].value = h; });
  }
  const bodyRows = gridBody.querySelectorAll('tr');
  dataRows.forEach((dr, ri) => {
    if (!bodyRows[ri]) return;
    const cells = bodyRows[ri].querySelectorAll('td:not(.row-num) input');
    dr.forEach((v, ci) => { if (cells[ci]) cells[ci].value = v; });
  });
  validateAllCells();
}

csvUpload.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => populateGridFromCSV(ev.target.result);
  reader.readAsText(file);
});

/* ===== Statistics ===== */
function toNumbers(arr) { return arr.map(Number).filter(v => !isNaN(v)); }

function computeStats(nums) {
  const sorted = [...nums].sort((a, b) => a - b);
  const n = sorted.length;
  if (!n) return null;
  const sum = sorted.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  const median = quantile(sorted, 0.5);
  const q1 = quantile(sorted, 0.25);
  const q3 = quantile(sorted, 0.75);
  const iqr = q3 - q1;
  const variance = sorted.reduce((s, v) => s + (v - mean) ** 2, 0) / n;
  const std = Math.sqrt(variance);
  const se = std / Math.sqrt(n);
  const freq = {};
  sorted.forEach(v => { freq[v] = (freq[v] || 0) + 1; });
  const maxFreq = Math.max(...Object.values(freq));
  const modes = Object.keys(freq).filter(k => freq[k] === maxFreq).map(Number);
  const mode = maxFreq === 1 ? 'None' : modes.join(', ');
  const m3 = sorted.reduce((s, v) => s + ((v - mean) / (std || 1)) ** 3, 0) / n;
  const m4 = sorted.reduce((s, v) => s + ((v - mean) / (std || 1)) ** 4, 0) / n - 3;

  // Confidence intervals
  const z90 = 1.645, z95 = 1.960, z99 = 2.576;
  const ci90 = [r(mean - z90 * se), r(mean + z90 * se)];
  const ci95 = [r(mean - z95 * se), r(mean + z95 * se)];
  const ci99 = [r(mean - z99 * se), r(mean + z99 * se)];

  // Normality test (D'Agostino-Pearson omnibus)
  const normality = dagostinoPearson(m3, m4, n);

  // Outlier detection (IQR method)
  const lowerFence = q1 - 1.5 * iqr;
  const upperFence = q3 + 1.5 * iqr;
  const outliers = sorted.filter(v => v < lowerFence || v > upperFence);

  return {
    Count: n, Mean: r(mean), Median: r(median), Mode: mode,
    'Std Dev': r(std), Variance: r(variance),
    Min: r(sorted[0]), Max: r(sorted[n - 1]), Range: r(sorted[n - 1] - sorted[0]),
    Q1: r(q1), Q2: r(median), Q3: r(q3), IQR: r(iqr),
    Skewness: r(m3), Kurtosis: r(m4),
    'CI 90%': `[${ci90[0]}, ${ci90[1]}]`,
    'CI 95%': `[${ci95[0]}, ${ci95[1]}]`,
    'CI 99%': `[${ci99[0]}, ${ci99[1]}]`,
    Normality: normality.label,
    Outliers: outliers.length ? outliers.map(v => r(v)).join(', ') : 'None',
    _raw: { mean, std, m3, m4, n, q1, q3, iqr, outliers, normality, sorted, lowerFence, upperFence }
  };
}

/* D'Agostino-Pearson normality test */
function dagostinoPearson(skew, kurt, n) {
  if (n < 8) return { normal: null, p: null, label: 'N too small' };
  // Z-score for skewness
  const Y = skew * Math.sqrt((n + 1) * (n + 3) / (6 * (n - 2)));
  const b2s = 3 * (n * n + 27 * n - 70) * (n + 1) * (n + 3) / ((n - 2) * (n + 5) * (n + 7) * (n + 9));
  const W2 = -1 + Math.sqrt(2 * (b2s - 1));
  const delta = 1 / Math.sqrt(Math.log(Math.sqrt(W2)));
  const alpha = Math.sqrt(2 / (W2 - 1));
  const Zs = delta * Math.log(Y / alpha + Math.sqrt((Y / alpha) ** 2 + 1));

  // Z-score for kurtosis (excess kurtosis)
  const Ek = kurt; // already excess
  const varK = 24 * n * (n - 2) * (n - 3) / ((n + 1) ** 2 * (n + 3) * (n + 5));
  const Zk = Ek / Math.sqrt(varK);

  const K2 = Zs * Zs + Zk * Zk; // chi-squared with df=2
  // approximate p-value from chi-squared(2)
  const p = Math.exp(-K2 / 2);
  const normal = p > 0.05;
  return { normal, p: r(p), label: `${normal ? 'Yes' : 'No'} (p≈${r(p)})` };
}

function quantile(sorted, p) {
  const idx = p * (sorted.length - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}

function r(v) { return typeof v === 'number' ? +v.toFixed(4) : v; }

/* ===== Correlation ===== */
function pearsonR(xs, ys) {
  const n = Math.min(xs.length, ys.length);
  if (n < 2) return NaN;
  let sx = 0, sy = 0, sxy = 0, sx2 = 0, sy2 = 0;
  for (let i = 0; i < n; i++) {
    sx += xs[i]; sy += ys[i]; sxy += xs[i] * ys[i];
    sx2 += xs[i] ** 2; sy2 += ys[i] ** 2;
  }
  const denom = Math.sqrt((n * sx2 - sx ** 2) * (n * sy2 - sy ** 2));
  return denom === 0 ? 0 : (n * sxy - sx * sy) / denom;
}

function buildCorrelationMatrix(cols) {
  const names = Object.keys(cols);
  const numCols = {};
  names.forEach(n => {
    const nums = toNumbers(cols[n]);
    if (nums.length >= 2) numCols[n] = nums;
  });
  const keys = Object.keys(numCols);
  if (keys.length < 2) return null;
  const matrix = {};
  keys.forEach(a => {
    matrix[a] = {};
    keys.forEach(b => {
      matrix[a][b] = r(pearsonR(numCols[a], numCols[b]));
    });
  });
  return { keys, matrix };
}

function renderCorrelationTable(corrData) {
  if (!corrData) {
    correlationSection.classList.add('hidden');
    return;
  }
  correlationSection.classList.remove('hidden');
  const { keys, matrix } = corrData;
  let html = '<table class="corr-table"><thead><tr><th></th>';
  keys.forEach(k => { html += `<th>${k}</th>`; });
  html += '</tr></thead><tbody>';
  keys.forEach(row => {
    html += `<tr><th>${row}</th>`;
    keys.forEach(col => {
      const val = matrix[row][col];
      const abs = Math.abs(val);
      const hue = val >= 0 ? 210 : 0;
      const bg = `hsla(${hue},70%,50%,${abs * 0.4})`;
      html += `<td style="background:${bg}">${val}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  correlationContent.innerHTML = html;
}

/* ===== Regression ===== */
function leastSquares(xs, ys) {
  const n = Math.min(xs.length, ys.length);
  if (n < 2) return null;
  let sx = 0, sy = 0, sxy = 0, sx2 = 0, sy2 = 0;
  for (let i = 0; i < n; i++) {
    sx += xs[i]; sy += ys[i]; sxy += xs[i] * ys[i];
    sx2 += xs[i] ** 2; sy2 += ys[i] ** 2;
  }
  const m = (n * sxy - sx * sy) / (n * sx2 - sx ** 2);
  const b = (sy - m * sx) / n;
  // R²
  const yMean = sy / n;
  let ssTot = 0, ssRes = 0;
  for (let i = 0; i < n; i++) {
    ssTot += (ys[i] - yMean) ** 2;
    ssRes += (ys[i] - (m * xs[i] + b)) ** 2;
  }
  const r2 = ssTot === 0 ? 1 : 1 - ssRes / ssTot;
  return { m: +m.toFixed(4), b: +b.toFixed(4), r2: +r2.toFixed(4) };
}

/* ===== Insights ===== */
function generateInsights(allStats, corrData) {
  const items = [];
  const colNames = Object.keys(allStats);

  colNames.forEach(name => {
    const s = allStats[name];
    if (!s || !s._raw) return;
    const raw = s._raw;

    // One-line summary per column: shape + normality + outliers
    const parts = [];
    if (Math.abs(raw.m3) < 0.5) parts.push('symmetric');
    else parts.push(raw.m3 > 0 ? `right-skewed (${r(raw.m3)})` : `left-skewed (${r(raw.m3)})`);

    if (raw.normality.normal === true) parts.push('normal');
    else if (raw.normality.normal === false) parts.push('non-normal');

    if (raw.outliers.length) parts.push(`${raw.outliers.length} outlier${raw.outliers.length > 1 ? 's' : ''}`);

    items.push(`<span class="insight-col">${name}</span> ${parts.join(' · ')}`);
  });

  // Correlation highlights
  if (corrData && corrData.keys.length >= 2) {
    for (let i = 0; i < corrData.keys.length; i++) {
      for (let j = i + 1; j < corrData.keys.length; j++) {
        const a = corrData.keys[i], b = corrData.keys[j];
        const val = corrData.matrix[a][b];
        if (Math.abs(val) > 0.7) {
          items.push(`<span class="insight-col">${a} ↔ ${b}</span> ${val > 0 ? '+' : '−'}corr (r=${val})`);
        }
      }
    }
  }

  if (!items.length) return '<div class="insight-line">No notable patterns.</div>';
  return items.map(s => `<div class="insight-line">${s}</div>`).join('');
}

/* ===== Outlier cell highlighting ===== */
function highlightOutliers(allStats) {
  // clear old highlights
  gridBody.querySelectorAll('td:not(.row-num) input.outlier').forEach(inp => inp.classList.remove('outlier'));

  const headers = Array.from(gridHead.querySelectorAll('th:not(.row-num-header) input')).map(i => i.value.trim());
  const bodyRows = gridBody.querySelectorAll('tr');

  headers.forEach((h, colIdx) => {
    const stats = allStats[h];
    if (!stats || !stats._raw) return;
    const { lowerFence, upperFence } = stats._raw;
    bodyRows.forEach(tr => {
      const inp = tr.querySelectorAll('td:not(.row-num) input')[colIdx];
      if (!inp) return;
      const v = Number(inp.value.trim());
      if (!isNaN(v) && inp.value.trim() !== '' && (v < lowerFence || v > upperFence)) {
        inp.classList.add('outlier');
      }
    });
  });
}

/* ===== Chart selection & rendering ===== */
function detectType(cols) {
  const names = Object.keys(cols);
  const numericCols = names.filter(n => {
    const nums = toNumbers(cols[n]);
    return nums.length > cols[n].length * 0.5;
  });
  const catCols = names.filter(n => !numericCols.includes(n));

  if (numericCols.length >= 2) return { type: 'scatter', x: numericCols[0], y: numericCols[1] };
  if (numericCols.length === 1 && catCols.length === 1) return { type: 'catbar', cat: catCols[0], num: numericCols[0] };
  if (numericCols.length === 1) {
    const vals = toNumbers(cols[numericCols[0]]);
    const unique = new Set(vals).size;
    if (unique <= 8) return { type: 'pie', col: numericCols[0] };
    if (isTimeSeries(vals)) return { type: 'line', col: numericCols[0] };
    if (vals.length <= 20) return { type: 'bar', col: numericCols[0] };
    return { type: 'histogram', col: numericCols[0] };
  }
  if (names.length === 1) return { type: 'pie', col: names[0] };
  return { type: 'bar', col: names[0] };
}

function isTimeSeries(nums) {
  if (nums.length < 5) return false;
  let incr = 0;
  for (let i = 1; i < nums.length; i++) if (nums[i] >= nums[i - 1]) incr++;
  return incr / (nums.length - 1) > 0.85;
}

/* ===== Chart card management ===== */
/* Types that support axis variable selection */
const AXIS_TYPES = ['scatter', 'line', 'bar', 'area', 'histogram', 'kde', 'bubble', 'catbar'];

function getColumnNames() {
  return Object.keys(parsedColumns);
}

function buildAxisSelectors(id) {
  const row = document.createElement('div');
  row.className = 'axis-selector-row';

  const makeSelect = (label, cls, includeNone) => {
    const wrap = document.createElement('label');
    wrap.className = 'axis-label';
    wrap.textContent = label;
    const sel = document.createElement('select');
    sel.className = cls;
    if (includeNone) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '(auto)';
      sel.appendChild(opt);
    }
    getColumnNames().forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', () => renderSingleChart(id));
    wrap.appendChild(sel);
    return wrap;
  };

  row.appendChild(makeSelect('X', 'axis-x-select', true));
  row.appendChild(makeSelect('Y', 'axis-y-select', true));
  return row;
}

function updateAxisOptions(card) {
  const colNames = getColumnNames();
  card.querySelectorAll('.axis-x-select, .axis-y-select').forEach(sel => {
    const cur = sel.value;
    // keep (auto) option, rebuild the rest
    while (sel.options.length > 1) sel.remove(1);
    colNames.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      if (name === cur) opt.selected = true;
      sel.appendChild(opt);
    });
  });
}

function addChartCard(selectedType) {
  const id = ++chartIdCounter;
  const card = document.createElement('div');
  card.className = 'chart-card';
  card.dataset.chartId = id;

  const header = document.createElement('div');
  header.className = 'chart-card-header';

  const select = document.createElement('select');
  CHART_TYPES.forEach(t => {
    const opt = document.createElement('option');
    opt.value = t.value;
    opt.textContent = t.label;
    if (t.value === selectedType) opt.selected = true;
    select.appendChild(opt);
  });
  select.addEventListener('change', () => {
    const axisRow = card.querySelector('.axis-selector-row');
    const type = select.value === 'auto' ? detectType(parsedColumns).type : select.value;
    if (AXIS_TYPES.includes(type)) {
      if (axisRow) axisRow.classList.remove('hidden');
    } else {
      if (axisRow) axisRow.classList.add('hidden');
    }
    renderSingleChart(id);
  });

  const removeBtn = document.createElement('button');
  removeBtn.className = 'pill-btn-remove';
  removeBtn.innerHTML = '&times;';
  removeBtn.addEventListener('click', () => removeChartCard(id));

  header.appendChild(select);
  header.appendChild(removeBtn);

  const axisRow = buildAxisSelectors(id);
  const autoType = selectedType === 'auto' ? detectType(parsedColumns).type : selectedType;
  if (!AXIS_TYPES.includes(autoType)) axisRow.classList.add('hidden');

  const label = document.createElement('p');
  label.className = 'chart-type-label';

  const canvas = document.createElement('canvas');

  card.appendChild(header);
  card.appendChild(axisRow);
  card.appendChild(label);
  card.appendChild(canvas);
  chartsContainer.appendChild(card);

  chartInstances.push({ id, type: selectedType, instance: null });
  renderSingleChart(id);
}

function removeChartCard(id) {
  const idx = chartInstances.findIndex(c => c.id === id);
  if (idx !== -1) {
    if (chartInstances[idx].instance) chartInstances[idx].instance.destroy();
    chartInstances.splice(idx, 1);
  }
  const card = chartsContainer.querySelector(`[data-chart-id="${id}"]`);
  if (card) card.remove();
}

function destroyAllCharts() {
  chartInstances.forEach(c => { if (c.instance) c.instance.destroy(); });
  chartInstances = [];
  chartsContainer.innerHTML = '';
  chartIdCounter = 0;
}

/* ===== KDE (Gaussian Kernel Density Estimation) ===== */
function computeKDE(nums, nPoints) {
  if (!nPoints) nPoints = 100;
  const n = nums.length;
  if (n < 2) return { points: [], densities: [] };
  const sorted = [...nums].sort((a, b) => a - b);
  const mean = sorted.reduce((a, b) => a + b, 0) / n;
  const std = Math.sqrt(sorted.reduce((s, v) => s + (v - mean) ** 2, 0) / n) || 1;
  // Silverman's rule of thumb for bandwidth
  const iqr = quantile(sorted, 0.75) - quantile(sorted, 0.25);
  const h = 0.9 * Math.min(std, (iqr || std) / 1.34) * Math.pow(n, -0.2);
  const pad = 3 * h;
  const xMin = sorted[0] - pad;
  const xMax = sorted[n - 1] + pad;
  const step = (xMax - xMin) / (nPoints - 1);
  const points = [];
  const densities = [];
  for (let i = 0; i < nPoints; i++) {
    const x = xMin + i * step;
    points.push(x);
    let sum = 0;
    for (let j = 0; j < n; j++) {
      const u = (x - nums[j]) / h;
      sum += Math.exp(-0.5 * u * u); // Gaussian kernel (without constant, normalized below)
    }
    densities.push(sum / (n * h * Math.sqrt(2 * Math.PI)));
  }
  return { points, densities };
}

/* ===== Group data by categorical column ===== */
const GROUP_COLORS = [
  '#4a90d9', '#e55', '#34c759', '#ff9500', '#af52de',
  '#ff2d55', '#5ac8fa', '#ffcc00', '#8e8e93', '#30b0c7'
];

function groupByCategory(cols, catCol, numericCols) {
  const catValues = cols[catCol];
  const groups = {};
  catValues.forEach((label, i) => {
    if (!groups[label]) groups[label] = {};
    numericCols.forEach(nc => {
      if (!groups[label][nc]) groups[label][nc] = [];
      const val = cols[nc][i];
      if (val !== undefined) groups[label][nc].push(val);
    });
  });
  return groups; // { "GroupA": { "Col1": [...], "Col2": [...] }, ... }
}

function getGroupedRawRows(cols, catCol) {
  // returns rows as array of { cat, vals: {colName: rawVal} } aligned by row index
  const catValues = cols[catCol];
  const otherCols = Object.keys(cols).filter(n => n !== catCol);
  const maxLen = Math.max(...Object.values(cols).map(a => a.length));
  const rows = [];
  for (let i = 0; i < maxLen; i++) {
    const row = { cat: catValues[i] || '' };
    otherCols.forEach(c => { row[c] = cols[c][i]; });
    rows.push(row);
  }
  return rows;
}

function buildConfig(requestedType, cols, axisOverride) {
  const ax = axisOverride || {};
  const names = Object.keys(cols);
  let numericCols = names.filter(n => toNumbers(cols[n]).length > cols[n].length * 0.5);
  const catCols = names.filter(n => !numericCols.includes(n));

  // Reorder numericCols based on axis override
  if (ax.x && numericCols.includes(ax.x)) {
    numericCols = [ax.x, ...numericCols.filter(n => n !== ax.x)];
  }
  if (ax.y && numericCols.includes(ax.y) && numericCols[0] !== ax.y) {
    // ensure Y is second
    numericCols = [numericCols[0], ax.y, ...numericCols.filter(n => n !== numericCols[0] && n !== ax.y)];
  }

  const firstNumCol = ax.x && cols[ax.x] ? ax.x : numericCols[0];
  const allNums = firstNumCol ? toNumbers(cols[firstNumCol]) : [];
  const canGroup = catCols.length >= 1 && numericCols.length >= 1;
  const groupCol = canGroup ? catCols[0] : null;
  const groups = canGroup ? groupByCategory(cols, groupCol, numericCols) : null;
  const groupKeys = groups ? Object.keys(groups) : [];

  function freqMap(arr) {
    const f = {};
    arr.forEach(v => { f[v] = (f[v] || 0) + 1; });
    return f;
  }

  function histBins(nums) {
    const min = Math.min(...nums), max = Math.max(...nums);
    const binCount = Math.ceil(Math.sqrt(nums.length));
    const binW = (max - min) / binCount || 1;
    const bins = Array(binCount).fill(0);
    nums.forEach(v => { const i = Math.min(Math.floor((v - min) / binW), binCount - 1); bins[i]++; });
    const labels = bins.map((_, i) => r(min + i * binW) + '–' + r(min + (i + 1) * binW));
    return { labels, bins };
  }

  let type = requestedType;
  let resolvedLabel = type;

  if (type === 'auto') {
    const info = detectType(cols);
    type = info.type;
    resolvedLabel = 'Auto → ' + type;
  }

  let config;

  if (type === 'scatter') {
    if (numericCols.length >= 2) {
      const datasets = [];
      if (canGroup && groupKeys.length > 1) {
        // Grouped scatter: one dataset per category
        groupKeys.forEach((gk, gi) => {
          const xV = toNumbers(groups[gk][numericCols[0]] || []);
          const yV = toNumbers(groups[gk][numericCols[1]] || []);
          const data = xV.map((x, i) => ({ x, y: yV[i] })).filter(d => d.y !== undefined);
          datasets.push({ label: gk, data, backgroundColor: GROUP_COLORS[gi % GROUP_COLORS.length] });
        });
      } else {
        const xV = toNumbers(cols[numericCols[0]]);
        const yV = toNumbers(cols[numericCols[1]]);
        const data = xV.map((x, i) => ({ x, y: yV[i] })).filter(d => d.y !== undefined);
        datasets.push({ label: `${numericCols[0]} vs ${numericCols[1]}`, data, backgroundColor: '#4a90d9' });
      }

      // Regression lines (on all data)
      const xAll = toNumbers(cols[numericCols[0]]);
      const yAll = toNumbers(cols[numericCols[1]]);
      const reg = leastSquares(xAll, yAll);
      if (reg) {
        const xMin = Math.min(...xAll), xMax = Math.max(...xAll);
        datasets.push({
          label: `y = ${reg.m}x + ${reg.b}, R² = ${reg.r2}`,
          data: [{ x: xMin, y: reg.m * xMin + reg.b }, { x: xMax, y: reg.m * xMax + reg.b }],
          type: 'line',
          borderColor: '#e55',
          borderWidth: 2,
          borderDash: [6, 3],
          pointRadius: 0,
          fill: false
        });
      }
      // Polynomial regression (degree 2)
      const poly2 = polyFit(xAll, yAll, 2);
      if (poly2 && xAll.length > 3) {
        const xMin = Math.min(...xAll), xMax = Math.max(...xAll);
        const nPts = 50;
        const step = (xMax - xMin) / (nPts - 1);
        const polyData = [];
        for (let i = 0; i < nPts; i++) {
          const px = xMin + i * step;
          let py = 0;
          poly2.coeffs.forEach((c, j) => { py += c * Math.pow(px, j); });
          polyData.push({ x: +px.toFixed(4), y: +py.toFixed(4) });
        }
        datasets.push({
          label: `Poly2: R²=${poly2.r2}`,
          data: polyData,
          type: 'line',
          borderColor: '#34c759',
          borderWidth: 2,
          borderDash: [3, 3],
          pointRadius: 0,
          fill: false
        });
      }
      config = { type: 'scatter', data: { datasets } };
    } else {
      const data = allNums.map((v, i) => ({ x: i + 1, y: v }));
      config = { type: 'scatter', data: { datasets: [{ label: firstNumCol || 'Data', data, backgroundColor: '#4a90d9' }] } };
    }
  } else if (type === 'catbar') {
    if (catCols.length && numericCols.length) {
      config = { type: 'bar', data: { labels: cols[catCols[0]], datasets: [{ label: numericCols[0], data: toNumbers(cols[numericCols[0]]), backgroundColor: '#4a90d9' }] } };
    } else {
      config = { type: 'bar', data: { labels: allNums.map((_, i) => i + 1), datasets: [{ label: firstNumCol || 'Data', data: allNums, backgroundColor: '#4a90d9' }] } };
    }
  } else if (type === 'bar') {
    if (canGroup && numericCols.length >= 2 && groupKeys.length > 1) {
      // Grouped bar: categories on x-axis, one dataset per numeric column
      const uniqueCats = [...new Set(cols[groupCol])];
      const barDs = numericCols.map((nc, ci) => {
        const catMap = {};
        cols[groupCol].forEach((cat, i) => {
          const v = Number(cols[nc][i]);
          if (!isNaN(v)) {
            if (!catMap[cat]) catMap[cat] = [];
            catMap[cat].push(v);
          }
        });
        return {
          label: nc,
          data: uniqueCats.map(c => catMap[c] ? catMap[c].reduce((a, b) => a + b, 0) / catMap[c].length : 0),
          backgroundColor: GROUP_COLORS[ci % GROUP_COLORS.length]
        };
      });
      config = { type: 'bar', data: { labels: uniqueCats, datasets: barDs } };
    } else if (catCols.length && numericCols.length) {
      config = { type: 'bar', data: { labels: cols[catCols[0]], datasets: [{ label: numericCols[0], data: toNumbers(cols[numericCols[0]]), backgroundColor: '#4a90d9' }] } };
    } else {
      config = { type: 'bar', data: { labels: allNums.map((_, i) => i + 1), datasets: [{ label: firstNumCol || 'Data', data: allNums, backgroundColor: '#4a90d9' }] } };
    }
  } else if (type === 'line') {
    if (canGroup && groupKeys.length > 1 && numericCols.length >= 1) {
      // Grouped line: one line per category
      const maxLen = Math.max(...groupKeys.map(gk => toNumbers(groups[gk][firstNumCol] || []).length));
      const labels = Array.from({ length: maxLen }, (_, i) => i + 1);
      const lineGroupDs = groupKeys.map((gk, gi) => ({
        label: gk,
        data: toNumbers(groups[gk][firstNumCol] || []),
        borderColor: GROUP_COLORS[gi % GROUP_COLORS.length],
        fill: false,
        tension: 0.2
      }));
      config = { type: 'line', data: { labels, datasets: lineGroupDs } };
    } else {
      const lineDs = [{ label: firstNumCol || 'Data', data: allNums, borderColor: '#4a90d9', fill: false, tension: 0.2 }];
      // Moving average overlay
      if (allNums.length >= 5) {
        const ma = movingAverage(allNums);
        lineDs.push({ label: 'Moving Avg', data: ma.map(v => +v.toFixed(4)), borderColor: '#ff9500', borderWidth: 2, borderDash: [5, 3], pointRadius: 0, fill: false, tension: 0.3 });
      }
      config = { type: 'line', data: { labels: allNums.map((_, i) => i + 1), datasets: lineDs } };
    }
  } else if (type === 'area') {
    config = { type: 'line', data: { labels: allNums.map((_, i) => i + 1), datasets: [{ label: firstNumCol || 'Data', data: allNums, borderColor: '#4a90d9', backgroundColor: 'rgba(74,144,217,0.15)', fill: true, tension: 0.2 }] } };
    resolvedLabel = 'area';
  } else if (type === 'histogram') {
    if (canGroup && groupKeys.length > 1) {
      // Grouped histogram: shared bins, one dataset per group
      const allVals = toNumbers(cols[firstNumCol]);
      const min = Math.min(...allVals), max = Math.max(...allVals);
      const binCount = Math.ceil(Math.sqrt(allVals.length));
      const binW = (max - min) / binCount || 1;
      const labels = Array.from({ length: binCount }, (_, i) => r(min + i * binW) + '–' + r(min + (i + 1) * binW));
      const histGroupDs = groupKeys.map((gk, gi) => {
        const nums = toNumbers(groups[gk][firstNumCol] || []);
        const bins = Array(binCount).fill(0);
        nums.forEach(v => { const idx = Math.min(Math.floor((v - min) / binW), binCount - 1); bins[idx]++; });
        return { label: gk, data: bins, backgroundColor: GROUP_COLORS[gi % GROUP_COLORS.length] + '99' };
      });
      config = { type: 'bar', data: { labels, datasets: histGroupDs } };
    } else {
      const h = histBins(allNums);
      const histDs = [{ label: 'Frequency', data: h.bins, backgroundColor: '#4a90d9' }];
      // Normal PDF overlay
      if (allNums.length >= 5) {
        const mean = allNums.reduce((a, b) => a + b, 0) / allNums.length;
        const std = Math.sqrt(allNums.reduce((s, v) => s + (v - mean) ** 2, 0) / allNums.length) || 1;
        const min = Math.min(...allNums), max = Math.max(...allNums);
        const binW = (max - min) / h.bins.length || 1;
        const pdfData = h.labels.map((_, i) => {
          const midpoint = min + (i + 0.5) * binW;
          return +(normalPDF(midpoint, mean, std) * allNums.length * binW).toFixed(4);
        });
        histDs.push({ label: 'Normal fit', data: pdfData, type: 'line', borderColor: '#e55', borderWidth: 2, pointRadius: 0, fill: false, tension: 0.4 });
      }
      config = { type: 'bar', data: { labels: h.labels, datasets: histDs } };
    }
  } else if (type === 'pie') {
    const col = catCols[0] || firstNumCol || names[0];
    const freq = freqMap(cols[col]);
    const labels = Object.keys(freq);
    const data = Object.values(freq);
    const colors = labels.map((_, i) => `hsl(${(i * 360 / labels.length) % 360},65%,55%)`);
    config = { type: 'doughnut', data: { labels, datasets: [{ data, backgroundColor: colors }] } };
  } else if (type === 'boxplot') {
    const boxLabels = [];
    const mins = [], q1s = [], medians = [], q3s = [], maxs = [];

    if (canGroup && groupKeys.length > 1) {
      // Grouped boxplot: one box per group (for the first numeric column)
      groupKeys.forEach(gk => {
        const sorted = [...toNumbers(groups[gk][firstNumCol] || [])].sort((a, b) => a - b);
        if (!sorted.length) return;
        boxLabels.push(gk);
        mins.push(sorted[0]);
        q1s.push(quantile(sorted, 0.25));
        medians.push(quantile(sorted, 0.5));
        q3s.push(quantile(sorted, 0.75));
        maxs.push(sorted[sorted.length - 1]);
      });
    } else {
      const targets = numericCols.length ? numericCols : (firstNumCol ? [firstNumCol] : []);
      targets.forEach(col => {
        const sorted = [...toNumbers(cols[col])].sort((a, b) => a - b);
        if (!sorted.length) return;
        boxLabels.push(col);
        mins.push(sorted[0]);
        q1s.push(quantile(sorted, 0.25));
        medians.push(quantile(sorted, 0.5));
        q3s.push(quantile(sorted, 0.75));
        maxs.push(sorted[sorted.length - 1]);
      });
    }

    const iqrColors = boxLabels.map((_, i) => GROUP_COLORS[i % GROUP_COLORS.length] + 'aa');
    const rangeColors = boxLabels.map((_, i) => GROUP_COLORS[i % GROUP_COLORS.length] + '33');
    const borderColors = boxLabels.map((_, i) => GROUP_COLORS[i % GROUP_COLORS.length]);

    config = {
      type: 'bar',
      data: {
        labels: boxLabels,
        datasets: [
          { label: 'IQR', data: boxLabels.map((_, i) => [q1s[i], q3s[i]]), backgroundColor: iqrColors, borderColor: borderColors, borderWidth: 1 },
          { label: 'Range', data: boxLabels.map((_, i) => [mins[i], maxs[i]]), backgroundColor: rangeColors, borderColor: borderColors, borderWidth: 1 },
        ]
      },
      options: {
        responsive: true,
        indexAxis: 'x',
        scales: {
          y: {
            min: Math.floor(Math.min(...mins) - (Math.max(...maxs) - Math.min(...mins)) * 0.1),
            max: Math.ceil(Math.max(...maxs) + (Math.max(...maxs) - Math.min(...mins)) * 0.1),
            beginAtZero: false
          }
        },
        plugins: {
          tooltip: {
            callbacks: {
              afterBody: (ctx) => {
                const i = ctx[0].dataIndex;
                return `Min: ${r(mins[i])} | Q1: ${r(q1s[i])} | Med: ${r(medians[i])} | Q3: ${r(q3s[i])} | Max: ${r(maxs[i])}`;
              }
            }
          }
        }
      }
    };
  } else if (type === 'radar') {
    const labels = allNums.map((_, i) => `${i + 1}`);
    if (numericCols.length >= 2) {
      const radarDs = numericCols.map((col, ci) => ({
        label: col,
        data: toNumbers(cols[col]),
        borderColor: `hsl(${ci * 120},65%,50%)`,
        backgroundColor: `hsla(${ci * 120},65%,50%,0.1)`,
        fill: true
      }));
      const maxLen = Math.max(...radarDs.map(d => d.data.length));
      config = { type: 'radar', data: { labels: Array.from({ length: maxLen }, (_, i) => `${i + 1}`), datasets: radarDs } };
    } else {
      config = { type: 'radar', data: { labels, datasets: [{ label: firstNumCol || 'Data', data: allNums, borderColor: '#4a90d9', backgroundColor: 'rgba(74,144,217,0.15)', fill: true }] } };
    }
  } else if (type === 'polarArea') {
    const col = catCols[0] || firstNumCol || names[0];
    const freq = freqMap(cols[col]);
    const labels = Object.keys(freq);
    const data = Object.values(freq);
    const colors = labels.map((_, i) => `hsl(${(i * 360 / labels.length) % 360},65%,55%)`);
    config = { type: 'polarArea', data: { labels, datasets: [{ data, backgroundColor: colors }] } };
  } else if (type === 'bubble') {
    let data;
    if (numericCols.length >= 3) {
      const xV = toNumbers(cols[numericCols[0]]);
      const yV = toNumbers(cols[numericCols[1]]);
      const rV = toNumbers(cols[numericCols[2]]);
      data = xV.map((x, i) => ({ x, y: yV[i] || 0, r: Math.abs(rV[i] || 5) }));
    } else if (numericCols.length >= 2) {
      const xV = toNumbers(cols[numericCols[0]]);
      const yV = toNumbers(cols[numericCols[1]]);
      data = xV.map((x, i) => ({ x, y: yV[i] || 0, r: 5 }));
    } else {
      data = allNums.map((v, i) => ({ x: i + 1, y: v, r: 5 }));
    }
    config = { type: 'bubble', data: { datasets: [{ label: 'Data', data, backgroundColor: 'rgba(74,144,217,0.5)' }] } };
  } else if (type === 'stem') {
    const sorted = [...allNums].sort((a, b) => a - b);
    config = { type: 'bar', data: { labels: sorted.map(v => r(v)), datasets: [{ label: 'Value', data: sorted, backgroundColor: '#4a90d9', barThickness: 4 }] } };
    resolvedLabel = 'stem (bar)';
  } else if (type === 'cumulative') {
    const sorted = [...allNums].sort((a, b) => a - b);
    const cumData = sorted.map((_, i) => (i + 1) / sorted.length);
    config = { type: 'line', data: { labels: sorted.map(v => r(v)), datasets: [{ label: 'Cumulative Freq', data: cumData, borderColor: '#4a90d9', fill: false, stepped: true }] } };
  } else if (type === 'pareto') {
    const h = histBins(allNums);
    const total = h.bins.reduce((a, b) => a + b, 0);
    let cum = 0;
    const cumLine = h.bins.map(b => { cum += b; return +(cum / total * 100).toFixed(1); });
    config = {
      type: 'bar',
      data: {
        labels: h.labels,
        datasets: [
          { label: 'Frequency', data: h.bins, backgroundColor: '#4a90d9', yAxisID: 'y' },
          { label: 'Cumulative %', data: cumLine, type: 'line', borderColor: '#e55', fill: false, yAxisID: 'y1' }
        ]
      },
      options: {
        responsive: true,
        scales: {
          y: { beginAtZero: true, position: 'left' },
          y1: { beginAtZero: true, max: 100, position: 'right', grid: { drawOnChartArea: false } }
        }
      }
    };
  } else if (type === 'kde') {
    // KDE: Gaussian kernel density estimation
    if (canGroup && groupKeys.length > 1) {
      // Grouped KDE: one density curve per category
      // Use shared x-axis from all data
      const allVals = toNumbers(cols[firstNumCol]);
      const globalKDE = computeKDE(allVals);
      const sharedPoints = globalKDE.points;
      const kdeGroupDs = groupKeys.map((gk, gi) => {
        const nums = toNumbers(groups[gk][firstNumCol] || []);
        if (nums.length < 2) return null;
        const kde = computeKDE(nums);
        return {
          label: gk,
          data: kde.densities,
          borderColor: GROUP_COLORS[gi % GROUP_COLORS.length],
          backgroundColor: GROUP_COLORS[gi % GROUP_COLORS.length] + '22',
          fill: true,
          tension: 0.4,
          pointRadius: 0
        };
      }).filter(Boolean);
      config = {
        type: 'line',
        data: { labels: sharedPoints.map(v => r(v)), datasets: kdeGroupDs },
        options: { responsive: true, scales: { y: { beginAtZero: true } } }
      };
    } else {
      const kdeResult = computeKDE(allNums);
      const kdeDs = [
        { label: 'Density', data: kdeResult.densities, borderColor: '#4a90d9', backgroundColor: 'rgba(74,144,217,0.15)', fill: true, tension: 0.4, pointRadius: 0 }
      ];
      // overlay multiple columns if available
      if (numericCols.length >= 2) {
        numericCols.slice(1).forEach((col, ci) => {
          const nums2 = toNumbers(cols[col]);
          if (nums2.length < 2) return;
          const kde2 = computeKDE(nums2);
          kdeDs.push({
            label: col + ' density',
            data: kde2.densities,
            borderColor: GROUP_COLORS[(ci + 1) % GROUP_COLORS.length],
            backgroundColor: 'transparent',
            fill: false,
            tension: 0.4,
            pointRadius: 0
          });
        });
        kdeDs[0].label = numericCols[0] + ' density';
      }
      config = {
        type: 'line',
        data: { labels: kdeResult.points.map(v => r(v)), datasets: kdeDs },
        options: { responsive: true, scales: { y: { beginAtZero: true } } }
      };
    }
    resolvedLabel = 'KDE (density)';
  } else {
    config = { type: 'bar', data: { labels: allNums.map((_, i) => i + 1), datasets: [{ label: firstNumCol || 'Data', data: allNums, backgroundColor: '#4a90d9' }] } };
  }

  if (!config.options) config.options = {};
  config.options.responsive = true;
  config.options.aspectRatio = 1.4;
  if (!config.options.plugins) config.options.plugins = {};
  if (!config.options.scales) config.options.scales = {};

  // Axis titles — infer from column names and chart type
  const noAxisTypes = ['doughnut', 'pie', 'radar', 'polarArea'];
  if (!noAxisTypes.includes(config.type)) {
    const xLabel = _axisLabel('x', type, numericCols, catCols, firstNumCol);
    const yLabel = _axisLabel('y', type, numericCols, catCols, firstNumCol);
    if (xLabel) {
      if (!config.options.scales.x) config.options.scales.x = {};
      config.options.scales.x.title = { display: true, text: xLabel };
    }
    if (yLabel) {
      if (!config.options.scales.y) config.options.scales.y = {};
      config.options.scales.y.title = { display: true, text: yLabel };
    }
  }

  // Pinch-to-zoom plugin
  if (typeof ChartZoom !== 'undefined' || (Chart.registry && Chart.registry.plugins.get('zoom'))) {
    config.options.plugins.zoom = {
      pan: { enabled: true, mode: 'xy' },
      zoom: {
        pinch: { enabled: true },
        wheel: { enabled: true },
        mode: 'xy'
      }
    };
  }

  return { config, resolvedLabel };
}

/* ===== Axis title helper ===== */
function _axisLabel(axis, type, numericCols, catCols, firstNumCol) {
  if (type === 'scatter' || type === 'bubble') {
    return axis === 'x' ? (numericCols[0] || 'X') : (numericCols[1] || 'Y');
  }
  if (type === 'histogram') {
    return axis === 'x' ? (firstNumCol || 'Value') : 'Frequency';
  }
  if (type === 'kde') {
    return axis === 'x' ? 'Value' : 'Density';
  }
  if (type === 'cumulative') {
    return axis === 'x' ? (firstNumCol || 'Value') : 'Cumulative Frequency';
  }
  if (type === 'pareto') {
    return axis === 'x' ? (firstNumCol || 'Bin') : (axis === 'y' ? 'Frequency' : null);
  }
  if (type === 'catbar' || type === 'bar') {
    return axis === 'x' ? (catCols[0] || 'Index') : (numericCols[0] || 'Value');
  }
  if (type === 'line' || type === 'area') {
    return axis === 'x' ? 'Index' : (firstNumCol || 'Value');
  }
  if (type === 'boxplot') {
    return axis === 'x' ? 'Column' : 'Value';
  }
  if (type === 'stem') {
    return axis === 'x' ? (firstNumCol || 'Value') : 'Value';
  }
  return null;
}

function renderSingleChart(id) {
  const entry = chartInstances.find(c => c.id === id);
  if (!entry) return;
  const card = chartsContainer.querySelector(`[data-chart-id="${id}"]`);
  if (!card) return;

  const select = card.querySelector('select');
  const chosenType = select.value;
  entry.type = chosenType;

  if (entry.instance) { entry.instance.destroy(); entry.instance = null; }

  const cols = parsedColumns;
  if (!Object.keys(cols).length) return;

  // Update axis dropdown options in case columns changed
  updateAxisOptions(card);

  // Read axis selections
  const xSel = card.querySelector('.axis-x-select');
  const ySel = card.querySelector('.axis-y-select');
  const axisOverride = {
    x: xSel ? xSel.value || null : null,
    y: ySel ? ySel.value || null : null
  };

  const { config, resolvedLabel } = buildConfig(chosenType, cols, axisOverride);
  const label = card.querySelector('.chart-type-label');
  label.textContent = resolvedLabel;

  const canvas = card.querySelector('canvas');
  entry.instance = new Chart(canvas.getContext('2d'), config);
  setupChartClickHighlight(entry.instance);
}

function renderAllCharts() {
  chartInstances.forEach(c => renderSingleChart(c.id));
}

/* ===== Export CSV ===== */
function exportCSV() {
  if (!Object.keys(parsedColumns).length) return;
  const headers = Object.keys(parsedColumns);
  const maxLen = Math.max(...headers.map(h => parsedColumns[h].length));
  let csv = headers.join(',') + '\n';
  for (let i = 0; i < maxLen; i++) {
    csv += headers.map(h => parsedColumns[h][i] || '').join(',') + '\n';
  }
  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'stats-data.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ===== Export PNG ===== */
function exportPNG() {
  const firstChart = chartInstances.find(c => c.instance);
  if (!firstChart) return;
  const url = firstChart.instance.toBase64Image();
  const a = document.createElement('a');
  a.href = url;
  a.download = 'chart.png';
  a.click();
}

/* ===== Dataset tabs ===== */
function renderDatasetTabs() {
  datasetTabsEl.innerHTML = '';
  Object.keys(savedDatasets).forEach(name => {
    const tab = document.createElement('button');
    tab.className = 'dataset-tab' + (name === activeDataset ? ' active' : '');
    tab.innerHTML = name;
    if (Object.keys(savedDatasets).length > 1) {
      const close = document.createElement('span');
      close.className = 'tab-close';
      close.textContent = '×';
      close.addEventListener('click', e => { e.stopPropagation(); deleteDataset(name); });
      tab.appendChild(close);
    }
    tab.addEventListener('click', () => switchDataset(name));
    datasetTabsEl.appendChild(tab);
  });
}

function switchDataset(name) {
  // save current
  savedDatasets[activeDataset] = snapshotGrid();
  activeDataset = name;
  const snap = savedDatasets[name];
  if (snap) restoreSnapshot(snap);
  else initGrid(10, 2, true);
  renderDatasetTabs();
  saveToLocalStorage();
}

function addDataset() {
  savedDatasets[activeDataset] = snapshotGrid();
  let n = Object.keys(savedDatasets).length + 1;
  let name = `Dataset ${n}`;
  while (savedDatasets[name]) { n++; name = `Dataset ${n}`; }
  savedDatasets[name] = null;
  activeDataset = name;
  initGrid(10, 2, true);
  renderDatasetTabs();
  saveToLocalStorage();
}

function deleteDataset(name) {
  const keys = Object.keys(savedDatasets);
  if (keys.length <= 1) return;
  delete savedDatasets[name];
  if (activeDataset === name) {
    activeDataset = Object.keys(savedDatasets)[0];
    const snap = savedDatasets[activeDataset];
    if (snap) restoreSnapshot(snap); else initGrid(10, 2, true);
  }
  renderDatasetTabs();
  saveToLocalStorage();
}

/* ===== Formula column ===== */
function showFormulaModal() {
  const modal = document.createElement('div');
  modal.className = 'formula-modal';
  modal.innerHTML = `<div class="formula-modal-inner">
    <h3>Add Formula Column</h3>
    <input type="text" id="formula-name" placeholder="Column name" value="Result">
    <input type="text" id="formula-expr" placeholder="e.g. Col1 / Col2">
    <div class="formula-hint">Use column names. Supports: + - * / ** sqrt() log() abs() round() min() max()</div>
    <div class="formula-btns">
      <button class="pill-btn" id="formula-cancel">Cancel</button>
      <button class="pill-btn" id="formula-ok" style="background:var(--primary);color:#fff">Add</button>
    </div>
  </div>`;
  document.body.appendChild(modal);
  modal.querySelector('#formula-cancel').addEventListener('click', () => modal.remove());
  modal.addEventListener('click', e => { if (e.target === modal) modal.remove(); });
  modal.querySelector('#formula-ok').addEventListener('click', () => {
    const name = modal.querySelector('#formula-name').value.trim() || 'Result';
    const expr = modal.querySelector('#formula-expr').value.trim();
    modal.remove();
    if (!expr) return;
    applyFormula(name, expr);
  });
}

function applyFormula(colName, expr) {
  pushUndo();
  // Get current headers and data
  const headers = Array.from(gridHead.querySelectorAll('th:not(.row-num-header) input')).map(i => i.value.trim());
  const bodyRows = gridBody.querySelectorAll('tr');

  // Add new column
  addColumn();
  const newIdx = gridCols - 1;
  const hInputs = gridHead.querySelectorAll('th:not(.row-num-header) input');
  hInputs[newIdx].value = colName;

  // Safe math functions
  const mathFns = { sqrt: Math.sqrt, log: Math.log, abs: Math.abs, round: Math.round, min: Math.min, max: Math.max, pow: Math.pow, exp: Math.exp, ceil: Math.ceil, floor: Math.floor };

  bodyRows.forEach(tr => {
    const cells = tr.querySelectorAll('td:not(.row-num) input');
    // Build variable map from row values
    const vars = {};
    headers.forEach((h, i) => {
      const v = cells[i] ? Number(cells[i].value.trim()) : NaN;
      vars[h] = isNaN(v) ? 0 : v;
    });

    try {
      // Replace column names with values in expression
      let safeExpr = expr;
      // Sort headers by length descending to avoid partial replacements
      const sortedHeaders = [...headers].sort((a, b) => b.length - a.length);
      sortedHeaders.forEach(h => {
        safeExpr = safeExpr.replace(new RegExp(h.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), vars[h]);
      });
      // Replace math function names
      Object.keys(mathFns).forEach(fn => {
        safeExpr = safeExpr.replace(new RegExp(`\\b${fn}\\b`, 'g'), `__fn.${fn}`);
      });
      const result = new Function('__fn', `"use strict"; return (${safeExpr});`)(mathFns);
      if (cells[newIdx] && isFinite(result)) {
        cells[newIdx].value = +result.toFixed(4);
      }
    } catch(e) { /* skip invalid rows */ }
  });
  validateAllCells();
  saveToLocalStorage();
}

/* ===== Hypothesis testing ===== */
// One-sample t-test
function tTest1Sample(nums, mu0) {
  const n = nums.length;
  if (n < 2) return null;
  const mean = nums.reduce((a, b) => a + b, 0) / n;
  const std = Math.sqrt(nums.reduce((s, v) => s + (v - mean) ** 2, 0) / (n - 1));
  const t = (mean - mu0) / (std / Math.sqrt(n));
  const df = n - 1;
  const p = 2 * tDistP(Math.abs(t), df);
  return { test: '1-sample t', t: r(t), df, p: r(p), sig: p < 0.05, mean: r(mean), mu0 };
}

// Two-sample independent t-test (Welch's)
function tTest2Sample(a, b) {
  const n1 = a.length, n2 = b.length;
  if (n1 < 2 || n2 < 2) return null;
  const m1 = a.reduce((s, v) => s + v, 0) / n1;
  const m2 = b.reduce((s, v) => s + v, 0) / n2;
  const v1 = a.reduce((s, v) => s + (v - m1) ** 2, 0) / (n1 - 1);
  const v2 = b.reduce((s, v) => s + (v - m2) ** 2, 0) / (n2 - 1);
  const se = Math.sqrt(v1 / n1 + v2 / n2);
  if (se === 0) return null;
  const t = (m1 - m2) / se;
  const df = Math.floor((v1 / n1 + v2 / n2) ** 2 / ((v1 / n1) ** 2 / (n1 - 1) + (v2 / n2) ** 2 / (n2 - 1)));
  const p = 2 * tDistP(Math.abs(t), df);
  return { test: '2-sample t (Welch)', t: r(t), df, p: r(p), sig: p < 0.05 };
}

// One-way ANOVA
function anova(groupArrays) {
  const k = groupArrays.length;
  if (k < 2) return null;
  const all = groupArrays.flat();
  const N = all.length;
  const grandMean = all.reduce((a, b) => a + b, 0) / N;
  let ssBetween = 0, ssWithin = 0;
  groupArrays.forEach(g => {
    const gMean = g.reduce((a, b) => a + b, 0) / g.length;
    ssBetween += g.length * (gMean - grandMean) ** 2;
    g.forEach(v => { ssWithin += (v - gMean) ** 2; });
  });
  const dfB = k - 1, dfW = N - k;
  if (dfW <= 0) return null;
  const msB = ssBetween / dfB, msW = ssWithin / dfW;
  const F = msW === 0 ? Infinity : msB / msW;
  const p = 1 - fDistCDF(F, dfB, dfW);
  return { test: 'One-way ANOVA', F: r(F), dfB, dfW, p: r(p), sig: p < 0.05 };
}

// Chi-squared goodness of fit (observed vs uniform expected)
function chiSquaredTest(observed) {
  const n = observed.length;
  if (n < 2) return null;
  const total = observed.reduce((a, b) => a + b, 0);
  const expected = total / n;
  let chi2 = 0;
  observed.forEach(o => { chi2 += (o - expected) ** 2 / expected; });
  const df = n - 1;
  const p = 1 - chi2CDF(chi2, df);
  return { test: 'Chi-squared', chi2: r(chi2), df, p: r(p), sig: p < 0.05 };
}

// Approximate t-distribution p-value (one-tail) using normal approx
function tDistP(t, df) {
  // Use normal approximation (works well for all practical df)
  // Abramowitz & Stegun approximation via cube-root transform
  if (df <= 0) return 0.5;
  const A = 1 - 1 / (4 * df) + 1 / (32 * df * df);
  const B = df * (1 + t * t / df);
  const z = Math.abs(t) * (1 - 1 / (4 * df)) / Math.sqrt(t * t / df * (1 + 1 / (2 * df)) + 1) || 0;
  // Simpler: for df > 3, use Wilson-Hilferty
  const v = df;
  const z2 = (Math.pow(Math.abs(t) * Math.abs(t) / v, 1/3) - (1 - 2 / (9 * v))) / Math.sqrt(2 / (9 * v));
  return 0.5 * erfc(z2 / Math.sqrt(2));
}

function erfc(x) {
  const t = 1 / (1 + 0.3275911 * Math.abs(x));
  const poly = t * (0.254829592 + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
  const result = poly * Math.exp(-x * x);
  return x >= 0 ? result : 2 - result;
}

function lgamma(x) {
  // Lanczos approximation
  const c = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  let y = x, tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) ser += c[j] / ++y;
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}

// F-distribution CDF approximation
function fDistCDF(f, d1, d2) {
  if (f <= 0) return 0;
  const x = d1 * f / (d1 * f + d2);
  return regularizedBeta(x, d1 / 2, d2 / 2);
}

// Regularized incomplete beta via continued fraction
function regularizedBeta(x, a, b) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const lnBeta = lgamma(a) + lgamma(b) - lgamma(a + b);
  const front = Math.exp(a * Math.log(x) + b * Math.log(1 - x) - lnBeta);
  // Use continued fraction
  if (x < (a + 1) / (a + b + 2)) {
    return front * betaCF(x, a, b) / a;
  }
  return 1 - front * betaCF(1 - x, b, a) / b;
}

function betaCF(x, a, b) {
  const maxIter = 200;
  const eps = 1e-10;
  let qab = a + b, qap = a + 1, qam = a - 1;
  let c = 1, d = 1 - qab * x / qap;
  if (Math.abs(d) < 1e-30) d = 1e-30;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= maxIter; m++) {
    let m2 = 2 * m;
    let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < 1e-30) d = 1e-30; d = 1 / d;
    c = 1 + aa / c; if (Math.abs(c) < 1e-30) c = 1e-30;
    h *= d * c;
    aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < 1e-30) d = 1e-30; d = 1 / d;
    c = 1 + aa / c; if (Math.abs(c) < 1e-30) c = 1e-30;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < eps) break;
  }
  return h;
}

// Chi-squared CDF via regularized gamma
function chi2CDF(x, df) {
  if (x <= 0) return 0;
  return regularizedGamma(df / 2, x / 2);
}

function regularizedGamma(a, x) {
  if (x < a + 1) {
    // Series expansion
    let sum = 1 / a, term = 1 / a;
    for (let n = 1; n < 200; n++) {
      term *= x / (a + n);
      sum += term;
      if (Math.abs(term) < 1e-10) break;
    }
    return sum * Math.exp(-x + a * Math.log(x) - lgamma(a));
  }
  // Continued fraction
  let f = 1e-30, c = 1e-30, d = x + 1 - a; d = 1 / d;
  let h = d;
  for (let n = 1; n < 200; n++) {
    const an = -n * (n - a);
    const bn = x + 2 * n + 1 - a;
    d = bn + an * d; if (Math.abs(d) < 1e-30) d = 1e-30; d = 1 / d;
    c = bn + an / c; if (Math.abs(c) < 1e-30) c = 1e-30;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-10) break;
  }
  return 1 - h * Math.exp(-x + a * Math.log(x) - lgamma(a));
}

function runHypothesisTests(cols) {
  const names = Object.keys(cols);
  const numericCols = names.filter(n => toNumbers(cols[n]).length > cols[n].length * 0.5);
  const catCols = names.filter(n => !numericCols.includes(n));
  const results = [];

  // One-sample t-tests (H0: mu = 0) for each numeric column
  numericCols.forEach(name => {
    const nums = toNumbers(cols[name]);
    if (nums.length >= 2) {
      const res = tTest1Sample(nums, 0);
      if (res) results.push({ ...res, col: name });
    }
  });

  // If we have a categorical grouping variable + numeric columns → 2-sample t / ANOVA
  if (catCols.length >= 1 && numericCols.length >= 1) {
    const groups = groupByCategory(cols, catCols[0], numericCols);
    const groupKeys = Object.keys(groups);

    numericCols.forEach(nc => {
      const validKeys = groupKeys.filter(gk => toNumbers(groups[gk][nc] || []).length >= 2);
      const groupArrays = validKeys.map(gk => toNumbers(groups[gk][nc] || []));
      // Compute per-group means for explanations
      const groupMeans = {};
      validKeys.forEach((gk, i) => {
        const arr = groupArrays[i];
        groupMeans[gk] = r(arr.reduce((a, b) => a + b, 0) / arr.length);
      });

      if (groupArrays.length === 2) {
        const res = tTest2Sample(groupArrays[0], groupArrays[1]);
        if (res) results.push({ ...res, col: `${nc} by ${catCols[0]}`, variable: nc, catVar: catCols[0], groups: `${validKeys[0]} vs ${validKeys[1]}`, groupMeans });
      }
      if (groupArrays.length >= 2) {
        const res = anova(groupArrays);
        if (res) results.push({ ...res, col: `${nc} by ${catCols[0]}`, variable: nc, catVar: catCols[0], groupKeys: validKeys, groupMeans });
      }
    });

    // Chi-squared on category frequencies
    const freq = {};
    cols[catCols[0]].forEach(v => { freq[v] = (freq[v] || 0) + 1; });
    const observed = Object.values(freq);
    if (observed.length >= 2) {
      const res = chiSquaredTest(observed);
      if (res) results.push({ ...res, col: catCols[0] + ' (frequencies)' });
    }
  }

  return results;
}

function hypExplain(res) {
  const meansStr = res.groupMeans
    ? Object.entries(res.groupMeans).map(([g, m]) => `<b>${g}</b>: ${m}`).join(', ')
    : '';
  if (res.test === '1-sample t') {
    return res.sig
      ? `The mean of <b>${res.col}</b> (${res.mean}) is significantly different from ${res.mu0} (p=${res.p}). There is strong evidence the true average is not ${res.mu0}.`
      : `The mean of <b>${res.col}</b> (${res.mean}) is not significantly different from ${res.mu0} (p=${res.p}). Not enough evidence to reject that the true average equals ${res.mu0}.`;
  }
  if (res.test.startsWith('2-sample')) {
    const base = res.sig
      ? `<b>${res.variable}</b> differs significantly between the two levels of <b>${res.catVar}</b> (p=${res.p}). The difference is unlikely due to chance.`
      : `<b>${res.variable}</b> does not differ significantly between the two levels of <b>${res.catVar}</b> (p=${res.p}). The difference could be random variation.`;
    return meansStr ? `${base}<br>Group means — ${meansStr}` : base;
  }
  if (res.test === 'One-way ANOVA') {
    const base = res.sig
      ? `At least one level of <b>${res.catVar}</b> has a significantly different mean for <b>${res.variable}</b> (p=${res.p}).`
      : `No significant difference in <b>${res.variable}</b> across the levels of <b>${res.catVar}</b> (p=${res.p}). The groups have similar averages.`;
    return meansStr ? `${base}<br>Group means — ${meansStr}` : base;
  }
  if (res.test === 'Chi-squared') {
    return res.sig
      ? `The category frequencies for <b>${res.col}</b> are significantly unequal (p=${res.p}). The distribution is not uniform.`
      : `The category frequencies for <b>${res.col}</b> are not significantly different from uniform (p=${res.p}).`;
  }
  return '';
}

function renderHypothesisTests(results) {
  if (!results.length) { hypothesisSection.classList.add('hidden'); return; }
  hypothesisSection.classList.remove('hidden');
  let html = '<table class="hyp-table"><thead><tr><th>Variable</th><th>Test</th><th>Statistic</th><th>df</th><th>p-value</th><th>Result</th></tr></thead><tbody>';
  results.forEach(res => {
    const stat = res.t !== undefined ? `t=${res.t}` : res.F !== undefined ? `F=${res.F}` : `χ²=${res.chi2}`;
    const df = res.dfB !== undefined ? `${res.dfB},${res.dfW}` : res.df;
    const cls = res.sig ? 'sig' : 'not-sig';
    const explain = hypExplain(res);
    html += `<tr><td>${res.col}</td><td>${res.test}</td><td>${stat}</td><td>${df}</td><td>${res.p}</td><td class="${cls}">${res.sig ? 'Significant' : 'Not sig.'}</td></tr>`;
    html += `<tr class="hyp-explain-row"><td colspan="6">${explain}</td></tr>`;
  });
  html += '</tbody></table>';
  hypothesisContent.innerHTML = html;
}

/* ===== Effect size ===== */
function computeEffectSizes(cols) {
  const names = Object.keys(cols);
  const numericCols = names.filter(n => toNumbers(cols[n]).length > cols[n].length * 0.5);
  const catCols = names.filter(n => !numericCols.includes(n));
  const results = [];

  if (catCols.length < 1 || numericCols.length < 1) return results;

  const groups = groupByCategory(cols, catCols[0], numericCols);
  const groupKeys = Object.keys(groups);
  if (groupKeys.length < 2) return results;

  numericCols.forEach(nc => {
    const groupArrays = groupKeys.map(gk => toNumbers(groups[gk][nc] || [])).filter(a => a.length >= 2);

    // Cohen's d (for 2 groups)
    if (groupArrays.length === 2) {
      const [a, b] = groupArrays;
      const m1 = a.reduce((s, v) => s + v, 0) / a.length;
      const m2 = b.reduce((s, v) => s + v, 0) / b.length;
      const s1 = Math.sqrt(a.reduce((s, v) => s + (v - m1) ** 2, 0) / (a.length - 1));
      const s2 = Math.sqrt(b.reduce((s, v) => s + (v - m2) ** 2, 0) / (b.length - 1));
      const pooled = Math.sqrt(((a.length - 1) * s1 * s1 + (b.length - 1) * s2 * s2) / (a.length + b.length - 2));
      const d = pooled === 0 ? 0 : (m1 - m2) / pooled;
      const mag = Math.abs(d) < 0.2 ? 'negligible' : Math.abs(d) < 0.5 ? 'small' : Math.abs(d) < 0.8 ? 'medium' : 'large';
      results.push({ col: nc, measure: "Cohen's d", value: r(d), magnitude: mag, groups: `${groupKeys[0]} vs ${groupKeys[1]}` });
    }

    // Eta-squared (from ANOVA)
    if (groupArrays.length >= 2) {
      const all = groupArrays.flat();
      const grandMean = all.reduce((a, b) => a + b, 0) / all.length;
      let ssBetween = 0, ssTotal = 0;
      groupArrays.forEach(g => {
        const gMean = g.reduce((a, b) => a + b, 0) / g.length;
        ssBetween += g.length * (gMean - grandMean) ** 2;
      });
      all.forEach(v => { ssTotal += (v - grandMean) ** 2; });
      const eta2 = ssTotal === 0 ? 0 : ssBetween / ssTotal;
      const mag = eta2 < 0.01 ? 'negligible' : eta2 < 0.06 ? 'small' : eta2 < 0.14 ? 'medium' : 'large';
      results.push({ col: nc, measure: 'Eta-squared (η²)', value: r(eta2), magnitude: mag });
    }
  });
  return results;
}

function effectExplain(res) {
  if (res.measure === "Cohen's d") {
    const dir = res.value > 0 ? 'higher' : 'lower';
    const absD = Math.abs(res.value);
    if (res.magnitude === 'negligible') return `The difference between groups (${res.groups}) for <b>${res.col}</b> is negligible (d=${res.value}). The means are practically identical.`;
    if (res.magnitude === 'small') return `A small difference exists between groups (${res.groups}) for <b>${res.col}</b> (d=${res.value}). The first group is slightly ${dir}.`;
    if (res.magnitude === 'medium') return `A moderate difference between groups (${res.groups}) for <b>${res.col}</b> (d=${res.value}). The effect is noticeable and practically meaningful.`;
    return `A large difference between groups (${res.groups}) for <b>${res.col}</b> (d=${res.value}). The groups are clearly separated on this variable.`;
  }
  if (res.measure.includes('Eta')) {
    const pct = (res.value * 100).toFixed(1);
    if (res.magnitude === 'negligible') return `The grouping variable explains only ${pct}% of variance in <b>${res.col}</b> — virtually none.`;
    if (res.magnitude === 'small') return `The grouping variable explains ${pct}% of variance in <b>${res.col}</b> — a small but detectable effect.`;
    if (res.magnitude === 'medium') return `The grouping variable explains ${pct}% of variance in <b>${res.col}</b> — a moderate, practically relevant effect.`;
    return `The grouping variable explains ${pct}% of variance in <b>${res.col}</b> — a large effect. Group membership strongly predicts this variable.`;
  }
  return '';
}

function renderEffectSizes(results) {
  if (!results.length) { effectSizeSection.classList.add('hidden'); return; }
  effectSizeSection.classList.remove('hidden');
  let html = '<table class="effect-table"><thead><tr><th>Variable</th><th>Measure</th><th>Value</th><th>Magnitude</th></tr></thead><tbody>';
  results.forEach(res => {
    const explain = effectExplain(res);
    html += `<tr><td>${res.col}</td><td>${res.measure}</td><td>${res.value}</td><td>${res.magnitude}</td></tr>`;
    html += `<tr class="hyp-explain-row"><td colspan="4">${explain}</td></tr>`;
  });
  html += '</tbody></table>';
  effectSizeContent.innerHTML = html;
}

/* ===== Descriptive stats per group ===== */
function renderGroupStats(cols) {
  const names = Object.keys(cols);
  const numericCols = names.filter(n => toNumbers(cols[n]).length > cols[n].length * 0.5);
  const catCols = names.filter(n => !numericCols.includes(n));
  if (!catCols.length || !numericCols.length) { groupStatsSection.classList.add('hidden'); return; }

  const groups = groupByCategory(cols, catCols[0], numericCols);
  const groupKeys = Object.keys(groups);
  if (groupKeys.length < 2) { groupStatsSection.classList.add('hidden'); return; }

  groupStatsSection.classList.remove('hidden');
  const statKeys = ['Count', 'Mean', 'Median', 'Std Dev', 'Min', 'Max'];
  let html = '';

  numericCols.forEach(nc => {
    html += `<h3 style="font-size:0.85rem;margin:8px 0 4px;color:var(--primary)">${nc}</h3>`;
    html += '<table class="group-stats-table"><thead><tr><th>Group</th>';
    statKeys.forEach(k => { html += `<th>${k}</th>`; });
    html += '</tr></thead><tbody>';
    groupKeys.forEach(gk => {
      const nums = toNumbers(groups[gk][nc] || []);
      if (!nums.length) return;
      const s = computeStats(nums);
      html += `<tr><td style="font-weight:600">${gk}</td>`;
      statKeys.forEach(k => { html += `<td>${s[k]}</td>`; });
      html += '</tr>';
    });
    html += '</tbody></table>';
  });
  groupStatsContent.innerHTML = html;
}

/* ===== Polynomial regression ===== */
function polyFit(xs, ys, degree) {
  // Fit polynomial of given degree using normal equations
  const n = Math.min(xs.length, ys.length);
  if (n <= degree) return null;
  const size = degree + 1;
  // Build Vandermonde-like system: X^T X a = X^T y
  const XtX = Array.from({ length: size }, () => Array(size).fill(0));
  const XtY = Array(size).fill(0);
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < size; j++) {
      XtY[j] += Math.pow(xs[i], j) * ys[i];
      for (let k = 0; k < size; k++) {
        XtX[j][k] += Math.pow(xs[i], j + k);
      }
    }
  }
  // Gaussian elimination
  const aug = XtX.map((row, i) => [...row, XtY[i]]);
  for (let col = 0; col < size; col++) {
    let maxRow = col;
    for (let row = col + 1; row < size; row++) {
      if (Math.abs(aug[row][col]) > Math.abs(aug[maxRow][col])) maxRow = row;
    }
    [aug[col], aug[maxRow]] = [aug[maxRow], aug[col]];
    if (Math.abs(aug[col][col]) < 1e-12) return null;
    for (let row = col + 1; row < size; row++) {
      const f = aug[row][col] / aug[col][col];
      for (let j = col; j <= size; j++) aug[row][j] -= f * aug[col][j];
    }
  }
  const coeffs = Array(size).fill(0);
  for (let i = size - 1; i >= 0; i--) {
    coeffs[i] = aug[i][size];
    for (let j = i + 1; j < size; j++) coeffs[i] -= aug[i][j] * coeffs[j];
    coeffs[i] /= aug[i][i];
  }
  // R²
  const yMean = ys.reduce((a, b) => a + b, 0) / n;
  let ssTot = 0, ssRes = 0;
  for (let i = 0; i < n; i++) {
    let yPred = 0;
    coeffs.forEach((c, j) => { yPred += c * Math.pow(xs[i], j); });
    ssTot += (ys[i] - yMean) ** 2;
    ssRes += (ys[i] - yPred) ** 2;
  }
  const r2 = ssTot === 0 ? 1 : 1 - ssRes / ssTot;
  return { coeffs, r2: +r2.toFixed(4) };
}

function polyLabel(coeffs) {
  return coeffs.map((c, i) => {
    const cv = +c.toFixed(3);
    if (i === 0) return cv;
    if (i === 1) return `${cv}x`;
    return `${cv}x²${i > 2 ? `^${i}` : ''}`;
  }).reverse().join(' + ').replace(/\+ -/g, '- ');
}

/* ===== Moving average ===== */
function movingAverage(nums, window) {
  if (!window || window < 2) window = Math.max(3, Math.floor(nums.length / 10));
  const result = [];
  for (let i = 0; i < nums.length; i++) {
    const start = Math.max(0, i - Math.floor(window / 2));
    const end = Math.min(nums.length, start + window);
    const slice = nums.slice(start, end);
    result.push(slice.reduce((a, b) => a + b, 0) / slice.length);
  }
  return result;
}

/* ===== Normal PDF overlay for histograms ===== */
function normalPDF(x, mean, std) {
  return Math.exp(-0.5 * ((x - mean) / std) ** 2) / (std * Math.sqrt(2 * Math.PI));
}

/* ===== Interactive chart → grid row highlight ===== */
function setupChartClickHighlight(chartInstance) {
  const canvas = chartInstance.canvas;
  canvas.addEventListener('click', e => {
    const points = chartInstance.getElementsAtEventForMode(e, 'nearest', { intersect: true }, false);
    // Clear previous highlights
    gridBody.querySelectorAll('tr.highlight-row').forEach(tr => tr.classList.remove('highlight-row'));
    if (!points.length) return;
    const idx = points[0].index;
    const rows = gridBody.querySelectorAll('tr');
    if (rows[idx]) {
      rows[idx].classList.add('highlight-row');
      rows[idx].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  });
}

/* ===== Dark mode ===== */
function initDarkMode() {
  const saved = localStorage.getItem('theme');
  if (saved) {
    document.documentElement.setAttribute('data-theme', saved);
  } else if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
    document.documentElement.setAttribute('data-theme', 'dark');
  }
}

function toggleDarkMode() {
  const current = document.documentElement.getAttribute('data-theme');
  const next = current === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('theme', next);
}

/* ===== Animated transitions ===== */
Chart.defaults.animation.duration = 600;
Chart.defaults.animation.easing = 'easeInOutQuart';

/* ===== Relationship Analysis (EDA) ===== */
let scatterMatrixCharts = [];

function renderScatterMatrix(cols) {
  // Destroy previous charts
  scatterMatrixCharts.forEach(c => c.destroy());
  scatterMatrixCharts = [];
  scatterMatrixEl.innerHTML = '';

  const names = Object.keys(cols);
  const numCols = names.filter(n => toNumbers(cols[n]).length > cols[n].length * 0.5);
  const selected = numCols.slice(0, 6);
  const n = selected.length;
  if (n < 2) return;

  scatterMatrixEl.style.gridTemplateColumns = `repeat(${n}, 150px)`;

  const numData = {};
  selected.forEach(name => { numData[name] = toNumbers(cols[name]); });

  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const axisColor = isDark ? '#888' : '#aaa';
  const pointColor = isDark ? 'rgba(90,159,230,0.5)' : 'rgba(74,144,217,0.5)';

  for (let row = 0; row < n; row++) {
    for (let col = 0; col < n; col++) {
      if (row === col) {
        // Diagonal: label
        const label = document.createElement('div');
        label.className = 'scatter-matrix-label';
        label.textContent = selected[row];
        scatterMatrixEl.appendChild(label);
        continue;
      }
      const canvas = document.createElement('canvas');
      canvas.width = 150;
      canvas.height = 150;
      scatterMatrixEl.appendChild(canvas);

      const xData = numData[selected[col]];
      const yData = numData[selected[row]];
      const data = [];
      const len = Math.min(xData.length, yData.length);
      for (let i = 0; i < len; i++) data.push({ x: xData[i], y: yData[i] });

      const chart = new Chart(canvas, {
        type: 'scatter',
        data: { datasets: [{ data, backgroundColor: pointColor, pointRadius: 2 }] },
        options: {
          responsive: false,
          plugins: { legend: { display: false }, zoom: { zoom: { wheel: { enabled: false } }, pan: { enabled: false } } },
          scales: {
            x: { display: row === n - 1, ticks: { font: { size: 8 }, color: axisColor, maxTicksLimit: 3 }, grid: { display: false } },
            y: { display: col === 0, ticks: { font: { size: 8 }, color: axisColor, maxTicksLimit: 3 }, grid: { display: false } }
          },
          animation: false
        }
      });
      scatterMatrixCharts.push(chart);
    }
  }
}

function renderParallelCoords(cols) {
  const names = Object.keys(cols);
  const numCols = names.filter(n => toNumbers(cols[n]).length > cols[n].length * 0.5);
  const catCols = names.filter(n => !numCols.includes(n));
  const selected = numCols.slice(0, 6);
  if (selected.length < 2) return;

  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const catCol = catCols[0] || null;

  // Build row data
  const numData = {};
  selected.forEach(name => { numData[name] = cols[name].map(Number); });
  const rowCount = cols[selected[0]].length;

  // Category colors
  const categories = catCol ? [...new Set(cols[catCol])] : ['all'];
  const palette = ['#4a90d9', '#e55c5c', '#34c759', '#f5a623', '#9b59b6', '#1abc9c', '#e67e22', '#3498db'];
  const catColorMap = {};
  categories.forEach((c, i) => { catColorMap[c] = palette[i % palette.length]; });

  // Min/max per axis
  const mins = {}, maxs = {};
  selected.forEach(name => {
    const vals = numData[name].filter(v => !isNaN(v));
    mins[name] = Math.min(...vals);
    maxs[name] = Math.max(...vals);
    if (mins[name] === maxs[name]) { mins[name] -= 1; maxs[name] += 1; }
  });

  const canvas = parallelCoordsCanvas;
  const w = Math.max(500, selected.length * 120);
  const h = 300;
  canvas.width = w;
  canvas.height = h;
  canvas.style.height = h + 'px';

  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, w, h);

  const padTop = 30, padBottom = 30, padLeft = 40, padRight = 40;
  const plotH = h - padTop - padBottom;
  const axisSpacing = (w - padLeft - padRight) / (selected.length - 1);

  const textColor = isDark ? '#e5e5e7' : '#333';
  const axisLineColor = isDark ? '#555' : '#ccc';

  // Draw axes
  ctx.strokeStyle = axisLineColor;
  ctx.lineWidth = 1;
  selected.forEach((name, i) => {
    const x = padLeft + i * axisSpacing;
    ctx.beginPath();
    ctx.moveTo(x, padTop);
    ctx.lineTo(x, padTop + plotH);
    ctx.stroke();
    // Label
    ctx.fillStyle = textColor;
    ctx.font = 'bold 10px -apple-system, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(name.slice(0, 10), x, padTop - 8);
    // Min/max
    ctx.font = '9px -apple-system, sans-serif';
    ctx.fillText(maxs[name].toFixed(1), x, padTop - 0);
    ctx.fillText(mins[name].toFixed(1), x, padTop + plotH + 14);
  });

  // Draw lines
  ctx.lineWidth = 1;
  ctx.globalAlpha = 0.3;
  for (let r = 0; r < rowCount; r++) {
    const cat = catCol ? cols[catCol][r] : 'all';
    ctx.strokeStyle = catColorMap[cat] || '#4a90d9';
    ctx.beginPath();
    let started = false;
    selected.forEach((name, i) => {
      const val = numData[name][r];
      if (isNaN(val)) return;
      const x = padLeft + i * axisSpacing;
      const t = (val - mins[name]) / (maxs[name] - mins[name]);
      const y = padTop + plotH - t * plotH;
      if (!started) { ctx.moveTo(x, y); started = true; }
      else ctx.lineTo(x, y);
    });
    ctx.stroke();
  }
  ctx.globalAlpha = 1.0;

  // Legend
  if (catCol && categories.length > 1) {
    let lx = padLeft;
    ctx.font = '10px -apple-system, sans-serif';
    categories.forEach(cat => {
      ctx.fillStyle = catColorMap[cat];
      ctx.fillRect(lx, h - 14, 10, 10);
      ctx.fillStyle = textColor;
      ctx.fillText(cat, lx + 13, h - 5);
      lx += ctx.measureText(cat).width + 24;
    });
  }
}

function renderEDA(cols) {
  const names = Object.keys(cols);
  const numCols = names.filter(n => toNumbers(cols[n]).length > cols[n].length * 0.5);
  if (numCols.length < 2) {
    edaSection.classList.add('hidden');
    return;
  }
  edaSection.classList.remove('hidden');
  try { renderScatterMatrix(cols); } catch(e) { console.warn('Scatter matrix error:', e); }
  try { renderParallelCoords(cols); } catch(e) { console.warn('Parallel coords error:', e); }
}

/* ===== Main ===== */
analyzeBtn.addEventListener('click', () => {
  parsedColumns = getGridData();
  if (!Object.keys(parsedColumns).length) return;

  validateAllCells();

  // Stats
  statsGrid.innerHTML = '';
  statsSection.classList.remove('hidden');
  const allStats = {};
  for (const name of Object.keys(parsedColumns)) {
    const nums = toNumbers(parsedColumns[name]);
    if (!nums.length) continue;
    const stats = computeStats(nums);
    if (!stats) continue;
    allStats[name] = stats;

    const keys = ['Count','Mean','Median','Mode','Std Dev','Variance',
      'Min','Max','Range','Q1','Q3','IQR','Skewness','Kurtosis',
      'Normality','CI 95%','Outliers'];

    let html = '<div class="stats-column-block">';
    const multiCol = Object.keys(parsedColumns).length > 1;
    if (multiCol) html += `<div class="stats-col-title">${name}</div>`;
    html += '<div class="stats-compact">';
    keys.forEach(k => {
      if (stats[k] === undefined) return;
      html += `<div class="stats-row"><span class="s-key">${k}</span><span class="s-val">${stats[k]}</span></div>`;
    });
    html += '</div></div>';
    statsGrid.innerHTML += html;
  }

  // Outlier highlighting
  try { highlightOutliers(allStats); } catch(e) { console.warn('Outlier highlight error:', e); }

  // Correlation
  let corrData = null;
  try {
    corrData = buildCorrelationMatrix(parsedColumns);
    renderCorrelationTable(corrData);
  } catch(e) { console.warn('Correlation error:', e); }

  // EDA relationship charts
  try { renderEDA(parsedColumns); } catch(e) { console.warn('EDA error:', e); }

  // Insights
  try {
    insightsContent.innerHTML = generateInsights(allStats, corrData);
    insightsSection.classList.remove('hidden');
  } catch(e) { console.warn('Insights error:', e); }

  // Hypothesis testing
  try {
    const hypResults = runHypothesisTests(parsedColumns);
    renderHypothesisTests(hypResults);
  } catch(e) { console.warn('Hypothesis tests error:', e); }

  // Group stats
  try {
    renderGroupStats(parsedColumns);
  } catch(e) { console.warn('Group stats error:', e); }

  // Charts
  chartsSection.classList.remove('hidden');
  if (!chartInstances.length) {
    addChartCard('auto');
  } else {
    renderAllCharts();
  }
});

/* ===== Toolbar buttons ===== */
document.getElementById('add-col-btn').addEventListener('click', addColumn);
document.getElementById('add-row-btn').addEventListener('click', addRow);
document.getElementById('clear-btn').addEventListener('click', clearGrid);
document.getElementById('add-chart-btn').addEventListener('click', () => addChartCard('auto'));
document.getElementById('undo-btn').addEventListener('click', undo);
document.getElementById('redo-btn').addEventListener('click', redo);
document.getElementById('export-csv-btn').addEventListener('click', exportCSV);
document.getElementById('export-png-btn').addEventListener('click', exportPNG);
document.getElementById('dark-toggle').addEventListener('click', toggleDarkMode);
document.getElementById('test-data-btn').addEventListener('click', loadTestData);
document.getElementById('add-formula-btn').addEventListener('click', showFormulaModal);
document.getElementById('add-dataset-btn').addEventListener('click', addDataset);

/* ===== Synthetic test data ===== */
function loadTestData() {
  pushUndo();

  // Seeded pseudo-random (simple LCG so data is reproducible)
  let seed = 42;
  function rand() { seed = (seed * 1664525 + 1013904223) & 0x7fffffff; return seed / 0x7fffffff; }
  function randNorm(mu, sigma) {
    // Box-Muller
    const u1 = rand() || 0.001;
    const u2 = rand();
    return mu + sigma * Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
  }

  const groups = ['Alpha', 'Beta', 'Gamma'];
  const n = 60; // 20 per group
  const headers = ['Group', 'Height', 'Weight', 'Score'];
  const rows = [];

  for (let g = 0; g < groups.length; g++) {
    for (let i = 0; i < 20; i++) {
      const height = +(randNorm(170 + g * 5, 8)).toFixed(1);
      const weight = +(0.6 * height - 30 + randNorm(0, 5)).toFixed(1);
      let score = +(randNorm(70 + g * 10, 12)).toFixed(1);
      // inject a couple of outliers per group
      if (i === 0) score = +(score + 45).toFixed(1);
      if (i === 19) score = +(score - 40).toFixed(1);
      rows.push([groups[g], String(height), String(weight), String(score)]);
    }
  }

  initGrid(rows.length, headers.length, true);

  // Set headers
  const hInputs = gridHead.querySelectorAll('th:not(.row-num-header) input');
  headers.forEach((h, i) => { if (hInputs[i]) hInputs[i].value = h; });

  // Fill data
  const bodyRows = gridBody.querySelectorAll('tr');
  rows.forEach((row, ri) => {
    if (!bodyRows[ri]) return;
    const cells = bodyRows[ri].querySelectorAll('td:not(.row-num) input');
    row.forEach((v, ci) => { if (cells[ci]) cells[ci].value = v; });
  });

  validateAllCells();
}

/* ===== Keyboard shortcuts ===== */
document.addEventListener('keydown', e => {
  const mod = e.metaKey || e.ctrlKey;
  if (mod && !e.shiftKey && e.key === 'z') { e.preventDefault(); undo(); }
  if (mod && e.shiftKey && (e.key === 'z' || e.key === 'Z')) { e.preventDefault(); redo(); }
});

/* ===== Init ===== */
initDarkMode();
if (!loadFromLocalStorage()) {
  initGrid(10, 2, true);
}
// Init dataset tabs
if (!Object.keys(savedDatasets).length) {
  savedDatasets[activeDataset] = snapshotGrid();
}
renderDatasetTabs();

/* ===== Service Worker ===== */
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('./sw.js');
}
