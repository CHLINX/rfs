(function () {
  'use strict';

  const state = { records: [], fileName: '', sourceSheet: '', rawCount: 0, result: null, usingSample: false };
  const $ = (id) => document.getElementById(id);
  const fmt = (n) => Number(n || 0).toLocaleString('zh-TW', { maximumFractionDigits: 8 });
  const showToast = (message) => {
    const el = $('toast'); el.textContent = message; el.classList.add('show');
    clearTimeout(showToast.timer); showToast.timer = setTimeout(() => el.classList.remove('show'), 3000);
  };

  function cleanText(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\u00a0/g, ' ').trim();
  }
  function normalizeHeader(value) {
    return cleanText(value).toLowerCase().replace(/[\s　_\-:：()（）【】\[\]、，,]/g, '');
  }
  function parseNumber(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const text = cleanText(value).replace(/,/g, '').replace(/[，]/g, '').replace(/公頃|平方公尺|㎡|m²/gi, '').replace(/[^0-9.+\-eE]/g, '');
    if (!text || text === '.' || text === '-' || text === '+') return 0;
    const num = Number(text);
    return Number.isFinite(num) ? num : 0;
  }
  function keyOf(record) {
    return [record.serial, record.section, record.landNo].map(v => cleanText(v)).join('\u001f');
  }
  function aggregateRecords(inputRows) {
    const map = new Map();
    let acceptedRows = 0;
    for (const row of inputRows) {
      const serial = cleanText(row.serial), section = cleanText(row.section), landNo = cleanText(row.landNo);
      if (!serial || !section || !landNo) continue;
      acceptedRows += 1;
      const record = { serial, section, landNo, unitArea: parseNumber(row.unitArea), contractArea: parseNumber(row.contractArea) };
      const key = keyOf(record);
      if (map.has(key)) map.get(key).contractArea += record.contractArea;
      else map.set(key, record);
    }
    return { records: Array.from(map.values()), acceptedRows };
  }

  function findHeader(rows) {
    const aliases = {
      serial: ['流水號', '流水序號'],
      section: ['土地座落地段', '土地座落', '地段'],
      landNo: ['地號'],
      unitArea: ['本筆面積'],
      contractArea: ['契作面積公頃', '契作面積']
    };
    for (let r = 0; r < rows.length; r += 1) {
      const values = rows[r] || [];
      const normalized = values.map(normalizeHeader);
      const indexes = {};
      for (const [field, names] of Object.entries(aliases)) {
        indexes[field] = normalized.findIndex(cell => names.some(alias => {
          const a = normalizeHeader(alias);
          if (field === 'section') return cell === a || (a === '土地座落地段' && cell.includes(a));
          if (field === 'contractArea') return cell === a || (a === '契作面積' && cell === '契作面積');
          return cell === a;
        }));
      }
      if (Object.values(indexes).every(i => i >= 0)) return { rowIndex: r, indexes };
    }
    return null;
  }

  function extractWorkbook(workbook) {
    let found = null;
    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      if (!ws || !ws['!ref']) continue;
      const range = XLSX.utils.decode_range(ws['!ref']);
      const rows = [];
      for (let rr = range.s.r; rr <= range.e.r; rr += 1) {
        const line = [];
        for (let cc = range.s.c; cc <= range.e.c; cc += 1) {
          const cell = ws[XLSX.utils.encode_cell({ r: rr, c: cc })];
          line.push(cell ? (cell.w !== undefined && cell.w !== '' ? cell.w : cell.v) : null);
        }
        rows.push(line);
      }
      const header = findHeader(rows);
      if (header) { found = { sheetName, ws, rows, header, range }; break; }
    }
    if (!found) throw new Error('找不到同時包含五項必要欄位的標題列。請確認欄位名稱及清冊工作表。');

    const { ws, rows, header, range } = found;
    const { indexes, rowIndex } = header;
    const actualCol = (col) => range.s.c + col;
    const entries = [];
    let carrySerial = '', carrySection = '';
    let formulaBlankCount = 0;
    for (let r = rowIndex + 1; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const actualRow = range.s.r + r;
      const rowText = row.map(cleanText).join(' ');
      if (/小\s*計|合\s*計|總\s*計/.test(rowText)) {
        carrySerial = ''; carrySection = '';
        continue;
      }
      const cellValue = (col) => {
        if (col < 0) return '';
        const cell = ws[XLSX.utils.encode_cell({ r: actualRow, c: actualCol(col) })];
        if (!cell) return row[col] === null || row[col] === undefined ? '' : row[col];
        return cell.w !== undefined && cell.w !== '' ? cell.w : cell.v;
      };
      const serialCell = cleanText(cellValue(indexes.serial));
      const sectionCell = cleanText(cellValue(indexes.section));
      if (serialCell) carrySerial = serialCell;
      if (sectionCell) carrySection = sectionCell;
      const landNo = cleanText(cellValue(indexes.landNo));
      if (!landNo) continue;
      const unitCell = ws[XLSX.utils.encode_cell({ r: actualRow, c: actualCol(indexes.unitArea) })];
      const contractCell = ws[XLSX.utils.encode_cell({ r: actualRow, c: actualCol(indexes.contractArea) })];
      const unitArea = unitCell && unitCell.v !== undefined ? unitCell.v : row[indexes.unitArea];
      const contractArea = contractCell && contractCell.v !== undefined ? contractCell.v : row[indexes.contractArea];
      if (contractCell && contractCell.f && (contractCell.v === undefined || contractCell.v === null || contractCell.v === '')) formulaBlankCount += 1;
      entries.push({ serial: carrySerial, section: carrySection, landNo, unitArea, contractArea });
    }
    const aggregated = aggregateRecords(entries);
    if (!aggregated.records.length) throw new Error('已找到標題列，但標題列下方沒有可辨識的土地資料。');
    return { ...aggregated, sheetName: found.sheetName, formulaBlankCount };
  }

  function randomIndex(length, rng) {
    if (!length) throw new Error('沒有可供抽選的土地。');
    if (rng) return Math.floor(rng() * length);
    if (window.crypto && window.crypto.getRandomValues) {
      const limit = Math.floor(0x100000000 / length) * length;
      const buf = new Uint32Array(1);
      let value;
      do { window.crypto.getRandomValues(buf); value = buf[0]; } while (value >= limit);
      return value % length;
    }
    return Math.floor(Math.random() * length);
  }
  function takeRandom(list, rng) {
    const index = randomIndex(list.length, rng);
    return list.splice(index, 1)[0];
  }

  function runSelection(records, config, rng) {
    const minSerials = Number(config.minSerials);
    const maxPerSerial = Number(config.maxPerSerial);
    const totalCount = Number(config.totalCount);
    for (const [name, value] of Object.entries({ '最低流水號數': minSerials, '每號最多土地數': maxPerSerial, '總抽選數': totalCount })) {
      if (!Number.isInteger(value) || value < 1) throw new Error(`${name}必須是大於或等於 1 的整數。`);
    }
    const groups = new Map();
    records.forEach(record => {
      const serial = cleanText(record.serial);
      if (!groups.has(serial)) groups.set(serial, []);
      groups.get(serial).push({ ...record });
    });
    const availableSerials = Array.from(groups.keys()).filter(serial => groups.get(serial).length > 0);
    if (availableSerials.length < minSerials) throw new Error(`清冊只有 ${availableSerials.length} 個可抽選流水號，無法達成至少 ${minSerials} 個流水號的條件。`);
    if (totalCount < minSerials) throw new Error(`總抽選數 ${totalCount} 筆小於最低流水號數 ${minSerials} 個，無法讓第一批每個流水號至少抽中 1 筆。請調高總抽選數或降低最低流水號數。`);

    const remaining = new Map(Array.from(groups, ([serial, items]) => [serial, items.slice()]));
    const participating = new Set();
    const representatives = new Map();
    const selected = [];
    const perSerialCount = new Map();
    const log = [];

    // 第一批：從尚未參與的所有土地中隨機取代表土地，形成指定數量的流水號。
    for (let i = 0; i < minSerials; i += 1) {
      const pool = [];
      for (const serial of availableSerials) {
        if (participating.has(serial)) continue;
        for (let j = 0; j < remaining.get(serial).length; j += 1) pool.push({ serial, index: j });
      }
      if (!pool.length) break;
      const choice = pool[randomIndex(pool.length, rng)];
      const representative = remaining.get(choice.serial).splice(choice.index, 1)[0];
      participating.add(choice.serial);
      representatives.set(choice.serial, representative);
      perSerialCount.set(choice.serial, 1);
      selected.push(representative);
      log.push(`第一批流水號 ${i + 1}/${minSerials}：由土地「${representative.section} ${representative.landNo}」隨機納入流水號 ${choice.serial}。`);
    }

    function completeSerial(serial, isFirstBatch) {
      const rep = representatives.get(serial);
      const already = perSerialCount.get(serial) || 0;
      const targetForSerial = Math.min(maxPerSerial, (groups.get(serial) || []).length);
      let additional = 0;
      while (selected.length < totalCount && (perSerialCount.get(serial) || 0) < targetForSerial && remaining.get(serial).length) {
        const record = takeRandom(remaining.get(serial), rng);
        selected.push(record);
        perSerialCount.set(serial, (perSerialCount.get(serial) || 0) + 1);
        additional += 1;
      }
      const actual = perSerialCount.get(serial) || already;
      const reachedTotal = selected.length >= totalCount;
      const condition = actual >= maxPerSerial ? `達每號上限 ${maxPerSerial} 筆`
        : actual >= targetForSerial ? '該流水號土地已抽選完畢'
        : reachedTotal ? `總抽選數已達 ${totalCount} 筆` : '該流水號可抽土地已用盡';
      if (reachedTotal) log.push(`流水號 ${serial} 抽中 ${actual} 筆（${condition}），立即停止。`);
      else log.push(`${isFirstBatch ? '第一批完成' : '新增流水號完成'}：${serial} 抽中 ${actual} 筆（${condition}）。`);
      return additional;
    }

    // 第一批流水號全部形成後，才逐一抽選各流水號其餘土地。
    for (const serial of Array.from(participating)) {
      if (selected.length >= totalCount) break;
      completeSerial(serial, true);
    }

    // 第一批完成仍不足時，每次只新增一個尚未參與的流水號。
    while (selected.length < totalCount) {
      const pool = [];
      for (const serial of availableSerials) {
        if (participating.has(serial)) continue;
        const items = remaining.get(serial);
        for (let j = 0; j < items.length; j += 1) pool.push({ serial, index: j });
      }
      if (!pool.length) break;
      const choice = pool[randomIndex(pool.length, rng)];
      const representative = remaining.get(choice.serial).splice(choice.index, 1)[0];
      participating.add(choice.serial);
      representatives.set(choice.serial, representative);
      perSerialCount.set(choice.serial, 1);
      selected.push(representative);
      log.push(`新增流水號：由土地「${representative.section} ${representative.landNo}」隨機納入流水號 ${choice.serial}。`);
      if (selected.length < totalCount) completeSerial(choice.serial, false);
      else log.push(`流水號 ${choice.serial} 已納入 1 筆；總數已達 ${totalCount} 筆，立即停止。`);
    }

    const keys = selected.map(keyOf);
    const unique = new Set(keys);
    if (unique.size !== selected.length) throw new Error('抽選結果檢核失敗：發現重複土地紀錄，已停止匯出。');
    const sumBySerial = Array.from(perSerialCount.values()).reduce((a, b) => a + b, 0);
    if (sumBySerial !== selected.length) throw new Error('抽選結果檢核失敗：各流水號抽選數合計不一致。');
    if (selected.length > totalCount) throw new Error('抽選結果檢核失敗：實際抽選數超過設定總數。');
    const shortfall = Math.max(0, totalCount - selected.length);
    return {
      selected, log, participatingSerials: participating.size, perSerialCount,
      requested: totalCount, actual: selected.length, shortfall,
      status: shortfall ? '不足額' : '已達標', checks: { unique: unique.size === selected.length, perSerialSum: sumBySerial === selected.length, noOverdraw: selected.length <= totalCount }
    };
  }

  function makeSampleRows() {
    const sections = ['松柏段', '新興段', '東安段', '富田段', '仁和段', '大同段', '永豐段'];
    const rows = [];
    for (let s = 0; s < 7; s += 1) {
      for (let p = 1; p <= 6; p += 1) {
        rows.push({ serial: `R${String(s + 1).padStart(3, '0')}`, section: sections[s], landNo: `${String(120 + s * 20 + p).padStart(4, '0')}-0000`, unitArea: 0.045 + p * 0.006, contractArea: 0.018 + p * 0.003 });
      }
    }
    // 故意加入一列重複唯一鍵，示範契作面積會先加總。
    rows.push({ ...rows[2], contractArea: 0.012 });
    return rows;
  }

  function formatArea(value) {
    return Number(value || 0).toLocaleString('zh-TW', { maximumFractionDigits: 12 });
  }
  function renderTable(targetId, records, limit) {
    const tbody = $(targetId); tbody.innerHTML = '';
    const visible = records.slice(0, limit || 50);
    for (const record of visible) {
      const tr = document.createElement('tr');
      [record.serial, record.section, record.landNo, formatArea(record.unitArea), formatArea(record.contractArea)].forEach(value => {
        const td = document.createElement('td'); td.textContent = String(value); tr.appendChild(td);
      });
      tbody.appendChild(tr);
    }
    return visible.length;
  }
  function renderSource() {
    $('dataOverview').hidden = false;
    $('sourceSheet').textContent = `來源工作表：${state.sourceSheet || '內建範例'}；已辨識並完成土地彙整。`;
    $('rawCount').textContent = fmt(state.rawCount);
    $('uniqueCount').textContent = fmt(state.records.length);
    $('serialCount').textContent = fmt(new Set(state.records.map(r => r.serial)).size);
    const shown = renderTable('sourcePreview', state.records, 8);
    $('previewFoot').textContent = `目前預覽 ${shown} 筆／共 ${fmt(state.records.length)} 筆彙整後土地。`;
    $('runDraw').disabled = false;
  }
  function renderResult(result) {
    state.result = result;
    $('resultPanel').hidden = false;
    $('resultDescription').textContent = result.status === '已達標' ? `已抽選 ${result.actual} 筆土地，符合設定總數。` : `可抽土地已用盡，保留已抽中的 ${result.actual} 筆結果。`;
    const metrics = [
      ['設定抽選數', result.requested, '筆'], ['實際抽選數', result.actual, '筆'], ['納入流水號', result.participatingSerials, '個'], ['尚缺筆數', result.shortfall, '筆']
    ];
    $('resultMetrics').innerHTML = '';
    metrics.forEach(([label, value, unit]) => {
      const div = document.createElement('div'); div.className = 'result-metric';
      const small = document.createElement('small'); small.textContent = label;
      const strong = document.createElement('strong'); strong.textContent = fmt(value) + unit;
      div.append(small, strong); $('resultMetrics').appendChild(div);
    });
    $('resultWarning').hidden = !result.shortfall;
    $('resultWarning').textContent = result.shortfall ? `不足額提醒：設定 ${result.requested} 筆，實際抽選 ${result.actual} 筆，差額 ${result.shortfall} 筆。已抽選結果仍可下載。` : '';
    const log = $('drawLog'); log.innerHTML = '';
    result.log.forEach(item => { const li = document.createElement('li'); li.textContent = item; log.appendChild(li); });
    const shown = renderTable('resultPreview', result.selected, 100);
    $('resultFoot').textContent = `目前預覽 ${shown} 筆／實際抽選 ${fmt(result.actual)} 筆。唯一鍵重複檢核、各流水號合計及未超抽檢核均通過。`;
    $('resultPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name)) { setNotice('檔案格式不符，請選擇 Excel 檔案。', 'error'); return; }
    if (!window.XLSX) { setNotice('Excel 讀取元件載入失敗。請確認網路連線後重新整理頁面。', 'error'); return; }
    try {
      setNotice('正在讀取 Excel 並搜尋所有工作表……');
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, cellNF: true, cellText: true });
      const extracted = extractWorkbook(workbook);
      state.records = extracted.records; state.rawCount = extracted.acceptedRows; state.fileName = file.name;
      state.sourceSheet = extracted.sheetName; state.usingSample = false; state.result = null;
      $('resultPanel').hidden = true;
      $('fileName').textContent = file.name;
      $('fileMeta').textContent = `${(file.size / 1024).toFixed(1)} KB · 已完成讀取`;
      renderSource();
      const formulaNote = extracted.formulaBlankCount ? `；另有 ${extracted.formulaBlankCount} 個公式面積儲存格沒有快取值，請留意來源檔案是否已重新計算並儲存。` : '';
      setNotice(`讀取成功：${extracted.sheetName}，來源有效資料 ${extracted.acceptedRows} 列，彙整後 ${extracted.records.length} 筆。${formulaNote}`, 'success');
    } catch (error) {
      state.records = []; $('runDraw').disabled = true;
      setNotice(error.message || '讀取失敗，請檢查 Excel 清冊格式。', 'error');
    }
  }
  function setNotice(message, type) {
    const el = $('parseNotice'); el.textContent = message; el.className = 'notice' + (type ? ` ${type}` : '');
  }
  function readConfig() {
    return { minSerials: $('minSerials').value, maxPerSerial: $('maxPerSerial').value, totalCount: $('totalCount').value };
  }
  function executeDraw() {
    if (!state.records.length) { showToast('請先匯入清冊或載入範例資料。'); return; }
    try {
      const result = runSelection(state.records, readConfig());
      renderResult(result);
      showToast(result.status === '已達標' ? '抽選完成，檢核通過。' : `抽選完成，但尚缺 ${result.shortfall} 筆。`);
    } catch (error) { showToast(error.message); }
  }

  function decimalFormat(value) {
    const text = String(value);
    if (!text.includes('.')) return '0';
    const decimals = text.split('.')[1].replace(/0+$/, '').length;
    return decimals ? `0.${'#'.repeat(Math.min(decimals, 14))}` : '0';
  }
  function exportResult() {
    if (!state.result || !window.XLSX) return;
    const columns = ['流水號', '土地座落地段', '地號', '本筆面積', '契作面積（公頃）'];
    const rows = [columns, ...state.result.selected.map(r => [r.serial, r.section, r.landNo, Number(r.unitArea), Number(r.contractArea)])];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const headerStyle = { fill: { patternType: 'solid', fgColor: { rgb: '17365D' } }, font: { bold: true, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: { bottom: { style: 'thin', color: { rgb: '9FBAD0' } } } };
    columns.forEach((_, c) => { const cell = ws[XLSX.utils.encode_cell({ r: 0, c })]; cell.s = headerStyle; });
    for (let r = 1; r < rows.length; r += 1) {
      for (const c of [3, 4]) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell) { cell.t = 'n'; cell.v = Number(cell.v); cell.z = decimalFormat(cell.v); }
      }
    }
    ws['!cols'] = [
      { wch: 16 }, { wch: 24 }, { wch: 18 },
      { wch: Math.min(255, Math.max(12, ...rows.slice(1).map(r => formatArea(r[3]).length + 2))) },
      { wch: Math.min(255, Math.max(18, ...rows.slice(1).map(r => formatArea(r[4]).length + 2))) }
    ];
    ws['!rows'] = [{ hpt: 24 }];
    ws['!autofilter'] = { ref: `A1:E${rows.length}` };
    ws['!freeze'] = 'A2';
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '抽選結果');
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `抽選結果_${stamp}.xlsx`, { bookType: 'xlsx', compression: true });
    showToast('Excel 結果檔已產生。');
  }

  function loadSample() {
    const aggregated = aggregateRecords(makeSampleRows());
    state.records = aggregated.records; state.rawCount = aggregated.acceptedRows; state.fileName = 'RFS_虛構範例清冊'; state.sourceSheet = '範例資料'; state.usingSample = true; state.result = null;
    $('fileName').textContent = 'RFS_虛構範例清冊'; $('fileMeta').textContent = '內建虛構資料 · 不含真實個資'; $('resultPanel').hidden = true;
    renderSource(); setNotice(`範例載入完成：${aggregated.acceptedRows} 列來源資料，彙整後 ${aggregated.records.length} 筆土地（含 1 組重複地號示範）。`, 'success');
    $('tool').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function downloadSample() {
    if (!window.XLSX) { showToast('Excel 元件尚未載入。'); return; }
    const rows = makeSampleRows();
    const data = [['流水號', '土地座落地段', '地號', '本筆面積', '契作面積（公頃）'], ...rows.map(r => [r.serial, r.section, r.landNo, r.unitArea, r.contractArea])];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 16 }, { wch: 24 }, { wch: 18 }, { wch: 14 }, { wch: 20 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '土地清冊範例');
    XLSX.writeFile(wb, 'RFS_虛構土地清冊範例.xlsx');
    showToast('已產生虛構範例清冊。');
  }

  if (typeof document !== 'undefined') {
    $('fileInput').addEventListener('change', e => handleFile(e.target.files[0]));
    $('loadSample').addEventListener('click', loadSample);
    $('exampleRun').addEventListener('click', () => { loadSample(); executeDraw(); });
    $('downloadSample').addEventListener('click', downloadSample);
    $('runDraw').addEventListener('click', executeDraw);
    $('downloadResult').addEventListener('click', exportResult);
    const dropzone = $('dropzone');
    ['dragenter', 'dragover'].forEach(name => dropzone.addEventListener(name, e => { e.preventDefault(); dropzone.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach(name => dropzone.addEventListener(name, e => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
    dropzone.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
  }

  // Export core functions for automated local tests (no dependency on DOM).
  if (typeof module !== 'undefined' && module.exports) module.exports = { aggregateRecords, runSelection, makeSampleRows, parseNumber, keyOf };
})();
