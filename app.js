(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const state = {
    dgRows: [],
    productRows: [],
    productMap: new Map(),
    ships: [],
    selectedShips: new Set(),
    errors: [],
    mapping: loadJson('v7-mapping', {}),
    settings: loadJson('v7-settings', { normalize: true, removeVietgap: true, dedupe: true }),
    templateBuffer: null,
  };

  const REQUIRED_DG = ['code', 'tenkhach', 'diemgiao', 'ship'];
  const REQUIRED_PRODUCT = ['code', 'tensanpham'];
  const HEADER_ALIASES = {
    code: ['code', 'ma', 'masp', 'makhach'],
    tenkhach: ['tenkhach', 'tenkhachhang', 'khachhang', 'customer'],
    diemgiao: ['diemgiao', 'diachigiao', 'noigiao', 'deliverypoint'],
    ship: ['ship', 'machuyen', 'chuyenhang'],
    tensanpham: ['tensanpham', 'sanpham', 'product', 'productname'],
  };

  function loadJson(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; }
  }
  function saveJson(key, value) { localStorage.setItem(key, JSON.stringify(value)); }
  function normalizeHeader(value) {
    return String(value ?? '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toLowerCase().replace(/đ/g, 'd').replace(/[^a-z0-9]/g, '');
  }
  function cleanText(value) { return String(value ?? '').replace(/\s+/g, ' ').trim(); }
  function showToast(message) {
    const el = $('toast'); el.textContent = message; el.classList.remove('hidden');
    clearTimeout(showToast.timer); showToast.timer = setTimeout(() => el.classList.add('hidden'), 2600);
  }
  function showStatus(message, isError = false) {
    const el = $('status'); el.textContent = message; el.classList.remove('hidden', 'error');
    if (isError) el.classList.add('error');
  }
  function setProgress(percent, text) {
    $('progress').classList.remove('hidden');
    $('progressPercent').textContent = `${percent}%`;
    $('progressText').textContent = text;
    $('progressBar').style.width = `${percent}%`;
  }
  function sleep(ms = 0) { return new Promise(resolve => setTimeout(resolve, ms)); }

  function parseCsv(text) {
    const rows = []; let row = [], field = '', quoted = false;
    const input = String(text).replace(/^\uFEFF/, '');
    for (let i = 0; i < input.length; i++) {
      const ch = input[i];
      if (quoted) {
        if (ch === '"' && input[i + 1] === '"') { field += '"'; i++; }
        else if (ch === '"') quoted = false;
        else field += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === ',') { row.push(field); field = ''; }
      else if (ch === '\n') { row.push(field.replace(/\r$/, '')); rows.push(row); row = []; field = ''; }
      else field += ch;
    }
    if (field.length || row.length) { row.push(field.replace(/\r$/, '')); rows.push(row); }
    return rows;
  }

  function columnIndexFromRef(ref) {
    const letters = String(ref).match(/^[A-Z]+/i)?.[0]?.toUpperCase() || 'A';
    let value = 0; for (const ch of letters) value = value * 26 + ch.charCodeAt(0) - 64;
    return value - 1;
  }

  async function readSimpleXlsx(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const workbookXml = await zip.file('xl/workbook.xml').async('text');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');
    const workbookDoc = parseXml(workbookXml), relsDoc = parseXml(relsXml);
    const sheetNodes = localElements(workbookDoc, 'sheet');
    const selectedSheet = sheetNodes.find(node => node.getAttribute('state') !== 'hidden') || sheetNodes[0];
    if (!selectedSheet) throw new Error('File Excel không có worksheet.');
    const relationId = selectedSheet.getAttribute('r:id') || [...selectedSheet.attributes].find(a => a.localName === 'id')?.value;
    const relation = localElements(relsDoc, 'Relationship').find(node => node.getAttribute('Id') === relationId);
    if (!relation) throw new Error('Không xác định được worksheet đầu tiên.');
    let target = relation.getAttribute('Target').replace(/^\//, '');
    const sheetPath = target.startsWith('xl/') ? target : `xl/${target.replace(/^\.\.\//, '')}`;

    let sharedStrings = [];
    if (zip.file('xl/sharedStrings.xml')) {
      const sharedDoc = parseXml(await zip.file('xl/sharedStrings.xml').async('text'));
      sharedStrings = localElements(sharedDoc, 'si').map(si => localElements(si, 't').map(t => t.textContent).join(''));
    }
    const sheetDoc = parseXml(await zip.file(sheetPath).async('text'));
    const matrix = []; let maxCol = 0;
    for (const rowNode of localElements(firstLocal(sheetDoc, 'sheetData'), 'row')) {
      const rowNumber = Number(rowNode.getAttribute('r') || matrix.length + 1);
      const row = [];
      for (const cell of localElements(rowNode, 'c')) {
        const col = columnIndexFromRef(cell.getAttribute('r')); maxCol = Math.max(maxCol, col);
        const type = cell.getAttribute('t'); let value = '';
        if (type === 'inlineStr') value = localElements(cell, 't').map(t => t.textContent).join('');
        else {
          const v = firstLocal(cell, 'v')?.textContent ?? '';
          value = type === 's' ? (sharedStrings[Number(v)] ?? '') : v;
        }
        row[col] = value;
      }
      while (matrix.length < rowNumber - 1) matrix.push([]);
      matrix[rowNumber - 1] = row;
    }
    for (const row of matrix) for (let i = 0; i <= maxCol; i++) if (row[i] == null) row[i] = '';
    return matrix;
  }

  async function readTabularFile(file) {
    if (file.name.toLowerCase().endsWith('.csv')) return parseCsv(await file.text());
    return readSimpleXlsx(await file.arrayBuffer());
  }

  function identifyHeaderMap(row, required) {
    const normalized = row.map(normalizeHeader);
    const map = {};
    for (const canonical of required) {
      const aliases = HEADER_ALIASES[canonical];
      const index = normalized.findIndex(item => aliases.includes(item));
      if (index < 0) return null;
      map[canonical] = index;
    }
    return map;
  }

  function tableToObjects(matrix, required) {
    const maxScan = Math.min(matrix.length, 60);
    let headerIndex = -1, headerMap = null;
    for (let i = 0; i < maxScan; i++) {
      const map = identifyHeaderMap(matrix[i] || [], required);
      if (map) { headerIndex = i; headerMap = map; break; }
    }
    if (headerIndex < 0) throw new Error(`Không tìm thấy hàng tiêu đề bắt buộc: ${required.join(', ')}`);
    const rows = [];
    for (let i = headerIndex + 1; i < matrix.length; i++) {
      const row = matrix[i] || [];
      const obj = { __row: i + 1 };
      let hasData = false;
      for (const key of required) {
        obj[key] = cleanText(row[headerMap[key]]);
        if (obj[key]) hasData = true;
      }
      if (hasData) rows.push(obj);
    }
    return rows;
  }

  async function loadInputs() {
    const dgFile = $('dgFile').files[0];
    const productFiles = [...$('productFiles').files];
    if (!dgFile || !productFiles.length) return;
    try {
      setProgress(5, 'Đang đọc file DG');
      state.dgRows = tableToObjects(await readTabularFile(dgFile), REQUIRED_DG);
      state.productRows = [];
      for (let i = 0; i < productFiles.length; i++) {
        setProgress(10 + Math.round((i + 1) / productFiles.length * 30), `Đang đọc file sản phẩm ${i + 1}/${productFiles.length}`);
        const rows = tableToObjects(await readTabularFile(productFiles[i]), REQUIRED_PRODUCT);
        rows.forEach(row => { row.__file = productFiles[i].name; });
        state.productRows.push(...rows);
      }
      buildProductMap();
      validateData();
      buildShipList();
      updateSummary();
      setProgress(100, 'Đọc dữ liệu hoàn tất');
      setTimeout(() => $('progress').classList.add('hidden'), 700);
      showToast(`Đã đọc ${state.dgRows.length} dòng DG và ${state.productRows.length} dòng sản phẩm`);
    } catch (error) {
      showStatus(error.message, true);
      $('progress').classList.add('hidden');
    }
  }

  function buildProductMap() {
    state.productMap = new Map();
    const exactSeen = new Set();
    for (const row of state.productRows) {
      const code = cleanText(row.code);
      const product = cleanText(row.tensanpham);
      if (!code || !product) continue;
      const exactKey = `${code}\u0000${product.toLocaleLowerCase('vi')}`;
      if (state.settings.dedupe && exactSeen.has(exactKey)) continue;
      exactSeen.add(exactKey);
      if (!state.productMap.has(code)) state.productMap.set(code, []);
      state.productMap.get(code).push(product);
    }
  }

  function validateData() {
    const errors = [];
    for (const row of state.dgRows) {
      for (const field of REQUIRED_DG) {
        if (!row[field]) errors.push({ level: 'critical', type: `Thiếu ${field}`, detail: `Dòng DG ${row.__row}`, ship: row.ship, code: row.code });
      }
      if (row.code && !state.productMap.has(row.code)) {
        errors.push({ level: 'critical', type: 'Code không có trong file sản phẩm', detail: `Dòng DG ${row.__row}`, ship: row.ship, code: row.code });
      }
      if (state.settings.normalize && row.tenkhach && !findMappedName(row.tenkhach)) {
        errors.push({ level: 'warning', type: 'Tên khách chưa có trong bảng chuẩn hoá', detail: row.tenkhach, ship: row.ship, code: row.code });
      }
    }
    for (const row of state.productRows) {
      if (!row.code || !row.tensanpham) errors.push({ level: 'critical', type: 'Thiếu Code hoặc Tên sản phẩm', detail: `${row.__file || 'File sản phẩm'} dòng ${row.__row}`, code: row.code });
    }
    state.errors = errors;
    renderErrors();
  }

  function findMappedName(name) {
    const exact = state.mapping[cleanText(name).toLocaleLowerCase('vi')];
    return exact || '';
  }
  function outputCustomerName(name) {
    if (!state.settings.normalize) return cleanText(name);
    return findMappedName(name) || cleanText(name);
  }
  function cleanProductName(name) {
    let value = cleanText(name);
    if (state.settings.removeVietgap) value = value.replace(/viet\s*gap/gi, ' ');
    return value.replace(/\s*[-–—]\s*$/g, '').replace(/\s+/g, ' ').trim();
  }
  function productRank(name) {
    const value = name.toLocaleLowerCase('vi');
    if (/5\s*kg/.test(value)) return 1;
    if (/2\s*kg/.test(value)) return 2;
    if (/1\s*kg/.test(value)) return 3;
    if (/khay/.test(value)) return 4;
    if (/(^|\s)tf($|\s)/i.test(value)) return 5;
    if (/500\s*g/.test(value)) return 6;
    return 99;
  }

  function buildShipList() {
    const map = new Map();
    for (const row of state.dgRows) {
      if (!row.ship) continue;
      if (!map.has(row.ship)) map.set(row.ship, { rows: 0, customers: new Set(), products: 0 });
      const info = map.get(row.ship); info.rows++;
      info.customers.add(`${row.tenkhach}\u0000${row.diemgiao}`);
      info.products += (state.productMap.get(row.code) || []).length;
    }
    state.ships = [...map.entries()].map(([ship, info]) => ({ ship, rows: info.rows, customers: info.customers.size, products: info.products })).sort((a,b) => a.ship.localeCompare(b.ship, 'vi'));
    const validShips = new Set(state.ships.map(item => item.ship));
    state.selectedShips = new Set([...state.selectedShips].filter(ship => validShips.has(ship)));
    renderShipList();
  }

  function renderShipList() {
    const query = cleanText($('shipSearch').value).toLocaleLowerCase('vi');
    const items = state.ships.filter(item => item.ship.toLocaleLowerCase('vi').includes(query));
    const root = $('shipList'); root.innerHTML = '';
    if (!items.length) { root.className = 'ship-list empty'; root.textContent = state.ships.length ? 'Không tìm thấy Ship.' : 'Chưa đọc dữ liệu.'; return; }
    root.className = 'ship-list';
    for (const item of items) {
      const label = document.createElement('label'); label.className = 'ship-item';
      const checkbox = document.createElement('input'); checkbox.type = 'checkbox'; checkbox.checked = state.selectedShips.has(item.ship);
      checkbox.addEventListener('change', () => { checkbox.checked ? state.selectedShips.add(item.ship) : state.selectedShips.delete(item.ship); updateShipSummary(); updateSummary(); });
      const span = document.createElement('span');
      span.innerHTML = `<b>${escapeHtml(item.ship)}</b><small>${item.customers} khách · ${item.products} sản phẩm</small>`;
      label.append(checkbox, span); root.append(label);
    }
    updateShipSummary();
  }
  function updateShipSummary() { $('shipSummary').textContent = `${state.selectedShips.size} đã chọn`; }
  function escapeHtml(value) { return String(value).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }

  function buildGroups(shipFilter) {
    const groups = new Map();
    for (const row of state.dgRows) {
      if (shipFilter && !shipFilter.has(row.ship)) continue;
      const customer = outputCustomerName(row.tenkhach);
      const delivery = cleanText(row.diemgiao);
      const key = `${row.ship}\u0000${customer}\u0000${delivery}`;
      if (!groups.has(key)) groups.set(key, { ship: row.ship, customer, delivery, products: [], seen: new Set() });
      const group = groups.get(key);
      for (const rawProduct of state.productMap.get(row.code) || []) {
        const product = cleanProductName(rawProduct); if (!product) continue;
        const productKey = product.toLocaleLowerCase('vi');
        if (state.settings.dedupe && group.seen.has(productKey)) continue;
        group.seen.add(productKey); group.products.push(product);
      }
    }
    const output = [...groups.values()].filter(group => group.products.length);
    for (const group of output) group.products.sort((a,b) => productRank(a) - productRank(b) || a.localeCompare(b, 'vi'));
    return output;
  }

  function chooseBestSubset(groups, capacity) {
    const n = groups.length;
    const dp = Array(capacity + 1).fill(null); dp[0] = [];
    for (let i = 0; i < n; i++) {
      const size = groups[i].products.length;
      if (size > capacity) continue;
      for (let total = capacity; total >= size; total--) {
        if (dp[total - size] && !dp[total]) dp[total] = [...dp[total - size], i];
      }
    }
    for (let total = capacity; total >= 0; total--) if (dp[total]) return dp[total];
    return [];
  }

  function packGroups(groups) {
    const remaining = groups.map(group => ({ ...group, products: [...group.products] }));
    const forms = [];
    let capacity = 24;
    while (remaining.length) {
      const oversizedIndex = remaining.findIndex(group => group.products.length > capacity);
      if (oversizedIndex >= 0) {
        const group = remaining[oversizedIndex];
        const chunk = group.products.splice(0, capacity);
        forms.push([{ ...group, products: chunk }]);
        if (!group.products.length) remaining.splice(oversizedIndex, 1);
        capacity = 25;
        continue;
      }
      const indices = chooseBestSubset(remaining, capacity);
      if (!indices.length) {
        const group = remaining.shift(); forms.push([group]); capacity = 25; continue;
      }
      const selected = indices.map(index => remaining[index]);
      for (const index of [...indices].sort((a,b) => b-a)) remaining.splice(index, 1);
      forms.push(selected); capacity = 25;
    }
    return forms;
  }

  function flattenFormGroups(formGroups) {
    const rows = [];
    for (const group of formGroups) {
      const label = `${group.customer} - ${group.delivery}`;
      group.products.forEach((product, index) => rows.push([index === 0 ? label : '', product]));
    }
    return rows;
  }

  function updateSummary() {
    const filter = state.selectedShips.size ? state.selectedShips : null;
    const groups = buildGroups(filter);
    const products = groups.reduce((sum, group) => sum + group.products.length, 0);
    const forms = groups.length ? packGroups(groups).length : 0;
    $('customerCount').textContent = groups.length;
    $('productCount').textContent = products;
    $('formCount').textContent = forms;
  }

  function renderErrors() {
    const critical = state.errors.filter(error => error.level === 'critical').length;
    const warning = state.errors.filter(error => error.level === 'warning').length;
    $('criticalCount').textContent = critical; $('warningCount').textContent = warning; $('totalErrorCount').textContent = state.errors.length;
    const root = $('errorList'); root.innerHTML = '';
    if (!state.errors.length) { root.className = 'error-list empty'; root.textContent = state.dgRows.length ? 'Không phát hiện lỗi.' : 'Chưa có dữ liệu kiểm tra.'; return; }
    root.className = 'error-list';
    for (const error of state.errors.slice(0, 500)) {
      const item = document.createElement('div'); item.className = `error-item ${error.level}`;
      item.innerHTML = `<b>${escapeHtml(error.type)}</b><small>${escapeHtml(error.detail || '')}${error.ship ? ` · Ship ${escapeHtml(error.ship)}` : ''}${error.code ? ` · Code ${escapeHtml(error.code)}` : ''}</small>`;
      root.append(item);
    }
  }

  async function getTemplateBuffer() {
    const custom = $('templateFile').files[0];
    if (custom) return custom.arrayBuffer();
    if (state.templateBuffer) return state.templateBuffer.slice(0);
    try {
      const response = await fetch('BM-QC-26-template.xlsx');
      if (!response.ok) throw new Error('Không tải được file mẫu');
      state.templateBuffer = await response.arrayBuffer();
      return state.templateBuffer.slice(0);
    } catch {
      throw new Error('Trình duyệt không tải được mẫu kèm theo. Hãy chọn file BM-QC-26-template.xlsx trong mục “Mẫu Excel tuỳ chọn”, hoặc chạy app qua GitHub Pages/localhost.');
    }
  }

  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_DRAW = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  function localElements(root, name) { return [...root.getElementsByTagNameNS('*', name)]; }
  function firstLocal(root, name) { return root.getElementsByTagNameNS('*', name)[0] || null; }
  function parseXml(text) { return new DOMParser().parseFromString(text, 'application/xml'); }
  function serializeXml(doc) { return new XMLSerializer().serializeToString(doc); }
  function rowNumberFromRef(ref) { return Number(String(ref).match(/\d+$/)?.[0] || 0); }
  function shiftCellRef(ref, offset) { return String(ref).replace(/(\d+)$/, (_, n) => String(Number(n) + offset)); }
  function shiftRangeRef(ref, offset) { return String(ref).split(':').map(part => shiftCellRef(part, offset)).join(':'); }

  async function patchTemplate(templateBuffer, forms) {
    const zip = await JSZip.loadAsync(templateBuffer);
    const sheetPath = 'xl/worksheets/sheet2.xml';
    const drawingPath = 'xl/drawings/drawing1.xml';
    const workbookPath = 'xl/workbook.xml';
    const stringsPath = 'xl/sharedStrings.xml';
    const [sheetText, drawingText, workbookText, stringsText] = await Promise.all([
      zip.file(sheetPath).async('text'), zip.file(drawingPath).async('text'), zip.file(workbookPath).async('text'), zip.file(stringsPath).async('text')
    ]);
    const sheetDoc = parseXml(sheetText), drawingDoc = parseXml(drawingText), workbookDoc = parseXml(workbookText), stringsDoc = parseXml(stringsText);
    const sheetData = firstLocal(sheetDoc, 'sheetData');
    const mergeCells = firstLocal(sheetDoc, 'mergeCells');
    const dimension = firstLocal(sheetDoc, 'dimension');
    const originalRows = localElements(sheetData, 'row').map(row => row.cloneNode(true));
    const sourceSecondRows = originalRows.filter(row => Number(row.getAttribute('r')) >= 37 && Number(row.getAttribute('r')) <= 72);
    const sourceSecondMerges = localElements(mergeCells, 'mergeCell').filter(node => {
      const start = rowNumberFromRef(node.getAttribute('ref').split(':')[0]);
      return start >= 37 && start <= 72;
    }).map(node => node.getAttribute('ref'));

    // Remove trailing blank row 73 and any previously generated rows.
    for (const row of localElements(sheetData, 'row')) if (Number(row.getAttribute('r')) >= 73) row.remove();
    for (const node of [...localElements(mergeCells, 'mergeCell')]) if (rowNumberFromRef(node.getAttribute('ref').split(':')[0]) >= 73) node.remove();

    const actualFormCount = Math.max(2, forms.length);
    for (let formIndex = 2; formIndex < actualFormCount; formIndex++) {
      const offset = 36 * (formIndex - 1);
      for (const sourceRow of sourceSecondRows) {
        const clone = sourceRow.cloneNode(true);
        clone.setAttribute('r', Number(sourceRow.getAttribute('r')) + offset);
        for (const cell of localElements(clone, 'c')) cell.setAttribute('r', shiftCellRef(cell.getAttribute('r'), offset));
        sheetData.appendChild(clone);
      }
      for (const ref of sourceSecondMerges) {
        const node = sheetDoc.createElementNS(NS_MAIN, 'mergeCell'); node.setAttribute('ref', shiftRangeRef(ref, offset)); mergeCells.appendChild(node);
      }
    }
    mergeCells.setAttribute('count', String(localElements(mergeCells, 'mergeCell').length));
    dimension.setAttribute('ref', `A2:X${actualFormCount * 36}`);

    // Shared strings index.
    const sst = stringsDoc.documentElement;
    const stringMap = new Map();
    localElements(sst, 'si').forEach((si, index) => {
      const text = localElements(si, 't').map(t => t.textContent).join('');
      if (!stringMap.has(text)) stringMap.set(text, index);
    });
    const getStringIndex = (text) => {
      const value = String(text ?? '');
      if (stringMap.has(value)) return stringMap.get(value);
      const si = stringsDoc.createElementNS(NS_MAIN, 'si');
      const t = stringsDoc.createElementNS(NS_MAIN, 't'); t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); t.textContent = value;
      si.appendChild(t); sst.appendChild(si);
      const index = localElements(sst, 'si').length - 1; stringMap.set(value, index); return index;
    };

    const rowMap = new Map(localElements(sheetData, 'row').map(row => [Number(row.getAttribute('r')), row]));
    function findCell(rowNumber, column) {
      const row = rowMap.get(rowNumber); if (!row) throw new Error(`Thiếu dòng ${rowNumber} trong biểu mẫu`);
      const ref = `${column}${rowNumber}`;
      let cell = localElements(row, 'c').find(item => item.getAttribute('r') === ref);
      if (!cell) {
        cell = sheetDoc.createElementNS(NS_MAIN, 'c'); cell.setAttribute('r', ref);
        row.appendChild(cell);
      }
      return cell;
    }
    function setCellText(rowNumber, column, value) {
      const cell = findCell(rowNumber, column);
      while (cell.firstChild) cell.removeChild(cell.firstChild);
      if (!value) { cell.removeAttribute('t'); return; }
      cell.setAttribute('t', 's');
      const v = sheetDoc.createElementNS(NS_MAIN, 'v'); v.textContent = String(getStringIndex(value)); cell.appendChild(v);
    }

    // Clear and fill each form. Form 1 has 24 data lines, later forms have 25.
    for (let formIndex = 0; formIndex < actualFormCount; formIndex++) {
      const start = formIndex === 0 ? 9 : 43 + 36 * (formIndex - 1);
      const capacity = formIndex === 0 ? 24 : 25;
      for (let i = 0; i < capacity; i++) { setCellText(start + i, 'J', ''); setCellText(start + i, 'K', ''); }
      const rows = forms[formIndex] ? flattenFormGroups(forms[formIndex]) : [];
      rows.slice(0, capacity).forEach((pair, i) => { setCellText(start + i, 'J', pair[0]); setCellText(start + i, 'K', pair[1]); });
    }
    sst.setAttribute('count', String(localElements(sst, 'si').length));
    sst.setAttribute('uniqueCount', String(localElements(sst, 'si').length));

    // Page breaks after every form except the last.
    let rowBreaks = firstLocal(sheetDoc, 'rowBreaks');
    if (!rowBreaks) {
      rowBreaks = sheetDoc.createElementNS(NS_MAIN, 'rowBreaks');
      const drawing = firstLocal(sheetDoc, 'drawing'); sheetDoc.documentElement.insertBefore(rowBreaks, drawing || null);
    }
    while (rowBreaks.firstChild) rowBreaks.removeChild(rowBreaks.firstChild);
    for (let i = 1; i < actualFormCount; i++) {
      const brk = sheetDoc.createElementNS(NS_MAIN, 'brk'); brk.setAttribute('id', String(i * 36)); brk.setAttribute('max', '23'); brk.setAttribute('man', '1'); rowBreaks.appendChild(brk);
    }
    rowBreaks.setAttribute('count', String(Math.max(0, actualFormCount - 1)));
    rowBreaks.setAttribute('manualBreakCount', String(Math.max(0, actualFormCount - 1)));

    // Rebuild exactly two signature textboxes per form from the first two original anchors.
    const drawingRoot = drawingDoc.documentElement;
    const baseAnchors = localElements(drawingRoot, 'oneCellAnchor').slice(0, 2).map(anchor => anchor.cloneNode(true));
    while (drawingRoot.firstChild) drawingRoot.removeChild(drawingRoot.firstChild);
    let shapeId = 2;
    for (let formIndex = 0; formIndex < actualFormCount; formIndex++) {
      const offset = 36 * formIndex;
      for (const base of baseAnchors) {
        const clone = base.cloneNode(true);
        const row = firstLocal(firstLocal(clone, 'from'), 'row'); row.textContent = String(Number(row.textContent) + offset);
        const cNvPr = firstLocal(clone, 'cNvPr'); cNvPr.setAttribute('id', String(shapeId)); cNvPr.setAttribute('name', `TextBox ${shapeId}`); shapeId++;
        const extLst = firstLocal(cNvPr, 'extLst'); if (extLst) extLst.remove();
        drawingRoot.appendChild(clone);
      }
    }

    // Update print area for Xe to.
    for (const definedName of localElements(workbookDoc, 'definedName')) {
      if (definedName.getAttribute('name') === '_xlnm.Print_Area' && definedName.getAttribute('localSheetId') === '1') {
        definedName.textContent = `'Xe to'!$A$1:$X$${actualFormCount * 36}`;
      }
    }

    zip.file(sheetPath, serializeXml(sheetDoc));
    zip.file(drawingPath, serializeXml(drawingDoc));
    zip.file(workbookPath, serializeXml(workbookDoc));
    zip.file(stringsPath, serializeXml(stringsDoc));
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  function safeFileName(value) { return String(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z0-9_-]+/g, '_').replace(/^_+|_+$/g, ''); }
  function dateStamp() { const d = new Date(); return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`; }
  function downloadBlob(blob, name) { const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1500); }

  async function exportReport() {
    if (!state.dgRows.length || !state.productRows.length) return showStatus('Hãy tải file DG và ít nhất một file sản phẩm.', true);
    if (!state.selectedShips.size) return showStatus('Hãy chọn ít nhất một Ship.', true);
    const critical = state.errors.filter(error => error.level === 'critical' && (!error.ship || state.selectedShips.has(error.ship)));
    if (critical.length) return showStatus(`Còn ${critical.length} lỗi cần sửa. Mở menu “Kiểm tra lỗi” để xem.`, true);
    try {
      setProgress(5, 'Đang tải biểu mẫu gốc'); await sleep(20);
      const template = await getTemplateBuffer();
      const mode = $('multiShipMode').value;
      if (mode === 'separate' && state.selectedShips.size > 1) {
        const outerZip = new JSZip(); let index = 0;
        for (const ship of state.selectedShips) {
          index++; setProgress(10 + Math.round(index/state.selectedShips.size*80), `Đang tạo báo cáo Ship ${ship}`); await sleep(10);
          const groups = buildGroups(new Set([ship])); const forms = packGroups(groups);
          const blob = await patchTemplate(template.slice(0), forms);
          outerZip.file(`Bao_cao_${safeFileName(ship)}_${dateStamp()}.xlsx`, blob);
        }
        const zipBlob = await outerZip.generateAsync({ type:'blob', compression:'DEFLATE' });
        downloadBlob(zipBlob, `Bao_cao_xuat_hang_${dateStamp()}.zip`);
      } else {
        const groups = buildGroups(state.selectedShips); const forms = packGroups(groups);
        setProgress(35, `Đang chia ${forms.length} form`); await sleep(20);
        const blob = await patchTemplate(template, forms);
        const shipName = state.selectedShips.size === 1 ? [...state.selectedShips][0] : 'nhieu_ship';
        downloadBlob(blob, `Bao_cao_${safeFileName(shipName)}_${dateStamp()}.xlsx`);
      }
      setProgress(100, 'Hoàn thành'); showStatus('Đã tạo file Excel từ đúng biểu mẫu gốc.');
      setTimeout(() => $('progress').classList.add('hidden'), 1000);
    } catch (error) {
      console.error(error); $('progress').classList.add('hidden'); showStatus(error.message || 'Không thể tạo báo cáo.', true);
    }
  }

  function saveMapping() {
    const map = {};
    for (const line of $('mappingText').value.split(/\r?\n/)) {
      const [left, ...rightParts] = line.split('=>');
      const right = rightParts.join('=>');
      if (cleanText(left) && cleanText(right)) map[cleanText(left).toLocaleLowerCase('vi')] = cleanText(right);
    }
    state.mapping = map; saveJson('v7-mapping', map); validateData(); updateSummary();
    $('mappingStatus').textContent = `Đã lưu ${Object.keys(map).length} tên chuẩn hoá.`; $('mappingStatus').classList.remove('hidden');
  }
  function renderMapping() { $('mappingText').value = Object.entries(state.mapping).map(([key, value]) => `${key} => ${value}`).join('\n'); }
  function applySettingsToUi() { $('normalizeToggle').checked = state.settings.normalize; $('vietgapToggle').checked = state.settings.removeVietgap; $('dedupeToggle').checked = state.settings.dedupe; }
  function saveSettings() {
    state.settings = { normalize: $('normalizeToggle').checked, removeVietgap: $('vietgapToggle').checked, dedupe: $('dedupeToggle').checked };
    saveJson('v7-settings', state.settings); buildProductMap(); validateData(); updateSummary(); showToast('Đã lưu cài đặt');
  }
  function downloadErrors() {
    const rows = [['Mức độ','Loại lỗi','Chi tiết','Ship','Code'], ...state.errors.map(e => [e.level,e.type,e.detail||'',e.ship||'',e.code||''])];
    const csv = rows.map(row => row.map(value => `"${String(value).replace(/"/g,'""')}"`).join(',')).join('\r\n');
    downloadBlob(new Blob(['\ufeff'+csv], {type:'text/csv;charset=utf-8'}), `danh_sach_loi_${dateStamp()}.csv`);
  }

  function showScreen(name) {
    document.querySelectorAll('.screen').forEach(el => el.classList.toggle('active', el.id === `screen-${name}`));
    document.querySelectorAll('.nav-button').forEach(el => el.classList.toggle('active', el.dataset.screen === name));
    $('drawer').classList.add('hidden'); $('backdrop').classList.add('hidden');
  }
  function resetSession() {
    state.dgRows=[]; state.productRows=[]; state.productMap=new Map(); state.ships=[]; state.selectedShips.clear(); state.errors=[];
    $('dgFile').value=''; $('productFiles').value=''; $('dgFileName').textContent='Chọn một file'; $('productFileName').textContent='Có thể chọn nhiều file';
    buildShipList(); renderErrors(); updateSummary(); $('status').classList.add('hidden');
  }

  $('menuButton').addEventListener('click', () => { $('drawer').classList.remove('hidden'); $('backdrop').classList.remove('hidden'); });
  $('closeMenu').addEventListener('click', () => { $('drawer').classList.add('hidden'); $('backdrop').classList.add('hidden'); });
  $('backdrop').addEventListener('click', () => $('closeMenu').click());
  document.querySelectorAll('.nav-button').forEach(button => button.addEventListener('click', () => showScreen(button.dataset.screen)));
  $('dgFile').addEventListener('change', () => { $('dgFileName').textContent = $('dgFile').files[0]?.name || 'Chọn một file'; loadInputs(); });
  $('productFiles').addEventListener('change', () => { const files=[...$('productFiles').files]; $('productFileName').textContent = files.length ? `${files.length} file: ${files.map(f=>f.name).join(', ')}` : 'Có thể chọn nhiều file'; loadInputs(); });
  $('shipToggle').addEventListener('click', () => $('shipPanel').classList.toggle('hidden'));
  $('shipSearch').addEventListener('input', renderShipList);
  $('selectAllButton').addEventListener('click', () => { state.ships.forEach(item => state.selectedShips.add(item.ship)); renderShipList(); updateSummary(); });
  $('clearShipButton').addEventListener('click', () => { state.selectedShips.clear(); renderShipList(); updateSummary(); });
  $('exportButton').addEventListener('click', exportReport);
  $('resetButton').addEventListener('click', resetSession);
  $('saveMapping').addEventListener('click', saveMapping);
  $('clearMapping').addEventListener('click', () => { state.mapping={}; saveJson('v7-mapping',{}); renderMapping(); validateData(); updateSummary(); });
  $('saveSettings').addEventListener('click', saveSettings);
  $('downloadErrors').addEventListener('click', downloadErrors);

  renderMapping(); applySettingsToUi(); renderErrors(); buildShipList(); updateSummary();
})();
