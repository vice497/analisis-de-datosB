(function(){
  let rawRows = [];      // array of objects, raw from sheet
  let headers = [];
  let movimientos = [];  // {date: Date, monto: number} monto>0 ingreso, monto<0 gasto — ya filtrado
  let allMovimientos = []; // todos los movimientos, sin aplicar filtros
  let currentTab = 'day';
  let skippedCount = 0;

  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const fileNameEl = document.getElementById('fileName');
  const errorNote = document.getElementById('errorNote');
  const mapCard = document.getElementById('mapCard');
  const previewTable = document.getElementById('previewTable');
  const detectionNote = document.getElementById('detectionNote');
  const colFecha = document.getElementById('colFecha');
  const colMontoUnico = document.getElementById('colMontoUnico');
  const colIngreso = document.getElementById('colIngreso');
  const colGasto = document.getElementById('colGasto');
  const colMontoTipo = document.getElementById('colMontoTipo');
  const colTipoOperacion = document.getElementById('colTipoOperacion');
  const tipoChipsEl = document.getElementById('tipoChips');
  const singleModeFields = document.getElementById('singleModeFields');
  const splitModeFields = document.getElementById('splitModeFields');
  const typeModeFields = document.getElementById('typeModeFields');
  const processBtn = document.getElementById('processBtn');
  const skippedNote = document.getElementById('skippedNote');
  const resultsSection = document.getElementById('resultsSection');
  const ledgerEl = document.getElementById('ledger');
  const personTabBtn = document.getElementById('personTabBtn');
  const personListEl = document.getElementById('personList');
  const filterDesde = document.getElementById('filterDesde');
  const filterHasta = document.getElementById('filterHasta');
  const filterBusqueda = document.getElementById('filterBusqueda');
  const filterBusquedaWrap = document.getElementById('filterBusquedaWrap');
  const clearFiltersBtn = document.getElementById('clearFiltersBtn');
  const filterEmptyNote = document.getElementById('filterEmptyNote');

  let montoMode = 'single';
  let tipoClassMap = {}; // { "TE PAGO": "ingreso", "PAGASTE": "gasto" }
  let nameConfig = { mode: 'none' }; // 'dual' (origen/destino), 'single' (una columna de nombre), o 'none'

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault();
    dropzone.classList.remove('drag');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', e => {
    if (e.target.files.length) handleFile(e.target.files[0]);
  });

  const resetBtn = document.getElementById('resetBtn');
  resetBtn.addEventListener('click', () => location.reload());

  document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.mode-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      montoMode = btn.dataset.mode;
      singleModeFields.style.display = montoMode === 'single' ? 'grid' : 'none';
      splitModeFields.style.display = montoMode === 'split' ? 'grid' : 'none';
      typeModeFields.style.display = montoMode === 'type' ? 'block' : 'none';
      if (montoMode === 'type' && colTipoOperacion.value) buildTipoChips(colTipoOperacion.value);
    });
  });

  colTipoOperacion.addEventListener('change', () => buildTipoChips(colTipoOperacion.value));

  function showError(msg){
    errorNote.textContent = msg;
    errorNote.style.display = 'block';
  }
  function clearError(){ errorNote.style.display = 'none'; }

  function handleFile(file){
    clearError();
    const validExt = /\.(xlsx|xls|csv)$/i.test(file.name);
    if (!validExt){
      showError('Formato no soportado. Usa un archivo .xlsx, .xls o .csv.');
      return;
    }
    fileNameEl.textContent = file.name;
    const reader = new FileReader();
    reader.onload = function(e){
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[firstSheetName];
        const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
        if (!aoa.length){
          showError('El archivo no tiene filas de datos en la primera hoja.');
          return;
        }
        const headerRowIdx = findHeaderRowIndex(aoa);
        const rawHeaders = aoa[headerRowIdx].map(h => String(h).trim());
        const dataAoa = aoa.slice(headerRowIdx + 1);

        rawRows = dataAoa.map(row => {
          const obj = {};
          rawHeaders.forEach((h, idx) => {
            if (h === '') return; // ignora columnas sin nombre (celdas vacías de relleno)
            obj[h] = row[idx] !== undefined ? row[idx] : '';
          });
          return obj;
        }).filter(obj => Object.values(obj).some(v => v !== ''));

        headers = rawHeaders.filter(h => h !== '');

        if (!rawRows.length){
          showError('No se encontraron filas de datos debajo del encabezado detectado.');
          return;
        }
        populateMapping();
      } catch(err){
        showError('No se pudo leer el archivo. Verifica que sea un Excel válido.');
      }
    };
    reader.onerror = function(){ showError('Ocurrió un error al leer el archivo.'); };
    reader.readAsArrayBuffer(file);
  }

  function findHeaderRowIndex(aoa){
    // busca la primera fila con varias celdas no vacías, cuya fila siguiente también tenga datos alineados
    // (así se salta títulos sueltos o filas en blanco antes del encabezado real)
    const limit = Math.min(aoa.length, 20);
    for (let i = 0; i < limit; i++){
      const row = aoa[i] || [];
      const nonEmpty = row.filter(c => c !== '' && c !== null && c !== undefined).length;
      if (nonEmpty >= 2){
        const next = aoa[i+1] || [];
        const nextNonEmpty = next.filter(c => c !== '' && c !== null && c !== undefined).length;
        if (nextNonEmpty >= Math.max(2, nonEmpty - 1)) return i;
      }
    }
    return 0;
  }

  function populateMapping(){
    [colFecha, colMontoUnico, colIngreso, colGasto, colMontoTipo, colTipoOperacion].forEach(sel => sel.innerHTML = '');
    headers.forEach(h => {
      [colFecha, colMontoUnico, colIngreso, colGasto, colMontoTipo, colTipoOperacion].forEach(sel => {
        const opt = document.createElement('option');
        opt.value = h; opt.textContent = h;
        sel.appendChild(opt.cloneNode(true));
      });
    });

    const detection = autoDetectColumns();
    colFecha.value = detection.fecha || headers[0];

    if (detection.mode === 'type'){
      document.querySelector('.mode-btn[data-mode="type"]').click();
      colMontoTipo.value = detection.monto;
      colTipoOperacion.value = detection.tipo;
      buildTipoChips(detection.tipo, detection.tipoClassMap);
      nameConfig = detectNameColumns(detection.fecha, [detection.monto, detection.tipo]);
    } else if (detection.mode === 'split'){
      document.querySelector('.mode-btn[data-mode="split"]').click();
      colIngreso.value = detection.ingreso;
      colGasto.value = detection.gasto;
      nameConfig = detectNameColumns(detection.fecha, [detection.ingreso, detection.gasto]);
    } else {
      document.querySelector('.mode-btn[data-mode="single"]').click();
      colMontoUnico.value = detection.montoUnico || headers[headers.length-1];
      nameConfig = detectNameColumns(detection.fecha, [colMontoUnico.value]);
    }

    personTabBtn.style.display = nameConfig.mode === 'none' ? 'none' : '';

    // preview table: header + up to 5 rows
    let thead = '<thead><tr>' + headers.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr></thead>';
    let tbody = '<tbody>' + rawRows.slice(0,5).map(row =>
      '<tr>' + headers.map(h => `<td>${escapeHtml(String(row[h]))}</td>`).join('') + '</tr>'
    ).join('') + '</tbody>';
    previewTable.innerHTML = thead + tbody;

    detectionNote.style.display = 'block';
    detectionNote.textContent = detection.confident
      ? 'Columnas detectadas automáticamente. Revisa la vista previa y, si algo no coincide, ajústalo abajo.'
      : 'No estoy seguro de haber detectado bien las columnas — revisa y corrige antes de procesar.';

    mapCard.style.display = 'block';

    // auto-run with the detected columns; the user can adjust and click "Reprocesar" si algo no coincide
    processBtn.click();
  }

  function normalizeText(s){
    return String(s).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  const TIPO_INGRESO_KEYS = [
    'te pago','te presto','recib','abono','deposit','cobr','ingreso','te transfir','recarga recibida',
    'te yape','yapearon','te plin','plinearon','te enviaron','te transfirieron'
  ];
  const TIPO_GASTO_KEYS = [
    'pagaste','pago a','enviaste','envio','retiro','compra','cargo','gasto','pago de servicio','transferiste',
    'yapeaste','plineaste','pago de recarga','recarga a'
  ];

  function matchTipoKeyword(val){
    const t = normalizeText(val);
    if (TIPO_INGRESO_KEYS.some(k => t.includes(k))) return 'ingreso';
    if (TIPO_GASTO_KEYS.some(k => t.includes(k))) return 'gasto';
    return null;
  }

  function classifyTipoValue(val){
    return matchTipoKeyword(val) || 'gasto'; // sin coincidencia clara, por defecto gasto (editable con un clic)
  }

  function tipoConfidence(header){
    const vals = sampleValues(header);
    if (!vals.length) return 0;
    const hits = vals.filter(v => matchTipoKeyword(v) !== null).length;
    return hits / vals.length;
  }

  function buildTipoChips(header, presetMap){
    const freq = {};
    rawRows.forEach(row => {
      const v = row[header];
      if (v === '' || v === null || v === undefined) return;
      const key = String(v).trim();
      freq[key] = (freq[key] || 0) + 1;
    });
    // mostrar los valores más frecuentes primero, para que los tipos reales
    // (aunque haya cientos de textos distintos) aparezcan arriba
    const topValues = Object.keys(freq).sort((a,b) => freq[b] - freq[a]).slice(0, 20);

    tipoClassMap = {};
    topValues.forEach(val => {
      tipoClassMap[val] = (presetMap && presetMap[val]) || classifyTipoValue(val);
    });

    renderTipoChips();
  }

  function renderTipoChips(){
    tipoChipsEl.innerHTML = Object.keys(tipoClassMap).map(val => {
      const cls = tipoClassMap[val];
      return `
        <div class="tipo-chip" data-val="${escapeHtml(val)}">
          <span class="tipo-name">${escapeHtml(val)}</span>
          <button type="button" class="tipo-flag ${cls}" data-val="${escapeHtml(val)}">${cls === 'ingreso' ? 'Ingreso' : 'Gasto'}</button>
        </div>`;
    }).join('');

    tipoChipsEl.querySelectorAll('.tipo-flag').forEach(btn => {
      btn.addEventListener('click', () => {
        const val = btn.dataset.val;
        tipoClassMap[val] = tipoClassMap[val] === 'ingreso' ? 'gasto' : 'ingreso';
        renderTipoChips();
      });
    });
  }

  // --- Auto-detección de columnas por contenido, no solo por el nombre ---

  function sampleValues(header){
    const vals = [];
    for (let i = 0; i < rawRows.length && vals.length < 40; i++){
      const v = rawRows[i][header];
      if (v !== '' && v !== null && v !== undefined) vals.push(v);
    }
    return vals;
  }

  function dateScore(header){
    const vals = sampleValues(header);
    if (!vals.length) return 0;
    const hits = vals.filter(v => parseDateVal(v) !== null).length;
    return hits / vals.length;
  }

  function numberScore(header){
    const vals = sampleValues(header);
    if (!vals.length) return 0;
    const hits = vals.filter(v => !isNaN(parseAmount(v))).length;
    return hits / vals.length;
  }

  function hasNegatives(header){
    return sampleValues(header).some(v => parseAmount(v) < 0);
  }

  function matchesKeywords(header, keywords){
    const h = header.toLowerCase();
    return keywords.some(k => h.includes(k));
  }

  function uniqueCount(header){
    const seen = new Set();
    sampleValues(header).forEach(v => seen.add(normalizeText(v)));
    return seen.size;
  }

  function autoDetectColumns(){
    // 1. Fecha: la columna con mayor proporción de valores interpretables como fecha
    let fecha = null, bestDateScore = 0;
    headers.forEach(h => {
      const s = dateScore(h);
      if (s > bestDateScore){ bestDateScore = s; fecha = h; }
    });

    const candidates = headers.filter(h => h !== fecha);
    const excludeAsBalance = h => matchesKeywords(h, ['saldo','balance','total acumulado']);

    // 2. Buscar columnas explícitas de ingreso y gasto por nombre + que contengan números
    const ingresoKeys = ['ingreso','abono','credito','crédito','deposito','depósito','haber'];
    const gastoKeys = ['gasto','cargo','debito','débito','retiro','egreso','debe'];
    const ingresoCol = candidates.find(h => matchesKeywords(h, ingresoKeys) && numberScore(h) > 0.3);
    const gastoCol = candidates.find(h => matchesKeywords(h, gastoKeys) && numberScore(h) > 0.3);

    if (ingresoCol && gastoCol && ingresoCol !== gastoCol){
      return { fecha, mode: 'split', ingreso: ingresoCol, gasto: gastoCol, confident: bestDateScore > 0.6 };
    }

    // 3. Buscar un patrón "monto sin signo + columna de tipo de operación" (típico de Yape/Plin/apps móviles).
    //    No exige pocos valores únicos: la columna puede traer texto libre (ej. "Pagaste a Juan Pérez"),
    //    basta con que la mayoría del texto contenga palabras como "pagaste" o "te pago".
    const montoCol = candidates.find(h => matchesKeywords(h, ['monto','importe','valor','amount']) && numberScore(h) > 0.5);
    const tipoCol = candidates.find(h => h !== montoCol && tipoConfidence(h) > 0.6);
    if (montoCol && tipoCol && !hasNegatives(montoCol)){
      return { fecha, mode: 'type', monto: montoCol, tipo: tipoCol, confident: bestDateScore > 0.6 };
    }

    // 4. Si no hay dos columnas separadas ni patrón de tipo, buscar una sola columna de monto con signo:
    //    preferir nombre típico, si no, la columna numérica (que no sea saldo) con más valores negativos
    const numericCandidates = candidates.filter(h => !excludeAsBalance(h) && numberScore(h) > 0.5);
    let montoUnico = numericCandidates.find(h => matchesKeywords(h, ['monto','importe','valor','amount','movimiento']));
    if (!montoUnico){
      montoUnico = numericCandidates.find(h => hasNegatives(h)) ||
                   numericCandidates.sort((a,b) => numberScore(b) - numberScore(a))[0];
    }

    return {
      fecha,
      mode: 'single',
      montoUnico,
      confident: bestDateScore > 0.6 && !!montoUnico
    };
  }

  // --- Detección de columnas de nombre/contacto, para poder agrupar "por persona" ---
  function detectNameColumns(fechaCol, usedCols){
    const candidates = headers.filter(h => h !== fechaCol && !usedCols.includes(h));

    const origenCol = candidates.find(h => matchesKeywords(h, ['origen','remitente','pagador']));
    const destinoCol = candidates.find(h => matchesKeywords(h, ['destino','destinatario','beneficiario']));
    if (origenCol && destinoCol && origenCol !== destinoCol){
      return { mode: 'dual', origen: origenCol, destino: destinoCol };
    }

    const singleKeys = ['nombre','contacto','beneficiario','destinatario','cliente','proveedor','remitente','quien','descripcion','concepto'];
    const nameCol = candidates.find(h =>
      matchesKeywords(h, singleKeys) && numberScore(h) < 0.3 && dateScore(h) < 0.3
    );
    if (nameCol) return { mode: 'single', columna: nameCol };

    return { mode: 'none' };
  }

  function getQuien(row, monto){
    if (nameConfig.mode === 'dual'){
      const val = monto < 0 ? row[nameConfig.destino] : row[nameConfig.origen];
      return String(val ?? '').trim();
    }
    if (nameConfig.mode === 'single'){
      return String(row[nameConfig.columna] ?? '').trim();
    }
    return '';
  }

  function escapeHtml(str){
    return str.replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  function parseAmount(val){
    if (typeof val === 'number') return val;
    if (val === null || val === undefined) return NaN;
    let s = String(val).trim();
    if (s === '') return NaN;
    // quitar símbolos de moneda pegados al número (S/., S/, PEN, $, USD) antes de interpretar
    // separadores decimales, para no confundir el punto del símbolo con un decimal (ej. "S/.30.00")
    s = s.replace(/^(S\/\.?|PEN|USD|US\$|\$)\s*/i, '').trim();
    let negative = false;
    if (/^\(.*\)$/.test(s)){ negative = true; s = s.slice(1,-1); }
    s = s.replace(/[^0-9.,\-]/g, '');
    if (s.includes('-')) { negative = negative || s.trim().startsWith('-'); s = s.replace(/-/g,''); }
    const lastComma = s.lastIndexOf(',');
    const lastDot = s.lastIndexOf('.');
    if (lastComma > -1 && lastDot > -1){
      if (lastComma > lastDot){ s = s.replace(/\./g,'').replace(',', '.'); }
      else { s = s.replace(/,/g,''); }
    } else if (lastComma > -1){
      const decimals = s.length - lastComma - 1;
      if (decimals === 2) s = s.replace(',', '.');
      else s = s.replace(/,/g,'');
    }
    const num = parseFloat(s);
    if (isNaN(num)) return NaN;
    return negative ? -Math.abs(num) : num;
  }

  function parseDateVal(val){
    if (val instanceof Date && !isNaN(val)) return val;
    if (typeof val === 'number'){
      const d = XLSX.SSF ? XLSX.SSF.parse_date_code(val) : null;
      if (d) return new Date(d.y, d.m-1, d.d);
    }
    if (typeof val === 'string'){
      const s = val.trim();
      // fecha con u sin hora pegada: "12/08/2026" o "12/08/2026 21:53:58"
      let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
      if (m){
        let [_, a, b, y, hh, mm, ss] = m;
        if (y.length === 2) y = '20'+y;
        let day = parseInt(a,10), month = parseInt(b,10);
        if (month > 12) { [day, month] = [month, day]; }
        const d = new Date(
          parseInt(y,10), month-1, day,
          hh ? parseInt(hh,10) : 0,
          mm ? parseInt(mm,10) : 0,
          ss ? parseInt(ss,10) : 0
        );
        if (!isNaN(d)) return d;
      }
      const d2 = new Date(s);
      if (!isNaN(d2)) return d2;
    }
    return null;
  }

  processBtn.addEventListener('click', () => {
    skippedCount = 0;
    allMovimientos = [];
    const fCol = colFecha.value;
    rawRows.forEach(row => {
      const date = parseDateVal(row[fCol]);
      if (!date){ skippedCount++; return; }

      if (montoMode === 'single'){
        const monto = parseAmount(row[colMontoUnico.value]);
        if (isNaN(monto)){ skippedCount++; return; }
        allMovimientos.push({date, monto, quien: getQuien(row, monto)});
      } else if (montoMode === 'type'){
        const monto = parseAmount(row[colMontoTipo.value]);
        const tipoVal = String(row[colTipoOperacion.value] ?? '').trim();
        if (isNaN(monto) || monto === 0 || !tipoVal){ skippedCount++; return; }
        const clasificacion = Object.prototype.hasOwnProperty.call(tipoClassMap, tipoVal)
          ? tipoClassMap[tipoVal]
          : classifyTipoValue(tipoVal);
        const signo = clasificacion === 'ingreso' ? 1 : -1;
        const montoFinal = signo * Math.abs(monto);
        allMovimientos.push({date, monto: montoFinal, quien: getQuien(row, montoFinal)});
      } else {
        const ingresoRaw = row[colIngreso.value];
        const gastoRaw = row[colGasto.value];
        const ingreso = parseAmount(ingresoRaw);
        const gasto = parseAmount(gastoRaw);
        const hasIngreso = !isNaN(ingreso) && ingreso !== 0;
        const hasGasto = !isNaN(gasto) && gasto !== 0;
        if (!hasIngreso && !hasGasto){ skippedCount++; return; }
        if (hasIngreso) allMovimientos.push({date, monto: Math.abs(ingreso), quien: getQuien(row, Math.abs(ingreso))});
        if (hasGasto) allMovimientos.push({date, monto: -Math.abs(gasto), quien: getQuien(row, -Math.abs(gasto))});
      }
    });

    if (!allMovimientos.length){
      skippedNote.style.display = 'block';
      skippedNote.textContent = 'No se pudo interpretar ninguna fila. Revisa las columnas seleccionadas.';
      resultsSection.style.display = 'none';
      return;
    }

    skippedNote.style.display = skippedCount > 0 ? 'block' : 'none';
    if (skippedCount > 0){
      skippedNote.textContent = `${skippedCount} fila(s) se omitieron por fecha o monto no reconocido.`;
    }

    // rango de fechas disponible, para guiar los selectores de filtro
    const fechas = allMovimientos.map(m => m.date);
    const minDate = new Date(Math.min(...fechas));
    const maxDate = new Date(Math.max(...fechas));
    filterDesde.min = filterHasta.min = toDateInputValue(minDate);
    filterDesde.max = filterHasta.max = toDateInputValue(maxDate);
    filterDesde.value = '';
    filterHasta.value = '';
    filterBusqueda.value = '';
    filterBusquedaWrap.style.display = nameConfig.mode === 'none' ? 'none' : '';

    currentTab = 'day';
    applyFilters();
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({behavior:'smooth', block:'start'});
  });

  function toDateInputValue(d){
    const y = d.getFullYear(), m = String(d.getMonth()+1).padStart(2,'0'), day = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${day}`;
  }

  function applyFilters(){
    const desde = filterDesde.value ? new Date(filterDesde.value + 'T00:00:00') : null;
    const hasta = filterHasta.value ? new Date(filterHasta.value + 'T23:59:59') : null;
    const busqueda = normalizeText(filterBusqueda.value.trim());

    movimientos = allMovimientos.filter(m => {
      if (desde && m.date < desde) return false;
      if (hasta && m.date > hasta) return false;
      if (busqueda && !normalizeText(m.quien || '').includes(busqueda)) return false;
      return true;
    });

    filterEmptyNote.style.display = movimientos.length === 0 ? 'block' : 'none';

    renderSummary();
    if (currentTab === 'person'){
      renderByPerson();
    } else {
      renderGroup(currentTab);
    }
  }

  const applyFiltersBtn = document.getElementById('applyFiltersBtn');
  applyFiltersBtn.addEventListener('click', applyFilters);
  filterBusqueda.addEventListener('keydown', e => { if (e.key === 'Enter') applyFilters(); });
  clearFiltersBtn.addEventListener('click', () => {
    filterDesde.value = '';
    filterHasta.value = '';
    filterBusqueda.value = '';
    applyFilters();
  });

  function fmt(n){
    return n.toLocaleString('es-PE', {minimumFractionDigits:2, maximumFractionDigits:2});
  }

  function renderSummary(){
    let totalIn = 0, totalOut = 0;
    movimientos.forEach(m => { if (m.monto >= 0) totalIn += m.monto; else totalOut += -m.monto; });
    document.getElementById('totalIngresos').textContent = fmt(totalIn);
    document.getElementById('totalGastos').textContent = fmt(totalOut);
    const neto = totalIn - totalOut;
    const netoEl = document.getElementById('totalNeto');
    netoEl.textContent = (neto >= 0 ? '' : '−') + fmt(Math.abs(neto));
    netoEl.classList.toggle('pos', neto >= 0);
    netoEl.classList.toggle('neg', neto < 0);
  }

  function isoWeekKey(d){
    const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dayNum = (date.getUTCDay() + 6) % 7;
    date.setUTCDate(date.getUTCDate() - dayNum + 3);
    const firstThursday = new Date(Date.UTC(date.getUTCFullYear(),0,4));
    const weekNum = 1 + Math.round(((date - firstThursday) / 86400000 - 3 + ((firstThursday.getUTCDay()+6)%7)) / 7);
    return `${date.getUTCFullYear()}-S${String(weekNum).padStart(2,'0')}`;
  }

  function groupKey(date, mode){
    const y = date.getFullYear();
    const m = String(date.getMonth()+1).padStart(2,'0');
    const d = String(date.getDate()).padStart(2,'0');
    if (mode === 'day') return `${y}-${m}-${d}`;
    if (mode === 'week') return isoWeekKey(date);
    if (mode === 'month') return `${y}-${m}`;
    return `${y}`;
  }

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentTab = btn.dataset.group;
      if (currentTab === 'person'){
        ledgerEl.style.display = 'none';
        personListEl.style.display = 'block';
        renderByPerson();
      } else {
        personListEl.style.display = 'none';
        ledgerEl.style.display = '';
        renderGroup(currentTab);
      }
    });
  });

  function renderByPerson(){
    const groups = {};
    movimientos.forEach(m => {
      const key = m.quien && m.quien.trim() ? m.quien.trim() : 'Sin nombre identificado';
      if (!groups[key]) groups[key] = { ingreso: 0, gasto: 0, moves: [] };
      if (m.monto >= 0) groups[key].ingreso += m.monto;
      else groups[key].gasto += -m.monto;
      groups[key].moves.push(m);
    });

    const keys = Object.keys(groups).sort((a,b) =>
      (groups[b].ingreso + groups[b].gasto) - (groups[a].ingreso + groups[a].gasto)
    );

    personListEl.innerHTML = keys.map(name => {
      const g = groups[name];
      const moves = g.moves.slice().sort((a,b) => b.date - a.date);
      const movesHtml = moves.map(m => {
        const isIn = m.monto >= 0;
        return `
          <div class="person-move-row">
            <span class="person-move-date">${m.date.toLocaleDateString('es-PE')}</span>
            <span class="person-move-amount ${isIn ? 'in' : 'out'}">${isIn ? '+' : '−'}${fmt(Math.abs(m.monto))}</span>
          </div>`;
      }).join('');

      return `
        <details class="person-card">
          <summary>
            <span class="person-name">${escapeHtml(name)}</span>
            <span class="person-totals">
              ${g.ingreso > 0 ? `<span class="p-in">+${fmt(g.ingreso)}</span>` : ''}
              ${g.gasto > 0 ? `<span class="p-out">−${fmt(g.gasto)}</span>` : ''}
            </span>
          </summary>
          <div class="person-moves">${movesHtml}</div>
        </details>`;
    }).join('');
  }

  function renderGroup(mode){
    const groups = {};
    movimientos.forEach(m => {
      const key = groupKey(m.date, mode);
      if (!groups[key]) groups[key] = {ingreso:0, gasto:0, moves:[]};
      if (m.monto >= 0) groups[key].ingreso += m.monto;
      else groups[key].gasto += -m.monto;
      groups[key].moves.push(m);
    });
    const keys = Object.keys(groups).sort().reverse();
    const maxVal = Math.max(1, ...keys.map(k => Math.max(groups[k].ingreso, groups[k].gasto)));

    ledgerEl.innerHTML = keys.map(k => {
      const g = groups[k];
      const neto = g.ingreso - g.gasto;
      const inPct = (g.ingreso / maxVal) * 100;
      const outPct = (g.gasto / maxVal) * 100;
      const moves = g.moves.slice().sort((a,b) => b.date - a.date);
      const movesHtml = moves.map(m => {
        const isIn = m.monto >= 0;
        const who = m.quien && m.quien.trim() ? m.quien.trim() : '';
        const dateLabel = m.date.toLocaleDateString('es-PE', {day:'2-digit', month:'2-digit', year:'numeric'});
        return `
          <div class="ledger-move-row">
            <span class="ledger-move-who">${dateLabel}${who ? ' — ' + escapeHtml(who) : ''}</span>
            <span class="ledger-move-amount ${isIn ? 'in' : 'out'}">${isIn ? '+' : '−'}${fmt(Math.abs(m.monto))}</span>
          </div>`;
      }).join('');

      return `
        <details class="ledger-row">
          <summary>
            <div class="ledger-period">${k}</div>
            <div class="ledger-bars">
              <div class="bar-track"><div class="bar-fill ingreso" style="width:${inPct}%"></div></div>
              <div class="bar-track"><div class="bar-fill gasto" style="width:${outPct}%"></div></div>
            </div>
            <div class="ledger-values">
              <div class="v-in">+${fmt(g.ingreso)}</div>
              <div class="v-out">−${fmt(g.gasto)}</div>
              <div class="v-net">${neto >= 0 ? '' : '−'}${fmt(Math.abs(neto))} neto</div>
            </div>
            <div class="ledger-expand-icon"></div>
          </summary>
          <div class="ledger-moves">${movesHtml}</div>
        </details>`;
    }).join('');
  }
})();
