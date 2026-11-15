(function () {
  'use strict';
  const { STATUSES, DEFAULT_CODES, parseWorkbook, exportWorkbook } = window.ListaExcel;
  const LABELS = { A: 'Asistencia', F: 'Falta', R: 'Retardo', J: 'Justificado' };
  const STORE_KEY = 'paseDeLista.v1';
  const $ = (s) => document.querySelector(s);

  // ---------- Estado ----------
  // { fileName, fileB64, groups, codes, ui: { group, date } }
  let state = load();

  function load() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) { /* almacenamiento no disponible */ }
    return null;
  }
  function save() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
    catch (e) { toast('No se pudo guardar en el dispositivo'); }
  }

  const todayISO = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };
  const group = () => state.groups[state.ui.group];
  const dayRecords = (create) => {
    const g = group();
    if (!g.records[state.ui.date] && create) g.records[state.ui.date] = {};
    return g.records[state.ui.date] || {};
  };
  const fmtDate = (iso) => {
    const [y, m, d] = iso.split('-').map(Number);
    return new Date(y, m - 1, d).toLocaleDateString('es-MX', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
  };

  // ---------- Base64 <-> ArrayBuffer ----------
  function bufToB64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  function b64ToBuf(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
  }

  // ---------- Subir archivo ----------
  $('#file-input').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const codes = (state && state.codes) || { ...DEFAULT_CODES };
      const groups = parseWorkbook(wb, codes);
      if (!groups.length) {
        toast('No encontré una columna “NOMBRE” en ninguna hoja');
        return;
      }
      state = { fileName: file.name, fileB64: bufToB64(buf), groups, codes, ui: { group: 0, date: todayISO() } };
      save();
      const total = groups.reduce((n, g) => n + g.students.length, 0);
      toast(`${groups.length} grupo(s), ${total} alumnos cargados`);
      render();
    } catch (err) {
      console.error(err);
      toast('No se pudo leer el archivo. ¿Es un .xlsx?');
    }
  });

  // ---------- Render ----------
  function render() {
    const has = !!(state && state.groups && state.groups.length);
    $('#screen-upload').hidden = has;
    $('#screen-main').hidden = !has;
    if (!has) return;

    const sel = $('#group-select');
    sel.innerHTML = state.groups
      .map((g, i) => `<option value="${i}">${escapeHtml(g.sheet)} · ${g.students.length} alumnos</option>`)
      .join('');
    sel.value = state.ui.group;
    $('#date-input').value = state.ui.date;
    startCard();
    refresh();
  }

  // ---------- Vistas: uno por uno (por defecto) / lista ----------
  let view = 'card';
  let cardIdx = 0;   // índice del alumno en pantalla; = número de alumnos cuando ya terminó
  let busy = false;  // evita doble toque mientras avanza

  function refresh() {
    const opt = $('#group-select').options[state.ui.group];
    if (opt) opt.textContent = `${group().sheet} · ${group().students.length} alumnos`;
    $('#card-view').hidden = view !== 'card';
    $('#list-view').hidden = view !== 'list';
    if (view === 'card') renderCard(false);
    else renderStudents();
    renderSummary();
  }

  function firstPending() {
    const g = group();
    const rec = dayRecords(false);
    return g.students.findIndex((s) => !rec[s.row]);
  }
  // Al abrir un grupo/fecha empieza en el primer alumno sin marcar.
  function startCard() {
    const i = firstPending();
    cardIdx = i < 0 ? group().students.length : i;
  }

  function renderCard(animate) {
    const g = group();
    const rec = dayRecords(false);
    const n = g.students.length;
    const marked = g.students.filter((s) => rec[s.row]).length;
    const finished = cardIdx >= n;

    $('#progress-fill').style.width = `${(marked / n) * 100}%`;
    ['#card', '#big-choices', '.card-nav'].forEach((sel) => ($(sel).hidden = finished));
    $('#done').hidden = !finished;

    if (finished) {
      $('#card-pos').textContent = `${marked} de ${n} marcados`;
      const c = { A: 0, F: 0, R: 0, J: 0 };
      g.students.forEach((s) => rec[s.row] && c[rec[s.row]]++);
      $('#done-summary').innerHTML = STATUSES.map((k) => `<span class="${k}">${LABELS[k]}: ${c[k]}</span>`).join('');
      const pending = n - marked;
      $('#done-pending').textContent = pending ? `Quedaron ${pending} alumno(s) sin marcar.` : 'Todos los alumnos tienen marca.';
      $('#done-restart').textContent = pending ? 'Ir a los pendientes' : 'Volver al inicio';
      return;
    }

    const s = g.students[cardIdx];
    const st = rec[s.row] || '';
    $('#card-pos').textContent = `Alumno ${cardIdx + 1} de ${n}`;
    $('#card-no').textContent = s.no ? `No. ${s.no}` : '';
    $('#card-name').innerHTML = escapeHtml(s.name) + (s.added ? ' <span class="badge-new">NUEVO</span>' : '');
    $('#card-status').textContent = st ? `Marcado: ${LABELS[st]}` : '';
    const card = $('#card');
    card.className = `card ${st}`;
    if (animate) { void card.offsetWidth; card.classList.add('slide'); }
    document.querySelectorAll('.big-choice').forEach((b) => b.classList.toggle('on', b.dataset.s === st));
    $('#card-prev').disabled = cardIdx === 0;
  }

  // Pide autorización antes de cambiar (o quitar) una marca que ya existía.
  function confirmChange(name, from, to) {
    return new Promise((resolve) => {
      const d = $('#confirm-change');
      const pill = (sel, st) => {
        const el = $(sel);
        el.className = `pill-lg ${st || 'none'}`;
        el.textContent = st ? LABELS[st] : 'Sin marca';
      };
      $('#confirm-name').textContent = name;
      pill('#confirm-from', from);
      pill('#confirm-to', to);
      $('#confirm-ok').textContent = to ? 'Sí, cambiar' : 'Sí, quitar';
      d.returnValue = '';
      d.addEventListener('close', () => resolve(d.returnValue === 'ok'), { once: true });
      d.showModal();
    });
  }

  async function markCard(status) {
    const g = group();
    if (busy || cardIdx >= g.students.length) return;
    const s = g.students[cardIdx];
    const prev = dayRecords(false)[s.row];
    if (prev && prev !== status) {
      busy = true;
      const ok = await confirmChange(s.name, prev, status);
      busy = false;
      if (!ok) return; // se queda en el mismo alumno, sin cambios
      toast(`Cambiado a ${LABELS[status]}`);
    }
    const rec = dayRecords(true);
    rec[s.row] = status;
    save();
    document.querySelectorAll('.big-choice').forEach((b) => b.classList.toggle('on', b.dataset.s === status));
    if (navigator.vibrate) navigator.vibrate(15);
    busy = true;
    setTimeout(() => {   // breve pausa para ver la selección y pasa al siguiente
      busy = false;
      cardIdx++;
      renderCard(true);
      renderSummary();
    }, 200);
  }

  $('#big-choices').addEventListener('click', (e) => {
    const b = e.target.closest('.big-choice');
    if (b) markCard(b.dataset.s);
  });
  $('#card-prev').addEventListener('click', () => { if (cardIdx > 0) { cardIdx--; renderCard(true); } });
  $('#card-next').addEventListener('click', () => { cardIdx++; renderCard(true); });
  $('#done-review').addEventListener('click', () => { view = 'list'; refresh(); });
  $('#done-restart').addEventListener('click', () => {
    const i = firstPending();
    cardIdx = i < 0 ? 0 : i;
    renderCard(true);
  });
  $('#to-list').addEventListener('click', () => { view = 'list'; refresh(); });
  $('#to-card').addEventListener('click', () => { view = 'card'; startCard(); refresh(); });

  // Teclado (en computadora): A/F/R/J para marcar, flechas para moverse.
  document.addEventListener('keydown', (e) => {
    if (view !== 'card' || $('#screen-main').hidden || document.querySelector('dialog[open]')) return;
    if (/^(INPUT|SELECT|TEXTAREA)$/.test(e.target.tagName)) return;
    const k = e.key.toUpperCase();
    if (STATUSES.includes(k)) markCard(k);
    else if (e.key === 'ArrowLeft') $('#card-prev').click();
    else if (e.key === 'ArrowRight') $('#card-next').click();
  });

  function renderStudents() {
    const g = group();
    const rec = dayRecords(false);
    const q = normalize($('#search').value);
    const list = $('#students');
    list.innerHTML = g.students
      .map((s, i) => [s, i])
      .filter(([s]) => !q || normalize(s.name).includes(q))
      .map(([s, i]) => {
        const st = rec[s.row] || '';
        const btns = STATUSES.map(
          (k) => `<button class="choice ${st === k ? 'on' : ''}" data-s="${k}" aria-pressed="${st === k}">${k}<small>${LABELS[k]}</small></button>`
        ).join('');
        return `<li class="student ${st}" data-row="${s.row}" data-idx="${i}">
          <div class="student-name"><span class="no">${escapeHtml(s.no)}</span><span>${escapeHtml(s.name)}${s.added ? '<span class="badge-new">NUEVO</span>' : ''}</span></div>
          <div class="choices">${btns}</div></li>`;
      })
      .join('');
    renderSummary();
  }

  function renderSummary() {
    const g = group();
    const rec = dayRecords(false);
    const counts = { A: 0, F: 0, R: 0, J: 0 };
    Object.values(rec).forEach((s) => counts[s]++);
    const marked = Object.keys(rec).length;
    $('#summary').innerHTML =
      `<span>${fmtDate(state.ui.date)}</span><span>${marked}/${g.students.length} marcados</span>` +
      STATUSES.map((k) => `<span class="${k}">${k}: ${counts[k]}</span>`).join('');
  }

  // ---------- Pasar lista ----------
  // ---------- Quitar alumno (solo los agregados desde la app) ----------
  // Está fuera del pase de lista a propósito: se busca al alumno, se elige y se confirma.
  let removeRow = null;

  function openRemoveStudent() {
    removeRow = null;
    $('#remove-search').value = '';
    renderRemoveList();
    $('#remove-student').showModal();
  }

  function renderRemoveList() {
    const g = group();
    const q = normalize($('#remove-search').value);
    const added = g.students.filter((s) => s.added);
    const shown = added.filter((s) => !q || normalize(s.name).includes(q));
    $('#remove-list').innerHTML = !added.length
      ? '<li class="empty">No hay alumnos agregados en este grupo. Los alumnos que vienen del Excel original no se pueden quitar.</li>'
      : !shown.length
        ? '<li class="empty">Ningún alumno coincide con la búsqueda.</li>'
        : shown.map((s) => `<li data-row="${s.row}" class="${s.row === removeRow ? 'sel' : ''}"><span class="no">${escapeHtml(s.no)}</span>${escapeHtml(s.name)}</li>`).join('');
    const sel = g.students.find((s) => s.row === removeRow);
    const ok = $('#remove-ok');
    ok.disabled = !sel;
    ok.textContent = sel ? 'Quitar alumno elegido' : 'Elige un alumno';
  }

  $('#list-remove').addEventListener('click', openRemoveStudent);
  $('#remove-search').addEventListener('input', renderRemoveList);
  $('#remove-list').addEventListener('click', (e) => {
    const li = e.target.closest('li[data-row]');
    if (!li) return;
    removeRow = Number(li.dataset.row);
    renderRemoveList();
  });
  $('#remove-ok').addEventListener('click', () => {
    const g = group();
    const idx = g.students.findIndex((s) => s.row === removeRow && s.added);
    if (idx < 0) return;
    const s = g.students[idx];
    g.students.splice(idx, 1);
    Object.values(g.records).forEach((day) => delete day[s.row]);
    save();
    if (cardIdx > g.students.length) cardIdx = g.students.length;
    $('#remove-student').close();
    refresh();
    toast(`${s.name} quitado de ${g.sheet}`);
  });

  $('#students').addEventListener('click', async (e) => {
    const nameEl = e.target.closest('.student-name');
    if (nameEl) {   // tocar el nombre: pasar lista uno por uno desde ese alumno
      cardIdx = Number(nameEl.closest('.student').dataset.idx);
      view = 'card';
      refresh();
      return;
    }
    const btn = e.target.closest('.choice');
    if (!btn) return;
    const li = btn.closest('.student');
    const row = li.dataset.row;
    const s = btn.dataset.s;
    const prev = dayRecords(false)[row];
    const next = prev === s ? undefined : s; // tocar la misma otra vez = quitar marca
    if (prev) {
      const name = group().students.find((x) => String(x.row) === row).name;
      if (!(await confirmChange(name, prev, next))) return;
    }
    const rec = dayRecords(true);
    if (next) rec[row] = next;
    else delete rec[row];
    save();
    li.className = `student ${rec[row] || ''}`;
    li.querySelectorAll('.choice').forEach((b) => {
      const on = b.dataset.s === rec[row];
      b.classList.toggle('on', on);
      b.setAttribute('aria-pressed', on);
    });
    renderSummary();
  });


  // Cambió el grupo o la fecha: vuelve a empezar desde el primer pendiente.
  function changed() {
    save();
    startCard();
    refresh();
  }
  $('#group-select').addEventListener('change', (e) => {
    state.ui.group = Number(e.target.value);
    $('#search').value = '';
    changed();
  });
  $('#date-input').addEventListener('change', (e) => {
    if (!e.target.value) return;
    state.ui.date = e.target.value;
    changed();
  });
  function shiftDay(n) {
    const [y, m, d] = state.ui.date.split('-').map(Number);
    const dt = new Date(y, m - 1, d + n);
    state.ui.date = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
    $('#date-input').value = state.ui.date;
    changed();
  }
  $('#prev-day').addEventListener('click', () => shiftDay(-1));
  $('#next-day').addEventListener('click', () => shiftDay(1));
  $('#search').addEventListener('input', renderStudents);

  // ---------- Exportar ----------
  $('#export-btn').addEventListener('click', () => doExport($('#export-btn')));
  $('#done-export').addEventListener('click', () => doExport($('#done-export')));
  $('#preview-export').addEventListener('click', () => doExport($('#preview-export')));
  $('#done-preview').addEventListener('click', openPreview);
  $('#list-preview').addEventListener('click', openPreview);

  // ---------- Agregar alumno ----------
  $('#list-add').addEventListener('click', openAddStudent);

  function openAddStudent() {
    const g = group();
    const nums = g.students.map((s) => Number(s.no)).filter(Number.isFinite);
    $('#add-group').textContent = `Grupo ${g.sheet} · ${g.students.length} alumnos`;
    $('#add-name').value = '';
    $('#add-no').value = nums.length ? Math.max(...nums) + 1 : g.students.length + 1;
    $('#add-error').textContent = '';
    const d = $('#add-student');
    d.returnValue = '';
    d.showModal();
    $('#add-name').focus();
  }

  $('#add-form').addEventListener('submit', (e) => {
    if (e.submitter && e.submitter.value === 'cancel') return; // cierra sin agregar
    e.preventDefault();
    const g = group();
    const name = $('#add-name').value.replace(/\s+/g, ' ').trim().toUpperCase();
    const no = $('#add-no').value.trim();
    if (!name) { $('#add-error').textContent = 'Escribe el nombre del alumno.'; return; }
    if (g.students.some((s) => normalize(s.name) === normalize(name))) {
      $('#add-error').textContent = 'Ya hay un alumno con ese nombre en este grupo.';
      return;
    }
    // Va en la siguiente fila libre debajo del último alumno.
    const row = Math.max(g.headerRow, ...g.students.map((s) => s.row)) + 1;
    g.students.push({ row, no, name, added: true });
    save();
    $('#add-student').close();
    // Lo deja en pantalla para pasarle lista de inmediato.
    view = 'card';
    cardIdx = g.students.length - 1;
    refresh();
    toast(`${name} agregado a ${g.sheet}`);
  });

  // ---------- Vista previa ----------
  // Se genera el mismo Excel que al exportar y se vuelve a leer, así lo que se ve es exactamente lo que se descarga.
  async function openPreview() {
    const g = group();
    let ws;
    try {
      const out = await exportWorkbook(ExcelJS, b64ToBuf(state.fileB64), [g], state.codes);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(out);
      ws = wb.getWorksheet(g.sheet);
    } catch (err) {
      console.error(err);
      toast('No se pudo generar la vista previa');
      return;
    }

    const dates = Object.keys(g.records).filter((d) => Object.keys(g.records[d]).length).sort();
    const lastCol = Math.max(g.nameCol, g.startCol + dates.length - 1);
    const curCol = dates.includes(state.ui.date) ? g.startCol + dates.indexOf(state.ui.date) : -1;
    const colLetter = (c) => { let s = ''; for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };
    const statusOf = (v) => STATUSES.find((k) => String(state.codes[k]).toUpperCase() === String(v).toUpperCase()) || '';
    const text = (v) => {
      if (v instanceof Date) return `${String(v.getUTCDate()).padStart(2, '0')}/${String(v.getUTCMonth() + 1).padStart(2, '0')}/${String(v.getUTCFullYear()).slice(-2)}`;
      return ListaExcel.cellText(v).replace(/\s+/g, ' ').trim();
    };

    let html = '<thead><tr><th class="rownum"></th>';
    for (let c = 1; c <= lastCol; c++) html += `<th class="${c <= g.nameCol ? 'stick' : ''}">${colLetter(c)}</th>`;
    html += '</tr></thead><tbody>';
    const rows = [g.headerRow, ...g.students.map((s) => s.row)];
    rows.forEach((r, i) => {
      const isHdr = i === 0;
      html += `<tr class="${isHdr ? 'hdr' : ''}"><td class="rownum">${r}</td>`;
      for (let c = 1; c <= lastCol; c++) {
        const v = ws.getCell(r, c).value;
        const t = text(v);
        const cls = [];
        if (c <= g.nameCol) cls.push('stick');
        if (c === g.nameCol) cls.push('name');
        const isDateCol = c >= g.startCol && c < g.startCol + dates.length;
        if (isDateCol && isHdr) cls.push('date-h');
        if (isDateCol && !isHdr) cls.push('v', statusOf(t));
        if (c === curCol) cls.push('cur');
        html += `<td class="${cls.join(' ')}">${escapeHtml(t)}</td>`;
      }
      html += '</tr>';
    });
    html += '</tbody>';
    $('#preview-table').innerHTML = html;

    $('#preview-title').textContent = `Vista previa · ${g.sheet}`;
    $('#preview-sub').textContent = dates.length
      ? `${g.students.length} alumnos · ${dates.length} fecha(s) de pase de lista`
      : 'Aún no hay pases de lista en este grupo.';
    $('#preview-legend').innerHTML = STATUSES.map((k) => `<span class="summary-${k}">${escapeHtml(String(state.codes[k]))} = ${LABELS[k]}</span>`).join('');

    const d = $('#preview');
    d.showModal();
    // Columnas fijas (No. y Nombre) al desplazar horizontalmente.
    const firstRow = d.querySelectorAll('#preview-table tbody tr:first-child td');
    let left = 0;
    const lefts = [];
    firstRow.forEach((td, idx) => { if (idx <= g.nameCol) { lefts.push(left); left += td.offsetWidth; } });
    d.querySelectorAll('#preview-table tr').forEach((tr) =>
      [...tr.children].forEach((cell, idx) => { if (idx <= g.nameCol) cell.style.left = `${lefts[idx]}px`; })
    );
    // Desplaza hasta la fecha actual (o la última) para verla directo.
    const target = d.querySelector('#preview-table td.cur') || d.querySelector('#preview-table td.date-h:last-of-type');
    if (target) d.querySelector('.preview-scroll').scrollLeft = Math.max(0, target.offsetLeft - left - 40);
  }

  async function doExport(btn) {
    if (btn) { btn.disabled = true; btn.textContent = 'Generando…'; }
    try {
      const out = await exportWorkbook(ExcelJS, b64ToBuf(state.fileB64), state.groups, state.codes);
      const name = state.fileName.replace(/\.xlsx$/i, '') + ' - asistencia.xlsx';
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const file = new File([blob], name, { type: blob.type });
      // En celular: menú de compartir (WhatsApp, correo, Drive…). En PC: descarga.
      if (navigator.canShare && navigator.canShare({ files: [file] }) && /Android|iPhone|iPad/i.test(navigator.userAgent)) {
        await navigator.share({ files: [file], title: name }).catch(() => {});
      } else {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = name;
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 2000);
      }
      toast('Excel generado');
    } catch (err) {
      console.error(err);
      toast('Error al generar el Excel');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Exportar Excel'; }
    }
  }

  // ---------- Menú y diálogos ----------
  const menu = $('#menu');
  $('#menu-btn').addEventListener('click', () => menu.showModal());
  document.querySelectorAll('dialog').forEach((d) =>
    d.addEventListener('click', (e) => {
      if (e.target === d || e.target.dataset.action === 'close') d.close();
    })
  );

  menu.addEventListener('click', (e) => {
    const action = e.target.dataset.action;
    if (!action || action === 'close') return;
    menu.close();
    if (action === 'export') doExport(null);
    if (action === 'preview') openPreview();
    if (action === 'add-student') openAddStudent();
    if (action === 'history') openHistory();
    if (action === 'settings') openSettings();
    if (action === 'clear-day') {
      if (confirm(`¿Borrar el pase de lista del ${fmtDate(state.ui.date)} en ${group().sheet}?`)) {
        delete group().records[state.ui.date];
        changed();
      }
    }
    if (action === 'new-file') {
      if (confirm('Se reemplazará la lista actual y su asistencia guardada. Exporta antes si la necesitas. ¿Continuar?')) {
        $('#file-input').click();
      }
    }
  });

  function openHistory() {
    const g = group();
    const dates = Object.keys(g.records).filter((d) => Object.keys(g.records[d]).length).sort().reverse();
    $('#history-list').innerHTML = dates.length
      ? dates.map((d) => {
          const c = { A: 0, F: 0, R: 0, J: 0 };
          Object.values(g.records[d]).forEach((s) => c[s]++);
          return `<li data-date="${d}"><span>${fmtDate(d)}</span><span class="counts">A ${c.A} · F ${c.F} · R ${c.R} · J ${c.J}</span></li>`;
        }).join('')
      : '<li>Aún no hay pases de lista en este grupo.</li>';
    $('#history').showModal();
  }
  $('#history-list').addEventListener('click', (e) => {
    const li = e.target.closest('li[data-date]');
    if (!li) return;
    state.ui.date = li.dataset.date;
    $('#date-input').value = state.ui.date;
    changed();
    $('#history').close();
  });

  function openSettings() {
    const f = $('#settings-form');
    STATUSES.forEach((k) => (f.elements[k].value = state.codes[k]));
    $('#settings').showModal();
  }
  $('#settings').addEventListener('close', () => {
    if ($('#settings').returnValue !== 'save') return;
    const f = $('#settings-form');
    const vals = STATUSES.map((k) => f.elements[k].value.trim());
    if (vals.some((v) => v === '') || new Set(vals.map((v) => v.toUpperCase())).size !== vals.length) {
      toast('Cada estado necesita un valor distinto');
      return;
    }
    STATUSES.forEach((k, i) => (state.codes[k] = vals[i]));
    save();
    toast('Valores guardados');
  });

  // ---------- Utilidades ----------
  function normalize(s) {
    return (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 2600);
  }

  // ---------- Inicio ----------
  if (state) {
    if (!state.ui) state.ui = { group: 0, date: todayISO() };
    if (state.ui.group >= state.groups.length) state.ui.group = 0;
    state.ui.date = todayISO(); // al abrir la app siempre empieza en hoy
  }
  render();

  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
})();
