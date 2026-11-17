(function () {
  'use strict';
  const { STATUSES, READ, DEFAULT_CODES, TYPES, ASIST_ID, sectionsOf, newSectionId, secMin, parseWorkbook, exportWorkbook, groupDates, sheetLayout,
    DEFAULT_GRADING, finalGrade, gradingOn, weightOf, round1, worksOf, normalizeGroup, studentReport, exportStudent, groupReport, exportGroupReport } = window.ListaExcel;
  const LABELS = { A: 'Asistencia', F: 'Falta', R: 'Retardo', J: 'Justificado' };
  // Opciones de captura según la plantilla de la sección: valor guardado, texto del botón y color (clase A/F/R/J/N).
  const OPTIONS = {
    sino: [
      { v: 'S', btn: 'Sí', color: 'A' },
      { v: 'N', btn: 'No', color: 'F' },
    ],
    trab: [
      { v: 0, btn: '0', color: 'F' },
      { v: 50, btn: '50', color: 'R' },
      { v: 100, btn: '100', color: 'A' },
    ],
    // Disciplina: todos tienen 10 sin capturar nada; solo se guarda el 0.
    disc: [
      { v: 10, btn: '10', color: 'N' },
      { v: 0, btn: '0', color: 'F' },
    ],
  };
  const DISC_DEFAULT = 10;
  // Valores del botón "restantes" según la plantilla.
  const FILL_REST = { sino: ['S', 'N'], trab: [0] };
  // Secciones del grupo actual (cada maestro las agrega, quita y renombra).
  const secs = () => sectionsOf(group());
  const secById = (id) => secs().find((x) => x.id === id);
  const tab = () => (state.ui.tab && secById(state.ui.tab) ? state.ui.tab : ASIST_ID);
  const typeOf = (id) => (secById(id) || { type: 'asist' }).type;
  const optionOf = (id, v) => {
    const sec = secById(id) || { type: 'asist', name: 'Asistencia' };
    if (sec.type === 'asist') return v ? { v, label: LABELS[v], color: v } : null;
    if (sec.type === 'num') return v != null ? { v, label: `${sec.name}: ${v}`, color: 'N' } : null;
    const o = OPTIONS[sec.type].find((x) => x.v === v);
    return o ? { ...o, label: `${sec.name}: ${o.btn}` } : null;
  };
  // Datos de una sección para el día seleccionado.
  // Trabajos: puede haber varios por día; se usa el trabajo seleccionado (state.ui.work).
  const works = (create) => {
    const g = group();
    const id = tab();
    g[id] = g[id] || {};
    if (!g[id][state.ui.date] && create) g[id][state.ui.date] = [];
    return g[id][state.ui.date] || [];
  };
  const workIdx = () => Math.max(0, Math.min(state.ui.work || 0, Math.max(works(false).length - 1, 0)));
  const curWork = (create) => {
    const ws = works(create);
    const i = workIdx();
    if (!ws[i] && create) ws[i] = { title: '', vals: {} };
    return ws[i] || null;
  };
  const dayData = (id, create) => {
    if (typeOf(id) === 'trab') { const w = curWork(create); return w ? w.vals : {}; }
    const g = group();
    g[id] = g[id] || {};
    if (!g[id][state.ui.date] && create) g[id][state.ui.date] = {};
    return g[id][state.ui.date] || {};
  };
  // Porcentajes de un grupo, completos (con los valores por defecto que falten).
  const fixGrading = (gr) => JSON.parse(JSON.stringify({ ...DEFAULT_GRADING, ...(gr || {}), w: { ...DEFAULT_GRADING.w, ...((gr || {}).w || {}) } }));
  // Nombre del grupo para mostrar: hoja del Excel y, si el maestro la puso, la materia (p. ej. "V205 · Ética").
  const gname = (g) => (g.subject ? `${g.sheet} · ${g.subject}` : g.sheet);
  // En el selector: con materia se omite el número de alumnos para que quepa el nombre.
  const glabel = (g) => (g.subject ? gname(g) : `${g.sheet} · ${g.students.length} alumnos`);
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
      // Si el grupo ya existía (misma hoja), conserva sus secciones, porcentajes y materia.
      const old = (state && state.groups) || [];
      const prevOf = (sheet) => old.find((o) => o.sheet === sheet);
      const groups = parseWorkbook(wb, codes, (sheet) => (prevOf(sheet) || {}).sections);
      if (!groups.length) {
        toast('No encontré una columna “NOMBRE” en ninguna hoja');
        return;
      }
      groups.forEach((g) => {
        const prev = prevOf(g.sheet);
        normalizeGroup(g);
        g.grading = fixGrading(prev && prev.grading);
        if (prev && prev.subject) g.subject = prev.subject; // la materia también se conserva
      });
      state = { fileName: file.name, fileB64: bufToB64(buf), groups, codes: { ...DEFAULT_CODES, ...codes }, ui: { group: 0, date: todayISO(), tab: ASIST_ID } };
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
      .map((g, i) => `<option value="${i}">${escapeHtml(glabel(g))}</option>`)
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
    if (opt) opt.textContent = glabel(group());
    const t = tab();
    renderTabs();
    const att = typeOf(t) === 'asist';
    $('#card-view').hidden = !att || view !== 'card';
    $('#list-view').hidden = !att || view !== 'list';
    $('#compact-view').hidden = att;
    if (!att) renderCompact();
    else if (view === 'card') renderCard(false);
    else renderStudents();
    renderSummary();
  }

  // ---------- Pestañas (una por sección del grupo) ----------
  function renderTabs() {
    const t = tab();
    const box = $('#tabs');
    const html = secs().map((x) => `<button class="tab ${x.id === t ? 'on' : ''}" role="tab" aria-selected="${x.id === t}" data-tab="${x.id}">${escapeHtml(x.name)}</button>`).join('');
    if (box.innerHTML !== html) {
      box.innerHTML = html;
      const on = box.querySelector('.tab.on');
      if (on) box.scrollLeft = Math.max(0, on.offsetLeft - box.clientWidth / 2 + on.offsetWidth / 2);
    }
  }
  $('#tabs').addEventListener('click', (e) => {
    const b = e.target.closest('.tab');
    if (!b || b.dataset.tab === tab()) return;
    state.ui.tab = b.dataset.tab;
    save();
    state.ui.work = 0;
    if (typeOf(state.ui.tab) === 'asist') startCard();
    refresh();
    window.scrollTo(0, 0);
  });

  // ---------- Vista compacta (lecturas / trabajos) ----------
  function renderCompact() {
    const g = group();
    const cat = tab();
    const t = typeOf(cat);
    const sec = secById(cat);
    const day = dayData(cat, false);
    const att = dayData(ASIST_ID, false);
    renderWorkBar();
    const disc = t === 'disc';
    const exam = t === 'num';
    $('#compact-search-bar').hidden = !disc && !exam;
    $('#compact-hint').innerHTML = exam
      ? `Escribe la calificación de <b>${escapeHtml(sec.name)}</b> (0 a 100). Con <b>Enter</b> pasas al siguiente alumno.`
      : 'Todos tienen 10. Busca al alumno y toca <b>0</b> si se portó mal.';
    $('#compact-list').innerHTML = g.students.map((s) => {
      const cur = optionOf(cat, disc ? (day[s.row] ?? DISC_DEFAULT) : day[s.row]);
      const a = att[s.row];
      const btns = exam
        ? `<input class="exam-in" type="text" inputmode="decimal" enterkeyhint="next" maxlength="5" value="${day[s.row] ?? ''}" aria-label="${escapeHtml(sec.name)} de ${escapeHtml(s.name)}" placeholder="—">`
        : OPTIONS[t].map((o) => {
        const on = cur && cur.v === o.v;
        return `<button class="cb ${t === 'trab' && o.v === 100 ? 'wide' : ''} ${on ? 'on c' + o.color : ''}" data-v="${o.v}" aria-pressed="${!!on}" aria-label="${escapeHtml(sec.name)}: ${o.btn}">${o.btn}</button>`;
      }).join('');
      return `<li class="crow ${cur ? 'c' + cur.color : ''} ${a === 'F' ? 'absent' : ''}" data-row="${s.row}" data-name="${escapeHtml(normalize(s.name))}">
        <span class="cno">${escapeHtml(s.no)}</span>
        <span class="cname">${a ? `<i class="att ${a}" title="${LABELS[a]}">${a}</i>` : ''}${escapeHtml(s.name)}</span>
        <span class="cbtns">${btns}</span></li>`;
    }).join('');
    if (disc || exam) filterCompact();
    updateFillRest();
  }

  // ---------- Botón "restantes" (lecturas: No · trabajos: 0) ----------
  const pendingRows = () => {
    const day = dayData(tab(), false);
    return group().students.filter((s) => day[s.row] == null);
  };
  // Sí/No: si todos tienen la misma marca, botón para cambiarlos todos a la otra.
  const flipTarget = () => {
    if (typeOf(tab()) !== 'sino') return null;
    const day = dayData(tab(), false);
    const vals = new Set(group().students.map((s) => day[s.row]));
    if (vals.size !== 1 || !group().students.length) return null;
    const [only] = vals;
    return only === 'S' ? 'N' : only === 'N' ? 'S' : null;
  };
  function updateFillRest() {
    const t = typeOf(tab());
    const box = $('#fill-rest');
    const n = t in FILL_REST ? pendingRows().length : 0;
    const flip = n ? null : flipTarget();
    box.hidden = !n && !flip;
    if (flip) {
      const total = group().students.length;
      box.innerHTML = `<button class="btn fill-rest ${flip === 'S' ? 'yes' : ''}" data-all="${flip}">Cambiar a los ${total} a “${flip === 'S' ? 'Sí' : 'No'}”</button>`;
      return;
    }
    if (!n) return;
    const text = { S: `Poner “Sí” a los ${n} sin marcar`, N: `Poner “No” a los ${n} sin marcar`, 0: `Poner 0 a los ${n} sin calificar` };
    box.innerHTML = FILL_REST[t].map((v) => `<button class="btn fill-rest ${v === 'S' ? 'yes' : ''}" data-v="${v}">${text[v]}</button>`).join('');
  }
  $('#fill-rest').addEventListener('click', async (e) => {
    const all = e.target.closest('button[data-all]');
    if (all) {
      const to = all.dataset.all;
      if (flipTarget() !== to) return;
      const label = to === 'S' ? 'Sí' : 'No';
      const students = group().students;
      if (!(await confirmBulk(`¿Cambiar a los ${students.length} alumnos a “${label}”?`,
        `Todos tienen “${to === 'S' ? 'No' : 'Sí'}” en este día y quedarán con “${label}”.`, 'Sí, cambiar'))) return;
      const day = dayData(tab(), true);
      students.forEach((s) => (day[s.row] = to));
      save();
      if (navigator.vibrate) navigator.vibrate(15);
      refresh();
      toast(`${students.length} alumno(s) con ${label}`);
      return;
    }
    const btn = e.target.closest('button[data-v]');
    const cat = tab();
    const t = typeOf(cat);
    const rows = pendingRows();
    if (!btn || !(t in FILL_REST) || !rows.length) return;
    const v = t === 'trab' ? Number(btn.dataset.v) : btn.dataset.v;
    const what = { S: '“Sí”', N: '“No”' }[v] || '0';
    if (!(await confirmBulk(`¿Poner ${what} a ${rows.length} alumno(s)?`, 'Solo a los que siguen en blanco. Los que ya marcaste no cambian.'))) return;
    const day = dayData(cat, true);
    rows.forEach((s) => (day[s.row] = v));
    save();
    if (navigator.vibrate) navigator.vibrate(15);
    refresh();
    toast(`${rows.length} alumno(s) con ${{ S: 'Sí', N: 'No' }[v] || '0'}`);
  });
  // Confirmación propia de la app (los avisos del navegador pueden bloquearse en el celular).
  function confirmBulk(title, text, okText = 'Sí, ponerlo', danger = false) {
    return new Promise((resolve) => {
      const d = $('#confirm-bulk');
      $('#bulk-title').textContent = title;
      $('#bulk-text').textContent = text;
      $('#bulk-ok').textContent = okText;
      $('#bulk-ok').className = `btn ${danger ? 'danger' : 'primary'}`;
      d.returnValue = '';
      d.addEventListener('close', () => resolve(d.returnValue === 'ok'), { once: true });
      d.showModal();
    });
  }

  // Buscador de la pestaña Disciplina.
  function filterCompact() {
    const q = normalize($('#compact-search').value);
    let shown = 0;
    document.querySelectorAll('#compact-list .crow').forEach((li) => {
      const ok = !q || li.dataset.name.includes(q);
      li.hidden = !ok;
      if (ok) shown++;
    });
    $('#compact-empty').hidden = shown > 0;
  }
  $('#compact-search').addEventListener('input', filterCompact);

  $('#compact-list').addEventListener('click', async (e) => {
    const btn = e.target.closest('.cb');
    if (!btn) return;
    const cat = tab();
    const t = typeOf(cat);
    const li = btn.closest('.crow');
    const row = li.dataset.row;
    const disc = t === 'disc';
    const v = t === 'sino' ? btn.dataset.v : Number(btn.dataset.v);
    let prev = dayData(cat, false)[row];
    let next = prev === v ? undefined : v; // tocar la misma otra vez = quitar
    if (disc) {
      // Sin captura = 10. Poner 0 no pide confirmación; regresar a 10 sí.
      prev = prev ?? DISC_DEFAULT;
      if (v === prev) return;
      next = v === 0 ? 0 : undefined;
      if (next == null) {
        const name = group().students.find((x) => String(x.row) === row).name;
        if (!(await confirmChange(name, prev, DISC_DEFAULT, cat))) return;
      }
    } else if (prev != null) {
      const name = group().students.find((x) => String(x.row) === row).name;
      if (!(await confirmChange(name, prev, next, cat))) return;
    }
    const day = dayData(cat, true);
    if (next != null) day[row] = next;
    else delete day[row];
    save();
    if (navigator.vibrate) navigator.vibrate(10);
    const cur = optionOf(cat, disc ? (day[row] ?? DISC_DEFAULT) : day[row]);
    li.classList.remove('cA', 'cF', 'cR', 'cJ', 'cN');
    if (cur) li.classList.add('c' + cur.color);
    li.querySelectorAll('.cb').forEach((b) => {
      const o = OPTIONS[t].find((x) => String(x.v) === b.dataset.v);
      const on = cur && cur.v === o.v;
      b.classList.remove('on', 'cA', 'cF', 'cR', 'cN');
      if (on) b.classList.add('on', 'c' + o.color);
      b.setAttribute('aria-pressed', !!on);
    });
    updateFillRest();
    renderSummary();
  });
  // Examen: se guarda al salir del campo o con Enter.
  $('#compact-list').addEventListener('change', async (e) => {
    const inp = e.target.closest('.exam-in');
    if (!inp || typeOf(tab()) !== 'num') return;
    const li = inp.closest('.crow');
    const row = li.dataset.row;
    const prev = dayData(tab(), false)[row];
    const txt = inp.value.trim().replace(',', '.');
    let next;
    if (txt !== '') {
      next = Number(txt);
      if (!Number.isFinite(next) || next < 0 || next > 100) {
        toast('La calificación va de 0 a 100');
        inp.value = prev ?? '';
        return;
      }
    }
    if (next === prev) { inp.value = prev ?? ''; return; }
    if (prev != null) {
      const name = group().students.find((x) => String(x.row) === row).name;
      if (!(await confirmChange(name, prev, next, tab()))) { inp.value = prev; return; }
    }
    const day = dayData(tab(), true);
    if (next != null) day[row] = next;
    else delete day[row];
    inp.value = next ?? '';
    li.classList.toggle('cN', next != null);
    save();
    renderSummary();
  });
  $('#compact-list').addEventListener('keydown', (e) => {
    if (e.key !== 'Enter' || !e.target.classList.contains('exam-in')) return;
    e.preventDefault();
    const all = [...document.querySelectorAll('#compact-list .crow:not([hidden]) .exam-in')];
    const nxt = all[all.indexOf(e.target) + 1];
    if (nxt) nxt.focus(); else e.target.blur();
  });
  // ---------- Varios trabajos por día ----------
  function renderWorkBar() {
    const bar = $('#work-bar');
    bar.hidden = typeOf(tab()) !== 'trab';
    if (bar.hidden) return;
    const ws = works(false);
    const total = Math.max(ws.length, 1);
    const i = workIdx();
    state.ui.work = i;
    const w = ws[i];
    $('#work-pos').textContent = `${secById(tab()).name} ${i + 1} de ${total}`;
    $('#work-prev').disabled = i === 0;
    $('#work-next').disabled = i >= total - 1;
    // El nombre se refresca al cambiar de trabajo o de día (no mientras se escribe en el mismo trabajo).
    const title = $('#work-title');
    const key = `${state.ui.group}|${state.ui.date}|${i}`;
    if (title.dataset.key !== key || document.activeElement !== title) title.value = w ? w.title : '';
    title.dataset.key = key;
    $('#work-del').hidden = !w || (ws.length <= 1 && !Object.keys(w.vals).length && !w.title);
  }
  function goWork(i) {
    // Siempre se muestra el nombre del trabajo al que se cambia.
    $('#work-title').blur();
    $('#work-title').dataset.key = '';
    state.ui.work = i;
    save();
    refresh();
  }
  $('#work-prev').addEventListener('click', () => goWork(workIdx() - 1));
  $('#work-next').addEventListener('click', () => goWork(workIdx() + 1));
  $('#work-add').addEventListener('click', () => {
    const ws = works(true);
    const last = ws[ws.length - 1];
    // Si el último trabajo sigue vacío no se crea otro: se usa ese.
    if (last && !Object.keys(last.vals).length && !last.title.trim()) {
      goWork(ws.length - 1);
    } else {
      if (!ws.length) ws.push({ title: '', vals: {} }); // el trabajo 1 que se estaba viendo
      ws.push({ title: '', vals: {} });
      goWork(ws.length - 1);
      toast(`Trabajo ${ws.length} creado`);
    }
    $('#work-title').focus();
  });
  $('#work-title').addEventListener('input', (e) => {
    const w = curWork(true);
    w.title = e.target.value;
    save();
  });
  $('#work-title').addEventListener('keydown', (e) => { if (e.key === 'Enter') e.target.blur(); });
  $('#work-del').addEventListener('click', async () => {
    const ws = works(false);
    const i = workIdx();
    const w = ws[i];
    if (!w) return;
    const n = Object.keys(w.vals).length;
    const name = w.title.trim() ? `“${w.title.trim()}”` : `Trabajo ${i + 1}`;
    const ok = await confirmBulk(
      `¿Seguro que quieres eliminar el trabajo ${name}?`,
      n ? `Se borrarán también sus ${n} calificación(es) de este día. No se puede deshacer.` : 'Este trabajo no tiene calificaciones.',
      'Sí, eliminar', true);
    if (!ok) return;
    ws.splice(i, 1);
    if (!ws.length) delete group()[tab()][state.ui.date];
    goWork(Math.max(0, i - 1));
    toast(`${name} eliminado`);
  });

  $('#compact-preview').addEventListener('click', openPreview);
  $('#compact-export').addEventListener('click', () => doExport($('#compact-export')));

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
  function confirmChange(name, from, to, cat = ASIST_ID) {
    return new Promise((resolve) => {
      const d = $('#confirm-change');
      const pill = (sel, v) => {
        const el = $(sel);
        const o = optionOf(cat, v);
        el.className = `change-val ${o ? o.color : 'none'}`;
        el.textContent = o ? o.label : 'Sin marca';
      };
      $('#confirm-name').textContent = name;
      pill('#confirm-from', from);
      pill('#confirm-to', to);
      $('#confirm-ok').textContent = to != null ? 'Sí, cambiar' : 'Sí, quitar';
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
    if (typeOf(tab()) !== 'asist' || view !== 'card' || $('#screen-main').hidden || document.querySelector('dialog[open]')) return;
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
    const cat = tab();
    const t = typeOf(cat);
    if (t === 'num') {
      const day = dayData(cat, false);
      const vals = g.students.map((s) => day[s.row]).filter((v) => v != null);
      const avg = vals.length ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 10) / 10 : null;
      $('#summary').innerHTML =
        `<span>${fmtDate(state.ui.date)}</span><span>${vals.length}/${g.students.length} capturados</span>` +
        (avg != null ? `<span>Promedio: ${avg}</span>` : '');
      return;
    }
    if (t === 'disc') {
      const day = dayData(cat, false);
      const zeros = g.students.filter((s) => day[s.row] === 0).length;
      $('#summary').innerHTML =
        `<span>${fmtDate(state.ui.date)}</span><span class="F">Con 0: ${zeros}</span><span>Los demás: 10</span>`;
      return;
    }
    if (t !== 'asist') {
      const day = dayData(cat, false);
      const vals = g.students.map((s) => day[s.row]).filter((v) => v != null);
      $('#summary').innerHTML =
        `<span>${vals.length}/${g.students.length} capturados</span>` +
        OPTIONS[t].map((o) => `<span class="${o.color}">${o.btn}: ${vals.filter((v) => v === o.v).length}</span>`).join('');
      return;
    }
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
    sectionsOf(g).forEach((x) => Object.values(g[x.id] || {}).forEach((day) => {
      if (Array.isArray(day)) day.forEach((w) => delete w.vals[s.row]); // trabajos
      else delete day[s.row];
    }));
    save();
    if (cardIdx > g.students.length) cardIdx = g.students.length;
    $('#remove-student').close();
    refresh();
    toast(`${s.name} quitado de ${gname(g)}`);
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
    state.ui.work = 0; // al cambiar de día o grupo se empieza en el primer trabajo
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

  function openAddStudent() {
    const g = group();
    const nums = g.students.map((s) => Number(s.no)).filter(Number.isFinite);
    $('#add-group').textContent = `Grupo ${gname(g)} · ${g.students.length} alumnos`;
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
    state.ui.tab = ASIST_ID;
    save();
    view = 'card';
    cardIdx = g.students.length - 1;
    refresh();
    toast(`${name} agregado a ${gname(g)}`);
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

    const dates = groupDates(g);
    const layout = sheetLayout(g, g.grading);
    const byCol = new Map(layout.map((x) => [x.col, x]));
    const lastCol = Math.max(g.nameCol, ...layout.map((x) => x.col));
    const isCur = (c) => byCol.has(c) && byCol.get(c).date === state.ui.date;
    const colLetter = (c) => { let s = ''; for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };
    // Se toma de lo capturado (no del número), porque A, R y J pueden tener el mismo valor.
    const colorOf = (x, r) => {
      if (x.kind === 'disc-count') return discCount(x.sec, r) > 0 ? 'F' : '';
      if (x.kind !== 'day' && x.kind !== 'trab') return '';
      const v = x.kind === 'trab' ? (worksOf(g, x.date, x.sec.id)[x.work].vals || {})[r] : ((g[x.sec.id] || {})[x.date] || {})[r];
      if (v == null) return '';
      const o = optionOf(x.sec.id, v);
      return o ? o.color : '';
    };
    const discCount = (sec, r) => Object.values(g[sec.id] || {}).filter((d) => d[r] === 0).length;
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
        const x = byCol.get(c);
        if (x) {
          cls.push(isHdr ? 'date-h' : 'v');
          if (!isHdr) cls.push(colorOf(x, r));
          if (x.first) cls.push('day-start');
          if (x.last) cls.push('block-end');
          if (x.kind === 'day' && x.sec.type === 'sino') cls.push('lect-col');
          if (x.kind === 'grade' && x.key === 'final') cls.push('final');
          if (x.kind === 'disc-dates') cls.push('disc-dates');
          if (isCur(c)) cls.push('cur');
        }
        // El encabezado lleva fecha y apartado en dos renglones, como en el Excel.
        const shown = isHdr && x ? escapeHtml(ListaExcel.cellText(v)).replace(/\n/g, '<br>') : escapeHtml(t);
        html += `<td class="${cls.join(' ')}">${shown}</td>`;
      }
      html += '</tr>';
    });
    html += '</tbody>';
    $('#preview-table').innerHTML = html;

    $('#preview-title').textContent = `Vista previa · ${gname(g)}`;
    $('#preview-sub').textContent = layout.length
      ? `${g.students.length} alumnos · ${dates.length} día(s) · Por fecha: ${secs().filter((x) => x.type === 'asist' || x.type === 'sino').map((x) => x.name).join(' + ')}; las demás secciones aparte a la derecha`
      : 'Aún no hay registros en este grupo.';
    $('#preview-legend').innerHTML =
      STATUSES.map((k) => `<span class="summary-${k}">${escapeHtml(String(state.codes[k]))} = ${LABELS[k]}</span>`).join('') +
      (secs().some((x) => x.type === 'sino') ? READ.map((k) => `<span class="summary-${k === 'S' ? 'A' : 'F'}">${escapeHtml(String(state.codes[k]))} = ${k === 'S' ? 'Sí' : 'No'}</span>`).join('') : '');

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
      await deliverFile(out, state.fileName.replace(/\.xlsx$/i, '') + ' - asistencia.xlsx');
      toast('Excel generado');
    } catch (err) {
      console.error(err);
      toast('Error al generar el Excel');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Exportar Excel'; }
    }
  }

  async function deliverFile(out, name) {
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
  }

  // ---------- Cuadernillo (PDF por grupo) ----------
  // El PDF es grande para localStorage: se guarda en IndexedDB, una entrada por grupo (nombre de la hoja),
  // así sigue ahí aunque se vuelva a subir el Excel. Se muestra con PDF.js para que se vea igual en Android y iPhone.
  const bookDB = (() => {
    let p;
    const open = () => (p = p || new Promise((res, rej) => {
      const r = indexedDB.open('paseDeLista', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('cuadernillos');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }));
    const run = (mode, fn) => open().then((db) => new Promise((res, rej) => {
      const tx = db.transaction('cuadernillos', mode);
      const req = fn(tx.objectStore('cuadernillos'));
      tx.oncomplete = () => res(req && req.result);
      tx.onerror = () => rej(tx.error);
      tx.onabort = () => rej(tx.error);
    }));
    return {
      get: (k) => run('readonly', (s) => s.get(k)),
      put: (k, v) => run('readwrite', (s) => s.put(v, k)),
      del: (k) => run('readwrite', (s) => s.delete(k)),
    };
  })();
  let pdfjsReady = null;
  const loadPdfJs = () => (pdfjsReady = pdfjsReady || new Promise((res, rej) => {
    const sc = document.createElement('script');
    sc.src = 'vendor/pdfjs/pdf.min.js';
    sc.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js'; res(window.pdfjsLib); };
    sc.onerror = () => { pdfjsReady = null; rej(new Error('No se pudo cargar el visor de PDF')); };
    document.head.appendChild(sc);
  }));

  const ZOOMS = [1, 1.25, 1.5, 2, 2.5, 3];
  const book = { key: null, doc: null, rec: null, zoom: 0, page: 1, observer: null, saveTimer: null, gen: 0 };
  const fmtSize = (n) => (n > 1048576 ? `${(n / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`);

  function bookShow(mode) {
    $('#book-empty').hidden = mode !== 'empty';
    $('#book-bar').hidden = mode !== 'doc';
    $('#book-pages').hidden = mode !== 'doc';
    $('#book-actions').hidden = mode !== 'doc';
    $('#book-loading').hidden = mode !== 'loading';
  }
  function closeBookDoc() {
    book.gen++;
    if (book.observer) book.observer.disconnect();
    book.observer = null;
    if (book.doc) book.doc.destroy();
    book.doc = null;
    $('#book-pages').innerHTML = '';
  }
  async function openBook() {
    const g = group();
    book.key = g.sheet;
    closeBookDoc();
    $('#book-sub').textContent = gname(g);
    $('#book').showModal();
    bookShow('loading');
    let rec;
    try { rec = await bookDB.get(book.key); } catch (err) { console.error(err); }
    if (!rec) {
      $('#book-empty-text').textContent = `Aún no hay cuadernillo para ${gname(g)}. Sube el PDF una sola vez y aquí lo verás siempre, también sin internet.`;
      bookShow('empty');
      return;
    }
    showBookDoc(rec);
  }
  async function showBookDoc(rec) {
    const gen = ++book.gen;
    bookShow('loading');
    book.rec = rec;
    try {
      const pdfjsLib = await loadPdfJs();
      // PDF.js se queda con el buffer que recibe: se le pasa una copia.
      const doc = await pdfjsLib.getDocument({ data: new Uint8Array(rec.data.slice(0)) }).promise;
      if (gen !== book.gen) { doc.destroy(); return; }
      book.doc = doc;
    } catch (err) {
      console.error(err);
      toast('No se pudo abrir el PDF');
      $('#book-empty-text').textContent = 'El PDF guardado no se pudo abrir. Sube otro.';
      bookShow('empty');
      return;
    }
    $('#book-sub').textContent = `${gname(group())} · ${rec.name} · ${fmtSize(rec.size)}`;
    $('#book-total').textContent = book.doc.numPages;
    bookShow('doc');
    await layoutBook(Math.min(rec.page || 1, book.doc.numPages));
  }
  // Coloca un espacio por página (con su proporción) y dibuja solo las que están cerca de la vista.
  async function layoutBook(goTo) {
    const gen = book.gen;
    const box = $('#book-pages');
    if (book.observer) book.observer.disconnect();
    const first = await book.doc.getPage(1);
    if (gen !== book.gen) return;
    const vp = first.getViewport({ scale: 1 });
    const width = Math.floor((box.clientWidth - 16) * ZOOMS[book.zoom]);
    let html = '';
    for (let n = 1; n <= book.doc.numPages; n++) {
      html += `<div class="book-sheet" data-n="${n}" style="width:${width}px;height:${Math.round((width * vp.height) / vp.width)}px"><span>${n}</span></div>`;
    }
    box.innerHTML = html;
    $('#book-zoom-val').textContent = `${Math.round(ZOOMS[book.zoom] * 100)} %`;
    $('#book-fpage').textContent = `${book.page} / ${book.doc.numPages}`;
    book.observer = new IntersectionObserver((entries) => entries.forEach((e) => {
      if (e.isIntersecting) drawPage(e.target, gen);
      else { const c = e.target.querySelector('canvas'); if (c) { c.width = 0; c.remove(); } } // libera memoria
    }), { root: box, rootMargin: '800px 0px' });
    box.querySelectorAll('.book-sheet').forEach((el) => book.observer.observe(el));
    goPage(goTo);
  }
  async function drawPage(el, gen) {
    if (el.querySelector('canvas') || el.dataset.busy) return;
    el.dataset.busy = '1';
    try {
      const page = await book.doc.getPage(Number(el.dataset.n));
      if (gen !== book.gen) return;
      const css = el.clientWidth;
      const vp1 = page.getViewport({ scale: 1 });
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const vp = page.getViewport({ scale: (css / vp1.width) * dpr });
      const c = document.createElement('canvas');
      c.width = Math.floor(vp.width);
      c.height = Math.floor(vp.height);
      el.style.height = `${Math.round(vp.height / dpr)}px`; // por si la página tiene otro tamaño
      await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      if (gen !== book.gen || el.querySelector('canvas')) return;
      el.appendChild(c);
    } catch (err) {
      if (gen === book.gen) console.error(err);
    } finally {
      delete el.dataset.busy;
    }
  }
  function goPage(n) {
    const box = $('#book-pages');
    const el = box.querySelector(`.book-sheet[data-n="${n}"]`);
    if (!el) return;
    box.scrollTop = el.offsetTop - box.offsetTop - 8;
    setBookPage(n);
  }
  function setBookPage(n) {
    book.page = n;
    if (document.activeElement !== $('#book-page')) $('#book-page').value = n;
    if (book.doc) $('#book-fpage').textContent = `${n} / ${book.doc.numPages}`;
    // Recuerda la página para abrir ahí la próxima vez.
    clearTimeout(book.saveTimer);
    const rec = book.rec;
    const key = book.key;
    book.saveTimer = setTimeout(() => {
      if (!rec || rec.page === n) return;
      rec.page = n;
      bookDB.put(key, rec).catch(() => {});
    }, 800);
  }
  $('#book-pages').addEventListener('scroll', () => {
    const box = $('#book-pages');
    const mid = box.scrollTop + box.clientHeight / 3 + box.offsetTop;
    let n = 1;
    for (const el of box.children) { if (el.offsetTop <= mid) n = Number(el.dataset.n); else break; }
    if (n !== book.page) setBookPage(n);
  }, { passive: true });
  $('#book-page').addEventListener('change', (e) => {
    const n = Math.round(Number(e.target.value));
    if (!book.doc || !(n >= 1 && n <= book.doc.numPages)) { e.target.value = book.page; toast(`Escribe una página del 1 al ${book.doc ? book.doc.numPages : 1}`); return; }
    e.target.blur();
    goPage(n);
  });
  $('#book-page').addEventListener('keydown', (e) => { if (e.key === 'Enter') e.target.dispatchEvent(new Event('change')); });
  const zoomBook = (d) => {
    const z = Math.max(0, Math.min(ZOOMS.length - 1, book.zoom + d));
    if (z === book.zoom || !book.doc) return;
    book.zoom = z;
    book.gen++;
    layoutBook(book.page);
  };
  $('#book-zoom-in').addEventListener('click', () => zoomBook(1));
  $('#book-zoom-out').addEventListener('click', () => zoomBook(-1));
  // Pantalla completa: solo el PDF, con controles flotantes mínimos.
  // En Android usa la pantalla completa del sistema; en iPhone (no la permite) ocupa toda la ventana de la app.
  const bookDlg = $('#book');
  const isReading = () => bookDlg.classList.contains('reading');
  async function setReading(on) {
    if (on === isReading() || !book.doc) return;
    const page = book.page;
    bookDlg.classList.toggle('reading', on);
    try {
      if (on && bookDlg.requestFullscreen && !document.fullscreenElement) await bookDlg.requestFullscreen({ navigationUI: 'hide' });
      if (!on && document.fullscreenElement) await document.exitFullscreen();
    } catch (err) { /* sin pantalla completa del sistema: basta con ocultar lo demás */ }
    await new Promise((r) => requestAnimationFrame(() => r()));
    book.gen++;
    layoutBook(page); // el ancho cambió: se vuelven a acomodar las páginas en la misma página
  }
  $('#book-full').addEventListener('click', () => setReading(true));
  $('#book-exit').addEventListener('click', () => setReading(false));
  $('#book-fzoom-in').addEventListener('click', () => zoomBook(1));
  $('#book-fzoom-out').addEventListener('click', () => zoomBook(-1));
  // Botón "atrás"/Esc del sistema: primero sale de pantalla completa, no cierra el cuadernillo.
  document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && isReading()) setReading(false); });
  bookDlg.addEventListener('cancel', (e) => { if (isReading()) { e.preventDefault(); setReading(false); } });
  $('#book-upload').addEventListener('click', () => $('#book-file').click());
  $('#book-change').addEventListener('click', () => $('#book-file').click());
  $('#book-file').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') { toast('Elige un archivo PDF'); return; }
    const rec = { name: file.name, size: file.size, data: await file.arrayBuffer(), page: 1, added: Date.now() };
    if (String.fromCharCode(...new Uint8Array(rec.data.slice(0, 5))) !== '%PDF-') { toast('Ese archivo no es un PDF válido'); return; }
    try {
      if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
      await bookDB.put(book.key, rec);
    } catch (err) {
      console.error(err);
      toast('No hay espacio suficiente en el dispositivo para guardar el PDF');
      return;
    }
    closeBookDoc();
    book.zoom = 0;
    toast(`Cuadernillo guardado para ${gname(group())}`);
    showBookDoc(rec);
  });
  $('#book-remove').addEventListener('click', async () => {
    const key = book.key;
    const name = book.rec ? book.rec.name : 'el cuadernillo';
    if (!(await confirmBulk(`¿Quitar el cuadernillo de ${gname(group())}?`, `Se borra “${name}” de este dispositivo. Puedes volver a subirlo cuando quieras.`, 'Sí, quitar', true))) return;
    await bookDB.del(key).catch(() => {});
    closeBookDoc();
    book.rec = null;
    $('#book-sub').textContent = gname(group());
    $('#book-empty-text').textContent = `Aún no hay cuadernillo para ${gname(group())}. Sube el PDF una sola vez y aquí lo verás siempre, también sin internet.`;
    bookShow('empty');
    toast('Cuadernillo quitado');
  });
  $('#book').addEventListener('close', () => { $('#book').classList.remove('reading'); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); clearTimeout(book.saveTimer); if (book.rec && book.rec.page !== book.page && book.doc) { book.rec.page = book.page; bookDB.put(book.key, book.rec).catch(() => {}); } closeBookDoc(); });

  // ---------- Materia del grupo ----------
  // Para maestros que tienen el mismo grupo dos veces en el Excel (una hoja por materia).
  function openSubject() {
    const g = group();
    $('#subject-group').textContent = `Hoja del Excel: ${g.sheet} · ${g.students.length} alumnos`;
    $('#subject-input').value = g.subject || '';
    $('#subject-clear').hidden = !g.subject;
    $('#subject').showModal();
    setTimeout(() => $('#subject-input').focus(), 50);
  }
  function setSubject(v) {
    const g = group();
    g.subject = v.replace(/\s+/g, ' ').trim().slice(0, 40);
    if (!g.subject) delete g.subject;
    save();
    $('#subject').close();
    refresh();
    toast(g.subject ? `${g.sheet} ahora dice “${g.subject}”` : `${g.sheet} quedó sin materia`);
  }
  $('#subject-form').addEventListener('submit', (e) => { e.preventDefault(); setSubject($('#subject-input').value); });
  $('#subject-clear').addEventListener('click', () => setSubject(''));

  // ---------- Resumen del grupo (semáforo) ----------
  let grFilter = 'all';
  const ST_LABEL = { rojo: 'En riesgo', amarillo: 'Observación', verde: 'Bien', none: 'Sin datos' };
  function openGroupReport() {
    grFilter = 'all';
    renderGroupReport();
    $('#group-report').showModal();
    $('#gr-list').scrollTop = 0;
  }
  function renderGroupReport() {
    const g = group();
    const r = groupReport(g, g.grading);
    const st = r.stats;
    $('#gr-sub').textContent = `${gname(g)}${r.graded ? ` · calificación sobre ${r.escala}` : ''}`;
    const tile = (v, label, cls = '') => `<div class="gr-tile ${cls}"><b>${v}</b><span>${label}</span></div>`;
    $('#gr-tiles').innerHTML =
      tile(st.alumnos, 'Alumnos') +
      tile(st.asistencia == null ? '—' : `${round1(st.asistencia)} %`, 'Asistencia promedio') +
      tile(st.promedio == null ? '—' : st.promedio, r.graded ? 'Promedio general' : 'Promedio (falta definir %)') +
      tile(st.rojo, 'Alumnos en riesgo', st.rojo ? 'rojo' : '') +
      tile(st.faltas, 'Faltas en total');
    const chip = (k, n, label) => `<button class="gr-chip ${k} ${grFilter === k ? 'on' : ''}" data-f="${k}">${label} <b>${n}</b></button>`;
    $('#gr-chips').innerHTML = chip('all', st.alumnos, 'Todos') + chip('rojo', st.rojo, 'Rojo') + chip('amarillo', st.amarillo, 'Amarillo') + chip('verde', st.verde, 'Verde') +
      (st.none ? chip('none', st.none, 'Sin datos') : '');
    $('#gr-criteria').innerHTML = r.criteria.map((t) => `<li>${escapeHtml(t)}</li>`).join('');
    // Primero los que más atención necesitan.
    const order = { rojo: 0, amarillo: 1, verde: 2, none: 3 };
    const rows = r.rows.filter((x) => grFilter === 'all' || x.status === grFilter)
      .sort((a, b) => order[a.status] - order[b.status] || (a.final ?? 99) - (b.final ?? 99) || (a.att ?? 101) - (b.att ?? 101));
    $('#gr-list').innerHTML = rows.length
      ? rows.map((x) => `<li data-row="${x.student.row}" class="${x.status}">` +
          `<span class="gr-dot" title="${ST_LABEL[x.status]}"></span>` +
          `<span class="gr-name"><b>${escapeHtml(x.student.name)}</b><small>${x.reasons.length ? escapeHtml(x.reasons.join(' · ')) : ST_LABEL[x.status]}</small></span>` +
          `<span class="gr-num"><b>${x.final ?? '—'}</b><small>calif.</small></span>` +
          `<span class="gr-num"><b>${x.att == null ? '—' : `${Math.round(x.att)}%`}</b><small>asist.</small></span>` +
          `<span class="gr-num"><b>${x.counts.F}</b><small>faltas</small></span></li>`).join('')
      : '<li class="empty">No hay alumnos en este color.</li>';
  }
  $('#gr-chips').addEventListener('click', (e) => {
    const b = e.target.closest('[data-f]');
    if (!b) return;
    grFilter = b.dataset.f;
    renderGroupReport();
  });
  // Tocar un alumno abre su resumen individual.
  $('#gr-list').addEventListener('click', (e) => {
    const li = e.target.closest('li[data-row]');
    if (!li) return;
    $('#group-report').close();
    openStudent();
    studentRow = Number(li.dataset.row);
    renderStudent();
  });
  // Vista previa: genera el mismo Excel que se exporta y lo dibuja celda por celda (colores, negritas, celdas unidas).
  $('#gr-preview').addEventListener('click', async () => {
    const g = group();
    const btn = $('#gr-preview');
    btn.disabled = true;
    try {
      const buf = await exportGroupReport(ExcelJS, g, g.grading);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      $('#gr-sheet-table').innerHTML = sheetToHtml(wb.worksheets[0]);
      $('#gr-sheet-sub').textContent = `Resumen ${gname(g)}.xlsx`;
      $('#gr-sheet').showModal();
      $('#gr-sheet .preview-scroll').scrollTop = 0;
    } catch (err) {
      console.error(err);
      toast('No se pudo generar la vista previa');
    } finally {
      btn.disabled = false;
    }
  });
  $('#gr-sheet-export').addEventListener('click', () => $('#gr-export').click());
  function sheetToHtml(ws) {
    const nCols = ws.columnCount;
    const nRows = ws.rowCount;
    const colName = (c) => { let s = ''; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
    // Celdas unidas: la primera lleva colspan/rowspan y las demás no se dibujan.
    const span = {};
    const skip = new Set();
    (ws.model.merges || []).forEach((rg) => {
      const [a, b] = rg.split(':').map((ref) => ws.getCell(ref));
      span[`${a.row},${a.col}`] = { cs: b.col - a.col + 1, rs: b.row - a.row + 1 };
      for (let r = a.row; r <= b.row; r++) for (let c = a.col; c <= b.col; c++) if (r !== a.row || c !== a.col) skip.add(`${r},${c}`);
    });
    const argb = (x) => (x && x.argb ? `#${x.argb.slice(-6)}` : null);
    const text = (v) => {
      if (v == null) return '';
      if (typeof v === 'object' && v.richText) return v.richText.map((t) => t.text).join('');
      return String(v);
    };
    let html = '<thead><tr><th class="rownum"></th>';
    for (let c = 1; c <= nCols; c++) {
      // Ancho proporcional al de Excel, más compacto para que en celular se vean los números junto a los nombres.
      const w = Math.min(Math.round((ws.getColumn(c).width || 10) * 5.5), 190);
      html += `<th style="min-width:${w}px;max-width:${w}px">${colName(c)}</th>`;
    }
    html += '</tr></thead><tbody>';
    for (let r = 1; r <= nRows; r++) {
      html += `<tr><td class="rownum">${r}</td>`;
      for (let c = 1; c <= nCols; c++) {
        const key = `${r},${c}`;
        if (skip.has(key)) continue;
        const cell = ws.getCell(r, c);
        const st = [];
        const bg = cell.fill && cell.fill.type === 'pattern' ? argb(cell.fill.fgColor) : null;
        if (bg) st.push(`background:${bg}`);
        const f = cell.font || {};
        if (f.bold) st.push('font-weight:700');
        if (f.color && argb(f.color)) st.push(`color:${argb(f.color)}`);
        if (f.size && f.size > 12) st.push(`font-size:${f.size}px`);
        const al = (cell.alignment || {}).horizontal;
        if (al === 'center') st.push('text-align:center');
        if (cell.border && cell.border.top) st.push('border:1px solid #999');
        const sp = span[key];
        html += `<td${sp ? ` colspan="${sp.cs}" rowspan="${sp.rs}"` : ''}${st.length ? ` style="${st.join(';')}"` : ''}>${escapeHtml(text(cell.value)).replace(/\n/g, '<br>')}</td>`;
      }
      html += '</tr>';
    }
    return html + '</tbody>';
  }
  $('#gr-export').addEventListener('click', async () => {
    const g = group();
    const btn = $('#gr-export');
    btn.disabled = true;
    btn.textContent = 'Generando…';
    try {
      const out = await exportGroupReport(ExcelJS, g, g.grading);
      const safe = (t) => String(t).replace(/[\\/:*?"<>|]+/g, ' ').trim();
      await deliverFile(out, `Resumen ${safe(gname(g))}.xlsx`);
      toast('Resumen del grupo generado');
    } catch (err) {
      console.error(err);
      toast('Error al generar el Excel');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Exportar Excel';
    }
  });

  // ---------- Resumen por alumno ----------
  // Para quejas o dudas del alumno o de los padres: todo lo capturado de un alumno y su Excel individual.
  let studentRow = null;
  function openStudent() {
    studentRow = null;
    $('#student-search').value = '';
    renderStudent();
    $('#student').showModal();
  }
  function renderStudent() {
    const g = group();
    const picked = g.students.find((s) => s.row === studentRow);
    $('#student-pick').hidden = !!picked;
    $('#student-report').hidden = !picked;
    $('#student-back').hidden = !picked;
    $('#student-export').hidden = !picked;
    $('#student-close').hidden = !!picked;
    if (!picked) {
      $('#student-sub').textContent = `${gname(g)} · elige un alumno`;
      const q = normalize($('#student-search').value);
      const shown = g.students.filter((s) => !q || normalize(s.name).includes(q));
      $('#student-list').innerHTML = shown.length
        ? shown.map((s) => {
            const f = finalGrade(g, s.row, g.grading).final;
            return `<li data-row="${s.row}"><span class="no">${escapeHtml(s.no)}</span><span class="nm">${escapeHtml(s.name)}</span>${f != null ? `<span class="cal">${f}</span>` : ''}</li>`;
          }).join('')
        : '<li class="empty">Ningún alumno coincide con la búsqueda.</li>';
      return;
    }
    const r = studentReport(g, studentRow, g.grading);
    $('#student-sub').textContent = gname(g);
    const pct = (v) => (v == null ? 'Sin datos' : `${round1(v)} %`);
    $('#student-report').innerHTML =
      `<div class="sr-head"><div><b>${escapeHtml(picked.name)}</b><small>No. ${escapeHtml(picked.no)} · ${escapeHtml(gname(g))}</small></div>` +
      `<div class="sr-final">${r.final == null ? '—' : r.final}<small>${r.graded ? `sobre ${r.grading.escala}` : 'sin porcentajes'}</small></div></div>` +
      '<table class="sheet sr-parts"><thead><tr><th>Sección</th><th>Vale</th><th>Resultado</th></tr></thead><tbody>' +
      r.sections.map((p) => `<tr><td><b>${escapeHtml(p.name)}</b><small>${escapeHtml(r.detail[p.id])}</small></td>` +
        `<td class="v">${weightOf(r.grading, p) ? `${weightOf(r.grading, p)} %` : '—'}</td><td class="v">${pct(r.parts[p.id])}</td></tr>`).join('') +
      '</tbody></table>' +
      '<h3 class="sr-title">Detalle por día</h3>' +
      (r.days.length
        ? '<div class="preview-scroll"><table class="sheet sr-days"><thead><tr><th>Fecha</th>' +
          r.sections.map((p) => `<th class="wrap">${escapeHtml(p.name)}</th>`).join('') + '</tr></thead><tbody>' +
          r.days.map((d) => `<tr><td class="nowrap">${d.label.replace(/\/\d\d(\d\d)$/, '/$1')}</td>` +
            r.sections.map((p) => {
              const c = d.cells[p.id];
              if (!c) return '<td></td>';
              if (p.type === 'trab') return `<td>${c.list.map((w) => `<span class="${Number(w.v) === 0 ? 'F' : ''}">${escapeHtml(w.title)}: ${w.v}</span>`).join('<br>')}</td>`;
              return `<td class="v ${c.tone || ''}">${escapeHtml(String(c.short))}</td>`;
            }).join('') + '</tr>').join('') +
          '</tbody></table></div>'
        : '<p class="hint">Aún no hay datos capturados de este alumno.</p>');
  }
  $('#student-search').addEventListener('input', renderStudent);
  $('#student-list').addEventListener('click', (e) => {
    const li = e.target.closest('li[data-row]');
    if (!li) return;
    studentRow = Number(li.dataset.row);
    renderStudent();
    $('#student-report').scrollTop = 0;
  });
  $('#student-back').addEventListener('click', () => { studentRow = null; renderStudent(); });
  $('#student-export').addEventListener('click', async () => {
    const g = group();
    const s = g.students.find((x) => x.row === studentRow);
    if (!s) return;
    const btn = $('#student-export');
    btn.disabled = true;
    btn.textContent = 'Generando…';
    try {
      const out = await exportStudent(ExcelJS, g, s.row, g.grading);
      const safe = (t) => String(t).replace(/[\\/:*?"<>|]+/g, ' ').trim();
      await deliverFile(out, `${safe(s.name)} - ${safe(gname(g))} - resumen.xlsx`);
      toast('Excel del alumno generado');
    } catch (err) {
      console.error(err);
      toast('Error al generar el Excel');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Exportar Excel';
    }
  });

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
    if (action === 'remove-student') openRemoveStudent();
    if (action === 'grades') openGrades();
    if (action === 'student') openStudent();
    if (action === 'group-report') openGroupReport();
    if (action === 'book') openBook();
    if (action === 'subject') openSubject();
    if (action === 'settings') openSettings();
    if (action === 'new-file') {
      if (confirm('Se reemplazará la lista actual y su asistencia guardada. Exporta antes si la necesitas. ¿Continuar?')) {
        $('#file-input').click();
      }
    }
  });

  // ---------- Secciones y calificaciones ----------
  // Cada grupo tiene sus secciones (nombre + plantilla) y cuánto vale cada una. Todo se edita en un
  // borrador: la tabla se actualiza al instante, pero solo se guarda con "Guardar".
  let gradeDraft = null; // { sections, grading }
  const snapshot = (g) => JSON.parse(JSON.stringify({ sections: sectionsOf(g), grading: g.grading }));
  const sameDraft = () => JSON.stringify(gradeDraft) === JSON.stringify(snapshot(group()));
  // Grupo "de prueba" con el borrador, para calcular la tabla sin tocar el grupo real.
  const draftGroup = () => Object.assign({}, group(), { sections: gradeDraft.sections });
  const typeLabel = (t) => TYPES[t].label;
  const draftTotal = () => round1(gradeDraft.sections.reduce((n, x) => n + weightOf(gradeDraft.grading, x), 0));
  // Cuántos datos tiene una sección (para avisar antes de borrarla).
  function secDataCount(g, sec) {
    const d = g[sec.id] || {};
    return Object.keys(d).filter((k) => (Array.isArray(d[k]) ? d[k].some((w) => Object.keys(w.vals || {}).length) : Object.keys(d[k]).length)).length;
  }

  function openGrades() {
    gradeDraft = snapshot(group());
    $('#grades-scope').textContent = `Secciones de ${gname(group())} y cuánto vale cada una. Cada grupo tiene las suyas.`;
    const f = $('#grades-form');
    f.elements.retardo.value = String(gradeDraft.grading.retardo);
    f.elements.escala.value = String(gradeDraft.grading.escala);
    renderSecList();
    renderGrades();
    $('#grades').showModal();
  }
  function renderSecList() {
    const g = group();
    $('#sec-list').innerHTML = gradeDraft.sections.map((x, i) => {
      const isNew = !sectionsOf(g).some((o) => o.id === x.id);
      const w = gradeDraft.grading.w[x.id];
      const typeBox = isNew
        ? `<select data-f="type" aria-label="Plantilla">${Object.keys(TYPES).filter((t) => t !== 'asist').map((t) => `<option value="${t}" ${t === x.type ? 'selected' : ''}>${escapeHtml(typeLabel(t))}</option>`).join('')}</select>`
        : `<span class="sec-type">${escapeHtml(typeLabel(x.type))}</span>`;
      return `<div class="sec-row ${isNew ? 'new' : ''}" data-i="${i}">
        <div class="sec-top">
          <input class="sec-name" data-f="name" maxlength="30" value="${escapeHtml(x.name)}" placeholder="Nombre de la sección" aria-label="Nombre de la sección">
          <span class="sec-w"><input data-f="w" inputmode="decimal" maxlength="5" value="${w || ''}" aria-label="Porcentaje de ${escapeHtml(x.name)}">%</span>
          ${x.type === 'asist' ? '<span class="sec-del-ph"></span>' : `<button type="button" class="icon-btn sec-del" data-del="${i}" aria-label="Quitar ${escapeHtml(x.name)}">✕</button>`}
        </div>
        <div class="sec-sub">${typeBox}${x.type === 'sino' ? `<label class="sec-min">100 % con <input data-f="min" class="mini-num" inputmode="numeric" maxlength="3" value="${secMin(x)}"> o más “Sí”</label>` : ''}</div>
      </div>`;
    }).join('');
  }
  $('#sec-list').addEventListener('input', (e) => {
    const el = e.target.closest('[data-f]');
    const row = e.target.closest('.sec-row');
    if (!el || !row) return;
    const x = gradeDraft.sections[Number(row.dataset.i)];
    const f = el.dataset.f;
    if (f === 'name') x.name = el.value.replace(/[:\n]/g, ' ').replace(/\s+/g, ' ').trimStart();
    if (f === 'w') {
      const n = Number(el.value.replace(',', '.'));
      if (Number.isFinite(n) && n > 0) gradeDraft.grading.w[x.id] = n; else delete gradeDraft.grading.w[x.id];
    }
    if (f === 'min') { const n = Math.round(Number(el.value)); if (Number.isFinite(n) && n >= 1) x.min = n; }
    if (f === 'type') {
      x.type = el.value;
      if (x.type === 'sino') x.min = x.min || 4; else delete x.min;
      renderSecList();
    }
    renderGrades();
  });
  $('#sec-list').addEventListener('click', (e) => {
    const b = e.target.closest('[data-del]');
    if (!b) return;
    const x = gradeDraft.sections[Number(b.dataset.del)];
    gradeDraft.sections.splice(Number(b.dataset.del), 1);
    delete gradeDraft.grading.w[x.id];
    renderSecList();
    renderGrades();
    const n = secDataCount(group(), x);
    toast(n ? `“${x.name}” se quitará al guardar (tiene datos de ${n} día(s))` : `“${x.name || 'Sección'}” quitada`);
  });
  $('#sec-add').addEventListener('click', () => {
    gradeDraft.sections.push({ id: newSectionId(), name: '', type: 'num' });
    renderSecList();
    renderGrades();
    const rows = document.querySelectorAll('#sec-list .sec-row');
    const last = rows[rows.length - 1];
    last.scrollIntoView({ block: 'nearest' });
    last.querySelector('.sec-name').focus();
  });
  $('#grades-form').addEventListener('change', (e) => {
    if (e.target.name === 'retardo') gradeDraft.grading.retardo = Number(e.target.value);
    if (e.target.name === 'escala') gradeDraft.grading.escala = Number(e.target.value);
    if (e.target.name === 'retardo' || e.target.name === 'escala') renderGrades();
  });
  $('#grades-save').addEventListener('click', async () => {
    const g = group();
    const list = gradeDraft.sections;
    list.forEach((x) => (x.name = x.name.trim()));
    const empty = list.find((x) => !x.name);
    if (empty) { toast('Ponle nombre a todas las secciones'); return; }
    const seen = new Set();
    const dup = list.find((x) => { const k = normalize(x.name); if (seen.has(k)) return true; seen.add(k); return false; });
    if (dup) { toast(`Hay dos secciones llamadas “${dup.name}”`); return; }
    const total = draftTotal();
    if (total !== 100 && total !== 0) { toast(`Los porcentajes suman ${total} %; deben sumar 100 %`); return; }
    // Secciones quitadas que tenían datos: se avisa antes de borrarlas.
    const gone = sectionsOf(g).filter((o) => !list.some((x) => x.id === o.id));
    const withData = gone.map((o) => [o, secDataCount(g, o)]).filter(([, n]) => n);
    if (withData.length && !(await confirmBulk(
      `¿Quitar ${withData.map(([o]) => `“${o.name}”`).join(', ')}?`,
      `Se borrarán sus datos capturados (${withData.map(([o, n]) => `${o.name}: ${n} día(s)`).join(', ')}). No se puede deshacer.`,
      'Sí, quitar', true))) return;
    gone.forEach((o) => delete g[o.id]);
    g.sections = JSON.parse(JSON.stringify(list));
    g.grading = JSON.parse(JSON.stringify(gradeDraft.grading));
    Object.keys(g.grading.w).forEach((id) => { if (!g.sections.some((x) => x.id === id)) delete g.grading.w[id]; });
    normalizeGroup(g);
    gradeDraft = snapshot(g);
    if (!secById(state.ui.tab)) state.ui.tab = ASIST_ID;
    save();
    renderSecList();
    renderGrades();
    refresh();
    toast(`Secciones y porcentajes de ${gname(g)} guardados`);
  });
  // Al cerrar sin guardar se descartan los cambios (avisando).
  $('#grades').addEventListener('close', () => {
    if (gradeDraft && !sameDraft()) toast('No se guardaron los cambios');
    gradeDraft = null;
  });
  function renderGrades() {
    const g = group();
    const dg = draftGroup();
    const gr = gradeDraft.grading;
    const total = draftTotal();
    const saved = sameDraft();
    const tot = $('#weights-total');
    const okTotal = total === 100 || total === 0;
    tot.textContent = total === 0
      ? (saved ? 'Sin porcentajes · Guardado ✓' : 'Sin porcentajes (no se calcula calificación) · Falta guardar')
      : total !== 100
        ? `Total: ${total} % — debe sumar 100 %`
        : saved ? 'Total: 100 % · Guardado ✓' : 'Total: 100 % · Falta guardar';
    tot.className = 'weights-total ' + (!okTotal ? 'bad' : saved ? 'ok' : 'warn');
    const btn = $('#grades-save');
    btn.disabled = saved;
    btn.textContent = saved ? 'Guardado ✓' : 'Guardar';
    $('#grades-sub').textContent = `${gname(g)} · ${g.students.length} alumnos · sobre ${gr.escala}`;
    const parts = dg.sections.filter((x) => weightOf(gr, x) > 0);
    if (!parts.length) {
      $('#grades-table').innerHTML = '<tbody><tr><td class="grades-empty">Escribe cuánto vale cada sección para ver las calificaciones.</td></tr></tbody>';
      return;
    }
    const fmt = (v) => (v == null ? '—' : round1(v));
    // La calificación final va junto al nombre para que se vea sin desplazar en el celular.
    let html = '<thead><tr><th class="stick">No.</th><th class="stick name">Nombre</th><th class="final">Calif.</th>' +
      parts.map((x) => `<th>${escapeHtml(x.name || 'Sin nombre')}<br><small>${weightOf(gr, x)}%${x.type === 'sino' ? ` · mín. ${secMin(x)}` : ''}</small></th>`).join('') +
      '</tr></thead><tbody>';
    for (const st of g.students) {
      const r = finalGrade(dg, st.row, gr);
      html += `<tr><td class="stick">${escapeHtml(st.no)}</td><td class="stick name">${escapeHtml(st.name)}</td><td class="v final">${fmt(r.final)}</td>` +
        parts.map((x) => `<td class="v">${fmt(r.parts[x.id])}</td>`).join('') +
        '</tr>';
    }
    const t = $('#grades-table');
    t.innerHTML = html + '</tbody>';
    // No. y Nombre fijos al desplazar a la derecha.
    const w = t.querySelector('thead th.stick').offsetWidth;
    t.querySelectorAll('tr').forEach((tr) => {
      tr.children[0].style.left = '0px';
      tr.children[1].style.left = `${w}px`;
    });
  }

  function openSettings() {
    const f = $('#settings-form');
    [...STATUSES, ...READ].forEach((k) => (f.elements[k].value = state.codes[k]));
    $('#settings').showModal();
  }
  $('#settings').addEventListener('close', () => {
    if ($('#settings').returnValue !== 'save') return;
    const f = $('#settings-form');
    const vals = (keys) => keys.map((k) => f.elements[k].value.trim());
    if ([...vals(STATUSES), ...vals(READ)].some((v) => v === '')) {
      toast('Ninguna opción puede quedar vacía');
      return;
    }
    // Asistencia puede repetir valores (p. ej. A = R = J = 1); Sí y No deben ser distintos.
    const [si, no] = vals(READ);
    if (si.toUpperCase() === no.toUpperCase()) {
      toast('Sí y No necesitan valores distintos');
      return;
    }
    [...STATUSES, ...READ].forEach((k) => (state.codes[k] = f.elements[k].value.trim()));
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
    if (!state.ui.tab) state.ui.tab = ASIST_ID;
    state.codes = { ...DEFAULT_CODES, ...state.codes };
    // Versiones anteriores: porcentajes de todo el archivo y 5 secciones fijas. Cada grupo queda con su copia
    // y con sus secciones (las de siempre: Asistencia, Lecturas, Trabajos, Disciplina y Examen).
    state.groups.forEach((g) => {
      g.grading = fixGrading(g.grading || state.grading);
      normalizeGroup(g, g.grading);
      delete g.grading.lectMin; // el mínimo de "Sí" ahora es de cada sección
    });
    delete state.grading;
  }
  render();

  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
})();
