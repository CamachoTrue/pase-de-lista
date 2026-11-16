(function () {
  'use strict';
  const { STATUSES, READ, WORK, CATS, DEFAULT_CODES, parseWorkbook, exportWorkbook, groupDates, columnLayout, discCols, examCols, shortDate,
    GRADE_PARTS, DEFAULT_GRADING, finalGrade, gradeCols, round1, worksOf, normalizeWorks } = window.ListaExcel;
  // Todos los apartados que se guardan por día (disciplina no lleva columna diaria en el Excel).
  const KEYS = [...CATS.map((c) => c.key), 'disc', 'exam'];
  const LABELS = { A: 'Asistencia', F: 'Falta', R: 'Retardo', J: 'Justificado' };
  // Opciones de cada apartado: valor guardado, texto del botón, texto largo y color (clase A/F/R/J).
  const OPTIONS = {
    lect: [
      { v: 'S', btn: 'Sí', label: 'Sí leyó', color: 'A' },
      { v: 'N', btn: 'No', label: 'No leyó', color: 'F' },
    ],
    trab: [
      { v: 0, btn: '0', label: 'Trabajo: 0', color: 'F' },
      { v: 50, btn: '50', label: 'Trabajo: 50', color: 'R' },
      { v: 100, btn: '100', label: 'Trabajo: 100', color: 'A' },
    ],
    // Disciplina: todos tienen 10 sin capturar nada; solo se guarda el 0.
    disc: [
      { v: 10, btn: '10', label: 'Disciplina: 10', color: 'N' },
      { v: 0, btn: '0', label: 'Disciplina: 0', color: 'F' },
    ],
  };
  const DISC_DEFAULT = 10;
  // Valor del botón "restantes" en lecturas y trabajos.
  const FILL_REST = { lect: 'N', trab: 0 };
  const optionOf = (cat, v) => {
    if (cat === 'records') return v ? { v, label: LABELS[v], color: v } : null;
    if (cat === 'exam') return v != null ? { v, label: `Examen: ${v}`, color: 'N' } : null;
    return OPTIONS[cat].find((o) => o.v === v) || null;
  };
  const tab = () => state.ui.tab || 'records';
  // Datos de un apartado para el día seleccionado.
  // Trabajos: puede haber varios por día; se usa el trabajo seleccionado (state.ui.work).
  const works = (create) => {
    const g = group();
    g.trab = g.trab || {};
    if (!g.trab[state.ui.date] && create) g.trab[state.ui.date] = [];
    return g.trab[state.ui.date] || [];
  };
  const workIdx = () => Math.max(0, Math.min(state.ui.work || 0, Math.max(works(false).length - 1, 0)));
  const curWork = (create) => {
    const ws = works(create);
    const i = workIdx();
    if (!ws[i] && create) ws[i] = { title: '', vals: {} };
    return ws[i] || null;
  };
  const dayData = (cat, create) => {
    if (cat === 'trab') { const w = curWork(create); return w ? w.vals : {}; }
    const g = group();
    g[cat] = g[cat] || {};
    if (!g[cat][state.ui.date] && create) g[cat][state.ui.date] = {};
    return g[cat][state.ui.date] || {};
  };
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
      const grading = (state && state.grading) || JSON.parse(JSON.stringify(DEFAULT_GRADING));
      state = { fileName: file.name, fileB64: bufToB64(buf), groups, codes: { ...DEFAULT_CODES, ...codes }, grading, ui: { group: 0, date: todayISO(), tab: 'records' } };
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
    const t = tab();
    document.querySelectorAll('.tab').forEach((b) => {
      b.classList.toggle('on', b.dataset.tab === t);
      b.setAttribute('aria-selected', b.dataset.tab === t);
    });
    const att = t === 'records';
    $('#card-view').hidden = !att || view !== 'card';
    $('#list-view').hidden = !att || view !== 'list';
    $('#compact-view').hidden = att;
    if (!att) renderCompact();
    else if (view === 'card') renderCard(false);
    else renderStudents();
    renderSummary();
  }

  // ---------- Pestañas ----------
  $('#tabs').addEventListener('click', (e) => {
    const b = e.target.closest('.tab');
    if (!b || b.dataset.tab === tab()) return;
    state.ui.tab = b.dataset.tab;
    save();
    if (state.ui.tab === 'records') startCard();
    refresh();
    window.scrollTo(0, 0);
  });

  // ---------- Vista compacta (lecturas / trabajos) ----------
  function renderCompact() {
    const g = group();
    const cat = tab();
    const day = dayData(cat, false);
    const att = dayData('records', false);
    renderWorkBar();
    const disc = cat === 'disc';
    const exam = cat === 'exam';
    $('#compact-search-bar').hidden = !disc && !exam;
    $('#compact-hint').innerHTML = exam
      ? 'Escribe el porcentaje del examen (0 a 100). Con <b>Enter</b> pasas al siguiente alumno.'
      : 'Todos tienen 10. Busca al alumno y toca <b>0</b> si se portó mal.';
    $('#compact-list').innerHTML = g.students.map((s) => {
      const cur = optionOf(cat, disc ? (day[s.row] ?? DISC_DEFAULT) : day[s.row]);
      const a = att[s.row];
      const btns = exam
        ? `<input class="exam-in" type="text" inputmode="decimal" enterkeyhint="next" maxlength="5" value="${day[s.row] ?? ''}" aria-label="Examen de ${escapeHtml(s.name)}" placeholder="—">`
        : OPTIONS[cat].map((o) => {
        const on = cur && cur.v === o.v;
        return `<button class="cb ${cat === 'trab' && o.v === 100 ? 'wide' : ''} ${on ? 'on c' + o.color : ''}" data-v="${o.v}" aria-pressed="${!!on}" aria-label="${escapeHtml(o.label)}">${o.btn}</button>`;
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
  function updateFillRest() {
    const cat = tab();
    const b = $('#fill-rest');
    const n = cat in FILL_REST ? pendingRows().length : 0;
    b.hidden = n === 0;
    if (n) b.textContent = cat === 'lect' ? `Poner “No” a los ${n} sin marcar` : `Poner 0 a los ${n} sin calificar`;
  }
  $('#fill-rest').addEventListener('click', async () => {
    const cat = tab();
    const rows = pendingRows();
    if (!(cat in FILL_REST) || !rows.length) return;
    const v = FILL_REST[cat];
    const what = cat === 'lect' ? '“No leyó”' : '0 de trabajo';
    if (!(await confirmBulk(`¿Poner ${what} a ${rows.length} alumno(s)?`, 'Solo a los que siguen en blanco. Los que ya marcaste no cambian.'))) return;
    const day = dayData(cat, true);
    rows.forEach((s) => (day[s.row] = v));
    save();
    if (navigator.vibrate) navigator.vibrate(15);
    refresh();
    toast(`${rows.length} alumno(s) con ${cat === 'lect' ? 'No' : '0'}`);
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
    const li = btn.closest('.crow');
    const row = li.dataset.row;
    const disc = cat === 'disc';
    const v = cat === 'lect' ? btn.dataset.v : Number(btn.dataset.v);
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
      const o = OPTIONS[cat].find((x) => String(x.v) === b.dataset.v);
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
    if (!inp || tab() !== 'exam') return;
    const li = inp.closest('.crow');
    const row = li.dataset.row;
    const prev = dayData('exam', false)[row];
    const txt = inp.value.trim().replace(',', '.');
    let next;
    if (txt !== '') {
      next = Number(txt);
      if (!Number.isFinite(next) || next < 0 || next > 100) {
        toast('El examen va de 0 a 100');
        inp.value = prev ?? '';
        return;
      }
    }
    if (next === prev) { inp.value = prev ?? ''; return; }
    if (prev != null) {
      const name = group().students.find((x) => String(x.row) === row).name;
      if (!(await confirmChange(name, prev, next, 'exam'))) { inp.value = prev; return; }
    }
    const day = dayData('exam', true);
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
    bar.hidden = tab() !== 'trab';
    if (bar.hidden) return;
    const ws = works(false);
    const total = Math.max(ws.length, 1);
    const i = workIdx();
    state.ui.work = i;
    const w = ws[i];
    $('#work-pos').textContent = `Trabajo ${i + 1} de ${total}`;
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
    if (!ws.length) delete group().trab[state.ui.date];
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
  function confirmChange(name, from, to, cat = 'records') {
    return new Promise((resolve) => {
      const d = $('#confirm-change');
      const pill = (sel, v) => {
        const el = $(sel);
        const o = optionOf(cat, v);
        el.className = `pill-lg ${o ? o.color : 'none'}`;
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
    if (tab() !== 'records' || view !== 'card' || $('#screen-main').hidden || document.querySelector('dialog[open]')) return;
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
    if (cat === 'exam') {
      const day = dayData(cat, false);
      const vals = g.students.map((s) => day[s.row]).filter((v) => v != null);
      const avg = vals.length ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 10) / 10 : null;
      $('#summary').innerHTML =
        `<span>${fmtDate(state.ui.date)}</span><span>${vals.length}/${g.students.length} capturados</span>` +
        (avg != null ? `<span>Promedio: ${avg}</span>` : '');
      return;
    }
    if (cat === 'disc') {
      const day = dayData(cat, false);
      const zeros = g.students.filter((s) => day[s.row] === 0).length;
      $('#summary').innerHTML =
        `<span>${fmtDate(state.ui.date)}</span><span class="F">Con 0: ${zeros}</span><span>Los demás: 10</span>`;
      return;
    }
    if (cat !== 'records') {
      const day = dayData(cat, false);
      const vals = g.students.map((s) => day[s.row]).filter((v) => v != null);
      $('#summary').innerHTML =
        `<span>${vals.length}/${g.students.length} capturados</span>` +
        OPTIONS[cat].map((o) => `<span class="${o.color}">${o.btn}: ${vals.filter((v) => v === o.v).length}</span>`).join('');
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
    KEYS.forEach((k) => Object.values(g[k] || {}).forEach((day) => {
      if (Array.isArray(day)) day.forEach((w) => delete w.vals[s.row]); // trabajos
      else delete day[s.row];
    }));
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
    state.ui.tab = 'records';
    save();
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
      const out = await exportWorkbook(ExcelJS, b64ToBuf(state.fileB64), [g], state.codes, state.grading);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(out);
      ws = wb.getWorksheet(g.sheet);
    } catch (err) {
      console.error(err);
      toast('No se pudo generar la vista previa');
      return;
    }

    const dates = groupDates(g);
    const layout = columnLayout(g);
    const byCol = new Map(layout.map((x) => [x.col, x]));
    const dc = discCols(g);
    const ex = examCols(g);
    const exCol = new Map(ex.map((x, i) => [x.col, { ...x, first: i === 0, last: i === ex.length - 1 }]));
    const gcs = gradeCols(g, state.grading);
    const gcCol = new Map(gcs.map((x, i) => [x.col, { ...x, first: i === 0, last: i === gcs.length - 1 }]));
    const lastCol = Math.max(g.nameCol, ...layout.map((x) => x.col), dc ? dc.dates : 0, ...ex.map((x) => x.col), ...gcs.map((x) => x.col));
    const isCur = (c) => byCol.has(c) && byCol.get(c).date === state.ui.date;
    const colLetter = (c) => { let s = ''; for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };
    // Se toma de lo capturado (no del número), porque A, R y J pueden tener el mismo valor.
    const colorOf = (c, r) => {
      const { cat, date, work } = byCol.get(c);
      const v = cat === 'trab' ? (worksOf(g, date)[work].vals || {})[r] : ((g[cat] || {})[date] || {})[r];
      if (v == null) return '';
      if (cat === 'records') return v;
      const o = optionOf(cat, v);
      return o ? o.color : '';
    };
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
        const isDateCol = byCol.has(c);
        if (isDateCol && isHdr) cls.push('date-h');
        if (isDateCol && !isHdr) cls.push('v', colorOf(c, r));
        if (isDateCol && byCol.get(c).first) cls.push('day-start');
        if (isDateCol && byCol.get(c).cat === 'lect') cls.push('lect-col');
        if (isDateCol && !byCol.has(c + 1)) cls.push('block-end');
        if (isDateCol && isCur(c)) cls.push('cur');
        const isExam = exCol.has(c);
        if (isExam) {
          const x = exCol.get(c);
          cls.push(isHdr ? 'date-h' : 'v');
          if (x.first) cls.push('day-start');
          if (x.last) cls.push('block-end');
          if (x.date === state.ui.date) cls.push('cur');
        }
        const isGrade = gcCol.has(c);
        if (isGrade) {
          const x = gcCol.get(c);
          cls.push(isHdr ? 'date-h' : 'v');
          if (x.first || x.key === 'final') cls.push('day-start');
          if (x.last) cls.push('block-end');
          if (x.key === 'final') cls.push('final');
        }
        const isDisc = (dc && (c === dc.count || c === dc.dates)) || isExam || isGrade;
        if (isDisc && isHdr) cls.push('date-h');
        if (dc && c === dc.count) cls.push('day-start');
        if (dc && c === dc.dates) cls.push('block-end', 'disc-dates');
        if (dc && c === dc.count && !isHdr) cls.push('v', Number(t) > 0 ? 'F' : '');
        // El encabezado lleva fecha y apartado en dos renglones, como en el Excel.
        const shown = isHdr && (isDateCol || isDisc) ? escapeHtml(ListaExcel.cellText(v)).replace(/\n/g, '<br>') : escapeHtml(t);
        html += `<td class="${cls.join(' ')}">${shown}</td>`;
      }
      html += '</tr>';
    });
    html += '</tbody>';
    $('#preview-table').innerHTML = html;

    $('#preview-title').textContent = `Vista previa · ${g.sheet}`;
    $('#preview-sub').textContent = dates.length || dc
      ? `${g.students.length} alumnos · ${dates.length} día(s) · Asistencia y Lectura juntas; Trabajos y Disciplina aparte a la derecha`
      : 'Aún no hay registros en este grupo.';
    $('#preview-legend').innerHTML =
      STATUSES.map((k) => `<span class="summary-${k}">${escapeHtml(String(state.codes[k]))} = ${LABELS[k]}</span>`).join('') +
      READ.map((k) => { const o = optionOf('lect', k); return `<span class="summary-${o.color}">${escapeHtml(String(state.codes[k]))} = ${o.label}</span>`; }).join('');

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
      const out = await exportWorkbook(ExcelJS, b64ToBuf(state.fileB64), state.groups, state.codes, state.grading);
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
    if (action === 'grades') openGrades();
    if (action === 'history') openHistory();
    if (action === 'settings') openSettings();
    if (action === 'clear-day') {
      confirmBulk(`¿Borrar los registros del ${fmtDate(state.ui.date)}?`,
        `Se borran asistencia, lecturas, trabajos, disciplina y examen de ese día en ${group().sheet}. No se puede deshacer.`,
        'Sí, borrar', true).then((ok) => {
        if (!ok) return;
        KEYS.forEach((k) => { if (group()[k]) delete group()[k][state.ui.date]; });
        changed();
      });
    }
    if (action === 'new-file') {
      if (confirm('Se reemplazará la lista actual y su asistencia guardada. Exporta antes si la necesitas. ¿Continuar?')) {
        $('#file-input').click();
      }
    }
  });

  function openHistory() {
    const g = group();
    const count = (cat, d) => (cat === 'trab'
      ? worksOf(g, d).filter((w) => Object.keys(w.vals).length).length
      : Object.keys((g[cat] || {})[d] || {}).length);
    const dates = [...new Set([...groupDates(g), ...['disc', 'exam'].flatMap((k) => Object.keys(g[k] || {}).filter((d) => count(k, d)))])].sort().reverse();
    $('#history-list').innerHTML = dates.length
      ? dates.map((d) => {
          const c = { A: 0, F: 0, R: 0, J: 0 };
          Object.values(g.records[d] || {}).forEach((s) => c[s]++);
          return `<li data-date="${d}"><span>${fmtDate(d)}</span><span class="counts">A ${c.A} · F ${c.F} · R ${c.R} · J ${c.J}<br>Lecturas ${count('lect', d)} · Trabajos ${count('trab', d)} (${worksOf(g, d).reduce((n, w) => n + Object.keys(w.vals).length, 0)} calif.)${count('disc', d) ? ` · Disciplina 0: ${count('disc', d)}` : ''}${count('exam', d) ? ` · Examen: ${count('exam', d)}` : ''}</span></li>`;
        }).join('')
      : '<li>Aún no hay registros en este grupo.</li>';
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

  // ---------- Calificaciones ----------
  // Los porcentajes son del archivo (todos los grupos); la calificación se recalcula sola con cada dato.
  // Lo que se escribe es un borrador: la tabla se actualiza al instante, pero solo se guarda con "Guardar".
  let gradeDraft = null;
  const sameGrading = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  function openGrades() {
    const f = $('#grades-form');
    gradeDraft = JSON.parse(JSON.stringify(state.grading));
    const gr = gradeDraft;
    GRADE_PARTS.forEach((p) => (f.elements[p.key].value = gr.w[p.key] || ''));
    f.elements.retardo.value = String(gr.retardo);
    f.elements.escala.value = String(gr.escala);
    f.elements.lectMin.value = gr.lectMin;
    renderGrades();
    $('#grades').showModal();
  }
  $('#grades-form').addEventListener('input', () => {
    const f = $('#grades-form');
    GRADE_PARTS.forEach((p) => {
      const n = Number(f.elements[p.key].value.replace(',', '.'));
      gradeDraft.w[p.key] = Number.isFinite(n) && n > 0 ? n : 0;
    });
    gradeDraft.retardo = Number(f.elements.retardo.value);
    gradeDraft.escala = Number(f.elements.escala.value);
    const lm = Math.round(Number(f.elements.lectMin.value));
    gradeDraft.lectMin = Number.isFinite(lm) && lm >= 1 ? lm : state.grading.lectMin;
    renderGrades();
  });
  const draftTotal = () => round1(GRADE_PARTS.reduce((n, p) => n + (gradeDraft.w[p.key] || 0), 0));
  $('#grades-save').addEventListener('click', () => {
    const total = draftTotal();
    if (total !== 100) { toast(`Los porcentajes suman ${total} %; deben sumar 100 %`); return; }
    state.grading = JSON.parse(JSON.stringify(gradeDraft));
    save();
    renderGrades();
    toast('Porcentajes guardados. La calificación saldrá siempre en la vista previa y el Excel.');
  });
  // Al cerrar sin guardar se descartan los cambios (avisando).
  $('#grades').addEventListener('close', () => {
    if (gradeDraft && !sameGrading(gradeDraft, state.grading)) toast('No se guardaron los cambios en los porcentajes');
    gradeDraft = null;
  });
  function renderGrades() {
    const g = group();
    const gr = gradeDraft;
    const total = draftTotal();
    const saved = sameGrading(gradeDraft, state.grading);
    const tot = $('#weights-total');
    tot.textContent = total !== 100
      ? `Total: ${total} % — debe sumar 100 %`
      : saved ? 'Total: 100 % · Guardado ✓' : 'Total: 100 % · Falta guardar';
    tot.className = 'weights-total ' + (total !== 100 ? 'bad' : saved ? 'ok' : 'warn');
    const btn = $('#grades-save');
    btn.disabled = saved;
    btn.textContent = saved ? 'Guardado ✓' : 'Guardar porcentajes';
    $('#grades-sub').textContent = `${g.sheet} · ${g.students.length} alumnos · sobre ${gr.escala}`;
    const parts = GRADE_PARTS.filter((p) => gr.w[p.key] > 0);
    if (!parts.length) {
      $('#grades-table').innerHTML = '<tbody><tr><td class="grades-empty">Escribe cuánto vale cada apartado para ver las calificaciones.</td></tr></tbody>';
      return;
    }
    const fmt = (v) => (v == null ? '—' : round1(v));
    // La calificación final va junto al nombre para que se vea sin desplazar en el celular.
    let html = '<thead><tr><th class="stick">No.</th><th class="stick name">Nombre</th><th class="final">Calif.</th>' +
      parts.map((p) => `<th>${p.label}<br><small>${gr.w[p.key]}%${p.key === 'lect' ? ` · mín. ${gr.lectMin}` : ''}</small></th>`).join('') +
      '</tr></thead><tbody>';
    for (const s of g.students) {
      const r = finalGrade(g, s.row, gr);
      html += `<tr><td class="stick">${escapeHtml(s.no)}</td><td class="stick name">${escapeHtml(s.name)}</td><td class="v final">${fmt(r.final)}</td>` +
        parts.map((p) => `<td class="v">${fmt(r.parts[p.key])}</td>`).join('') +
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
    if (!state.ui.tab) state.ui.tab = 'records';
    // Datos guardados con la versión anterior (solo asistencia).
    state.codes = { ...DEFAULT_CODES, ...state.codes };
    state.groups.forEach((g) => { g.lect = g.lect || {}; normalizeWorks(g); g.disc = g.disc || {}; g.exam = g.exam || {}; });
    state.grading = { ...DEFAULT_GRADING, ...state.grading, w: { ...DEFAULT_GRADING.w, ...(state.grading || {}).w } };
  }
  render();

  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
})();
