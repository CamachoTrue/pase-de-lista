import { ASIST_ID } from '../lib/excel-core.js';
import { confirmBulk, confirmChange } from './dialogs.js';
import { refresh, renderSummary } from './main-view.js';
import { DISC_DEFAULT, FILL_REST, LABELS, OPTIONS, curWork, dayData, group, lista, optionOf, save, secById, state, tab, typeOf, workIdx, works } from './state.js';
import { $, escapeHtml, normalize, toast } from './util.js';

// ---------- Vista compacta (lecturas / trabajos) ----------
export function renderCompact() {
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
  $('#compact-list').innerHTML = lista().map((s) => {
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
export const pendingRows = () => {
  const day = dayData(tab(), false);
  return lista().filter((s) => day[s.row] == null);
};
// Sí/No: si todos tienen la misma marca, botón para cambiarlos todos a la otra.
export const flipTarget = () => {
  if (typeOf(tab()) !== 'sino') return null;
  const day = dayData(tab(), false);
  const vals = new Set(lista().map((s) => day[s.row]));
  if (vals.size !== 1 || !lista().length) return null;
  const [only] = vals;
  return only === 'S' ? 'N' : only === 'N' ? 'S' : null;
};
export function updateFillRest() {
  const t = typeOf(tab());
  const box = $('#fill-rest');
  const n = t in FILL_REST ? pendingRows().length : 0;
  const flip = n ? null : flipTarget();
  box.hidden = !n && !flip;
  if (flip) {
    const total = lista().length;
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
    const students = lista();
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
// Buscador de la pestaña Disciplina.
export function filterCompact() {
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
export function renderWorkBar() {
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
export function goWork(i) {
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

