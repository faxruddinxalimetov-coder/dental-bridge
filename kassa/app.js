/* DentBridge · Kassa
   Klinika kassasining oylik Excel jadvali asosida: oylik varaq + kunlik varaqlar,
   doktor ulushi (Plan), avans (Oldi), Zoir, foyda, olinmagan pul, kassa apparati.
   Formula hisoblagich Kassa-web namunasidan olingan. Summalar ming so'mda. */
(() => {
  'use strict';

  const STORAGE_KEY = 'dentbridge-kassa-v1';
  const MONTHS = ['Yanvar', 'Fevral', 'Mart', 'Aprel', 'May', 'Iyun', 'Iyul', 'Avgust', 'Sentabr', 'Oktabr', 'Noyabr', 'Dekabr'];

  /* Excel kunlik varaqdagi katak manzillari va formulalari */
  const DEFAULT_FORMULAS = {
    day_cash: '=SUM(DoctorCash)',
    day_card: '=SUM(DoctorCard)+OtherCard',
    day_plan_cash: '=SUM(DoctorPlanCash)',
    day_plan_card: '=SUM(DoctorPlanCard)',
    day_submitted: '=A46+C46',
    day_adv: '=SUM(DoctorAdvance)',
    day_bank: '=Bank',
    day_exp: '=SUM(ExpenseAmounts)',
    day_zoir: '=A46-AJ36-K46-P47+OtherCash',
    day_profit_cash: '=S47+K46',
    day_profit_card: '=C46-H46',
    day_profit: '=X46+Z46',
    day_kassa: '=A46-AJ36-P47+OtherCash',
    day_machine_due: '=K46-MachineCash',
    day_unpaid: '=F46-AJ36',
    day_bank_opt: '=(A46-F46)/2',
    day_patients: '=SUM(DoctorPatients)',
    month_cash: '=SUM(DayCash)',
    month_card: '=SUM(DayCard)',
    month_bank: '=SUM(DayBank)',
    month_exp: '=SUM(DayExp)',
    month_zoir: '=SUM(DayZoir)',
    month_profit_cash: '=SUM(DayProfitCash)',
    month_profit_card: '=SUM(DayProfitCard)',
    month_profit: '=AO41+AP41',
    month_adv: '=SUM(DayAdv)',
    month_unpaid: '=SUM(DayUnpaid)',
    month_other_cash: '=SUM(DayOtherCash)',
    month_other_card: '=SUM(DayOtherCard)',
    month_patients: '=SUM(DayPatients)',
    month_machine_card: '=SUM(DayMachineCard)'
  };
  const DAY_DEFS = [
    ['day_cash', 'A46', 'Jami topshirdi — naqd'],
    ['day_card', 'C46', 'Jami topshirdi — plastik (boshqa plastik bilan)'],
    ['day_plan_cash', 'F46', 'Plan (doktor ulushi) — naqd'],
    ['day_plan_card', 'H46', 'Plan (doktor ulushi) — plastik'],
    ['day_submitted', 'A47', 'Jami topshirdi'],
    ['day_adv', 'AJ36', 'Oldi (avanslar) jami'],
    ['day_bank', 'K46', 'Bank'],
    ['day_exp', 'P47', 'Xarajat'],
    ['day_zoir', 'S47', 'Zoir'],
    ['day_profit_cash', 'X46', 'Jami kunlik foyda — naqd'],
    ['day_profit_card', 'Z46', 'Jami kunlik foyda — plastik'],
    ['day_profit', 'X47', 'Jami kunlik foyda'],
    ['day_kassa', 'C50', 'Kassada hozirgi summa'],
    ['day_machine_due', 'J50', 'Kassa aparatga kiritilish kerak summa'],
    ['day_unpaid', 'P50', 'Olinmagan pul'],
    ['day_bank_opt', 'U50', 'Bankga beriladigan optimal summa'],
    ['day_patients', 'B109', 'Qayd qilingan kasallar soni']
  ].map(([key, address, label]) => ({ key, address, label, scope: 'Kunlik' }));
  const MONTH_DEFS = [
    ['month_cash', 'AJ41', 'Jami topshirdi — naqd'],
    ['month_card', 'AK41', 'Jami topshirdi — plastik'],
    ['month_bank', 'AL41', 'Bank'],
    ['month_exp', 'AM41', 'Xarajat'],
    ['month_zoir', 'AN41', 'Zoir'],
    ['month_profit_cash', 'AO41', 'Foyda — naqd'],
    ['month_profit_card', 'AP41', 'Foyda — plastik'],
    ['month_profit', 'AQ41', 'Jami foyda'],
    ['month_adv', 'AR41', 'Oldi (avanslar)'],
    ['month_unpaid', 'AS41', 'Olinmagan pul'],
    ['month_other_cash', 'AT41', 'Boshqa kirim — naqd'],
    ['month_other_card', 'AU41', 'Boshqa kirim — plastik'],
    ['month_patients', 'AV41', 'Qayd qilingan kasallar soni'],
    ['month_machine_card', 'AW41', 'Kassa aparatga kiritilgan plastik']
  ].map(([key, address, label]) => ({ key, address, label, scope: 'Oylik' }));
  const FDEF = Object.fromEntries([...DAY_DEFS, ...MONTH_DEFS].map(d => [d.key, d]));

  /* ---------- Holat ---------- */
  const nowYM = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; };
  function initialState() {
    return { version: 1, selectedMonth: nowYM(), selectedDay: new Date().getDate(), view: 'daily', doctors: [], expenseTypes: ['Material', 'Oylik', 'Kommunal', 'Boshqa'], formulas: { ...DEFAULT_FORMULAS }, months: {} };
  }
  function loadState() {
    try {
      const p = JSON.parse(localStorage.getItem(STORAGE_KEY));
      if (!p || !Array.isArray(p.doctors) || typeof p.months !== 'object') return initialState();
      return { ...initialState(), ...p, formulas: { ...DEFAULT_FORMULAS, ...(p.formulas || {}) } };
    } catch (_) { return initialState(); }
  }
  let state = loadState();
  let importPreview = null;

  function persist() {
    const el = document.getElementById('saveState');
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); if (el) el.innerHTML = '<i></i> Saqlandi'; return true; }
    catch (_) { if (el) el.innerHTML = '<i style="background:#ff8a80"></i> Saqlashda xato'; toast("Brauzer xotirasi to'ldi. Zaxira nusxa oling."); return false; }
  }
  let toastTimer;
  function toast(msg) {
    let n = document.querySelector('.toast');
    if (!n) { n = document.createElement('div'); n.className = 'toast'; n.setAttribute('role', 'status'); document.body.appendChild(n); }
    n.textContent = msg; n.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => n.classList.remove('show'), 2600);
  }

  /* ---------- Yordamchilar ---------- */
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const num = v => { const x = Number(String(v ?? '').replace(/\s/g, '').replace(',', '.')); return Number.isFinite(x) ? x : 0; };
  const sum = a => (a || []).reduce((s, x) => s + num(x), 0);
  const fmt = v => !Number.isFinite(v) ? '#XATO' : (Math.round(v * 100) / 100).toLocaleString('ru-RU', { maximumFractionDigits: 2 });
  const cell = (v, cls = '') => `<td class="${cls}${v < 0 ? ' neg' : ''}">${v ? fmt(v) : ''}</td>`;
  const pct = x => Math.round((x ?? 0) * 1000) / 10;
  const daysIn = ym => { const [y, m] = ym.split('-').map(Number); return new Date(y, m, 0).getDate(); };
  const monthLabel = ym => { const [y, m] = ym.split('-').map(Number); return `${MONTHS[m - 1]} ${y}`; };
  // "150+200 80" yoki "=150+200" -> [150, 200, 80]
  const parseList = s => String(s ?? '').replace(/^=/, '').split(/[+\s;]+/).map(x => x.replace(',', '.')).filter(x => x !== '' && Number.isFinite(Number(x)) && Number(x) !== 0).map(Number);
  const listText = a => (a || []).map(x => fmt(x).replace(/\s/g, '')).join('+');
  const uid = () => 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

  const makeDay = () => ({ doctors: {}, bank: 0, otherCash: 0, otherCard: 0, otherWho: '', machineCash: 0, machineCard: 0, expenses: [] });
  const getDay = (ym = state.selectedMonth, d = state.selectedDay) => state.months[ym]?.days?.[String(d)] || null;
  function ensureDay(ym = state.selectedMonth, d = state.selectedDay) {
    const m = state.months[ym] || (state.months[ym] = { days: {} });
    return m.days[String(d)] || (m.days[String(d)] = makeDay());
  }
  const entry = (day, id) => day?.doctors?.[id] || { cash: [], card: [], adv: [], chair: 0 };
  function dayHasData(day) {
    if (!day) return false;
    if (day.bank || day.otherCash || day.otherCard || day.machineCash || day.machineCard || (day.expenses || []).some(e => num(e.amount))) return true;
    return Object.values(day.doctors || {}).some(e => (e.cash || []).length || (e.card || []).length || (e.adv || []).length || e.chair);
  }
  const docShown = (doc, days) => doc.active || days.some(d => { const e = d?.doctors?.[doc.id]; return e && ((e.cash || []).length || (e.card || []).length || (e.adv || []).length); });

  /* ---------- Formula hisoblagich (Kassa-web namunasidan) ---------- */
  const colNum = name => { let t = 0; for (const c of name.toUpperCase()) t = t * 26 + c.charCodeAt(0) - 64; return t; };
  const colName = n => { let o = ''; while (n > 0) { const r = (n - 1) % 26; o = String.fromCharCode(65 + r) + o; n = Math.floor((n - 1) / 26); } return o; };
  const normCell = r => String(r).replace(/\$/g, '').toUpperCase();
  function splitCell(ref) { const m = normCell(ref).match(/^([A-Z]+)(\d+)$/); if (!m) throw new Error(`Katak manzili noto'g'ri: ${ref}`); return { col: colNum(m[1]), row: +m[2] }; }
  function tokenize(expr) {
    const src = String(expr || '').trim().replace(/^=/, ''), out = [];
    const re = /\s*(>=|<=|<>|\$?[A-Z]{1,3}\$?\d+|\d+(?:\.\d+)?(?:[eE][+-]?\d+)?|[A-Za-z_][A-Za-z0-9_]*|[()+\-*/^,:<>=])/y;
    let i = 0;
    while (i < src.length) { re.lastIndex = i; const m = re.exec(src); if (!m) { if (/^\s+$/.test(src.slice(i))) break; throw new Error(`Formulani o'qib bo'lmadi: ${src.slice(i, i + 12)}`); } out.push(m[1]); i = re.lastIndex; }
    return out;
  }
  function evaluate(expr, ctx) {
    const t = tokenize(expr); let k = 0;
    const peek = () => t[k], take = () => t[k++];
    const nv = v => { if (typeof v === 'boolean') return v ? 1 : 0; if (v == null || v === '') return 0; const x = Number(v); if (!Number.isFinite(x)) throw new Error('Avvalgi katakda xato bor'); return x; };
    const flat = a => a.flat(Infinity).filter(v => v != null && v !== '');
    function cmp() { let l = add(); while (['=', '<>', '>', '<', '>=', '<='].includes(peek())) { const op = take(), a = nv(l), b = nv(add()); l = op === '=' ? a === b : op === '<>' ? a !== b : op === '>' ? a > b : op === '<' ? a < b : op === '>=' ? a >= b : a <= b; } return l; }
    function add() { let l = mul(); while (peek() === '+' || peek() === '-') { const op = take(), r = nv(mul()); l = op === '+' ? nv(l) + r : nv(l) - r; } return l; }
    function mul() { let l = pow(); while (peek() === '*' || peek() === '/') { const op = take(), r = nv(pow()); if (op === '/' && r === 0) throw new Error("Nolga bo'lish mumkin emas"); l = op === '*' ? nv(l) * r : nv(l) / r; } return l; }
    function pow() { let l = un(); if (peek() === '^') { take(); l = Math.pow(nv(l), nv(pow())); } return l; }
    function un() { if (peek() === '+') { take(); return nv(un()); } if (peek() === '-') { take(); return -nv(un()); } return prim(); }
    function prim() {
      const x = take();
      if (x === undefined) throw new Error("Formula bo'sh");
      if (x === '(') { const r = cmp(); if (take() !== ')') throw new Error('Qavs yopilmagan'); return r; }
      if (/^\d/.test(x)) return Number(x);
      if (/^\$?[A-Z]{1,3}\$?\d+$/i.test(x)) { if (peek() === ':') { take(); const e = take(); if (!/^\$?[A-Z]{1,3}\$?\d+$/i.test(e || '')) throw new Error("Oraliq noto'g'ri"); return ctx.range(x, e); } return ctx.cell(x); }
      if (/^[A-Za-z_]/.test(x)) {
        if (peek() === '(') { take(); const args = []; if (peek() !== ')') { do { args.push(cmp()); if (peek() !== ',') break; take(); } while (true); } if (take() !== ')') throw new Error('Funksiya qavsi yopilmagan'); return fn(x, args); }
        return ctx.name(x);
      }
      throw new Error(`Tushunarsiz belgi: ${x}`);
    }
    function fn(name, args) {
      const f = flat(args);
      switch (name.toUpperCase()) {
        case 'SUM': return f.reduce((s, v) => s + nv(v), 0);
        case 'COUNT': return f.filter(v => Number.isFinite(Number(v)) && Number(v) !== 0).length;
        case 'AVERAGE': return f.length ? f.reduce((s, v) => s + nv(v), 0) / f.length : 0;
        case 'MIN': return f.length ? Math.min(...f.map(nv)) : 0;
        case 'MAX': return f.length ? Math.max(...f.map(nv)) : 0;
        case 'ABS': return Math.abs(nv(args[0]));
        case 'ROUND': return Number(nv(args[0]).toFixed(Math.max(0, Math.min(8, nv(args[1] || 0)))));
        case 'IF': return args[0] ? args[1] : args[2];
        default: throw new Error(`Funksiya qo'llab-quvvatlanmaydi: ${name}`);
      }
    }
    const r = cmp();
    if (k < t.length) throw new Error(`Ortiqcha qism: ${t.slice(k).join(' ')}`);
    return r;
  }
  function engine(defs, names) {
    const byAddr = new Map(defs.map(d => [normCell(d.address), d])), cache = {}, errs = {}, visiting = new Set();
    const text = key => state.formulas[key] ?? DEFAULT_FORMULAS[key] ?? '';
    const ctx = {
      name: n => { const m = Object.keys(names).find(k => k.toLowerCase() === n.toLowerCase()); if (m === undefined) throw new Error(`Noma'lum nom: ${n}`); return names[m]; },
      cell: r => get(r),
      range: (a, b) => { const A = splitCell(a), B = splitCell(b), out = []; for (let r = Math.min(A.row, B.row); r <= Math.max(A.row, B.row); r++) for (let c = Math.min(A.col, B.col); c <= Math.max(A.col, B.col); c++) out.push(get(colName(c) + r)); return out; }
    };
    function get(ref) {
      const a = normCell(ref);
      if (a in cache) { if (Number.isNaN(cache[a])) throw new Error(errs[a] || `Xato: ${a}`); return cache[a]; }
      if (visiting.has(a)) throw new Error(`Aylanma formula: ${a}`);
      const d = byAddr.get(a);
      if (!d) return 0;
      visiting.add(a);
      try { cache[a] = Number(evaluate(text(d.key), ctx)); } catch (e) { cache[a] = NaN; errs[a] = e.message; throw e; } finally { visiting.delete(a); }
      return cache[a];
    }
    const values = {}, errors = {};
    for (const d of defs) { try { values[d.key] = get(d.address); if (!Number.isFinite(values[d.key])) errors[d.key] = 'Natija son emas'; } catch (e) { values[d.key] = NaN; errors[d.key] = e.message; } }
    return { values, errors };
  }

  /* ---------- Hisob ---------- */
  function docRows(day) {
    return state.doctors.map(doc => {
      const e = entry(day, doc.id), cash = sum(e.cash), card = sum(e.card), adv = sum(e.adv);
      const pc = doc.pct ?? 0.5, pp = doc.pp ?? pc;
      return { doc, e, cash, card, adv, planCash: cash * pc, planCard: card * pp, patients: (e.cash || []).length + (e.card || []).length, chair: num(e.chair) };
    });
  }
  function calcDay(day) {
    day = day || makeDay();
    const rows = docRows(day);
    const names = {
      DoctorCash: rows.map(r => r.cash), DoctorCard: rows.map(r => r.card), DoctorAdvance: rows.map(r => r.adv),
      DoctorPlanCash: rows.map(r => r.planCash), DoctorPlanCard: rows.map(r => r.planCard), DoctorPatients: rows.map(r => r.patients),
      ExpenseAmounts: (day.expenses || []).map(e => num(e.amount)),
      Bank: num(day.bank), OtherCash: num(day.otherCash), OtherCard: num(day.otherCard), MachineCash: num(day.machineCash), MachineCard: num(day.machineCard)
    };
    return { rows, ...engine(DAY_DEFS, names) };
  }
  function calcMonth(ym = state.selectedMonth) {
    const n = daysIn(ym), days = [];
    for (let d = 1; d <= n; d++) { const day = getDay(ym, d); days.push({ d, day, c: calcDay(day) }); }
    const pick = k => days.map(x => x.c.values[k]);
    const names = {
      DayCash: pick('day_cash'), DayCard: pick('day_card'), DayBank: pick('day_bank'), DayExp: pick('day_exp'), DayZoir: pick('day_zoir'),
      DayProfitCash: pick('day_profit_cash'), DayProfitCard: pick('day_profit_card'), DayAdv: pick('day_adv'), DayUnpaid: pick('day_unpaid'),
      DayOtherCash: days.map(x => num(x.day?.otherCash)), DayOtherCard: days.map(x => num(x.day?.otherCard)),
      DayPatients: pick('day_patients'), DayMachineCard: days.map(x => num(x.day?.machineCard))
    };
    const res = engine(MONTH_DEFS, names), errors = { ...res.errors };
    days.forEach(x => Object.entries(x.c.errors).forEach(([k, m]) => errors[`${x.d}-kun · ${FDEF[k].address}`] = m));
    return { days, values: res.values, errors };
  }
  const fx = (key, v, cls = '') => `<td class="fx ${cls}${v < 0 ? ' neg' : ''}" data-fx="${key}">${fmt(v)}</td>`;

  /* ---------- Ko'rinishlar ---------- */
  function renderDaily() {
    const ym = state.selectedMonth, d = state.selectedDay, day = getDay(ym, d), c = calcDay(day), v = c.values;
    const rows = c.rows.filter(r => r.doc.active || r.cash || r.card || r.adv);
    if (!state.doctors.length) return `<div class="note">Doktorlar hali yo'q. <b>Sozlamalar</b> bo'limida doktor qo'shing yoki <b>Excel import</b> orqali eski kassa fayllaringizni yuklang.</div>`;
    const tN = rows.reduce((s, r) => s + r.cash, 0), tP = rows.reduce((s, r) => s + r.card, 0);
    const head = `<tr><th class="lbl">${String(d).padStart(2, '0')}.${ym.slice(5)}.${ym.slice(0, 4)}</th>${rows.map(r => `<th colspan="2" class="name" title="Ulush: naqd ${pct(r.doc.pct ?? .5)}%, plastik ${pct(r.doc.pp ?? r.doc.pct ?? .5)}%">${esc(r.doc.name)}</th>`).join('')}<th colspan="2" class="jami">Jami</th></tr>
      <tr><th class="lbl"></th>${rows.map(() => '<th class="naqd">Naqd</th><th class="plastik">Plastik</th>').join('')}<th class="naqd">Naqd</th><th class="plastik">Plastik</th></tr>`;
    const inp = (r, f, cls = '') => `<td class="in ${cls}"><input id="i-${r.doc.id}-${f}" data-k="doc" data-doc="${esc(r.doc.id)}" data-f="${f}" value="${esc(f === 'chair' ? (r.chair || '') : listText(r.e[f]))}" aria-label="${esc(r.doc.name)} ${f}" inputmode="decimal"></td>`;
    const body = [
      `<tr><td class="lbl">To'lovlar</td>${rows.map(r => inp(r, 'cash') + inp(r, 'card')).join('')}<td></td><td></td></tr>`,
      `<tr><td class="lbl">Kasallar soni</td>${rows.map(r => `<td class="cnt">${(r.e.cash || []).length || ''}</td><td class="cnt">${(r.e.card || []).length || ''}</td>`).join('')}${fx('day_patients', v.day_patients)}<td></td></tr>`,
      `<tr><td class="lbl top">Topshirdi</td>${rows.map(r => cell(r.cash, 'top') + cell(r.card, 'top2')).join('')}${cell(tN, 'top')}${cell(tP, 'top2')}</tr>`,
      `<tr><td class="lbl plann">Plan</td>${rows.map(r => cell(r.planCash, 'plann') + cell(r.planCard, 'planp')).join('')}${fx('day_plan_cash', v.day_plan_cash, 'plann')}${fx('day_plan_card', v.day_plan_card, 'planp')}</tr>`,
      `<tr><td class="lbl">Oldi (avans)</td>${rows.map(r => `<td class="in" colspan="2"><input id="i-${r.doc.id}-adv" data-k="doc" data-doc="${esc(r.doc.id)}" data-f="adv" value="${esc(listText(r.e.adv))}" aria-label="${esc(r.doc.name)} avans" inputmode="decimal"></td>`).join('')}<td colspan="2"></td></tr>`,
      `<tr><td class="lbl umumiy">Umumiy</td>${rows.map(r => `<td colspan="2" class="umumiy">${r.adv ? fmt(r.adv) : ''}</td>`).join('')}${fx('day_adv', v.day_adv, 'umumiy')}<td></td></tr>`,
      `<tr><td class="lbl oladi">Oladi:</td>${rows.map(r => cell(r.planCash - r.adv, 'peach') + '<td></td>').join('')}${fx('day_unpaid', v.day_unpaid, 'peach')}<td></td></tr>`,
      `<tr><td class="lbl">Qoldi:</td>${rows.map(r => cell(r.cash - r.adv, 'naqd') + cell(r.card - r.planCard, 'plastik')).join('')}${cell(tN - v.day_adv, 'naqd')}${cell(tP - v.day_plan_card, 'plastik')}</tr>`,
      `<tr><td class="lbl">Kresloda o'tirgan</td>${rows.map(r => inp(r, 'chair') + '<td></td>').join('')}<td>${fmt(rows.reduce((s, r) => s + r.chair, 0)) || ''}</td><td></td></tr>`,
      `<tr><td class="lbl">Kassaga to'lamagan</td>${rows.map(r => { const x = r.chair ? r.chair - r.patients : 0; return cell(x, x > 0 ? 'peach' : '') + '<td></td>'; }).join('')}<td></td><td></td></tr>`
    ].join('');
    const exps = (day?.expenses || []).map((e, i) => `<tr><td class="in"><input class="txt" id="e-${i}-paidTo" data-k="exp" data-i="${i}" data-f="paidTo" value="${esc(e.paidTo)}" placeholder="Nimaga sarflandi" aria-label="Nimaga sarflandi"></td>
      <td class="in"><select data-k="exp" data-i="${i}" data-f="type" aria-label="Xarajat turi" style="width:100%;height:24px;border:0;background:transparent"><option value="" ${e.type ? '' : 'selected'}>—</option>${[...new Set([...state.expenseTypes, e.type || ''])].filter(Boolean).map(t => `<option ${t === e.type ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select></td>
      <td class="in"><input id="e-${i}-amount" data-k="exp" data-i="${i}" data-f="amount" value="${num(e.amount) || ''}" inputmode="decimal" aria-label="Summa"></td><td><button class="x" data-action="delExp" data-i="${i}" aria-label="O'chirish">×</button></td></tr>`).join('');
    const dIn = (f, label, txt) => `<td class="in"><input id="f-${f}" data-k="day" data-f="${f}" value="${esc(txt ? (day?.[f] || '') : (num(day?.[f]) || ''))}" ${txt ? 'class="txt" placeholder="Kim berdi"' : 'inputmode="decimal"'} aria-label="${label}"></td>`;
    const sumRows = [
      ['Jami topshirdi', 'jami', fx('day_cash', v.day_cash, 'sn'), fx('day_card', v.day_card, 'yp')],
      ['Plan (doktorlar ulushi)', 'plann', fx('day_plan_cash', v.day_plan_cash, 'plann'), fx('day_plan_card', v.day_plan_card, 'planp')],
      ['Oldi (avanslar)', 'umumiy', fx('day_adv', v.day_adv, 'umumiy'), '<td></td>'],
      ['Xarajat', 'kul', fx('day_exp', v.day_exp), '<td></td>'],
      ['Bankaga berilgan summa', 'bank', dIn('bank', 'Bankka berilgan summa'), '<td></td>'],
      ['Zoir', 'naqd', fx('day_zoir', v.day_zoir), '<td></td>'],
      ['Jami kunlik foyda', 'jami', fx('day_profit_cash', v.day_profit_cash, 'sn'), fx('day_profit_card', v.day_profit_card, 'yp')],
      ['Jami kunlik foyda (naqd + plastik)', 'jami', fx('day_profit', v.day_profit), '<td></td>'],
      ['Boshqa kirimlar', 'boshqa', dIn('otherCash', 'Boshqa kirim naqd'), dIn('otherCard', 'Boshqa kirim plastik')],
      ['Kim berdi', 'plann', `<td class="in" colspan="2"><input id="f-otherWho" class="txt" data-k="day" data-f="otherWho" value="${esc(day?.otherWho || '')}" placeholder="Kim berdi / izoh" aria-label="Kim berdi"></td>`, ''],
      ['Kassa aparatga kiritildi', 'apparat', dIn('machineCash', 'Kassa apparati naqd'), dIn('machineCard', 'Kassa apparati plastik')],
      ['Kassada hozirgi summa', 'summa', fx('day_kassa', v.day_kassa, 'summa'), '<td></td>'],
      ['Kassa aparatga kiritilish kerak summa', 'apparat', fx('day_machine_due', v.day_machine_due), '<td></td>'],
      ['Olinmagan pul', 'peach', fx('day_unpaid', v.day_unpaid, 'peach'), '<td></td>'],
      ['Bankga beriladigan optimal summa', 'summa', fx('day_bank_opt', v.day_bank_opt), '<td></td>']
    ].map(([l, c, a, b]) => `<tr><td class="lbl ${c}">${l}</td>${a}${b}</tr>`).join('');
    const errs = Object.keys(c.errors).length ? `<div class="note err">Formula xatosi: ${Object.entries(c.errors).map(([k, m]) => `${FDEF[k].address} — ${esc(m)}`).join('; ')}</div>` : '';
    return `${errs}
      <div class="cards">${[['Jami topshirdi', v.day_submitted, 'jami'], ['Zoir', v.day_zoir, 'naqd'], ['Jami kunlik foyda', v.day_profit, 'summa'], ['Kassada hozirgi summa', v.day_kassa, 'summa'], ['Olinmagan pul', v.day_unpaid, 'peach'], ['Kasallar soni', v.day_patients, 'sn']].map(([k, x, cl]) => `<div class="card xl"><div class="k ${cl}" style="margin:-8px -10px 6px;padding:4px 10px">${k}</div><div class="v ${x < 0 ? 'neg' : ''}">${fmt(x)}</div></div>`).join('')}</div>
      <div class="grid-wrap"><table class="xl"><thead>${head}</thead><tbody>${body}</tbody></table></div>
      <div class="two">
        <section class="blk"><div class="blk-head"><h2>Xarajat</h2><button class="rb-btn" data-action="addExp">+ Qator qo'shish</button></div>
          <div class="grid-wrap"><table class="xl" style="width:100%"><thead><tr><th class="kul" style="text-align:left">Nimaga sarflandi</th><th class="peach">Turi</th><th class="summa">Summa</th><th></th></tr></thead><tbody>${exps || '<tr><td colspan="4" class="muted" style="text-align:center">Xarajat yo\'q</td></tr>'}<tr class="tot"><td>Jami</td><td></td>${fx('day_exp', v.day_exp, 'summa')}<td></td></tr></tbody></table></div></section>
        <section class="blk"><div class="blk-head"><h2>Kunlik hisob</h2></div>
          <div class="grid-wrap"><table class="xl" style="width:100%"><thead><tr><th class="lbl"></th><th class="sn">Naqd</th><th class="yp">Plastik</th></tr></thead><tbody>${sumRows}</tbody></table></div></section>
      </div>`;
  }

  function renderMonthly() {
    const ym = state.selectedMonth, m = calcMonth(ym), v = m.values, all = m.days.map(x => x.day);
    const docs = state.doctors.filter(doc => docShown(doc, all));
    const extra = [['Naqd', 'oyn', 'day_cash'], ['Plastik', 'bank', 'day_card'], ['Bank', '', 'day_bank'], ['Xarajat', 'kul', 'day_exp'], ['Zoir', 'naqd', 'day_zoir'], ['Foyda naqd', 'summa', 'day_profit_cash'], ['Foyda plastik', 'summa', 'day_profit_card'], ['Jami foyda', 'jami', 'day_profit'], ['Oldi', 'umumiy', 'day_adv'], ['Olinmagan', 'peach', 'day_unpaid'], ['Boshqa naqd', 'boshqa', null, 'otherCash'], ['Boshqa plastik', 'boshqa', null, 'otherCard'], ['Kasallar', 'sn', 'day_patients'], ['Kassa ap. plastik', 'apparat', null, 'machineCard']];
    const head = `<tr><th class="lbl" rowspan="2">${esc(monthLabel(ym))}</th>${docs.map(d => `<th colspan="2" class="name">${esc(d.name)}</th>`).join('')}${extra.map(([t, c]) => `<th rowspan="2" class="${c}">${t}</th>`).join('')}</tr><tr>${docs.map(() => '<th class="naqd">Naqd</th><th class="plastik">Plastik</th>').join('')}</tr>`;
    const tot = {};
    const body = m.days.map(({ d, day, c }) => {
      const has = dayHasData(day);
      const dc = docs.map(doc => { const r = c.rows.find(x => x.doc.id === doc.id); tot[doc.id] = tot[doc.id] || [0, 0]; tot[doc.id][0] += r.cash; tot[doc.id][1] += r.card; return cell(r.cash) + cell(r.card); }).join('');
      const ex = extra.map(([, cl, k, f]) => cell(k ? c.values[k] : num(day?.[f]), cl === 'jami' || cl === 'oyn' ? cl : '')).join('');
      return `<tr${has ? '' : ' class="muted"'}><td class="lbl"><button class="day-link" data-open-day="${d}">${MONTHS[+ym.slice(5) - 1]} (${d})</button></td>${dc}${ex}</tr>`;
    }).join('');
    const mk = { day_cash: 'month_cash', day_card: 'month_card', day_bank: 'month_bank', day_exp: 'month_exp', day_zoir: 'month_zoir', day_profit_cash: 'month_profit_cash', day_profit_card: 'month_profit_card', day_profit: 'month_profit', day_adv: 'month_adv', day_unpaid: 'month_unpaid', otherCash: 'month_other_cash', otherCard: 'month_other_card', day_patients: 'month_patients', machineCard: 'month_machine_card' };
    const totRow = `<tr class="tot"><td class="lbl top">Topshirdi</td>${docs.map(d => cell(tot[d.id]?.[0] || 0, 'top') + cell(tot[d.id]?.[1] || 0, 'umumiy')).join('')}${extra.map(([, cl, k, f]) => fx(mk[k || f], v[mk[k || f]], cl)).join('')}</tr>`;
    const errN = Object.keys(m.errors).length;
    return `${errN ? `<div class="note err">${errN} ta formula xatosi bor. <button class="rb-btn" data-view="formulas">Formulalarni ko'rish</button></div>` : ''}
      <div class="cards">${[['Jami topshirdi (naqd)', v.month_cash, 'oyn'], ['Jami topshirdi (plastik)', v.month_card, 'bank'], ['Bank', v.month_bank, 'apparat'], ['Xarajat', v.month_exp, 'kul'], ['Zoir', v.month_zoir, 'naqd'], ['Jami foyda', v.month_profit, 'jami'], ['Olinmagan pul', v.month_unpaid, 'peach'], ['Kasallar soni', v.month_patients, 'sn']].map(([k, x, cl]) => `<div class="card xl"><div class="k ${cl}" style="margin:-8px -10px 6px;padding:4px 10px">${k}</div><div class="v ${x < 0 ? 'neg' : ''}">${fmt(x)}</div></div>`).join('')}</div>
      ${state.doctors.length ? `<div class="grid-wrap"><table class="xl"><thead>${head}</thead><tbody>${body}${totRow}</tbody></table></div><p class="muted">Kun nomini bosing — kunlik varaq ochiladi. Yashil qalin raqamlar formula bilan hisoblanadi (ustiga olib boring).</p>` : '<div class="note">Doktorlar yo\'q. Sozlamalar yoki Excel import orqali qo\'shing.</div>'}`;
  }

  function renderFormulas() {
    const dc = calcDay(getDay()), mc = calcMonth();
    const rows = (defs, res) => defs.map(d => { const e = res.errors[d.key]; return `<tr><td class="lbl" style="font-weight:400">${esc(d.label)}</td><td style="text-align:center;font-family:Consolas,monospace">${d.address}</td><td class="in"><input class="formula-input" data-formula="${d.key}" value="${esc(state.formulas[d.key] ?? DEFAULT_FORMULAS[d.key])}" aria-label="${esc(d.label)} formulasi"></td><td class="${e ? 'neg' : 'fx'}" title="${esc(e || '')}">${e ? '#XATO' : fmt(res.values[d.key])}</td><td><button class="rb-btn" data-action="saveFormula" data-key="${d.key}">Saqlash</button></td></tr>`; }).join('');
    const t = (title, defs, res) => `<section class="blk"><div class="blk-head"><h2>${title}</h2></div><div class="grid-wrap"><table class="xl" style="width:100%"><thead><tr><th class="lbl">Nomi</th><th>Katak</th><th>Formula</th><th>Natija</th><th></th></tr></thead><tbody>${rows(defs, res)}</tbody></table></div></section>`;
    return `<div class="note ok">Formulalar Excel kunlik varaqdagi kataklar bilan bir xil manzilda. Nomlar: <code>DoctorCash</code>, <code>DoctorCard</code>, <code>DoctorPlanCash</code>, <code>DoctorPlanCard</code>, <code>DoctorAdvance</code>, <code>DoctorPatients</code>, <code>ExpenseAmounts</code>, <code>Bank</code>, <code>OtherCash</code>, <code>OtherCard</code>, <code>MachineCash</code>, <code>MachineCard</code>. Funksiyalar: SUM, COUNT, AVERAGE, MIN, MAX, ABS, ROUND, IF.</div>
      ${t(`Kunlik varaq · ${state.selectedDay}-kun`, DAY_DEFS, dc)}${t(`Oylik varaq · ${monthLabel(state.selectedMonth)}`, MONTH_DEFS, mc)}
      <p><button class="rb-btn danger" data-action="resetFormulas">Formulalarni asl holiga qaytarish</button></p>`;
  }

  function renderImport() {
    const r = importPreview;
    let prev = '';
    if (r?.busy) prev = '<div class="note">Fayllar o\'qilmoqda…</div>';
    else if (r?.err) prev = `<div class="note err">${esc(r.err)}</div>`;
    else if (r) {
      const ds = Object.keys(r.days).sort();
      prev = `<div class="cards">${[['Fayllar', r.files], ['Kunlik varaqlar', r.sheets], ['Kunlar', ds.length], ["O'tkazib yuborildi", r.skipped]].map(([k, x]) => `<div class="card xl"><div class="k jami" style="margin:-8px -10px 6px;padding:4px 10px">${k}</div><div class="v">${x}</div></div>`).join('')}</div>
        <p>Davr: <b>${esc(ds[0] || '—')}</b> — <b>${esc(ds.at(-1) || '—')}</b></p>
        <div class="grid-wrap" style="max-width:520px"><table class="xl"><thead><tr><th class="lbl name">Doktor</th><th class="plann">Naqd ulush %</th><th class="planp">Plastik ulush %</th><th>Bazada</th></tr></thead><tbody>${r.people.map(p => `<tr><td class="lbl">${esc(p.n)}</td><td>${pct(p.pct)}</td><td>${pct(p.pp ?? p.pct)}</td><td style="text-align:center">${state.doctors.some(d => d.name.toLowerCase() === p.n.toLowerCase()) ? 'bor' : 'yangi'}</td></tr>`).join('')}</tbody></table></div>
        <p class="row"><label class="row"><input type="checkbox" id="impReplace" checked> Mavjud kunlarni almashtirish</label></p>
        <p><button class="rb-btn primary" data-action="impGo">Kassaga qo'shish</button></p>`;
    }
    return `<div class="note ok">Oylik kassa Excel fayllarini (<b>.xlsx</b>) tanlang — bir nechtasini birdaniga. Har bir «2026 Avgust (4)» ko'rinishidagi kunlik varaqdan har bir bemor to'lovi (naqd/plastik), Plan ulushi, Oldi, xarajatlar va izohi, bank, boshqa kirimlar, kassa apparati va kreslo soni olinadi. Fayllar kompyuterdan chiqmaydi.</div>
      <p><button class="rb-btn primary" data-action="pickXlsx">Excel fayllarni tanlash</button></p>${prev}`;
  }

  function renderSettings() {
    const docs = state.doctors.map(d => `<tr><td class="in"><input class="txt" data-k="docset" data-doc="${esc(d.id)}" data-f="name" value="${esc(d.name)}" aria-label="Doktor ismi"></td>
      <td class="in"><input data-k="docset" data-doc="${esc(d.id)}" data-f="pct" value="${pct(d.pct ?? .5)}" inputmode="decimal" aria-label="${esc(d.name)} naqd ulushi"></td>
      <td class="in"><input data-k="docset" data-doc="${esc(d.id)}" data-f="pp" value="${pct(d.pp ?? d.pct ?? .5)}" inputmode="decimal" aria-label="${esc(d.name)} plastik ulushi"></td>
      <td style="text-align:center"><button class="rb-btn" data-action="toggleDoc" data-doc="${esc(d.id)}">${d.active ? 'Faol' : 'Arxivda'}</button></td></tr>`).join('');
    return `<div class="two">
      <section><div class="blk-head"><h2>Doktorlar va ulush (Plan)</h2></div>
        <form class="row" id="addDoc" style="margin-bottom:8px"><input class="inp" id="newDoc" placeholder="Yangi doktor ismi" required maxlength="50"><input class="inp num" id="newPct" value="50" aria-label="Ulush foizi"><span>%</span><button class="rb-btn primary">+ Qo'shish</button></form>
        <div class="grid-wrap"><table class="xl" style="width:100%"><thead><tr><th class="name" style="text-align:left">Ism</th><th class="plann">Naqd %</th><th class="planp">Plastik %</th><th>Holat</th></tr></thead><tbody>${docs || '<tr><td colspan="4" class="muted" style="text-align:center">Doktor yo\'q</td></tr>'}</tbody></table></div>
        <p class="muted">Ulush: doktor topshirgan pulning necha foizi unga tegishli (Excel'dagi «Plan» qatori). Masalan 50% = <code>/2</code>, 45% = <code>*0.45</code>.</p></section>
      <section><div class="blk-head"><h2>Xarajat turlari</h2></div>
        <form class="row" id="addType" style="margin-bottom:8px"><input class="inp" id="newType" placeholder="Masalan: dori-darmon" required maxlength="50"><button class="rb-btn primary">+ Qo'shish</button></form>
        <div class="list">${state.expenseTypes.map((t, i) => `<div class="row" style="justify-content:space-between;border:1px solid var(--grid);padding:4px 8px"><span>${esc(t)}</span><button class="x rb-btn danger" data-action="delType" data-i="${i}">O'chirish</button></div>`).join('')}</div>
        <div class="blk-head blk"><h2>Ma'lumotlar</h2></div>
        <p class="row"><button class="rb-btn" data-action="exportJson">JSON zaxira</button><button class="rb-btn" data-action="importJson">Zaxiradan tiklash</button><button class="rb-btn" data-action="exportCsv">Joriy oy CSV</button><button class="rb-btn danger" data-action="resetAll">Hammasini o'chirish</button></p>
        <p class="muted">Ma'lumotlar faqat shu brauzerda saqlanadi. Muntazam zaxira nusxa oling.</p></section></div>`;
  }

  function renderSheets() {
    const ym = state.selectedMonth, n = daysIn(ym), mn = MONTHS[+ym.slice(5) - 1];
    const tabs = [`<button class="sheet oy ${state.view === 'monthly' ? 'active' : ''}" data-view="monthly">${mn}</button>`];
    for (let d = n; d >= 1; d--) tabs.push(`<button class="sheet ${dayHasData(getDay(ym, d)) ? '' : 'blank'} ${state.view === 'daily' && state.selectedDay === d ? 'active' : ''}" data-open-day="${d}">${ym.slice(0, 4)} ${mn} (${d})</button>`);
    return tabs.join('');
  }

  function render() {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(state.selectedMonth)) state.selectedMonth = nowYM();
    state.selectedDay = Math.min(Math.max(1, state.selectedDay || 1), daysIn(state.selectedMonth));
    const focusId = document.activeElement?.id;
    document.getElementById('monthPicker').value = state.selectedMonth;
    document.getElementById('fileName').textContent = `${monthLabel(state.selectedMonth)}.xlsx — ${({ monthly: 'oylik varaq', daily: `${state.selectedDay}-kun`, formulas: 'formulalar', import: 'Excel import', settings: 'sozlamalar' })[state.view] || ''}`;
    document.querySelectorAll('.xl-tab[data-view]').forEach(t => t.classList.toggle('active', t.dataset.view === state.view));
    const app = document.getElementById('app');
    app.innerHTML = ({ monthly: renderMonthly, formulas: renderFormulas, import: renderImport, settings: renderSettings })[state.view]?.() ?? renderDaily();
    const sh = document.getElementById('sheets');
    sh.innerHTML = renderSheets();
    sh.querySelector('.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    if (focusId) { const el = document.getElementById(focusId); if (el && app.contains(el)) { el.focus(); if (el.select) el.select(); } }
  }
  const rerender = () => setTimeout(render, 0);

  /* ---------- Eksport / import ---------- */
  function download(blob, name) { const a = document.createElement('a'), u = URL.createObjectURL(blob); a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 1000); }
  function exportCsv() {
    const ym = state.selectedMonth, m = calcMonth(ym), all = m.days.map(x => x.day), docs = state.doctors.filter(d => docShown(d, all));
    const q = x => `"${String(x ?? '').replace(/"/g, '""')}"`;
    const head = ['Kun', ...docs.flatMap(d => [`${d.name} naqd`, `${d.name} plastik`]), 'Jami naqd', 'Jami plastik', 'Bank', 'Xarajat', 'Zoir', 'Foyda naqd', 'Foyda plastik', 'Jami foyda', 'Oldi', 'Olinmagan', 'Boshqa naqd', 'Boshqa plastik', 'Kasallar', 'Kassa ap. plastik'];
    const lines = [head.map(q).join(';')];
    for (const { d, day, c } of m.days) {
      const v = c.values;
      lines.push([d, ...docs.flatMap(doc => { const r = c.rows.find(x => x.doc.id === doc.id); return [r.cash, r.card]; }), v.day_cash, v.day_card, v.day_bank, v.day_exp, v.day_zoir, v.day_profit_cash, v.day_profit_card, v.day_profit, v.day_adv, v.day_unpaid, num(day?.otherCash), num(day?.otherCard), v.day_patients, num(day?.machineCard)].map(q).join(';'));
    }
    const v = m.values;
    lines.push(['Jami', ...docs.flatMap(doc => [m.days.reduce((s, x) => s + x.c.rows.find(r => r.doc.id === doc.id).cash, 0), m.days.reduce((s, x) => s + x.c.rows.find(r => r.doc.id === doc.id).card, 0)]), v.month_cash, v.month_card, v.month_bank, v.month_exp, v.month_zoir, v.month_profit_cash, v.month_profit_card, v.month_profit, v.month_adv, v.month_unpaid, v.month_other_cash, v.month_other_card, v.month_patients, v.month_machine_card].map(q).join(';'));
    download(new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), `Kassa-${ym}.csv`);
    toast('Oylik jadval CSV formatida yuklandi (Excel ochadi).');
  }
  function exportJson() { download(new Blob([JSON.stringify(state)], { type: 'application/json' }), `Kassa-zaxira-${new Date().toISOString().slice(0, 10)}.json`); toast('Zaxira nusxa yuklandi.'); }
  function importJson(file) {
    file.text().then(t => {
      const p = JSON.parse(t);
      if (!p || !Array.isArray(p.doctors) || typeof p.months !== 'object') throw new Error();
      if (!confirm("Joriy kassa ma'lumotlari zaxira bilan almashtiriladi. Davom etasizmi?")) return;
      state = { ...initialState(), ...p, formulas: { ...DEFAULT_FORMULAS, ...(p.formulas || {}) } };
      persist(); render(); toast('Zaxira tiklandi.');
    }).catch(() => toast('Bu Kassa zaxira fayli emas.'));
  }
  function loadXLSX() {
    return window.XLSX ? Promise.resolve() : new Promise((ok, no) => { const s = document.createElement('script'); s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'; s.onload = ok; s.onerror = () => no(new Error('Excel kutubxonasi yuklanmadi. Internetni tekshiring.')); document.head.appendChild(s); });
  }
  async function readXlsx(files) {
    importPreview = { busy: true }; render();
    try {
      await loadXLSX();
      let acc;
      for (const f of files) acc = KassaImport.parseWorkbook(XLSX.read(await f.arrayBuffer(), { type: 'array', cellFormula: true }), acc);
      const r = KassaImport.finish(acc); r.files = files.length;
      if (!Object.keys(r.days).length) r.err = "Kunlik varaqlar topilmadi. Varaq nomi «2026 Avgust (4)» ko'rinishida bo'lishi kerak.";
      importPreview = r;
    } catch (e) { importPreview = { err: e.message || String(e) }; }
    if (state.view === 'import') render();
  }
  function importMerge() {
    const r = importPreview, replace = document.getElementById('impReplace')?.checked !== false, K = 1000, ids = {};
    for (const p of r.people) {
      let d = state.doctors.find(x => x.name.toLowerCase() === p.n.toLowerCase());
      if (!d) { d = { id: uid(), name: p.n, active: true, pct: p.pct, pp: p.pp ?? p.pct }; state.doctors.push(d); }
      else { d.pct = p.pct; d.pp = p.pp ?? p.pct; }
      ids[p.n] = d.id;
    }
    const idOf = n => ids[n] || (ids[n] = (() => { const d = { id: uid(), name: n, active: true, pct: .5, pp: .5 }; state.doctors.push(d); return d.id; })());
    let added = 0, kept = 0;
    for (const [date, x] of Object.entries(r.days)) {
      const ym = date.slice(0, 7), dn = String(+date.slice(8));
      if (getDay(ym, dn) && !replace) { kept++; continue; }
      const day = makeDay();
      const ent = n => day.doctors[idOf(n)] || (day.doctors[idOf(n)] = { cash: [], card: [], adv: [], chair: 0 });
      if (x.lists) for (const [n, l] of Object.entries(x.lists)) { ent(n).cash = l.n.map(s => s / K); ent(n).card = l.p.map(s => s / K); }
      else for (const [n, m, s] of x.in) ent(n)[m === 'n' ? 'cash' : 'card'].push(s / K);
      for (const [n, s] of x.adv) ent(n).adv = [s / K];
      for (const [n, c] of Object.entries(x.chair || {})) ent(n).chair = c;
      day.expenses = x.exp.map(([t, s]) => ({ paidTo: t, type: '', amount: s / K }));
      day.bank = (x.bank || 0) / K;
      day.otherCash = x.oth.filter(o => o[1] === 'n').reduce((s, o) => s + o[2], 0) / K;
      day.otherCard = x.oth.filter(o => o[1] === 'p').reduce((s, o) => s + o[2], 0) / K;
      day.otherWho = [...new Set(x.oth.map(o => o[0]))].join(', ');
      if (x.mach) { day.machineCash = x.mach[0] / K; day.machineCard = x.mach[1] / K; }
      (state.months[ym] || (state.months[ym] = { days: {} })).days[dn] = day;
      added++;
    }
    const last = Object.keys(r.days).sort().at(-1);
    if (last) { state.selectedMonth = last.slice(0, 7); state.selectedDay = +last.slice(8); }
    importPreview = null; state.view = 'monthly';
    if (persist()) toast(`${added} kun qo'shildi${kept ? `, ${kept} kun o'zgarishsiz qoldi` : ''}.`);
    render();
  }

  /* ---------- Hodisalar ---------- */
  document.addEventListener('click', e => {
    const v = e.target.closest('[data-view]');
    if (v) { state.view = v.dataset.view; persist(); return render(); }
    const od = e.target.closest('[data-open-day]');
    if (od) { state.selectedDay = +od.dataset.openDay; state.view = 'daily'; persist(); return render(); }
    const b = e.target.closest('[data-action]');
    if (!b) return;
    const a = b.dataset.action;
    if (a === 'exportCsv') return exportCsv();
    if (a === 'exportJson') return exportJson();
    if (a === 'importJson') return document.getElementById('importFile').click();
    if (a === 'pickXlsx') { const f = document.getElementById('xlsxFile'); f.value = ''; return f.click(); }
    if (a === 'impGo') return importMerge();
    if (a === 'addExp') { ensureDay().expenses.push({ paidTo: '', type: state.expenseTypes[0] || '', amount: 0 }); persist(); render(); document.getElementById(`e-${ensureDay().expenses.length - 1}-paidTo`)?.focus(); return; }
    if (a === 'delExp') { ensureDay().expenses.splice(+b.dataset.i, 1); persist(); return render(); }
    if (a === 'saveFormula') { const i = document.querySelector(`[data-formula="${b.dataset.key}"]`); state.formulas[b.dataset.key] = (i?.value || '').trim() || DEFAULT_FORMULAS[b.dataset.key]; persist(); render(); return toast('Formula saqlandi va qayta hisoblandi.'); }
    if (a === 'resetFormulas') { if (!confirm('Barcha formulalar asl holiga qaytarilsinmi?')) return; state.formulas = { ...DEFAULT_FORMULAS }; persist(); return render(); }
    if (a === 'toggleDoc') { const d = state.doctors.find(x => x.id === b.dataset.doc); if (d) d.active = !d.active; persist(); return render(); }
    if (a === 'delType') { state.expenseTypes.splice(+b.dataset.i, 1); persist(); return render(); }
    if (a === 'resetAll') { if (!confirm("Barcha kassa ma'lumotlari o'chiriladi. Avval zaxira nusxa olganingizga ishonchingiz komilmi?")) return; state = initialState(); persist(); render(); return toast("Ma'lumotlar o'chirildi."); }
  });
  document.addEventListener('change', e => {
    const t = e.target;
    if (t.id === 'monthPicker') { if (!t.value) return; state.selectedMonth = t.value; persist(); return render(); }
    if (t.id === 'importFile') { if (t.files[0]) importJson(t.files[0]); t.value = ''; return; }
    if (t.id === 'xlsxFile') { const fs = [...t.files]; if (fs.length) readXlsx(fs); return; }
    const k = t.dataset.k;
    if (!k) return;
    if (k === 'doc') {
      const day = ensureDay(), en = day.doctors[t.dataset.doc] || (day.doctors[t.dataset.doc] = { cash: [], card: [], adv: [], chair: 0 });
      if (t.dataset.f === 'chair') en.chair = Math.max(0, Math.round(num(t.value))); else en[t.dataset.f] = parseList(t.value);
    } else if (k === 'day') {
      const day = ensureDay(); day[t.dataset.f] = t.dataset.f === 'otherWho' ? t.value.trim() : num(t.value);
    } else if (k === 'exp') {
      const x = ensureDay().expenses[+t.dataset.i]; if (!x) return;
      x[t.dataset.f] = t.dataset.f === 'amount' ? num(t.value) : t.value;
    } else if (k === 'docset') {
      const d = state.doctors.find(x => x.id === t.dataset.doc); if (!d) return;
      if (t.dataset.f === 'name') { const n = t.value.trim(); if (n) d.name = n; }
      else d[t.dataset.f] = Math.min(100, Math.max(0, num(t.value))) / 100;
    }
    persist(); rerender();
  });
  document.addEventListener('submit', e => {
    e.preventDefault();
    if (e.target.id === 'addDoc') {
      const n = document.getElementById('newDoc').value.trim(), p = Math.min(100, Math.max(0, num(document.getElementById('newPct').value))) / 100;
      if (!n) return;
      if (state.doctors.some(d => d.name.toLowerCase() === n.toLowerCase())) return toast("Bu doktor ro'yxatda bor.");
      state.doctors.push({ id: uid(), name: n, active: true, pct: p, pp: p }); persist(); render(); toast(`${n} qo'shildi.`);
    }
    if (e.target.id === 'addType') {
      const n = document.getElementById('newType').value.trim();
      if (!n || state.expenseTypes.some(x => x.toLowerCase() === n.toLowerCase())) return;
      state.expenseTypes.push(n); persist(); render();
    }
  });
  // Formula qatori (fx): katak ustiga borganda yoki input tanlanganda ko'rsatiladi
  function showFx(el) {
    const name = document.getElementById('fxName'), text = document.getElementById('fxText');
    const k = el?.closest?.('[data-fx]')?.dataset.fx;
    if (k) { name.textContent = FDEF[k].address; text.textContent = `${state.formulas[k] ?? DEFAULT_FORMULAS[k]}   — ${FDEF[k].label}`; return; }
    if (el?.matches?.('input[data-k]')) { name.textContent = el.getAttribute('aria-label') || ''; text.textContent = el.value; }
  }
  document.addEventListener('mouseover', e => showFx(e.target));
  document.addEventListener('focusin', e => showFx(e.target));
  document.addEventListener('input', e => { if (e.target.matches('input[data-k]')) document.getElementById('fxText').textContent = e.target.value; });
  document.addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.target.matches('td.in input')) { e.preventDefault(); const all = [...document.querySelectorAll('#app td.in input')], i = all.indexOf(e.target); (all[i + 1] || e.target).focus(); }
  });

  render();
})();
