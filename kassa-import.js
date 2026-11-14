/* DentBridge: kassa Excel (oylik .xlsx) fayllarini baza formatiga o'tkazish.
   Brauzerda ham (window.KassaImport), Node'da ham (module.exports) ishlaydi.
   Excel ichidagi summalar ming so'mda yozilgan, bazada so'mda saqlanadi.

   Natija:
   { people: [{n, pct}],
     days: { 'YYYY-MM-DD': { in:[[ism,'n'|'p',som]], adv:[[ism,som]], exp:[[izoh,som]], bank:som, oth:[[izoh,'n'|'p',som]] } } } */
(function (root) {
  const MONTHS = { yanvar: 1, fevral: 2, mart: 3, aprel: 4, may: 5, iyun: 6, iyul: 7, avgust: 8, sentabr: 9, sentyabr: 9, oktabr: 10, oktyabr: 10, noyabr: 11, dekabr: 12 };
  const K = 1000;

  function colName(c) { let s = ''; c++; while (c) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }
  function sheetDate(name) {
    const m = /^\s*(\d{4})\s+([A-Za-z']+)\s*\((\d{1,2})\)\s*$/.exec(name);
    if (!m) return null;
    const mo = MONTHS[m[2].toLowerCase()];
    if (!mo) return null;
    const d = +m[3];
    if (d < 1 || d > 31) return null;
    return `${m[1]}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }

  // "Plan" formulasidan ulushni olish: =B25/2 -> 0.5, =F25*0.45 -> 0.45, =R25*2/3 -> 0.667
  function planPct(x) {
    const f = x && x.f;
    if (!f) return 0.5;
    const e = String(f).replace(/^=/, '').replace(/^\$?[A-Z]{1,3}\$?\d+/, '1');
    if (!/^[0-9.*\/+\-() ]+$/.test(e)) return 0.5;
    try { const v = Function('return (' + e + ')')(); return isFinite(v) && v >= 0 && v <= 1 ? Math.round(v * 1000) / 1000 : 0.5; } catch (_) { return 0.5; }
  }

  function parseSheet(ws) {
    const cell = (c, r) => ws[colName(c) + r];
    const val = (c, r) => { const x = cell(c, r); return x ? x.v : undefined; };
    const num = (c, r) => { const v = val(c, r); return typeof v === 'number' && isFinite(v) ? v : 0; };
    const str = (c, r) => { const v = val(c, r); return typeof v === 'string' ? v.trim() : ''; };
    // A ustunidagi yorliqlar bo'yicha qatorlarni topish
    const rowOf = {};
    for (let r = 1; r <= 80; r++) { const a = str(0, r).toLowerCase(); if (a && !(a in rowOf)) rowOf[a] = r; }
    const rTop = rowOf['topshirdi'], rPlan = rowOf['plan'], rOldi = rowOf['oldi'], rUm = rowOf['umumiy'], rXar = rowOf['xarajat'];
    if (!rTop) return null;
    // 7-qatordagi ismlar (B, D, F ... juft ustunlar: Naqd / Plastik)
    const people = [];
    for (let c = 1; c < 60; c += 2) {
      const n = str(c, 7);
      if (!n) continue;
      if (/jami|bank|xarajat|zoir|foida|kirim|kassa/i.test(n)) break;
      people.push({ n, c, pct: planPct(rPlan && cell(c, rPlan)), pp: planPct(rPlan && cell(c + 1, rPlan)) });
    }
    const day = { in: [], adv: [], exp: [], bank: 0, oth: [] };
    for (const p of people) {
      let n = 0, pl = 0;
      for (let r = 9; r < rTop; r++) { n += num(p.c, r); pl += num(p.c + 1, r); }
      if (n) day.in.push([p.n, 'n', Math.round(n * K)]);
      if (pl) day.in.push([p.n, 'p', Math.round(pl * K)]);
      if (rOldi && rUm) {
        let a = 0;
        for (let r = rOldi; r < rUm; r++) { a += num(p.c, r) + num(p.c + 1, r); }
        if (a) day.adv.push([p.n, Math.round(a * K)]);
      }
    }
    // Xarajat: izoh 39..42 qatorlarda, summa "Summa:" qatorida
    if (rXar) {
      let rSum = 0;
      for (let r = rXar; r < rXar + 8; r++) if (/summa/i.test(str(1, r))) { rSum = r; break; }
      if (rSum) for (let c = 3; c < 32; c += 2) {
        const s = num(c, rSum);
        if (!s) continue;
        const t = [];
        for (let r = rXar; r < rSum; r++) { const x = str(c, r); if (x) t.push(x); }
        day.exp.push([t.join(' ').replace(/\s+/g, ' ') || 'Xarajat', Math.round(s * K)]);
      }
    }
    // Bank va boshqa kirimlar: yorliq bo'yicha qidirish
    for (let r = 40; r <= 60; r++) for (let c = 0; c < 50; c++) {
      const s = str(c, r);
      if (/^bankaga berilgan summa/i.test(s)) { for (let k = c + 1; k < c + 6; k++) { const v = num(k, r); if (v) { day.bank = Math.round(v * K); break; } } }
      if (/^boshqa kirimlar/i.test(s)) {
        for (let k = 33; k < 44; k += 2) {
          const nq = num(k, r + 2), pq = num(k, r + 3);
          const who = str(k, r + 1) || 'Boshqa kirim';
          if (nq) day.oth.push([who, 'n', Math.round(nq * K)]);
          if (pq) day.oth.push([who, 'p', Math.round(pq * K)]);
        }
      }
    }
    return { people, day };
  }

  function empty(d) { return !d.in.length && !d.adv.length && !d.exp.length && !d.bank && !d.oth.length; }
  function weight(d) { return d.in.length + d.adv.length + d.exp.length + d.oth.length + (d.bank ? 1 : 0); }

  /* wb: XLSX.read(...) natijasi. acc: oldingi natija (bir nechta faylni yig'ish uchun) */
  function parseWorkbook(wb, acc) {
    acc = acc || { people: [], days: {}, sheets: 0, skipped: 0 };
    const pmap = new Map(acc.people.map(p => [p.n, p]));
    for (const name of wb.SheetNames) {
      const date = sheetDate(name);
      if (!date) { acc.skipped++; continue; }
      const res = parseSheet(wb.Sheets[name]);
      if (!res) { acc.skipped++; continue; }
      acc.sheets++;
      for (const p of res.people) {
        const cur = pmap.get(p.n);
        if (!cur) { const np = { n: p.n, pct: p.pct, pp: p.pp, last: date }; pmap.set(p.n, np); acc.people.push(np); }
        else if (date >= cur.last) { cur.pct = p.pct; cur.pp = p.pp; cur.last = date; }
      }
      if (empty(res.day)) continue;
      // Bir sana ikki faylda bo'lsa, ma'lumoti ko'prog'i olinadi
      const old = acc.days[date];
      if (!old || weight(res.day) > weight(old)) acc.days[date] = res.day;
    }
    return acc;
  }

  function finish(acc) {
    return { people: acc.people.map(p => ({ n: p.n, pct: p.pct, pp: p.pp })), days: acc.days, sheets: acc.sheets, skipped: acc.skipped };
  }

  const api = { parseWorkbook, finish, sheetDate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.KassaImport = api;
})(typeof window !== 'undefined' ? window : this);
