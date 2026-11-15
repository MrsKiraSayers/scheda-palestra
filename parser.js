// Lettore Excel -> scheda. Riconosce: giorni come fogli, come righe-titolo o come colonna "Giorno".
(function (root) {
  const RX = {
    day: /^(giorno|day|seduta|sessione|allenamento|workout)\b/i,
    name: /eserciz|exercise|^nome$|movimento/i,
    sets: /^(n[°.]?\s*)?(serie|sets?)\b/i,
    reps: /^(n[°.]?\s*)?(rip|rep)/i,
    note: /note|^rec\b|recuper|rest|pausa|riposo|tempo|tut|carico|peso|kg|rpe|buffer/i,
    dayTitle: /^(giorno|day|seduta|sessione|allenamento|workout|scheda)\b|^(luned|marted|mercoled|gioved|venerd|sabato|domenica)/i
  };

  function detectHeader(r) {
    const c = { name: -1, sets: -1, reps: -1, day: -1, notes: [] };
    r.forEach((v, i) => {
      if (!v) return;
      if (c.name < 0 && RX.name.test(v)) c.name = i;
      else if (c.sets < 0 && RX.sets.test(v)) c.sets = i;
      else if (c.reps < 0 && RX.reps.test(v)) c.reps = i;
      else if (c.day < 0 && RX.day.test(v) && r.filter(Boolean).length > 2) c.day = i;
      else if (RX.note.test(v)) c.notes.push({ i, label: v });
    });
    const hits = (c.name >= 0) + (c.sets >= 0) + (c.reps >= 0);
    return hits >= 2 ? c : null;
  }

  function parseRows(rows, sheetName, days) {
    let cols = null, day = null;
    const newDay = n => { day = { name: n, exercises: [] }; days.push(day); return day; };
    rows.forEach(raw => {
      const r = raw.map(c => String(c == null ? '' : c).trim());
      const idx = r.map((v, i) => v ? i : -1).filter(i => i >= 0);
      if (!idx.length) return;
      const h = detectHeader(r);
      if (h) { cols = h; return; }
      // riga-titolo: una sola cella, oppure inizia con "Giorno A", "Lunedì"...
      if (idx.length === 1 || (idx.length <= 2 && RX.dayTitle.test(r[idx[0]]) && !/\d+\s*[x×]/i.test(r[idx[1]] || ''))) {
        if (idx.length === 1 && !RX.dayTitle.test(r[idx[0]]) && cols && cols.name === idx[0] && day) {
          day.exercises.push({ name: r[idx[0]], sets: '', reps: '', note: '' }); // esercizio senza serie/rip
        } else newDay(idx.map(i => r[i]).join(' '));
        return;
      }
      const combo = !cols && /^\d+\s*[x×]/i.test(r[idx[1]] || '');
      const c = cols || { name: idx[0], sets: idx[1], reps: combo ? -1 : idx[2], day: -1, notes: idx.slice(combo ? 2 : 3).map(i => ({ i, label: '' })) };
      const name = r[c.name >= 0 ? c.name : idx[0]];
      if (!name) return;
      let sets = c.sets >= 0 ? r[c.sets] || '' : '', reps = c.reps >= 0 && c.reps !== undefined ? r[c.reps] || '' : '';
      const m = /^(\d+)\s*[x×]\s*(.+)$/i.exec(sets) || (!sets && /^(\d+)\s*[x×]\s*(.+)$/i.exec(reps));
      if (m) { sets = m[1]; reps = m[2]; }
      const note = c.notes.filter(n => r[n.i]).map(n => (n.label && !/^note$/i.test(n.label) ? n.label + ': ' : '') + r[n.i]).join(' · ');
      if (c.day >= 0 && r[c.day] && (!day || day.name !== r[c.day])) day = days.find(d => d.name === r[c.day]) || newDay(r[c.day]);
      if (!day) newDay(sheetName);
      day.exercises.push({ name, sets, reps, note });
    });
  }

  function parseWorkbook(XLSX, wb) {
    const days = [];
    wb.SheetNames.forEach(sn => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, blankrows: false, defval: '', raw: false });
      const start = days.length;
      parseRows(rows, sn, days);
      // più fogli e nessun titolo interno: il nome del foglio è il giorno
      if (wb.SheetNames.length > 1 && days.length - start === 1 && !RX.dayTitle.test(days[start].name)) days[start].name = sn;
    });
    return days.filter(d => d.exercises.length);
  }

  root.parseWorkbook = parseWorkbook;
  if (typeof module !== 'undefined') module.exports = { parseWorkbook };
})(typeof self !== 'undefined' ? self : globalThis);
