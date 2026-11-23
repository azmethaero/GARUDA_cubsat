/* ===========================================================================
   G.A.R.U.D.A. shared helpers
   Here's where your team details are remembered and where every download gets
   its header: team, team ID, lab, CubeSat name, date and time.
   =========================================================================== */
(function (g) {
  const KEY = 'garuda-names';
  const DEF = { team: '', teamId: '', lab: '', cubesat: '', operator: '' };

  // The details you typed on the Welcome page, saved on this computer only.
  function names() { try { return { ...DEF, ...(JSON.parse(localStorage.getItem(KEY)) || {}) }; } catch (e) { return { ...DEF }; } }
  function setNames(p) {
    const n = { ...names(), ...p };
    try { localStorage.setItem(KEY, JSON.stringify(n)); } catch (e) { /* storage blocked: lasts for this visit only */ }
    return n;
  }

  // Which questions each section asks before it opens.
  const FIELDS = {
    lab: [['team', 'Team name', 'e.g. Team Vega'], ['teamId', 'Team ID', 'e.g. T-042'], ['lab', 'Lab name', 'e.g. Bench Lab 1'], ['cubesat', 'CubeSat name', 'e.g. Nano-1']],
    telemetry: [['team', 'Team name', 'e.g. Team Vega'], ['teamId', 'Team ID', 'e.g. T-042'], ['operator', 'Your name (person on duty)', 'e.g. Niranjan CS'], ['cubesat', 'CubeSat name', 'e.g. Nano-1']]
  };
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  function formHtml(mode) {
    const n = names();
    return FIELDS[mode].map(([k, l, ph]) => `<label class="lbl">${l}<input class="field" data-k="${k}" placeholder="${ph}" value="${esc(n[k] || '')}" maxlength="60" autocomplete="off"></label>`).join('');
  }
  function readForm(root) { const o = {}; root.querySelectorAll('input[data-k]').forEach(i => { o[i.dataset.k] = i.value.trim(); }); return o; }

  // The Testing Lab and Telemetry only open after the details were asked in this visit.
  function gate(mode) {
    try { if (sessionStorage.getItem('garuda-ok-' + mode) === '1') return true; } catch (e) { return true; }
    location.replace('index.html?go=' + mode);
    return false;
  }
  function pass(mode) { try { sessionStorage.setItem('garuda-ok-' + mode, '1'); } catch (e) { /* ignore */ } }

  function stamp() {
    const d = new Date(), p = x => String(x).padStart(2, '0');
    return { date: `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`, time: `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}` };
  }
  function slug(s) { return (s || 'CubeSat').trim().replace(/[^A-Za-z0-9_-]+/g, '-').replace(/^-+|-+$/g, '') || 'CubeSat'; }
  function fileName(base, ext) { const t = stamp(); return `${slug(names().cubesat)}_${base}_${t.date}_${t.time.replace(/:/g, '')}.${ext}`; }
  function meta() {
    const n = names(), t = stamp(), d = v => v || '(not given)';
    return { team: d(n.team), team_id: d(n.teamId), lab: d(n.lab), cubesat: n.cubesat || '(not named yet)', operator: d(n.operator), exported_date: t.date, exported_time: t.time, app: 'G.A.R.U.D.A.' };
  }
  const HEAD = m => [['Team', m.team], ['Team ID', m.team_id], ['Testing lab', m.lab], ['CubeSat', m.cubesat], ['Operator', m.operator], ['Exported date', m.exported_date], ['Exported time', m.exported_time]];
  const csvCell = v => { const s = v == null ? '' : String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  function toCsv(rows) { if (!rows.length) return ''; const h = Object.keys(rows[0]); return [h.join(','), ...rows.map(r => h.map(k => csvCell(r[k])).join(','))].join('\n'); }
  function saveBlob(name, text, mime) {
    const url = URL.createObjectURL(new Blob([text], { type: mime || 'text/plain' }));
    const a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  // Every CSV, JSON and log download goes through here, so the details are always inside the file too.
  function download(filename, content, mime) {
    const ext = (filename.split('.').pop() || 'txt').toLowerCase(), base = filename.replace(/\.[^.]+$/, ''), m = meta();
    let out;
    if (ext === 'csv') out = HEAD(m).concat([[]]).map(r => r.map(csvCell).join(',')).join('\n') + '\n' + content;
    else if (ext === 'json') { let data = content; try { data = JSON.parse(content); } catch (e) { /* keep text */ } out = JSON.stringify({ exported: m, data }, null, 2); }
    else out = HEAD(m).map(r => '# ' + r.join(': ')).join('\n') + '\n' + content;
    saveBlob(fileName(base, ext), out, mime);
  }
  // A real .xlsx Excel file, with the details above the table.
  function exportXlsx(rows, base) {
    if (!g.XLSX) { download(base + '.csv', toCsv(rows), 'text/csv'); return; }
    const m = meta(), headers = rows.length ? Object.keys(rows[0]) : ['(no data yet)'];
    const aoa = HEAD(m).concat([[], headers]);
    rows.forEach(r => aoa.push(headers.map(h => (r[h] == null ? '' : r[h]))));
    const wb = g.XLSX.utils.book_new();
    g.XLSX.utils.book_append_sheet(wb, g.XLSX.utils.aoa_to_sheet(aoa), 'Data');
    g.XLSX.writeFile(wb, fileName(base, 'xlsx'));
  }
  g.GARUDA = { names, setNames, formHtml, readForm, gate, pass, stamp, meta, download, exportXlsx, toCsv, fileName, slug };
})(window);
