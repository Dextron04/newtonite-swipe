// Newtonite Swipe: spreadsheet -> column roles -> swipeable candidate cards. No build step, no backend.
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const store = {
  get: (k, d) => { try { return JSON.parse(localStorage.getItem('ns:' + k)) ?? d; } catch { return d; } },
  set: (k, v) => { try { localStorage.setItem('ns:' + k, JSON.stringify(v)); } catch {} },
};

const ROLES = {
  name: 'Full name', email: 'Email', phone: 'Phone', college: 'College', course: 'Course', year: 'Graduation year',
  ghuser: 'GitHub username', repo: 'Repository URL', live: 'Live URL', stack: 'Tech stack', writeup: 'Engineering write-up',
  confirm: 'Repo confirmation', checklist: 'Submission checklist', time: 'Timestamp', other: 'Other (ignored)',
};
const THRESH = 110;
const DECISIONS = { accept: 'Accepted', reject: 'Rejected', maybe: 'Maybe' };

let columns = [], rows = [], cands = [], cursor = 0, allChecks = [];
let decisions = store.get('decisions', {});
let filter = 'all';

/* ---------- 1. column analysis ---------- */
function classify(header, vals) {
  const h = header.toLowerCase(), v = vals.filter(x => x != null && x !== '');
  const frac = fn => v.length ? v.filter(fn).length / v.length : 0;
  const avg = v.length ? v.reduce((a, x) => a + String(x).length, 0) / v.length : 0;
  if (frac(x => x instanceof Date) > .8 || /timestamp/.test(h)) return 'time';
  if (/e-?mail/.test(h) || frac(x => /^[^@\s]+@[^@\s]+$/.test(String(x).trim())) > .8) return 'email';
  if (/phone|mobile|contact/.test(h)) return 'phone';
  if (/github.*user|username/.test(h)) return 'ghuser';
  if (/repo.*url|repository url/.test(h) || (frac(x => /github\.com\/[^/\s]+\/[^/\s]+/.test(x)) > .6 && !/live/.test(h))) return 'repo';
  if (/live|deploy|demo/.test(h)) return 'live';
  if (/tech|stack/.test(h)) return 'stack';
  if (/checklist/.test(h)) return 'checklist';
  if (/confirm/.test(h)) return 'confirm';
  if (/decision|describe|explain|why|summary/.test(h) || avg > 200) return 'writeup';
  if (/college|institution|universit|school/.test(h)) return 'college';
  if (/course|special|branch|degree|major/.test(h)) return 'course';
  if (/year|batch|graduat/.test(h)) return 'year';
  if (/name/.test(h)) return 'name';
  return 'other';
}

function loadSheet(buf) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames.find(n => wb.Sheets[n]['!ref'] && !/^A1:A1$/.test(wb.Sheets[n]['!ref'])) || wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }).filter(r => r.some(c => c != null && c !== ''));
  if (grid.length < 2) throw new Error('No data rows found in that file.');
  const heads = grid[0].map(h => String(h ?? '').trim());
  rows = grid.slice(1);
  columns = heads.map((h, i) => ({ i, header: h || `Column ${i + 1}`, role: classify(h, rows.map(r => r[i])), sample: rows.map(r => r[i]).find(x => x != null && x !== '') }));
  renderColumns();
  go(['review', 'board'].includes(location.hash.slice(1)) ? location.hash.slice(1) : 'columns');
}

function renderColumns() {
  $('#cols').innerHTML = columns.map((c, n) => {
    const s = c.sample instanceof Date ? c.sample.toISOString().slice(0, 16).replace('T', ' ') : String(c.sample ?? '').replace(/\s+/g, ' ');
    return `<div class="col" style="--i:${n}"><i>${String(n + 1).padStart(2, '0')}</i><b>${esc(c.header)}</b><span title="${esc(s)}">${esc(s.slice(0, 90)) || '—'}</span>
      <select data-i="${n}">${Object.entries(ROLES).map(([k, l]) => `<option value="${k}"${k === c.role ? ' selected' : ''}>${l}</option>`).join('')}</select></div>`;
  }).join('');
}

/* ---------- 2. candidates ---------- */
const urls = s => String(s ?? '').match(/https?:\/\/[^\s,;]+/g) || [];
function parseRepo(u) {
  const m = String(u ?? '').match(/github\.com\/([^/\s]+)(?:\/([^/\s?#]+))?/i);
  return m ? { owner: m[1], repo: m[2] ? m[2].replace(/\.git$/, '') : null } : null;
}
function buildCands() {
  const col = r => columns.find(c => c.role === r)?.i;
  const idx = Object.fromEntries(Object.keys(ROLES).map(r => [r, col(r)]));
  const get = (row, r) => { const v = idx[r] == null ? null : row[idx[r]]; return v == null ? '' : v instanceof Date ? v : String(v).trim(); };
  const by = new Map();
  rows.forEach(row => {
    const name = String(get(row, 'name')).replace(/\s+/g, ' '), email = String(get(row, 'email')).toLowerCase();
    const id = email || name.toLowerCase();
    const gh = parseRepo(get(row, 'repo'));
    const c = {
      id, name: name || '(no name)', email, phone: get(row, 'phone'), college: String(get(row, 'college')).replace(/\s+/g, ' '),
      course: get(row, 'course'), year: get(row, 'year'), time: get(row, 'time'),
      ghuser: String(get(row, 'ghuser')).replace(/^.*github\.com\//i, '').replace(/^@/, '') || gh?.owner || '',
      repoUrl: urls(get(row, 'repo'))[0] || '', gh, live: urls(get(row, 'live')),
      stack: String(get(row, 'stack')).split(/\n|,|;|\|/).map(s => s.trim()).filter(Boolean),
      writeup: String(get(row, 'writeup')), confirm: get(row, 'confirm'),
      checks: String(get(row, 'checklist')).split(/\.\s*,\s*/).map(s => s.replace(/\.\s*$/, '').trim()).filter(Boolean),
      submissions: 1,
    };
    const prev = by.get(id);
    if (prev) { c.submissions = prev.submissions + 1; if (!(c.time >= prev.time)) { prev.submissions = c.submissions; return; } }
    by.set(id, c);
  });
  cands = [...by.values()];
  allChecks = [...new Set(cands.flatMap(c => c.checks))];
  cursor = Math.max(0, cands.findIndex(c => !decisions[c.id]));
  if (cands.every(c => decisions[c.id])) cursor = cands.length;
}

/* ---------- 3. GitHub enrichment (public API, cached) ---------- */
const ghCache = new Map();
async function ghGet(path) {
  const token = store.get('token', '');
  const r = await fetch('https://api.github.com' + path, { headers: { Accept: 'application/vnd.github+json', ...(token && { Authorization: 'Bearer ' + token }) } });
  if (r.status === 403 || r.status === 429) { $('#token-note').hidden = false; throw new Error('rate'); }
  if (!r.ok) return null;
  return r.json();
}
function ghData(c) {
  const key = (c.ghuser + '/' + (c.gh?.repo || '')).toLowerCase();
  if (ghCache.has(key)) return ghCache.get(key);
  const cached = store.get('gh:' + key);
  const p = cached && Date.now() - cached.t < 12 * 36e5 ? Promise.resolve(cached.d) : (async () => {
    const [user, repo] = await Promise.all([c.ghuser ? ghGet('/users/' + c.ghuser) : null, c.gh?.repo ? ghGet(`/repos/${c.gh.owner}/${c.gh.repo}`) : null]);
    const tree = repo ? await ghGet(`/repos/${c.gh.owner}/${c.gh.repo}/git/trees/${repo.default_branch}?recursive=1`) : null;
    const d = { user, repo, files: tree?.tree?.map(t => t.path) ?? null, truncated: !!tree?.truncated };
    store.set('gh:' + key, { t: Date.now(), d });
    return d;
  })();
  ghCache.set(key, p);
  p.catch(() => ghCache.delete(key));
  return p;
}
const ago = iso => { const d = (Date.now() - new Date(iso)) / 864e5; return d < 1 ? 'today' : d < 30 ? `${Math.round(d)}d ago` : d < 365 ? `${Math.round(d / 30)}mo ago` : `${(d / 365).toFixed(1)}y ago`; };

function ghHtml(c, d) {
  if (!c.ghuser) return `<p class="note">No GitHub username given.</p>`;
  const u = d?.user, r = d?.repo;
  let h = `<div class="gh-user">${u ? `<img src="${esc(u.avatar_url)}&s=88" alt="">` : ''}<div><b>${esc(u.name || c.ghuser)}</b><small><a href="https://github.com/${esc(c.ghuser)}" target="_blank" rel="noopener">@${esc(c.ghuser)}</a>${u ? ` · joined ${new Date(u.created_at).getFullYear()}` : ' · profile not found'}</small></div></div>`;
  if (u) h += `<div class="stats"><div><b>${u.public_repos}</b><small>Repos</small></div><div><b>${u.followers}</b><small>Followers</small></div><div><b>${r ? r.stargazers_count : '–'}</b><small>Stars</small></div><div><b>${r ? Math.round(r.size / 1024 * 10) / 10 : '–'}</b><small>MB</small></div></div>`;
  h += `<div class="graph"><img src="https://ghchart.rshah.org/2E5BFF/${encodeURIComponent(c.ghuser)}" alt="Contribution graph" onload="this.classList.add('loaded')" onerror="this.parentNode.innerHTML='<span class=note>Contribution graph unavailable</span>'"></div>`;
  if (c.gh?.repo) {
    if (!r) h += `<div class="signals"><span class="sig n">repo not reachable / private</span></div>`;
    else {
      const f = d.files, has = re => f?.some(p => re.test(p));
      const sig = (ok, t) => `<span class="sig ${ok ? 'y' : 'n'}">${t}</span>`;
      h += `<div class="signals">${r.language ? `<span class="sig w">${esc(r.language)}</span>` : ''}<span class="sig w">pushed ${ago(r.pushed_at)}</span>${f ? `<span class="sig w">${f.length}${d.truncated ? '+' : ''} files</span>
        ${sig(has(/^readme/i), 'README')}${sig(has(/(^|\/)(tests?|__tests__|spec)\/|\.(test|spec)\.[jt]sx?$|test_.*\.py$/i), 'tests')}${sig(has(/decisions?|adr/i), 'decisions doc')}${sig(has(/dockerfile|compose\.ya?ml/i), 'docker')}${sig(has(/^\.github\/workflows\//), 'CI')}` : ''}</div>`;
    }
  }
  return h;
}

/* ---------- 4. cards ---------- */
const cardEls = new Map();
function cardHtml(c) {
  const links = [
    c.repoUrl && `<a class="chip" href="${esc(c.repoUrl)}" target="_blank" rel="noopener">Repository ↗</a>`,
    ...c.live.map((u, i) => `<a class="chip" href="${esc(u)}" target="_blank" rel="noopener">${c.live.length > 1 ? `Live ${i + 1}` : 'Live demo'} ↗</a>`),
    c.ghuser && `<a class="chip" href="https://github.com/${esc(c.ghuser)}" target="_blank" rel="noopener">GitHub ↗</a>`,
    c.email && `<a class="chip" href="mailto:${esc(c.email)}">${esc(c.email)}</a>`,
    c.phone && `<a class="chip" href="tel:${esc(c.phone)}">${esc(c.phone)}</a>`,
  ].filter(Boolean).join('');
  const missing = allChecks.filter(x => !c.checks.includes(x));
  return `<div class="stamp stamp--ok">Accept</div><div class="stamp stamp--bad">Reject</div><div class="stamp stamp--warn">Maybe</div>
  <header class="card-head"><span class="meta">${[c.year && `Batch ${esc(c.year)}`, c.course && esc(c.course)].filter(Boolean).join(' · ')}</span>
    <h2>${esc(c.name)}</h2><span class="sub">${esc(c.college)}</span>${c.submissions > 1 ? `<span class="dupe">Submitted ${c.submissions}× · latest shown</span>` : ''}</header>
  <div class="card-body">
    <div class="pane"><div><span class="lbl">Links</span><div class="links" data-nodrag>${links || '<span class="note">None</span>'}</div></div>
      <div><span class="lbl">GitHub</span><div class="gh" data-gh data-nodrag><div class="skel"></div><div class="skel" style="width:60%"></div></div></div></div>
    <div class="pane"><div><span class="lbl">Stack</span><div class="links">${c.stack.slice(0, 16).map(s => `<span class="chip chip--sm">${esc(s.length > 44 ? s.slice(0, 42) + '…' : s)}</span>`).join('') || '<span class="note">None</span>'}</div></div>
      <div><span class="lbl">Engineering decisions</span><div class="writeup">${esc(c.writeup) || '<span class="note">None given</span>'}</div></div>
      <div><span class="lbl">Checklist · ${c.checks.length}/${allChecks.length} ticked</span><ul class="ticks">${c.checks.map(x => `<li>${esc(x)}</li>`).join('')}${missing.map(x => `<li class="no">${esc(x)}</li>`).join('')}</ul></div></div>
  </div>`;
}

function syncStack() {
  const want = cands.slice(cursor, cursor + 3);
  const stack = $('#stack');
  for (const [id, el] of cardEls) if (!want.some(c => c.id === id) && !el.classList.contains('out')) { el.remove(); cardEls.delete(id); }
  want.forEach((c, depth) => {
    let el = cardEls.get(c.id);
    if (!el) {
      el = document.createElement('article'); el.className = 'card'; el.innerHTML = cardHtml(c); el.dataset.id = c.id;
      if (depth === 0) el.classList.add('enter');
      cardEls.set(c.id, el);
    }
    el.dataset.depth = depth; el.style.zIndex = 3 - depth;
    if (decisions[c.id] && depth === 0) { el.dataset.d = decisions[c.id]; el.style.setProperty('--tint', .4); }
    else { delete el.dataset.d; el.style.setProperty('--tint', 0); }
    stack.append(el);
    if (depth === 0 && !el.dataset.loaded) { el.dataset.loaded = 1; fillGh(el, c); }
  });
  $('#empty').hidden = want.length > 0;
  $('.deck-bar').style.visibility = $('.keys').style.visibility = want.length ? '' : 'hidden';
  $('#undo').hidden = cursor === 0;
  tally();
}
async function fillGh(el, c) {
  const box = $('[data-gh]', el);
  try { box.innerHTML = ghHtml(c, await ghData(c)); }
  catch { box.innerHTML = `<p class="note">GitHub details unavailable (rate limit). ${c.ghuser ? `<a href="https://github.com/${esc(c.ghuser)}" target="_blank" rel="noopener">Open profile</a>` : ''}</p>`; delete el.dataset.loaded; }
}

function decide(d) {
  const c = cands[cursor], el = c && cardEls.get(c.id);
  if (!el || el.classList.contains('out')) return;
  decisions[c.id] = d; store.set('decisions', decisions);
  el.dataset.d = d; el.style.setProperty('--tint', 1);
  $$('.stamp', el).forEach(s => s.style.opacity = '');
  const W = innerWidth * 1.2;
  el.style.setProperty('--dx', d === 'accept' ? W + 'px' : d === 'reject' ? -W + 'px' : (el.style.getPropertyValue('--dx') || '0px'));
  el.style.setProperty('--dy', d === 'maybe' ? innerHeight + 'px' : el.style.getPropertyValue('--dy') || '0px');
  el.style.setProperty('--rot', d === 'accept' ? '22deg' : d === 'reject' ? '-22deg' : '0deg');
  el.classList.remove('drag', 'enter'); el.classList.add('out');
  cursor++; $('#stack').style.setProperty('--lift', 0);
  syncStack();
  setTimeout(() => { el.remove(); cardEls.delete(c.id); }, 400);
}
function undo() {
  if (cursor === 0) return;
  cursor--; delete decisions[cands[cursor].id]; store.set('decisions', decisions);
  const old = cardEls.get(cands[cursor].id); if (old) { old.remove(); cardEls.delete(old.dataset.id); }
  syncStack();
}

/* drag */
let drag = null;
$('#stack').addEventListener('pointerdown', e => {
  const el = e.target.closest('.card'); if (!el || el.dataset.depth !== '0' || e.target.closest('a,button,.writeup,[data-nodrag] img')) return;
  drag = { el, x: e.clientX, y: e.clientY, on: false, id: e.pointerId };
});
addEventListener('pointermove', e => {
  if (!drag || e.pointerId !== drag.id) return;
  const dx = e.clientX - drag.x, dy = Math.max(0, e.clientY - drag.y);
  if (!drag.on) { if (Math.hypot(dx, dy) < 6) return; drag.on = true; drag.el.classList.add('drag'); drag.el.classList.remove('enter'); drag.el.setPointerCapture?.(e.pointerId); getSelection()?.removeAllRanges(); }
  const down = dy > Math.abs(dx) * 1.2, dir = down ? 'maybe' : dx > 0 ? 'accept' : 'reject';
  const tint = Math.min(1, (down ? dy : Math.abs(dx)) / THRESH);
  const s = drag.el.style; s.setProperty('--dx', dx + 'px'); s.setProperty('--dy', dy * .6 + 'px'); s.setProperty('--rot', dx / 18 + 'deg'); s.setProperty('--tint', tint);
  drag.el.dataset.d = dir; drag.dir = dir; drag.tint = tint;
  $$('.stamp', drag.el).forEach(st => st.style.opacity = st.classList.contains({ accept: 'stamp--ok', reject: 'stamp--bad', maybe: 'stamp--warn' }[dir]) ? Math.min(1, tint * 1.3) : 0);
  $('#stack').style.setProperty('--lift', tint);
});
const release = e => {
  if (!drag || e.pointerId !== drag.id) return;
  const { el, on, dir, tint } = drag; drag = null;
  if (!on) return;
  el.classList.remove('drag');
  if (tint >= 1) return decide(dir);
  ['--dx', '--dy', '--rot', '--tint'].forEach((p, i) => el.style.setProperty(p, i < 2 ? '0px' : i === 2 ? '0deg' : 0));
  $$('.stamp', el).forEach(s => s.style.opacity = ''); $('#stack').style.setProperty('--lift', 0);
  if (decisions[cands[cursor]?.id]) el.style.setProperty('--tint', .4); else delete el.dataset.d;
};
addEventListener('pointerup', release); addEventListener('pointercancel', release);

$('.deck-bar').addEventListener('click', e => { const b = e.target.closest('[data-act]'); if (b) decide(b.dataset.act); });
$('#undo').onclick = undo;
addEventListener('keydown', e => {
  if (currentView !== 'review' || e.target.matches?.('input,select,textarea') || e.metaKey || e.ctrlKey) return;
  const k = { ArrowRight: 'accept', ArrowLeft: 'reject', ArrowDown: 'maybe' }[e.key];
  if (k) { e.preventDefault(); decide(k); } else if (e.key.toLowerCase() === 'z') undo();
});

/* ---------- 5. board ---------- */
function renderBoard() {
  const counts = { all: cands.length, accept: 0, maybe: 0, reject: 0, none: 0 };
  cands.forEach(c => counts[decisions[c.id] || 'none']++);
  $('#filters').innerHTML = [['all', 'All'], ['accept', 'Accepted'], ['maybe', 'Maybe'], ['reject', 'Rejected'], ['none', 'Pending']].map(([k, l]) => `<button data-f="${k}" class="${filter === k ? 'on' : ''}">${l} ${counts[k]}</button>`).join('');
  $('#rows').innerHTML = cands.map((c, i) => [c, i]).filter(([c]) => filter === 'all' || (decisions[c.id] || 'none') === filter).map(([c, i], n) => {
    const d = decisions[c.id] || 'none';
    return `<div class="row" style="--i:${n}" data-d="${d}" data-i="${i}"><span class="st">${DECISIONS[d] || 'Pending'}</span><div><b>${esc(c.name)}</b><br><small>${esc(c.college)}</small></div>
      <small>${esc(c.repoUrl.replace(/^https?:\/\/(www\.)?github\.com\//, ''))}</small>
      <div class="mini"><button data-set="reject" title="Reject">✕</button><button data-set="maybe" title="Maybe">↓</button><button data-set="accept" title="Accept">✓</button><button data-open title="Open card">↗</button></div></div>`;
  }).join('') || '<p class="muted" style="padding:24px 0">Nothing here yet.</p>';
}
$('#filters').addEventListener('click', e => { const b = e.target.closest('[data-f]'); if (b) { filter = b.dataset.f; renderBoard(); } });
$('#rows').addEventListener('click', e => {
  const row = e.target.closest('.row'); if (!row) return;
  const i = +row.dataset.i, c = cands[i];
  const set = e.target.closest('[data-set]')?.dataset.set;
  if (set) { decisions[c.id] = set; store.set('decisions', decisions); renderBoard(); tally(); }
  else if (e.target.closest('[data-open]')) { cursor = i; $$('.card').forEach(x => x.remove()); cardEls.clear(); go('review'); }
});
$('#export').onclick = () => {
  const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csv = ['Decision,Name,Email,Phone,College,Year,Repo,Live'].concat(cands.map(c => [decisions[c.id] || 'pending', c.name, c.email, c.phone, c.college, c.year, c.repoUrl, c.live.join(' ')].map(q).join(','))).join('\n');
  const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(new Blob([csv], { type: 'text/csv' })), download: 'decisions.csv' }); a.click();
};
$('#reset').onclick = () => { if (confirm('Clear all decisions?')) { decisions = {}; store.set('decisions', decisions); cursor = 0; $$('.card').forEach(x => x.remove()); cardEls.clear(); renderBoard(); tally(); } };

/* ---------- 6. routing + wiring ---------- */
let currentView = 'load';
function go(v) { location.hash = v === 'load' ? '' : v; show(v); }
function show(v) {
  if (v !== 'load' && !columns.length) v = 'load';
  if (v === 'review' && !cands.length) buildCands();
  currentView = v;
  $$('.view').forEach(s => s.hidden = s.id !== 'v-' + v);
  $('#nav').hidden = v === 'load';
  $$('#nav a').forEach(a => a.classList.toggle('on', a.dataset.view === v));
  if (v === 'review') syncStack();
  if (v === 'board') { if (!cands.length) buildCands(); renderBoard(); }
  tally();
}
addEventListener('hashchange', () => show(location.hash.slice(1) || 'load'));
function tally() {
  if (!cands.length) { $('#tally').textContent = ''; return; }
  const n = k => cands.filter(c => decisions[c.id] === k).length;
  $('#tally').textContent = `✓ ${n('accept')}  ~ ${n('maybe')}  ✕ ${n('reject')}  · ${cands.filter(c => !decisions[c.id]).length} left`;
}
$('#cols').addEventListener('change', e => { columns[+e.target.dataset.i].role = e.target.value; cands = []; });
$('#start').onclick = () => { buildCands(); go('review'); };
$('#token-save').onclick = () => { store.set('token', $('#token').value.trim()); $('#token-note').hidden = true; ghCache.clear(); $$('.card').forEach(x => x.remove()); cardEls.clear(); syncStack(); };

const drop = $('#drop'), fileIn = $('#file');
const readFile = f => f.arrayBuffer().then(loadSheet).catch(err => { $('#load-err').textContent = err.message; });
fileIn.onchange = () => fileIn.files[0] && readFile(fileIn.files[0]);
['dragover', 'dragenter'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => e.dataTransfer.files[0] && readFile(e.dataTransfer.files[0]));

// Convenience: a local data.xlsx next to index.html loads automatically (gitignored).
fetch('data.xlsx').then(r => r.ok ? r.arrayBuffer() : Promise.reject()).then(loadSheet).catch(() => show('load'));
