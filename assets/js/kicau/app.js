/*
 * Kicau: Rasa Kenari's chat app, where each Excel lesson arrives as a request from a colleague.
 *
 * One page, plain hash routes (artifact frames pass only #name): #desk, #chat-<person>, #notes, #career, #you.
 * Progress lives in this browser only (localStorage "kicau.v1").
 * Answers are judged by core/check.js on two sheets, with the real formula engine, and every verdict
 * comes back as something the colleague would say. ?fast skips the typing delays (for tests),
 * ?reset starts over.
 */
import { load, evalOn, evalFill, evalSpill } from '../core/engine.js';

const { PEOPLE, avatar } = window.KicauPeople;
const MS = window.KicauMissions;
const Check = window.KenariCheck;
const VBA = window.KenariVBA;
const LIST = MS.LIST;
const byId = Object.fromEntries(LIST.map((m) => [m.id, m]));

const $ = (s, r = document) => r.querySelector(s);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const Q = new URLSearchParams(location.search);
const FAST = Q.has('fast');
const wait = (ms) => new Promise((r) => setTimeout(r, FAST ? 0 : ms));
const LAP = matchMedia('(min-width:1000px)');
const coarse = matchMedia('(pointer:coarse)').matches;
const I = (id, cls) => `<svg class="ic${cls ? ' ' + cls : ''}" aria-hidden="true"><use href="#i-${id}"/></svg>`;
const fmt = (v) => (typeof v === 'number' ? v.toLocaleString('en-US', { maximumFractionDigits: 2 }) : String(v));
const pad = (n) => String(n).padStart(2, '0');
const clock = () => { const d = new Date(); return pad(d.getHours()) + ':' + pad(d.getMinutes()); };
const localDay = (d = new Date()) => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const dateText = (serial) => { const d = new Date(Date.UTC(1899, 11, 30) + serial * 864e5); return pad(d.getUTCDate()) + ' ' + MON[d.getUTCMonth()]; };
const strip = (h) => String(h).replace(/<[^>]+>/g, '').replace(/&amp;/g, '&');

/* ---------- Saved progress ---------- */
const KEY = 'kicau.v1';
const fresh = () => ({ v: 1, name: '', mood: '', started: '', logs: {}, missions: {}, closed: [], coffee: { day: '', used: 0 }, days: [], opened: {}, promoted: false, practice: { solved: {}, tried: {} } });
let ST;
try { ST = Object.assign(fresh(), JSON.parse(localStorage.getItem(KEY) || 'null') || {}); } catch (e) { ST = fresh(); }
if (Q.has('reset')) { ST = fresh(); try { localStorage.removeItem(KEY); } catch (e) { /* private mode */ } }
const save = () => { try { localStorage.setItem(KEY, JSON.stringify(ST)); } catch (e) { /* private mode: progress lasts this visit */ } };
const MST = (id) => (ST.missions[id] = ST.missions[id] || { attempts: 0, coffees: 0, hint: 0, draft: '', started: false, solved: null });
const logOf = (p) => (ST.logs[p] = ST.logs[p] || []);
const PR = () => (ST.practice = ST.practice && ST.practice.solved ? ST.practice : { solved: {}, tried: {} });
/* The formula being typed belongs to a practice problem when one is open, otherwise to the current request. */
const activeM = () => UI.prac || current();

/* ---------- Progress helpers ---------- */
const isClosed = (m) => ST.closed.includes(m.id);
const unlocked = (m) => { const i = LIST.indexOf(m); return i === 0 || isClosed(LIST[i - 1]); };
const current = () => LIST.find((m) => unlocked(m) && !isClosed(m));
const activeFor = (p) => { const c = current(); return c && c.from === p ? c : null; };
const rankOf = (n) => MS.RANKS.reduce((r, k) => (n >= k.at ? k : r), MS.RANKS[0]);
const nextRank = (n) => MS.RANKS.find((k) => k.at > n);
function coffeesLeft() { if (ST.coffee.day !== localDay()) ST.coffee = { day: localDay(), used: 0 }; return Math.max(0, 3 - ST.coffee.used); }
function markDay() { const d = localDay(); if (!ST.days.includes(d)) ST.days.push(d); }
function streak() {
  const set = new Set(ST.days);
  let n = 0;
  const d = new Date();
  if (!set.has(localDay(d))) d.setDate(d.getDate() - 1);
  while (set.has(localDay(d))) { n++; d.setDate(d.getDate() - 1); }
  return n;
}
const dayNumber = () => { if (!ST.started) return 1; return Math.max(1, Math.round((Date.parse(localDay()) - Date.parse(ST.started)) / 864e5) + 1); };
function fill(s, extra) {
  const vals = Object.assign({ name: esc(ST.name || 'kamu'), target: esc(MS.target), other: esc(MS.other), altD5: MS.ALT_D5 }, extra || {});
  return String(s).replace(/\{(\w+)\}/g, (_, k) => (vals[k] === undefined ? '' : vals[k]));
}
function greeting() { const h = new Date().getHours(); return h < 11 ? 'Selamat pagi' : h < 15 ? 'Selamat siang' : h < 18 ? 'Selamat sore' : 'Selamat malam'; }

/* ---------- The formula engine ---------- */
const ENG = { api: null, failed: false, ex: {} };
/* The engine (10 MB) waits until the opening ends or is skipped, so it never competes with the animation for the
   CPU. Nobody needs it before the welcome screen, and it loads while they type their name. */
let engineGo;
const engineGate = new Promise((r) => { engineGo = r; });
const enginePromise = engineGate.then(load).then((api) => {
  ENG.api = api;
  document.documentElement.dataset.engine = 'ready';
  document.documentElement.dataset.engineMs = api.loadMs;
  document.querySelectorAll('[data-exres]').forEach(paintExample);
  const e = $('#engmsg'); if (e) e.hidden = true;
  if (UI.pendingSend) { UI.pendingSend = false; send(); }
  return api;
}).catch((err) => {
  ENG.failed = true;
  document.documentElement.dataset.engine = 'failed';
  const e = $('#engmsg');
  if (e) { e.hidden = false; e.textContent = 'The practice sheet did not load, which can happen on a strict office network. Reload to try again.'; }
  console.warn('Engine load failed', err);
});
function paintExample(el) {
  /* A lesson's example result is worked out the first time the lesson shows, not for all 97 at start-up. */
  const id = el.dataset.exres;
  if (ENG.ex[id] === undefined && ENG.api && byId[id] && byId[id].lesson && byId[id].lesson.exFormula) ENG.ex[id] = evalOn(ENG.api, byId[id].cells, 'Z1', byId[id].lesson.exFormula);
  const v = ENG.ex[id];
  const L = byId[el.dataset.exres].lesson;
  if (v !== undefined) el.innerHTML = `Its result: <b>${esc(L.exDate && typeof v === 'number' ? dateText(v) : fmt(v))}</b>${L.exLabel ? ' ' + esc(L.exLabel) : ''}.`;
}

/* ---------- Screen state ---------- */
const UI = { route: '', person: null, busy: {}, dock: 'open', pendingSend: false, promo: false, filter: '' };
const app = $('#app');
const live = document.createElement('div');
live.className = 'sr';
live.setAttribute('aria-live', 'polite');
document.body.append(live);
const say = (t) => { live.textContent = strip(t); };

/* ---------- Shell pieces ---------- */
function threadList() {
  const out = [];
  const cur = current();
  const seenP = new Set();
  const people = [];
  if (cur) people.push(cur.from);
  for (const m of LIST.slice().reverse()) if (unlocked(m) && !people.includes(m.from)) people.push(m.from);
  people.push('sari');
  for (const p of people) {
    if (seenP.has(p)) continue;
    seenP.add(p);
    const log = logOf(p);
    const act = activeFor(p);
    const unread = act && !MST(act.id).started;
    let pv = '', t = '';
    if (unread) { pv = fill(act.intro[0]); t = act.due; }
    else if (log.length) { const last = [...log].reverse().find((e) => e.k === 'text' || e.k === 'answer' || e.k === 'closed'); if (last) { pv = last.k === 'answer' ? 'You: ' + last.formula : last.k === 'closed' ? '✓ Request ' + last.mission + ' closed' : (last.who === 'me' ? 'You: ' : '') + last.html; t = last.t || ''; } }
    out.push({ p, pv: strip(pv), t, unread, typing: !!UI.busy[p] });
  }
  if (ST.promoted) for (const p of ['citra', 'dimas']) if (!seenP.has(p)) out.push({ p, pv: strip(fill((logOf(p).slice(-1)[0] || {}).html || '')), t: (logOf(p).slice(-1)[0] || {}).t || '', unread: !ST.opened[p], typing: false });
  return out;
}
function rowHTML(r) {
  const P = PEOPLE[r.p];
  return `<a class="row${UI.person === r.p ? ' on' : ''}${r.unread ? ' unread' : ''}" href="#chat-${r.p}">${avatar(r.p, 42, { on: true })}<div class="body"><b>${esc(P.name)}${P.bank ? ` <span class="ext">${I('bank')}Bank</span>` : ''}<small>${esc(r.t)}</small></b><span class="pv">${r.typing ? `<span class="typing">typing<span class="dots"><i></i><i></i><i></i></span></span>` : esc(r.pv || P.role)}</span></div>${r.unread ? '<span class="ub">1</span>' : ''}</a>`;
}
function bankLockRow() {
  if (ST.promoted) return '';
  return `<div class="row lk">${avatar('citra', 42, { dim: true })}<div class="body"><b>Citra <span class="ext">${I('bank')}Bank Kenari</span></b><span class="pv">Opens when you reach Analyst</span></div>${I('lock')}</div>`;
}
const unreadCount = () => threadList().filter((r) => r.unread).length;
function rail(on) {
  const n = unreadCount();
  const a = (href, id, label, key) => `<a href="${href}" class="${on === key ? 'on' : ''}" aria-label="${label}" title="${label}">${I(id)}${key === 'chats' && n ? `<i class="n">${n}</i>` : ''}</a>`;
  return `<nav class="rail" aria-label="Kicau"><span class="logo" style="color:var(--brand)">${I('bird')}</span>${a('#desk', 'chat', 'Chats', 'chats')}${a('#practice', 'target', 'Practice', 'practice')}${a('#notes', 'note', 'Notes', 'notes')}${a('#career', 'brief', 'Career', 'career')}${a('#you', 'user', 'You', 'me')}<a class="me" href="#you" aria-label="You">${avatar('me', 38)}</a></nav>`;
}
function tabs(on) {
  const n = unreadCount();
  const a = (href, id, label, key) => `<a href="${href}" class="${on === key ? 'on' : ''}">${I(id)}${label}${key === 'chats' && n ? `<i class="n">${n}</i>` : ''}</a>`;
  return `<nav class="tabs" aria-label="Kicau">${a('#desk', 'chat', 'Chats', 'chats')}${a('#practice', 'target', 'Practice', 'practice')}${a('#notes', 'note', 'Notes', 'notes')}${a('#career', 'brief', 'Career', 'career')}${a('#you', 'user', 'You', 'me')}</nav>`;
}
function side() {
  const n = ST.closed.length, nx = nextRank(n), rk = rankOf(n);
  const pct = nx ? Math.round(((n - rk.at) / (nx.at - rk.at)) * 100) : 100;
  return `<aside class="side"><div class="hd"><h2>Chats</h2><span class="sp"></span><span class="org">Rasa Kenari</span></div>
    <label class="search">${I('search')}<input id="tsearch" placeholder="Search people" value="${esc(UI.filter)}" aria-label="Search people"></label>
    <a class="minicar" href="#career"><div class="r"><span>${esc(rk.title)}</span><b>${nx ? `${n} of ${nx.at}` : n}</b></div><div class="bar"><i style="width:${pct}%"></i></div></a>
    <div class="list"><h2 class="sec">Direct messages</h2>${threadList().filter((r) => !UI.filter || PEOPLE[r.p].name.toLowerCase().includes(UI.filter.toLowerCase())).map(rowHTML).join('')}
    ${ST.promoted ? '' : `<h2 class="sec">Bank Kenari</h2>${bankLockRow()}`}</div></aside>`;
}
function shell(on, inner) {
  /* A practice problem works like a chat: full height on a phone (no bottom tabs), Practice lit in the rail. */
  const nav = on === 'prob' ? 'practice' : on;
  if (LAP.matches) return `<div class="lap${UI.max && (on === 'chat' || on === 'prob') ? ' max' : ''}">${rail(nav)}${side()}<main class="main">${inner}</main></div>`;
  return `<div class="ph">${inner}${on === 'chat' || on === 'prob' ? '' : tabs(nav)}</div>`;
}
function ring(n, of, cls) {
  const r = 26, c = 2 * Math.PI * r, p = of ? Math.min(1, n / of) : 1;
  return `<span class="ring ${cls || ''}"><svg viewBox="0 0 62 62"><circle cx="31" cy="31" r="${r}" fill="none" stroke="currentColor" stroke-opacity=".18" stroke-width="6"/><circle cx="31" cy="31" r="${r}" fill="none" stroke="currentColor" stroke-width="6" stroke-linecap="round" stroke-dasharray="${c * p} ${c}"/></svg><span class="t">${n}${of ? `<em>of ${of}</em>` : ''}</span></span>`;
}
function hero() {
  const n = ST.closed.length, rk = rankOf(n), nx = nextRank(n);
  const notes = ST.closed.length;
  return `<a class="hero" href="#career"><div class="top"><div style="flex:1;min-width:0"><small>Your role at Rasa Kenari</small><span class="title">${esc(rk.title)}</span><span class="team">Sales Finance team</span></div>${ring(n, nx ? nx.at : 0)}</div>
    <p class="next">${nx ? `${nx.at - n} more ${nx.at - n === 1 ? 'request' : 'requests'} to <b>${esc(nx.title)}</b>${nx.title === 'Analyst' ? ', when Bank Kenari starts writing to you.' : '.'}` : 'Top of the ladder.'}</p>
    <div class="stats"><span>${I('coffee')}<span><b>${coffeesLeft()}</b> coffees</span></span><span>${I('note')}<span><b>${notes}</b> notes</span></span><span>${I('flame')}<span>${streak() ? `<b>${streak()}</b>-day streak` : 'Streak starts today'}</span></span></div></a>`;
}
function replyCard() {
  const m = current();
  if (!m) return `<div class="empty">All caught up. ${ST.promoted ? 'Bank Kenari sends Chapter 2 soon.' : ''}</div>`;
  const P = PEOPLE[m.from];
  const started = MST(m.id).started;
  return `<a class="card" href="#chat-${m.from}">${avatar(m.from, 46, { on: true })}<div class="body"><div class="top"><b>${esc(P.name)}</b><small>${esc(m.due)}</small></div><div class="role">${esc(P.role)}</div><p class="pv${started ? '' : ' unread'}">${strip(fill(m.intro[started ? 1 : 0]))}</p><div class="tags"><span class="tag due">Due ${esc(m.due)}</span><span class="tag"><span class="xl${m.kind === 'macro' ? ' m' : ''}">X</span>${esc(m.file)}</span>${started ? '<span class="tag">In progress</span>' : ''}</div></div></a>`;
}
function laterCard() {
  if (ST.promoted) return `<a class="card" href="#chat-citra">${avatar('citra', 46, { on: true })}<div class="body"><div class="top"><b>Citra</b></div><div class="role">Head of Collection, <span class="ext">${I('bank')}Bank Kenari</span></div><p class="pv${ST.opened.citra ? '' : ' unread'}">${strip(fill(logOf('citra')[0] ? logOf('citra')[0].html : ''))}</p></div></a>`;
  const left = 8 - ST.closed.length;
  return `<div class="lock">${avatar('citra', 46, { dim: true })}<div class="body"><b>Bank Kenari wants to meet you</b>Citra, Head of Collection. Opens when you reach Analyst, ${left} ${left === 1 ? 'request' : 'requests'} from now.<div class="bar"><i style="width:${(ST.closed.length / 8) * 100}%"></i></div></div>${I('lock')}</div>`;
}

/* ---------- Home: the desk ---------- */
function renderHome() {
  UI.person = null;
  const hello = `<div class="hello"><small>${DAYS[new Date().getDay()]}, ${new Date().getDate()} ${MON[new Date().getMonth()]} · Day ${dayNumber()} at Rasa Kenari</small><h1>${greeting()}, ${esc(ST.name)}.</h1><p>${current() ? `${esc(PEOPLE[current().from].short)} needs you.` : 'Your inbox is clear.'}</p></div>`;
  const notes = ST.closed.map((id) => byId[id].note.fn).join(', ').split(', ').filter(Boolean);
  if (LAP.matches) {
    app.innerHTML = shell('chats', `<div class="desk">${hello}<div class="grid2"><div>${hero()}<h2 class="sec">Opens later</h2>${laterCard()}</div>
      <div><h2 class="sec">Needs your reply ${current() ? '<em>1</em>' : ''}</h2>${replyCard()}<h2 class="sec">Your notes</h2>${notes.length ? `<div class="notechips">${notes.map((n) => `<a class="chip" href="#notes">${esc(n)}</a>`).join('')}</div>` : '<div class="empty">Each request you close adds a note here.</div>'}</div></div></div>`);
  } else {
    app.innerHTML = shell('chats', `<header class="apphd"><span class="logo">${I('bird')}Kicau</span><span class="org">Rasa Kenari</span><span class="sp"></span><a class="icb" href="#you" aria-label="You">${avatar('me', 34)}</a></header>
      <div class="scroll">${hello}${hero()}<h2 class="sec">Needs your reply ${current() ? '<em>1</em>' : ''}</h2>${replyCard()}
      <h2 class="sec">All chats</h2><div class="card" style="display:block;padding:4px 0">${threadList().map(rowHTML).join('')}${bankLockRow()}</div>
      <h2 class="sec">Opens later</h2>${laterCard()}</div>`);
  }
  wireSide();
}
function wireSide() {
  const s = $('#tsearch');
  if (s) s.oninput = () => { UI.filter = s.value; const pos = s.selectionStart; render(); const n = $('#tsearch'); n.focus(); n.setSelectionRange(pos, pos); };
}

/* ---------- Chat ---------- */
function renderChat(p) {
  UI.person = p;
  const P = PEOPLE[p];
  if (!P) { location.hash = '#desk'; return; }
  if (P.bank && !ST.promoted) { location.hash = '#desk'; return; }
  ST.opened[p] = true;
  save();
  document.querySelectorAll('.toast').forEach((t) => { if (t.getAttribute('href') === '#chat-' + p) t.remove(); });
  /* After a request closes, its sheet stays up for review until the learner presses Continue. */
  const held = ST.hold ? LIST.find((x) => x.id === ST.hold && x.from === p) : null;
  const act = held ? null : activeFor(p);
  if (act && !MST(act.id).started) UI.dock = 'min';
  else if (UI.dock === 'min') UI.dock = 'open';
  const head =`<header class="chathd">${LAP.matches ? '' : `<a class="icb" href="#desk" aria-label="Back to chats">${I('back')}</a>`}${avatar(p, 40, { on: true })}<div class="who"><b>${esc(P.name)}</b><small class="${UI.busy[p] ? 'live' : ''}" id="whostate">${UI.busy[p] ? 'typing…' : esc(P.role + ', ' + P.org)}</small></div></header>`;
  const pin = act ? `<button class="pin" id="pin">${I('spark')}<span class="txt"><b>Request ${esc(act.id)}</b> ${strip(act.ask)}</span><span class="tag due">Due ${esc(act.due)}</span></button>` : '';
  const inner = `<div class="chatwrap"><section class="chatcol">${head}${pin}<div class="wall" id="wall" role="log" aria-label="Messages"></div>${!act && !held && LIST.some((m) => m.from === p) && !LAP.matches ? `<div class="foot" style="padding:10px 12px calc(12px + env(safe-area-inset-bottom));background:var(--surface);border-top:1px solid var(--line)"><a class="btn wide" href="${current() ? '#chat-' + current().from : '#desk'}">${current() ? `Next: ${esc(PEOPLE[current().from].short)}` : 'Back to chats'}</a></div>` : ''}</section>${act ? dockHTML(act) : held ? reviewHTML(held) : LAP.matches ? idlePanel(p) : ''}</div>`;
  app.innerHTML = shell('chat', inner);
  wireSide();
  if ($('#pin')) $('#pin').onclick = (e) => e.currentTarget.classList.toggle('open');
  renderWall();
  if (act) { wireDock(act); if (!LAP.matches) setDock(UI.dock); if (!MST(act.id).started) startMission(act); }
  if (held) wireReview(held);
  wireMax();
  mountMini(p);
}
/** The line under a file card: how many rows, and the other workbooks it comes with. */
function fileNote(m) {
  const rows = m.first && m.last ? m.last - m.first + 1 : MS.rows.length;
  const more = m.books && m.books.length > 1 ? ' · with ' + m.books.slice(1).map((b) => b.file).join(', ') : '';
  return rows + ' rows' + more;
}
/** Whether a Continue press is waiting, and where it leads. */
function continueLabel() {
  const n = current();
  if (!n) return LIST.every(isClosed) ? 'Continue' : 'Back to chats';
  if (ST.promoDue) return 'Continue';
  const h = ST.hold && LIST.find((x) => x.id === ST.hold);
  return h && n.from === h.from ? 'Continue: next request' : 'Continue to ' + PEOPLE[n.from].short;
}
/** The learner pressed Continue: only now does Kicau move on. */
function cont() {
  ST.hold = null;
  save();
  const n = current();
  if (ST.promoDue) { ST.promoDue = false; save(); showPromo(); return; }
  if (!n) { if (LIST.every(isClosed)) showFinale(); else location.hash = '#desk'; return; }
  if (location.hash === '#chat-' + n.from) renderChat(n.from); else location.hash = '#chat-' + n.from;
}
/** The closed request's sheet, read-only, with the formula that solved it and a Continue button. */
function reviewHTML(m) {
  const st = MST(m.id);
  const formulaKind = m.kind === 'formula' || (m.kind === 'excel' && !m.pivotCheck && !m.wantValues);
  const body = formulaKind
    ? `<div class="fbar ro"><span class="nbox">${esc(m.answerCell)}</span><span class="fxi" aria-hidden="true">fx</span><div class="fxro" title="${esc(st.solved || m.solution)}">${esc(st.solved || m.solution)}</div></div><div class="gw" id="gw">${sheetGrid(m, true)}</div>${sheetTabsHTML(m)}`
    : `<div class="page"><div class="empty">${I('check')} Request ${esc(m.id)} is closed. Take your time, then continue.</div></div>`;
  return `<section class="dock open review" id="dock" aria-label="Your finished sheet">
    <div class="shd">${bookTabsHTML(m, '· sent to ' + PEOPLE[m.from].short)}${maxButton()}</div>
    ${body}
    <div class="comp"><small class="rv">Review your formula, then continue.</small><span class="sp"></span><button class="send" data-continue>${I('send')}<span>${esc(continueLabel())}</span></button></div></section>`;
}
function wireReview(m) {
  document.querySelectorAll('#dock [data-continue]').forEach((b) => { b.onclick = cont; });
  wireTabs(m, true);
}
/* ---------- Maximize the sheet on a laptop ---------- */
const maxButton = () => (LAP.matches ? `<button class="icb" id="dockmax" aria-label="${UI.max ? 'Show the chat again' : 'Maximize the workbook'}" title="${UI.max ? 'Show the chat' : 'Maximize the workbook'}">${I(UI.max ? 'collapse' : 'expand')}</button>` : '');
function wireMax() {
  const b = $('#dockmax');
  if (!b) return;
  b.onclick = () => {
    UI.max = !UI.max;
    const lap = document.querySelector('.lap');
    if (lap) lap.classList.toggle('max', UI.max);
    b.classList.remove('unread');
    b.innerHTML = I(UI.max ? 'collapse' : 'expand');
    b.setAttribute('aria-label', UI.max ? 'Show the chat again' : 'Maximize the workbook');
    b.title = UI.max ? 'Show the chat' : 'Maximize the workbook';
    if (!UI.max) { UI.miniOpen = false; const mi = $('#mini'); if (mi) mi.classList.remove('open', 'unread'); }
  };
}
/*
 * The small chat window. When the sheet is expanded (Maximize on a laptop, full screen on a phone) the chat is
 * out of sight, so a "Chat with ..." button floats over the sheet. Closed, it shows the newest line and a red dot
 * when something new arrives. Open, it is a small window: the request pinned on top, then the latest messages.
 */
function mountMini(p) {
  const d = $('#dock');
  if (!d || !PEOPLE[p] || $('#mini')) return;
  const el = document.createElement('div');
  el.className = 'mini' + (UI.miniOpen ? ' open' : '');
  el.id = 'mini';
  el.innerHTML = `<button class="mini-pill" id="minipill" aria-expanded="${!!UI.miniOpen}" aria-controls="miniwin">${avatar(p, 28)}<span class="mp-t" id="minitext">Chat with ${esc(PEOPLE[p].short)}</span><i class="dot" aria-hidden="true"></i></button>
    <div class="mini-win" id="miniwin" role="dialog" aria-label="Chat with ${esc(PEOPLE[p].name)}"><div class="mini-hd">${avatar(p, 28)}<b>${esc(PEOPLE[p].name)}</b><span class="sp"></span><button class="icb" id="minifull" title="Show the full chat" aria-label="Show the full chat">${I('collapse')}</button><button class="icb" id="miniclose" aria-label="Close the small chat">×</button></div><div class="mini-ask" id="miniask"></div><div class="mini-log" id="minilog" role="log" aria-label="Latest messages"></div></div>`;
  d.appendChild(el);
  const setOpen = (on) => { UI.miniOpen = on; el.classList.toggle('open', on); $('#minipill').setAttribute('aria-expanded', String(on)); if (on) { el.classList.remove('unread'); $('#minipill').innerHTML = `${avatar(p, 28)}<span class="mp-t" id="minitext">Chat with ${esc(PEOPLE[p].short)}</span><i class="dot" aria-hidden="true"></i>`; const b = $('#dockmax'); if (b) b.classList.remove('unread'); renderMini(); } };
  $('#minipill').onclick = () => setOpen(true);
  $('#miniclose').onclick = () => setOpen(false);
  $('#minifull').onclick = () => { setOpen(false); if (LAP.matches) { const b = $('#dockmax'); if (b && UI.max) b.click(); } else setDock('open'); };
  renderMini();
}
/** One message as a short line for the small window. */
function miniLine(e) {
  if (e.k === 'closed') return `<p class="mm sys">${I('check')} Request ${esc(e.mission)} closed</p>`;
  const me = e.who === 'me';
  const who = me ? 'You' : e.who && PEOPLE[e.who] ? PEOPLE[e.who].short : '';
  let t;
  if (e.k === 'answer') t = `<code>${esc(e.formula)}</code>`;
  else if (e.k === 'upload') t = esc(e.name + ', ' + e.summary);
  else if (e.k === 'pick') t = esc(fmt(e.value) + ' distributors need review');
  else if (e.k === 'nudge') t = `<i>${esc(e.title)}.</i> ${esc(e.plain || '')}`;
  else if (e.k === 'lesson') { const m = LIST.find((x) => x.id === e.mission); t = m && m.lesson ? '<i>Lesson:</i> ' + esc(strip(m.lesson.idea)) : ''; }
  else if (e.k === 'text' || e.k === 'hint' || e.k === 'sari') t = esc(strip(fill(e.html || '')));
  else return '';
  return t ? `<p class="mm${me ? ' me' : ''}"><b>${esc(who)}</b> ${t}</p>` : '';
}
function renderMini() {
  const box = $('#minilog');
  if (!box || !UI.person) return;
  const m = activeFor(UI.person) || (ST.hold ? LIST.find((x) => x.id === ST.hold) : null);
  $('#miniask').innerHTML = m ? `<b>Request ${esc(m.id)}</b> ${strip(m.ask)}` : '';
  box.innerHTML = logOf(UI.person).slice(-12).map(miniLine).join('') + (UI.busy[UI.person] ? `<p class="mm"><span class="typing"><span class="dots"><i></i><i></i><i></i></span></span></p>` : '');
  box.scrollTop = box.scrollHeight;
}
/** A new line from a colleague: the closed small chat shows it with a red dot. */
function miniNew(e) {
  const el = $('#mini');
  if (!el || !e.who || e.who === 'me') return;
  const expanded = (LAP.matches && UI.max) || (!LAP.matches && UI.dock === 'full');
  if (!expanded || UI.miniOpen) return;
  el.classList.add('unread');
  const pill = $('#minipill');
  if (pill) pill.innerHTML = `${avatar(e.who, 28)}<span class="mp-t" id="minitext">${esc(PEOPLE[e.who].short + ': ' + strip(fill(e.html || e.title || 'Lesson')))}</span><i class="dot" aria-hidden="true"></i>`;
  const b = $('#dockmax');
  if (b) b.classList.add('unread');
}
function idlePanel(p) {
  const last = [...LIST].reverse().find((m) => m.from === p && isClosed(m));
  if (!last || last.kind !== 'formula') return `<section class="dock open"><div class="shd">${I('bird')}<span class="nm">Nothing open with ${esc(PEOPLE[p].short)}</span></div><div class="page"><div class="empty">${current() ? `${esc(PEOPLE[current().from].name)} has the next request.` : 'All caught up.'}</div>${current() ? `<p style="margin-top:12px"><a class="btn" href="#chat-${current().from}">Open ${esc(PEOPLE[current().from].short)}'s chat</a></p>` : ''}</div></section>`;
  return `<section class="dock open"><div class="shd"><span class="xl">X</span><span class="nm">${esc(last.file)}</span><small>· your last answer for ${esc(PEOPLE[p].short)}</small></div><div class="gw">${gridHTML(last, true)}</div>${SHEET_TABS}<div class="fxlabel" style="padding-top:10px">Formula in ${esc(last.answerCell)} <b>${esc(MST(last.id).solved || '')}</b></div></section>`;
}

/** Messages from someone, one at a time, with a typing bubble between them. */
async function play(p, entries, gap) {
  UI.busy[p] = true;
  for (const e of entries) {
    if (UI.person === p) { const w = $('#whostate'); if (w) { w.textContent = 'typing…'; w.className = 'live'; } }
    renderWall(true);
    await wait(e.who === 'sari' || e.k === 'lesson' || e.k === 'nudge' || e.k === 'hint' ? 500 : Math.min(1500, 450 + strip(e.html || '').length * 12));
    e.t = e.t || clock();
    logOf(p).push(e);
    save();
    if (UI.person === p) miniNew(e);
    say((e.who && e.who !== 'me' ? PEOPLE[e.who].short + ': ' : '') + (e.html || e.title || ''));
    renderWall();
    await wait(gap || 160);
  }
  UI.busy[p] = false;
  if (UI.person === p) { const w = $('#whostate'); if (w) { w.textContent = PEOPLE[p].role + ', ' + PEOPLE[p].org; w.className = ''; } }
  renderWall();
}
async function startMission(m) {
  if (ST.hold) { ST.hold = null; save(); }
  /* A new request is read in the chat, so a maximized sheet steps back. */
  if (UI.max) { UI.max = false; const lap = document.querySelector('.lap'); if (lap) lap.classList.remove('max'); const b = $('#dockmax'); if (b) b.innerHTML = I('expand'); }
  const st = MST(m.id);
  st.started = true;
  markDay();
  logOf(m.from).push({ k: 'divider', mission: m.id });
  save();
  const lines = m.intro.map((l) => (l === '@file' ? { k: 'file', who: m.from, mission: m.id } : { k: 'text', who: m.from, html: fill(l) }));
  await play(m.from, lines);
  await play(m.from, [{ k: 'lesson', who: 'sari', mission: m.id, open: true }]);
  if (!LAP.matches) setDock(m.kind === 'macro' ? 'full' : 'open');
}

function entryHTML(e, i, log) {
  const prev = log[i - 1];
  const m = e.mission ? byId[e.mission] : null;
  switch (e.k) {
    case 'divider': return `<div class="divider${isClosed(m) ? ' done' : ''}">${I(isClosed(m) ? 'check' : 'spark')}Request ${esc(m.id)} · ${esc(m.title)}</div>`;
    case 'text': {
      const me = e.who === 'me';
      const cont = prev && (prev.k === 'text' || prev.k === 'file') && prev.who === e.who;
      return `<div class="b ${me ? 'me' : 'them'}${cont ? ' cont' : ''}${e.react ? ' has-react' : ''}">${e.html}<div class="t">${esc(e.t)}${me ? I('check2') : ''}</div>${e.react ? `<span class="react">${e.react}<b>1</b></span>` : ''}</div>`;
    }
    case 'file': {
      const macro = m.kind === 'macro';
      return `<div class="b them cont"><button class="file" data-opendock>${`<span class="xl${macro ? ' m' : ''}">X</span>`}<span><b>${esc(m.file)}</b><small>${macro ? '6 distributors, 1 macro' : fileNote(m)}</small></span><span class="open">Open</span></button><div class="t">${esc(e.t)}</div></div>`;
    }
    case 'lesson': {
      const L = m.lesson;
      const ex = L.ex ? `<ol>${L.ex.map(([c, n]) => `<li><code>${esc(c)}</code><span>${esc(n)}</span></li>`).join('')}</ol><p class="exres" data-exres="${m.id}">Working out its result…</p>` : '';
      return `<div class="mentor lesson">${avatar('sari', 34)}<div class="body"><small>Mbak Sari · ${isClosed(m) ? 'what you used' : 'before you start'}</small><p class="idea">${esc(L.idea)}</p>${L.teach}<details${e.open && !isClosed(m) ? ' open' : ''} class="more"><summary>${L.ex ? 'Example, line by line' : 'Where you have seen this'}</summary>${ex}${L.bridge ? `<p class="bridge"><b>From your ledger days:</b> ${esc(L.bridge)}</p>` : ''}</details></div></div>`;
    }
    case 'answer': {
      let rows = '';
      if (e.ok && (m.id === '1.6' || m.id === '1.5')) {
        const hit = MS.rows.filter((r) => r.distributor === MS.target);
        rows = `<table><tbody>${hit.slice(0, 2).map((r) => `<tr><td>${esc(r.product)}</td><td class="n">${m.id === '1.6' ? fmt(r.net) : dateText(window.KenariData.toSerial(r.date))}</td></tr>`).join('')}<tr class="more"><td>and ${hit.length - 2} more rows</td><td class="n">${hit.length} of ${MS.rows.length}</td></tr></tbody></table>`;
      }
      return `<div class="b me"><span class="code">${esc(e.formula)}</span><div class="ans"><div class="rh${m.spill ? ' col' : ''}"><span>${m.spill ? 'Spills from ' + esc(m.spill.anchor) : m.fill ? esc(m.fill.col + m.fill.first + ':' + m.fill.col + m.fill.last) + ', ' + (e.rows || '') + ' rows' : esc(m.answerCell || 'Answer')}${e.ok && m.id === '1.6' ? ', ' + esc(MS.target) : ''}</span><b${m.spill ? ' class="sum"' : ''}>${esc(fmt(e.value))}</b></div>${rows}</div><div class="t">${esc(e.t)}${I('check2')}</div></div>`;
    }
    case 'upload': return `<div class="b me"><div class="file" style="margin:0"><span class="xl">X</span><span><b>${esc(e.name)}</b><small>${esc(e.summary)}</small></span></div><div class="t">${esc(e.t)}${I('check2')}</div></div>`;
    case 'built': return `<div class="b me"><span class="code">${esc(e.text)}</span>${e.html ? `<div class="ans bans">${e.html}</div>` : ''}<div class="t">${esc(e.t)}${I('check2')}</div></div>`;
    case 'pick': return `<div class="b me"><b>${esc(fmt(e.value))}</b> distributors need review.<div class="t">${esc(e.t)}${I('check2')}</div></div>`;
    case 'nudge': return `<div class="mentor warn">${avatar('sari', 34)}<div class="body"><small>Mbak Sari${e.pre ? ' · before you send' : ''}</small><p><b>${esc(e.title)}</b></p><p>${esc(e.plain)}</p>${e.fix ? `<div class="acts"><button class="btn sm" data-fix="${esc(e.fix.formula)}">${esc(e.fix.label)}</button></div>` : ''}</div></div>`;
    case 'hint': return `<div class="mentor">${avatar('sari', 34)}<div class="body"><small>${I('coffee')} Coffee with Mbak Sari${e.n ? ' · number ' + e.n + ' today' : ''}</small><p>${e.html}</p></div></div>`;
    case 'sari': return `<div class="mentor">${avatar('sari', 34)}<div class="body"><small>Mbak Sari</small><p>${e.html}</p></div></div>`;
    case 'choices': {
      const done = log.some((x) => x.k === 'pick' && x.mission === e.mission && x.ok);
      if (done) return '';
      return `<div class="choices">${m.choices.map((n) => `<button class="btn sm" data-pick="${n}">${n} need review</button>`).join('')}</div>`;
    }
    case 'closed': return closedHTML(e, m);
    default: return '';
  }
}
function closedHTML(e, m) {
  const n = e.n, nx = nextRank(n - 1) || nextRank(n);
  const rk = rankOf(n - 1);
  const segs = nx ? nx.at - rk.at : 8;
  const filled = n - rk.at;
  const fresh = Date.now() - (e.at || 0) < 6000;
  const nextM = LIST[LIST.indexOf(m) + 1];
  const tags = [`<span class="tag ok">${I('note')}${esc(m.note.fn)} saved to Notes</span>`];
  if (m.kind === 'formula') tags.push('<span class="tag">Checked on 2 sheets</span>');
  tags.push(`<span class="tag">${e.coffees ? `${e.coffees} ${e.coffees === 1 ? 'coffee' : 'coffees'} with Sari` : 'No coffee needed'}</span>`);
  if (e.attempts === 1) tags.push(`<span class="tag ok">${I('bolt')}First try</span>`);
  const btn = ST.hold === m.id ? `<button class="btn sm" data-continue>${esc(continueLabel())}</button>` : '';
  return `<div class="done">${fresh ? confetti(24) : ''}<div class="ck">${I('check')}</div><h3>Request closed</h3><p>${esc(PEOPLE[m.from].short)} has what they needed.</p>
    <div class="ladder">${Array.from({ length: segs }, (_, i) => `<i class="${i < filled - 1 ? 'on' : i === filled - 1 ? (fresh ? 'new' : 'on') : ''}"></i>`).join('')}</div>
    <div class="ladder-l"><span>${esc(rk.title)}</span><span>${n} of ${nx ? nx.at : n}</span><span>${esc(nx ? nx.title : '')}</span></div>
    <div class="tags">${tags.join('')}</div>${btn ? `<div class="acts"><a class="btn sm ghost" href="#notes">${I('note')}Notes</a>${btn}</div>` : ''}</div>`;
}
function confetti(n) {
  const cols = ['#3E5BF2', '#F2B705', '#22C55E', '#EF4444', '#8B7CF6', '#06B6D4', '#F472B6'];
  let h = '<div class="conf" aria-hidden="true">';
  for (let i = 0; i < n; i++) {
    const x = Math.random() * 100, y = 50 + Math.random() * 120, r = (Math.random() * 720 - 360) | 0, d = (Math.random() * 0.5).toFixed(2);
    h += `<i style="left:${x}%;background:${cols[i % cols.length]};--y:${y}px;--r:${r}deg;animation-delay:${d}s"></i>`;
  }
  return h + '</div>';
}
function renderWall(typing) {
  const w = $('#wall');
  if (!w || !UI.person) return;
  const log = logOf(UI.person);
  const head = !log.length && !UI.busy[UI.person] ? `<span class="day">${esc(PEOPLE[UI.person].short)} has not written yet</span>` : `<span class="day">${DAYS[new Date().getDay()]}</span>`;
  w.innerHTML = head + log.map((e, i) => entryHTML(e, i, log)).join('') + ((typing || UI.busy[UI.person]) ? `<div class="b them typingb" aria-label="typing"><span class="typing"><span class="dots"><i></i><i></i><i></i></span></span></div>` : '');
  w.querySelectorAll('[data-exres]').forEach(paintExample);
  w.scrollTop = w.scrollHeight;
  w.querySelectorAll('[data-opendock]').forEach((b) => { b.onclick = () => { setDock(LAP.matches ? 'open' : 'open'); const g = $('#gw'); if (g) g.scrollTop = 0; }; });
  w.querySelectorAll('[data-fix]').forEach((b) => { b.onclick = () => { setFormula(b.dataset.fix); }; });
  w.querySelectorAll('[data-pick]').forEach((b) => { b.onclick = () => pick(+b.dataset.pick); });
  w.querySelectorAll('[data-continue]').forEach((b) => { b.onclick = cont; });
  renderMini();
}

/* ---------- The workbook dock ---------- */
/* The sheet tab strip under the grid, as in desktop Excel. Decorative: each request has one sheet. */
const SHEET_TABS = '<div class="sheettabs" aria-hidden="true"><span class="arr">◀ ▶</span><b>Sheet1</b></div>';
const COLS = ['A', 'B', 'C', 'D', 'E'];
const colAt = (c, k) => String.fromCharCode(c.charCodeAt(0) + k);
const spillRefs = (m) => { const a = /^([A-Z])(\d+)$/.exec(m.spill.anchor); const out = []; for (let r = 0; r < m.spill.rows; r++) for (let c = 0; c < m.spill.cols; c++) out.push(colAt(a[1], c) + (+a[2] + r)); return out; };
const fillRefs = (m) => (m.spill ? spillRefs(m) : m.fill ? Array.from({ length: m.fill.last - m.fill.first + 1 }, (_, i) => m.fill.col + (m.fill.first + i)) : [m.answerCell]);
function gridHTML(m, solved) {
  const c = m.cells;
  const cols = m.cols || COLS, num = m.num || ['D', 'E'], dates = m.dateCols || ['A'];
  const st = MST(m.id);
  let shownVals = null;
  if (solved && st.solved && ENG.api) shownVals = m.fill ? fillValues(m, st.solved, 'base') : [evalOn(ENG.api, m.cells, m.answerCell, st.solved)];
  else if (st.shownVals) shownVals = st.shownVals;
  else if (st.shown !== undefined && !m.fill) shownVals = [st.shown];
  const answers = fillRefs(m);
  const act = /^([A-Z]+)(\d+)$/.exec(m.answerCell || '') || [];
  const head = `<tr><th class="cn" aria-hidden="true"></th>${cols.map((col) => `<th scope="col"${col === act[1] ? ' class="on"' : ''}><button type="button" data-col="${col}" aria-label="Column ${col}">${col}</button></th>`).join('')}</tr>`;
  let body = '';
  for (let r = 1; r <= m.last; r++) {
    body += `<tr><th scope="row"${String(r) === act[2] ? ' class="on"' : ''}>${r}</th>`;
    for (const col of cols) {
      const ref = col + r;
      let v = c[ref], cls = '';
      if (r === 4) cls = 'hd';
      else if (r < 4 && col === 'A') cls = 'lb';
      else if (r >= m.first && num.includes(col)) { cls = 'n'; v = fmt(v); }
      else if (r >= m.first && dates.includes(col)) v = dateText(v);
      else if ((m.dateCells || []).includes(ref) && typeof v === 'number') v = dateText(v);
      else if (typeof v === 'number') { cls = 'n'; v = fmt(v); }
      const ai = answers.indexOf(ref);
      if (ai >= 0) {
        if (shownVals) { cls = (solved || st.shownOk) ? 'got' : 'miss'; v = shownVals[ai] === undefined ? '' : fmtAns(m, shownVals[ai]); }
        else cls = ai === 0 ? 'ask' : 'askfill';
      }
      body += `<td data-ref="${ref}" class="${cls}"${typeof v === "string" && v.length > 16 ? ` title="${esc(v)}"` : ""}>${v === undefined ? '' : esc(v)}</td>`;
    }
    body += '</tr>';
  }
  return `<table class="gx${m.cols ? ' wide' : ''}" aria-label="${esc(m.file)}"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}
/* ---------- Several workbooks ----------
   A request can have more than one workbook open (m.books, Chapter 11). Each book has sheets; the request's own
   sheet is the first sheet of the first book. The learner switches with the workbook tabs in the title bar and
   the sheet tabs under the grid, and tapping a cell there writes the reference as Excel does:
   Diskon!A5 for a sheet of the same file, [master_produk.xlsx]Produk!$A$5 for another file. */
const SHEETS_K = Symbol.for('kenari.sheets');
const booksOf = (m) => m.books || [{ file: m.file, sheets: [{ name: m.sheetName || 'Sheet1', cols: m.cols || COLS, main: true, prefix: '' }] }];
const viewOf = (m) => { if (!UI.view || UI.view.id !== m.id) UI.view = { id: m.id, b: 0, s: 0 }; return UI.view; };
const viewSheet = (m) => { const v = viewOf(m), B = booksOf(m), b = B[v.b] || B[0]; return { book: b, sheet: b.sheets[v.s] || b.sheets[0] }; };
const engName = (m, file, sheet) => (file === m.file ? sheet : file + '|' + sheet);
function sheetCells(m, file, name, variant) {
  const cells = m.variantCells ? m.variantCells[variant || 'base'] : m.cells;
  const s = (cells[SHEETS_K] || []).find((x) => x.name === engName(m, file, name));
  return s ? s.cells : {};
}
/** The grid for whatever sheet is in view: the request's own, or one of the other open sheets. */
function sheetGrid(m, solved) {
  const { book, sheet } = viewSheet(m);
  if (sheet.main) return gridHTML(m, solved);
  const c = sheetCells(m, book.file, sheet.name);
  const head = `<tr><th class="cn" aria-hidden="true"></th>${sheet.cols.map((col) => `<th scope="col"><button type="button" data-col="${col}" aria-label="Column ${col}">${col}</button></th>`).join('')}</tr>`;
  let body = '';
  for (let r = 1; r <= sheet.last; r++) {
    body += `<tr><th scope="row">${r}</th>`;
    for (const col of sheet.cols) {
      const ref = col + r;
      let v = c[ref], cls = '';
      if (r === 4) cls = 'hd';
      else if (r >= sheet.first && (sheet.dateCols || []).includes(col) && typeof v === 'number') v = dateText(v);
      else if (typeof v === 'number') { cls = 'n'; v = fmt(v); }
      body += `<td data-ref="${ref}" class="${cls}"${typeof v === 'string' && v.length > 16 ? ` title="${esc(v)}"` : ''}>${v === undefined ? '' : esc(v)}</td>`;
    }
    body += '</tr>';
  }
  return `<table class="gx" aria-label="${esc(book.file + ', sheet ' + sheet.name)}"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}
/** The title bar: one tab per open workbook, or the file name when there is only one. */
function bookTabsHTML(m, sub) {
  const B = booksOf(m), v = viewOf(m);
  if (B.length < 2) return `<span class="xl">X</span><span class="nm">${esc(m.file)}</span><small>${esc(sub)}</small>`;
  return `<div class="bktabs" id="bktabs" role="tablist" aria-label="Open workbooks">${B.map((b, i) => `<button class="bk${i === v.b ? ' on' : ''}" role="tab" aria-selected="${i === v.b}" data-book="${i}"><span class="xl">X</span>${esc(b.file)}</button>`).join('')}</div>`;
}
/** The sheet tabs under the grid, for the workbook in view. */
function sheetTabsHTML(m) {
  if (!m.books) return SHEET_TABS;
  const { book, sheet } = viewSheet(m);
  return `<div class="sheettabs" id="sheettabs" role="tablist" aria-label="Sheets in ${esc(book.file)}"><span class="arr" aria-hidden="true">◀ ▶</span>${book.sheets.map((s, i) => `<button class="st${s === sheet ? ' on' : ''}" role="tab" aria-selected="${s === sheet}" data-sheet="${i}">${esc(s.name)}</button>`).join('')}</div>`;
}
/** Switch workbook or sheet without touching the formula being typed. */
function wireTabs(m, solved) {
  const d = $('#dock');
  if (!d || !m.books) return;
  d.addEventListener('click', (e) => {
    const bk = e.target.closest('[data-book]'), sh = e.target.closest('[data-sheet]');
    if (!bk && !sh) return;
    const v = viewOf(m);
    if (bk) { v.b = +bk.dataset.book; v.s = 0; } else v.s = +sh.dataset.sheet;
    const gw = $('#gw');
    if (gw) { gw.innerHTML = sheetGrid(m, solved); gw.scrollTop = 0; gw.scrollLeft = 0; }
    const t1 = $('#bktabs'), t2 = $('#sheettabs');
    if (t1) t1.outerHTML = bookTabsHTML(m, '');
    if (t2) t2.outerHTML = sheetTabsHTML(m);
    if (!solved) paint();
    say(viewSheet(m).book.file + ', sheet ' + viewSheet(m).sheet.name);
  });
}
/** A cell or range in the sheet in view, written from the request's own sheet the way Excel writes it. */
function refFromView(m, a, b, colRange) {
  const { book, sheet } = viewSheet(m);
  if (sheet.main) return colRange || (b ? a + ':' + b : a);
  const ext = book.file !== m.file;
  const pin = (r) => (ext ? r.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2') : r);
  return sheet.prefix + (colRange ? pin(colRange[0]) + ':' + pin(colRange[1]) : pin(a) + (b ? ':' + pin(b) : ''));
}
function dockHTML(m) {
  if (m.kind === 'macro') return macroDockHTML(m);
  if (m.kind === 'excel') return excelDockHTML(m);
  if (BUILDERS.includes(m.kind)) return buildDockHTML(m);
  if (m.kind === 'vbafix') return vbaFixHTML(m);
  const chipBtns = `<button class="chip fn" data-key="${esc(m.fnKey)}">${esc(m.fnKey)}</button>${m.chips.map((ch) => `<button class="chip" data-key="${esc(ch)}">${esc(ch)}</button>`).join('')}`;
  const keys = coarse ? `<div class="keys" id="keys">${chipBtns}<span class="sep"></span>${['=', ',', '(', ')', ':', '$', '"', '>', '<', '>=', '*', '/', '+', '-'].map((k) => `<button type="button" data-key="${esc(k)}">${esc(k)}</button>`).join('')}<button type="button" class="wd" data-act="back" aria-label="Delete one character">⌫</button><button type="button" class="wd" data-act="clear">Clear</button><button type="button" class="wd" data-act="abc" aria-label="Show the phone keyboard">ABC</button></div>` : '';
  const left = coffeesLeft();
  return `<section class="dock ${LAP.matches ? 'open' : UI.dock}" id="dock" aria-label="Workbook">
    <button class="grab" id="grab" aria-label="Make the workbook bigger or smaller"></button>
    <div class="shd">${bookTabsHTML(m, '· from ' + PEOPLE[m.from].short)}${dlButton()}${LAP.matches ? maxButton() : `<button class="icb" id="dockbig" aria-label="Full screen sheet">${I(UI.dock === 'full' ? 'collapse' : 'expand')}</button>`}</div>
    <div class="fbar"><span class="nbox" title="Name box: the cell your answer goes in">${esc(m.answerCell)}</span><span class="fxi" aria-hidden="true">fx</span><div class="fxbox"><div class="paint" id="paint" aria-hidden="true"></div><input id="fx" autocomplete="off" autocapitalize="characters" spellcheck="false" aria-label="Formula for ${esc(m.answerCell)}" ${coarse ? 'inputmode="none"' : ''}></div></div>
    ${m.spill ? `<div class="fxlabel">Type one formula in <b>${esc(m.answerCell)}</b>. Its answer spills down and across by itself.</div>` : m.fill ? `<div class="fxlabel">Type it in <b>${esc(m.answerCell)}</b>. When you send, Kicau copies it down to <b>${esc(m.fill.col + m.fill.last)}</b>.</div>` : LAP.matches ? `<div class="fxlabel">Answer cell <b>${esc(m.answerCell)}</b> Tap cells, or drag across them. F4 fixes a reference with $.</div>` : ''}
    <div class="gw" id="gw">${sheetGrid(m)}</div>
    ${sheetTabsHTML(m)}
    ${coarse ? '' : `<div class="chips" id="chips"><small>Tap a cell, or</small>${chipBtns}</div>`}
    ${keys}
    <div class="engmsg" id="engmsg" ${ENG.api ? 'hidden' : ''}>${ENG.failed ? 'The practice sheet did not load. Reload to try again.' : 'Warming up the practice sheet…'}</div>
    <div class="comp"><button class="coffee${left ? '' : ' out'}" id="coffee" aria-label="Coffee with Mbak Sari, ${left} left today" title="Ask Mbak Sari for a hint">${I('coffee')}<i>${left}</i></button>
      <span class="sp"></span><button class="send" id="send" aria-label="Send to ${esc(PEOPLE[m.from].short)}">${I('send')}<span>Send to ${esc(PEOPLE[m.from].short)}</span></button></div></section>`;
}
function setDock(state) {
  if (LAP.matches) return;
  UI.dock = state;
  const d = $('#dock');
  if (!d) return;
  d.classList.remove('open', 'full', 'min');
  d.classList.add(state);
  const w = d.closest('.chatwrap');
  if (w) w.classList.toggle('dock-full', state === 'full');
  const b = $('#dockbig');
  if (b) b.innerHTML = I(state === 'full' ? 'collapse' : 'expand');
}

/** Find every reference in a formula, in order, outside quoted text. */
function refsIn(f) {
  const out = [];
  const plain = f.replace(/"[^"]*"?/g, (s) => ' '.repeat(s.length));
  const re = /((?:'[^']+'|\[[^\]]+\][A-Za-z0-9_.]+|[A-Za-z_][A-Za-z0-9_.]*)!)?(?:\$?([A-Z]{1,2})\$?(\d+)(?::\$?([A-Z]{1,2})\$?(\d+))?|\b([A-Z]):([A-Z])\b)/g;
  let m;
  while ((m = re.exec(plain))) {
    /* Which sheet: none means the request's own; [book]Sheet another workbook; Sheet the same workbook. */
    let book = null, sheet = null;
    if (m[1]) {
      const q = m[1].slice(0, -1).replace(/^'|'$/g, '');
      const b = /^\[([^\]]+)\](.+)$/.exec(q);
      if (b) { book = b[1]; sheet = b[2]; } else sheet = q;
    }
    if (m[6]) out.push({ at: m.index, len: m[0].length, c1: m[6], c2: m[7], r1: 1, r2: 9999, book, sheet });
    else {
      const before = plain[m.index - 1];
      if (!m[1] && before && /[A-Za-z]/.test(before)) continue;
      out.push({ at: m.index, len: m[0].length, c1: m[2], c2: m[4] || m[2], r1: +m[3], r2: +(m[5] || m[3]), book, sheet });
    }
  }
  return out;
}
function paint() {
  const inp = $('#fx'), p = $('#paint');
  if (!inp || !p) return;
  const f = inp.value;
  if (!f) { p.innerHTML = '<span class="ph">Start with =</span>'; return; }
  const refs = refsIn(f);
  let h = '', at = 0;
  refs.forEach((r, i) => { h += esc(f.slice(at, r.at)) + `<span class="p${i % 4}">${esc(f.substr(r.at, r.len))}</span>`; at = r.at + r.len; });
  p.innerHTML = h + esc(f.slice(at)) + ' ';
  p.scrollLeft = inp.scrollLeft;
  /* Tint the cells each reference points at, in the same color, as Excel does while you edit. */
  document.querySelectorAll('#gw td').forEach((td) => td.classList.remove('k0', 'k1', 'k2', 'k3', 'kt', 'kb', 'one'));
  const cm = activeM();
  const inView = (r) => {
    if (!cm || !cm.books) return !r.sheet;
    const { book, sheet } = viewSheet(cm);
    const file = r.book || cm.file, name = r.sheet || cm.sheetName;
    return file === book.file && name === sheet.name;
  };
  refs.forEach((r, i) => {
    if (!inView(r)) return;
    const a = Math.min(r.c1.charCodeAt(0), r.c2.charCodeAt(0)), z = Math.max(r.c1.charCodeAt(0), r.c2.charCodeAt(0));
    const top = Math.min(r.r1, r.r2), bot = Math.min(Math.max(r.r1, r.r2), 60);
    for (let c = a; c <= z; c++) for (let row = top; row <= bot; row++) {
      const td = document.querySelector(`#gw td[data-ref="${String.fromCharCode(c)}${row}"]`);
      if (!td || td.classList.contains('ask')) continue;
      td.classList.add('k' + (i % 4));
      if (top === bot) td.classList.add('one');
      else { if (row === top) td.classList.add('kt'); if (row === bot) td.classList.add('kb'); }
    }
  });
}
function setFormula(f) {
  const inp = $('#fx');
  if (!inp) return;
  inp.value = f;
  afterEdit();
  if (!coarse) { inp.focus({ preventScroll: true }); inp.setSelectionRange(f.length, f.length); }
}
function afterEdit() {
  const inp = $('#fx');
  const m = activeM();
  if (m && inp) { MST(m.id).draft = inp.value; save(); }
  paint();
}
function insert(text) {
  const inp = $('#fx');
  const a = inp.selectionStart == null ? inp.value.length : inp.selectionStart;
  const b = inp.selectionEnd == null ? a : inp.selectionEnd;
  inp.value = inp.value.slice(0, a) + text + inp.value.slice(b);
  const at = a + text.length;
  inp.focus({ preventScroll: true });
  inp.setSelectionRange(at, at);
  inp.scrollLeft = inp.scrollWidth;
  afterEdit();
}
function cycleAbs() {
  const inp = $('#fx');
  const at = inp.selectionStart;
  const left = inp.value.slice(0, at);
  const m = /(\$?)([A-Z]+)(\$?)(\d+)$/.exec(left);
  if (!m) return;
  const next = !m[1] && !m[3] ? `$${m[2]}$${m[4]}` : m[1] && m[3] ? `${m[2]}$${m[4]}` : !m[1] && m[3] ? `$${m[2]}${m[4]}` : `${m[2]}${m[4]}`;
  const start = left.length - m[0].length;
  inp.value = left.slice(0, start) + next + inp.value.slice(at);
  inp.setSelectionRange(start + next.length, start + next.length);
  afterEdit();
}
function wireDock(m) {
  const dl = $('#dlx');
  if (dl) dl.onclick = () => downloadXlsx(m);
  if (m.kind === 'macro') { wireMacro(m); return; }
  if (m.kind === 'excel') { wireExcel(m); return; }
  if (BUILDERS.includes(m.kind)) { wireBuild(m); return; }
  if (m.kind === 'vbafix') { wireVbaFix(m); return; }
  const inp = $('#fx');
  const st = MST(m.id);
  inp.value = st.draft || m.starter;
  paint();
  inp.addEventListener('input', afterEdit);
  inp.addEventListener('scroll', () => { $('#paint').scrollLeft = inp.scrollLeft; });
  inp.addEventListener('keyup', () => { $('#paint').scrollLeft = inp.scrollLeft; });
  inp.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); if (UI.prac) submitPractice(UI.prac); else send(); }
    if (e.key === 'F4') { e.preventDefault(); cycleAbs(); }
  });
  if (!coarse) setTimeout(() => { inp.focus({ preventScroll: true }); inp.setSelectionRange(inp.value.length, inp.value.length); }, 0);
  const gw = $('#gw');
  let from = null, dragged = false;
  gw.addEventListener('mousedown', (e) => { e.preventDefault(); const td = e.target.closest('td[data-ref]'); from = td ? td.dataset.ref : null; dragged = false; });
  gw.addEventListener('mouseup', (e) => { const td = e.target.closest('td[data-ref]'); if (from && td && td.dataset.ref !== from) { dragged = true; insert(refFromView(m, from, td.dataset.ref)); } from = null; });
  gw.addEventListener('click', (e) => {
    if (dragged) { dragged = false; return; }
    const { sheet } = viewSheet(m);
    const col = e.target.closest('[data-col]');
    if (col) { const c = col.dataset.col; insert(sheet.main ? m.colRange[c] : refFromView(m, null, null, [c + sheet.first, c + sheet.last])); return; }
    const td = e.target.closest('td[data-ref]');
    if (!td || (sheet.main && td.dataset.ref === m.answerCell)) return;
    insert(refFromView(m, td.dataset.ref));
    td.classList.add('flash');
    setTimeout(() => td.classList.remove('flash'), 520);
  });
  for (const id of ['chips', 'keys']) {
    const box = $('#' + id);
    if (!box) continue;
    box.addEventListener('mousedown', (e) => e.preventDefault());
    box.addEventListener('click', (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.key) { insert(b.dataset.key); return; }
      if (b.dataset.act === 'back') {
        const a = inp.selectionStart, z = inp.selectionEnd;
        if (a === z && a > 0) { inp.value = inp.value.slice(0, a - 1) + inp.value.slice(z); inp.setSelectionRange(a - 1, a - 1); }
        else if (a !== z) { inp.value = inp.value.slice(0, a) + inp.value.slice(z); inp.setSelectionRange(a, a); }
        inp.focus({ preventScroll: true });
        afterEdit();
      }
      if (b.dataset.act === 'clear') { inp.value = '='; inp.focus({ preventScroll: true }); inp.setSelectionRange(1, 1); afterEdit(); }
      if (b.dataset.act === 'abc') {
        const on = inp.getAttribute('inputmode') === 'none';
        inp.setAttribute('inputmode', on ? 'text' : 'none');
        b.textContent = on ? 'Hide' : 'ABC';
        inp.blur(); inp.focus();
      }
    });
  }
  $('#send').onclick = () => (UI.prac ? submitPractice(UI.prac) : send());
  if ($('#coffee')) $('#coffee').onclick = coffee;
  if ($('#grab')) $('#grab').onclick = () => setDock(UI.dock === 'open' ? 'full' : UI.dock === 'full' ? 'min' : 'open');
  const big = $('#dockbig');
  if (big) big.onclick = () => setDock(UI.dock === 'full' ? 'open' : 'full');
  inp.addEventListener('focus', () => { if (!LAP.matches && UI.dock === 'min') setDock('open'); });
  wireTabs(m, false);
}

/* ---------- Sending an answer ---------- */
function fillValues(m, formula, variant) {
  if (m.spill) return evalSpill(ENG.api, m.variantCells[variant], m.spill.anchor, formula, m.spill.cols, m.spill.rows);
  return evalFill(ENG.api, m.variantCells[variant], m.fill.col, m.fill.first, m.fill.last, formula, Check.shiftRows);
}
const run = (m) => (formula, variant) => (m.fill || m.spill ? Check.columnKey(fillValues(m, formula, variant), m.exactCase) : evalOn(ENG.api, m.variantCells[variant], m.answerCell, formula));
/** How an answer reads: dates as dates, text with its spaces shown. */
const fmtAns = (m, v) => (m.ansDate && typeof v === 'number' ? dateText(v) : typeof v === 'string' && m.exactCase && /^\s|\s$|\s\s/.test(v) ? '"' + v + '"' : fmt(v));
/** A spilled answer in a few words: how many rows it filled, and the first values. */
const spillSummary = (m, vals) => { const rows = []; for (let r = 0; r < m.spill.rows; r++) { const row = vals.slice(r * m.spill.cols, (r + 1) * m.spill.cols); if (row.some((v) => v !== '')) rows.push(row); } return rows.length + (rows.length === 1 ? ' row: ' : ' rows: ') + rows.slice(0, 2).map((r) => r.filter((v) => v !== '').map((v) => fmtAns(m, v)).join(' · ')).join(' / ') + (rows.length > 2 ? ' …' : ''); };
const preview = (vals, m) => vals.slice(0, 3).map((v) => (m ? fmtAns(m, v) : fmt(v))).join(', ') + (vals.length > 3 ? ' …' : '');
function replyFor(m, v) {
  if (v.state === 'pass') return null;
  const trap = (m.traps || []).find((t) => t.code === v.code);
  if (trap) return trap.reply;
  if (v.code === 'technique') return (m.requires[0] || {}).reply;
  if (v.code === 'frozen') return m.replies.frozen;
  if (v.code === 'typed') return m.replies.typed;
  return 'Hmm, {value} does not match my numbers 🤔 Can you check which column and which rows you used?';
}
async function send() {
  const m = current();
  if (!m || m.from !== UI.person || UI.busy[m.from]) return;
  const inp = $('#fx');
  if (!ENG.api) {
    const e = $('#engmsg');
    if (ENG.failed) { if (e) { e.hidden = false; } return; }
    UI.pendingSend = true;
    if (e) { e.hidden = false; e.textContent = 'Warming up the practice sheet. Your answer goes as soon as it is ready.'; }
    return;
  }
  const typed = inp.value;
  const v = Check.judge(m, typed, run(m));
  window.__lastVerdict = v;
  const st = MST(m.id);
  if (v.state === 'empty' || v.state === 'error') {
    await play(m.from, [{ k: 'nudge', who: 'sari', pre: true, title: v.title, plain: v.plain, fix: v.fix }]);
    return;
  }
  st.attempts++;
  const formula = Check.normalize(typed);
  let shown = v.value !== undefined ? v.value : v.typed;
  let colVals = null;
  if ((m.fill || m.spill) && formula[0] === '=') { colVals = fillValues(m, formula, 'base'); shown = m.spill ? spillSummary(m, colVals) : preview(colVals, m); }
  st.shown = shown;
  st.shownVals = colVals;
  st.shownOk = v.state === 'pass';
  logOf(m.from).push({ k: 'answer', who: 'me', mission: m.id, formula, value: shown, rows: colVals ? colVals.length : 0, ok: v.state === 'pass', t: clock() });
  save();
  fillRefs(m).forEach((ref, i) => {
    const td = $(`#gw td[data-ref="${ref}"]`);
    if (!td) return;
    td.textContent = colVals ? fmtAns(m, colVals[i]) : fmtAns(m, shown);
    td.className = v.state === 'pass' ? 'got' : 'miss';
  });
  renderWall();
  const vals = { value: esc(colVals ? fmtAns(m, colVals[0]) : fmtAns(m, shown)) };
  if (v.state === 'pass') {
    st.solved = formula;
    const lines = m.replies.pass.map((l, i, a) => ({ k: 'text', who: m.from, html: fill(l, vals), react: i === a.length - 1 ? (m.from === 'bayu' ? '🚚👍' : '🙏🎉') : '' }));
    await play(m.from, lines);
    await closeMission(m);
  } else {
    await play(m.from, [{ k: 'text', who: m.from, html: fill(replyFor(m, v), vals) }, { k: 'nudge', who: 'sari', title: v.title, plain: v.plain }]);
  }
}
async function closeMission(m) {
  const st = MST(m.id);
  if (!isClosed(m)) ST.closed.push(m.id);
  markDay();
  const promotedNow = ST.closed.length >= 8 && !ST.promoted;
  if (promotedNow) {
    ST.promoted = true;
    ST.logs.citra = [{ k: 'text', who: 'citra', html: 'Halo {name}! Citra, Head of Collection at Bank Kenari 👋', t: clock() }, { k: 'text', who: 'citra', html: 'Bu Lestari told me you read macros now. Selamat on Analyst!', t: clock() }, { k: 'text', who: 'citra', html: 'Our collection files start in Chapter 2. I will send the first one soon 😉', t: clock() }];
    ST.logs.dimas = [{ k: 'text', who: 'dimas', html: 'Selamat, Analyst! Dimas here, Portfolio Analyst at the bank.', t: clock() }, { k: 'text', who: 'dimas', html: 'In Chapter 2 we join Rasa Kenari\'s sales with our loan book. That is the early-warning report everyone talks about.', t: clock() }];
  }
  if (promotedNow) ST.promoDue = true;
  ST.hold = m.id;
  logOf(m.from).push({ k: 'closed', mission: m.id, n: ST.closed.length, coffees: st.coffees, attempts: st.attempts, at: Date.now(), t: clock() });
  save();
  /* The finished sheet stays up with the formula. Nothing moves on until the learner presses Continue. */
  if (UI.person === m.from) renderChat(m.from); else renderWall();
  if (m.after) { await wait(900); await play(m.from, [{ k: 'sari', who: 'sari', html: esc(m.after) }]); }
}
async function coffee() {
  const m = current();
  if (!m || UI.busy[m.from]) return;
  const st = MST(m.id);
  const left = coffeesLeft();
  ST.coffee.used++;
  st.coffees++;
  const i = Math.min(st.hint, (m.hints || []).length - 1);
  st.hint++;
  save();
  const btn = $('#coffee');
  if (btn) { const l = coffeesLeft(); btn.querySelector('i').textContent = l; btn.classList.toggle('out', !l); }
  const extra = left ? '' : 'That is coffee number ' + ST.coffee.used + ' today, I will need a nap 😴 ';
  await play(m.from, [{ k: 'hint', who: 'sari', html: extra + esc(m.hints[i]), n: ST.coffee.used }]);
}

/* ---------- The macro request ---------- */
const MACRO = `Sub MarkOverdue()
    ' Flag every distributor more than 30 days late
    Dim r As Long, lastRow As Long, flagged As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Cells(r, 3).Value > 30 Then
            Cells(r, 4).Value = "REVIEW"
            Cells(r, 4).Interior.Color = vbYellow
            flagged = flagged + 1
        Else
            Cells(r, 4).Value = "OK"
        End If
    Next r
    Range("F1").Value = flagged
    MsgBox flagged & " distributors need review"
End Sub`;
let MAC = null;
const MCOLS = ['A', 'B', 'C', 'D', 'E', 'F'];
/** Scroll a box so el shows, without moving the page around it. Sticky headers take top and left. */
function within(box, el, top = 0, left = 0) {
  const b = box.getBoundingClientRect(), r = el.getBoundingClientRect();
  if (r.top < b.top + top) box.scrollTop -= b.top + top - r.top;
  else if (r.bottom > b.bottom) box.scrollTop += r.bottom - b.bottom;
  if (r.left < b.left + left) box.scrollLeft -= b.left + left - r.left;
  else if (r.right > b.right) box.scrollLeft += r.right - b.right;
}
function macroSheet(m) {
  const s = { A1: 'Distributor', B1: 'Outstanding (Rp)', C1: 'Days late', D1: 'Status', E1: 'To review' };
  m.rows.forEach((r, i) => { const n = i + 2; s['A' + n] = r.distributor; s['B' + n] = r.outstanding; s['C' + n] = r.dpd; });
  return s;
}
function macroDockHTML(m) {
  if (!LAP.matches && UI.dock !== 'min') UI.dock = 'full';
  return `<section class="dock ${LAP.matches ? 'open' : UI.dock}" id="dock" aria-label="Macro workbook">
    <button class="grab" id="grab" aria-label="Make the workbook bigger or smaller"></button>
    <div class="shd"><span class="xl m">X</span><span class="nm">${esc(m.file)}</span><small>· MarkOverdue</small>${LAP.matches ? maxButton() : `<button class="icb" id="dockbig" aria-label="Full screen">${I(UI.dock === 'full' ? 'collapse' : 'expand')}</button>`}</div>
    <div class="mac"><p class="say" id="say" role="status"></p><div class="split"><div class="code-v" id="codev" aria-label="The macro"></div><div class="gw" id="mgw"></div></div><div class="vars" id="vars" aria-label="Variables"></div></div>
    <div class="mac-acts"><button class="btn ghost" id="mback" aria-label="Back one line">${I('undo')}</button><button class="btn ghost" id="mend" aria-label="Run to the end">${I('ff')}End</button><button class="btn" id="mstep">${I('play')}Run the first line</button></div>
    <div class="msgbox" id="mmsg" hidden role="alertdialog" aria-label="Message box"><small>Microsoft Excel</small><span id="mmsgt"></span><button class="btn wide sm" id="mok" style="margin-top:10px">OK</button></div></section>`;
}
const KW = /\b(Sub|End Sub|Dim|As|Long|For|To|Next|If|Then|Else|End If|MsgBox)\b/g;
function vbPaint(line) {
  const c = line.indexOf("'");
  if (c >= 0) return esc(line.slice(0, c)) + '<span class="cm">' + esc(line.slice(c)) + '</span>';
  return esc(line).replace(/&quot;[^&]*&quot;/g, (s) => '<span class="st">' + s + '</span>').replace(/(^|[^\w">])(\d+)\b/g, '$1<span class="nu">$2</span>').replace(KW, '<span class="kw">$1</span>');
}
function wireMacro(m) {
  MAC = VBA.create(MACRO, macroSheet(m));
  $('#codev').innerHTML = MAC.lines.map((l, i) => `<div class="ln" data-i="${i}"><i>${i + 1}</i><span>${vbPaint(l) || ' '}</span></div>`).join('');
  let seen = 0;
  const last = m.rows.length + 1;
  const draw = () => {
    const st = MAC.state;
    const at = st.done ? st.line : st.pc;
    document.querySelectorAll('#codev .ln').forEach((el) => { const on = +el.dataset.i === at; el.classList.toggle('now', on); if (on) within($('#codev'), el); });
    $('#say').innerHTML = `<small>${st.steps === 0 ? 'About to run' : st.done ? 'Finished' : 'What that line did'}</small>${esc(st.say)}`;
    const vars = Object.values(st.vars);
    $('#vars').innerHTML = vars.length ? vars.map((x) => `<span>${esc(x.name)}<b>${esc(x.v)}</b></span>`).join('') : '<span>No variables yet</span>';
    let h = `<table class="gx"><thead><tr><th class="cn"></th>${MCOLS.map((c) => `<th>${c}</th>`).join('')}</tr></thead><tbody>`;
    for (let r = 1; r <= last; r++) {
      h += `<tr><th>${r}</th>`;
      for (const c of MCOLS) { const ref = c + r, v = st.sheet[ref]; const cls = [typeof v === 'number' ? 'n' : '', r === 1 ? 'hd' : '', st.fill[ref] === 'yellow' ? 'yellow' : '', st.changed.includes(ref) ? 'changed' : ''].filter(Boolean).join(' '); h += `<td class="${cls}" data-ref="${ref}">${v === undefined ? '' : esc(fmt(v))}</td>`; }
      h += '</tr>';
    }
    $('#mgw').innerHTML = h + '</tbody></table>';
    const ch = document.querySelector('#mgw td.changed');
    if (ch) within($('#mgw'), ch, 40, 34);
    const box = $('#mmsg');
    box.hidden = !st.msgs.length || seen === st.msgs.length;
    if (st.msgs.length) $('#mmsgt').textContent = st.msgs[st.msgs.length - 1];
    $('#mstep').disabled = st.done;
    $('#mstep').innerHTML = I('play') + (st.done ? 'Finished' : st.steps === 0 ? 'Run the first line' : 'Run the next line');
    $('#mend').disabled = st.done;
    $('#mback').disabled = !MAC.canBack;
    window.__vba = st;
  };
  const ask = async () => {
    const log = logOf(m.from);
    if (log.some((e) => e.k === 'choices' && e.mission === m.id)) return;
    await play(m.from, [{ k: 'text', who: m.from, html: 'So? How many need review? 👀' }, { k: 'choices', mission: m.id }]);
  };
  $('#mstep').onclick = () => { MAC.step(); draw(); };
  $('#mend').onclick = () => { MAC.runToEnd(); draw(); };
  $('#mback').onclick = () => { MAC.back(); seen = Math.min(seen, MAC.state.msgs.length); draw(); };
  $('#mok').onclick = () => { seen = MAC.state.msgs.length; draw(); setDock('open'); ask(); };
  $('#grab').onclick = () => setDock(UI.dock === 'full' ? 'open' : 'full');
  const big = $('#dockbig');
  if (big) big.onclick = () => setDock(UI.dock === 'full' ? 'open' : 'full');
  draw();
  if (logOf(m.from).some((e) => e.k === 'choices' && e.mission === m.id)) { MAC.runToEnd(); seen = MAC.state.msgs.length; draw(); UI.dock = 'open'; }
}
async function pick(n) {
  const m = current();
  if (!m || m.kind !== 'macro' || UI.busy[m.from]) return;
  const st = MST(m.id);
  st.attempts++;
  const ok = n === m.answer;
  logOf(m.from).push({ k: 'pick', who: 'me', mission: m.id, value: n, ok, t: clock() });
  save();
  renderWall();
  if (ok) {
    await play(m.from, m.replies.pass.map((l, i, a) => ({ k: 'text', who: m.from, html: fill(l, { value: n }), react: i === a.length - 1 ? '🙏🎉' : '' })));
    await closeMission(m);
  } else {
    await play(m.from, [{ k: 'text', who: m.from, html: fill(m.replies.wrong) }, { k: 'sari', who: 'sari', html: 'Run it to the end, then look at <b>F1</b>. The macro writes its count there.' }]);
  }
}

/* ---------- Real Excel: download a request's sheet, upload it back ---------- */
const EXCELJS = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
let excelLoading = null;
function loadExcel() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (!excelLoading) excelLoading = new Promise((ok, fail) => { const sc = document.createElement('script'); sc.src = EXCELJS; sc.onload = () => ok(window.ExcelJS); sc.onerror = () => { excelLoading = null; fail(new Error('ExcelJS did not load')); }; document.head.append(sc); });
  return excelLoading;
}
const dlButton = () => `<button class="icb dl" id="dlx" aria-label="Download for Excel" title="Download for Excel">${I('download')}</button>`;
/** The request's sheet as a real .xlsx: the same cells, headers in bold, the answer cells marked. */
/** Write a sheet's cells into an ExcelJS worksheet: bold headers, dates as dates, thousands separators. */
function fillWorksheet(ws, cells, o) {
  for (const [ref, v] of Object.entries(cells)) {
    const cell = ws.getCell(ref);
    const row = +ref.replace(/\D/g, ''), col = ref.replace(/\d/g, '');
    if (((row >= o.first && o.dates.includes(col)) || (o.dateCells || []).includes(ref)) && typeof v === 'number') { cell.value = new Date(Date.UTC(1899, 11, 30) + v * 864e5); cell.numFmt = 'dd mmm'; }
    else cell.value = v;
    if (row === 4) cell.font = { bold: true };
    if (row < 4 && col === 'A') cell.font = { bold: true };
    if (typeof v === 'number' && v >= 1000) cell.numFmt = '#,##0';
  }
  o.cols.forEach((c, i) => { ws.getColumn(i + 1).width = c === 'A' ? 16 : c === 'B' ? 22 : 14; });
}
const xlsxBlob = async (wb) => new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
/** The request's sheet as a real .xlsx: the same cells, headers in bold, the answer cells marked. */
async function buildXlsx(m) {
  const X = await loadExcel();
  const wb = new X.Workbook();
  wb.creator = 'Kicau, Kenari Excel Lab';
  const ws = wb.addWorksheet(m.sheetName || 'Sheet1');
  fillWorksheet(ws, m.cells, { first: m.first, dates: m.dateCols || ['A'], dateCells: m.dateCells, cols: m.cols || COLS });
  const isReport = BUILDERS.includes(m.kind) || m.pivotCheck;
  if (!isReport) for (const ref of fillRefs(m)) ws.getCell(ref).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3B0' } };
  /* The other sheets of the same workbook come along. */
  for (const b of booksOf(m)) if (b.file === m.file) for (const sh of b.sheets) if (!sh.main) fillWorksheet(wb.addWorksheet(sh.name), sheetCells(m, b.file, sh.name), { first: sh.first, dates: sh.dateCols || [], cols: sh.cols });
  const req = wb.addWorksheet('Request');
  req.getCell('A1').value = PEOPLE[m.from].name + ', ' + PEOPLE[m.from].role;
  req.getCell('A2').value = strip(m.ask);
  req.getCell('A3').value = isReport ? 'Build it in Excel from the data on ' + (m.sheetName || 'Sheet1') + '.' : m.spill ? 'Type one formula in ' + m.answerCell + ' and let it spill.' : m.fill ? 'Type the formula in ' + m.answerCell + ', then copy it down to ' + m.fill.col + m.fill.last + '.' : 'Type the formula in ' + m.answerCell + '.';
  if (m.books && m.books.length > 1) req.getCell('A4').value = 'Open these workbooks too, with their names unchanged: ' + m.books.slice(1).map((b) => b.file).join(', ') + '. Then click into them while you type, and Excel writes the links.';
  req.getColumn(1).width = 90;
  return xlsxBlob(wb);
}
/** Every workbook the request has open, each as its own file, named as in the request so Excel links match. */
async function buildBooks(m) {
  const out = [{ filename: m.file, blob: await buildXlsx(m) }];
  const X = await loadExcel();
  for (const b of booksOf(m).slice(1)) {
    const wb = new X.Workbook();
    wb.creator = 'Kicau, Kenari Excel Lab';
    for (const sh of b.sheets) fillWorksheet(wb.addWorksheet(sh.name), sheetCells(m, b.file, sh.name), { first: sh.first, dates: sh.dateCols || [], cols: sh.cols });
    out.push({ filename: b.file, blob: await xlsxBlob(wb) });
  }
  return out;
}
/** Hand the file over: through the viewer's save prompt inside claude.ai, or a normal download elsewhere. */
async function downloadXlsx(m) {
  const btn = $('#dlx');
  try {
    if (btn) btn.disabled = true;
    /* One request, one file; a request with several workbooks gives each its real name, so Excel's links match. */
    const files = m.books && m.books.length > 1 ? await buildBooks(m) : [{ filename: 'kicau-' + m.id + '-' + m.file.replace(/\.xls[xm]$/, '') + '.xlsx', blob: await buildXlsx(m) }];
    const host = window.claude && window.claude.use ? await window.claude.use('downloads') : null;
    for (const f of files) {
      if (host) {
        try { await host.save({ filename: f.filename, data: f.blob }); flash('Saved ' + f.filename); }
        catch (err) { if (err && err.code !== 'declined') flash(err.code === 'unavailable' || err.code === 'not_granted' ? 'Downloads are off in this view.' : 'The download did not go through.'); if (err && (err.code === 'declined' || err.code === 'unavailable' || err.code === 'not_granted')) break; }
      } else if (window.KENARI_ARTIFACT) {
        flash('Downloads are off in this view.'); break;
      } else {
        const url = URL.createObjectURL(f.blob);
        const a = document.createElement('a');
        a.href = url; a.download = f.filename; document.body.append(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 4000);
        flash(files.length > 1 ? 'Downloaded ' + files.map((x) => x.filename).join(' and ') : 'Downloaded ' + f.filename);
        if (files.length > 1) await wait(400);
      }
    }
  } catch (err) {
    flash('Excel export needs the internet the first time.');
  } finally { if (btn) btn.disabled = false; }
}
function flash(text) {
  document.querySelectorAll('.flashmsg').forEach((x) => x.remove());
  const el = document.createElement('div');
  el.className = 'flashmsg';
  el.setAttribute('role', 'status');
  el.textContent = text;
  document.body.append(el);
  setTimeout(() => el.remove(), 3200);
}
function excelDockHTML(m) {
  return `<section class="dock ${LAP.matches ? 'open' : UI.dock === 'min' ? 'open' : UI.dock}" id="dock" aria-label="Real Excel request">
    <button class="grab" id="grab" aria-label="Make the workbook bigger or smaller"></button>
    <div class="shd"><span class="xl">X</span><span class="nm">${esc(m.file)}</span><small>· real Excel</small>${maxButton()}</div>
    <div class="xsteps"><div class="xstep"><b>1</b><span>${m.books && m.books.length > 1 ? `Download the ${m.books.length} files and open them all in Excel` : 'Download the file'}</span><button class="btn sm" id="xdl">${I('download')}Download for Excel</button></div>
    <div class="xstep"><b>2</b><span>${m.pivotCheck ? 'Insert a PivotTable: Distributor in Rows, Sum of Net in Values, then save' : m.spill ? `Type one formula in <code>${esc(m.answerCell)}</code> in Excel and let it spill, then save` : m.fill ? `Type the formula in <code>${esc(m.answerCell)}</code> and copy it down to <code>${esc(m.fill.col + m.fill.last)}</code> in Excel, then save` : `Type the formula in <code>${esc(m.answerCell)}</code> in Excel, then save`}</span></div>
    <div class="xstep"><b>3</b><span>Send the file back</span><label class="btn sm" for="xup">${I('upload')}Upload .xlsx</label><input type="file" id="xup" accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12" class="sr"></div></div>
    <div class="gw" id="gw">${gridHTML(m)}</div></section>`;
}
/** Read an uploaded workbook and judge column I: every grade right, and every cell a formula. */
async function checkUpload(m, buf, name) {
  if (!m || m.kind !== 'excel' || UI.busy[m.from]) return;
  if (m.pivotCheck) return checkPivotUpload(m, buf, name);
  const st = MST(m.id);
  let got = [], formulas = 0, readErr = null;
  try {
    const X = await loadExcel();
    const wb = new X.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.getWorksheet(m.sheetName || 'Sheet1') || wb.worksheets[0];
    for (const ref of fillRefs(m)) {
      const v = ws.getCell(ref).value;
      if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) { formulas++; got.push(v.result === undefined ? 0 : v.result && typeof v.result === 'object' && v.result.error ? v.result.error : v.result); /* Some tools skip storing a result of 0. */ }
      else got.push(v == null && m.spill ? '' : v);
    }
  } catch (err) { readErr = err; }
  st.attempts++;
  const want = m.wantValues ? m.wantValues : !ENG.api ? [] : m.fill || m.spill ? fillValues(m, m.solution, 'base') : [evalOn(ENG.api, m.variantCells.base, m.answerCell, m.solution)];
  const wrongRow = got.findIndex((v, i) => !Check.same(typeof v === 'string' && /^\d+$/.test(v) ? +v : v, want[i], m.precision, m.exactCase));
  const needF = m.valuesOnly ? 0 : m.spill ? 1 : got.length;
  const ok = !readErr && formulas >= needF && wrongRow < 0;
  logOf(m.from).push({ k: 'upload', who: 'me', mission: m.id, name: name || 'workbook.xlsx', summary: readErr ? 'could not be read' : `${formulas} of ${got.length} cells are formulas`, t: clock() });
  save();
  renderWall();
  if (ok) {
    st.solved = m.solution;
    st.shownVals = want; st.shownOk = true;
    await play(m.from, m.replies.pass.map((l, i, a) => ({ k: 'text', who: m.from, html: fill(l), react: i === a.length - 1 ? '🙏🎉' : '' })));
    await closeMission(m);
    return;
  }
  const row = !m.fill && !m.spill ? 4 : m.spill ? +m.spill.anchor.slice(1) + Math.floor(Math.max(0, wrongRow) / m.spill.cols) : m.fill.first + Math.max(0, wrongRow);
  const reply = readErr ? 'Hmm, I cannot open that file 🤔 Is it the .xlsx you downloaded, saved as Excel Workbook?'
    : formulas < needF && wrongRow < 0 ? m.replies.typed
    : m.rowReply ? fill(m.rowReply, { row, got: esc(fmt(got[wrongRow] == null ? 'nothing' : got[wrongRow])), want: esc(fmt(want[wrongRow])), key: esc(m.cells[(m.keyCol || 'B') + row]) })
    : `Row ${row} says ${esc(fmt(got[wrongRow] === undefined || got[wrongRow] === null ? 'nothing' : got[wrongRow]))}, but ${esc(m.cells['B' + row])} is ${m.cells['F' + row]} days late, so our rule gives ${esc(fmt(want[wrongRow]))}.`;
  await play(m.from, [{ k: 'text', who: m.from, html: reply }, { k: 'sari', who: 'sari', html: readErr ? 'Save it from Excel as <b>Excel Workbook (.xlsx)</b>, then upload again.' : 'Fix it in Excel, save, and upload again. Your formula should start in <b>' + m.answerCell + '</b>' + (m.spill ? '.' : ' and be copied down with the fill handle.') }]);
}
/** A real PivotTable: Excel saves its visible cells like any others, so find each distributor and the total beside it. */
async function checkPivotUpload(m, buf, name) {
  const want = m.pairs || MS.DATA_TOTALS, found = {};
  let readErr = null, sheets = 0;
  try {
    const X = await loadExcel();
    const wb = new X.Workbook();
    await wb.xlsx.load(buf);
    sheets = wb.worksheets.length;
    wb.worksheets.forEach((ws) => ws.eachRow((row) => {
      row.eachCell((cell, col) => {
        const k = typeof cell.value === 'string' ? cell.value.trim() : null;
        if (!k || !(k in want)) return;
        const nv = row.getCell(col + 1).value;
        const n = nv && typeof nv === 'object' ? nv.result : nv;
        if (typeof n === 'number' && Math.abs(n - want[k]) < 1) found[k] = true;
      });
    }));
  } catch (err) { readErr = err; }
  MST(m.id).attempts++;
  const missing = Object.keys(want).filter((k) => !found[k]);
  logOf(m.from).push({ k: 'upload', who: 'me', mission: m.id, name: name || 'workbook.xlsx', summary: readErr ? 'could not be read' : `${sheets} sheets, ${Object.keys(want).length - missing.length} of ${Object.keys(want).length} totals found`, t: clock() });
  save();
  renderWall();
  if (!readErr && !missing.length) {
    MST(m.id).solved = 'PivotTable';
    await play(m.from, m.replies.pass.map((l, i, a) => ({ k: 'text', who: m.from, html: fill(l), react: i === a.length - 1 ? '🙏🎉' : '' })));
    await closeMission(m);
    return;
  }
  const reply = readErr ? 'Hmm, I cannot open that file 🤔 Is it saved as Excel Workbook (.xlsx)?' : `I cannot find the right total for <b>${esc(missing[0])}</b> 🤔 Is Net summed, and Distributor in Rows?`;
  await play(m.from, [{ k: 'text', who: m.from, html: reply }, { k: 'sari', who: 'sari', html: 'In the PivotTable field list, tick <b>Distributor</b> (it goes to Rows) and <b>Net</b> (it goes to Values as Sum). Save, then upload again.' }]);
}
function wireExcel(m) {
  $('#xdl').onclick = () => downloadXlsx(m);
  $('#xup').onchange = async (e) => { const f = e.target.files && e.target.files[0]; if (!f) return; await checkUpload(m, await f.arrayBuffer(), f.name); e.target.value = ''; };
  $('#grab').onclick = () => setDock(UI.dock === 'full' ? 'open' : 'full');
}

/* ---------- Power Query: a list of applied steps, run on a copy of the source table ---------- */
const STEP_LABEL = { promote: 'Use first row as headers', blank: 'Remove blank rows', trim: 'Trim', number: 'Change type to number', filter: 'Keep rows where', remove: 'Remove column', append: 'Append', merge: 'Merge', unpivot: 'Unpivot other columns', group: 'Group by' };
const NEEDS_COL = ['trim', 'number', 'filter', 'remove', 'merge', 'unpivot', 'group'];
/** Run steps on a copy of the source. A table is {head, rows}; before "use first row as headers" head is null. */
function runQuery(m, steps) {
  let t = JSON.parse(JSON.stringify(m.sources.main));
  const idx = (c) => (t.head || []).indexOf(c);
  for (const s of steps || []) {
    const i = idx(s.col);
    if (s.op === 'promote') { t = { head: t.rows[0].map(String), rows: t.rows.slice(1) }; }
    else if (s.op === 'blank') t.rows = t.rows.filter((r) => r.some((v) => String(v ?? '').trim() !== ''));
    else if (s.op === 'trim' && i >= 0) t.rows.forEach((r) => { r[i] = String(r[i] ?? '').replace(/\s+/g, ' ').trim(); });
    else if (s.op === 'number' && i >= 0) t.rows.forEach((r) => { r[i] = Number(String(r[i] ?? '').replace(/[^\d-]/g, '')) || 0; });
    else if (s.op === 'filter' && i >= 0) t.rows = t.rows.filter((r) => String(r[i]) === String(s.val));
    else if (s.op === 'remove' && i >= 0) { t.head.splice(i, 1); t.rows.forEach((r) => r.splice(i, 1)); }
    else if (s.op === 'append' && m.sources.other) t.rows = t.rows.concat(JSON.parse(JSON.stringify(m.sources.other.rows)));
    else if (s.op === 'merge' && i >= 0 && m.sources.lookup) { const L = m.sources.lookup; t.head.push(m.mergeCol); t.rows.forEach((r) => { const hit = L.rows.find((x) => x[0] === r[i]); r.push(hit ? hit[L.head.indexOf(m.mergeCol)] : null); }); }
    else if (s.op === 'unpivot' && i >= 0) { const head = t.head; const rows = []; t.rows.forEach((r) => head.forEach((h, j) => { if (j !== i) rows.push([r[i], h, r[j]]); })); t = { head: [head[i], m.unpivotNames[0], m.unpivotNames[1]], rows }; }
    else if (s.op === 'group' && i >= 0) { const j = idx(m.groupSum); const map = new Map(); t.rows.forEach((r) => map.set(r[i], (map.get(r[i]) || 0) + (Number(r[j]) || 0))); t = { head: [s.col, 'Sum of ' + m.groupSum], rows: [...map.entries()].sort((a, b) => String(a[0]).localeCompare(String(b[0]))) }; }
  }
  return t;
}
function queryTableHTML(t, max) {
  const n = max || 8;
  const head = t.head ? `<tr>${t.head.map((h) => `<th>${esc(h)}</th>`).join('')}</tr>` : `<tr>${(t.rows[0] || []).map((_, j) => `<th>Column${j + 1}</th>`).join('')}</tr>`;
  const body = t.rows.slice(0, n).map((r) => `<tr>${r.map((v) => `<td class="${typeof v === 'number' ? 'n' : ''}">${v === '' || v == null ? '<i class="nul">null</i>' : esc(typeof v === 'number' ? fmt(v) : String(v).replace(/^ +| +$| {2,}/g, (sp) => '·'.repeat(sp.length)))}</td>`).join('')}</tr>`).join('');
  return `<p class="bsum">${t.rows.length} rows, ${(t.head || t.rows[0] || []).length} columns</p><div class="qwrap"><table class="pt">${head}${body}</table></div>${t.rows.length > n ? `<p class="bempty">and ${t.rows.length - n} more rows</p>` : ''}`;
}
function queryControls(m, c) {
  const steps = c.steps || [];
  const t = runQuery(m, steps);
  const chip = (g, v, label, on) => `<button class="chip${on ? ' on' : ''}" data-g="${g}" data-v="${esc(v)}">${esc(label || v)}</button>`;
  const applied = `<ol class="steps"><li class="src">Source: ${esc(m.sources.main.name)}</li>${steps.map((st, k) => `<li>${esc(STEP_LABEL[st.op])}${st.col ? ' <b>' + esc(st.col) + '</b>' : ''}${st.val ? ' = <b>' + esc(st.val) + '</b>' : ''}${k === steps.length - 1 ? '<button data-g="undo" data-v="1" aria-label="Remove this step">×</button>' : ''}</li>`).join('')}</ol>`;
  let pick = '';
  if (UI.qOp && NEEDS_COL.includes(UI.qOp)) {
    pick = `<div class="brow put"><span>${esc(STEP_LABEL[UI.qOp])}: which column?</span>${(t.head || []).map((h) => chip('qcol', h, h)).join('')}</div>`;
    if (UI.qOp === 'filter' && UI.qCol) pick = `<div class="brow put"><span>Keep rows where <b>${esc(UI.qCol)}</b> is</span>${[...new Set(t.rows.map((r) => r[t.head.indexOf(UI.qCol)]))].slice(0, 8).map((v) => chip('qval', v, String(v))).join('')}</div>`;
  }
  return `<div class="bsec"><small>Applied steps</small>${applied}<small>Add a step</small><div class="brow">${m.ops.map((o) => chip('qop', o, STEP_LABEL[o], UI.qOp === o)).join('')}</div>${pick}</div>`;
}
const sameTable = (a, b) => JSON.stringify(a.head) === JSON.stringify(b.head) && JSON.stringify(a.rows) === JSON.stringify(b.rows);

/* ---------- Macros with one line to write: run the whole macro, compare what it did to the sheet ---------- */
const fullCode = (m, line) => m.code.map((l) => (l.trim() === '___' ? l.replace('___', line) : l)).join('\n');
/** Run a macro to the end on a copy of the sheet. Returns {sheet, fill, msgs} or {error}. */
function runMacro(m, line) {
  try { const mac = VBA.create(fullCode(m, line), JSON.parse(JSON.stringify(m.sheet))); const st = mac.runToEnd(); return { sheet: st.sheet, fill: st.fill, msgs: st.msgs }; }
  catch (err) { return { error: String(err && err.message || err) }; }
}
function macroSheetHTML(m, res) {
  const cols = m.sheetCols, sh = (res && res.sheet) || m.sheet, fl = (res && res.fill) || {};
  let h = `<table class="gx"><thead><tr><th class="cn"></th>${cols.map((c) => `<th>${c}</th>`).join('')}</tr></thead><tbody>`;
  for (let r = 1; r <= m.sheetRows; r++) { h += `<tr><th>${r}</th>${cols.map((c) => { const v = sh[c + r]; return `<td class="${typeof v === 'number' ? 'n' : ''}${r === 1 ? ' hd' : ''}${fl[c + r] ? ' yellow' : ''}">${v === undefined ? '' : esc(fmt(v))}</td>`; }).join('')}</tr>`; }
  return h + '</tbody></table>';
}
function vbaFixHTML(m) {
  const st = MST(m.id);
  const line = st.line || '';
  const code = m.code.map((l, i) => {
    if (l.trim() === '___') { const ind = l.match(/^\s*/)[0]; return `<div class="ln blank"><i>${i + 1}</i><span>${esc(ind)}<input id="vline" spellcheck="false" autocapitalize="off" autocomplete="off" placeholder="write this line" value="${esc(line)}" aria-label="The missing line"></span></div>`; }
    return `<div class="ln"><i>${i + 1}</i><span>${vbPaint(l) || ' '}</span></div>`;
  }).join('');
  return `<section class="dock ${LAP.matches ? 'open' : UI.dock === 'min' ? 'full' : UI.dock}" id="dock" aria-label="Macro">
    <button class="grab" id="grab" aria-label="Make the workbook bigger or smaller"></button>
    <div class="shd"><span class="xl m">X</span><span class="nm">${esc(m.file)}</span><small>· ${esc(m.macroName)}</small>${maxButton()}</div>
    <div class="mac"><div class="split"><div class="code-v" id="codev">${code}</div><div class="gw" id="mgw">${macroSheetHTML(m, st.res)}</div></div>
    <div class="brow">${m.choices.map((c) => `<button class="chip" data-line="${esc(c)}">${esc(c.trim())}</button>`).join('')}</div>
    ${st.res && st.res.msgs && st.res.msgs.length ? `<p class="bsum">MsgBox: ${esc(st.res.msgs[st.res.msgs.length - 1])}</p>` : ''}</div>
    <div class="mac-acts"><button class="btn ghost" id="vrun">${I('play')}Run it</button><button class="btn" id="vsend">${I('send')}Send to ${esc(PEOPLE[m.from].short)}</button></div></section>`;
}
function wireVbaFix(m) {
  const st = MST(m.id);
  const inp = $('#vline');
  inp.oninput = () => { st.line = inp.value; save(); };
  $('#dock').querySelectorAll('[data-line]').forEach((b) => { b.onclick = () => { inp.value = b.dataset.line.trim(); st.line = inp.value; save(); }; });
  $('#vrun').onclick = () => { st.res = runMacro(m, inp.value); save(); if (st.res.error) flash('Kicau could not run it: ' + st.res.error); $('#mgw').innerHTML = macroSheetHTML(m, st.res); };
  $('#vsend').onclick = () => sendVba(m);
  $('#grab').onclick = () => setDock(UI.dock === 'full' ? 'open' : 'full');
}
async function sendVba(m) {
  if (!m || UI.busy[m.from]) return;
  const st = MST(m.id);
  const line = (st.line || '').trim();
  if (!line) { flash('Write the missing line first.'); return; }
  st.attempts++;
  const mine = runMacro(m, line), want = runMacro(m, m.solutionLine);
  st.res = mine.error ? null : mine;
  const ok = !mine.error && JSON.stringify(mine.sheet) === JSON.stringify(want.sheet) && JSON.stringify(mine.fill) === JSON.stringify(want.fill) && JSON.stringify(mine.msgs) === JSON.stringify(want.msgs);
  logOf(m.from).push({ k: 'built', who: 'me', mission: m.id, text: line, html: mine.error ? '' : macroSheetHTML(m, mine), ok, t: clock() });
  save();
  renderWall();
  const g = $('#mgw'); if (g && !mine.error) g.innerHTML = macroSheetHTML(m, mine);
  if (ok) {
    st.solved = line;
    await play(m.from, m.replies.pass.map((l, i, a) => ({ k: 'text', who: m.from, html: fill(l), react: i === a.length - 1 ? '🙏🎉' : '' })));
    await closeMission(m);
    return;
  }
  await play(m.from, mine.error
    ? [{ k: 'sari', who: 'sari', html: 'Kicau could not run that line: ' + esc(mine.error) + ' Check the spelling and the brackets.' }]
    : [{ k: 'text', who: m.from, html: fill(m.replies.wrong) }, { k: 'sari', who: 'sari', html: m.nudge }]);
}

/* ---------- Report builders: PivotTable, chart, and rules (conditional formatting, data validation) ---------- */
const BUILDERS = ['pivot', 'chart', 'rule', 'query', 'goal'];
const FIELDS = ['Date', 'Distributor', 'Product', 'Category', 'Cartons', 'Net'];
const NUMERIC = ['Cartons', 'Net'];
const cfgOf = (m) => { const st = MST(m.id); st.cfg = st.cfg || JSON.parse(JSON.stringify(m.startCfg || {})); return st.cfg; };
/** Group the data the way a PivotTable does. Returns {rowKeys, colKeys, cell(r,c), total}. */
function pivot(data, c) {
  if (!c.rows || !c.agg || (c.agg !== 'Count' && !NUMERIC.includes(c.val))) return null;
  const keys = (f) => [...new Set(data.map((d) => d[f]))].sort((a, b) => (typeof a === 'number' ? a - b : String(a).localeCompare(String(b))));
  const rowKeys = keys(c.rows), colKeys = c.cols ? keys(c.cols) : [null];
  const agg = (set) => { if (c.agg === 'Count') return set.length; const s = set.reduce((a, d) => a + d[c.val], 0); return c.agg === 'Average' ? (set.length ? s / set.length : 0) : s; };
  const all = agg(data);
  const cell = (r, k) => { const v = agg(data.filter((d) => (r === null || d[c.rows] === r) && (k === null || d[c.cols] === k))); return c.pct ? v / all : v; };
  return { rowKeys, colKeys, cell, total: c.pct ? 1 : all };
}
const pv = (v, c) => (c.pct ? (v * 100).toFixed(1) + '%' : fmt(Math.round(v * 100) / 100));
function pivotHTML(m, c) {
  const p = pivot(m.data, c);
  if (!p) return '<p class="bempty">Put a field in <b>Rows</b> and a number in <b>Values</b> to see the PivotTable.</p>';
  const head = `<tr><th>${esc(c.rows)}</th>${p.colKeys.map((k) => `<th>${k === null ? esc((c.agg || '') + ' of ' + (c.agg === 'Count' ? 'rows' : c.val)) : esc(k)}</th>`).join('')}${c.cols ? '<th>Grand total</th>' : ''}</tr>`;
  const body = p.rowKeys.map((r) => `<tr><td>${esc(r)}</td>${p.colKeys.map((k) => `<td class="n">${pv(p.cell(r, k), c)}</td>`).join('')}${c.cols ? `<td class="n">${pv(p.cell(r, null), c)}</td>` : ''}</tr>`).join('');
  const tot = `<tr class="tot"><td>Grand total</td>${p.colKeys.map((k) => `<td class="n">${pv(p.cell(null, k), c)}</td>`).join('')}${c.cols ? `<td class="n">${pv(p.cell(null, null), c)}</td>` : ''}</tr>`;
  return `<table class="pt"><thead>${head}</thead><tbody>${body}${tot}</tbody></table>`;
}
const CHART_COLORS = ['#3E5BF2', '#F2B705', '#0F8A5F', '#E2603A', '#8B7CF6', '#06B6D4'];
function chartSVG(m, c) {
  if (!c.type || !c.cat || !c.val) return '<p class="bempty">Pick a chart type, a category and a value.</p>';
  const p = pivot(m.data, { rows: c.cat, val: c.val, agg: 'Sum' });
  if (!p) return '<p class="bempty">The value has to be a number column.</p>';
  const items = p.rowKeys.map((k) => ({ k, v: p.cell(k, null) }));
  const max = Math.max(...items.map((i) => i.v)) || 1;
  const W = 320, H = 190;
  if (c.type === 'Pie') {
    const tot = items.reduce((a, i) => a + i.v, 0);
    let a0 = -Math.PI / 2, out = '';
    items.forEach((it, i) => { const a1 = a0 + (it.v / tot) * Math.PI * 2; const big = a1 - a0 > Math.PI ? 1 : 0; const x0 = 95 + 80 * Math.cos(a0), y0 = 95 + 80 * Math.sin(a0), x1 = 95 + 80 * Math.cos(a1), y1 = 95 + 80 * Math.sin(a1); out += `<path d="M95 95L${x0.toFixed(1)} ${y0.toFixed(1)}A80 80 0 ${big} 1 ${x1.toFixed(1)} ${y1.toFixed(1)}Z" fill="${CHART_COLORS[i % 6]}"/>`; a0 = a1; });
    const legend = items.map((it, i) => `<g transform="translate(200 ${20 + i * 22})"><rect width="12" height="12" rx="3" fill="${CHART_COLORS[i % 6]}"/><text x="18" y="11">${esc(String(it.k).slice(0, 16))} ${Math.round((it.v / tot) * 100)}%</text></g>`).join('');
    return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Pie chart">${out}${legend}</svg>`;
  }
  const n = items.length, pad = 26, bw = (W - pad - 8) / n;
  if (c.type === 'Bar') {
    const bh = (H - 10) / n;
    return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Bar chart">${items.map((it, i) => `<rect x="110" y="${5 + i * bh + 3}" width="${((W - 120) * it.v) / max}" height="${bh - 6}" rx="3" fill="${CHART_COLORS[0]}"/><text x="104" y="${5 + i * bh + bh / 2 + 4}" text-anchor="end">${esc(String(it.k).slice(0, 15))}</text>`).join('')}</svg>`;
  }
  if (c.type === 'Line') {
    const pts = items.map((it, i) => `${pad + i * bw + bw / 2},${H - 24 - ((H - 40) * it.v) / max}`).join(' ');
    return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Line chart"><polyline points="${pts}" fill="none" stroke="${CHART_COLORS[0]}" stroke-width="3"/>${items.map((it, i) => `<text x="${pad + i * bw + bw / 2}" y="${H - 8}" text-anchor="middle">${esc(String(it.k).split(' ').pop().slice(0, 8))}</text>`).join('')}</svg>`;
  }
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Column chart">${items.map((it, i) => { const h = ((H - 40) * it.v) / max; return `<rect x="${pad + i * bw + 6}" y="${H - 24 - h}" width="${bw - 12}" height="${h}" rx="3" fill="${CHART_COLORS[0]}"/><text x="${pad + i * bw + bw / 2}" y="${H - 8}" text-anchor="middle">${esc(String(it.k).split(' ').pop().slice(0, 8))}</text>`; }).join('')}</svg>`;
}
/** Which rows a rule catches: a conditional format highlights them, data validation flags them. */
function ruleHits(m, c) {
  if (!c.col || !c.op || c.a === '' || c.a == null) return null;
  const a = +c.a, b = +c.b;
  const test = (v) => (c.op === '>' ? v > a : c.op === '<' ? v < a : c.op === '>=' ? v >= a : v >= a && v <= b);
  return m.data.map((d) => (m.rule === 'dv' ? !test(d[c.col]) : test(d[c.col])));
}
function ruleHTML(m, c) {
  const hits = ruleHits(m, c);
  const style = c.style || 'green';
  const rowsH = m.data.map((d, i) => `<tr class="${hits && hits[i] ? (m.rule === 'dv' ? 'bad' : 'hl-' + style) : ''}"><td>${esc(d.Distributor)}</td><td>${esc(d.Product)}</td><td class="n">${fmt(d.Cartons)}</td><td class="n">${fmt(d.Net)}</td></tr>`).join('');
  const sum = hits ? `<p class="bsum">${m.rule === 'dv' ? `${hits.filter(Boolean).length} rows break this rule and would be refused` : `${hits.filter(Boolean).length} cells highlighted`}</p>` : '<p class="bempty">Pick a column, a condition and a value.</p>';
  return `${sum}<table class="pt"><thead><tr><th>Distributor</th><th>Product</th><th>Cartons</th><th>Net</th></tr></thead><tbody>${rowsH}</tbody></table>`;
}
/** Goal Seek on the loan calculator: solve the input that makes PMT hit the target. */
function goalHTML(m, c) {
  const L = m.cells, r = L.B2 / 12, n = L.B3, P = L.B1;
  const pmt = (loan, rate, months) => (loan * rate) / (1 - Math.pow(1 + rate, -months));
  if (c.set !== 'B4' || !c.to || !c.change) return `<p class="bempty">Now: B4 = Rp ${fmt(Math.round(pmt(P, r, n)))} a month. Fill in the three boxes.</p>`;
  const t = +c.to;
  let out = '';
  if (c.change === 'B1') out = 'B1 = Rp ' + fmt(Math.round((t * (1 - Math.pow(1 + r, -n))) / r));
  else if (c.change === 'B3') out = 'B3 = ' + (Math.log(1 / (1 - (P * r) / t)) / Math.log(1 + r)).toFixed(1) + ' months';
  else out = c.change + ' cannot reach that by itself.';
  return `<p class="bsum">Goal Seek found a solution: ${esc(out)}</p><p class="bempty">B4 = Rp ${fmt(t)} a month.</p>`;
}
function buildControls(m, c) {
  if (m.kind === 'goal') return `<div class="bsec"><small>Set cell</small><div class="brow">${['B1', 'B3', 'B4'].map((v) => `<button class="chip${c.set === v ? ' on' : ''}" data-g="set" data-v="${v}">${v}</button>`).join('')}</div>
    <small>To value</small><div class="brow"><label class="numf"><input id="gto" inputmode="numeric" value="${esc(c.to ?? '')}"></label></div>
    <small>By changing cell</small><div class="brow">${['B1', 'B2', 'B3'].map((v) => `<button class="chip${c.change === v ? ' on' : ''}" data-g="change" data-v="${v}">${v}</button>`).join('')}</div></div>`;
  const chip = (group, val, on, label) => `<button class="chip${on ? ' on' : ''}" data-g="${group}" data-v="${esc(val)}">${esc(label || val)}</button>`;
  if (m.kind === 'pivot') {
    const zone = (z, label) => `<div class="zone"><small>${label}</small>${c[z] ? `<span class="tagf">${esc(z === 'val' && c.agg === 'Count' ? 'Count of rows' : c[z])}<button data-clear="${z}" aria-label="Remove">×</button></span>` : '<em>empty</em>'}</div>`;
    return `<div class="bsec"><small>Fields: tap one, then where it goes</small><div class="brow">${FIELDS.map((f) => chip('field', f, UI.bField === f)).join('')}</div>
      ${UI.bField ? `<div class="brow put"><span>Put <b>${esc(UI.bField)}</b> in</span>${chip('put', 'rows', false, 'Rows')}${m.noCols ? '' : chip('put', 'cols', false, 'Columns')}${chip('put', 'val', false, 'Values')}</div>` : ''}</div>
      <div class="zones">${zone('rows', 'Rows')}${m.noCols ? '' : zone('cols', 'Columns')}${zone('val', 'Values')}</div>
      <div class="brow">${['Sum', 'Count', 'Average'].map((a) => chip('agg', a, c.agg === a)).join('')}${m.showPct ? chip('pct', 'pct', !!c.pct, '% of grand total') : ''}</div>`;
  }
  if (m.kind === 'chart') return `<div class="bsec"><small>Chart type</small><div class="brow">${['Column', 'Bar', 'Line', 'Pie'].map((t) => chip('type', t, c.type === t)).join('')}</div>
    <small>Categories (axis)</small><div class="brow">${['Distributor', 'Product', 'Category'].map((f) => chip('cat', f, c.cat === f)).join('')}</div>
    <small>Values</small><div class="brow">${NUMERIC.map((f) => chip('val', f, c.val === f)).join('')}</div></div>`;
  return `<div class="bsec"><small>Apply to column</small><div class="brow">${NUMERIC.map((f) => chip('col', f, c.col === f)).join('')}</div>
    <small>${m.rule === 'dv' ? 'Allow only values' : 'Format cells where the value is'}</small><div class="brow">${(m.rule === 'dv' ? [['between', 'between']] : [['>', 'greater than'], ['<', 'less than'], ['>=', 'at least']]).map(([o, l]) => chip('op', o, c.op === o, l)).join('')}</div>
    <div class="brow"><label class="numf">${c.op === 'between' ? 'from' : 'value'} <input id="ra" inputmode="numeric" value="${esc(c.a ?? '')}"></label>${c.op === 'between' ? `<label class="numf">to <input id="rb" inputmode="numeric" value="${esc(c.b ?? '')}"></label>` : ''}</div>
    ${m.rule === 'dv' ? '' : `<small>Format</small><div class="brow">${[['green', 'Green fill'], ['red', 'Red fill'], ['bold', 'Bold']].map(([v, l]) => chip('style', v, (c.style || '') === v, l)).join('')}</div>`}</div>`;
}
function buildDockHTML(m) {
  const c = cfgOf(m);
  const preview = m.kind === 'goal' ? goalHTML(m, c) : m.kind === 'query' ? queryTableHTML(runQuery(m, c.steps)) : m.kind === 'pivot' ? pivotHTML(m, c) : m.kind === 'chart' ? chartSVG(m, c) : ruleHTML(m, c);
  const label = { goal: 'Goal Seek', query: 'Power Query', pivot: 'PivotTable', chart: 'Chart', rule: m.rule === 'dv' ? 'Data validation' : 'Conditional formatting' }[m.kind];
  return `<section class="dock ${LAP.matches ? 'open' : UI.dock === 'min' ? 'open' : UI.dock}" id="dock" aria-label="${label} builder">
    <button class="grab" id="grab" aria-label="Make the workbook bigger or smaller"></button>
    <div class="shd"><span class="xl">X</span><span class="nm">${esc(m.file)}</span><small>· ${label}</small>${dlButton()}${maxButton()}</div>
    <div class="bwrap" id="bwrap"><div id="bctl">${m.kind === 'query' ? queryControls(m, c) : buildControls(m, c)}</div><div class="bprev" id="bprev">${preview}</div></div>
    <div class="comp"><button class="coffee${coffeesLeft() ? '' : ' out'}" id="coffee" aria-label="Coffee with Mbak Sari">${I('coffee')}<i>${coffeesLeft()}</i></button><span class="sp"></span><button class="btn" id="bsend">${I('send')}Send to ${esc(PEOPLE[m.from].short)}</button></div></section>`;
}
function wireBuild(m) {
  const redraw = () => {
    const c = cfgOf(m);
    $('#bctl').innerHTML = m.kind === 'query' ? queryControls(m, c) : buildControls(m, c);
    $('#bprev').innerHTML = m.kind === 'goal' ? goalHTML(m, c) : m.kind === 'query' ? queryTableHTML(runQuery(m, c.steps)) : m.kind === 'pivot' ? pivotHTML(m, c) : m.kind === 'chart' ? chartSVG(m, c) : ruleHTML(m, c);
    save();
    wireInputs();
  };
  const wireInputs = () => {
    const c = cfgOf(m);
    const g = $('#gto'); if (g) g.oninput = () => { c.to = g.value.replace(/[^\d]/g, ''); save(); $('#bprev').innerHTML = goalHTML(m, c); };
    for (const id of ['ra', 'rb']) { const el = $('#' + id); if (el) el.oninput = () => { c[id === 'ra' ? 'a' : 'b'] = el.value.replace(/[^\d.]/g, ''); save(); $('#bprev').innerHTML = ruleHTML(m, c); }; }
  };
  $('#bctl').onclick = (e) => {
    const c = cfgOf(m);
    const clr = e.target.closest('[data-clear]');
    if (clr) { c[clr.dataset.clear] = null; redraw(); return; }
    const b = e.target.closest('[data-g]');
    if (!b) return;
    const g = b.dataset.g, v = b.dataset.v;
    if (m.kind === 'query') {
      c.steps = c.steps || [];
      if (g === 'undo') c.steps.pop();
      else if (g === 'qop') { if (NEEDS_COL.includes(v)) { UI.qOp = UI.qOp === v ? null : v; UI.qCol = null; } else { c.steps.push({ op: v }); UI.qOp = null; } }
      else if (g === 'qcol') { if (UI.qOp === 'filter') UI.qCol = v; else { c.steps.push({ op: UI.qOp, col: v }); UI.qOp = null; } }
      else if (g === 'qval') { c.steps.push({ op: 'filter', col: UI.qCol, val: v }); UI.qOp = null; UI.qCol = null; }
      redraw();
      return;
    }
    if (g === 'field') UI.bField = UI.bField === v ? null : v;
    else if (g === 'put') { c[v] = UI.bField; if (v === 'val' && !c.agg) c.agg = NUMERIC.includes(UI.bField) ? 'Sum' : 'Count'; UI.bField = null; }
    else if (g === 'pct') c.pct = !c.pct;
    else c[g] = v;
    redraw();
  };
  wireInputs();
  $('#bsend').onclick = () => sendBuild(m);
  $('#coffee').onclick = coffee;
  $('#grab').onclick = () => setDock(UI.dock === 'full' ? 'open' : 'full');
}
/** Compare the built report with what the colleague asked for, one setting at a time. */
function buildVerdict(m, c) {
  if (m.kind === 'query') {
    const t = runQuery(m, c.steps), w = runQuery(m, m.solutionSteps);
    if (sameTable(t, w)) return null;
    if (!t.head || JSON.stringify(t.head) !== JSON.stringify(w.head)) return 'cols';
    if (t.rows.length !== w.rows.length) return 'rows';
    return 'values';
  }
  for (const k of Object.keys(m.want)) {
    let a = c[k], b = m.want[k];
    if (k === 'val' && m.want.agg === 'Count') continue;
    if (k === 'a' || k === 'b' || k === 'to') { a = +a; b = +b; }
    if ((a || null) !== (b || null) && !(k === 'pct' && !a && !b)) return k;
  }
  return null;
}
function buildSummary(m, c) {
  if (m.kind === 'query') return (c.steps || []).length ? (c.steps || []).map((st) => STEP_LABEL[st.op] + (st.col ? ' ' + st.col : '') + (st.val ? ' = ' + st.val : '')).join(' → ') : 'No steps yet';
  if (m.kind === 'pivot') return `Rows: ${c.rows || '–'}${c.cols ? ', Columns: ' + c.cols : ''}, Values: ${c.agg === 'Count' ? 'Count' : (c.agg || '') + ' of ' + (c.val || '–')}${c.pct ? ', % of total' : ''}`;
  if (m.kind === 'goal') return `Set ${c.set || '?'} to ${c.to ? fmt(+c.to) : '?'} by changing ${c.change || '?'}`;
  if (m.kind === 'chart') return `${c.type || '?'} chart of ${c.val || '?'} by ${c.cat || '?'}`;
  return `${c.col || '?'} ${c.op === 'between' ? 'between ' + c.a + ' and ' + c.b : (c.op || '?') + ' ' + fmt(+c.a)}${m.rule === 'dv' ? '' : ', ' + (c.style || 'green')}`;
}
async function sendBuild(m) {
  if (!m || UI.busy[m.from]) return;
  const c = cfgOf(m), st = MST(m.id);
  st.attempts++;
  const miss = buildVerdict(m, c);
  logOf(m.from).push({ k: 'built', who: 'me', mission: m.id, text: buildSummary(m, c), html: m.kind === 'chart' ? chartSVG(m, c) : m.kind === 'pivot' ? pivotHTML(m, c) : m.kind === 'query' ? queryTableHTML(runQuery(m, c.steps), 5) : '', ok: !miss, t: clock() });
  save();
  renderWall();
  if (!miss) {
    st.solved = buildSummary(m, c);
    await play(m.from, m.replies.pass.map((l, i, a) => ({ k: 'text', who: m.from, html: fill(l), react: i === a.length - 1 ? '🙏🎉' : '' })));
    await closeMission(m);
    return;
  }
  await play(m.from, [{ k: 'text', who: m.from, html: fill(m.replies[miss] || m.replies.other) }, { k: 'sari', who: 'sari', html: m.nudges[miss] || m.nudges.other }]);
}

/* ---------- Toasts and the promotion ---------- */
function toast(m) {
  document.querySelectorAll('.toast').forEach((t) => t.remove());
  const P = PEOPLE[m.from];
  const t = document.createElement('a');
  t.className = 'toast';
  t.href = '#chat-' + m.from;
  t.innerHTML = `${avatar(m.from, 40, { on: true })}<div class="body"><b>${esc(P.name)}</b><span>${esc(strip(fill(m.intro[0])))}</span></div>${I('chat')}`;
  document.body.append(t);
  say('New message from ' + P.name);
  if (navigator.vibrate) try { navigator.vibrate(18); } catch (e) { /* no vibration */ }
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 320); }, FAST ? 700 : 5200);
  if (!location.hash || location.hash === '#desk') render();
}
/** The end of the course: every request closed. */
function showFinale() {
  if ($('.promo')) return;
  const firstTry = LIST.filter((m) => MST(m.id).attempts === 1).length;
  const cups = LIST.reduce((a, m) => a + MST(m.id).coffees, 0);
  const el = document.createElement('div');
  el.className = 'promo';
  el.setAttribute('role', 'dialog');
  el.setAttribute('aria-label', 'Course complete');
  el.innerHTML = `${confetti(90)}<div class="in"><div class="ring-av">${avatar('me', 108)}</div><small>Kenari Excel Lab · complete</small><h1>${esc(rankOf(ST.closed.length).title)}</h1><p>${LIST.length} requests, ${MS.CHAPTERS.length} chapters, two companies, and files that talk to each other. From SUM to a report that warned a bank in time.</p>
    <div class="kp"><span><b>${LIST.length}</b>requests</span><span><b>${firstTry}</b>first try</span><span><b>${cups}</b>${cups === 1 ? 'coffee' : 'coffees'}</span></div>
    <div class="unlock">${avatar('sari', 44, { on: true })}<div><b>Mbak Sari:</b> I had nothing left to teach you by Chapter 8 😄 Selamat, ${esc(ST.name)}.</div></div>
    <button class="btn" id="promook">Back to Kicau</button></div>`;
  document.body.append(el);
  $('#promook').focus();
  $('#promook').onclick = () => { el.remove(); location.hash = '#career'; render(); };
}
function showPromo() {
  if ($('.promo')) return;
  const firstTry = LIST.filter((m) => MST(m.id).attempts === 1).length;
  const cups = LIST.reduce((a, m) => a + MST(m.id).coffees, 0);
  const el = document.createElement('div');
  el.className = 'promo';
  el.setAttribute('role', 'dialog');
  el.setAttribute('aria-label', 'Promotion');
  el.innerHTML = `${confetti(70)}<div class="in"><div class="ring-av">${avatar('me', 108)}</div><small>Rasa Kenari · end of week one</small><h1>You are now an Analyst</h1><p>Eight requests closed for four colleagues.</p>
    <div class="kp"><span><b>${ST.closed.length}</b>requests</span><span><b>${firstTry}</b>first try</span><span><b>${cups}</b>${cups === 1 ? 'coffee' : 'coffees'}</span></div>
    <div class="unlock">${avatar('citra', 44, { on: true })}<div><b>Bank Kenari can message you now.</b><br>Citra and Dimas just wrote.</div></div>
    <button class="btn" id="promook">Open Kicau</button></div>`;
  document.body.append(el);
  if (navigator.vibrate) try { navigator.vibrate([20, 60, 30]); } catch (e) { /* no vibration */ }
  $('#promook').focus();
  $('#promook').onclick = () => { el.remove(); location.hash = '#chat-citra'; render(); };
}

/* ---------- Notes, career, you ---------- */
function renderNotes() {
  UI.person = null;
  const q = UI.notesQ || '';
  const items = LIST.map((m) => {
    const done = isClosed(m);
    if (q && !(m.note.fn + ' ' + m.note.what).toLowerCase().includes(q.toLowerCase())) return '';
    if (!done) return `<div class="note locked"><div class="h">${I('lock')}<b>${esc(m.note.fn)}</b><small>Request ${esc(m.id)}</small></div><p>Unlocks when you answer ${esc(PEOPLE[m.from].short)}.</p></div>`;
    const yours = MST(m.id).solved;
    return `<div class="note"><div class="h"><b>${esc(m.note.fn)}</b><small>Request ${esc(m.id)}</small></div><p>${esc(m.note.what)}</p><code class="ex">${esc(yours || m.note.shape)}</code><div class="from">${avatar(m.from, 24)}For ${esc(PEOPLE[m.from].name)}${MST(m.id).attempts === 1 ? ` · ${I('bolt')}first try` : ''}</div></div>`;
  }).join('');
  app.innerHTML = shell('notes', `<div class="page"><h1>Notes</h1><p class="lead">Every function you have used for a colleague, with the formula you sent. ${ST.closed.length} of ${LIST.length} so far.</p>
    <label class="search">${I('search')}<input id="nsearch" placeholder="Search notes" value="${esc(q)}" aria-label="Search notes"></label><div class="notes">${items || '<div class="empty">No notes match.</div>'}</div></div>`);
  wireSide();
  const s = $('#nsearch');
  s.oninput = () => { UI.notesQ = s.value; const pos = s.selectionStart; renderNotes(); const n = $('#nsearch'); n.focus(); n.setSelectionRange(pos, pos); };
}
const BADGES = [
  { id: 'first', icon: 'check', name: 'First reply', what: 'Close your first request', got: () => ST.closed.length >= 1 },
  { id: 'firsttry', icon: 'bolt', name: 'First try', what: 'Pass on your first send', got: () => LIST.some((m) => isClosed(m) && MST(m.id).attempts === 1) },
  { id: 'nocoffee', icon: 'coffee', name: 'No coffee', what: 'Close one without a hint', got: () => LIST.some((m) => isClosed(m) && MST(m.id).coffees === 0) },
  { id: 'pointer', icon: 'target', name: 'Pointer', what: 'Point at B1 instead of typing a name', got: () => isClosed(byId['1.5']) || isClosed(byId['1.6']) },
  { id: 'macro', icon: 'code', name: 'Macro reader', what: 'Step through a macro', got: () => isClosed(byId['1.8']) },
  { id: 'analyst', icon: 'trophy', name: 'Analyst', what: 'Finish week one', got: () => ST.promoted },
  { id: 'dollar', icon: 'lock', name: 'Pinned', what: 'Fix a reference with $', got: () => byId['2.2'] && isClosed(byId['2.2']) },
  { id: 'excel', icon: 'grid', name: 'Real Excel', what: 'Send back a file done in Excel', got: () => byId['2.11'] && isClosed(byId['2.11']) },
  { id: 'bank', icon: 'bank', name: 'Bank side', what: 'Finish Chapter 2', got: () => MS.CHAPTERS[1].ids.every((id) => isClosed(byId[id])) },
  { id: 'lookup', icon: 'search', name: 'Lookup', what: 'Finish Chapter 3', got: () => MS.CHAPTERS[2].ids.every((id) => isClosed(byId[id])) },
  { id: 'arrays', icon: 'spark', name: 'Spill', what: 'Finish Chapter 5', got: () => MS.CHAPTERS[4].ids.every((id) => isClosed(byId[id])) },
  { id: 'query', icon: 'undo', name: 'Clean data', what: 'Finish Chapter 7', got: () => MS.CHAPTERS[6] && MS.CHAPTERS[6].ids.every((id) => isClosed(byId[id])) },
  { id: 'loans', icon: 'bank', name: 'Loan pricer', what: 'Finish Chapter 8', got: () => MS.CHAPTERS[7] && MS.CHAPTERS[7].ids.every((id) => isClosed(byId[id])) },
  { id: 'macros', icon: 'code', name: 'Macro writer', what: 'Finish Chapter 9', got: () => MS.CHAPTERS[8] && MS.CHAPTERS[8].ids.every((id) => isClosed(byId[id])) },
  { id: 'grad', icon: 'trophy', name: 'Graduate', what: 'Finish all ten chapters', got: () => LIST.every(isClosed) },
  { id: 'dates', icon: 'flame', name: 'Aging', what: 'Finish Chapter 4', got: () => MS.CHAPTERS[3].ids.every((id) => isClosed(byId[id])) }
];
function renderCareer() {
  UI.person = null;
  const n = ST.closed.length;
  const firstTry = LIST.filter((m) => isClosed(m) && MST(m.id).attempts === 1).length;
  const cups = LIST.reduce((a, m) => a + MST(m.id).coffees, 0);
  const ranks = MS.RANKS.map((r, i) => {
    const nx = MS.RANKS[i + 1];
    const state = n >= (nx ? nx.at : Infinity) ? 'done' : n >= r.at ? 'now' : '';
    const sub = i === 0 ? 'Week one: Rasa Kenari requests' : r.title === 'Analyst' ? 'Bank Kenari starts writing. Chapter 2.' : r.title === 'Senior Analyst' ? 'Reports, PivotTables and charts.' : 'Macros and the early-warning report.';
    return `<div class="rank ${state}"><span class="dot">${state === 'done' ? I('check') : state === 'now' ? I('spark') : I('lock')}</span><div class="body"><b>${esc(r.title)}</b><small>${r.at ? `From ${r.at} requests. ` : ''}${esc(sub)}</small></div></div>`;
  }).join('');
  app.innerHTML = shell('career', `<div class="page"><h1>Career</h1><p class="lead">Your path at Rasa Kenari, and later Bank Kenari.</p>${hero()}
    <h2 class="sec">This week</h2><div class="kpis"><div class="kpi"><b>${n}</b><small>requests closed</small></div><div class="kpi"><b>${firstTry}</b><small>passed first try</small></div><div class="kpi"><b>${cups}</b><small>coffees with Sari</small></div><div class="kpi"><b>${streak()}</b><small>day streak</small></div></div>
    <h2 class="sec">Badges</h2><div class="badges">${BADGES.map((b) => `<div class="badge${b.got() ? '' : ' off'}"><span class="md">${I(b.icon)}</span>${esc(b.name)}<small>${esc(b.what)}</small></div>`).join('')}</div>
    <h2 class="sec">The ladder</h2><div class="ranks">${ranks}</div></div>`);
  wireSide();
}
const DEFAULT_MOOD = 'b'; /* Kenari Yellow is what a new learner sees first. */
const MOODS = [
  ['a', 'Kicau Blue', 'Light and calm', ['#F2F4F9', '#3E5BF2', '#0E1630']],
  ['b', 'Kenari Yellow', 'Light and warm', ['#F6F4EF', '#F2B705', '#1C1A14']],
  ['c', 'Night Shift', 'Dark, easy at night', ['#0B1020', '#7B93FF', '#18213A']]
];
function moodButtons(sel) {
  return `<div class="moods" role="group" aria-label="Color mood">${MOODS.map(([k, name, d, c]) => `<button class="mood" data-mood="${k}" aria-pressed="${sel === k}"><span class="sw" style="background:${c[0]}"><i style="left:8px;top:8px;width:34px;height:10px;border-radius:5px;background:${c[2]}"></i><i style="right:8px;top:24px;width:40px;height:12px;border-radius:6px;background:${c[1]}"></i><i style="left:8px;bottom:8px;width:30px;height:10px;border-radius:5px;background:${k === 'c' ? '#243052' : '#FFFFFF'};box-shadow:0 0 0 1px rgba(0,0,0,.06)"></i></span><span><b>${esc(name)}</b><small>${esc(d)}</small></span></button>`).join('')}</div>`;
}
function renderMe() {
  UI.person = null;
  app.innerHTML = shell('me', `<div class="page"><h1>You</h1><p class="lead">Junior Analysts change their mind about colors all the time. That is fine.</p>
    <div class="field"><label for="nm">What colleagues call you</label><input id="nm" maxlength="24" value="${esc(ST.name)}" autocomplete="given-name"></div>
    <h2 class="sec">Color mood</h2>${moodButtons(ST.mood)}
    <h2 class="sec">Progress</h2><div class="empty">Saved in this browser only. Another phone or laptop starts fresh. ${ST.closed.length} of ${LIST.length} requests closed.</div>
    <p style="margin-top:10px"><button class="btn ghost danger" id="wipe">Start week one again</button></p>
    <div class="about"><p>Kicau is part of Kenari Excel Lab. Rasa Kenari, Bank Kenari and everyone here are made up.</p><p>The practice sheet runs Formualizer 0.13.0 in your browser. Your name, answers and progress stay on this device. There is no account and no tracking.</p><p>Made by <a href="https://jeremysubiakto.github.io/links/" target="_blank" rel="noopener">Jeremy Subiakto</a> · Also try <a href="https://jeremysubiakto.github.io/kenari-sql-lab/" target="_blank" rel="noopener">Kenari Collection SQL Lab</a></p><p><a href="https://github.com/JeremySubiakto/kenari-excel-lab/issues/new?template=problem.md" target="_blank" rel="noopener">Report a problem</a> · Code under MIT, lessons under CC BY-NC 4.0.</p></div></div>`);
  wireSide();
  $('#nm').oninput = (e) => { ST.name = e.target.value.trim() || ST.name; save(); };
  document.querySelectorAll('[data-mood]').forEach((b) => { b.onclick = () => { ST.mood = b.dataset.mood; document.documentElement.dataset.v = ST.mood; save(); renderMe(); }; });
  $('#wipe').onclick = (e) => {
    const b = e.currentTarget;
    if (!b.dataset.sure) { b.dataset.sure = '1'; b.textContent = 'Tap again to clear notes and progress'; setTimeout(() => { if (b.isConnected) { delete b.dataset.sure; b.textContent = 'Start week one again'; } }, 5000); return; }
    { const keep = { name: ST.name, mood: ST.mood }; ST = Object.assign(fresh(), keep, { started: localDay() }); seedSari(); save(); location.hash = '#desk'; render(); } };
}

/* ---------- First run ---------- */
function seedSari() {
  ST.logs.sari = [
    { k: 'text', who: 'sari', html: 'Hai {name}! I am Sari, Senior Analyst. I will be your mentor this week 👋', t: clock() },
    { k: 'text', who: 'sari', html: 'Colleagues will send you requests here. Each one comes with a workbook. Tap cells to build a formula, then send it.', t: clock() },
    { k: 'text', who: 'sari', html: 'Stuck? Tap the ☕ next to the formula bar and I will come over. Three coffees a day, and I still help after that, I just get sleepy 😴', t: clock() },
    { k: 'text', who: 'sari', html: 'Tika from Sales Finance is first. Good luck!', t: clock() }
  ];
}
function renderWelcome() {
  const pickMood = UI.pickMood = ST.mood || UI.pickMood || DEFAULT_MOOD;
  document.documentElement.dataset.v = pickMood;
  app.innerHTML = `<div class="welcome"><div class="inner"><div class="mark">${I('bird')}</div><h1>Selamat datang di Rasa Kenari</h1><p class="sub">Your first week as a Junior Analyst starts now. Colleagues send requests in Kicau, and you answer them in Excel.</p>
    <div class="bubbles"><div class="b them">${avatar('sari', 26)} <b>Mbak Sari</b><br>Hai! I am your mentor. Before you meet the team, two quick things 👇</div></div>
    <div class="field"><label for="wname">What should colleagues call you?</label><input id="wname" maxlength="24" placeholder="Your first name" value="${esc(UI.wname || ST.name || '')}" autocomplete="given-name"></div>
    <h2>Pick your color mood</h2>${moodButtons(pickMood)}
    <button class="btn wide" id="go" ${pickMood && (UI.wname || ST.name) ? '' : 'disabled'}>Go to my desk</button></div></div>`;
  const nm = $('#wname');
  const ready = () => { $('#go').disabled = !(nm.value.trim() && (UI.pickMood || ST.mood)); };
  nm.oninput = () => { UI.wname = nm.value; ready(); };
  document.querySelectorAll('[data-mood]').forEach((b) => { b.onclick = () => { UI.pickMood = b.dataset.mood; document.documentElement.dataset.v = UI.pickMood; document.querySelectorAll('[data-mood]').forEach((x) => x.setAttribute('aria-pressed', String(x === b))); ready(); }; });
  $('#go').onclick = () => {
    ST.name = nm.value.trim();
    ST.mood = UI.pickMood || ST.mood;
    ST.started = localDay();
    seedSari();
    save();
    location.hash = '#desk';
    render();
    setTimeout(() => toast(current()), FAST ? 0 : 900);
  };
  if (!coarse) nm.focus();
}

/* ---------- Routing ---------- */
/* ---------- Practice: 500 problems from easy to hard ----------
   The problems live in assets/js/kicau/practice/ and load the first time Practice opens, so the chat and the opening
   never wait for them. Each problem is judged like a request: on the sheet you see and on a hidden one. */
let practiceP = null;
function loadPractice() {
  if (window.KicauPractice && window.KicauPractice.__ready) return Promise.resolve(window.KicauPractice);
  if (practiceP) return practiceP;
  const base = new URL('assets/js/kicau/practice/', document.baseURI).href;
  const add = (f) => new Promise((res, rej) => { const s = document.createElement('script'); s.src = base + f; s.onload = res; s.onerror = () => rej(new Error('Could not load ' + f)); document.head.appendChild(s); });
  practiceP = add('core.js')
    /* One topic file failing to load never takes the rest of Practice down. */
    .then(() => Promise.all(window.KicauPractice.FILES.map((f) => add(f.file).catch((e) => console.warn(e.message)))))
    .then(() => add('index.js').catch(() => null))
    .then(() => { window.KicauPractice.__ready = true; return window.KicauPractice; })
    .catch((e) => { practiceP = null; throw e; });
  return practiceP;
}
const DIFF = { Easy: 'e', Medium: 'm', Hard: 'h' };
function practiceStats(L) {
  const pr = PR(), by = { Easy: [0, 0], Medium: [0, 0], Hard: [0, 0] };
  for (const x of L) { by[x.difficulty][1]++; if (pr.solved[x.id]) by[x.difficulty][0]++; }
  return { by, solved: Object.keys(pr.solved).length, total: L.length };
}
function renderPractice() {
  UI.person = null; UI.prac = null;
  app.innerHTML = shell('practice', `<div class="page practice" id="practice"><h1>Practice</h1><p class="lead">Loading the problems…</p></div>`);
  wireSide();
  loadPractice().then(drawPractice).catch(() => { const el = $('#practice'); if (el) el.innerHTML = '<h1>Practice</h1><p class="lead">The problems did not load. Check your connection and open Practice again.</p>'; });
}
function drawPractice(P) {
  const el = $('#practice');
  if (!el) return;
  const L = P.list(), st = practiceStats(L), pr = PR();
  const f = UI.pf = UI.pf || { d: 'All', s: 'All', t: '', q: '' };
  const topics = [...new Set(L.flatMap((x) => x.topics))].sort();
  const bar = (d) => { const [a, b] = st.by[d]; return `<div class="pbar ${DIFF[d]}"><span><b>${d}</b> ${a}<em>/${b}</em></span><i><s style="width:${b ? (a / b) * 100 : 0}%"></s></i></div>`; };
  const seg = (k, vals) => `<div class="seg" role="group">${vals.map((v) => `<button type="button" data-f="${k}" data-v="${v}" aria-pressed="${f[k] === v}">${v}</button>`).join('')}</div>`;
  el.innerHTML = `<div class="pr-hd"><div><h1>Practice</h1><p class="lead">${L.length} Excel problems, from easy to hard. Each one is checked on the sheet you see and on a hidden test, the same file a day later.</p></div>
    <div class="pr-stats">${ring(st.solved, st.total, 'pr-ring')}<div class="pbars">${bar('Easy')}${bar('Medium')}${bar('Hard')}</div></div></div>
    <div class="pr-tools"><input class="search" id="pq" type="search" placeholder="Search problems" value="${esc(f.q)}" aria-label="Search problems">
      ${seg('d', ['All', 'Easy', 'Medium', 'Hard'])}${seg('s', ['All', 'Todo', 'Solved'])}
      <select id="pt" aria-label="Topic"><option value="">All topics</option>${topics.map((t) => `<option${f.t === t ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select>
      <button type="button" class="btn sm" id="ppick">${I('spark')}Pick one</button></div>
    <div class="pr-list" id="plist" role="list"></div>`;
  const draw = () => {
    const q = f.q.trim().toLowerCase();
    const rows = L.filter((x) => (f.d === 'All' || x.difficulty === f.d) && (f.s === 'All' || (f.s === 'Solved') === !!pr.solved[x.id]) && (!f.t || x.topics.includes(f.t)) && (!q || String(x.n) === q || (x.title || '').toLowerCase().includes(q) || x.topics.some((t) => t.toLowerCase().includes(q))));
    $('#plist').innerHTML = rows.length ? rows.map((x) => `<a class="pr-row${pr.solved[x.id] ? ' done' : pr.tried[x.id] ? ' tried' : ''}" href="#p-${x.n}" role="listitem"><span class="pst" aria-label="${pr.solved[x.id] ? 'Solved' : pr.tried[x.id] ? 'Attempted' : 'Not started'}">${pr.solved[x.id] ? I('check') : pr.tried[x.id] ? '•' : ''}</span><span class="pn">${x.n}.</span><span class="pt">${esc(x.title || 'Problem ' + x.n)}</span><span class="ptags">${x.topics.slice(0, 2).map((t) => `<i>${esc(t)}</i>`).join('')}</span><span class="pd ${DIFF[x.difficulty]}">${x.difficulty}</span></a>`).join('') : '<div class="empty">No problem matches. Clear a filter.</div>';
  };
  draw();
  el.querySelectorAll('[data-f]').forEach((b) => { b.onclick = () => { f[b.dataset.f] = b.dataset.v; el.querySelectorAll(`[data-f="${b.dataset.f}"]`).forEach((x) => x.setAttribute('aria-pressed', String(x === b))); draw(); }; });
  $('#pq').oninput = (e) => { f.q = e.target.value; draw(); };
  $('#pt').onchange = (e) => { f.t = e.target.value; draw(); };
  $('#ppick').onclick = () => { const todo = L.filter((x) => !pr.solved[x.id] && (f.d === 'All' || x.difficulty === f.d)); const pick = todo.length ? todo[Math.floor(Math.random() * todo.length)] : L[0]; location.hash = '#p-' + pick.n; };
}
function renderProblem(n) {
  UI.person = null;
  app.innerHTML = shell('practice', `<div class="page practice" id="practice"><p class="lead">Loading problem ${n}…</p></div>`);
  wireSide();
  loadPractice().then((P) => {
    const meta = P.list()[n - 1];
    if (!meta) { location.hash = '#practice'; return; }
    drawProblem(P.get(meta.id), P);
  }).catch(() => { const el = $('#practice'); if (el) el.innerHTML = '<p class="lead">The problem did not load. Check your connection and try again.</p>'; });
}
function practiceDockHTML(m) {
  const chipBtns = `${m.fnKey ? `<button class="chip fn" data-key="${esc(m.fnKey)}">${esc(m.fnKey)}</button>` : ''}${(m.chips || []).map((ch) => `<button class="chip" data-key="${esc(ch)}">${esc(ch)}</button>`).join('')}`;
  const keys = coarse ? `<div class="keys" id="keys">${chipBtns}<span class="sep"></span>${['=', ',', '(', ')', ':', '$', '"', '>', '<', '>=', '*', '/', '+', '-', '&'].map((k) => `<button type="button" data-key="${esc(k)}">${esc(k)}</button>`).join('')}<button type="button" class="wd" data-act="back" aria-label="Delete one character">⌫</button><button type="button" class="wd" data-act="clear">Clear</button><button type="button" class="wd" data-act="abc" aria-label="Show the phone keyboard">ABC</button></div>` : '';
  const where = m.spill ? `Type one formula in <b>${esc(m.answerCell)}</b>. It spills by itself.` : m.fill ? `Type it in <b>${esc(m.answerCell)}</b>. Submit fills it down to <b>${esc(m.fill.col + m.fill.last)}</b>.` : `Answer cell <b>${esc(m.answerCell)}</b>`;
  return `<section class="dock ${LAP.matches ? 'open' : UI.dock === 'min' ? 'open' : UI.dock}" id="dock" aria-label="Workbook">
    <button class="grab" id="grab" aria-label="Make the workbook bigger or smaller"></button>
    <div class="shd"><span class="xl">X</span><span class="nm">${esc(m.file)}</span><small>· problem ${m.n}</small>${LAP.matches ? maxButton() : `<button class="icb" id="dockbig" aria-label="Full screen sheet">${I(UI.dock === 'full' ? 'collapse' : 'expand')}</button>`}</div>
    <div class="fbar"><span class="nbox" title="Name box: the cell your answer goes in">${esc(m.answerCell)}</span><span class="fxi" aria-hidden="true">fx</span><div class="fxbox"><div class="paint" id="paint" aria-hidden="true"></div><input id="fx" autocomplete="off" autocapitalize="characters" spellcheck="false" aria-label="Formula for ${esc(m.answerCell)}" ${coarse ? 'inputmode="none"' : ''}></div></div>
    <div class="fxlabel">${where}</div>
    <div class="gw" id="gw">${gridHTML(m)}</div>
    ${SHEET_TABS}
    ${coarse ? '' : `<div class="chips" id="chips"><small>Tap a cell, or</small>${chipBtns}</div>`}
    ${keys}
    <div class="engmsg" id="engmsg" ${ENG.api ? 'hidden' : ''}>${ENG.failed ? 'The practice sheet did not load. Reload to try again.' : 'Warming up the practice sheet…'}</div>
    <div class="comp"><button class="btn sm ghost" id="preset" type="button">${I('undo')}Reset</button><span class="sp"></span><button class="send" id="send" aria-label="Submit">${I('check')}<span>Submit</span></button></div></section>`;
}
function drawProblem(m, P) {
  UI.prac = m;
  const L = P.list(), pr = PR(), solved = pr.solved[m.id];
  const prev = L[m.n - 2], next = L[m.n];
  const desc = `<section class="chatcol pdesc" aria-label="Problem ${m.n}">
    <div class="pdesc-in">
      <a class="pback" href="#practice">${I('back')}All problems</a>
      <h1>${m.n}. ${esc(m.title)}</h1>
      <div class="pmeta"><span class="pd ${DIFF[m.difficulty]}">${m.difficulty}</span>${m.topics.map((t) => `<i class="ptag">${esc(t)}</i>`).join('')}${solved ? `<span class="pdone">${I('check')}Solved</span>` : ''}</div>
      <div class="pst-body">${m.statement}</div>
      <p class="phid">${I('lock')}Checked on this sheet and on a hidden test: the same file a day later, with some numbers changed.</p>
      <div class="pres" id="pres" role="status" aria-live="polite"></div>
      <div class="phints">${(m.hints || []).map((h, i) => `<details><summary>Hint ${i + 1}</summary><p>${esc(h)}</p></details>`).join('')}</div>
      <div class="psol" id="psol">${solved ? solutionHTML(m) : `<details class="locked"><summary>Solution</summary><p>Unlocks when you solve it, or <button type="button" class="linkish" id="pshow">show it now</button>.</p></details>`}</div>
      <nav class="pnav">${prev ? `<a class="btn sm ghost" href="#p-${prev.n}">${I('back')}${prev.n}</a>` : '<span></span>'}<button type="button" class="btn sm ghost" id="ppick2">${I('spark')}Pick one</button>${next ? `<a class="btn sm" href="#p-${next.n}">Next: ${next.n}</a>` : '<span></span>'}</nav>
    </div></section>`;
  app.innerHTML = shell('prob', `<div class="chatwrap prob">${desc}${practiceDockHTML(m)}</div>`);
  wireSide();
  wireDock(m);
  if (!LAP.matches) setDock(UI.dock === 'min' ? 'open' : UI.dock);
  wireMax();
  const ps = $('#pshow');
  if (ps) ps.onclick = () => { $('#psol').innerHTML = solutionHTML(m); };
  $('#preset').onclick = () => { const st = MST(m.id); st.draft = ''; st.shownVals = null; st.shown = undefined; st.shownOk = false; save(); setFormula(m.starter || '='); const gw = $('#gw'); if (gw) gw.innerHTML = gridHTML(m); $('#pres').innerHTML = ''; };
  $('#ppick2').onclick = () => { const todo = L.filter((x) => !pr.solved[x.id] && x.id !== m.id); const pick = todo.length ? todo[Math.floor(Math.random() * todo.length)] : L[0]; location.hash = '#p-' + pick.n; };
  if (solved && !MST(m.id).draft) setFormula(solved.formula);
}
function solutionHTML(m) {
  return `<h2>Solution</h2><p>${esc(m.explain || '')}</p><pre class="code"><code>${esc(m.solution)}</code></pre>${(m.alts || []).length ? `<p class="palts">Also accepted:</p><ul class="palts">${m.alts.map((a) => `<li><code>${esc(a)}</code></li>`).join('')}</ul>` : ''}`;
}
/** Submit: judge on both sheets, show the verdict, and draw your results on the sheet. Nothing moves on by itself. */
function submitPractice(m) {
  const box = $('#pres');
  if (!ENG.api) { const e = $('#engmsg'); if (e) { e.hidden = false; e.textContent = ENG.failed ? 'The practice sheet did not load. Reload to try again.' : 'Warming up the practice sheet. Submit again in a moment.'; } return; }
  const typed = $('#fx').value;
  const v = Check.judge(m, typed, run(m));
  window.__lastVerdict = v;
  const pr = PR(), st = MST(m.id);
  if (v.state === 'empty') { box.innerHTML = `<div class="verdict err"><b>Nothing to submit</b><p>${esc(v.plain)}</p></div>`; return; }
  if (v.state === 'error') { box.innerHTML = `<div class="verdict err"><b>${esc(v.title)}</b><p>${esc(v.plain)}</p>${v.fix ? `<p>Try <code>${esc(v.fix.formula)}</code></p>` : ''}</div>`; return; }
  pr.tried[m.id] = (pr.tried[m.id] || 0) + 1;
  const formula = Check.normalize(typed);
  let colVals = null, shown = v.value !== undefined ? v.value : v.typed;
  if ((m.fill || m.spill) && formula[0] === '=') colVals = fillValues(m, formula, 'base');
  st.shown = shown; st.shownVals = colVals; st.shownOk = v.state === 'pass'; st.draft = typed;
  fillRefs(m).forEach((ref, i) => { const td = $(`#gw td[data-ref="${ref}"]`); if (!td) return; td.textContent = colVals ? fmtAns(m, colVals[i]) : fmtAns(m, shown); td.className = v.state === 'pass' ? 'got' : 'miss'; });
  if (v.state === 'pass') {
    const first = !pr.solved[m.id];
    pr.solved[m.id] = { at: Date.now(), formula };
    save();
    box.innerHTML = `<div class="verdict ok"><b>${I('check')}Accepted</b><p>Right on this sheet and on the hidden test.${first ? ` ${Object.keys(pr.solved).length} of 500 solved.` : ''}</p></div>`;
    $('#psol').innerHTML = solutionHTML(m);
    const md = document.querySelector('.pmeta');
    if (md && !md.querySelector('.pdone')) md.insertAdjacentHTML('beforeend', `<span class="pdone">${I('check')}Solved</span>`);
    say('Accepted');
    return;
  }
  save();
  let detail = esc(v.plain || '');
  if (v.code === 'frozen') detail = 'Right on this sheet, wrong on the hidden test (the same file a day later). Point at cells and ranges instead of typing numbers or names.';
  else if (v.code !== 'typed' && v.target !== undefined) {
    if (colVals) {
      const want = fillValues(m, m.solution, 'base');
      const i = want.findIndex((w, k) => !Check.same(colVals[k], w, m.precision, m.exactCase));
      if (i >= 0) detail = m.spill ? `Your result differs from the expected one at item ${i + 1}: you have ${esc(fmtAns(m, colVals[i]))}, expected ${esc(fmtAns(m, want[i]))}.` : `Row ${m.fill.first + i}: you have ${esc(fmtAns(m, colVals[i]))}, expected ${esc(fmtAns(m, want[i]))}.`;
    } else detail = `Your result is ${esc(fmtAns(m, v.value))}. Expected ${esc(fmtAns(m, v.target))}.`;
  }
  box.innerHTML = `<div class="verdict bad"><b>${v.code === 'typed' ? 'Typed value' : v.code === 'frozen' ? 'Wrong answer on the hidden test' : 'Wrong answer'}</b><p>${detail}</p></div>`;
  say(v.code === 'frozen' ? 'Wrong answer on the hidden test' : 'Wrong answer');
}

function render() {
  if (!ST.name || !ST.mood) { renderWelcome(); return; }
  UI.prac = null;
  /* Older #/chat/tika style links still work. */
  const h = location.hash.replace(/^#\/?/, '').replace(/^chat\//, 'chat-');
  if (h.startsWith('chat-')) renderChat(h.slice(5));
  else if (h === 'practice') renderPractice();
  else if (/^p-\d+$/.test(h)) renderProblem(+h.slice(2));
  else if (h === 'notes') renderNotes();
  else if (h === 'career') renderCareer();
  else if (h === 'you' || h === 'me') renderMe();
  else renderHome();
}
window.addEventListener('hashchange', render);
/* A link to the screen already open still refreshes it, so "Next" works when the same colleague writes again. */
document.addEventListener('click', (e) => {
  const a = e.target.closest && e.target.closest('a[href^="#"]');
  if (a && a.getAttribute('href') === (location.hash || '#desk')) { e.preventDefault(); render(); }
});
LAP.addEventListener('change', render);
if (window.visualViewport) {
  const fit = () => { app.style.height = visualViewport.height + 'px'; };
  visualViewport.addEventListener('resize', fit);
}
/* The opening (Mix 1, "The whole day") plays once per visit, over the app as it loads. ?fast and ?nointro skip it,
   ?intro plays it again. A tap, any key or the Skip button ends it; reduced motion gets a calm fade. */
(function opening() {
  if (typeof window.kcIntro !== 'function' || FAST || Q.has('nointro')) return;
  let seen = false;
  try { seen = sessionStorage.getItem('kicau.intro') === '1'; } catch (e) { /* storage blocked: play it */ }
  if (seen && !Q.has('intro')) return;
  try { sessionStorage.setItem('kicau.intro', '1'); } catch (e) { /* storage blocked */ }
  try {
    window.kcIntro(() => { engineGo(); const n = $('#wname'); if (n && !coarse) n.focus({ preventScroll: true }); });
    UI.introPlaying = true;
    /* The overlay now covers the screen; the page underneath can show, so the curtain reveals the welcome screen. */
    requestAnimationFrame(() => document.documentElement.classList.remove('kc-pre'));
  } catch (e) { /* the app works without it */ }
})();
if (!UI.introPlaying) { engineGo(); document.documentElement.classList.remove('kc-pre'); } else setTimeout(engineGo, 7000); /* a paused background tab still gets its engine */
render();

/* For the browser tests. */
window.__kicau = { setLine: (l) => { MST(current().id).line = l; }, sendVba: () => sendVba(current()), runQuery, setSteps: (steps) => { cfgOf(current()).steps = JSON.parse(JSON.stringify(steps)); }, setCfg: (cfg) => { const m = current(); Object.assign(cfgOf(m), cfg); }, sendBuild: () => sendBuild(current()), pivot, get ST() { return ST; }, LIST, MS, send, setFormula, coffee, pick, current, engine: enginePromise, run: (m) => run(m), Check, fillValues, uploadBuffer: (buf, name) => checkUpload(current(), buf, name || 'koleksi_oktober.xlsx'), buildXlsx: (m) => buildXlsx(m || current()), buildBooks: (m) => buildBooks(m || current()), cont, refsIn, get UI() { return UI; }, loadPractice, submitPractice, PR };
