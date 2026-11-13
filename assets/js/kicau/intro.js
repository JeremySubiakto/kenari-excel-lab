/*
 * Kicau opening, Mix 1: "The whole day". One story in three acts, each seam carried by an element.
 * Act 1 (0 to 0.85 s, from Option B): on the dark stage, Tika, Bayu and Citra pop in with typing bubbles that open into
 *   short requests, while the red badge counts 1, 2, 3.
 * Seam 1: the stage brightens and the bubbles fly into a desktop-Excel window that tilts in. Tika's bubble reshapes into
 *   the formula bar and her words land in it; Bayu's and Citra's become the Orders and Loans sheet tabs.
 * Act 2 (to 2.1 s, from Option A): =SUMIFS types itself over Tika's words with its ranges lit, and the three sales fly into E1.
 * Seam 2: the answer cell fills with gold and reshapes into the canary (a clip-path polygon morph).
 * Act 3 (to 4.3 s, from Option C): the canary takes off under the three-yellow sweep, swoops back and perches; it chirps
 *   while "Kicau" and "Kenari Excel Lab" spring in; the yellow curtain tears open on the diagonal to the welcome screen.
 * window.kcIntro(onDone) builds #kc-intro, plays it (about 4.3 s, plus at most 0.07 s waiting for fonts), removes it, then calls onDone.
 * Only transform, opacity, clip-path and SVG stroke are animated (Web Animations API, no libraries).
 * Skip: the Skip button, any key, or a tap anywhere (250 ms fade). Reduced motion (or ?reduce=1): a 400 ms fade of the lockup.
 */
(function () {
  'use strict';

  var BIRD = 'M15.2 4.2a4.3 4.3 0 0 0-4.3 4.3v1.1C7.6 10 4.8 12.2 3.2 16c2.7-.7 4.8-.5 6.5.6L8.4 20.4h2.5l1.3-2.8c3.9-.5 6.9-3.6 6.9-7.5V9.2l2.6-1.4-2.4-1.5a4.3 4.3 0 0 0-4.1-2.1Z';
  var WING = 'M11.6 10.2C8.9 10.3 6.1 11.7 3.9 14.5l2.3-.5-.7 1.4 2.4-1-.2 1.3 2.5-1.8c.9-.7 1.4-2 1.4-3.7Z';
  var INK = '#1C1A14';

  var EXPO = 'cubic-bezier(.16,1,.3,1)';
  var QUART = 'cubic-bezier(.25,1,.5,1)';
  var SOFT = 'cubic-bezier(.22,1,.36,1)';
  var BACK = 'cubic-bezier(.34,1.56,.64,1)';
  var INOUT = 'cubic-bezier(.65,0,.35,1)';
  var IN = 'cubic-bezier(.5,0,.75,0)';
  var OUTIN = 'cubic-bezier(.4,0,1,1)';
  var IO = 'cubic-bezier(.45,0,.55,1)';
  var DOCK = 'cubic-bezier(.55,0,.15,1)';      // a request flying into the workbook
  var SWEEP = 'cubic-bezier(.62,0,.22,1)';
  var CURTAIN = 'cubic-bezier(.83,0,.17,1)';
  var HOLD = 'steps(1,end)';

  var LINEAR_OK = false;
  try { LINEAR_OK = !!(window.CSS && CSS.supports && CSS.supports('animation-timing-function', 'linear(0, 0.5, 1)')); } catch (e) { /* old engine */ }

  /* A real damped spring as a CSS linear() easing (zeta sets the bounce); falls back to a back-out curve. */
  function spring(zeta, ms) {
    if (!LINEAR_OK) return BACK;
    var T = ms / 1000, w = 6.4 / (zeta * T), wd = w * Math.sqrt(1 - zeta * zeta), pts = [];
    for (var i = 0; i <= 48; i++) {
      var t = T * i / 48;
      pts.push((1 - Math.exp(-zeta * w * t) * (Math.cos(wd * t) + (zeta * w / wd) * Math.sin(wd * t))).toFixed(4));
    }
    pts[48] = '1';
    return 'linear(' + pts.join(',') + ')';
  }
  /* The same spring sampled into keyframes, for values that map to several properties at once. */
  function springKF(dur, zeta, map, n) {
    n = n || 40;
    var w0 = 5.2 / (zeta * dur / 1000), wd = w0 * Math.sqrt(1 - zeta * zeta), kf = [];
    for (var i = 0; i <= n; i++) {
      var t = (i / n) * dur / 1000;
      var p = i === n ? 1 : 1 - Math.exp(-zeta * w0 * t) * (Math.cos(wd * t) + (zeta * w0 / wd) * Math.sin(wd * t));
      var k = map(p); k.offset = i / n; kf.push(k);
    }
    return kf;
  }
  /* A spring by stiffness and damping, with the time it takes to settle. */
  function springKD(k, c) {
    var x = 0, v = 0, dt = 1 / 300, t = 0, s = [0];
    while (t < 2) {
      v += (-k * (x - 1) - c * v) * dt; x += v * dt; t += dt; s.push(x);
      if (t > .15 && Math.abs(1 - x) < .004 && Math.abs(v) < .06) break;
    }
    var out = [];
    for (var i = 0; i <= 48; i++) out.push(+s[Math.round(i / 48 * (s.length - 1))].toFixed(4));
    out[48] = 1;
    return { e: LINEAR_OK ? 'linear(' + out.join(', ') + ')' : BACK, d: Math.round(t * 1000) };
  }
  /* [ms, props, easingToNext] stops into keyframes on a shared clock that runs to the last stop. */
  function kfs(stops) {
    var total = stops[stops.length - 1][0];
    return stops.map(function (st) {
      var f = { offset: st[0] / total };
      for (var key in st[1]) f[key] = st[1][key];
      if (st[2]) f.easing = st[2];
      return f;
    });
  }
  /* the flight curve of the sales chips, and the share of its duration at which it reaches progress p */
  var FLY = [0.45, 0, 0.3, 1], FLY_E = 'cubic-bezier(.45,0,.3,1)';
  function flyTime(p) {
    var lo = 0, hi = 1, s = 0.5, u;
    for (var n = 0; n < 32; n++) { s = (lo + hi) / 2; u = 1 - s; if (3 * u * u * s * FLY[1] + 3 * u * s * s * FLY[3] + s * s * s < p) lo = s; else hi = s; }
    u = 1 - s;
    return 3 * u * u * s * FLY[0] + 3 * u * s * s * FLY[2] + s * s * s;
  }
  var quartOut = function (u) { return 1 - Math.pow(1 - u, 4); };
  var backOut = function (u) { var c = 1.9; return 1 + (c + 1) * Math.pow(u - 1, 3) + c * Math.pow(u - 1, 2); };

  /* Faces follow assets/js/kicau/people.js */
  function face(f) {
    var s = function (c) { return 'style="fill:' + c + '"'; };
    var hijab = f.style === 'hijab', ey = hijab ? 19.4 : 17.6;
    var body = '<path ' + s(f.top) + ' d="M5 41c.6-7.8 6.8-12 15-12s14.4 4.2 15 12z"/>';
    var back = hijab ? '<path ' + s(f.hijab) + ' d="M10.4 22.4c0-7.6 4.2-13.2 9.6-13.2s9.6 5.6 9.6 13.2v6.2c-2.6 2.2-5.8 3.2-9.6 3.2s-7-1-9.6-3.2z"/>' : '';
    var head = hijab ? '<ellipse ' + s(f.skin) + ' cx="20" cy="20" rx="5.7" ry="6.6"/>'
      : '<rect ' + s(f.skin) + ' x="17.5" y="22" width="5" height="6" rx="1.4"/><circle ' + s(f.skin) + ' cx="20" cy="17.6" r="7.3"/>';
    var front = f.style === 'cap' ? '<path ' + s(f.cap) + ' d="M12.3 15.6c0-4.8 3.4-7.9 7.7-7.9s7.7 3.1 7.7 7.9z"/><path ' + s(f.cap) + ' d="M19 14.2h11.2c1 0 1 1.8 0 1.8H19z"/>' : '';
    var eyes = '<circle style="fill:#241C18" cx="17.4" cy="' + ey + '" r=".9"/><circle style="fill:#241C18" cx="22.6" cy="' + ey + '" r=".9"/><path d="M17.9 ' + (ey + 3.1) + 'q2.1 1.7 4.2 0" style="fill:none;stroke:#241C18;stroke-width:.95;stroke-linecap:round"/>';
    var cheeks = '<circle style="fill:#E86A5A;opacity:.18" cx="15.6" cy="' + (ey + 2.6) + '" r="1.5"/><circle style="fill:#E86A5A;opacity:.18" cx="24.4" cy="' + (ey + 2.6) + '" r="1.5"/>';
    return '<svg viewBox="0 0 40 40" aria-hidden="true"><rect ' + s(f.bg) + ' width="40" height="40"/>' + (hijab ? body + back : back + body) + head + front + eyes + cheeks + '</svg>';
  }

  /* Who asks what. Tika's request is the one the workbook answers; Bayu's and Citra's become sheet tabs. */
  var CREW = [
    { name: 'Tika', role: 'Sales Finance', col: '#B8452A', time: '09:12', msg: 'Total Mi Lekas sales this week?',
      f: { bg: '#FFE4D6', skin: '#C88B63', style: 'hijab', hijab: '#E2603A', top: '#B8452A' } },
    { name: 'Bayu', role: 'Warehouse', col: '#276749', time: '09:14', msg: 'Product names for the orders?',
      f: { bg: '#DCF2E3', skin: '#A86F48', style: 'cap', cap: '#2F855A', top: '#276749' } },
    { name: 'Citra', role: 'Bank Kenari', col: '#1E3A8A', time: '09:15', msg: 'Grade every loan 1 to 5?',
      f: { bg: '#DFECFF', skin: '#C99A72', style: 'hijab', hijab: '#2C5FB0', top: '#1E3A8A' } }
  ];

  /* Faint formulas behind the chat, at three depths (z: 0 far and soft, 1 near and brighter). */
  var C1 = '#7B9BFF', C2 = '#F87171', C3 = '#C084FC', C4 = '#34D399';
  var TOKENS = [
    { h: '=SUMIF(<b style="color:' + C1 + '">C2:C90</b>,"Mi Lekas",<b style="color:' + C4 + '">E2:E90</b>)', x: 5, y: 12, z: .9 },
    { h: '=IF(<b style="color:' + C1 + '">D2</b>&gt;90,5,IF(<b style="color:' + C1 + '">D2</b>&gt;60,4,3))', x: 58, y: 23, z: .35 },
    { h: '=VLOOKUP(<b style="color:' + C1 + '">B2</b>,<b style="color:' + C3 + '">Items!A:C</b>,2,FALSE)', x: 36, y: 85, z: .75 },
    { h: '<b style="color:' + C2 + '">F2:F48</b>', x: 7, y: 70, z: .2 },
    { h: 'fx', x: 82, y: 8, z: .55 },
    { h: '&Sigma;', x: 88, y: 61, z: 1 },
    { h: '=COUNTIFS(', x: -2, y: 41, z: .1 },
    { h: '<b style="color:' + C4 + '">$E$2</b>', x: 74, y: 76, z: .45 }
  ];

  function birdSvg(cls) {
    return '<svg class="' + cls + '" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="' + BIRD + '"/><circle cx="15.8" cy="8" r="1.15" fill="#fff"/></svg>';
  }
  /* the flying canary: the wing sits behind the body and only shows when it flaps */
  function canarySvg(c) {
    return '<svg class="k3-svg" viewBox="0 0 24 24" aria-hidden="true" focusable="false">' +
      '<g class="k3-wing"><path d="' + WING + '" fill="' + c.wing + '"/></g>' +
      '<path d="' + BIRD + '" fill="' + c.body + '"/>' +
      '<circle class="k3-eye" cx="15.8" cy="8" r="1.15" fill="' + c.eye + '"/></svg>';
  }
  function birdBox(cls, svg) {
    return '<div class="k3-bird ' + cls + '"><div class="k3-fx"><div class="k3-fy"><div class="k3-fr"><div class="k3-sq">' + svg + '</div></div></div></div></div>';
  }
  var SKIP = '<button type="button" class="kci-skip" aria-label="Skip intro">Skip<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" d="m6 6 6 6-6 6M13 6l6 6-6 6"/></svg></button>';

  window.kcIntro = function kcIntro(onDone) {
    var d = document, W = window;
    var old = d.getElementById('kc-intro');
    if (old) { if (old._kciEnd) old._kciEnd(true); else old.remove(); }

    var reduce = /(?:^|[?&])reduce=1(?:&|$)/.test(W.location.search.slice(1)) ||
      !!(W.matchMedia && W.matchMedia('(prefers-reduced-motion: reduce)').matches);
    var root = d.createElement('div');
    root.id = 'kc-intro';
    var anims = [], over = false, leaving = false, armed = false, timer = 0, remeasure = null;

    function A(el, kf, at, dur, ease, fill) {
      /* Give moving and fading pieces their own GPU layer up front, so the first frame of each move does not stall. */
      var k0 = kf && kf[0];
      if (k0 && (k0.transform !== undefined || k0.opacity !== undefined) && !el.style.willChange) el.style.willChange = k0.transform !== undefined ? 'transform, opacity' : 'opacity';
      var a = el.animate(kf, { delay: at || 0, duration: Math.max(1, dur), easing: ease || 'linear', fill: fill || 'both' });
      anims.push(a);
      return a;
    }
    function end(silent) {
      if (over) return;
      over = true;
      clearTimeout(timer);
      W.removeEventListener('keydown', onKey, true);
      if (remeasure) W.removeEventListener('resize', remeasure);
      root.removeEventListener('pointerdown', onTap);
      anims.forEach(function (a) { try { a.cancel(); } catch (e) { /* already gone */ } });
      root.remove();
      if (!silent && typeof onDone === 'function') onDone();
    }
    function skip() {
      if (over || leaving) return;
      leaving = true;
      /* the overlay keeps catching pointer events while it fades, so the click after a tap never lands on the app */
      var o = parseFloat(getComputedStyle(root).opacity);
      var a = root.animate([{ opacity: isNaN(o) ? 1 : o }, { opacity: 0 }], { duration: 250, easing: 'cubic-bezier(.33,0,.2,1)', fill: 'forwards' });
      anims.push(a);
      a.finished.then(function () { end(); }, function () {});
      clearTimeout(timer);
      timer = setTimeout(function () { end(); }, 450);
    }
    function onKey(e) {
      if (!armed) return;
      /* keep browser shortcuts (Cmd, Ctrl, Alt and the F keys) working; any other key only skips */
      if (!e.metaKey && !e.ctrlKey && !e.altKey && !/^F\d{1,2}$/.test(e.key || '')) e.preventDefault();
      skip();
    }
    function onTap() { if (armed) skip(); }
    root._kciEnd = end;

    function mount(html) {
      root.innerHTML = html;
      d.body.appendChild(root);
      W.addEventListener('keydown', onKey, true);
      root.addEventListener('pointerdown', onTap);
      /* arm on the next frames so the click or key that opened the intro does not also skip it */
      requestAnimationFrame(function () { requestAnimationFrame(function () { armed = true; }); });
      /* screen readers and switch access fire click without pointerdown */
      var btn = root.querySelector('.kci-skip');
      btn.addEventListener('click', skip);
      return btn;
    }
    var ctl = { skip: skip };

    /* ---------- reduced motion: the lockup fades in, holds, and the overlay fades away. Nothing moves. ---------- */
    if (reduce) {
      root.className = 'kci-rm';
      var skR = mount('<div class="kci-lay kci-rmbg"></div><div class="kci-rmlock" aria-hidden="true"><span class="kci-tile">' + birdSvg('') + '</span><span class="kci-rmw"><b>Kicau</b><small>Kenari Excel Lab</small></span></div><span class="kci-sr">Kicau, Kenari Excel Lab</span>' + SKIP);
      A(root.querySelector('.kci-rmlock'), [{ opacity: 0 }, { opacity: 1 }], 0, 400, 'ease');
      A(skR, [{ opacity: 0 }, { opacity: 1 }], 0, 400, 'ease');
      A(root, [{ opacity: 1 }, { opacity: 0 }], 1150, 400, 'ease', 'forwards').finished.then(function () { end(); }, function () {});
      timer = setTimeout(function () { end(); }, 2200);
      root.dataset.state = 'play';
      root.dataset.t0 = String(performance.now());
      return ctl;
    }
    if (!root.animate) { setTimeout(function () { if (typeof onDone === 'function') onDone(); }, 0); return ctl; }

    /* ---------- geometry of the workbook: one unit f, rounded so gridlines stay crisp ---------- */
    var vw = W.innerWidth, vh = W.innerHeight;
    var f = Math.floor(Math.max(11, Math.min((vw - 40) / 26.3, vh * 0.026, 20)));
    var RH = Math.round(f * 1.5);
    var KEYS = ['#', 'A', 'B', 'C', 'D', 'E', 'F'];
    var CW = { '#': 2.1, A: 3.7, B: 6.6, C: 5, D: 2.9, E: 4.4, F: 1.6 }, CX = {}, x = 0;
    KEYS.forEach(function (k) { CW[k] = Math.round(CW[k] * f); CX[k] = x; x += CW[k]; });
    var WBW = x, NR = 10, SH = (NR + 1) * RH + Math.round(f * 0.7), PAD = Math.round(f * 0.32);
    function R_(c, r, n) { return { l: CX[c], t: r * RH, w: CW[c], h: (n || 1) * RH }; }
    function P(o, g) { g = g || 0; return 'left:' + (o.l - g) + 'px;top:' + (o.t - g) + 'px;width:' + (o.w + 2 * g) + 'px;height:' + (o.h + 2 * g) + 'px'; }

    /* Tika's question and its data (Mi Lekas: 1,250 + 1,430 + 1,120 = 3,800) */
    var DATA = [
      ['3-Nov', 'Mi Lekas', 'Jakarta', '120', '1,250'],
      ['3-Nov', 'Kopi Sekejap', 'Bandung', '80', '980'],
      ['4-Nov', 'Mi Lekas', 'Surabaya', '140', '1,430'],
      ['4-Nov', 'Sambal Nyala', 'Jakarta', '60', '760'],
      ['5-Nov', 'Mi Lekas', 'Bandung', '110', '1,120'],
      ['5-Nov', 'Sabun Rimba', 'Medan', '40', '640']
    ];
    var HITS = [4, 6, 8], PART = ['1,250', '2,680', '3,800'];
    var cells = [{ c: 'A', r: 1, t: 'Product', b: 1 }, { c: 'B', r: 1, t: 'Mi Lekas' }, { c: 'D', r: 1, t: 'Total', b: 1, n: 1 }];
    ['Date', 'Product', 'Area', 'Qty', 'Sales'].forEach(function (t, i) { cells.push({ c: 'ABCDE'[i], r: 3, t: t, b: 1, n: i === 0 || i >= 3 }); });
    DATA.forEach(function (row, ri) { row.forEach(function (t, ci) { cells.push({ c: 'ABCDE'[ci], r: 4 + ri, t: t, n: ci === 0 || ci >= 3 }); }); });
    var FX = '=SUMIFS(E4:E9,B4:B9,B1)', N = FX.length;
    var FXC = '=SUMIFS(<i style="color:#2E6BD6">E4:E9</i>,<i style="color:#C2453A">B4:B9</i>,<i style="color:#7B4FB5">B1</i>)';

    /* ---------- markup: Act 1, the chat ---------- */
    var kTok = Math.max(.72, Math.min(1, vw / 1100));
    var tokens = TOKENS.map(function (t) {
      var size = (11 + t.z * 8) * kTok, blur = (1 - t.z) * 1.3;
      return '<span style="left:' + t.x + '%;top:' + t.y + '%;font-size:' + size.toFixed(1) + 'px' + (blur > .25 ? ';filter:blur(' + blur.toFixed(2) + 'px)' : '') + '">' + t.h + '</span>';
    }).join('');
    var who = function (p) { return '<div class="k1-who" style="color:' + p.col + '">' + p.name + '<span>' + p.role + '</span><time>' + p.time + '</time></div>'; };
    var rowsHtml = CREW.map(function (p) {
      return '<div class="k1-row"><div class="k1-av"><div class="k1-face">' + face(p.f) + '</div><i class="k1-on"></i></div>' +
        '<div class="k1-bw"><div class="k1-bub"><div class="k1-dots"><i></i><i></i><i></i></div><div class="k1-txt">' +
        who(p) + '<div class="k1-msg">' + p.msg + '</div></div></div></div></div>';
    }).join('');
    var flyHtml = CREW.map(function (p) {
      return '<i class="k1-sh o"></i><i class="k1-sh w"></i><div class="k1-cl k1-cw">' + who(p) + '</div><div class="k1-cl k1-cm"><div class="k1-msg">' + p.msg + '</div></div>';
    }).join('');

    /* ---------- markup: Act 2, the workbook ---------- */
    var S = [];
    S.push('<div class="k2-hrow" style="height:' + RH + 'px"></div>');
    S.push('<div class="k2-hcol" style="width:' + CW['#'] + 'px"></div>');
    S.push('<div class="k2-fill in" style="' + P(R_('B', 1)) + '"></div>');
    S.push('<div class="k2-fill th" style="left:' + CX.A + 'px;top:' + 3 * RH + 'px;width:' + (CX.F - CX.A) + 'px;height:' + RH + 'px"></div>');
    S.push('<div class="k2-hon c ha" style="left:' + CX.A + 'px;top:0;width:' + CW.A + 'px;height:' + RH + 'px"></div>');
    S.push('<div class="k2-hon c he" style="left:' + CX.E + 'px;top:0;width:' + CW.E + 'px;height:' + RH + 'px"></div>');
    S.push('<div class="k2-hon r" style="left:0;top:' + RH + 'px;width:' + CW['#'] + 'px;height:' + RH + 'px"></div>');
    HITS.forEach(function (r) {
      S.push('<div class="k2-hit b" style="' + P(R_('B', r)) + '"></div>');
      S.push('<div class="k2-hit e" style="' + P(R_('E', r)) + '"></div>');
    });
    for (var r = 2; r <= NR + 1; r++) S.push('<div class="k2-hl" style="top:' + (r * RH - 1) + 'px;background:linear-gradient(90deg,#D4D4D4 0 ' + CW['#'] + 'px,#E1E1E1 0)"></div>');
    ['B', 'C', 'D', 'E', 'F'].forEach(function (k) { S.push('<div class="k2-vl" style="left:' + (CX[k] - 1) + 'px;background:linear-gradient(#D4D4D4 0 ' + RH + 'px,#E1E1E1 0)"></div>'); });
    KEYS.slice(1).forEach(function (k) { S.push('<div class="k2-ch" style="left:' + CX[k] + 'px;top:0;width:' + CW[k] + 'px;height:' + RH + 'px">' + k + '</div>'); });
    for (r = 1; r <= NR + 1; r++) S.push('<div class="k2-rh" style="left:0;top:' + r * RH + 'px;width:' + CW['#'] + 'px;height:' + RH + 'px">' + r + '</div>');
    function rg(cls, c, r0, n, col, tint) { return '<div class="k2-rg ' + cls + '" style="' + P(R_(c, r0, n), 1) + ';--c:' + col + ';--t:' + tint + '"><i></i><i></i><i></i><i></i></div>'; }
    S.push('<div class="k2-rngs">' +
      rg('one r1', 'E', 4, 1, '#2E6BD6', 'rgba(46,107,214,.09)') + rg('full r1', 'E', 4, 6, '#2E6BD6', 'rgba(46,107,214,.09)') +
      rg('one r2', 'B', 4, 1, '#C2453A', 'rgba(194,69,58,.09)') + rg('full r2', 'B', 4, 6, '#C2453A', 'rgba(194,69,58,.09)') +
      rg('one r3', 'B', 1, 1, '#7B4FB5', 'rgba(123,79,181,.11)') + '</div>');
    cells.forEach(function (c) {
      S.push('<div class="k2-c' + (c.n ? ' n' : '') + (c.b ? ' b' : '') + '" data-r="' + c.r + '" data-c="' + c.c + '" style="' + P(R_(c.c, c.r)) + ';padding:0 ' + PAD + 'px"><span>' + c.t + '</span></div>');
    });
    PART.forEach(function (v, i) { S.push('<div class="k2-val' + (i === 2 ? ' fin' : '') + '" style="' + P(R_('E', 1)) + ';padding-right:' + PAD + 'px"><span>' + v + '</span></div>'); });
    S.push('<div class="k2-selw" style="' + P(R_('E', 1), 1) + '"><div class="k2-sel"></div></div>');
    S.push('<div class="k2-ring" style="' + P(R_('E', 1), 1) + '"></div>');
    /* the flying sales chips sit above the cell cursor, so they pass over its fill handle on the way into E1 */
    HITS.forEach(function (r) { S.push('<div class="k2-chip" style="' + P(R_('E', r)) + ';padding-right:' + Math.max(2, PAD - 3) + 'px"><b>' + DATA[r - 4][4] + '</b></div>'); });

    var ic = function (p) { return '<svg viewBox="0 0 12 12"><path fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" d="' + p + '"/></svg>'; };
    var X_ = 'M2.5 2.5l7 7M9.5 2.5l-7 7', V_ = 'M2 6.4l2.6 2.6L10 3.2';
    var titlebar = '<div class="k2-tb"><svg class="k2-xl" viewBox="0 0 16 16"><rect width="16" height="16" rx="3.2" fill="#fff"/><path d="M5.2 4.6l5.6 6.8M10.8 4.6l-5.6 6.8" stroke="#107C41" stroke-width="2" stroke-linecap="round"/></svg>' +
      '<span class="k2-fn">Sales_November.xlsx' + (WBW >= 400 ? ' <small>&middot; Rasa Kenari</small>' : '') + '</span>' +
      '<span class="k2-wc">' + ic('M2 6h8') + ic('M2.5 2.5h7v7h-7z') + ic(X_) + '</span></div>';
    var fbar = '<div class="k2-fbar"><div class="k2-nbox"><span class="k2-nb a">A1</span><span class="k2-nb e">E1</span><i class="k2-dd"></i></div>' +
      '<div class="k2-fxi"><span class="k2-ic x">' + ic(X_) + '<span class="on">' + ic(X_) + '</span></span><span class="k2-ic v">' + ic(V_) + '<span class="on">' + ic(V_) + '</span></span><em>fx</em></div>' +
      '<div class="k2-fx"><span class="k2-fw"><span class="k2-ftw">' + FXC + '<span class="plain">' + FX + '</span></span><i class="k2-cw"><i class="k2-caret"></i></i></span></div></div>';
    var tabs = '<div class="k2-tabs"><span class="arr">&lsaquo; &rsaquo;</span><b>Sales</b><span class="k2-tab o"><i>Orders</i></span><span class="k2-tab l"><i>Loans</i></span><span>+</span><span class="st"><i class="rdy">Ready</i><i class="ent">Enter</i></span></div>';

    /* ---------- markup: Act 3, the canary and the lockup ---------- */
    var word = 'Kicau'.split('').map(function (ch) { return '<span class="k3-ch">' + ch + '</span>'; }).join('');
    var sub = 'Kenari Excel Lab'.split('').map(function (ch) { return '<span class="k3-sc">' + ch + '</span>'; }).join('');
    var chirp = '<svg class="k3-chirp" viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="' + INK + '" stroke-width="1.9" stroke-linecap="round">' +
      '<path class="k3-dash k3-arc" pathLength="1" d="M4.6 16.1A6 6 0 0 0 1.6 14.4"/>' +
      '<path class="k3-dash k3-arc" pathLength="1" d="M9.1 12.5A11 11 0 0 0 3.3 9.6"/>' +
      '<path class="k3-dash k3-arc" pathLength="1" d="M13.6 8.9A16 16 0 0 0 5 4.8"/></svg>';
    /* the lockup (with the ink canary) exists three times in the same place: one whole copy on top, plus one in each
       curtain half, so at the end the logo itself splits on the diagonal and leaves with the curtain */
    var top = function (cls) {
      return '<div class="kci-lay k3-top ' + cls + '"><div class="kci-lay k3-stage"><div class="k3-lockup">' +
          '<div class="k3-row"><span class="k3-slot">' + chirp + '</span><span class="k3-word">' + word + '</span></div>' +
          '<div class="k3-sub"><i class="k3-rule k3-rl"></i><span class="k3-subtxt">' + sub + '</span><i class="k3-rule k3-rr"></i></div>' +
        '</div></div>' +
        birdBox('k3-bird-b', canarySvg({ body: INK, wing: '#4B4334', eye: '#FFFFFF' })) +
        '</div>';
    };
    var D = Math.round(Math.min(Math.min(vw, vh) * .5, 290));
    root.style.setProperty('--D', D + 'px');

    var sk = mount(
      '<span class="kci-sr">Kicau, Kenari Excel Lab</span>' +
      '<svg class="kci-probe" aria-hidden="true" focusable="false"><path d="' + BIRD + '"/></svg>' +
      '<div class="kci-lay k1-dark" aria-hidden="true"><i class="k1-glow"></i><i class="k1-grid"></i><div class="k1-tokens">' + tokens + '</div></div>' +
      '<i class="k2-iris" aria-hidden="true"></i>' +
      '<div class="kci-lay k2-bg" aria-hidden="true"><i class="k2-glow g1"></i><i class="k2-glow g2"></i><i class="k2-dots"></i></div>' +
      '<div class="kci-lay k2-stage" aria-hidden="true" style="font-size:' + f + 'px"><div class="k2-cam"><div class="k2-floor"></div>' +
        '<div class="k2-wb" style="width:' + WBW + 'px">' + titlebar + fbar + '<div class="k2-sheet" style="height:' + SH + 'px">' + S.join('') + '</div>' + tabs + '</div></div></div>' +
      '<div class="kci-lay k1-chat" aria-hidden="true"><div class="k1-feed"><div class="k1-head"><span class="k1-hl">New requests</span><span class="k1-badge"><span class="k1-strip"><i>1</i><i>2</i><i>3</i></span></span></div>' + rowsHtml + '</div></div>' +
      '<div class="kci-lay k1-fly" aria-hidden="true">' + flyHtml + '</div>' +
      '<div class="kci-lay k3-morph" aria-hidden="true"><div class="k3-gold"><span class="k3-num"><span>3,800</span></span></div></div>' +
      '<div class="kci-lay k3-wipe k3-s1" aria-hidden="true"></div>' +
      '<div class="kci-lay k3-wipe k3-s2" aria-hidden="true"></div>' +
      '<div class="kci-lay k3-night" aria-hidden="true">' + birdBox('k3-bird-a', canarySvg({ body: '#F2B705', wing: '#C28F00', eye: INK })) + '</div>' +
      '<div class="kci-lay k3-wipe k3-yel" aria-hidden="true">' +
        '<div class="kci-lay k3-lin k3-cl k3-lin-l"></div><div class="kci-lay k3-lin k3-cr k3-lin-r"></div>' +
        '<div class="kci-lay k3-half k3-cl"><div class="k3-grid"></div>' + top('k3-split') + '</div>' +
        '<div class="kci-lay k3-half k3-cr"><div class="k3-grid"></div>' + top('k3-split') + '</div>' +
        top('k3-solo') +
      '</div>' + SKIP);

    /* Wait for the typefaces (at most 500 ms on a first visit; cached afterwards), so text never changes font mid-move. */
    var fontsP = Promise.resolve();
    try {
      if (d.fonts && d.fonts.load) {
        fontsP = Promise.all(['500 16px "Plus Jakarta Sans"', '700 16px "Plus Jakarta Sans"', '800 64px "Plus Jakarta Sans"', '500 16px "JetBrains Mono"', '400 16px Carlito', '700 16px Carlito']
          .map(function (s) { return d.fonts.load(s).catch(function () {}); }));
      }
    } catch (e) { /* no font loading API */ }
    /* Then wait for the page to go quiet (the app finishing its start-up), so the first moves never stutter.
       Until then the still first frame, the dark stage, is on screen. */
    var calm = function (fn) {
      var go = function () { requestAnimationFrame(function () { requestAnimationFrame(fn); }); };
      if (W.requestIdleCallback) W.requestIdleCallback(go, { timeout: 250 }); else setTimeout(go, 80);
    };
    Promise.race([fontsP, new Promise(function (res) { setTimeout(res, 500); })]).then(function () {
      calm(function () {
        if (over || leaving) return;
        try { play(); } catch (err) { end(); if (W.console) console.error('kcIntro', err); }
      });
    });
    return ctl;

    function play() {
      try { performance.mark('kc-intro-play'); } catch (e) { /* no performance marks */ }
      var $ = function (s) { return root.querySelector(s); };
      var $$ = function (s) { return Array.prototype.slice.call(root.querySelectorAll(s)); };
      var BR = function (el) { return el.getBoundingClientRect(); };
      var rect = function (b, dy) { return { l: b.left, t: b.top + (dy || 0), w: b.width, h: b.height }; };
      var box = BR(root), Wd = box.width, H = box.height;

      /* ================= measure once, before anything moves ================= */
      var feed = $('.k1-feed'), head = $('.k1-head'), hl = $('.k1-hl'), hb = $('.k1-badge'), strip = $('.k1-strip');
      var rows = $$('.k1-row');
      var bubs = rows.map(function (rw) { return rw.querySelector('.k1-bub'); });
      /* freeze each bubble's box, so a late font swap cannot change the shapes the flights start from */
      bubs.forEach(function (b) { b.style.width = b.offsetWidth + 'px'; });
      var bsz = bubs.map(function (b) { return { w: b.offsetWidth, h: b.offsetHeight }; });
      var hr = BR(head), rr = rows.map(BR);
      /* feed offsets: keep the block of arrived messages centred, like a chat pushing up */
      var offs = rr.map(function (b) { return H / 2 - (hr.top + b.bottom + 13) / 2; });
      var OY = offs[2];
      var bR = bubs.map(function (b) { return rect(BR(b), OY); });
      var whoR = rows.map(function (rw) { return rect(BR(rw.querySelector('.k1-who')), OY); });
      var msgR = rows.map(function (rw) { return rect(BR(rw.querySelector('.k1-msg')), OY); });
      var msgFS = parseFloat(getComputedStyle(rows[0].querySelector('.k1-msg')).fontSize) || 16;

      var stage = $('.k2-stage'), wb = $('.k2-wb');
      var fxR = rect(BR($('.k2-fx'))), fwR = BR($('.k2-fw'));
      var tabEls = [$('.k2-tab.o'), $('.k2-tab.l')];
      var tabR = tabEls.map(function (t) { return rect(BR(t)); });
      var shR = BR($('.k2-sheet'));
      var E1 = { l: shR.left + CX.E, t: shR.top + RH, w: CW.E, h: RH };
      var ex = E1.l + E1.w / 2, ey = E1.t + E1.h / 2;
      var cam = $('.k2-cam'), camR = BR(cam);

      /* the bird outline, sampled into points for the cell-to-canary morph */
      var probe = $('.kci-probe path'), NPT = 72, LEN = probe.getTotalLength(), bp = [];
      for (var q = 0; q < NPT; q++) { var pt = probe.getPointAtLength(LEN * q / NPT); bp.push([pt.x, pt.y]); }

      /* ================= timeline (ms from the start of play) ================= */
      var ROWS = [40, 160, 280], TYPEIN = 190, EXPAND = 360;
      var EXP = ROWS.map(function (s) { return s + TYPEIN; });
      var TICK = EXP.map(function (e) { return e + 100; });
      var S1 = 850;                      // the requests leave the chat for the workbook; the stage brightens
      var W0 = S1 + 20;                  // the window tilts in under them, on the spreading light
      var FLY0 = [S1 + 10, S1 + 50, S1 + 85], FLYD = [400, 360, 360];
      var GRID = W0 + 40, VALS = W0 + 140, SEL = VALS + 240;
      var TYPE = S1 + 410;
      var tt = [], acc = 20;
      var JIT = [0, 2, -2, 3, -2, 1, -1, 2, 0, -2, 3, -2, 1, 1, -1, 2, -1, 0, 1, -2, 1, -1, 2];
      for (var i = 0; i < N; i++) {
        if (i) acc += 14 + JIT[i] + (i === 8 ? 40 : (i === 14 || i === 20) ? 22 : i === 22 ? 16 : 0);
        tt.push(acc);
      }
      var TD = tt[N - 1];
      var ENTER = TYPE + TD + 30;
      /* each sales chip flies up the E column and dissolves into E1 at the moment it reaches the cell's lower edge;
         the running total ticks at that instant. PE(k) is the flight progress where chip k meets E1. */
      var CHIPS = ENTER + 20, CH_GAP = 65, CH_DUR = 280;
      var PE = function (k) { return 0.6 + ((1 - 1 / (HITS[k] - 1)) - 0.6) / 0.36 * 0.28; };
      var ARR = function (k) { return CHIPS + k * CH_GAP + Math.round(flyTime(PE(k) + 0.06) * CH_DUR); };
      var RESULT = ARR(2);
      var GOLD = RESULT + 30, GD = 170;   // the answer cell fills with gold
      var MORPH = RESULT + 230, MD = 220; // and reshapes into the canary
      var T0 = MORPH + MD;               // the canary takes off
      var LAND = T0 + 760;               // and perches in the lockup (one continuous flight, turns included)
      var CURT = LAND + 640;             // the curtain tears open
      var END = CURT + 400;

      stage.style.visibility = 'visible';
      root.dataset.state = 'play';
      root.dataset.t0 = String(performance.now());
      root.dataset.end = String(END);

      /* ================= Act 1: requests arrive ================= */
      A($('.k1-glow'), [{ opacity: 0, transform: 'scale(.6)' }, { opacity: 1, transform: 'scale(1)' }], 0, 900, SOFT);
      A($('.k1-grid'), [{ opacity: 0 }, { opacity: 1 }], 60, 600, 'ease-out');
      A($('.k1-grid'), [{ transform: 'translateY(18px)' }, { transform: 'translateY(-22px)' }], 0, S1 + 450, 'linear');
      $$('.k1-tokens span').forEach(function (el, k) {
        var t = TOKENS[k], op = .06 + t.z * .1, drift = 14 + t.z * 36;
        A(el, [{ opacity: 0 }, { opacity: op }], 100 + k * 50, 600, 'ease-out');
        A(el, [{ transform: 'translateY(' + (drift * .5).toFixed(1) + 'px)' }, { transform: 'translateY(' + (-drift).toFixed(1) + 'px)' }], 0, S1 + 450, 'cubic-bezier(.3,.1,.5,1)');
      });
      A(sk, [{ opacity: 0 }, { opacity: 1 }], 200, 320, 'ease-out');

      A(hl, [{ opacity: 0, transform: 'translateY(8px)' }, { opacity: 1, transform: 'none' }], 30, 420, EXPO);
      A(head, [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(-8px)' }], S1 - 90, 190, OUTIN, 'forwards');

      /* overlapping eased steps added together, sampled into keyframes (for numbers that move again before they settle) */
      function glide(el, v0, steps, ease) {
        var endT = 0; steps.forEach(function (s) { endT = Math.max(endT, s.at + s.dur); });
        var kf = [], n = 48;
        for (var j = 0; j <= n; j++) {
          var t = endT * j / n, v = v0;
          steps.forEach(function (s) { var u = Math.min(1, Math.max(0, (t - s.at) / s.dur)); v += (s.v - s.from) * (u >= 1 ? 1 : ease(u)); });
          kf.push({ offset: j / n, transform: 'translateY(' + v.toFixed(2) + 'px)' });
        }
        return A(el, kf, 0, endT, 'linear');
      }
      glide(feed, offs[0], [
        { at: ROWS[1], dur: 540, from: offs[0], v: offs[1] },
        { at: ROWS[2], dur: 540, from: offs[1], v: offs[2] }
      ], quartOut);

      /* the red badge counts 1, 2, 3 as each request opens */
      A(hb, springKF(480, .45, function (p) { return { transform: 'scale(' + p.toFixed(4) + ')' }; }), TICK[0], 480, 'linear');
      glide(strip, 0, [
        { at: TICK[1], dur: 320, from: 0, v: -20 },
        { at: TICK[2], dur: 320, from: -20, v: -40 }
      ], backOut);
      var BUMP = [{ transform: 'scale(1)' }, { transform: 'scale(1.18)' }, { transform: 'scale(1)' }];
      A(hb, BUMP, TICK[1], 280, 'ease-out', 'none');
      A(hb, BUMP, TICK[2], 280, 'ease-out', 'none');

      var pill = function (b) { return 'inset(' + (b.h - 36) + 'px ' + (b.w - 62) + 'px 0px 0px round 18px 18px 18px 6px)'; };
      var full = 'inset(0px 0px 0px 0px round 18px 18px 18px 6px)';
      rows.forEach(function (row, k) {
        var s = ROWS[k], exp = EXP[k];
        var av = row.querySelector('.k1-av'), on = row.querySelector('.k1-on'), bw = row.querySelector('.k1-bw');
        var dots = row.querySelector('.k1-dots'), txt = row.querySelector('.k1-txt');
        A(av, springKF(560, .5, function (p) { return { transform: 'translateY(' + (10 * (1 - p)).toFixed(2) + 'px) scale(' + (.25 + .75 * p).toFixed(4) + ')', opacity: Math.min(1, p * 3) }; }), s, 560, 'linear');
        A(on, springKF(420, .45, function (p) { return { transform: 'scale(' + p.toFixed(4) + ')' }; }), s + 200, 420, 'linear');
        A(bw, [{ opacity: 0, transform: 'translateY(18px) scale(.94)' }, { opacity: 1, transform: 'none' }], s + 50, 440, EXPO);
        /* the typing pill opens into the full message */
        A(bubs[k], [{ clipPath: pill(bsz[k]) }, { clipPath: full }], exp, EXPAND, EXPO);
        Array.prototype.forEach.call(dots.querySelectorAll('i'), function (dot, j) {
          anims.push(dot.animate([{ transform: 'translateY(0)', opacity: .5 }, { transform: 'translateY(-4px)', opacity: 1, offset: .28 }, { transform: 'translateY(0)', opacity: .5, offset: .56 }, { transform: 'translateY(0)', opacity: .5 }],
            { duration: 600, delay: s + 50 + j * 100, easing: 'ease-in-out', iterations: Infinity, fill: 'both' }));
        });
        A(dots, [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'scale(.6)' }], exp, 90, 'ease-in');
        A(txt, [{ opacity: 0, transform: 'translateY(6px)' }, { opacity: 1, transform: 'none' }], exp + 40, 300, EXPO);

        /* seam 1: the bubble hands over to its flight copy, and the face bows out */
        A(bw, [{ opacity: 1 }, { opacity: 0 }], S1, 1, 'linear', 'forwards');
        A(av, [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(4px) scale(.3)' }], S1 - 40 + k * 20, 150, IN, 'forwards');
      });

      /* ================= Seam 1: the stage brightens; the requests fly into the workbook ================= */
      /* sized so the disc clears the corners at about 80% of its growth: the light fills the screen without creeping into the corners */
      var iris = $('.k2-iris'), IR = Math.ceil(Math.hypot(Wd / 2, H / 2) * 1.25 / 0.95) + 4;
      iris.style.cssText = 'width:' + 2 * IR + 'px;height:' + 2 * IR + 'px;margin:' + -IR + 'px 0 0 ' + -IR + 'px';
      A(iris, [{ transform: 'scale(0)' }, { transform: 'scale(1)' }], S1 - 70, 460, 'cubic-bezier(.3,.55,.25,1)');
      A($('.k2-bg'), [{ opacity: 0 }, { opacity: 1 }], S1 + 230, 260, 'ease-out');
      A($('.k1-dark'), [{ opacity: 1 }, { opacity: 0 }], S1 + 440, 1, 'linear', 'forwards');
      A(iris, [{ opacity: 1 }, { opacity: 0 }], S1 + 520, 1, 'linear', 'forwards');
      A($('.k2-glow.g1'), [{ transform: 'translate(20vw,-18vh) scale(.92)' }, { transform: 'translate(13vw,-11vh) scale(1.06)' }], S1 - 40, 2000, 'ease-in-out');
      A($('.k2-glow.g2'), [{ transform: 'translate(-24vw,22vh) scale(1)' }, { transform: 'translate(-18vw,16vh) scale(1.1)' }], S1 - 40, 2000, 'ease-in-out');
      A($('.k2-dots'), [{ opacity: 0, transform: 'scale(1.06)' }, { opacity: 1, transform: 'scale(1)' }], S1, 900, EXPO);

      var ins = function (b, g, rad) {
        return 'inset(' + (b.t - g).toFixed(1) + 'px ' + (Wd - b.l - b.w - g).toFixed(1) + 'px ' + (H - b.t - b.h - g).toFixed(1) + 'px ' + (b.l - g).toFixed(1) + 'px round ' + rad + ')';
      };
      var RBUB = function (g) { return (18 + g) + 'px ' + (18 + g) + 'px ' + (18 + g) + 'px ' + (6 + g) + 'px'; };
      var RBOX = function (g) { return (1 + g) + 'px ' + (1 + g) + 'px ' + (1 + g) + 'px ' + (1 + g) + 'px'; };
      var shO = $$('.k1-sh.o'), shW = $$('.k1-sh.w'), cW = $$('.k1-cw'), cM = $$('.k1-cm');
      /* The bubble shapes fly by clip-path (main thread) and the words by transform (compositor). A constant, non-clipping
         clip-path in the words' flight keeps that animation on the main thread too, so words and bubble move frame for frame. */
      var SYNC = 'inset(-200px)';
      var TGT = [fxR, tabR[0], tabR[1]];
      CREW.forEach(function (p, k) {
        var from = bR[k], to = TGT[k], at = FLY0[k], dur = FLYD[k], land = at + dur;
        [shO[k], shW[k]].forEach(function (el, j) {
          var g = j === 0 ? 1 : 0;
          A(el, [{ clipPath: ins(from, g, RBUB(g)) }, { clipPath: ins(to, g, RBOX(g)) }], at, dur, DOCK);
          A(el, [{ opacity: 0 }, { opacity: 1 }], S1, 1, 'linear');
          A(el, [{ opacity: 1 }, { opacity: 0 }], land - 40, 140, 'ease-out', 'forwards');
        });
        cW[k].style.cssText = 'left:' + whoR[k].l.toFixed(1) + 'px;top:' + whoR[k].t.toFixed(1) + 'px;width:' + whoR[k].w.toFixed(1) + 'px';
        cM[k].style.cssText = 'left:' + msgR[k].l.toFixed(1) + 'px;top:' + msgR[k].t.toFixed(1) + 'px;width:' + msgR[k].w.toFixed(1) + 'px';
        if (!k) {
          A(cW[k], [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(-3px)' }], S1, 100, 'ease-in', 'forwards');
        } else {
          /* Bayu's and Citra's words ride inside their bubbles, shrinking with them, and fade just before the bubble becomes
             its sheet tab (so no empty white box ever flies over the workbook) */
          /* the words live inside the white shape, so its clip keeps them in the bubble even if the main thread (clip-path)
             falls a frame behind the compositor (transform) on a slow device */
          shW[k].appendChild(cW[k]); shW[k].appendChild(cM[k]);
          var sc2 = Math.min(to.w / from.w, to.h / from.h), fcx = from.l + from.w / 2, fcy = from.t + from.h / 2, tcx = to.l + to.w / 2, tcy = to.t + to.h / 2;
          var span = land - S1, o1 = (at - S1 + dur * 0.3) / span, o2 = (at - S1 + dur * 0.62) / span;
          [[cW[k], whoR[k], 0.82], [cM[k], msgR[k], 1]].forEach(function (pr) {
            var el = pr[0], r0 = pr[1];
            var tx = tcx + (r0.l - fcx) * sc2 - r0.l, ty = tcy + (r0.t - fcy) * sc2 - r0.t;
            A(el, [{ transform: 'translate(0px,0px) scale(1)', clipPath: SYNC }, { transform: 'translate(' + tx.toFixed(1) + 'px,' + ty.toFixed(1) + 'px) scale(' + sc2.toFixed(4) + ')', clipPath: SYNC }], at, dur, DOCK);
            A(el, [{ opacity: 1, offset: 0, clipPath: SYNC }, { opacity: 1, offset: o1 * pr[2], clipPath: SYNC }, { opacity: 0, offset: o2 * pr[2], clipPath: SYNC }, { opacity: 0, offset: 1, clipPath: SYNC }], S1, span, 'linear', 'forwards');
          });
          /* Bayu's and Citra's requests land as the Orders and Loans sheets */
          A(tabEls[k - 1].firstChild, [{ opacity: 0, transform: 'scale(.6)' }, { opacity: 1, transform: 'none' }], land - 70, 280, BACK);
        }
      });
      /* Tika's words land in the formula bar, the line the formula will be typed on */
      var mr = msgR[0], ms = (f * 0.86) / msgFS;
      var mT = 'translate(' + (fwR.left - mr.l).toFixed(1) + 'px,' + (fxR.t + fxR.h / 2 - mr.h * ms / 2 - mr.t).toFixed(1) + 'px) scale(' + ms.toFixed(4) + ')';
      A(cM[0], [{ opacity: 1 }, { opacity: 1 }], S1, 1, 'linear', 'forwards');
      A(cM[0], [{ transform: 'translate(0px,0px) scale(1)', clipPath: SYNC }, { transform: mT, clipPath: SYNC }], FLY0[0], FLYD[0], DOCK);
      A(cM[0], [{ opacity: 1, transform: mT }, { opacity: 0, transform: mT + ' translateY(-55%)' }], TYPE - 40, 130, 'ease-in', 'forwards');

      /* ================= Act 2: the workbook answers ================= */
      /* the window arrives in depth: a short 3D tilt on the window, a longer 2D settle on the stage */
      A(stage, [{ transform: 'translateY(' + 3 * f + 'px) scale(.9)' }, { transform: 'none' }], W0, 640, EXPO);
      A(wb, [{ transform: 'perspective(1400px) rotateX(22deg)' }, { transform: 'none' }], W0, 460, EXPO);
      /* a short fade, so the window is never a see-through box under the flying requests */
      A(wb, [{ opacity: 0 }, { opacity: 1 }], W0 + 30, 110, 'ease-out');
      A($('.k2-floor'), [{ opacity: 0, transform: 'scaleX(.55)' }, { opacity: 1, transform: 'none' }], W0 + 40, 640, EXPO);
      $$('.k2-tb > *').forEach(function (el, k) { A(el, [{ opacity: 0, transform: 'translateX(-.4em)' }, { opacity: 1, transform: 'none' }], W0 + 140 + k * 45, 340, EXPO); });

      /* the grid draws itself */
      A($('.k2-hrow'), [{ transform: 'scaleX(0)' }, { transform: 'none' }], GRID, 440, EXPO);
      A($('.k2-hcol'), [{ transform: 'scaleY(0)' }, { transform: 'none' }], GRID + 30, 480, EXPO);
      $$('.k2-hl').forEach(function (el, k) { A(el, [{ transform: 'scaleX(0)' }, { transform: 'none' }], GRID + 50 + k * 22, 420, EXPO); });
      $$('.k2-vl').forEach(function (el, k) { A(el, [{ transform: 'scaleY(0)' }, { transform: 'none' }], GRID + 80 + k * 36, 480, EXPO); });
      $$('.k2-ch').forEach(function (el, k) { A(el, [{ opacity: 0, transform: 'translateY(-45%)' }, { opacity: 1, transform: 'none' }], GRID + 80 + k * 26, 320, EXPO); });
      $$('.k2-rh').forEach(function (el, k) { A(el, [{ opacity: 0, transform: 'translateX(-35%)' }, { opacity: 1, transform: 'none' }], GRID + 100 + k * 20, 320, EXPO); });
      A($('.k2-tabs'), [{ opacity: 0 }, { opacity: 1 }], GRID + 160, 300, 'ease-out');

      /* the numbers drop into their cells, row by row */
      A($('.k2-fill.in'), [{ opacity: 0 }, { opacity: 1 }], VALS + 20, 300, 'ease-out');
      A($('.k2-fill.th'), [{ transform: 'scaleX(0)' }, { transform: 'none' }], VALS + 50, 400, EXPO);
      $$('.k2-c').forEach(function (el) {
        var rr2 = +el.getAttribute('data-r'), col = 'ABCDE'.indexOf(el.getAttribute('data-c'));
        A(el, [{ opacity: 0, transform: 'translateY(-70%)' }, { opacity: 1, transform: 'none' }], VALS + (rr2 - 1) * 24 + col * 14, 320, BACK);
      });
      /* the cell cursor starts at A1, then glides to E1 */
      A($('.k2-sel'), [{ opacity: 0, transform: 'scale(1.25)' }, { opacity: 1, transform: 'none' }], VALS + 120, 240, BACK);
      A($('.k2-hon.ha'), [{ opacity: 0 }, { opacity: 1 }], VALS + 120, 160, 'ease-out');
      A($('.k2-hon.r'), [{ opacity: 0 }, { opacity: 1 }], VALS + 120, 160, 'ease-out');
      A($('.k2-selw'), [{ transform: 'translateX(' + (CX.A - CX.E) + 'px) scaleX(' + (CW.A / CW.E).toFixed(3) + ')' }, { transform: 'none' }], SEL, 360, EXPO);
      A($('.k2-hon.ha'), [{ opacity: 1 }, { opacity: 0 }], SEL + 40, 160, 'ease-out', 'forwards');
      A($('.k2-hon.he'), [{ opacity: 0 }, { opacity: 1 }], SEL + 110, 200, 'ease-out');
      /* the name box rolls like a slot: A1 leaves upward as E1 arrives from below */
      A($('.k2-nb.a'), [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(-75%)' }], SEL + 40, 200, INOUT, 'forwards');
      A($('.k2-nb.e'), [{ opacity: 0, transform: 'translateY(75%)' }, { opacity: 1, transform: 'none' }], SEL + 40, 200, INOUT);

      /* the formula types itself over Tika's words; its ranges light up as each reference is complete */
      var ftw = $('.k2-ftw');
      var clipK = [{ offset: 0, clipPath: 'inset(-3px 100% -3px 0)', easing: HOLD }], carK = [{ offset: 0, transform: 'translateX(0%)', easing: HOLD }];
      for (i = 0; i < N; i++) {
        var off = tt[i] / TD;
        clipK.push({ offset: off, clipPath: 'inset(-3px ' + ((N - 1 - i) / N * 100).toFixed(3) + '% -3px 0)', easing: HOLD });
        carK.push({ offset: off, transform: 'translateX(' + ((i + 1) / N * 100).toFixed(3) + '%)', easing: HOLD });
      }
      A(ftw, clipK, TYPE, TD, 'linear');
      A($('.k2-cw'), carK, TYPE, TD, 'linear');
      A($('.k2-caret'), [{ opacity: 0 }, { opacity: 1 }], TYPE - 30, 1, 'linear');
      A($('.k2-caret'), [{ opacity: 1 }, { opacity: 0 }], ENTER, 1, 'linear', 'forwards');
      $$('.k2-ic .on').forEach(function (el) {
        A(el, [{ opacity: 0 }, { opacity: 1 }], TYPE, 160, 'ease-out');
        A(el, [{ opacity: 1 }, { opacity: 0 }], ENTER + 120, 220, 'ease-out', 'forwards');
      });
      A($('.k2-ic.v'), [{ transform: 'scale(1)' }, { transform: 'scale(1.55)', offset: 0.35 }, { transform: 'scale(1)' }], ENTER - 20, 300, QUART, 'none');
      /* status bar: Ready, then Enter while typing, then Ready again */
      var stOut = [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateY(-60%)' }];
      var stIn = [{ opacity: 0, transform: 'translateY(60%)' }, { opacity: 1, transform: 'none' }];
      A($('.st .rdy'), stOut, TYPE - 40, 90, 'ease-in');
      A($('.st .ent'), stIn, TYPE + 50, 160, EXPO);
      A($('.st .ent'), stOut, ENTER + 40, 90, 'ease-in', 'forwards');
      A($('.st .rdy'), stIn, ENTER + 130, 160, EXPO, 'forwards');

      var one = function (el, at) { A(el, [{ opacity: 0, transform: 'scale(1.18)' }, { opacity: 1, transform: 'none' }], at, 180, QUART); };
      var grow = function (oneEl, fullEl, at) {
        var startClip = 'inset(-6px -6px ' + (fullEl.offsetHeight - RH - 4) + 'px -6px)';
        A(fullEl, [{ clipPath: startClip }, { clipPath: 'inset(-6px -6px -6px -6px)' }], at, 260, EXPO);
        A(fullEl, [{ opacity: 0 }, { opacity: 1 }], at, 1, 'linear');
        A(oneEl, [{ opacity: 1 }, { opacity: 0 }], at + 20, 1, 'linear', 'forwards');
      };
      one($('.k2-rg.one.r1'), TYPE + tt[9]);
      grow($('.k2-rg.one.r1'), $('.k2-rg.full.r1'), TYPE + tt[12]);
      one($('.k2-rg.one.r2'), TYPE + tt[15]);
      grow($('.k2-rg.one.r2'), $('.k2-rg.full.r2'), TYPE + tt[18]);
      one($('.k2-rg.one.r3'), TYPE + tt[21]);

      /* Enter: the matching rows light, the sales fly into E1 and add up */
      A($('.k2-sel'), [{ transform: 'scale(1)' }, { transform: 'scale(1.08)', offset: 0.35 }, { transform: 'scale(1)' }], ENTER, 320, QUART, 'none');
      A($('.k2-ftw .plain'), [{ opacity: 0 }, { opacity: 1 }], ENTER + 40, 240, 'ease-out');
      A($('.k2-rngs'), [{ opacity: 1 }, { opacity: 0 }], ENTER + 160, 280, 'ease-out', 'forwards');
      $$('.k2-hit').forEach(function (el, k) {
        A(el, [{ opacity: 0 }, { opacity: 1 }], ENTER + 30 + (k >> 1) * 40, 200, 'ease-out');
        A(el, [{ opacity: 1 }, { opacity: 0 }], RESULT + 80, 300, 'ease-out', 'forwards');
      });
      $$('.k2-c').forEach(function (el) {
        var rr2 = +el.getAttribute('data-r');
        if (rr2 === 5 || rr2 === 7 || rr2 === 9) A(el, [{ opacity: 1 }, { opacity: 0.3 }], ENTER + 40, 260, 'ease-out', 'forwards');
      });
      $$('.k2-chip').forEach(function (el, k) {
        var dy = (HITS[k] - 1) * RH, dx = Math.round(f * 1.1), at = CHIPS + k * CH_GAP, pillEl = el.firstChild, pe = PE(k);
        A(el, [
          { transform: 'translate(0px,0px)' },
          { transform: 'translate(0px,' + (-0.18 * f).toFixed(1) + 'px)', offset: 0.16 },
          { transform: 'translate(' + dx + 'px,' + (-dy * 0.6).toFixed(1) + 'px)', offset: 0.6 },
          { transform: 'translate(' + Math.round(dx * 0.25) + 'px,' + (-dy * 0.96).toFixed(1) + 'px)', offset: 0.88 },
          { transform: 'translate(0px,' + (-dy) + 'px)' }
        ], at, CH_DUR, FLY_E);
        A(pillEl, [
          { offset: 0, opacity: 0, transform: 'scale(.9)' },
          { offset: 0.05, opacity: 1, transform: 'scale(1.1)' },
          { offset: 0.3, opacity: 1, transform: 'scale(1)' },
          { offset: pe - 0.1, opacity: 1, transform: 'scale(.95)' },
          { offset: pe + 0.07, opacity: 0, transform: 'scale(.62)' },
          { offset: 1, opacity: 0, transform: 'scale(.62)' }
        ], at, CH_DUR, FLY_E);
      });
      var FIN = spring(0.42, 560);
      $$('.k2-val').forEach(function (el, k) {
        A(el, [{ opacity: 0 }, { opacity: 1 }], ARR(k), 1, 'linear');
        if (k < 2) {
          A(el.firstChild, [{ transform: 'translateY(.45em) scale(.9)' }, { transform: 'none' }], ARR(k), 200, EXPO);
          A(el, [{ opacity: 1 }, { opacity: 0 }], ARR(k + 1), 1, 'linear', 'forwards');
        } else {
          A(el.firstChild, [{ transform: 'scale(.45)' }, { transform: 'none' }], ARR(k), 560, FIN);
        }
      });
      A($('.k2-ring'), [{ opacity: 0.95, transform: 'scale(1)' }, { opacity: 0, transform: 'scale(1.22,2.1)' }], RESULT, 620, EXPO, 'forwards');

      /* ================= Seam 2: the answer cell fills with gold and becomes the canary ================= */
      /* the canary stands on E1: its feet near the cell's middle, its body about twice a row tall */
      var u = RH * 2.2 / 16.2;
      var bcx = ex - 0.45 * u, bcy = ey + 0.15 * RH - 8.4 * u;
      var birdPoly = bp.map(function (p) { return [bcx + (p[0] - 12) * u, bcy + (p[1] - 12) * u]; });
      /* the cell as the same number of points, walked the same way round (head, back, tail, feet, breast, beak),
         starting above the head, so the polygon morph reads as one shape changing */
      var cellPoly = (function () {
        var L0 = E1.l, T_ = E1.t, R1 = E1.l + E1.w, B_ = E1.t + E1.h;
        var sx = Math.max(L0 + 1, Math.min(R1 - 1, bcx + 3.2 * u));
        var V = [[sx, T_], [L0, T_], [L0, B_], [R1, B_], [R1, T_], [sx, T_]], seg = [], Pm = 0;
        for (var j = 0; j < 5; j++) { var l = Math.hypot(V[j + 1][0] - V[j][0], V[j + 1][1] - V[j][1]); seg.push(l); Pm += l; }
        var out = [];
        for (j = 0; j < NPT; j++) {
          var dd = Pm * j / NPT, sg = 0;
          while (sg < 4 && dd > seg[sg]) { dd -= seg[sg]; sg++; }
          var t = seg[sg] ? dd / seg[sg] : 0;
          out.push([V[sg][0] + (V[sg + 1][0] - V[sg][0]) * t, V[sg][1] + (V[sg + 1][1] - V[sg][1]) * t]);
        }
        /* snap the nearest samples onto the corners, so the cell starts square */
        var a2 = 0;
        for (j = 1; j <= 4; j++) { a2 += seg[j - 1]; out[Math.round(a2 / Pm * NPT) % NPT] = V[j].slice(); }
        return out;
      })();
      var flatPoly = cellPoly.map(function (p) { return [p[0], E1.t + E1.h]; });
      var poly = function (pts) { return 'polygon(' + pts.map(function (p) { return p[0].toFixed(1) + 'px ' + p[1].toFixed(1) + 'px'; }).join(',') + ')'; };
      var gold = $('.k3-gold'), morph = $('.k3-morph'), num = $('.k3-num');
      /* the camera pushes into E1 while the answer lands; the gold layer rides the same push, so it stays on the cell */
      /* the push stops where the whole window still fits on screen (on a phone that is a gentler push), and the pan toward
         the centre is kept inside the frame, so the file name and column A are never cut off */
      var EDGE = 8, Z = Math.max(1.06, Math.min(1.32, (Wd - 2 * EDGE) / camR.width));
      var wL = ex + (camR.left - ex) * Z, wR = ex + (camR.right - ex) * Z, PX = (Wd / 2 - ex) * 0.3;
      PX = Math.round(wR - wL <= Wd - 2 * EDGE ? Math.min(Wd - EDGE - wR, Math.max(EDGE - wL, PX)) : Wd / 2 - (wL + wR) / 2);
      var PY = Math.round((H / 2 - ey) * 0.2), PUSH = 'cubic-bezier(.45,0,.2,1)';
      cam.style.transformOrigin = (ex - camR.left).toFixed(1) + 'px ' + (ey - camR.top).toFixed(1) + 'px';
      morph.style.transformOrigin = ex.toFixed(1) + 'px ' + ey.toFixed(1) + 'px';
      [cam, morph].forEach(function (el) {
        A(el, [{ transform: 'translate(0px,0px) scale(1)' }, { transform: 'translate(' + PX + 'px,' + PY + 'px) scale(' + Z + ')' }], RESULT - 60, 560, PUSH);
      });
      num.style.cssText = 'left:' + E1.l + 'px;top:' + E1.t + 'px;width:' + E1.w + 'px;height:' + E1.h + 'px;padding-right:' + PAD + 'px;font-size:' + (f * 0.92).toFixed(2) + 'px';
      A(num.firstChild, [{ transform: 'scale(.45)' }, { transform: 'none' }], RESULT, 560, FIN);
      A(morph, [{ opacity: 0 }, { opacity: 1 }], GOLD, 1, 'linear');
      A(morph, [{ opacity: 1 }, { opacity: 0 }], T0, 1, 'linear', 'forwards');
      A(gold, [{ clipPath: poly(flatPoly) }, { clipPath: poly(cellPoly) }], GOLD, GD, 'cubic-bezier(.3,.7,.4,1)');
      A(gold, [{ clipPath: poly(cellPoly) }, { clipPath: poly(birdPoly) }], MORPH, MD, 'cubic-bezier(.6,0,.3,1)', 'forwards');
      gold.style.transformOrigin = bcx.toFixed(1) + 'px ' + (bcy + 4 * u).toFixed(1) + 'px';
      A(gold, [{ transform: 'none' }, { transform: 'scale(1.08)', offset: 0.7 }, { transform: 'none' }], MORPH, MD, 'ease-out', 'none');
      A(num, [{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'scale(.85)' }], MORPH, 70, 'ease-in', 'forwards');

      /* ================= Act 3: the canary takes off, swoops back and perches ================= */
      var birdA = $('.k3-bird-a');
      A(birdA, [{ opacity: 0 }, { opacity: 1 }], T0, 1, 'linear');
      A(birdA.querySelector('.k3-eye'), [{ transform: 'scale(0)' }, { transform: 'scale(1.4)', offset: 0.55 }, { transform: 'scale(1)' }], T0 - 20, 220, 'cubic-bezier(.2,.7,.3,1)');

      /* where the lockup's bird slot is (layout offsets ignore transforms, so a re-measure during the push still aims true) */
      function measure() {
        var el = root.querySelector('.k3-solo .k3-slot'), sx = 0, sy = 0, e = el, w = el.offsetWidth;
        while (e && e !== root) { sx += e.offsetLeft; sy += e.offsetTop; e = e.offsetParent; }
        return { vw: root.clientWidth || vw, vh: root.clientHeight || vh, sx: sx + w / 2, sy: sy + el.offsetHeight / 2, s: w, k: w / D };
      }
      /* aim the curtain's diagonal through the gap between "i" and "c", so the logo splits cleanly between letters */
      function aimCut() {
        var ch = root.querySelectorAll('.k3-solo .k3-ch'), cw = root.clientWidth || vw, chh = root.clientHeight || vh;
        var offp = function (el) { var ox = 0, oy = 0, e = el; while (e && e !== root) { ox += e.offsetLeft; oy += e.offsetTop; e = e.offsetParent; } return [ox, oy]; };
        var a = offp(ch[1]), c = offp(ch[2]);
        var gx = (a[0] + ch[1].offsetWidth + c[0]) / 2, gy = a[1] + ch[1].offsetHeight * .55;
        var t = Math.min(63, Math.max(53, gx / cw * 100 + 16 * gy / chh));
        root.style.setProperty('--ct', t.toFixed(2) + '%');
        root.style.setProperty('--cb', (t - 16).toFixed(2) + '%');
      }
      var m = measure();
      aimCut();

      /* the canary takes over from the gold shape exactly where the camera push left it */
      var s0 = 24 * u * Z / D, LR = LAND - T0, ER = END - T0;
      var bx0 = ex + PX + (bcx - ex) * Z, by0 = ey + PY + (bcy - ey) * Z;
      /* The flight, as tracks: x, y, and turn plus scale (the squash and wings ride on their own).
         One continuous flight that never leaves the screen: the canary lifts off the answer cell, rises, turns around
         at the top of its arc (its width passes through zero, which reads as a turn), glides down toward the lockup,
         turns back to face "Kicau" as it brakes, and perches. Over the sheet it shows gold; over the yellow it shows ink. */
      function flight(mm) {
        var h = D / 2, k = mm.k;
        var sApex = Math.max(s0 * 1.25, k * .78);
        var half = D * sApex * .6;
        var ax = Math.min(bx0 + mm.vw * .05, mm.vw - half), ay = Math.max(half, Math.min(by0, mm.sy) - mm.vh * .2);
        var px = mm.sx + mm.s * .4, py = mm.sy - mm.vh * .11;
        var X = function (v) { return { transform: 'translateX(' + (v - h).toFixed(1) + 'px)' }; };
        var Y = function (v) { return { transform: 'translateY(' + (v - h).toFixed(1) + 'px)' }; };
        var Rt = function (r, s, dir) { return { transform: 'rotate(' + r + 'deg) scale(' + ((dir || 1) * s).toFixed(4) + ',' + s.toFixed(4) + ')' }; };
        return {
          fx: kfs([[0, X(bx0), 'cubic-bezier(.3,0,.2,1)'], [70, X(bx0 - .05 * s0 * D), 'cubic-bezier(.3,.1,.3,1)'], [210, X(ax), 'cubic-bezier(.4,0,.6,1)'],
            [290, X(ax - mm.vw * .02), 'cubic-bezier(.35,0,.4,1)'], [560, X(px), 'cubic-bezier(.3,0,.3,1)'], [665, X(mm.sx + mm.s * .08), 'cubic-bezier(.3,0,.2,1)'], [LR, X(mm.sx)], [ER, X(mm.sx)]]),
          fy: kfs([[0, Y(by0), 'cubic-bezier(.3,0,.2,1)'], [70, Y(by0 + .04 * s0 * D), 'cubic-bezier(.2,.6,.4,1)'], [210, Y(ay), 'cubic-bezier(.4,0,.6,1)'],
            [290, Y(ay - mm.vh * .01), 'cubic-bezier(.45,0,.55,1)'], [560, Y(py), 'cubic-bezier(.3,0,.3,1)'], [665, Y(mm.sy - mm.vh * .025), 'cubic-bezier(.4,0,.2,1)'], [LR, Y(mm.sy)], [ER, Y(mm.sy)]]),
          fr: kfs([[0, Rt(0, s0, 1), 'cubic-bezier(.3,0,.2,1)'], [70, Rt(4, s0, 1), 'cubic-bezier(.2,.6,.3,1)'], [210, Rt(-18, sApex, 1), 'cubic-bezier(.5,0,.5,1)'],
            [290, Rt(-6, sApex * 1.02, -1), 'cubic-bezier(.3,0,.4,1)'], [560, Rt(-10, k * 1.08, -1), 'cubic-bezier(.5,0,.5,1)'], [665, Rt(-4, k * 1.06, 1), 'cubic-bezier(.3,1.5,.5,1)'],
            [LR + 50, Rt(0, k, 1)], [ER, Rt(0, k, 1)]])
        };
      }
      var SQ = function (a, b) { return { transform: 'scale(' + a + ',' + b + ')' }; };
      var SQUASH = kfs([[0, SQ(1, 1), 'cubic-bezier(.3,0,.2,1)'], [80, SQ(1.08, .86), 'cubic-bezier(.3,0,.3,1)'], [150, SQ(.92, 1.1), 'cubic-bezier(.4,0,.2,1)'], [270, SQ(1, 1)],
        [LR - 50, SQ(1, 1), 'cubic-bezier(.5,0,1,.5)'], [LR, SQ(1.13, .84), 'cubic-bezier(.3,0,.3,1)'], [LR + 100, SQ(.96, 1.05), 'cubic-bezier(.4,0,.2,1)'], [LR + 220, SQ(1, 1)], [ER, SQ(1, 1)]]);
      var Wg = function (r, o) { return { transform: 'rotate(' + r + 'deg)', opacity: o }; };
      var FLAPS = kfs([[0, Wg(-6, 0)], [40, Wg(-6, 0), IO], [75, Wg(24, 1), IO], [115, Wg(58, 1), IO], [165, Wg(-20, 1), IO], [215, Wg(58, 1), IO], [262, Wg(-20, 1), IO], [300, Wg(40, 1), IO],
        [340, Wg(50, 1), IO], [400, Wg(-20, 1), IO], [460, Wg(50, 1), IO], [520, Wg(-16, 1), IO], [585, Wg(44, 1), IO], [LR - 20, Wg(24, 1), IO], [LR + 30, Wg(-4, 1), IO], [LR + 90, Wg(-6, 0)], [ER, Wg(-6, 0)]]);
      var fl = flight(m), birdTracks = [];
      ['fx', 'fy', 'fr'].forEach(function (key) {
        $$('.k3-' + key).forEach(function (el) { birdTracks.push([key, A(el, fl[key], T0, ER, 'linear')]); });
      });
      $$('.k3-sq').forEach(function (el) { A(el, SQUASH, T0, ER, 'linear'); });
      $$('.k3-wing').forEach(function (el) { A(el, FLAPS, T0, ER, 'linear'); });

      /* the three-yellow sweep follows it across: pale, deep, then Kenari Yellow (which carries the ink canary and the lockup) */
      var POLY0 = 'polygon(-50% 0%,-1% 0%,-26% 100%,-50% 100%)', POLY1 = 'polygon(-50% 0%,126% 0%,101% 100%,-50% 100%)';
      A($('.k3-s1'), [{ clipPath: POLY0 }, { clipPath: POLY1 }], T0 + 30, 540, SWEEP);
      A($('.k3-s2'), [{ clipPath: POLY0 }, { clipPath: POLY1 }], T0 + 75, 540, SWEEP);
      A($('.k3-yel'), [{ clipPath: POLY0 }, { clipPath: POLY1 }], T0 + 120, 540, SWEEP);
      /* once the yellow covers everything, the layers under it go, so the curtain opens straight onto the app */
      $$('.k2-bg,.k2-stage,.k1-chat,.k1-fly,.k3-s1,.k3-s2,.k3-night').forEach(function (el) { A(el, [{ opacity: 1 }, { opacity: 0 }], T0 + 680, 1, 'linear', 'forwards'); });
      $$('.k3-grid').forEach(function (el) { A(el, [{ transform: 'translate(0,0)' }, { transform: 'translate(-24px,-10px)' }], T0, ER, 'linear'); });

      /* ================= the lockup: chirp, "Kicau", "Kenari Excel Lab" ================= */
      $$('.k3-arc').forEach(function (el, k) {
        A(el, [{ strokeDashoffset: 1, opacity: 1 }, { strokeDashoffset: 0, opacity: 1, offset: .45 }, { strokeDashoffset: 0, opacity: 0 }], LAND - 30 + (k % 3) * 55, 420, 'cubic-bezier(.3,.6,.4,1)');
      });
      var sp = springKD(420, 24);
      $$('.k3-ch').forEach(function (el, k) {
        A(el, [{ transform: 'translateY(108%) rotate(9deg)' }, { transform: 'translateY(0) rotate(0deg)' }], LAND - 50 + (k % 5) * 40, sp.d, sp.e);
      });
      var sc = $$('.k3-sc'), nsc = 'Kenari Excel Lab'.length, mid = (nsc - 1) / 2;
      sc.forEach(function (el, k) {
        A(el, [{ transform: 'translateX(' + ((k % nsc - mid) * .62).toFixed(2) + 'em)', opacity: 0 }, { opacity: 1, offset: .5 }, { transform: 'translateX(0)', opacity: 1 }], LAND + 60, 480, EXPO);
      });
      $$('.k3-rule').forEach(function (el) { A(el, [{ transform: 'scaleX(0)' }, { transform: 'scaleX(1)' }], LAND + 110, 420, SOFT); });
      $$('.k3-bird-b .k3-eye').forEach(function (el) { A(el, [{ transform: 'scaleY(1)' }, { transform: 'scaleY(.12)', offset: 0.45 }, { transform: 'scaleY(1)' }], LAND + 430, 170, 'ease-in-out', 'none'); });
      /* the finished lockup never sits dead: a slow camera push while it holds */
      $$('.k3-top').forEach(function (el) { A(el, [{ transform: 'scale(1)' }, { transform: 'scale(1.035)' }], LAND + 60, END - LAND - 60, 'cubic-bezier(.3,0,.4,1)'); });

      /* ================= exit: the curtain tears open on the diagonal and carries the two halves of the logo off ================= */
      /* on the frame the curtain starts (it has not moved yet) the whole lockup hands over to the two clipped copies */
      A(sk, [{ opacity: 1 }, { opacity: 0 }], CURT - 60, 160, 'ease-in', 'forwards');
      $$('.k3-split').forEach(function (el) { A(el, [{ opacity: 0 }, { opacity: 1 }], CURT, 1, 'linear'); });
      A($('.k3-solo'), [{ opacity: 1 }, { opacity: 0 }], CURT, 1, 'linear', 'forwards');
      A($('.k3-half.k3-cl'), [{ transform: 'translateX(0)' }, { transform: 'translateX(-64%)' }], CURT, 380, CURTAIN, 'forwards');
      A($('.k3-half.k3-cr'), [{ transform: 'translateX(0)' }, { transform: 'translateX(64%)' }], CURT, 380, CURTAIN, 'forwards');
      A($('.k3-lin-l'), [{ transform: 'translateX(0)' }, { transform: 'translateX(-64%)' }], CURT + 20, 380, CURTAIN, 'forwards');
      var last = A($('.k3-lin-r'), [{ transform: 'translateX(0)' }, { transform: 'translateX(64%)' }], CURT + 20, 380, CURTAIN, 'forwards');
      last.finished.then(function () { end(); }, function () {});
      /* hard cap, so a slow first frame can never keep the overlay up */
      timer = setTimeout(function () { end(); }, END + 200);

      /* if the web font arrives late, the slot moves: re-aim the landing and the cut */
      remeasure = function () {
        if (over) return;
        var t = birdTracks.length ? birdTracks[0][1].currentTime : 0;
        if (t < CURT - 10) aimCut();
        if (t > LAND - 120) return;
        var mm = measure();
        if (Math.abs(mm.sx - m.sx) < .5 && Math.abs(mm.sy - m.sy) < .5 && Math.abs(mm.vw - m.vw) < .5) return;
        m = mm; var f2 = flight(m);
        birdTracks.forEach(function (p) { p[1].effect.setKeyframes(f2[p[0]]); });
      };
      if (d.fonts && d.fonts.ready) d.fonts.ready.then(remeasure);
      W.addEventListener('resize', remeasure);
    }
  };
})();
