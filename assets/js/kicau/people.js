/*
 * The people of Kicau: Rasa Kenari's colleagues, the mentor, and Bank Kenari's team.
 * Faces are drawn in SVG so every colleague is a person, not a letter. Everyone here is invented.
 * Classic script: window.KicauPeople.
 */
(function (root) {
  'use strict';

  const PEOPLE = {
    tika: { name: 'Tika Rahayu', short: 'Tika', role: 'Sales Finance', org: 'Rasa Kenari', f: { bg: '#FFE4D6', skin: '#C88B63', style: 'hijab', hijab: '#E2603A', top: '#B8452A' } },
    bayu: { name: 'Bayu Prakoso', short: 'Bayu', role: 'Warehouse Supervisor', org: 'Rasa Kenari', f: { bg: '#DCF2E3', skin: '#A86F48', style: 'cap', cap: '#2F855A', top: '#276749' } },
    anton: { name: 'Pak Anton', short: 'Pak Anton', role: 'Area Sales Manager', org: 'Rasa Kenari', f: { bg: '#EEEBE5', skin: '#C08A64', style: 'short', hair: '#A7A29C', top: '#3F3F46', glasses: true } },
    lestari: { name: 'Bu Lestari', short: 'Bu Lestari', role: 'Credit Control', org: 'Rasa Kenari', f: { bg: '#FCE7F3', skin: '#B97F59', style: 'bob', hair: '#1F1A17', top: '#BE185D' } },
    sari: { name: 'Mbak Sari', short: 'Mbak Sari', role: 'Senior Analyst, your mentor', org: 'Rasa Kenari', f: { bg: '#E9E5FF', skin: '#D9A27C', style: 'bun', hair: '#3A2A22', top: '#6D5BD0', glasses: true } },
    citra: { name: 'Citra', short: 'Citra', role: 'Head of Collection', org: 'Bank Kenari', bank: true, f: { bg: '#DFECFF', skin: '#C99A72', style: 'hijab', hijab: '#2C5FB0', top: '#1E3A8A' } },
    dimas: { name: 'Dimas', short: 'Dimas', role: 'Portfolio Analyst', org: 'Bank Kenari', bank: true, f: { bg: '#E0F2FE', skin: '#B07A52', style: 'side', hair: '#1C1917', top: '#0369A1' } },
    me: { name: 'You', short: 'You', role: 'Junior Analyst', org: 'Rasa Kenari', f: { bg: 'var(--brand-soft)', skin: '#C99068', style: 'side', hair: '#231A15', top: 'var(--me)' } }
  };

  /** One drawn face, 40 by 40, clipped to a circle by #kc-clip in the page's sprite. */
  function face(f) {
    const s = (c) => `style="fill:${c}"`;
    const hair = f.hair || '#231A15';
    const hijab = f.style === 'hijab';
    const ey = hijab ? 19.4 : 17.6;
    const body = `<path ${s(f.top)} d="M5 41c.6-7.8 6.8-12 15-12s14.4 4.2 15 12z"/>`;
    let back = '', front = '';
    if (hijab) back = `<path ${s(f.hijab)} d="M10.4 22.4c0-7.6 4.2-13.2 9.6-13.2s9.6 5.6 9.6 13.2v6.2c-2.6 2.2-5.8 3.2-9.6 3.2s-7-1-9.6-3.2z"/>`;
    if (f.style === 'bob') back = `<path ${s(hair)} d="M11.4 19.4c0-6.8 3.8-10.8 8.6-10.8s8.6 4 8.6 10.8v5.4h-3.4v-7.4c-2.4-1-4.2-3-5.2-5-1 2-2.8 4-5.2 5v7.4h-3.4z"/>`;
    const head = hijab
      ? `<ellipse ${s(f.skin)} cx="20" cy="20" rx="5.7" ry="6.6"/>`
      : `<rect ${s(f.skin)} x="17.5" y="22" width="5" height="6" rx="1.4"/><circle ${s(f.skin)} cx="20" cy="17.6" r="7.3"/>`;
    if (f.style === 'short' || f.style === 'bun' || f.style === 'side') front = `<path ${s(hair)} d="M12.5 16.8c0-5.4 3.3-8.7 7.5-8.7s7.5 3.3 7.5 8.7c-1.4-2.8-4.4-4.2-7.5-4.2-2 0-4 .6-5 1.6-.8.8-1.8 1.6-2.5 2.6z"/>`;
    if (f.style === 'side') front += `<path ${s(hair)} d="M12.6 15.2c1.8-3.6 5-5 8.2-4.6-2.6 1.4-5.2 3.4-8.2 4.6z"/>`;
    if (f.style === 'bun') front += `<circle ${s(hair)} cx="20" cy="8.4" r="3.4"/>`;
    if (f.style === 'cap') front = `<path ${s(f.cap)} d="M12.3 15.6c0-4.8 3.4-7.9 7.7-7.9s7.7 3.1 7.7 7.9z"/><path ${s(f.cap)} d="M19 14.2h11.2c1 0 1 1.8 0 1.8H19z"/>`;
    const eyes = `<circle style="fill:#241C18" cx="17.4" cy="${ey}" r=".9"/><circle style="fill:#241C18" cx="22.6" cy="${ey}" r=".9"/><path d="M17.9 ${ey + 3.1}q2.1 1.7 4.2 0" style="fill:none;stroke:#241C18;stroke-width:.95;stroke-linecap:round"/>`;
    const gl = f.glasses ? `<g style="fill:none;stroke:#241C18;stroke-width:.8"><circle cx="17.4" cy="${ey}" r="2.3"/><circle cx="22.6" cy="${ey}" r="2.3"/><path d="M19.7 ${ey}h.6"/></g>` : '';
    const cheeks = `<circle style="fill:#E86A5A;opacity:.18" cx="15.6" cy="${ey + 2.6}" r="1.5"/><circle style="fill:#E86A5A;opacity:.18" cx="24.4" cy="${ey + 2.6}" r="1.5"/>`;
    return `<svg viewBox="0 0 40 40" class="face" aria-hidden="true"><g clip-path="url(#kc-clip)"><rect ${s(f.bg)} width="40" height="40"/>${hijab ? body + back : back + body}${head}${front}${eyes}${cheeks}${gl}</g></svg>`;
  }

  /** An avatar: the face, plus a green dot when online, or grey when locked. */
  function avatar(key, size, opt) {
    const o = opt || {};
    const p = PEOPLE[key];
    return `<span class="av${o.dim ? ' dim' : ''}" style="--s:${size}px" role="img" aria-label="${p.name}">${face(p.f)}${o.on ? '<i class="on"></i>' : ''}</span>`;
  }

  root.KicauPeople = { PEOPLE, face, avatar };
})(typeof globalThis !== 'undefined' ? globalThis : this);
