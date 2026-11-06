/*
 * The two made-up companies behind Kenari Excel Lab.
 *
 * Rasa Kenari sells noodles, sachet coffee, sambal and soap through distributors.
 * Bank Kenari lends working capital to those same distributors. Every company, person,
 * product and number here is invented. The generator is seeded, so the browser and the
 * Node tests always build exactly the same rows.
 *
 * Classic script in the browser (window.KenariData), CommonJS in Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.KenariData = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const AS_OF = '2026-10-31';
  const MONTHS = ['2026-05', '2026-06', '2026-07', '2026-08', '2026-09', '2026-10'];

  /* Web-searched on 2026-11-05: no brand found under these names. See docs/NAMES.md. */
  const PRODUCTS = [
    ['MLK-GOR', 'Mi Lekas Goreng', 'Noodles', 40, 112000],
    ['MLK-SOT', 'Mi Lekas Soto', 'Noodles', 40, 108000],
    ['MLK-KAR', 'Mi Lekas Kari', 'Noodles', 40, 110000],
    ['KSJ-SUS', 'Kopi Sekejap Susu', 'Coffee', 120, 186000],
    ['KSJ-HIT', 'Kopi Sekejap Hitam', 'Coffee', 120, 142000],
    ['KSJ-ARE', 'Kopi Sekejap Gula Aren', 'Coffee', 120, 198000],
    ['SNY-BAW', 'Sambal Nyala Bawang', 'Sambal', 24, 216000],
    ['SNY-IJO', 'Sambal Nyala Ijo', 'Sambal', 24, 228000],
    ['SNY-TER', 'Sambal Nyala Terasi', 'Sambal', 24, 222000],
    ['SRW-SER', 'Sabun Rimba Wangi Serai', 'Soap', 72, 252000],
    ['SRW-MEL', 'Sabun Rimba Wangi Melati', 'Soap', 72, 252000],
    ['SRW-KAY', 'Sabun Rimba Wangi Kayu', 'Soap', 72, 264000]
  ].map(([sku, name, category, packs_per_carton, carton_price]) => ({ sku, name, category, packs_per_carton, carton_price }));

  /* Named after Indonesian birds, to keep the canary company. `trouble` marks the planted early-warning story. */
  const DISTRIBUTORS = [
    ['D01', 'CV Nuri Makmur', 'Bekasi', 'Jabodetabek', 30],
    ['D02', 'UD Jalak Sentosa', 'Tangerang', 'Jabodetabek', 30],
    ['D03', 'CV Kutilang Abadi', 'Bandung', 'Jawa Barat', 45],
    ['D04', 'UD Murai Niaga', 'Cirebon', 'Jawa Barat', 30],
    ['D05', 'CV Kepodang Niaga', 'Semarang', 'Jawa Tengah', 45, true],
    ['D06', 'UD Seriti Jaya', 'Solo', 'Jawa Tengah', 30],
    ['D07', 'CV Rangkong Lestari', 'Surabaya', 'Jawa Timur', 45],
    ['D08', 'UD Pipit Karya', 'Malang', 'Jawa Timur', 30, true],
    ['D09', 'CV Tekukur Rejeki', 'Denpasar', 'Bali Nusra', 30],
    ['D10', 'UD Cucak Berkah', 'Mataram', 'Bali Nusra', 45],
    ['D11', 'CV Bangau Sejahtera', 'Medan', 'Sumatera', 45],
    ['D12', 'UD Srigunting Utama', 'Palembang', 'Sumatera', 30, true],
    ['D13', 'CV Kasuari Sentral', 'Makassar', 'Sulawesi', 45],
    ['D14', 'UD Kedasi Prima', 'Manado', 'Sulawesi', 30]
  ].map(([id, name, city, region, terms_days, trouble]) => ({ id, name, city, region, terms_days, trouble: !!trouble }));

  /* The people who ask for things. Rasa Kenari is new here. Citra and Dimas come from the SQL lab's bank. */
  const CAST = {
    tika: { name: 'Tika', role: 'Sales Finance, Rasa Kenari' },
    anton: { name: 'Pak Anton', role: 'Area Sales Manager, Rasa Kenari' },
    lestari: { name: 'Bu Lestari', role: 'Credit Control, Rasa Kenari' },
    citra: { name: 'Citra', role: 'Head of Collection, Bank Kenari' },
    dimas: { name: 'Dimas', role: 'Portfolio Analyst, Bank Kenari' }
  };

  const dayMs = 864e5;
  const toMs = (iso) => Date.parse(iso + 'T00:00:00Z');
  const toIso = (ms) => new Date(ms).toISOString().slice(0, 10);
  const addDays = (iso, n) => toIso(toMs(iso) + n * dayMs);
  /** ISO date to an Excel date number, so date formulas work as they do in Excel. */
  const toSerial = (iso) => Math.round((toMs(iso) - toMs('1899-12-30')) / dayMs);
  const kol = (d) => (d === 0 ? 1 : d <= 90 ? 2 : d <= 120 ? 3 : d <= 180 ? 4 : 5);
  const lastDay = (m) => toIso(Date.UTC(+m.slice(0, 4), +m.slice(5, 7), 0));

  let cache = null;

  /** Everything, as plain objects. The same every time. */
  function build() {
    if (cache) return cache;
    let s = 20261105;
    const R = () => {
      s |= 0; s = (s + 0x6D2B79F5) | 0;
      let t = Math.imul(s ^ (s >>> 15), 1 | s);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const int = (a, b) => a + Math.floor(R() * (b - a + 1));
    const pick = (a) => a[Math.floor(R() * a.length)];
    const round = (v, to) => Math.round(v / to) * to;

    /* Sales fall for the planted distributors over the last three months. */
    const FADE = [1, 1, 0.92, 0.72, 0.5, 0.36];
    const LATE = [0, 0, 14, 38, 67, 96];

    const sales = [];
    let inv = 1000;
    for (const d of DISTRIBUTORS) {
      const size = 0.7 + R() * 0.9;
      d.tier_discount = pick([0, 2, 2, 3, 5]);
      MONTHS.forEach((m, mi) => {
        const lines = int(5, 9);
        const fade = d.trouble ? FADE[mi] : 1;
        for (let i = 0; i < lines; i++) {
          const p = pick(PRODUCTS);
          const date = addDays(m + '-01', int(0, 27));
          const cartons = Math.max(2, Math.round(int(8, 48) * size * fade));
          const gross = cartons * p.carton_price;
          const net = round(gross * (1 - d.tier_discount / 100), 100);
          const due = addDays(date, d.terms_days);
          /* Paid late more often as trouble grows. Unpaid when the due date is after the data date or it is badly overdue. */
          const lateBy = d.trouble ? int(0, LATE[mi] + 8) : (R() < 0.12 ? int(1, 12) : 0);
          const paidOn = addDays(due, lateBy - int(0, 6));
          sales.push({
            invoice: 'INV-' + (++inv), date, distributor_id: d.id, distributor: d.name, region: d.region,
            sku: p.sku, product: p.name, category: p.category,
            cartons, gross, discount_pct: d.tier_discount, net, due_date: due,
            paid_date: paidOn <= AS_OF ? paidOn : null
          });
        }
      });
    }
    sales.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.invoice < b.invoice ? -1 : 1));

    /* Warungs: the small shops at the end of the chain. */
    const OUTLET_WORD = ['Warung Bu Ani', 'Warung Pojok', 'Toko Sumber', 'Warung Pak Dul', 'Toko Berkah', 'Warung Hijau', 'Kios Pasar', 'Warung Sudut'];
    const outlets = [];
    DISTRIBUTORS.forEach((d) => {
      const n = int(3, 6);
      for (let i = 0; i < n; i++) outlets.push({ outlet_id: 'W' + String(outlets.length + 1).padStart(3, '0'), outlet: pick(OUTLET_WORD) + ' ' + d.city, distributor_id: d.id, city: d.city, visits_per_month: int(1, 4) });
    });

    /* Bank Kenari: a working-capital line for everyone, a term loan for some. */
    const monthlyNet = (id) => sales.filter((x) => x.distributor_id === id).reduce((a, x) => a + x.net, 0) / MONTHS.length;
    const facilities = [];
    const snapshots = [];
    DISTRIBUTORS.forEach((d) => {
      const types = R() < 0.35 ? ['Working Capital', 'Term Loan'] : ['Working Capital'];
      types.forEach((type) => {
        const facility_id = 'F' + String(facilities.length + 1).padStart(3, '0');
        const limit = round(monthlyNet(d.id) * (type === 'Term Loan' ? 1.2 : 2.5), 50000000) || 50000000;
        let out = limit * (0.5 + R() * 0.3);
        let last = null;
        MONTHS.forEach((m, mi) => {
          const dpd = d.trouble ? Math.max(0, LATE[mi] + int(-3, 3)) : (R() < 0.08 ? int(1, 14) : 0);
          out = Math.min(limit, out * (d.trouble ? 1.05 : 0.97 + R() * 0.05));
          last = { month: m, month_end: lastDay(m), facility_id, distributor_id: d.id, outstanding: round(out, 100000), dpd, kol: kol(dpd) };
          snapshots.push(last);
        });
        facilities.push({
          facility_id, distributor_id: d.id, distributor: d.name, region: d.region, type,
          limit, outstanding: last.outstanding, rate_pct: 9 + int(0, 8) * 0.25,
          start_date: addDays('2024-01-01', int(0, 600)), dpd: last.dpd, kol: last.kol
        });
      });
    });

    cache = { products: PRODUCTS, distributors: DISTRIBUTORS, outlets, sales, facilities, snapshots };
    return cache;
  }

  /** The 20-row sheet for the sample SUMIFS mission: October sales for five distributors. */
  function missionSales() {
    const { sales } = build();
    const ids = ['D01', 'D02', 'D03', 'D04', 'D05'];
    const rows = sales.filter((x) => x.date.startsWith('2026-10') && ids.includes(x.distributor_id)).slice(0, 20);
    const count = {};
    rows.forEach((r) => { count[r.distributor] = (count[r.distributor] || 0) + 1; });
    const ranked = Object.entries(count).sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1)).map((e) => e[0]);
    return { rows, target: ranked[0], other: ranked[1], example: ranked[2] };
  }

  /** The bank's short list for the macro demo: one line per facility, first six. */
  function macroSheet() {
    const { facilities } = build();
    const wc = facilities.filter((f) => f.type === 'Working Capital');
    const chosen = [wc[0], wc[4], wc[1], wc[7], wc[2], wc[11]];
    return chosen.map((f) => ({ distributor: f.distributor, outstanding: f.outstanding, dpd: f.dpd }));
  }

  /** The joined view the course ends on: sales falling and days past due rising, per distributor. */
  function earlyWarning() {
    const { sales, snapshots } = build();
    return DISTRIBUTORS.map((d) => {
      const net = (m) => sales.filter((x) => x.distributor_id === d.id && x.date.startsWith(m)).reduce((a, x) => a + x.net, 0);
      const dpd = (m) => Math.max(...snapshots.filter((x) => x.distributor_id === d.id && x.month === m).map((x) => x.dpd));
      return { id: d.id, distributor: d.name, net_may: net('2026-05'), net_oct: net('2026-10'), dpd_may: dpd('2026-05'), dpd_oct: dpd('2026-10'), trouble: d.trouble };
    });
  }

  return { AS_OF, MONTHS, PRODUCTS, DISTRIBUTORS, CAST, build, missionSales, macroSheet, earlyWarning, toSerial, kol };
});
