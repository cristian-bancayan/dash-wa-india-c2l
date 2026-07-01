/* =====================================================================
   WhatsApp Delivery Monitor — engageSPARK
   Reads the weekly engageSPARK export (data/latest.xlsx) fully in the
   browser (SheetJS) and renders scorecards + charts (Chart.js).

   HOW TO UPDATE WEEKLY
   ---------------------
   1. Download the new report from engageSPARK.
   2. Rename it to exactly:  latest.xlsx
   3. Replace the file at:   data/latest.xlsx
   4. Commit + push. The dashboard will pick it up automatically.

   TUNABLE CONSTANTS
   ------------------
   Everything you're likely to need to adjust as the campaign evolves
   lives in the CONFIG block right below.
===================================================================== */

const CONFIG = {
  // Sheet names expected inside the workbook
  SHEET_CAMPAIGN: 'Campaign Report',
  SHEET_WHATSAPP: 'WhatsApp Log',

  // Text (lowercased, trimmed) that counts as a "Saturday prompt button"
  // click. engageSPARK/WhatsApp quick-reply buttons sometimes get
  // re-labeled over time — add new variants here if the click chart
  // looks like it's under-counting.
  BUTTON_REPLY_TEXTS: [
    'send my chat prompt',
    'send my prompt',
    'send prompt',
    "i'm ready to chat",
    'ready to chat',
  ],

  // A message is treated as part of the "Saturday" cadence (used to
  // compute the % of Saturday sends that got a button click) if its
  // Message Label contains this substring.
  SATURDAY_LABEL_HINT: 'Sat',

  // Enrollment rounds are auto-detected by clustering distinct
  // subscription dates: a gap larger than this many days starts a new
  // round.
  ROUND_GAP_DAYS: 3,
};

/* ===== Global state ===== */
const STATE = {
  campaignRows: [],      // [{contactId, subscriptionTime, status, round}]
  waRows: [],             // [{contactId, direction, label, time, status, error, week, isButtonClick, isSaturday}]
  rounds: [],              // [{id, label, start, end, contactIds:Set}]
  contactMeta: new Map(),  // contactId -> {firstName, round}
  weeks: [],                // [{n, label, start, end}]
  anchorDate: null,
  filters: {
    rounds: new Set(['all']),
    participants: new Set(),
    weeks: new Set(['all']),
  },
  chartModes: { delivery: 'count', blocked: 'count', button: 'count' },
  charts: {},
};

/* ===== Boot ===== */
document.addEventListener('DOMContentLoaded', () => {
  loadFromFetch();
  document.getElementById('fileInput').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (ev) => processWorkbook(ev.target.result);
      reader.readAsArrayBuffer(file);
    }
  });
});

function loadFromFetch() {
  fetch('data/latest.xlsx')
    .then((res) => {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.arrayBuffer();
    })
    .then((buf) => processWorkbook(buf))
    .catch(() => {
      document.getElementById('loadingState').classList.add('hidden');
      document.getElementById('errorState').classList.remove('hidden');
      setStatus('err', 'No report loaded');
    });
}

function setStatus(kind, text) {
  const dot = document.getElementById('statusDot');
  dot.className = 'brand-dot ' + kind;
  document.getElementById('reportMetaValue').textContent = text;
}

/* =====================================================================
   PARSING
===================================================================== */
function processWorkbook(arrayBuffer) {
  let wb;
  try {
    wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  } catch (e) {
    console.error(e);
    document.getElementById('loadingState').classList.add('hidden');
    document.getElementById('errorState').classList.remove('hidden');
    setStatus('err', 'Could not parse file');
    return;
  }

  const campSheet = wb.Sheets[CONFIG.SHEET_CAMPAIGN];
  const waSheet = wb.Sheets[CONFIG.SHEET_WHATSAPP];
  if (!campSheet || !waSheet) {
    document.getElementById('loadingState').classList.add('hidden');
    document.getElementById('errorState').classList.remove('hidden');
    document.querySelector('#errorState p').innerHTML =
      `<strong>Expected sheets not found.</strong> This dashboard needs "${CONFIG.SHEET_CAMPAIGN}" and "${CONFIG.SHEET_WHATSAPP}" tabs in the workbook.`;
    setStatus('err', 'Unexpected file format');
    return;
  }

  const campJson = XLSX.utils.sheet_to_json(campSheet, { defval: null });
  const waJson = XLSX.utils.sheet_to_json(waSheet, { defval: null });

  buildCampaignData(campJson);
  buildWhatsAppData(waJson);
  buildRounds();
  attachRoundsToMessages();
  buildWeeks();

  renderAll();

  document.getElementById('loadingState').classList.add('hidden');
  document.getElementById('errorState').classList.add('hidden');
  document.getElementById('dashboard').classList.remove('hidden');

  const lastMsg = STATE.waRows.reduce((max, r) => (r.time && (!max || r.time > max) ? r.time : max), null);
  setStatus('ok', lastMsg ? `Data through ${fmtDate(lastMsg)}` : 'Loaded');
}

function buildCampaignData(rows) {
  STATE.campaignRows = rows.map((r) => ({
    contactId: r['Contact ID'],
    firstName: r['First Name'] || null,
    subscriptionTime: toDate(r['Subscription Time (America/Chicago)']),
    status: r['Subscription Status'] || null,
  })).filter((r) => r.contactId != null);
}

function buildWhatsAppData(rows) {
  STATE.waRows = rows.map((r) => {
    const time = toDate(r['Time of Message (America/Chicago)']);
    const direction = (r['Direction'] || '').toLowerCase();
    const label = r['Message Label'] || '';
    const msgText = (r['Message'] || '').toString().trim().toLowerCase();
    return {
      contactId: r['Contact ID'],
      direction,
      label,
      time,
      status: r['Delivery Status'] || 'Unknown',
      error: r['Delivery Error'] || null,
      isButtonClick: direction === 'inbound' && CONFIG.BUTTON_REPLY_TEXTS.includes(msgText),
      isSaturdaySend: direction === 'outbound' && label.indexOf(CONFIG.SATURDAY_LABEL_HINT) === 0,
    };
  }).filter((r) => r.contactId != null && r.time != null);

  // anchor = earliest message timestamp in the whole log
  STATE.anchorDate = STATE.waRows.reduce(
    (min, r) => (!min || r.time < min ? r.time : min),
    null
  );

  STATE.waRows.forEach((r) => {
    r.week = weekNumberOf(r.time);
  });
}

function weekNumberOf(date) {
  if (!STATE.anchorDate || !date) return null;
  const diffDays = Math.floor((date - STATE.anchorDate) / (1000 * 60 * 60 * 24));
  return Math.floor(diffDays / 7) + 1;
}

/* ===== Enrollment round detection ===== */
function buildRounds() {
  const withDates = STATE.campaignRows.filter((r) => r.subscriptionTime);
  const dayKey = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const uniqueDays = [...new Set(withDates.map((r) => dayKey(r.subscriptionTime)))].sort((a, b) => a - b);

  const clusters = [];
  let current = [];
  uniqueDays.forEach((day, i) => {
    if (current.length === 0) {
      current.push(day);
    } else {
      const gapDays = (day - current[current.length - 1]) / (1000 * 60 * 60 * 24);
      if (gapDays > CONFIG.ROUND_GAP_DAYS) {
        clusters.push(current);
        current = [day];
      } else {
        current.push(day);
      }
    }
  });
  if (current.length) clusters.push(current);

  STATE.rounds = clusters.map((days, i) => {
    const start = new Date(Math.min(...days));
    const end = new Date(Math.max(...days));
    const contactIds = new Set(
      withDates.filter((r) => dayKey(r.subscriptionTime) >= days[0] && dayKey(r.subscriptionTime) <= days[days.length - 1]).map((r) => r.contactId)
    );
    return { id: 'r' + (i + 1), n: i + 1, start, end, contactIds };
  });

  // contact -> round lookup
  STATE.contactMeta = new Map();
  STATE.campaignRows.forEach((r) => {
    const round = STATE.rounds.find((rd) => r.subscriptionTime && rd.contactIds.has(r.contactId));
    STATE.contactMeta.set(r.contactId, {
      firstName: r.firstName,
      round: round ? round.id : null,
      status: r.status,
    });
  });
}

function attachRoundsToMessages() {
  STATE.waRows.forEach((r) => {
    const meta = STATE.contactMeta.get(r.contactId);
    r.round = meta ? meta.round : null;
  });
}

/* ===== Week list (for chips + x-axis labels) ===== */
function buildWeeks() {
  const nums = [...new Set(STATE.waRows.map((r) => r.week).filter((w) => w != null))].sort((a, b) => a - b);
  STATE.weeks = nums.map((n) => {
    const start = new Date(STATE.anchorDate.getTime() + (n - 1) * 7 * 86400000);
    const end = new Date(start.getTime() + 6 * 86400000);
    return { n, label: `Week ${n}`, sub: `${fmtDate(start)}–${fmtDate(end)}` };
  });
}

/* =====================================================================
   FILTERING HELPERS
===================================================================== */
function activeRoundIds() {
  if (STATE.filters.rounds.has('all') || STATE.filters.rounds.size === 0) return null; // null = no round restriction
  return [...STATE.filters.rounds];
}
function activeWeeks() {
  if (STATE.filters.weeks.has('all') || STATE.filters.weeks.size === 0) return null;
  return new Set([...STATE.filters.weeks].map(Number));
}
function activeParticipants() {
  return STATE.filters.participants.size ? STATE.filters.participants : null;
}

// Returns filtered rows respecting round + participant + week filters.
// `roundOverride` lets the delivery chart compute one subset per round
// for comparison mode.
function filteredRows(roundOverride) {
  const parts = activeParticipants();
  const weeks = activeWeeks();
  const rounds = roundOverride ? [roundOverride] : activeRoundIds();

  return STATE.waRows.filter((r) => {
    if (rounds && !rounds.includes(r.round)) return false;
    if (parts && !parts.has(r.contactId)) return false;
    if (weeks && !weeks.has(r.week)) return false;
    return true;
  });
}

/* =====================================================================
   RENDER: everything
===================================================================== */
function renderAll() {
  renderScorecards();
  renderFilterControls();
  renderDeliveryChart();
  renderBlockedChart();
  renderButtonChart();
}

/* ===== Scorecards (always overall totals, not affected by filters) ===== */
function renderScorecards() {
  const uniqueContacts = new Set(STATE.campaignRows.map((r) => r.contactId));
  document.getElementById('scoreEnrolled').textContent = uniqueContacts.size.toLocaleString();
  document.getElementById('scoreEnrolledFoot').textContent = `across ${STATE.rounds.length} enrollment round${STATE.rounds.length === 1 ? '' : 's'}`;

  const sent = STATE.waRows.filter((r) => r.direction === 'outbound').length;
  document.getElementById('scoreSent').textContent = sent.toLocaleString();
  document.getElementById('scoreSentFoot').textContent = `through ${STATE.weeks.length ? STATE.weeks[STATE.weeks.length - 1].label : '—'}`;

  const responded = new Set(STATE.waRows.filter((r) => r.direction === 'inbound').map((r) => r.contactId));
  document.getElementById('scoreResponded').textContent = responded.size.toLocaleString();
  const pct = uniqueContacts.size ? Math.round((responded.size / uniqueContacts.size) * 100) : 0;
  document.getElementById('scoreRespondedFoot').textContent = `${pct}% of enrolled participants`;
}

/* ===== Filter controls ===== */
function renderFilterControls() {
  // Round chips
  const roundWrap = document.getElementById('roundChips');
  roundWrap.innerHTML = '';
  roundWrap.appendChild(makeChip('All', 'all', STATE.filters.rounds.has('all'), true));
  STATE.rounds.forEach((rd) => {
    const label = `${rd.label ? rd.label : 'Round ' + rd.n} (${fmtDate(rd.start)}${rd.start.getTime() !== rd.end.getTime() ? '–' + fmtDate(rd.end) : ''}) · ${rd.contactIds.size}`;
    roundWrap.appendChild(makeChip(label, rd.id, STATE.filters.rounds.has(rd.id), false));
  });
  roundWrap.querySelectorAll('.chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const val = chip.dataset.value;
      if (val === 'all') {
        STATE.filters.rounds = new Set(['all']);
      } else {
        STATE.filters.rounds.delete('all');
        if (STATE.filters.rounds.has(val)) {
          STATE.filters.rounds.delete(val);
          if (STATE.filters.rounds.size === 0) STATE.filters.rounds.add('all');
        } else {
          STATE.filters.rounds.add(val);
        }
      }
      renderFilterControls();
      renderDeliveryChart();
      renderBlockedChart();
      renderButtonChart();
    });
  });

  // Week chips
  const weekWrap = document.getElementById('weekChips');
  weekWrap.innerHTML = '';
  weekWrap.appendChild(makeChip('All weeks', 'all', STATE.filters.weeks.has('all'), true));
  STATE.weeks.forEach((w) => {
    weekWrap.appendChild(makeChip(w.label, String(w.n), STATE.filters.weeks.has(String(w.n)), false));
  });
  weekWrap.querySelectorAll('.chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const val = chip.dataset.value;
      if (val === 'all') {
        STATE.filters.weeks = new Set(['all']);
      } else {
        STATE.filters.weeks.delete('all');
        if (STATE.filters.weeks.has(val)) {
          STATE.filters.weeks.delete(val);
          if (STATE.filters.weeks.size === 0) STATE.filters.weeks.add('all');
        } else {
          STATE.filters.weeks.add(val);
        }
      }
      renderFilterControls();
      renderDeliveryChart();
      renderBlockedChart();
      renderButtonChart();
    });
  });

  renderParticipantSelect();

  document.getElementById('resetFilters').onclick = () => {
    STATE.filters.rounds = new Set(['all']);
    STATE.filters.weeks = new Set(['all']);
    STATE.filters.participants = new Set();
    renderFilterControls();
    renderDeliveryChart();
    renderBlockedChart();
    renderButtonChart();
  };
}

function makeChip(label, value, active, isAll) {
  const el = document.createElement('button');
  el.className = 'chip' + (isAll ? ' all' : '') + (active ? ' active' : '');
  el.dataset.value = value;
  el.type = 'button';
  el.textContent = label;
  return el;
}

function renderParticipantSelect() {
  const searchInput = document.getElementById('participantSearch');
  const listEl = document.getElementById('participantList');
  const tagsEl = document.getElementById('participantTags');

  const all = STATE.campaignRows
    .filter((r) => r.contactId != null)
    .map((r) => ({ id: r.contactId, name: r.firstName || 'Unnamed participant' }))
    .sort((a, b) => a.name.localeCompare(b.name));

  function renderList(query) {
    const q = (query || '').trim().toLowerCase();
    const matches = all.filter(
      (p) => !q || p.name.toLowerCase().includes(q) || String(p.id).includes(q)
    ).slice(0, 60);
    listEl.innerHTML = '';
    matches.forEach((p) => {
      const item = document.createElement('div');
      item.className = 'searchable-item' + (STATE.filters.participants.has(p.id) ? ' picked' : '');
      item.innerHTML = `<span>${escapeHtml(p.name)}</span><span class="sid">#${p.id}</span>`;
      item.addEventListener('click', () => {
        if (STATE.filters.participants.has(p.id)) {
          STATE.filters.participants.delete(p.id);
        } else {
          STATE.filters.participants.add(p.id);
        }
        renderParticipantSelect();
        renderDeliveryChart();
        renderBlockedChart();
        renderButtonChart();
      });
      listEl.appendChild(item);
    });
  }
  renderList(searchInput.value);

  searchInput.oninput = () => {
    listEl.classList.add('open');
    renderList(searchInput.value);
  };
  searchInput.onfocus = () => listEl.classList.add('open');
  document.addEventListener('click', (e) => {
    if (!document.getElementById('participantSelect').contains(e.target)) {
      listEl.classList.remove('open');
    }
  });

  tagsEl.innerHTML = '';
  [...STATE.filters.participants].forEach((id) => {
    const meta = all.find((p) => p.id === id);
    const tag = document.createElement('span');
    tag.className = 'tag';
    tag.innerHTML = `${escapeHtml(meta ? meta.name : 'Unknown')} #${id} <button type="button">×</button>`;
    tag.querySelector('button').addEventListener('click', () => {
      STATE.filters.participants.delete(id);
      renderParticipantSelect();
      renderDeliveryChart();
      renderBlockedChart();
      renderButtonChart();
    });
    tagsEl.appendChild(tag);
  });
}

/* =====================================================================
   CHART 1 — Delivery status
===================================================================== */
function renderDeliveryChart() {
  const ctx = document.getElementById('deliveryChart');
  const mode = STATE.chartModes.delivery;
  const compareRounds = !STATE.filters.rounds.has('all') && STATE.filters.rounds.size > 0;
  const categories = ['Read', 'Delivered', 'Not Delivered', 'Unknown'];
  const categoryLabels = ['Read', 'Delivered only', 'Not Delivered', 'Unknown'];
  const colors = [CSS_VAR('--teal'), CSS_VAR('--indigo'), CSS_VAR('--red'), CSS_VAR('--muted')];

  let datasets, labels;

  function countFor(rows) {
    const outbound = rows.filter((r) => r.direction === 'outbound');
    const total = outbound.length || 1;
    return categories.map((cat) => {
      const n = outbound.filter((r) => r.status === cat).length;
      return mode === 'percent' ? +((n / total) * 100).toFixed(1) : n;
    });
  }

  if (compareRounds) {
    const roundIds = [...STATE.filters.rounds];
    labels = categoryLabels;
    datasets = roundIds.map((rid, i) => {
      const rd = STATE.rounds.find((r) => r.id === rid);
      const rows = filteredRows(rid);
      return {
        label: rd ? `Round ${rd.n}` : rid,
        data: countFor(rows),
        backgroundColor: PALETTE[i % PALETTE.length],
        borderRadius: 4,
      };
    });
  } else {
    labels = categoryLabels;
    const rows = filteredRows();
    datasets = [{
      label: 'Messages',
      data: countFor(rows),
      backgroundColor: colors,
      borderRadius: 4,
    }];
  }

  drawBarChart('deliveryChart', 'delivery', labels, datasets, mode === 'percent');
}

/* =====================================================================
   CHART 2 — Blocked by Meta (x axis = week)
===================================================================== */
function renderBlockedChart() {
  const mode = STATE.chartModes.blocked;
  const rows = filteredRows();
  const weeksToShow = activeWeeks() ? STATE.weeks.filter((w) => activeWeeks().has(w.n)) : STATE.weeks;

  const data = weeksToShow.map((w) => {
    const weekOutbound = rows.filter((r) => r.direction === 'outbound' && r.week === w.n);
    const blocked = weekOutbound.filter((r) => r.error === 'Blocked by Meta').length;
    if (mode === 'percent') {
      const total = weekOutbound.length || 1;
      return +((blocked / total) * 100).toFixed(1);
    }
    return blocked;
  });

  drawBarChart(
    'blockedChart', 'blocked',
    weeksToShow.map((w) => w.label),
    [{ label: mode === 'percent' ? '% blocked by Meta' : 'Blocked by Meta', data, backgroundColor: CSS_VAR('--red'), borderRadius: 4 }],
    mode === 'percent'
  );
}

/* =====================================================================
   CHART 3 — Saturday button clicks (x axis = week)
===================================================================== */
function renderButtonChart() {
  const mode = STATE.chartModes.button;
  const rows = filteredRows();
  const weeksToShow = activeWeeks() ? STATE.weeks.filter((w) => activeWeeks().has(w.n)) : STATE.weeks;

  const data = weeksToShow.map((w) => {
    const clicks = rows.filter((r) => r.isButtonClick && r.week === w.n).length;
    if (mode === 'percent') {
      const satSent = rows.filter((r) => r.isSaturdaySend && r.week === w.n).length || 1;
      return +((clicks / satSent) * 100).toFixed(1);
    }
    return clicks;
  });

  drawBarChart(
    'buttonChart', 'button',
    weeksToShow.map((w) => w.label),
    [{ label: mode === 'percent' ? '% of Sat. sends clicked' : 'Button clicks', data, backgroundColor: CSS_VAR('--amber'), borderRadius: 4 }],
    mode === 'percent'
  );
}

/* =====================================================================
   Chart.js helper
===================================================================== */
function drawBarChart(canvasId, key, labels, datasets, isPercent) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  if (STATE.charts[key]) STATE.charts[key].destroy();
  STATE.charts[key] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: datasets.length > 1, labels: { font: { family: 'Inter', size: 11 } } },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y}${isPercent ? '%' : ''}`,
          },
        },
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 11 } } },
        y: {
          beginAtZero: true,
          ticks: {
            font: { family: 'Inter', size: 11 },
            callback: (v) => v + (isPercent ? '%' : ''),
          },
          grid: { color: '#EEEAE0' },
        },
      },
    },
  });
}

/* ===== Toggle buttons (count / percent) ===== */
document.addEventListener('click', (e) => {
  const btn = e.target.closest('.toggle-btn');
  if (!btn) return;
  const group = btn.closest('.toggle-group');
  const key = group.dataset.toggle;
  group.querySelectorAll('.toggle-btn').forEach((b) => b.classList.remove('active'));
  btn.classList.add('active');
  STATE.chartModes[key] = btn.dataset.mode;
  if (key === 'delivery') renderDeliveryChart();
  if (key === 'blocked') renderBlockedChart();
  if (key === 'button') renderButtonChart();
});

/* =====================================================================
   Small utilities
===================================================================== */
// Chat2Learn brand palette (deep green, blue, lime, plus a couple of
// darker fallbacks in case more than 4 rounds are compared at once)
const PALETTE = ['#00A651', '#29ABE2', '#8DC63F', '#AC3C33', '#0E7A3D', '#1D7FA8'];

function CSS_VAR(name) {
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function fmtDate(d) {
  if (!d) return '—';
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
