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

   PRIVACY
   --------
   This dashboard never reads or displays the "First Name" column. The
   only per-person identifier used anywhere is Contact ID.

   WEEK NUMBERING
   ---------------
   "Week N" is always counted from each participant's OWN subscription
   date (Week 0 = their first 7 days), not a shared calendar date. This
   keeps weeks comparable across enrollment rounds that started on
   different calendar dates. Because of this, week numbers intentionally
   have no fixed date range shown next to them.

   TUNABLE CONSTANTS
   ------------------
===================================================================== */

const CONFIG = {
  SHEET_CAMPAIGN: 'Campaign Report',
  SHEET_WHATSAPP: 'WhatsApp Log',

  // Text (lowercased, trimmed) that counts as a "Saturday prompt button"
  // click. Add new variants here if the click chart looks like it's
  // under-counting.
  BUTTON_REPLY_TEXTS: [
    'send my chat prompt',
    'send my prompt',
    'send prompt',
    "i'm ready to chat",
    'ready to chat',
  ],

  // A message is treated as part of the "Saturday" cadence (used to
  // compute the % of Saturday sends that got a button click) if its
  // Message Label starts with this substring.
  SATURDAY_LABEL_HINT: 'Sat',

  // Enrollment rounds are auto-detected by clustering distinct
  // subscription dates: a gap larger than this many days starts a new
  // round.
  ROUND_GAP_DAYS: 3,
};

/* ===== Global state ===== */
const STATE = {
  campaignRows: [],      // [{contactId, subscriptionTime}]
  waRows: [],              // [{contactId, direction, label, time, status, error, week, message, isButtonClick, isSaturdaySend, round}]
  rounds: [],               // [{id, n, start, end, contactIds:Set}]
  contactMeta: new Map(),    // contactId -> {round, subscriptionTime}
  weeks: [],                   // [{n, label}]  distinct per-participant-relative week numbers observed
  slots: [],                     // [{label, order, sampleMessage, display, weekNum}]  distinct outbound message templates, chronological
  anchorDate: null,
  filters: {
    rounds: new Set(['all']),
    participants: new Set(),   // Set<contactId>
    weeks: new Set(['all']),
  },
  chartModes: { delivery: 'count', blocked: 'count', button: 'count', slot: 'count' },
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

  document.querySelectorAll('.tab-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach((b) => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach((p) => p.classList.add('hidden'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.remove('hidden');
    });
  });
});

function loadFromFetch() {
  fetch('data/latest.xlsx')
    .then((res) => {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.arrayBuffer();
    })
    .then((buf) => processWorkbook(buf))
    .catch((err) => {
      console.error(err);
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
  buildRounds();
  buildWhatsAppData(waJson); // needs contactMeta (built in buildRounds) for per-contact weeks
  buildWeeks();
  buildSlots();

  renderAll();

  document.getElementById('loadingState').classList.add('hidden');
  document.getElementById('errorState').classList.add('hidden');
  document.getElementById('dashboard').classList.remove('hidden');

  const lastMsg = STATE.waRows.reduce((max, r) => (r.time && (!max || r.time > max) ? r.time : max), null);
  setStatus('ok', lastMsg ? `Data through ${fmtDate(lastMsg)}` : 'Loaded');
}

// NOTE: First Name is intentionally never read from the workbook —
// Contact ID is the only identifier this dashboard touches.
function buildCampaignData(rows) {
  STATE.campaignRows = rows.map((r) => ({
    contactId: r['Contact ID'],
    subscriptionTime: toDate(r['Subscription Time (America/Chicago)']),
  })).filter((r) => r.contactId != null);
}

function buildWhatsAppData(rows) {
  STATE.waRows = rows.map((r) => {
    const time = toDate(r['Time of Message (America/Chicago)']);
    const direction = (r['Direction'] || '').toLowerCase();
    const label = r['Message Label'] || '';
    const message = (r['Message'] || '').toString().trim();
    const msgTextLower = message.toLowerCase();
    const contactId = r['Contact ID'];
    const meta = STATE.contactMeta.get(contactId);
    return {
      contactId,
      direction,
      label,
      time,
      status: r['Delivery Status'] || 'Unknown',
      error: r['Delivery Error'] || null,
      message,
      round: meta ? meta.round : null,
      week: meta && meta.subscriptionTime && time ? weekSinceSubscription(meta.subscriptionTime, time) : null,
      isButtonClick: direction === 'inbound' && CONFIG.BUTTON_REPLY_TEXTS.includes(msgTextLower),
      isSaturdaySend: direction === 'outbound' && label.indexOf(CONFIG.SATURDAY_LABEL_HINT) === 0,
    };
  }).filter((r) => r.contactId != null && r.time != null);

  STATE.anchorDate = STATE.waRows.reduce(
    (min, r) => (!min || r.time < min ? r.time : min),
    null
  );
}

// Week 0 = the participant's first 7 days after subscribing.
function weekSinceSubscription(subscriptionTime, msgTime) {
  const diffDays = Math.floor((msgTime - subscriptionTime) / 86400000);
  return Math.floor(diffDays / 7);
}

/* ===== Enrollment round detection + contact -> round/subscription lookup ===== */
function buildRounds() {
  const withDates = STATE.campaignRows.filter((r) => r.subscriptionTime);
  const dayKey = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const uniqueDays = [...new Set(withDates.map((r) => dayKey(r.subscriptionTime)))].sort((a, b) => a - b);

  const clusters = [];
  let current = [];
  uniqueDays.forEach((day) => {
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

  STATE.contactMeta = new Map();
  STATE.campaignRows.forEach((r) => {
    const round = STATE.rounds.find((rd) => r.subscriptionTime && rd.contactIds.has(r.contactId));
    STATE.contactMeta.set(r.contactId, {
      round: round ? round.id : null,
      subscriptionTime: r.subscriptionTime,
    });
  });
}

/* ===== Week list (for chart x-axes and chips) — participant-relative, no dates ===== */
function buildWeeks() {
  const nums = [...new Set(STATE.waRows.map((r) => r.week).filter((w) => w != null))].sort((a, b) => a - b);
  STATE.weeks = nums.map((n) => ({ n, label: `Week ${n}` }));
}

/* ===== Distinct outbound message slots (templates), in chronological order ===== */
const WEEKDAY_SHORT = { Tues: 'Tue', Weds: 'Wed', Thurs: 'Thu', Sat: 'Sat', Sun: 'Sun', Mon: 'Mon', Fri: 'Fri' };

function parseSlotLabel(label) {
  const m = label.match(/^([A-Za-z]+)\s+Week\s+(\d+)$/i);
  if (m) {
    const day = WEEKDAY_SHORT[m[1]] || m[1].slice(0, 3);
    return { display: `Week ${m[2]} ${day}`, weekNum: parseInt(m[2], 10) };
  }
  if (!label) return { display: 'Unlabeled', weekNum: null };
  return { display: label.charAt(0).toUpperCase() + label.slice(1), weekNum: null };
}

function buildSlots() {
  const byLabel = new Map();
  STATE.waRows
    .filter((r) => r.direction === 'outbound' && r.label)
    .forEach((r) => {
      if (!byLabel.has(r.label)) {
        byLabel.set(r.label, { label: r.label, order: r.time, sampleMessage: r.message || '' });
      } else {
        const entry = byLabel.get(r.label);
        if (r.time < entry.order) entry.order = r.time;
        if (!entry.sampleMessage && r.message) entry.sampleMessage = r.message;
      }
    });
  STATE.slots = [...byLabel.values()]
    .sort((a, b) => a.order - b.order)
    .map((s) => {
      const parsed = parseSlotLabel(s.label);
      return { ...s, display: parsed.display, weekNum: parsed.weekNum };
    });
}

/* =====================================================================
   FILTERING
===================================================================== */
function activeRoundIds() {
  if (STATE.filters.rounds.has('all') || STATE.filters.rounds.size === 0) return null;
  return [...STATE.filters.rounds];
}
function activeParticipants() {
  return STATE.filters.participants.size ? STATE.filters.participants : null;
}
function activeWeeks() {
  if (STATE.filters.weeks.has('all') || STATE.filters.weeks.size === 0) return null;
  return new Set([...STATE.filters.weeks].map(Number));
}

// Rows filtered by round + contact ID only (week filtering is applied
// separately, per chart, since weeks control which x-axis categories
// are drawn rather than which underlying rows are counted).
function filteredRows() {
  const rounds = activeRoundIds();
  const parts = activeParticipants();
  return STATE.waRows.filter((r) => {
    if (rounds && !rounds.includes(r.round)) return false;
    if (parts && !parts.has(r.contactId)) return false;
    return true;
  });
}

function weeksToShow() {
  const active = activeWeeks();
  return active ? STATE.weeks.filter((w) => active.has(w.n)) : STATE.weeks;
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
  renderSlotChart();
  renderScheduledMessages();
}

function refreshFilteredViews() {
  renderDeliveryChart();
  renderBlockedChart();
  renderButtonChart();
  renderSlotChart();
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
    const label = `Round ${rd.n} (${fmtDate(rd.start)}${rd.start.getTime() !== rd.end.getTime() ? '–' + fmtDate(rd.end) : ''}) · ${rd.contactIds.size}`;
    roundWrap.appendChild(makeChip(label, rd.id, STATE.filters.rounds.has(rd.id), false));
  });
  roundWrap.querySelectorAll('.chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      toggleChipFilter(STATE.filters.rounds, chip.dataset.value);
      renderFilterControls();
      refreshFilteredViews();
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
      toggleChipFilter(STATE.filters.weeks, chip.dataset.value);
      renderFilterControls();
      refreshFilteredViews();
    });
  });

  renderParticipantSelect();

  document.getElementById('resetFilters').onclick = () => {
    STATE.filters.rounds = new Set(['all']);
    STATE.filters.weeks = new Set(['all']);
    STATE.filters.participants = new Set();
    renderFilterControls();
    refreshFilteredViews();
  };
}

function toggleChipFilter(set, val) {
  if (val === 'all') {
    set.clear();
    set.add('all');
    return;
  }
  set.delete('all');
  if (set.has(val)) {
    set.delete(val);
    if (set.size === 0) set.add('all');
  } else {
    set.add(val);
  }
}

function makeChip(label, value, active, isAll) {
  const el = document.createElement('button');
  el.className = 'chip' + (isAll ? ' all' : '') + (active ? ' active' : '');
  el.dataset.value = value;
  el.type = 'button';
  el.textContent = label;
  return el;
}

/* ===== Contact ID filter ===== */
function renderParticipantSelect() {
  const searchInput = document.getElementById('participantSearch');
  const listEl = document.getElementById('participantList');
  const tagsEl = document.getElementById('participantTags');

  const allIds = [...new Set(STATE.campaignRows.map((r) => r.contactId))].sort((a, b) => a - b);

  function renderList(query) {
    const q = (query || '').trim();
    const matches = (q ? allIds.filter((id) => String(id).includes(q)) : allIds).slice(0, 60);
    listEl.innerHTML = '';
    matches.forEach((id) => {
      const item = document.createElement('div');
      item.className = 'searchable-item' + (STATE.filters.participants.has(id) ? ' picked' : '');
      item.innerHTML = `<span>Contact #${id}</span>`;
      item.addEventListener('click', () => {
        if (STATE.filters.participants.has(id)) {
          STATE.filters.participants.delete(id);
        } else {
          STATE.filters.participants.add(id);
        }
        renderParticipantSelect();
        refreshFilteredViews();
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
  [...STATE.filters.participants].sort((a, b) => a - b).forEach((id) => {
    const tag = document.createElement('span');
    tag.className = 'tag';
    tag.innerHTML = `Contact #${id} <button type="button">×</button>`;
    tag.querySelector('button').addEventListener('click', () => {
      STATE.filters.participants.delete(id);
      renderParticipantSelect();
      refreshFilteredViews();
    });
    tagsEl.appendChild(tag);
  });
}

/* =====================================================================
   CHART 1 — Delivery status, by week since enrollment (stacked)
===================================================================== */
function renderDeliveryChart() {
  const mode = STATE.chartModes.delivery;
  const rows = filteredRows();
  const weeks = weeksToShow();
  const categories = ['Read', 'Delivered', 'Not Delivered', 'Unknown'];
  const categoryLabels = ['Read', 'Delivered only', 'Not Delivered', 'Unknown'];
  const colors = [CSS_VAR('--teal'), CSS_VAR('--indigo'), CSS_VAR('--red'), CSS_VAR('--muted')];

  const datasets = categories.map((cat, i) => ({
    label: categoryLabels[i],
    backgroundColor: colors[i],
    borderRadius: 3,
    stack: 'delivery',
    data: weeks.map((w) => {
      const weekOutbound = rows.filter((r) => r.direction === 'outbound' && r.week === w.n);
      const n = weekOutbound.filter((r) => r.status === cat).length;
      if (mode === 'percent') {
        const total = weekOutbound.length || 1;
        return +((n / total) * 100).toFixed(1);
      }
      return n;
    }),
  }));

  drawBarChart('deliveryChart', 'delivery', weeks.map((w) => w.label), datasets, mode === 'percent', true);
}

/* =====================================================================
   CHART 2 — Blocked by Meta (x axis = week since enrollment)
===================================================================== */
function renderBlockedChart() {
  const mode = STATE.chartModes.blocked;
  const rows = filteredRows();
  const weeks = weeksToShow();

  const data = weeks.map((w) => {
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
    weeks.map((w) => w.label),
    [{ label: mode === 'percent' ? '% blocked by Meta' : 'Blocked by Meta', data, backgroundColor: CSS_VAR('--red'), borderRadius: 4 }],
    mode === 'percent'
  );
}

/* =====================================================================
   CHART 3 — Saturday button clicks (x axis = week since enrollment)
===================================================================== */
function renderButtonChart() {
  const mode = STATE.chartModes.button;
  const rows = filteredRows();
  const weeks = weeksToShow();

  const data = weeks.map((w) => {
    const clicks = rows.filter((r) => r.isButtonClick && r.week === w.n).length;
    if (mode === 'percent') {
      const satSent = rows.filter((r) => r.isSaturdaySend && r.week === w.n).length || 1;
      return +((clicks / satSent) * 100).toFixed(1);
    }
    return clicks;
  });

  drawBarChart(
    'buttonChart', 'button',
    weeks.map((w) => w.label),
    [{ label: mode === 'percent' ? '% of Sat. sends clicked' : 'Button clicks', data, backgroundColor: CSS_VAR('--amber'), borderRadius: 4 }],
    mode === 'percent'
  );
}

/* =====================================================================
   CHART 4 — Delivery status per individual scheduled message
===================================================================== */
function renderSlotChart() {
  const mode = STATE.chartModes.slot;
  const rows = filteredRows();
  const categories = ['Read', 'Delivered', 'Not Delivered', 'Unknown'];
  const categoryLabels = ['Read', 'Delivered only', 'Not Delivered', 'Unknown'];
  const colors = [CSS_VAR('--teal'), CSS_VAR('--indigo'), CSS_VAR('--red'), CSS_VAR('--muted')];

  const datasets = categories.map((cat, i) => ({
    label: categoryLabels[i],
    backgroundColor: colors[i],
    borderRadius: 2,
    stack: 'slot',
    data: STATE.slots.map((s) => {
      const atSlot = rows.filter((r) => r.direction === 'outbound' && r.label === s.label);
      const n = atSlot.filter((r) => r.status === cat).length;
      if (mode === 'percent') {
        const total = atSlot.length || 1;
        return +((n / total) * 100).toFixed(1);
      }
      return n;
    }),
  }));

  const wrap = document.getElementById('slotChartWrap');
  const minWidth = Math.max(wrap.clientWidth, STATE.slots.length * 46);
  document.getElementById('slotChart').style.minWidth = minWidth + 'px';

  drawBarChart(
    'slotChart', 'slot',
    STATE.slots.map((s) => s.display),
    datasets,
    mode === 'percent',
    true,
    (dataIndex) => STATE.slots[dataIndex] ? STATE.slots[dataIndex].sampleMessage : ''
  );
}

/* =====================================================================
   Chart.js helper
===================================================================== */
function wrapText(str, width) {
  if (!str) return [];
  const words = str.split(/\s+/);
  const lines = [];
  let line = '';
  words.forEach((w) => {
    if ((line + ' ' + w).trim().length > width) {
      lines.push(line.trim());
      line = w;
    } else {
      line = (line + ' ' + w).trim();
    }
  });
  if (line) lines.push(line);
  return lines.slice(0, 8);
}

function drawBarChart(canvasId, key, labels, datasets, isPercent, stacked, messageLookup) {
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
            label: (c) => `${c.dataset.label}: ${c.parsed.y}${isPercent ? '%' : ''}`,
            footer: messageLookup
              ? (items) => {
                  const idx = items[0].dataIndex;
                  const msg = messageLookup(idx);
                  return msg ? ['', 'Message sent:', ...wrapText(msg, 46)] : [];
                }
              : undefined,
          },
        },
      },
      scales: {
        x: {
          stacked: !!stacked,
          grid: { display: false },
          ticks: { font: { family: 'Inter', size: 11 } },
        },
        y: {
          stacked: !!stacked,
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
  if (key === 'slot') renderSlotChart();
});

/* =====================================================================
   SCHEDULED MESSAGES TAB — the ~36 message templates, not the 4k log
===================================================================== */
function renderScheduledMessages(query) {
  const tbody = document.getElementById('msgTableBody');
  const q = (query || '').trim().toLowerCase();
  const rows = STATE.slots.filter(
    (s) => !q || s.display.toLowerCase().includes(q) || s.sampleMessage.toLowerCase().includes(q)
  );
  tbody.innerHTML = '';
  rows.forEach((s) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="mono">${s.weekNum != null ? 'Week ' + s.weekNum : '—'}</td>
      <td class="mono">${escapeHtml(s.display)}</td>
      <td>${escapeHtml(s.sampleMessage || '—')}</td>
    `;
    tbody.appendChild(tr);
  });
  document.getElementById('msgCount').textContent = `${rows.length} scheduled message${rows.length === 1 ? '' : 's'}`;
}

document.addEventListener('DOMContentLoaded', () => {
  const search = document.getElementById('msgSearch');
  if (search) {
    search.addEventListener('input', () => renderScheduledMessages(search.value));
  }
});

/* =====================================================================
   Small utilities
===================================================================== */
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
  div.textContent = str == null ? '' : String(str);
  return div.innerHTML;
}
