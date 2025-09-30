/* ============================================================
   Unofficial Clubhouse Show Calendar (Simplified Script)
   ------------------------------------------------------------
   Cleaned-up, modular version of original JS
   ============================================================ */

/* ---------- Constants & State ---------- */
const DATA_SOURCE = './events.json';
const VENUE_TZ = 'America/Los_Angeles';

let events = [];
let filtered = [];
let selectedGenres = new Set();
let allGenres = [];

let currentView = 'calendar';
let calendarMode = 'month';
let cursorDate = new Date();

const url = new URL(window.location.href);

/* ---------- Shortcuts ---------- */
const $ = (sel, el = document) => el.querySelector(sel);
const $$ = (sel, el = document) => [...el.querySelectorAll(sel)];
const safe = (s) => (s ?? '').toString();

/* ---------- Helpers ---------- */
const ymd = (d) =>
  new Intl.DateTimeFormat('en-CA', { timeZone: VENUE_TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);

const ym = (d) =>
  new Intl.DateTimeFormat('en-CA', { timeZone: VENUE_TZ, year: 'numeric', month: '2-digit' }).format(d);

const fmtDate = (d) =>
  d.toLocaleString('en-US', { timeZone: VENUE_TZ, weekday: 'long', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit'});

const fmtDateNoTime = (d) =>
  d.toLocaleString('en-US', { timeZone: VENUE_TZ, weekday: 'long', month: 'long', day: 'numeric'});

const stageColor = (s) => {
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash = (hash * 31 + s.charCodeAt(i)) >>> 0;
  return `hsl(${hash % 360} 55% 50%)`;
};

const slugify = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
const eventId = (e) => `${ymd(e.date)}-${slugify(e.title)}`;

const setParam = (k, v) => {
  if (v) url.searchParams.set(k, v);
  else url.searchParams.delete(k);
  history.replaceState(null, '', url);
};

const debounce = (fn, ms = 200) => {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
};

/* ---------- Initialization ---------- */
document.addEventListener('DOMContentLoaded', () => loadEvents());

async function loadEvents() {
  try {
    const res = await fetch(`${DATA_SOURCE}?_=${Date.now()}`);
    const data = await res.json();

    if (!data?.events) throw new Error('Invalid data');

    events = data.events
      .map((e) => ({
        ...e,
        title: safe(e.title),
        stage: safe(e.stage),
        description: safe(e.description),
        genres: parseGenres(e.genres || e.genre),
        date: new Date(e.date),
        image: safe(e.image),
        email: safe(e.email),
        instagram: safe(e.instagram),
        website: safe(e.website),
        duration: safe(e.duration)
      }))
      .filter((e) => !isNaN(e.date))
      .sort((a, b) => a.date - b.date);

    filtered = [...events];

    $('#loadingMessage').classList.add('hidden');
    $('#controls').classList.remove('hidden');
    $('#calendarView').classList.remove('hidden');

    if (data.lastUpdated) {
      $('#lastUpdated').textContent = 'Last updated: ' + new Date(data.lastUpdated).toLocaleString();
      $('#lastUpdated').classList.remove('hidden');
    }
    $('#tzLabel').textContent = 'Times shown in ' + VENUE_TZ;
    $('#tzLabel').classList.remove('hidden');

    populateFilters();
    setupListeners();
    initFromURL();
    applyFilters();
  } catch (err) {
    console.error(err);
    $('#loadingMessage').classList.add('hidden');
    $('#errorMessage').classList.remove('hidden');
  }
}

/* ---------- Filtering ---------- */
function parseGenres(g) {
  if (!g) return [];
  if (Array.isArray(g)) return g.filter(Boolean);
  return g
    .split(',')
    .map((x) => x.trim())
    .filter(Boolean)
    .filter((v, i, a) => a.findIndex((x) => x.toLowerCase() === v.toLowerCase()) === i);
}

function populateFilters() {
  // Stage filter
  const stages = [...new Set(events.map((e) => e.stage).filter(Boolean))].sort();
  const stageSel = $('#stageFilter');
  stageSel.innerHTML = '<option value="">Stage</option>' + stages.map((s) => `<option>${s}</option>`).join('');

  // Genres
  const seen = new Set();
  allGenres = [];
  for (const e of events) {
    for (const g of e.genres || []) {
      const key = g.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        allGenres.push(g);
      }
    }
  }
  allGenres.sort((a, b) => a.localeCompare(b));
  buildGenreList(allGenres);
  updateGenreLabel();
}

function buildGenreList(list) {
  const cont = $('#genreList');
  cont.innerHTML = list
    .map((g) => `<label><input type="checkbox" value="${g}"> ${g}</label>`)
    .join('');

  $$('input[type="checkbox"]', cont).forEach((cb) => {
    cb.checked = selectedGenres.has(cb.value.toLowerCase());
    cb.addEventListener('change', () => {
      const key = cb.value.toLowerCase();
      cb.checked ? selectedGenres.add(key) : selectedGenres.delete(key);
      updateGenreLabel();
      applyFilters();
    });
  });
}

function updateGenreLabel() {
  const lbl = $('#genreBtnLabel');
  const count = selectedGenres.size;

  if (count === 0) lbl.textContent = 'Genre (any)';
  else if (count === 1) lbl.textContent = [...selectedGenres][0];
  else lbl.innerHTML = [...selectedGenres].map((g) => `<span class="pill">${g}</span>`).join(' ');

  setParam('genres', count ? [...selectedGenres].join(',') : null);
}

const applyFilters = debounce(() => {
  const q = $('#searchInput').value.toLowerCase();
  const stage = $('#stageFilter').value;
  const dateF = $('#dateFilter').value;

  const today = new Date();
  const todayKey = ymd(today);
  const weekEndKey = ymd(new Date(today.getTime() + 7 * 86400000));
  const monthKey = ym(today);

  filtered = events.filter((e) => {
    const text = (e.title + e.description + e.stage + (e.genres || []).join(' ')).toLowerCase();
    const matches =
      (!q || text.includes(q)) &&
      (!stage || e.stage === stage) &&
      (selectedGenres.size === 0 || e.genres.some((g) => selectedGenres.has(g.toLowerCase())));

    let dateMatch = true;
    const key = ymd(e.date);
    if (dateF === 'today') dateMatch = key === todayKey;
    else if (dateF === 'this-week') dateMatch = key >= todayKey && key <= weekEndKey;
    else if (dateF === 'this-month') dateMatch = ym(e.date) === monthKey;

    return matches && dateMatch;
  });

  currentView === 'list' ? renderList() : renderCalendar();
}, 150);

function clearFilters() {
  $('#searchInput').value = '';
  $('#stageFilter').value = '';
  $('#dateFilter').value = '';
  selectedGenres.clear();
  buildGenreList(allGenres);
  updateGenreLabel();
  applyFilters();
}

/* ---------- Calendar Rendering ---------- */
function renderCalendar() {
  const grid = $('#calendarGrid');
  grid.innerHTML = '';

  if (calendarMode === 'month') renderMonth(grid);
  else renderWeek(grid);
}

function renderMonth(grid) {
  const y = cursorDate.getFullYear();
  const m = cursorDate.getMonth();
  $('#monthLabel').textContent = cursorDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  const start = new Date(y, m, 1 - ((new Date(y, m, 1).getDay() + 6) % 7));
  const todayKey = ymd(new Date());

  for (let i = 0; i < 42; i++) {
    const day = new Date(start.getFullYear(), start.getMonth(), start.getDate() + i);
    const cell = document.createElement('div');
    cell.className = 'calendar-day' + (ymd(day) === todayKey ? ' today' : '');

    const dn = document.createElement('div');
    dn.className = 'day-number';
    dn.textContent = day.getDate();
    cell.appendChild(dn);

    const dayEvents = filtered.filter((e) => ymd(e.date) === ymd(day));
    for (const ev of dayEvents) {
      const item = document.createElement('div');
      item.className = 'event-item';
      item.innerHTML = `${ev.title} <span class="event-chip" style="background:${stageColor(ev.stage)}">${ev.stage}</span>`;
      item.addEventListener('click', () => openModal(ev));
      cell.appendChild(item);
    }
    grid.appendChild(cell);
  }
}

function renderWeek(grid) {
  grid.innerHTML = '';
  const start = startOfWeek(cursorDate);
  const end = new Date(start.getTime() + 6 * 86400000);
  $('#monthLabel').textContent = `Week of ${fmtDateNoTime(start)} – ${fmtDateNoTime(end)}`;

  for (let i = 0; i < 7; i++) {
    const day = new Date(start.getTime() + i * 86400000);
    const section = document.createElement('div');
    section.className = 'calendar-day';
    const dayLabel = document.createElement('div');
    dayLabel.className = 'day-number';
    dayLabel.textContent = fmtDate(day).split(',')[0];
    section.appendChild(dayLabel);

    const dayEvents = filtered.filter((e) => ymd(e.date) === ymd(day));
    if (dayEvents.length === 0) {
      const noE = document.createElement('div');
      noE.textContent = 'No events';
      noE.style.opacity = 0.5;
      section.appendChild(noE);
    } else {
      for (const ev of dayEvents) {
        const item = document.createElement('div');
        item.className = 'event-item';
        item.innerHTML = `${ev.title} <span class="event-chip" style="background:${stageColor(ev.stage)}">${ev.stage}</span>`;
        item.addEventListener('click', () => openModal(ev));
        section.appendChild(item);
      }
    }
    grid.appendChild(section);
  }
}

const startOfWeek = (d) => {
  const day = new Date(d);
  const wd = (day.getDay() + 6) % 7;
  day.setDate(day.getDate() - wd);
  return day;
};

/* ---------- List Rendering ---------- */
function renderList() {
  const wrap = $('#eventCards');
  wrap.innerHTML = '';

  filtered.forEach((e) => {
    const card = document.createElement('div');
    card.className = 'event-card';
    card.innerHTML = `
      <img src="${e.image}" alt="${e.title}" class="event-image" onerror="this.src='https://via.placeholder.com/800x420/2d3748/ffffff?text=${encodeURIComponent(e.title)}'">
      <div class="event-content">
        <div class="event-title">${e.title} <span class="event-chip" style="background:${stageColor(e.stage)}">${e.stage}</span></div>
        <div class="event-date">${fmtDate(e.date)}</div>
        <div class="event-description">${e.description || ''}</div>
        <div class="links">
          ${e.email ? `<a class="link" href="mailto:${e.email}">Email</a>` : ''}
          ${e.instagram ? `<a class="link" href="https://instagram.com/${e.instagram.replace('@','')}" target="_blank">Instagram</a>` : ''}
          ${e.website ? `<a class="link" href="${e.website.startsWith('http') ? e.website : 'https://' + e.website}" target="_blank">Website</a>` : ''}
          <a class="link" href="${googleCalendarUrl(e)}" target="_blank">Add to Calendar</a>
          <a class="link" href="${buildICS(e)}" download="${eventId(e)}.ics">.ics</a>
          <button class="link" type="button">Details</button>
        </div>
      </div>`;
    card.querySelector('button').addEventListener('click', () => openModal(e));
    card.querySelector('.event-image').addEventListener('click', () => openModal(e));
    wrap.appendChild(card);
  });
}

/* ---------- Modal ---------- */
function openModal(e) {
  $('#modalBackdrop').style.display = 'flex';
  $('#modalTitle').textContent = e.title;

  $('#modalBody').innerHTML = `
    <p><strong>Date:</strong> ${fmtDate(e.date)}</p>
    <p><strong>Stage:</strong> ${e.stage || '—'}</p>
    <p><strong>Genres:</strong> ${(e.genres || []).join(', ') || '—'}</p>
    <p><strong>Duration:</strong> ${e.duration || '—'}</p>
    <p class="event-description">${e.description || ''}</p>
    <div class="links" style="margin-top:10px">
      ${e.email ? `<a class="link" href="mailto:${e.email}">Email</a>` : ''}
      ${e.instagram ? `<a class="link" href="https://instagram.com/${e.instagram.replace('@','')}" target="_blank">Instagram</a>` : ''}
      ${e.website ? `<a class="link" href="${e.website.startsWith('http') ? e.website : 'https://' + e.website}" target="_blank">Website</a>` : ''}
      <a class="link" href="${googleCalendarUrl(e)}" target="_blank">Google Calendar</a>
      <a class="link" href="${buildICS(e)}" download="${eventId(e)}.ics">Download .ics</a>
    </div>`;
}

$('#modalClose').onclick = () => ($('#modalBackdrop').style.display = 'none');
$('#modalBackdrop').addEventListener('click', (e) => {
  if (e.target.id === 'modalBackdrop') $('#modalBackdrop').style.display = 'none';
});

/* ---------- Calendar Links ---------- */
function googleCalendarUrl(e) {
  const start = toISO(e.date);
  const end = toISO(endFromDuration(e.date, e.duration));
  return `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(e.title)}&dates=${start}/${end}&details=${encodeURIComponent(e.description || '')}&location=${encodeURIComponent(e.stage || 'Clubhouse')}`;
}

const toISO = (d) => new Date(d).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z');

function endFromDuration(start, dur) {
  const s = new Date(start);
  const m = /(\d+)\s*h|\b(\d+)\s*m/gi;
  let hours = 0,
    mins = 0,
    match;
  while ((match = m.exec(dur)) !== null) {
    if (match[1]) hours += +match[1];
    if (match[2]) mins += +match[2];
  }
  if (hours === 0 && mins === 0) hours = 1;
  s.setHours(s.getHours() + hours);
  s.setMinutes(s.getMinutes() + mins);
  return s;
}

function buildICS(e) {
  const uid = eventId(e) + '@clubhousecal';
  const dtStart = toISO(e.date);
  const dtEnd = toISO(endFromDuration(e.date, e.duration));
  return 'data:text/calendar;charset=utf-8,' + encodeURIComponent(`BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//clubhousecal//EN
BEGIN:VEVENT
UID:${uid}
DTSTAMP:${dtStart}
DTSTART:${dtStart}
DTEND:${dtEnd}
SUMMARY:${e.title}
DESCRIPTION:${e.description || ''}
LOCATION:${e.stage || 'Clubhouse'}
END:VEVENT
END:VCALENDAR`);
}

/* ---------- View Controls ---------- */
function setupListeners() {
  $('#listViewBtn').onclick = () => switchView('list');
  $('#calendarViewBtn').onclick = () => switchView('calendar');
  $('#searchInput').addEventListener('input', applyFilters);
  $('#stageFilter').addEventListener('change', applyFilters);
  $('#dateFilter').addEventListener('change', applyFilters);
  $('#clearFiltersBtn').onclick = clearFilters;

  $('#prevSpanBtn').onclick = () => changeSpan(-1);
  $('#nextSpanBtn').onclick = () => changeSpan(1);
  $('#todayBtn').onclick = () => {
    cursorDate = new Date();
    renderCalendar();
  };

  $('#modeMonthBtn').onclick = () => setMode('month');
  $('#modeWeekBtn').onclick = () => setMode('week');

  const multi = $('#genreMulti');
  const btn = $('#genreBtn');
  const panel = $('#genrePanel');
  const search = $('#genreSearch');

  btn.onclick = () => {
    const open = multi.classList.toggle('open');
    btn.setAttribute('aria-expanded', open);
    if (open) search.focus();
  };
  document.addEventListener('click', (e) => {
    if (!multi.contains(e.target)) multi.classList.remove('open');
  });
  search.addEventListener('input', () => {
    const q = search.value.toLowerCase();
    const filtered = allGenres.filter((g) => g.toLowerCase().includes(q));
    buildGenreList(filtered);
  });
}

function switchView(v) {
  currentView = v;
  setParam('view', v);

  const listBtn = $('#listViewBtn');
  const calBtn = $('#calendarViewBtn');

  if (v === 'list') {
    $('#calendarView').classList.add('hidden');
    $('#listView').classList.remove('hidden');
    listBtn.classList.add('active');
    calBtn.classList.remove('active');
    renderList();
  } else {
    $('#listView').classList.add('hidden');
    $('#calendarView').classList.remove('hidden');
    calBtn.classList.add('active');
    listBtn.classList.remove('active');
    renderCalendar();
  }
}

function setMode(mode) {
  calendarMode = mode;
  setParam('mode', mode);
  $('#modeMonthBtn').classList.toggle('active', mode === 'month');
  $('#modeWeekBtn').classList.toggle('active', mode === 'week');
  renderCalendar();
}

function changeSpan(dir) {
  cursorDate.setDate(cursorDate.getDate() + dir * (calendarMode === 'month' ? 30 : 7));
  renderCalendar();
}

/* ---------- URL Sync ---------- */
function initFromURL() {
  currentView = url.searchParams.get('view') || (window.innerWidth <= 768 ? 'list' : 'calendar');
  calendarMode = url.searchParams.get('mode') || (window.innerWidth <= 768 ? 'week' : 'month');
  $('#searchInput').value = url.searchParams.get('q') || '';
  $('#stageFilter').value = url.searchParams.get('stage') || '';
  $('#dateFilter').value = url.searchParams.get('date') || '';
  const g = url.searchParams.get('genres');
  if (g) g.split(',').forEach((x) => selectedGenres.add(x.toLowerCase()));

  $('#filters').classList.toggle('collapsed', window.innerWidth <= 768);
  switchView(currentView);
  setMode(calendarMode);
}
