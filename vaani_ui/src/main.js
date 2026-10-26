import {
  LEAD_STATUSES,
  PIPELINE_STAGES,
  STATUS_META,
  STORAGE_KEY,
  addLeadNote,
  createInitialState,
  createLead,
  filterLeads,
  formatDate,
  formatDateTime,
  formatRelative,
  getMetrics,
  getVisibleLeads,
  initials,
  isFollowupDue,
  makeImportedLeads,
  moveStatus,
  todayIso,
  updateLead,
  validateLead,
} from './core.js';

const app = document.querySelector('#app');
const toastRoot = document.querySelector('#toast-root');
const reducedMotion = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches ?? false;
const crmPort = Number(window.location.port || 0);
const apiMode = Boolean(window.VAANI_API) || (crmPort >= 8765 && crmPort <= 8774);
let apiReady = false;

const ICONS = {
  grid: '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
  users: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
  board: '<rect x="3" y="3" width="5.5" height="18" rx="1.5"/><rect x="9.75" y="3" width="5.5" height="12" rx="1.5"/><rect x="16.5" y="3" width="5.5" height="8" rx="1.5"/>',
  phone: '<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.13.96.36 1.9.7 2.8a2 2 0 0 1-.45 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.45c.9.34 1.84.57 2.8.7A2 2 0 0 1 22 16.9z"/>',
  calendar: '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  upload: '<path d="M12 16V4m0 0L7 9m5-5 5 5M5 20h14"/>',
  download: '<path d="M12 4v12m0 0 5-5m-5 5-5-5M5 20h14"/>',
  arrow: '<path d="M5 12h14M13 6l6 6-6 6"/>',
  chevron: '<path d="m9 18 6-6-6-6"/>',
  close: '<path d="m6 6 12 12M18 6 6 18"/>',
  check: '<path d="m5 12 4 4L19 6"/>',
  flame: '<path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"/>',
  sparkle: '<path d="m12 3-1.3 5.7L5 10l5.7 1.3L12 17l1.3-5.7L19 10l-5.7-1.3L12 3Z"/>',
  dots: '<circle cx="5" cy="12" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="19" cy="12" r="1"/>',
  external: '<path d="M14 4h6v6M20 4l-9 9"/><path d="M18 13v5a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h5"/>',
  note: '<path d="M4 4h16v16H4zM8 9h8M8 13h8M8 17h5"/>',
  filter: '<path d="M4 6h16M7 12h10M10 18h4"/>',
  menu: '<path d="M4 7h16M4 12h16M4 17h16"/>',
  settings: '<path d="M12 15.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7Z"/><path d="m19.4 15 .1.1a2 2 0 0 1-2.8 2.8l-.1-.1a2 2 0 0 0-3.4 1.4v.2a2 2 0 0 1-4 0v-.2a2 2 0 0 0-3.4-1.4l-.1.1A2 2 0 0 1 3 15.1l.1-.1a2 2 0 0 0-1.4-3.4h-.2a2 2 0 0 1 0-4h.2A2 2 0 0 0 3.1 4.2L3 4.1A2 2 0 0 1 5.8 1.3l.1.1a2 2 0 0 0 3.4-1.4v-.2a2 2 0 0 1 4 0V0a2 2 0 0 0 3.4 1.4l.1-.1A2 2 0 0 1 19.6 4l-.1.1A2 2 0 0 0 20.9 7.5h.2a2 2 0 0 1 0 4h-.2a2 2 0 0 0-1.5 3.5Z"/>',
};

const escapeHtml = (value = '') => String(value)
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')
  .replaceAll("'", '&#039;');

const icon = (name, size = 18) => `<svg class="icon" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name] || ''}</svg>`;

const state = loadState();
let searchTimer;
let toastTimer;

function loadState() {
  const base = createInitialState();
  if (apiMode) return base;
  try {
    const stored = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    if (!stored) return base;
    return {
      ...base,
      ...stored,
      filters: { ...base.filters, ...(stored.filters || {}) },
      leads: Array.isArray(stored.leads) && stored.leads.length ? stored.leads : base.leads,
      runs: Array.isArray(stored.runs) && stored.runs.length ? stored.runs : base.runs,
      selectedIds: [],
      selectedLeadId: null,
    };
  } catch {
    return base;
  }
}

function saveState() {
  state.lastSavedAt = new Date().toISOString();
  if (apiMode) {
    if (!apiReady) return;
    fetch('/api/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ leads: state.leads }),
    }).catch(() => showToast('The local database could not be reached.', 'danger'));
    return;
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    ...state,
    selectedIds: [],
    selectedLeadId: null,
    mobileNavOpen: false,
  }));
}

function normalizeApiLead(lead) {
  return {
    ...lead,
    id: lead.id || `lead-${lead.phone}`,
    vertical: lead.vertical || state.workspace,
    name: lead.name || 'Unnamed lead',
    contact: lead.contact || '',
    location: lead.location || lead.address || '',
    phone: lead.phone || '',
    email: lead.email || '',
    website: lead.website || '',
    score: Number(lead.score) || 0,
    reviews: Number(lead.reviews) || 0,
    rating: Number(lead.rating) || 0,
    tier: lead.tier || 'C',
    status: lead.status || 'Not called',
    notes: Array.isArray(lead.notes) ? lead.notes : [],
    signals: Array.isArray(lead.signals) ? lead.signals : [],
  };
}

async function hydrateFromApi() {
  const response = await fetch('/api/bootstrap', { cache: 'no-store' });
  if (!response.ok) throw new Error(`CRM API returned ${response.status}`);
  const payload = await response.json();
  state.workspace = payload.workspace || state.workspace;
  state.leads = Array.isArray(payload.leads) ? payload.leads.map(normalizeApiLead) : state.leads;
  state.runs = Array.isArray(payload.runs) ? payload.runs : state.runs;
  state.lastSavedAt = payload.lastSavedAt || new Date().toISOString();
  apiReady = true;
}

function todayLabel() {
  return new Intl.DateTimeFormat('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).format(new Date());
}

function greeting() {
  const hour = new Date().getHours();
  if (hour < 12) return 'Good morning';
  if (hour < 17) return 'Good afternoon';
  return 'Good evening';
}

function workspaceLeads() {
  return state.leads.filter((lead) => lead.vertical === state.workspace);
}

function leadById(id) {
  return state.leads.find((lead) => lead.id === id);
}

function phoneDigits(phone = '') {
  return phone.replace(/\D/g, '').replace(/^91(?=\d{10}$)/, '');
}

function websiteHref(website = '') {
  if (!website) return '';
  return /^https?:\/\//i.test(website) ? website : `https://${website}`;
}

function initialsAvatar(name, extra = '') {
  return `<span class="avatar ${extra}" aria-hidden="true">${escapeHtml(initials(name))}</span>`;
}

function statusPill(status) {
  const meta = STATUS_META[status] || STATUS_META['Not called'];
  return `<span class="status-pill status-pill--${meta.tone}"><span class="status-dot"></span>${escapeHtml(meta.label)}</span>`;
}

function tierBadge(tier) {
  return `<span class="tier-badge tier-badge--${escapeHtml(tier || 'C')}">${escapeHtml(tier || '—')}</span>`;
}

function scoreBadge(score) {
  const value = Number(score) || 0;
  const tone = value >= 75 ? 'high' : value >= 55 ? 'mid' : 'low';
  return `<span class="score-badge score-badge--${tone}"><strong>${value || '—'}</strong><small>/100</small></span>`;
}

function followupLabel(lead) {
  if (!lead.followup) return '<span class="muted">No follow-up</span>';
  const due = isFollowupDue(lead);
  const label = lead.followup === todayIso() ? 'Today' : formatDate(lead.followup);
  return `<span class="followup ${due ? 'followup--due' : ''}">${icon('calendar', 13)} ${label}</span>`;
}

function metricCard(label, value, note, iconName, tone = '') {
  return `<article class="metric-card">
    <div class="metric-card__top"><span class="metric-card__label">${escapeHtml(label)}</span><span class="metric-card__icon ${tone}">${icon(iconName, 17)}</span></div>
    <strong class="metric-card__value">${escapeHtml(value)}</strong>
    <span class="metric-card__note ${tone === 'danger' ? 'metric-card__note--danger' : ''}">${escapeHtml(note)}</span>
  </article>`;
}

function renderLayout() {
  const metrics = getMetrics(state.leads, state.workspace);
  const pageTitles = {
    dashboard: 'Overview',
    leads: 'Leads',
    pipeline: 'Pipeline',
    followups: 'Follow-ups',
    runs: 'Import runs',
    settings: 'Workspace settings',
  };
  const title = pageTitles[state.view] || 'Overview';
  const navItem = (view, label, iconName, badge = '') => `<a class="nav-item ${state.view === view ? 'is-active' : ''}" href="#${view}" data-view="${view}" data-nav="${view}">${icon(iconName, 18)}<span>${label}</span>${badge ? `<span class="nav-badge">${badge}</span>` : ''}</a>`;
  const workspaceLabel = state.workspace === 'Healthcare' ? 'Healthcare' : 'Restaurants';

  app.innerHTML = `<div class="app-frame ${state.mobileNavOpen ? 'nav-open' : ''}">
    <aside class="sidebar" data-sidebar>
      <div class="brand-lockup"><div class="brand-mark">V</div><div><strong>vaani</strong><span>lead workspace</span></div></div>
      <div class="workspace-picker">
        <label for="workspace-select">Workspace</label>
        <div class="workspace-select-wrap"><span class="workspace-dot"></span><select id="workspace-select" data-workspace aria-label="Choose workspace" ${apiMode ? 'disabled' : ''}><option ${workspaceLabel === 'Restaurants' ? 'selected' : ''}>Restaurants</option><option ${workspaceLabel === 'Healthcare' ? 'selected' : ''}>Healthcare</option></select>${icon('chevron', 15)}</div>
      </div>
      <div class="sidebar-section-label">Workspace</div>
      <nav class="primary-nav" aria-label="Primary navigation">
        ${navItem('dashboard', 'Overview', 'grid')}
        ${navItem('leads', 'All leads', 'users', metrics.total)}
        ${navItem('pipeline', 'Pipeline', 'board')}
        ${navItem('followups', 'Follow-ups', 'calendar', metrics.due || '')}
      </nav>
      <div class="sidebar-section-label sidebar-section-label--spaced">Operations</div>
      <nav class="primary-nav" aria-label="Operations navigation">
        ${navItem('runs', 'Import runs', 'upload')}
        ${navItem('settings', 'Settings', 'settings')}
      </nav>
      <div class="sidebar-spacer"></div>
       <div class="storage-card"><div class="storage-card__top"><span>${apiMode ? 'SQLite workspace' : 'Local workspace'}</span><span class="online-dot"></span></div><strong>${apiMode ? 'Connected to local CRM data' : 'Your data stays here'}</strong><span>${apiMode ? 'Changes save to the local database' : 'Saved automatically in this browser'}</span><div class="storage-bar"><i style="width: 68%"></i></div><small>${apiMode ? 'Database-backed workspace' : 'Demo database · 68% organised'}</small></div>
      <div class="profile-card"><span class="profile-avatar">Y</span><div><strong>Yash Tanna</strong><span>Owner workspace</span></div>${icon('dots', 16)}</div>
    </aside>
    <div class="sidebar-scrim" data-close-nav></div>
    <div class="main-shell">
      <header class="topbar">
        <button class="icon-button mobile-menu" data-mobile-menu aria-label="Open navigation">${icon('menu', 20)}</button>
        <div class="breadcrumb"><span>Workspace</span>${icon('chevron', 14)}<strong>${escapeHtml(title)}</strong></div>
        <div class="topbar-actions">
          <label class="global-search"><span>${icon('search', 17)}</span><input data-global-search value="${escapeHtml(state.query)}" placeholder="Search leads, locations…" aria-label="Search leads" /><kbd>⌘ K</kbd></label>
          <button class="button button--secondary topbar-import" data-open-import>${icon('upload', 16)}<span>Import</span></button>
          <button class="button button--primary topbar-add" data-open-add>${icon('plus', 16)}<span>Add lead</span></button>
          <button class="icon-button notification-button" aria-label="Notifications"><span class="notification-dot"></span>${icon('clock', 18)}</button>
          <span class="topbar-avatar">Y</span>
        </div>
      </header>
      <main id="main-content" class="page-wrap"><div id="page-content"></div></main>
      <footer class="app-footer"><span>Vaani CRM · ${escapeHtml(state.workspace)} workspace</span><span><span class="online-dot"></span> ${apiMode ? 'SQLite synced' : 'Saved locally'} ${formatRelative(state.lastSavedAt)}</span></footer>
    </div>
    <aside class="lead-drawer" id="lead-drawer" aria-hidden="true"></aside>
  </div>`;
  renderContent();
  renderDrawer();
}

function renderContent() {
  const content = document.querySelector('#page-content');
  if (!content) return;
  const renderers = {
    dashboard: renderDashboard,
    leads: renderLeads,
    pipeline: renderPipeline,
    followups: renderFollowups,
    runs: renderRuns,
    settings: renderSettings,
  };
  content.innerHTML = (renderers[state.view] || renderDashboard)();
  updateChrome();
}

function updateChrome() {
  const metrics = getMetrics(state.leads, state.workspace);
  document.querySelectorAll('[data-nav="followups"] .nav-badge').forEach((node) => { node.textContent = metrics.due || ''; });
}

function renderDashboard() {
  const metrics = getMetrics(state.leads, state.workspace);
  const visible = workspaceLeads();
  const fresh = [...visible].sort((a, b) => b.firstSeen.localeCompare(a.firstSeen) || b.score - a.score).slice(0, 5);
  const activity = visible.flatMap((lead) => (lead.notes || []).map((note) => ({ lead, note }))).sort((a, b) => b.note.createdAt.localeCompare(a.note.createdAt)).slice(0, 5);
  const stageCounts = PIPELINE_STAGES.map((stage) => ({ ...stage, count: visible.filter((lead) => stage.statuses.includes(lead.status)).length }));
  const maxStage = Math.max(...stageCounts.map((stage) => stage.count), 1);
  const tierTotal = Math.max(metrics.total, 1);

  return `<section class="page-heading page-heading--dashboard"><div><div class="eyebrow"><span class="eyebrow-dot"></span>${escapeHtml(state.workspace)} workspace</div><h1>${greeting()}, Yash.</h1><p>${todayLabel()} · Here’s the clearest next move for your lead engine.</p></div><div class="heading-actions"><button class="button button--secondary" data-open-import>${icon('upload', 16)} Import scrape</button><button class="button button--primary" data-start-calling>${icon('phone', 16)} Start calling <span class="button__shortcut">⌘ ↵</span></button></div></section>
    <section class="hero-banner"><div class="hero-banner__glow"></div><div class="hero-banner__copy"><span class="hero-kicker">Today’s operating view</span><h2>Turn fresh data into<br><em>the next conversation.</em></h2><p>Your highest-intent leads are already prioritised. Clear due follow-ups first, then work the A-tier queue.</p><div class="hero-banner__actions"><button class="button button--light" data-view="leads" data-filter-shortcut="attention">Review attention queue ${icon('arrow', 15)}</button><button class="text-button text-button--light" data-view="pipeline">Open pipeline ${icon('arrow', 14)}</button></div></div><div class="hero-banner__stats"><div><strong>${metrics.due}</strong><span>follow-ups due</span></div><div><strong>${metrics.tiers.A}</strong><span>A-tier ready</span></div><div><strong>${metrics.demos}</strong><span>demos in motion</span></div></div><div class="hero-banner__orb"><span>V</span><small>local<br>workspace</small></div></section>
    <section class="metrics-grid">${metricCard('Active leads', String(metrics.active), `${metrics.total} total in workspace`, 'users')}${metricCard('Due today', String(metrics.due), metrics.overdue ? `${metrics.overdue} overdue · clear first` : 'Nothing overdue', 'calendar', metrics.overdue ? 'danger' : '')}${metricCard('Average score', `${metrics.avgScore}/100`, `${metrics.tiers.A} A-tier leads`, 'flame')}${metricCard('Win rate', `${metrics.winRate}%`, metrics.closed ? `${metrics.won} won from ${metrics.closed} closed` : 'No closed leads yet', 'check', metrics.won ? 'success' : '')}</section>
    <section class="dashboard-grid dashboard-grid--primary"><article class="panel panel--focus"><div class="panel-heading"><div><span class="panel-kicker">Your next actions</span><h2>Today’s focus</h2></div><button class="text-button" data-view="followups">View all ${icon('arrow', 14)}</button></div>${metrics.dueLeads.length ? `<div class="focus-list">${metrics.dueLeads.slice(0, 4).map((lead) => focusRow(lead)).join('')}</div>` : emptyState('You’re clear for today', 'No overdue or due follow-ups. Add a date to any lead when you need a reminder.', 'Set a follow-up', 'leads')}</article><article class="panel panel--funnel"><div class="panel-heading"><div><span class="panel-kicker">Pipeline health</span><h2>Where deals sit</h2></div><button class="icon-button" data-view="pipeline" aria-label="Open pipeline">${icon('arrow', 16)}</button></div><div class="pipeline-bars">${stageCounts.map((stage) => `<button class="pipeline-bar" data-view="pipeline"><span class="pipeline-bar__label"><span>${escapeHtml(stage.label)}</span><strong>${stage.count}</strong></span><span class="pipeline-bar__track"><i style="width:${Math.max(stage.count ? 8 : 0, Math.round((stage.count / maxStage) * 100))}%" class="pipeline-bar__fill pipeline-bar__fill--${stage.id}"></i></span></button>`).join('')}</div><div class="funnel-footer"><span><i class="legend-dot legend-dot--a"></i>A-tier ${Math.round((metrics.tiers.A / tierTotal) * 100)}%</span><span><i class="legend-dot legend-dot--b"></i>B-tier ${Math.round((metrics.tiers.B / tierTotal) * 100)}%</span><span><i class="legend-dot legend-dot--c"></i>C-tier ${Math.round((metrics.tiers.C / tierTotal) * 100)}%</span></div></article></section>
    <section class="dashboard-grid dashboard-grid--secondary"><article class="panel panel--table"><div class="panel-heading"><div><span class="panel-kicker">Latest signal</span><h2>Fresh from your scrape</h2></div><button class="text-button" data-view="leads" data-filter-shortcut="new">See all new leads ${icon('arrow', 14)}</button></div><div class="table-wrap">${leadTable(fresh, { compact: true, showSelect: false })}</div></article><article class="panel panel--activity"><div class="panel-heading"><div><span class="panel-kicker">Recent activity</span><h2>What changed</h2></div><span class="live-indicator"><i></i>Live</span></div>${activity.length ? `<div class="activity-list">${activity.map(activityRow).join('')}</div>` : emptyState('No activity yet', 'Your first note or status change will appear here.', 'Open leads', 'leads')}</article></section>`;
}

function focusRow(lead) {
  const overdue = lead.followup < todayIso();
  return `<button class="focus-row" data-open-lead="${escapeHtml(lead.id)}"><span class="focus-row__icon ${overdue ? 'is-overdue' : ''}">${icon(overdue ? 'clock' : 'calendar', 16)}</span><span class="focus-row__body"><strong>${escapeHtml(lead.name)}</strong><small>${escapeHtml(lead.contact || lead.location)} · ${overdue ? `${Math.max(1, Math.round((Date.now() - new Date(`${lead.followup}T12:00:00`).valueOf()) / 86400000))}d overdue` : 'Due today'}</small></span>${statusPill(lead.status)}${icon('chevron', 15)}</button>`;
}

function activityRow({ lead, note }) {
  return `<button class="activity-row" data-open-lead="${escapeHtml(lead.id)}">${initialsAvatar(lead.name)}<span class="activity-row__body"><strong>${escapeHtml(note.text)}</strong><small>${escapeHtml(lead.name)} · ${formatRelative(note.createdAt)}</small></span><span class="activity-arrow">${icon('chevron', 14)}</span></button>`;
}

function renderLeads() {
  const rows = getVisibleLeads(state);
  const metrics = getMetrics(state.leads, state.workspace);
  const selectedCount = state.selectedIds.length;
  return `<section class="page-heading"><div><div class="eyebrow"><span class="eyebrow-dot"></span>Lead database</div><h1>All leads <span class="heading-count">${metrics.total}</span></h1><p>Every business Vaani has found, scored and kept ready for a useful next step.</p></div><div class="heading-actions"><button class="button button--secondary" data-export>${icon('download', 16)} Export CSV</button><button class="button button--primary" data-open-add>${icon('plus', 16)} Add lead</button></div></section>
    <section class="lead-toolbar"><div class="filter-tabs" role="tablist" aria-label="Lead shortcuts"><button class="filter-tab ${state.filters.date === 'all' && state.filters.tier === 'all' && state.filters.status === 'all' ? 'is-active' : ''}" data-filter-shortcut="all">All leads <span>${metrics.total}</span></button><button class="filter-tab ${state.filters.date === 'attention' ? 'is-active' : ''}" data-filter-shortcut="attention">Needs attention <span>${metrics.due}</span></button><button class="filter-tab ${state.filters.tier === 'A' ? 'is-active' : ''}" data-filter-shortcut="tierA">A-tier <span>${metrics.tiers.A}</span></button><button class="filter-tab ${state.filters.status === 'Not called' ? 'is-active' : ''}" data-filter-shortcut="new">Not contacted <span>${state.leads.filter((lead) => lead.vertical === state.workspace && lead.status === 'Not called').length}</span></button></div><button class="button button--secondary button--filter" data-toggle-filters>${icon('filter', 15)} Filters${activeFilterCount() ? `<span class="filter-count">${activeFilterCount()}</span>` : ''}</button></section>
    <section class="filter-panel ${activeFilterCount() ? 'is-visible' : ''}" data-filter-panel><label><span>Status</span><select data-status-filter>${statusSelectOptions(state.filters.status, true)}</select></label><label><span>Tier</span><select data-tier-filter><option value="all">All tiers</option><option value="A" ${state.filters.tier === 'A' ? 'selected' : ''}>A-tier</option><option value="B" ${state.filters.tier === 'B' ? 'selected' : ''}>B-tier</option><option value="C" ${state.filters.tier === 'C' ? 'selected' : ''}>C-tier</option></select></label><label><span>Added</span><select data-date-filter><option value="all" ${state.filters.date === 'all' ? 'selected' : ''}>Any time</option><option value="today" ${state.filters.date === 'today' ? 'selected' : ''}>Added today</option><option value="7d" ${state.filters.date === '7d' ? 'selected' : ''}>Last 7 days</option><option value="attention" ${state.filters.date === 'attention' ? 'selected' : ''}>Needs attention</option></select></label><label><span>Sort by</span><select data-sort><option value="score" ${state.sort === 'score' ? 'selected' : ''}>Best score</option><option value="followup" ${state.sort === 'followup' ? 'selected' : ''}>Follow-up</option><option value="newest" ${state.sort === 'newest' ? 'selected' : ''}>Newest added</option><option value="recent" ${state.sort === 'recent' ? 'selected' : ''}>Recently seen</option><option value="name" ${state.sort === 'name' ? 'selected' : ''}>Name</option></select></label><button class="text-button" data-clear-filters>Clear filters</button></section>
    ${selectedCount ? `<section class="bulk-bar"><span><strong>${selectedCount}</strong> selected</span><div><select data-bulk-status><option value="">Change status…</option>${statusSelectOptions('', false)}</select><button class="button button--secondary" data-clear-selection>Clear selection</button></div></section>` : ''}
    <section class="panel panel--leads-table"><div class="table-meta"><span><strong>${rows.length}</strong> shown${state.query ? ` for “${escapeHtml(state.query)}”` : ''}</span><span class="table-meta__hint">Click a lead to open its workspace</span></div><div class="table-wrap">${rows.length ? leadTable(rows, { showSelect: true }) : emptyState('No leads match this view', 'Try clearing a filter or import a new scrape to keep your pipeline moving.', 'Clear filters', 'clear-filters')}</div></section>`;
}

function activeFilterCount() {
  return [state.filters.status !== 'all', state.filters.tier !== 'all', state.filters.date !== 'all'].filter(Boolean).length;
}

function statusSelectOptions(selected = 'all', includeAll = false) {
  const all = includeAll ? `<option value="all">All statuses</option>` : '';
  return all + LEAD_STATUSES.map((status) => `<option value="${escapeHtml(status)}" ${selected === status ? 'selected' : ''}>${escapeHtml(status)}</option>`).join('');
}

function leadTable(rows, { compact = false, showSelect = true } = {}) {
  const header = `<thead><tr>${showSelect ? '<th class="table-check"><span class="sr-only">Select</span></th>' : ''}<th>Lead</th><th>Intent</th><th>Status</th><th>Next step</th><th>Source</th><th class="table-actions-heading"><span class="sr-only">Actions</span></th></tr></thead>`;
  const body = rows.map((lead) => `<tr class="lead-row" data-lead-row="${escapeHtml(lead.id)}">${showSelect ? `<td class="table-check"><input type="checkbox" data-select-lead="${escapeHtml(lead.id)}" ${state.selectedIds.includes(lead.id) ? 'checked' : ''} aria-label="Select ${escapeHtml(lead.name)}"></td>` : ''}<td><button class="lead-cell" data-open-lead="${escapeHtml(lead.id)}">${initialsAvatar(lead.name)}<span><strong>${escapeHtml(lead.name)}</strong><small>${escapeHtml(lead.location)}</small></span></button></td><td><div class="intent-cell">${scoreBadge(lead.score)}${tierBadge(lead.tier)}</div><small class="angle-text">${escapeHtml(lead.angle || 'New opportunity')}</small></td><td>${statusPill(lead.status)}</td><td>${followupLabel(lead)}</td><td><span class="source-cell">${escapeHtml(lead.source || 'Manual entry')}</span></td><td class="table-actions"><button class="icon-button icon-button--small" data-open-lead="${escapeHtml(lead.id)}" aria-label="Open ${escapeHtml(lead.name)}">${icon('arrow', 15)}</button></td></tr>`).join('');
  return `<table class="leads-table ${compact ? 'leads-table--compact' : ''}">${header}<tbody>${body}</tbody></table>`;
}

function renderPipeline() {
  const leads = workspaceLeads();
  return `<section class="page-heading"><div><div class="eyebrow"><span class="eyebrow-dot"></span>Sales funnel</div><h1>Pipeline</h1><p>Move leads forward without losing the context behind each conversation.</p></div><div class="heading-actions"><button class="button button--secondary" data-view="leads">Browse leads ${icon('arrow', 15)}</button><button class="button button--primary" data-start-calling>${icon('phone', 16)} Start calling</button></div></section><section class="pipeline-summary"><div><span>Active pipeline</span><strong>${leads.filter((lead) => !['Won', 'Lost', 'Not interested', 'Wrong number', 'Do not contact'].includes(lead.status)).length}</strong><small>leads still in motion</small></div><div><span>Weighted intent</span><strong>${Math.round(leads.reduce((sum, lead) => sum + lead.score, 0) / Math.max(leads.length, 1))}/100</strong><small>average lead score</small></div><div><span>Won this cycle</span><strong>${leads.filter((lead) => lead.status === 'Won').length}</strong><small>closed-won accounts</small></div><div><span>Needs attention</span><strong class="text-danger">${leads.filter((lead) => isFollowupDue(lead)).length}</strong><small>due or overdue</small></div></section><section class="pipeline-board" data-pipeline-board>${PIPELINE_STAGES.map((stage) => { const stageLeads = leads.filter((lead) => stage.statuses.includes(lead.status)); return `<div class="pipeline-column" data-stage="${stage.id}" data-status="${escapeHtml(stage.statuses[0])}"><div class="pipeline-column__head"><span class="pipeline-column__title"><i class="stage-dot stage-dot--${stage.id}"></i>${escapeHtml(stage.label)}</span><span class="column-count">${stageLeads.length}</span></div><div class="pipeline-column__body">${stageLeads.length ? stageLeads.map(pipelineCard).join('') : `<div class="pipeline-empty">Drop a lead here</div>`}</div></div>`; }).join('')}</section><div class="pipeline-hint">${icon('sparkle', 15)} Drag a lead to a new stage. Status and timeline update instantly.</div>`;
}

function pipelineCard(lead) {
  return `<article class="pipeline-card" draggable="true" data-drag-lead="${escapeHtml(lead.id)}" data-open-lead="${escapeHtml(lead.id)}"><div class="pipeline-card__top"><span>${tierBadge(lead.tier)}</span><button class="icon-button icon-button--small" data-open-lead="${escapeHtml(lead.id)}" aria-label="Open lead">${icon('dots', 15)}</button></div><strong>${escapeHtml(lead.name)}</strong><small>${escapeHtml(lead.location)}</small><div class="pipeline-card__meta"><span>${scoreBadge(lead.score)}</span>${lead.followup ? followupLabel(lead) : `<span class="muted">${escapeHtml(lead.angle)}</span>`}</div></article>`;
}

function renderFollowups() {
  const metrics = getMetrics(state.leads, state.workspace);
  const scoped = workspaceLeads();
  const today = scoped.filter((lead) => lead.followup === todayIso() && !['Won', 'Lost', 'Not interested', 'Wrong number', 'Do not contact'].includes(lead.status));
  const upcoming = scoped.filter((lead) => lead.followup > todayIso() && !['Won', 'Lost', 'Not interested', 'Wrong number', 'Do not contact'].includes(lead.status)).sort((a, b) => a.followup.localeCompare(b.followup));
  const overdue = scoped.filter((lead) => lead.followup && lead.followup < todayIso() && isFollowupDue(lead)).sort((a, b) => a.followup.localeCompare(b.followup));
  const completed = scoped.filter((lead) => ['Won', 'Lost', 'Not interested'].includes(lead.status)).sort((a, b) => b.lastSeen.localeCompare(a.lastSeen));
  const tab = state.followupTab;
  const shown = tab === 'today' ? [...overdue, ...today] : tab === 'upcoming' ? upcoming : tab === 'completed' ? completed : [...overdue, ...today, ...upcoming];
  return `<section class="page-heading"><div><div class="eyebrow"><span class="eyebrow-dot"></span>Next actions</div><h1>Follow-ups</h1><p>Keep every promise visible, with overdue work pulled to the top.</p></div><div class="heading-actions"><button class="button button--primary" data-start-calling>${icon('phone', 16)} Work the queue</button></div></section><section class="followup-summary"><div class="followup-summary__main"><span class="panel-kicker">Today</span><strong>${metrics.due}</strong><span>follow-ups need your attention</span></div><div class="followup-summary__stat"><span>Overdue</span><strong class="text-danger">${overdue.length}</strong></div><div class="followup-summary__stat"><span>Next 7 days</span><strong>${upcoming.filter((lead) => lead.followup <= new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10)).length}</strong></div><div class="followup-summary__stat"><span>Completed</span><strong>${completed.length}</strong></div></section><section class="subnav-tabs"><button class="subnav-tab ${tab === 'today' ? 'is-active' : ''}" data-followup-tab="today">Today <span>${overdue.length + today.length}</span></button><button class="subnav-tab ${tab === 'upcoming' ? 'is-active' : ''}" data-followup-tab="upcoming">Upcoming <span>${upcoming.length}</span></button><button class="subnav-tab ${tab === 'completed' ? 'is-active' : ''}" data-followup-tab="completed">Completed <span>${completed.length}</span></button><button class="subnav-tab ${tab === 'all' ? 'is-active' : ''}" data-followup-tab="all">All dates <span>${overdue.length + today.length + upcoming.length}</span></button></section><section class="followup-list">${shown.length ? shown.map((lead) => followupCard(lead, overdue.includes(lead))).join('') : emptyState(tab === 'today' ? 'No follow-ups due' : 'Nothing scheduled here', 'Set a follow-up from a lead drawer so the next action never gets lost.', 'Open leads', 'leads')}</section>`;
}

function followupCard(lead, overdue = false) {
  return `<article class="followup-card ${overdue ? 'followup-card--overdue' : ''}"><span class="followup-card__date"><strong>${formatDate(lead.followup, { weekday: 'short' })}</strong><small>${overdue ? `${Math.max(1, Math.round((Date.now() - new Date(`${lead.followup}T12:00:00`).valueOf()) / 86400000))}d late` : lead.followup === todayIso() ? 'Today' : 'Scheduled'}</small></span>${initialsAvatar(lead.name)}<div class="followup-card__body"><button class="link-button" data-open-lead="${escapeHtml(lead.id)}">${escapeHtml(lead.name)}</button><span>${escapeHtml(lead.contact || 'Decision-maker')} · ${escapeHtml(lead.location)}</span><small>${escapeHtml(lead.angle || 'Next conversation')}</small></div><div class="followup-card__status">${statusPill(lead.status)}${scoreBadge(lead.score)}</div><div class="followup-card__actions"><a class="quick-action" href="tel:${phoneDigits(lead.phone)}" aria-label="Call ${escapeHtml(lead.name)}">${icon('phone', 16)}</a><button class="quick-action" data-open-lead="${escapeHtml(lead.id)}" aria-label="Open ${escapeHtml(lead.name)}">${icon('arrow', 16)}</button></div></article>`;
}

function renderRuns() {
  const runs = state.runs.filter((run) => run.vertical === state.workspace);
  const totalFound = runs.reduce((sum, run) => sum + run.found, 0);
  const totalSaved = runs.reduce((sum, run) => sum + run.saved, 0);
  return `<section class="page-heading"><div><div class="eyebrow"><span class="eyebrow-dot"></span>Scrape history</div><h1>Import runs</h1><p>See where your lead database came from and keep the next search focused.</p></div><div class="heading-actions"><button class="button button--primary" data-open-import>${icon('upload', 16)} Import a scrape</button></div></section><section class="run-summary"><div><span>Total found</span><strong>${totalFound}</strong><small>across ${runs.length} runs</small></div><div><span>Saved to CRM</span><strong>${totalSaved}</strong><small>after quality filters</small></div><div><span>Contactable</span><strong>${runs.reduce((sum, run) => sum + run.contacts, 0)}</strong><small>with a mobile number</small></div></section><section class="runs-list">${runs.map(runCard).join('')}</section><section class="run-empty-cta"><div class="run-empty-cta__icon">${icon('sparkle', 22)}</div><div><strong>Have a new area in mind?</strong><p>Start with a location and category. Vaani will keep the run, score the results and add only new businesses to your workspace.</p></div><button class="button button--secondary" data-open-import>Start a new run ${icon('arrow', 15)}</button></section>`;
}

function runCard(run) {
  return `<article class="run-card"><div class="run-card__icon">${icon('upload', 19)}</div><div class="run-card__main"><div class="run-card__heading"><div><strong>${escapeHtml(run.location)}</strong><span>${escapeHtml(run.category)} · ${formatRelative(run.createdAt)}</span></div><span class="run-status"><i></i>${escapeHtml(run.status)}</span></div><div class="run-card__stats"><span><strong>${run.found}</strong> found</span><span><strong>${run.saved}</strong> saved</span><span><strong>${run.contacts}</strong> with contacts</span></div><div class="run-card__bar"><i style="width:${Math.min(100, Math.round((run.saved / Math.max(run.found, 1)) * 100))}%"></i></div></div><div class="run-card__actions"><button class="button button--secondary" data-run-review="${escapeHtml(run.id)}">Review leads ${icon('arrow', 14)}</button><button class="icon-button" data-run-again="${escapeHtml(run.location)}" aria-label="Run again for ${escapeHtml(run.location)}">${icon('dots', 17)}</button></div></article>`;
}

function renderSettings() {
  const metrics = getMetrics(state.leads, state.workspace);
  return `<section class="page-heading"><div><div class="eyebrow"><span class="eyebrow-dot"></span>Workspace controls</div><h1>Settings</h1><p>Keep this local workspace tidy, portable and easy to hand off.</p></div></section><section class="settings-grid"><article class="panel settings-card"><div class="settings-card__icon">${icon('users', 19)}</div><div><h2>Workspace identity</h2><p>Current section</p><strong>${escapeHtml(state.workspace)}</strong><small>${apiMode ? 'Launch the matching batch file to switch between restaurant and healthcare databases.' : 'Switch sections from the sidebar. Food and healthcare leads remain separate.'}</small></div><select data-workspace ${apiMode ? 'disabled' : ''}><option ${state.workspace === 'Restaurants' ? 'selected' : ''}>Restaurants</option><option ${state.workspace === 'Healthcare' ? 'selected' : ''}>Healthcare</option></select></article><article class="panel settings-card"><div class="settings-card__icon settings-card__icon--green">${icon('download', 19)}</div><div><h2>Data portability</h2><p>Export a clean snapshot of this workspace</p><strong>${metrics.total} leads ready</strong><small>CSV includes contact data, scores, statuses, follow-ups and source fields.</small></div><button class="button button--secondary" data-export>Export CSV ${icon('arrow', 14)}</button></article><article class="panel settings-card"><div class="settings-card__icon settings-card__icon--amber">${icon('clock', 19)}</div><div><h2>Autosave</h2><p>${apiMode ? 'SQLite sync is active' : 'Browser storage is active'}</p><strong>Saved ${formatRelative(state.lastSavedAt)}</strong><small>${apiMode ? 'Every status, note and follow-up change is saved to the local CRM database.' : 'Every status, note and follow-up change is saved to this browser.'}</small></div><span class="settings-live"><i></i>Active</span></article><article class="panel settings-card settings-card--danger"><div class="settings-card__icon settings-card__icon--red">${icon('settings', 19)}</div><div><h2>${apiMode ? 'Database backups' : 'Reset demo workspace'}</h2><p>${apiMode ? 'Backups are kept beside the local database' : 'Restore the sample data and clear local edits'}</p><strong>${apiMode ? 'Protected locally' : 'This cannot be undone'}</strong><small>${apiMode ? 'The CRM refreshes a rolling database backup on startup.' : 'Use this when you want to review the first-run experience again.'}</small></div>${apiMode ? '<span class="settings-live"><i></i>Protected</span>' : '<button class="button button--danger" data-reset-demo>Reset demo</button>'}</article></section>`;
}

function emptyState(title, body, actionLabel = '', action = '') {
  const actionAttr = ['dashboard', 'leads', 'pipeline', 'followups', 'runs', 'settings'].includes(action)
    ? `data-view="${escapeHtml(action)}"`
    : action.startsWith('#')
      ? `data-view="${escapeHtml(action.slice(1))}"`
      : `data-action="${escapeHtml(action)}"`;
  return `<div class="empty-state"><div class="empty-state__icon">${icon('sparkle', 20)}</div><strong>${escapeHtml(title)}</strong><p>${escapeHtml(body)}</p>${actionLabel ? `<button class="button button--secondary" ${actionAttr}>${escapeHtml(actionLabel)} ${icon('arrow', 14)}</button>` : ''}</div>`;
}

function renderDrawer() {
  const drawer = document.querySelector('#lead-drawer');
  if (!drawer) return;
  const lead = state.selectedLeadId ? leadById(state.selectedLeadId) : null;
  if (!lead) {
    drawer.className = 'lead-drawer';
    drawer.setAttribute('aria-hidden', 'true');
    drawer.innerHTML = '';
    return;
  }
  const notes = lead.notes || [];
  drawer.className = 'lead-drawer is-open';
  drawer.setAttribute('aria-hidden', 'false');
  drawer.innerHTML = `<div class="drawer-scrim" data-close-drawer></div><div class="drawer-panel"><div class="drawer-header"><div><span class="drawer-kicker">Lead workspace</span><span class="drawer-source">${escapeHtml(lead.source || 'Manual entry')}</span></div><button class="icon-button" data-close-drawer aria-label="Close lead drawer">${icon('close', 19)}</button></div><div class="drawer-content"><div class="drawer-lead-heading">${initialsAvatar(lead.name, 'avatar--large')}<div><h2>${escapeHtml(lead.name)}</h2><p>${escapeHtml(lead.location)}</p><div class="drawer-tags">${tierBadge(lead.tier)} ${scoreBadge(lead.score)} ${statusPill(lead.status)}</div></div></div><div class="drawer-actions"><a class="button button--primary" href="tel:${phoneDigits(lead.phone)}">${icon('phone', 16)} Call</a><a class="button button--secondary" href="https://wa.me/91${phoneDigits(lead.phone)}" target="_blank" rel="noreferrer">WhatsApp</a></div><div class="drawer-section drawer-section--details"><div class="drawer-section__heading"><h3>Contact details</h3><span class="muted">Scanned signals</span></div><div class="contact-grid"><div><span>Contact</span><strong>${escapeHtml(lead.contact || 'Unknown')}</strong></div><div><span>Mobile</span><strong>${escapeHtml(lead.phone)}</strong></div><div><span>Email</span><strong>${lead.email ? `<a href="mailto:${escapeHtml(lead.email)}">${escapeHtml(lead.email)}</a>` : '<em>Not found</em>'}</strong></div><div><span>Website</span><strong>${lead.website ? `<a href="${escapeHtml(websiteHref(lead.website))}" target="_blank" rel="noreferrer">${escapeHtml(lead.website)} ${icon('external', 12)}</a>` : '<em>Not found</em>'}</strong></div></div></div><div class="drawer-section drawer-section--next"><div class="drawer-section__heading"><h3>Next step</h3><span class="muted">Update saves instantly</span></div><div class="drawer-form-row"><label>Status<select data-drawer-status data-lead-id="${escapeHtml(lead.id)}">${statusSelectOptions(lead.status, false)}</select></label><label>Follow-up<input type="date" data-drawer-followup data-lead-id="${escapeHtml(lead.id)}" value="${escapeHtml(lead.followup || '')}"></label></div><div class="quick-dates"><button class="text-button" data-set-followup="1" data-lead-id="${escapeHtml(lead.id)}">Tomorrow</button><button class="text-button" data-set-followup="3" data-lead-id="${escapeHtml(lead.id)}">+3 days</button><button class="text-button" data-set-followup="7" data-lead-id="${escapeHtml(lead.id)}">Next week</button><button class="text-button text-button--muted" data-clear-followup data-lead-id="${escapeHtml(lead.id)}">Clear date</button></div></div><div class="drawer-section"><div class="drawer-section__heading"><h3>Why this lead</h3><span class="muted">Score evidence</span></div><div class="signal-list">${(lead.signals || []).map((signal) => `<span>${icon('check', 13)}${escapeHtml(signal)}</span>`).join('')}</div><p class="angle-callout"><strong>Pitch angle:</strong> ${escapeHtml(lead.angle || 'New opportunity')}</p></div><div class="drawer-section drawer-section--notes"><div class="drawer-section__heading"><h3>Notes & activity</h3><span class="muted">${notes.length} ${notes.length === 1 ? 'entry' : 'entries'}</span></div><form data-note-form data-lead-id="${escapeHtml(lead.id)}"><textarea name="note" rows="3" placeholder="What happened? Add context for the next call…"></textarea><div class="note-form__footer"><span>Private to this workspace</span><button class="button button--secondary" type="submit">Add note ${icon('arrow', 14)}</button></div></form><div class="timeline">${notes.length ? notes.map((note) => `<div class="timeline-item"><span class="timeline-marker ${note.source === 'status' ? 'timeline-marker--status' : ''}">${icon(note.source === 'status' ? 'check' : 'note', 12)}</span><div><strong>${escapeHtml(note.text)}</strong><small>${formatDateTime(note.createdAt)} · ${escapeHtml(note.source === 'status' ? 'Pipeline update' : 'Note')}</small></div></div>`).join('') : '<div class="timeline-empty">No notes yet. Add the context you will want on the next call.</div>'}</div></div></div></div>`;
}

function openLead(id) {
  if (!leadById(id)) return;
  state.selectedLeadId = id;
  history.replaceState(null, '', `#lead/${encodeURIComponent(id)}`);
  renderDrawer();
  document.body.classList.add('drawer-open');
  window.setTimeout(() => document.querySelector('.drawer-panel [data-close-drawer]')?.focus(), 30);
}

function closeDrawer() {
  state.selectedLeadId = null;
  const fallback = state.view || 'dashboard';
  history.replaceState(null, '', `#${fallback}`);
  renderDrawer();
  document.body.classList.remove('drawer-open');
}

function showToast(message, tone = 'default') {
  window.clearTimeout(toastTimer);
  toastRoot.innerHTML = `<div class="toast toast--${tone}"><span class="toast__icon">${icon(tone === 'danger' ? 'close' : 'check', 15)}</span><span>${escapeHtml(message)}</span></div>`;
  window.requestAnimationFrame(() => toastRoot.querySelector('.toast')?.classList.add('is-visible'));
  toastTimer = window.setTimeout(() => toastRoot.querySelector('.toast')?.classList.remove('is-visible'), 3200);
}

function navigate(view, { push = true } = {}) {
  closeDialog();
  state.view = view;
  state.selectedLeadId = null;
  state.mobileNavOpen = false;
  if (push) history.pushState(null, '', `#${view}`);
  document.body.classList.remove('drawer-open');
  renderLayout();
}

function syncRoute() {
  closeDialog();
  const hash = window.location.hash.replace(/^#/, '') || 'dashboard';
  if (hash.startsWith('lead/')) {
    state.view = 'leads';
    state.selectedLeadId = decodeURIComponent(hash.slice(5));
  } else {
    state.view = ['dashboard', 'leads', 'pipeline', 'followups', 'runs', 'settings'].includes(hash) ? hash : 'dashboard';
    state.selectedLeadId = null;
  }
  renderLayout();
  if (state.selectedLeadId) document.body.classList.add('drawer-open');
}

function updateLeadStatus(id, status, followup = undefined) {
  const lead = leadById(id);
  if (!lead || !LEAD_STATUSES.includes(status)) return;
  const oldStatus = lead.status;
  state.leads = moveStatus(state.leads, id, status, followup);
  if (oldStatus !== status) {
    state.leads = addLeadNote(state.leads, id, `Status moved from ${oldStatus} to ${status}.`, 'status');
  }
  saveState();
  renderContent();
  renderDrawer();
  showToast(oldStatus === status ? 'Lead saved' : `${lead.name} moved to ${STATUS_META[status]?.label || status}`);
}

function updateLeadFollowup(id, followup) {
  const lead = leadById(id);
  if (!lead) return;
  state.leads = updateLead(state.leads, id, { followup });
  saveState();
  renderContent();
  renderDrawer();
  showToast(followup ? `Follow-up set for ${formatDate(followup)}` : 'Follow-up cleared');
}

function applyShortcut(shortcut) {
  state.filters = { status: 'all', tier: 'all', date: 'all' };
  if (shortcut === 'attention') state.filters.date = 'attention';
  if (shortcut === 'tierA') state.filters.tier = 'A';
  if (shortcut === 'new') state.filters.status = 'Not called';
  state.view = 'leads';
  history.pushState(null, '', '#leads');
  renderLayout();
}

function openDialog(kind, preset = {}) {
  const dialogRoot = document.querySelector('#dialog-root') || document.body.appendChild(document.createElement('div'));
  dialogRoot.id = 'dialog-root';
  const isImport = kind === 'import';
  const importCopy = apiMode ? 'Run the scraper from Command Prompt with vaani.py, then reopen the CRM. The database-backed workspace will show the new leads automatically.' : 'Bring a location search into the workspace. Demo mode adds a realistic sample run and keeps every new record local.';
  dialogRoot.innerHTML = `<dialog class="modal" data-modal><form method="dialog" class="modal-card" data-modal-form data-modal-kind="${kind}"><button type="button" class="icon-button modal-close" data-close-modal aria-label="Close">${icon('close', 18)}</button><div class="eyebrow"><span class="eyebrow-dot"></span>${isImport ? 'New data source' : 'Manual entry'}</div><h2>${isImport ? 'Import a fresh scrape' : 'Add a lead'}</h2><p>${isImport ? importCopy : 'Add one business directly when you have a lead that did not come from a scrape.'}</p>${isImport ? `<div class="modal-form-grid"><label>Location<input name="location" ${apiMode ? '' : 'required'} value="${escapeHtml(preset.location || '')}" placeholder="e.g. Prahlad Nagar, Ahmedabad"></label><label>Category<select name="category"><option>Restaurants</option><option>Delivery-first restaurants</option><option>Cafés</option><option>Dental clinics</option><option>Physiotherapy clinics</option></select></label><label>Workspace<select name="vertical"><option ${state.workspace === 'Restaurants' ? 'selected' : ''}>Restaurants</option><option ${state.workspace === 'Healthcare' ? 'selected' : ''}>Healthcare</option></select></label><label>Places to keep<select name="count"><option value="4">4 leads</option><option value="6" selected>6 leads</option><option value="8">8 leads</option></select></label></div><label class="check-field"><input type="checkbox" name="emails" checked><span><strong>Check websites for emails</strong><small>Add public contact signals to the imported record.</small></span></label>` : `<div class="modal-form-grid"><label>Business name<input name="name" required placeholder="The new lead"></label><label>Contact name<input name="contact" placeholder="Owner or decision-maker"></label><label>Mobile number<input name="phone" required placeholder="+91 98…"></label><label>Email<input name="email" type="email" placeholder="hello@business.in"></label><label>Location<input name="location" placeholder="Area, city"></label><label>Website<input name="website" placeholder="business.in"></label><label>Workspace<select name="vertical"><option ${state.workspace === 'Restaurants' ? 'selected' : ''}>Restaurants</option><option ${state.workspace === 'Healthcare' ? 'selected' : ''}>Healthcare</option></select></label><label>Tier<select name="tier"><option>A</option><option selected>B</option><option>C</option></select></label></div>`}<div class="modal-actions"><button type="button" class="button button--secondary" data-close-modal>Cancel</button><button class="button button--primary" type="submit">${isImport ? (apiMode ? 'Close' : 'Import leads') : 'Save lead'} ${icon('arrow', 15)}</button></div></form></dialog>`;
  const dialog = dialogRoot.querySelector('dialog');
  if (typeof dialog.showModal === 'function') dialog.showModal();
  else dialog.setAttribute('open', '');
  dialog.querySelector('input')?.focus();
}

function closeDialog() {
  const dialog = document.querySelector('[data-modal]');
  if (!dialog) return;
  if (typeof dialog.close === 'function') dialog.close();
  document.querySelector('#dialog-root')?.remove();
}

function handleModalSubmit(form) {
  const kind = form.dataset.modalKind;
  const values = Object.fromEntries(new FormData(form).entries());
  if (kind === 'import') {
    if (apiMode) {
      closeDialog();
      showToast('Run vaani.py from Command Prompt, then reopen the CRM to see new leads.');
      return;
    }
    const count = Number(values.count) || 6;
    const vertical = values.vertical || state.workspace;
    const imported = makeImportedLeads({ location: values.location, category: values.category, vertical, count });
    state.leads = [...imported, ...state.leads];
    state.runs = [{ id: `run-${Date.now()}`, vertical, location: values.location, category: values.category, createdAt: new Date().toISOString(), found: count + 8, saved: imported.length, contacts: imported.filter((lead) => lead.phone).length, status: 'Complete' }, ...state.runs];
    state.workspace = vertical;
    state.view = 'leads';
    state.filters = { status: 'all', tier: 'all', date: 'today' };
    state.query = '';
    saveState();
    closeDialog();
    history.pushState(null, '', '#leads');
    renderLayout();
    showToast(`${imported.length} new leads added from ${values.location}`);
    return;
  }
  const result = validateLead(values);
  if (!result.valid) {
    showToast(Object.values(result.errors)[0], 'danger');
    return;
  }
  state.leads = [createLead(values), ...state.leads];
  state.workspace = values.vertical || state.workspace;
  state.view = 'leads';
  state.filters = { status: 'all', tier: 'all', date: 'all' };
  saveState();
  closeDialog();
  history.pushState(null, '', '#leads');
  renderLayout();
  showToast(`${values.name} added to your workspace`);
}

function exportCsv() {
  const rows = filterLeads(state.leads, { workspace: state.workspace });
  const headers = ['Business', 'Contact', 'Phone', 'Email', 'Location', 'Website', 'Score', 'Tier', 'Status', 'Follow-up', 'Pitch angle', 'Source', 'First seen', 'Last seen'];
  const csv = [headers, ...rows.map((lead) => [lead.name, lead.contact, lead.phone, lead.email, lead.location, lead.website, lead.score, lead.tier, lead.status, lead.followup, lead.angle, lead.source, lead.firstSeen, lead.lastSeen])]
    .map((row) => row.map((value) => `"${String(value ?? '').replaceAll('"', '""')}"`).join(','))
    .join('\n');
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = `vaani-${state.workspace.toLowerCase()}-leads.csv`;
  link.click();
  URL.revokeObjectURL(url);
  showToast(`${rows.length} leads exported as CSV`);
}

function startCalling() {
  const queue = getMetrics(state.leads, state.workspace).dueLeads;
  const first = queue[0] || getVisibleLeads({ ...state, filters: { status: 'Not called', tier: 'all', date: 'all' } })[0];
  if (!first) {
    showToast('Your calling queue is clear. Import a new scrape to keep moving.');
    return;
  }
  openLead(first.id);
  showToast(`Starting with ${first.name}`);
}

function handleClick(event) {
  const viewTarget = event.target.closest('[data-view]');
  if (viewTarget) {
    event.preventDefault();
    const shortcut = viewTarget.dataset.filterShortcut;
    if (shortcut) applyShortcut(shortcut);
    else navigate(viewTarget.dataset.view);
    return;
  }
  const openTarget = event.target.closest('[data-open-lead]');
  if (openTarget) {
    event.preventDefault();
    openLead(openTarget.dataset.openLead);
    return;
  }
  if (event.target.closest('[data-close-drawer]')) {
    closeDrawer();
    return;
  }
  if (event.target.closest('[data-mobile-menu]')) {
    state.mobileNavOpen = !state.mobileNavOpen;
    document.querySelector('.app-frame')?.classList.toggle('nav-open', state.mobileNavOpen);
    return;
  }
  if (event.target.closest('[data-close-nav]')) {
    state.mobileNavOpen = false;
    document.querySelector('.app-frame')?.classList.remove('nav-open');
    return;
  }
  if (event.target.closest('[data-open-import]')) {
    openDialog('import');
    return;
  }
  if (event.target.closest('[data-open-add]')) {
    openDialog('add');
    return;
  }
  if (event.target.closest('[data-close-modal]')) {
    closeDialog();
    return;
  }
  if (event.target.closest('[data-export]')) {
    exportCsv();
    return;
  }
  if (event.target.closest('[data-start-calling]')) {
    startCalling();
    return;
  }
  const shortcut = event.target.closest('[data-filter-shortcut]');
  if (shortcut) {
    applyShortcut(shortcut.dataset.filterShortcut);
    return;
  }
  if (event.target.closest('[data-toggle-filters]')) {
    document.querySelector('[data-filter-panel]')?.classList.toggle('is-visible');
    return;
  }
  if (event.target.closest('[data-clear-filters]') || event.target.closest('[data-action="clear-filters"]')) {
    state.filters = { status: 'all', tier: 'all', date: 'all' };
    state.sort = 'score';
    state.query = '';
    saveState();
    renderContent();
    const search = document.querySelector('[data-global-search]');
    if (search) search.value = '';
    return;
  }
  if (event.target.closest('[data-clear-selection]')) {
    state.selectedIds = [];
    renderContent();
    return;
  }
  if (event.target.closest('[data-followup-tab]')) {
    state.followupTab = event.target.closest('[data-followup-tab]').dataset.followupTab;
    renderContent();
    return;
  }
  const setFollowup = event.target.closest('[data-set-followup]');
  if (setFollowup) {
    const date = new Date();
    date.setDate(date.getDate() + Number(setFollowup.dataset.setFollowup));
    updateLeadFollowup(setFollowup.dataset.leadId, date.toISOString().slice(0, 10));
    return;
  }
  const clearFollowup = event.target.closest('[data-clear-followup]');
  if (clearFollowup) {
    updateLeadFollowup(clearFollowup.dataset.leadId, '');
    return;
  }
  if (event.target.closest('[data-reset-demo]')) {
    if (apiMode) {
      showToast('This CRM is connected to the local SQLite database.');
      return;
    }
    if (window.confirm('Reset this workspace to the original demo data?')) {
      localStorage.removeItem(STORAGE_KEY);
      Object.assign(state, createInitialState());
      renderLayout();
      showToast('Demo workspace reset');
    }
    return;
  }
  const runAgain = event.target.closest('[data-run-again]');
  if (runAgain) {
    openDialog('import', { location: runAgain.dataset.runAgain });
    return;
  }
  const runReview = event.target.closest('[data-run-review]');
  if (runReview) {
    state.query = state.runs.find((run) => run.id === runReview.dataset.runReview)?.location || '';
    state.view = 'leads';
    state.filters = { status: 'all', tier: 'all', date: 'all' };
    history.pushState(null, '', '#leads');
    renderLayout();
  }
}

function handleChange(event) {
  const target = event.target;
  if (target.matches('[data-workspace]')) {
    state.workspace = target.value;
    state.selectedIds = [];
    state.filters = { status: 'all', tier: 'all', date: 'all' };
    state.query = '';
    saveState();
    renderLayout();
    showToast(`${state.workspace} workspace selected`);
    return;
  }
  if (target.matches('[data-status-filter]')) {
    state.filters.status = target.value;
    renderContent();
    return;
  }
  if (target.matches('[data-tier-filter]')) {
    state.filters.tier = target.value;
    renderContent();
    return;
  }
  if (target.matches('[data-date-filter]')) {
    state.filters.date = target.value;
    renderContent();
    return;
  }
  if (target.matches('[data-sort]')) {
    state.sort = target.value;
    renderContent();
    return;
  }
  if (target.matches('[data-select-lead]')) {
    const id = target.dataset.selectLead;
    if (target.checked) state.selectedIds = [...new Set([...state.selectedIds, id])];
    else state.selectedIds = state.selectedIds.filter((selectedId) => selectedId !== id);
    renderContent();
    return;
  }
  if (target.matches('[data-bulk-status]') && target.value) {
    const updatedCount = state.selectedIds.length;
    state.selectedIds.forEach((id) => {
      const lead = leadById(id);
      if (lead) state.leads = moveStatus(state.leads, id, target.value);
    });
    state.leads = state.selectedIds.reduce((leads, id) => addLeadNote(leads, id, `Bulk status update: ${target.value}.`, 'status'), state.leads);
    state.selectedIds = [];
    saveState();
    renderContent();
    showToast(`Updated ${updatedCount} lead${updatedCount === 1 ? '' : 's'}`);
    return;
  }
  if (target.matches('[data-drawer-status]')) {
    updateLeadStatus(target.dataset.leadId, target.value);
    return;
  }
  if (target.matches('[data-drawer-followup]')) {
    updateLeadFollowup(target.dataset.leadId, target.value);
  }
}

function handleInput(event) {
  if (!event.target.matches('[data-global-search]')) return;
  state.query = event.target.value;
  window.clearTimeout(searchTimer);
  searchTimer = window.setTimeout(() => {
    if (state.view !== 'leads') {
      state.view = 'leads';
      history.replaceState(null, '', '#leads');
      renderLayout();
    } else {
      renderContent();
      document.querySelector('[data-global-search]')?.focus();
    }
  }, 120);
}

function handleSubmit(event) {
  const form = event.target;
  if (form.matches('[data-modal-form]')) {
    event.preventDefault();
    handleModalSubmit(form);
    return;
  }
  if (form.matches('[data-note-form]')) {
    event.preventDefault();
    const text = new FormData(form).get('note')?.toString().trim();
    if (!text) {
      showToast('Write a note before saving it.', 'danger');
      form.querySelector('textarea')?.focus();
      return;
    }
    state.leads = addLeadNote(state.leads, form.dataset.leadId, text);
    saveState();
    renderContent();
    renderDrawer();
    showToast('Note added to the lead');
  }
}

function handleKeydown(event) {
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') {
    event.preventDefault();
    document.querySelector('[data-global-search]')?.focus();
  }
  if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
    event.preventDefault();
    startCalling();
  }
  if (event.key === 'Escape') {
    if (document.querySelector('[data-modal]')) closeDialog();
    else if (state.selectedLeadId) closeDrawer();
  }
}

function handleDragStart(event) {
  const card = event.target.closest('[data-drag-lead]');
  if (!card) return;
  event.dataTransfer.effectAllowed = 'move';
  event.dataTransfer.setData('text/plain', card.dataset.dragLead);
  card.classList.add('is-dragging');
}

function handleDragEnd(event) {
  event.target.closest('[data-drag-lead]')?.classList.remove('is-dragging');
  document.querySelectorAll('.pipeline-column.is-over').forEach((column) => column.classList.remove('is-over'));
}

function handleDragOver(event) {
  const column = event.target.closest('[data-stage]');
  if (!column) return;
  event.preventDefault();
  column.classList.add('is-over');
}

function handleDrop(event) {
  const column = event.target.closest('[data-stage]');
  if (!column) return;
  event.preventDefault();
  column.classList.remove('is-over');
  const id = event.dataTransfer.getData('text/plain');
  if (id) updateLeadStatus(id, column.dataset.status);
}

async function init() {
  let apiError = null;
  if (apiMode) {
    try {
      await hydrateFromApi();
    } catch (error) {
      apiError = error;
    }
  }
  syncRoute();
  document.addEventListener('click', handleClick);
  document.addEventListener('change', handleChange);
  document.addEventListener('input', handleInput);
  document.addEventListener('submit', handleSubmit);
  document.addEventListener('keydown', handleKeydown);
  document.addEventListener('dragstart', handleDragStart);
  document.addEventListener('dragend', handleDragEnd);
  document.addEventListener('dragover', handleDragOver);
  document.addEventListener('drop', handleDrop);
  window.addEventListener('popstate', syncRoute);
  window.addEventListener('hashchange', syncRoute);
  window.__VAANI_CRM__ = { state, getMetrics, filterLeads, createLead, validateLead };
  if (apiError) showToast('Could not load the local CRM database.', 'danger');
}

init();
