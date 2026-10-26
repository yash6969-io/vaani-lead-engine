export const STORAGE_KEY = 'vaani-crm-workspace-v2';

export const LEAD_STATUSES = [
  'Not called', 'No answer', 'Staff reached', 'Decision-maker reached',
  'Conversation', 'Interested', 'Demo booked', 'Demo done', 'Deck sent',
  'Callback', 'Won', 'Lost', 'Not interested', 'Wrong number', 'Do not contact',
];

export const PIPELINE_STAGES = [
  { id: 'new', label: 'New', statuses: ['Not called'] },
  { id: 'contacted', label: 'Contacted', statuses: ['No answer', 'Staff reached', 'Decision-maker reached', 'Conversation'] },
  { id: 'interested', label: 'Interested', statuses: ['Interested'] },
  { id: 'demo', label: 'Demo', statuses: ['Demo booked', 'Demo done'] },
  { id: 'won', label: 'Won', statuses: ['Won'] },
  { id: 'lost', label: 'Lost', statuses: ['Lost', 'Not interested', 'Wrong number', 'Do not contact'] },
];

export const STATUS_META = {
  'Not called': { tone: 'neutral', label: 'Not called' },
  'No answer': { tone: 'slate', label: 'No answer' },
  'Staff reached': { tone: 'blue', label: 'Staff reached' },
  'Decision-maker reached': { tone: 'blue', label: 'Decision-maker' },
  Conversation: { tone: 'blue', label: 'Conversation' },
  Interested: { tone: 'amber', label: 'Interested' },
  'Demo booked': { tone: 'violet', label: 'Demo booked' },
  'Demo done': { tone: 'violet', label: 'Demo done' },
  'Deck sent': { tone: 'violet', label: 'Deck sent' },
  Callback: { tone: 'amber', label: 'Callback' },
  Won: { tone: 'green', label: 'Won' },
  Lost: { tone: 'red', label: 'Lost' },
  'Not interested': { tone: 'red', label: 'Not interested' },
  'Wrong number': { tone: 'red', label: 'Wrong number' },
  'Do not contact': { tone: 'dark', label: 'Do not contact' },
};

const localIso = (date) => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const daysFromToday = (offset) => {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  date.setDate(date.getDate() + offset);
  return localIso(date);
};

const note = (id, text, ageHours = 1, source = 'call') => ({
  id, text, source, createdAt: new Date(Date.now() - ageHours * 3600000).toISOString(),
});

export function seedLeads() {
  return [
    { id: 'lead-mango-tree', vertical: 'Restaurants', name: 'Mango Tree Kitchen', contact: 'Rohan Mehta', location: 'Vastrapur, Ahmedabad', phone: '+91 98250 18420', email: 'hello@mangotreekitchen.in', website: 'mangotreekitchen.in', rating: 4.6, reviews: 284, score: 88, tier: 'A', angle: 'Aggregator-dependent', status: 'Interested', followup: daysFromToday(0), source: 'Google Maps · 26 Oct', firstSeen: daysFromToday(-2), lastSeen: daysFromToday(0), timesSeen: 3, signals: ['Swiggy + Zomato', 'No direct ordering', '4.6 rating'], notes: [note('mango-1', 'Spoke with Rohan. Wants to see commission savings before the weekend.') ] },
    { id: 'lead-saffron-square', vertical: 'Restaurants', name: 'Saffron Square', contact: 'Neha Shah', location: 'Prahlad Nagar, Ahmedabad', phone: '+91 98980 44122', email: 'neha@saffronsquare.com', website: 'saffronsquare.com', rating: 4.4, reviews: 611, score: 84, tier: 'A', angle: 'Reputation-active', status: 'Demo booked', followup: daysFromToday(1), source: 'Google Maps · 25 Oct', firstSeen: daysFromToday(-3), lastSeen: daysFromToday(-1), timesSeen: 4, signals: ['611 reviews', 'Instagram active'], notes: [note('saffron-1', 'Demo booked for tomorrow at 11:30 AM.', 5, 'status')] },
    { id: 'lead-olive-table', vertical: 'Restaurants', name: 'The Olive Table', contact: 'Aarav Patel', location: 'Bodakdev, Ahmedabad', phone: '+91 90331 76548', email: '', website: '', rating: 4.2, reviews: 97, score: 76, tier: 'A', angle: 'Digitally weak - no website', status: 'Not called', followup: '', source: 'Google Maps · 26 Oct', firstSeen: daysFromToday(0), lastSeen: daysFromToday(0), timesSeen: 1, signals: ['97 reviews', 'No website'], notes: [] },
    { id: 'lead-dosa-district', vertical: 'Restaurants', name: 'Dosa District', contact: 'Kunal Desai', location: 'Satellite, Ahmedabad', phone: '+91 97274 11808', email: 'team@dosadistrict.co', website: 'dosadistrict.co', rating: 4.1, reviews: 431, score: 73, tier: 'B', angle: 'No direct ordering', status: 'No answer', followup: daysFromToday(-1), source: 'Google Maps · 24 Oct', firstSeen: daysFromToday(-5), lastSeen: daysFromToday(-2), timesSeen: 2, signals: ['Delivery + takeaway', 'Email found'], notes: [note('dosa-1', 'Called at 4:10 PM. Try again before lunch service.', 24)] },
    { id: 'lead-aurora-dental', vertical: 'Healthcare', name: 'Aurora Dental Studio', contact: 'Dr. Kavya Rao', location: 'C.G. Road, Ahmedabad', phone: '+91 98254 11229', email: 'care@auroradental.in', website: 'auroradental.in', rating: 4.8, reviews: 212, score: 86, tier: 'A', angle: 'Reputation-active', status: 'Interested', followup: daysFromToday(1), source: 'Google Maps · 25 Oct', firstSeen: daysFromToday(-4), lastSeen: daysFromToday(-1), timesSeen: 3, signals: ['212 reviews', 'Website found'], notes: [note('aurora-1', 'Doctor wants examples of appointment content.', 4)] },
    { id: 'lead-peak-physio', vertical: 'Healthcare', name: 'Peak Physio Care', contact: 'Sameer Bhat', location: 'Satellite, Ahmedabad', phone: '+91 98983 91220', email: '', website: 'peakphysio.care', rating: 4.4, reviews: 76, score: 68, tier: 'B', angle: 'Digitally weak - no website', status: 'Not called', followup: '', source: 'Google Maps · 26 Oct', firstSeen: daysFromToday(0), lastSeen: daysFromToday(0), timesSeen: 1, signals: ['76 reviews', 'Phone verified'], notes: [] },
  ];
}

export function seedRuns() {
  return [
    { id: 'run-1026', vertical: 'Restaurants', location: 'Ahmedabad', category: 'Restaurants', createdAt: new Date(Date.now() - 25 * 60000).toISOString(), found: 42, saved: 18, contacts: 16, status: 'Complete' },
    { id: 'run-1025', vertical: 'Restaurants', location: 'Prahlad Nagar, Ahmedabad', category: 'Delivery-first restaurants', createdAt: new Date(Date.now() - 24 * 3600000).toISOString(), found: 27, saved: 11, contacts: 10, status: 'Complete' },
    { id: 'run-1024', vertical: 'Healthcare', location: 'C.G. Road, Ahmedabad', category: 'Dental clinics', createdAt: new Date(Date.now() - 48 * 3600000).toISOString(), found: 19, saved: 9, contacts: 8, status: 'Complete' },
  ];
}

export function createInitialState() {
  return { view: 'dashboard', workspace: 'Restaurants', workspaces: [{ key: 'restaurants', label: 'Restaurants' }, { key: 'healthcare', label: 'Healthcare' }], leads: seedLeads(), runs: seedRuns(), selectedLeadId: null, query: '', filters: { status: 'all', tier: 'all', date: 'all' }, sort: 'score', selectedIds: [], followupTab: 'today', mobileNavOpen: false, lastSavedAt: new Date().toISOString() };
}

export function filterLeads(leads, { workspace = 'Restaurants', query = '', status = 'all', tier = 'all', date = 'all' } = {}) {
  const q = query.trim().toLowerCase();
  return leads.filter((lead) => {
    if (lead.vertical !== workspace) return false;
    if (status !== 'all' && lead.status !== status) return false;
    if (tier !== 'all' && lead.tier !== tier) return false;
    if (date === 'today' && lead.firstSeen !== daysFromToday(0)) return false;
    if (date === '7d' && lead.firstSeen < daysFromToday(-6)) return false;
    if (date === 'attention' && !isFollowupDue(lead)) return false;
    if (!q) return true;
    return [lead.name, lead.contact, lead.location, lead.phone, lead.email, lead.angle, lead.status, lead.source, ...(lead.signals || [])].join(' ').toLowerCase().includes(q);
  });
}

export function getVisibleLeads(state) {
  const rows = filterLeads(state.leads, { workspace: state.workspace, query: state.query, ...state.filters });
  return [...rows].sort((a, b) => state.sort === 'name' ? a.name.localeCompare(b.name) : state.sort === 'followup' ? (a.followup || '9999').localeCompare(b.followup || '9999') : state.sort === 'recent' ? b.lastSeen.localeCompare(a.lastSeen) : state.sort === 'newest' ? b.firstSeen.localeCompare(a.firstSeen) : (b.score || 0) - (a.score || 0));
}

export function isFollowupDue(lead, today = daysFromToday(0)) {
  return Boolean(lead.followup && lead.followup <= today && !['Won', 'Lost', 'Not interested', 'Wrong number', 'Do not contact'].includes(lead.status));
}

export function getMetrics(leads, workspace = 'Restaurants') {
  const scoped = leads.filter((lead) => lead.vertical === workspace);
  const closedStatuses = ['Won', 'Lost', 'Not interested', 'Wrong number', 'Do not contact'];
  const won = scoped.filter((lead) => lead.status === 'Won').length;
  const closed = scoped.filter((lead) => ['Won', 'Lost', 'Not interested'].includes(lead.status)).length;
  const tiers = { A: scoped.filter((lead) => lead.tier === 'A').length, B: scoped.filter((lead) => lead.tier === 'B').length, C: scoped.filter((lead) => lead.tier === 'C').length };
  const dueLeads = scoped.filter((lead) => isFollowupDue(lead)).sort((a, b) => (a.followup || '9999').localeCompare(b.followup || '9999'));
  const contacted = scoped.filter((lead) => lead.status !== 'Not called').length;
  return { total: scoped.length, active: scoped.filter((lead) => !closedStatuses.includes(lead.status)).length, due: dueLeads.length, overdue: dueLeads.filter((lead) => lead.followup < daysFromToday(0)).length, won, closed, demos: scoped.filter((lead) => ['Demo booked', 'Demo done'].includes(lead.status)).length, contacted, contactRate: scoped.length ? Math.round(contacted * 100 / scoped.length) : 0, winRate: closed ? Math.round(won * 100 / closed) : 0, avgScore: scoped.length ? Math.round(scoped.reduce((sum, lead) => sum + (Number(lead.score) || 0), 0) / scoped.length) : 0, tiers, dueLeads };
}

export function formatDate(value, options = {}) { if (!value) return '—'; const d = new Date(`${value}T12:00:00`); return Number.isNaN(d.valueOf()) ? value : new Intl.DateTimeFormat('en-IN', { day: 'numeric', month: 'short', ...options }).format(d); }
export function formatDateTime(value) { if (!value) return '—'; const d = new Date(value); return Number.isNaN(d.valueOf()) ? value : new Intl.DateTimeFormat('en-IN', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' }).format(d); }
export function formatRelative(value) { if (!value) return '—'; const diff = Math.round((Date.now() - new Date(value).valueOf()) / 60000); if (diff < 1) return 'Just now'; if (diff < 60) return `${diff}m ago`; if (diff < 1440) return `${Math.round(diff / 60)}h ago`; if (diff < 10080) return `${Math.round(diff / 1440)}d ago`; return formatDate(value); }
export function initials(name = '') { return name.split(/\s+/).filter(Boolean).slice(0, 2).map((part) => part[0]).join('').toUpperCase() || '??'; }

export function validateLead(values) {
  const errors = {};
  if (!String(values.name || '').trim()) errors.name = 'Add a business name.';
  if (!String(values.phone || '').trim()) errors.phone = 'Add a mobile number.';
  if (values.email && !/^\S+@\S+\.\S+$/.test(values.email)) errors.email = 'Use a valid email.';
  return { valid: Object.keys(errors).length === 0, errors };
}

export function createLead(values, now = new Date()) {
  return { id: `lead-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, vertical: values.vertical || 'Restaurants', name: String(values.name || '').trim(), contact: String(values.contact || '').trim(), location: String(values.location || '').trim(), phone: String(values.phone || '').trim(), email: String(values.email || '').trim(), website: String(values.website || '').trim(), rating: Number(values.rating) || 0, reviews: Number(values.reviews) || 0, score: Number(values.score) || 50, tier: values.tier || 'B', angle: values.angle || 'New opportunity', status: values.status || 'Not called', followup: values.followup || '', source: values.source || 'Manual entry', firstSeen: localIso(now), lastSeen: localIso(now), timesSeen: 1, signals: values.signals || ['Manually added'], notes: [] };
}

export function updateLead(leads, id, patch) { return leads.map((lead) => lead.id === id ? { ...lead, ...patch } : lead); }
export function addLeadNote(leads, id, text, source = 'note', now = new Date()) { const n = { id: `note-${Date.now()}`, text: String(text || '').trim(), source, createdAt: now.toISOString() }; return leads.map((lead) => lead.id === id ? { ...lead, notes: [n, ...(lead.notes || [])] } : lead); }
export function moveStatus(leads, id, status, followup = undefined) { return leads.map((lead) => lead.id === id ? { ...lead, status, ...(followup === undefined ? {} : { followup }), lastSeen: daysFromToday(0) } : lead); }
export function makeImportedLeads({ location = 'Ahmedabad', category = 'Restaurants', vertical = 'Restaurants', count = 4 } = {}) { return Array.from({ length: Math.max(1, Math.min(8, count)) }, (_, i) => createLead({ name: `${category} lead ${i + 1}`, location, phone: `+91 98${String(32000000 + i * 11321).slice(0, 8)}`, tier: i < 2 ? 'A' : 'B', score: 60 + i * 5, angle: 'New opportunity', source: `${category} · ${formatDate(daysFromToday(0))}`, vertical })); }
export const todayIso = () => daysFromToday(0);
