/* ── Taxonomy & state ──────────────────────────────────────────────────── */
const FIELD_MAP = {
  industry: 'Industry Focus',
  style:    'Investment style',
  mcap:     'Market capitalization',
  geo:      'Geography',
};

const selected = { industry: new Set(), style: new Set(), mcap: new Set(), geo: new Set() };
const files    = { contacts: null, ownership: null, fund: null, docs: [] };

/* ── Build CDF pickers ─────────────────────────────────────────────────── */
function buildPickers() {
  const tax = window.TAXONOMY || {};

  buildPicker('industry', tax[FIELD_MAP.industry] || [], true);
  buildPicker('style',    tax[FIELD_MAP.style]    || [], false);
  buildPicker('mcap',     tax[FIELD_MAP.mcap]     || [], false);
  buildPicker('geo',      tax[FIELD_MAP.geo]       || [], false);
}

function buildPicker(dim, items, grouped) {
  const list = document.getElementById(`list-${dim}`);
  list.innerHTML = '';

  if (!grouped) {
    items.forEach(item => list.appendChild(makeItem(dim, item.value, item.description)));
    return;
  }

  // Group by prefix before ':'
  const groups = {};
  const toplevel = [];
  items.forEach(item => {
    const colon = item.value.indexOf(':');
    if (colon === -1 || item.value.startsWith('*') || item.value === 'Agriculture' ||
        item.value === 'Infrastructure' || item.value === 'Packaging' || item.value === 'Utilities') {
      toplevel.push(item);
    } else {
      const prefix = item.value.substring(0, colon).trim();
      if (!groups[prefix]) groups[prefix] = [];
      groups[prefix].push(item);
    }
  });

  // Thematic group
  const thematic = [];
  const nonThematic = [];
  toplevel.forEach(item => {
    if (item.value.startsWith('Thematic')) thematic.push(item);
    else nonThematic.push(item);
  });

  nonThematic.forEach(item => {
    list.appendChild(makeItem(dim, item.value, item.description));
    if (groups[item.value]) {
      groups[item.value].forEach(sub => list.appendChild(makeItem(dim, sub.value, sub.description, true)));
    }
  });

  if (thematic.length) {
    const hdr = document.createElement('div');
    hdr.className = 'cdf-group-header';
    hdr.textContent = 'Thematic';
    list.appendChild(hdr);
    thematic.forEach(item => list.appendChild(makeItem(dim, item.value, item.description)));
  }
}

function makeItem(dim, value, description, indent) {
  const el = document.createElement('div');
  el.className = 'cdf-item';
  el.dataset.value = value;
  el.dataset.dim = dim;
  if (indent) el.style.paddingLeft = '26px';

  const box = document.createElement('div');
  box.className = 'cdf-checkbox';

  const label = document.createElement('div');
  label.className = 'cdf-item-text';
  label.textContent = value;

  el.appendChild(box);
  el.appendChild(label);

  if (description) {
    const tip = document.createElement('div');
    tip.className = 'cdf-tooltip';
    tip.textContent = description;
    el.appendChild(tip);
  }

  el.addEventListener('click', () => toggleItem(dim, value, el));
  return el;
}

function toggleItem(dim, value, el) {
  if (selected[dim].has(value)) {
    selected[dim].delete(value);
    el.classList.remove('selected');
  } else {
    selected[dim].add(value);
    el.classList.add('selected');
  }
  updateCount(dim);
  updateRunButton();
}

function updateCount(dim) {
  const n = selected[dim].size;
  document.getElementById(`count-${dim}`).textContent =
    n === 0 ? '0 selected' : `${n} selected`;
}

function setSelected(dim, values) {
  selected[dim].clear();
  values.forEach(v => selected[dim].add(v));

  const list = document.getElementById(`list-${dim}`);
  list.querySelectorAll('.cdf-item').forEach(el => {
    if (selected[dim].has(el.dataset.value)) el.classList.add('selected');
    else el.classList.remove('selected');
  });
  updateCount(dim);
}

/* ── Search ────────────────────────────────────────────────────────────── */
document.addEventListener('input', e => {
  if (!e.target.classList.contains('search-input')) return;
  const dim  = e.target.dataset.target;
  const term = e.target.value.toLowerCase();
  const list = document.getElementById(`list-${dim}`);

  list.querySelectorAll('.cdf-item').forEach(el => {
    const match = el.dataset.value.toLowerCase().includes(term);
    el.classList.toggle('hidden-item', !match);
  });
  list.querySelectorAll('.cdf-group-header').forEach(hdr => {
    const next = hdr.nextElementSibling;
    hdr.style.display = (!term) ? '' : 'none';
  });
});

/* ── File uploads ──────────────────────────────────────────────────────── */
function setupFileInput(inputId, fileKey, statusId) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener('change', () => {
    const f = input.files[0];
    if (!f) return;
    files[fileKey] = f;
    const status = document.getElementById(statusId);
    status.textContent = '✓ ' + truncate(f.name, 18);
    updateRunButton();
    updateAIButton();
  });
}

setupFileInput('input-contacts',  'contacts',  'status-contacts');
setupFileInput('input-ownership', 'ownership', 'status-ownership');
setupFileInput('input-fund',      'fund',      'status-fund');

document.getElementById('input-docs').addEventListener('change', function() {
  files.docs = Array.from(this.files);
  const status = document.getElementById('status-docs');
  status.textContent = files.docs.length
    ? `✓ ${files.docs.length} file${files.docs.length > 1 ? 's' : ''}`
    : '';
  updateAIButton();
});

function truncate(str, max) { return str.length > max ? str.slice(0, max) + '…' : str; }

function updateRunButton() {
  document.getElementById('run-btn').disabled = !files.contacts;
}

function updateAIButton() {
  document.getElementById('ai-btn').disabled = files.docs.length === 0;
}

/* ── AI analysis ───────────────────────────────────────────────────────── */
document.getElementById('ai-btn').addEventListener('click', async () => {
  if (!files.docs.length) return;
  const banner = document.getElementById('ai-banner');
  const loader = document.getElementById('ai-loader');
  const bannerText = document.getElementById('ai-banner-text');

  banner.classList.remove('hidden');
  loader.style.display = 'block';
  bannerText.textContent = 'Analyzing documents…';

  const form = new FormData();
  files.docs.forEach(f => form.append('documents', f));

  try {
    const resp = await fetch('/api/analyze', { method: 'POST', body: form });
    const data = await resp.json();

    if (data.error) throw new Error(data.error);

    if (data.industry) setSelected('industry', data.industry);
    if (data.style)    setSelected('style',    data.style);
    if (data.mcap)     setSelected('mcap',     data.mcap);
    if (data.geo)      setSelected('geo',      data.geo);

    loader.style.display = 'none';
    bannerText.textContent = '✓ AI recommendations applied — review and adjust before running';
    updateRunButton();
  } catch (err) {
    loader.style.display = 'none';
    bannerText.textContent = `Analysis failed: ${err.message}`;
    banner.style.background = 'var(--danger-bg)';
    banner.style.color = 'var(--danger)';
    banner.style.borderColor = 'rgba(192,57,43,0.2)';
  }
});

/* ── Run filter ────────────────────────────────────────────────────────── */
document.getElementById('run-btn').addEventListener('click', async () => {
  const btn = document.getElementById('run-btn');
  btn.classList.add('loading');
  btn.querySelector('.run-label').textContent = 'Running…';
  btn.disabled = true;

  hideResults();
  hideError();

  const form = new FormData();
  form.append('contacts',     files.contacts);
  if (files.ownership) form.append('ownership', files.ownership);
  if (files.fund)      form.append('fund_ownership', files.fund);

  selected.industry.forEach(v => form.append('industry', v));
  selected.style.forEach(v    => form.append('style', v));
  selected.mcap.forEach(v     => form.append('mcap', v));
  selected.geo.forEach(v      => form.append('geo', v));

  form.append('hf_treatment', document.querySelector('input[name=hf_treatment]:checked').value);
  form.append('company_name', document.getElementById('company-name').value.trim());
  form.append('ticker',       document.getElementById('ticker').value.trim().toUpperCase());

  try {
    const resp = await fetch('/api/run', { method: 'POST', body: form });
    const data = await resp.json();

    if (data.error) throw new Error(data.error);

    showResults(data);
  } catch (err) {
    showError(err.message);
  } finally {
    btn.classList.remove('loading');
    btn.querySelector('.run-label').textContent = 'Run filter';
    btn.disabled = false;
  }
});

/* ── Results ───────────────────────────────────────────────────────────── */
function showResults(data) {
  document.getElementById('stat-source').textContent  = fmt(data.total_source);
  document.getElementById('stat-matched').textContent = fmt(data.main_count);
  document.getElementById('stat-hf').textContent      = fmt(data.hf_count);
  document.getElementById('stat-dnc').textContent     = fmt(data.dnc_count);
  document.getElementById('stat-check').textContent   = fmt(data.check_count);
  document.getElementById('stat-quant').textContent   = fmt(data.quant_count);

  // Breakdown bars
  const barsEl = document.getElementById('breakdown-bars');
  barsEl.innerHTML = '';
  const breakdown = data.match_breakdown || {};
  const max = Math.max(...Object.values(breakdown), 1);
  const labels = { 1: '1 criterion', 2: '2 criteria', 3: '3 criteria', 4: '4 criteria' };
  [4, 3, 2, 1].forEach(k => {
    const n = breakdown[k] || 0;
    const row = document.createElement('div');
    row.className = 'breakdown-row';
    row.innerHTML = `
      <span class="breakdown-key">${labels[k] || k + ' criteria'}</span>
      <div class="breakdown-bar-wrap">
        <div class="breakdown-bar" style="width:${Math.round(n/max*100)}%"></div>
      </div>
      <span class="breakdown-val">${fmt(n)}</span>`;
    barsEl.appendChild(row);
  });

  // SharePoint link
  const spBtn = document.getElementById('sharepoint-btn');
  if (data.sharepoint_url) {
    spBtn.href = data.sharepoint_url;
    spBtn.classList.remove('hidden');
  } else {
    spBtn.classList.add('hidden');
  }

  document.getElementById('results-panel').classList.remove('hidden');
  document.getElementById('results-panel').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function hideResults() { document.getElementById('results-panel').classList.add('hidden'); }
function showError(msg) {
  const el = document.getElementById('error-panel');
  document.getElementById('error-text').textContent = msg;
  el.classList.remove('hidden');
}
function hideError() { document.getElementById('error-panel').classList.add('hidden'); }
function fmt(n) { return (n || 0).toLocaleString(); }

/* ── Init ──────────────────────────────────────────────────────────────── */
buildPickers();
