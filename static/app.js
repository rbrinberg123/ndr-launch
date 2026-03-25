/* ── NDR Launch app.js ─────────────────────────────────────────────────────── */

const TAX    = window.TAXONOMY       || {};
const CM_ICS = window.CITY_MAP_ICS   || [];
const CM_STS = window.CITY_MAP_STATES || [];

// ── CDF list rendering ────────────────────────────────────────────────────────

const DIM_MAP = {
  industry: 'Industry Focus',
  style:    'Investment style',
  mcap:     'Market capitalization',
  geo:      'Geography',
};

const selections = { industry: new Set(), style: new Set(), mcap: new Set(), geo: new Set() };

function renderCDFList(dim) {
  const taxKey  = DIM_MAP[dim];
  const items   = (TAX[taxKey] || []).map(i => i.value);
  const listEl  = document.getElementById(`list-${dim}`);
  const countEl = document.getElementById(`count-${dim}`);
  if (!listEl) return;

  listEl.innerHTML = '';

  if (dim === 'industry') {
    let lastGroup = null;
    items.forEach(val => {
      const group = val.includes(':') ? val.split(':')[0].trim() : (val.startsWith('*') ? 'Special' : val);
      if (group !== lastGroup) {
        const hdr = document.createElement('div');
        hdr.className = 'cdf-group-header';
        hdr.textContent = group;
        listEl.appendChild(hdr);
        lastGroup = group;
      }
      listEl.appendChild(makeItem(dim, val));
    });
  } else {
    items.forEach(val => listEl.appendChild(makeItem(dim, val)));
  }

  updateCount(dim, countEl);
}

function makeItem(dim, val) {
  const el = document.createElement('div');
  el.className = 'cdf-item' + (selections[dim].has(val) ? ' selected' : '');
  el.dataset.value = val;

  const cb  = document.createElement('div'); cb.className = 'cdf-checkbox';
  const txt = document.createElement('div'); txt.className = 'cdf-item-text'; txt.textContent = val;
  el.appendChild(cb); el.appendChild(txt);

  el.addEventListener('click', () => {
    if (selections[dim].has(val)) selections[dim].delete(val);
    else selections[dim].add(val);
    el.classList.toggle('selected', selections[dim].has(val));
    updateCount(dim, document.getElementById(`count-${dim}`));
    updateRunButton();
  });
  return el;
}

function updateCount(dim, el) {
  if (!el) return;
  const n = selections[dim].size;
  el.textContent = n === 0 ? 'All (skip)' : `${n} selected`;
  el.style.color = n > 0 ? 'var(--accent)' : 'var(--text-muted)';
}

// ── Select all / Clear per dimension ─────────────────────────────────────────

function selectAllDim(dim) {
  const taxKey = DIM_MAP[dim];
  const items  = (TAX[taxKey] || []).map(i => i.value);
  items.forEach(v => selections[dim].add(v));
  document.querySelectorAll(`#list-${dim} .cdf-item`).forEach(el => {
    el.classList.add('selected');
  });
  updateCount(dim, document.getElementById(`count-${dim}`));
  updateRunButton();
}

function clearDim(dim) {
  selections[dim].clear();
  document.querySelectorAll(`#list-${dim} .cdf-item`).forEach(el => {
    el.classList.remove('selected');
  });
  updateCount(dim, document.getElementById(`count-${dim}`));
  updateRunButton();
}

// ── Search filtering ──────────────────────────────────────────────────────────

document.querySelectorAll('.search-input').forEach(inp => {
  inp.addEventListener('input', () => {
    const dim   = inp.dataset.target;
    const query = inp.value.toLowerCase().trim();
    const list  = document.getElementById(`list-${dim}`);
    if (!list) return;
    list.querySelectorAll('.cdf-item').forEach(item => {
      item.classList.toggle('hidden-item', !item.dataset.value.toLowerCase().includes(query));
    });
    list.querySelectorAll('.cdf-group-header').forEach(hdr => {
      let next = hdr.nextElementSibling;
      let anyVisible = false;
      while (next && !next.classList.contains('cdf-group-header')) {
        if (!next.classList.contains('hidden-item')) anyVisible = true;
        next = next.nextElementSibling;
      }
      hdr.style.display = anyVisible ? '' : 'none';
    });
  });
});

// ── Accordion ────────────────────────────────────────────────────────────────

function toggleAcc(id) {
  const body  = document.getElementById(`acc-body-${id}`);
  const arrow = document.getElementById(`acc-arrow-${id}`);
  const open  = body.classList.contains('open');
  body.classList.toggle('open', !open);
  if (arrow) arrow.classList.toggle('open', !open);
}

// ── IC / State list rendering ─────────────────────────────────────────────────

function renderICList() {
  const grid = document.getElementById('ic-grid');
  if (!grid) return;
  grid.innerHTML = '';
  CM_ICS.forEach(ic => {
    const label = document.createElement('label');
    label.className = 'route-opt';
    label.innerHTML = `<input type="checkbox" name="selected_ics" value="${ic}" class="ic-check"><span title="${ic}">${ic}</span>`;
    grid.appendChild(label);
  });
}

function renderStateList() {
  const grid = document.getElementById('state-grid');
  if (!grid) return;
  grid.innerHTML = '';
  CM_STS.forEach(state => {
    const label = document.createElement('label');
    label.className = 'route-opt';
    label.innerHTML = `<input type="checkbox" name="selected_states" value="${state}" class="state-check"><span>${state}</span>`;
    grid.appendChild(label);
  });
}

// IC search
document.getElementById('ic-search')?.addEventListener('input', function () {
  const q = this.value.toLowerCase().trim();
  document.querySelectorAll('#ic-grid .route-opt').forEach(opt => {
    opt.style.display = opt.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
});

// ── Routing mode ──────────────────────────────────────────────────────────────

function setRoutingMode(mode, btn) {
  // Update segment buttons
  btn.closest('.seg-ctrl').querySelectorAll('.seg-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');

  // Update hidden input
  const inp = document.getElementById('city_mode_input');
  if (inp) inp.value = mode;

  // Show/hide sub-panels
  ['virtual', 'investment_center', 'cities', 'state'].forEach(m => {
    const p = document.getElementById(`rp-${m}`);
    if (p) p.classList.toggle('open', m === mode);
  });

  // Update routing badge
  updateRoutingBadge();
}

function setScope(scope, btn) {
  btn.closest('.scope-seg').querySelectorAll('.scope-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  const inp = document.getElementById('virtual_scope_input');
  if (inp) inp.value = scope;
  updateRoutingBadge();
}

function updateRoutingBadge() {
  const badge = document.getElementById('acc-badge-routing');
  if (!badge) return;
  const mode  = document.getElementById('city_mode_input')?.value || 'virtual';
  const scope = document.getElementById('virtual_scope_input')?.value || 'both';
  const modeLabels = { virtual: 'Virtual', investment_center: 'IC', cities: 'City', state: 'State' };
  const scopeLabels = { both: 'Both', nam: 'NAM', eur: 'EUR' };
  badge.textContent = `${modeLabels[mode] || mode} · ${scopeLabels[scope] || scope}`;
}

// ── AI pre-fill ───────────────────────────────────────────────────────────────

function applyAIResults(data) {
  const map = { industry: data.industry||[], style: data.style||[], mcap: data.mcap||[], geo: data.geo||[] };
  Object.entries(map).forEach(([dim, vals]) => {
    selections[dim].clear();
    vals.forEach(v => selections[dim].add(v));
    const list = document.getElementById(`list-${dim}`);
    if (list) {
      list.querySelectorAll('.cdf-item').forEach(item => {
        item.classList.toggle('selected', selections[dim].has(item.dataset.value));
      });
    }
    updateCount(dim, document.getElementById(`count-${dim}`));
  });
  updateRunButton();
}

document.getElementById('ai-btn')?.addEventListener('click', async () => {
  const input = document.getElementById('input-docs');
  if (!input?.files?.length) { showError('Please upload company documents first (10-K, deck, etc.)'); return; }

  const banner     = document.getElementById('ai-banner');
  const bannerText = document.getElementById('ai-banner-text');
  const loader     = document.getElementById('ai-loader');
  banner.classList.remove('hidden');
  bannerText.textContent = 'Analyzing documents…';
  loader.style.display = 'block';

  const fd = new FormData();
  Array.from(input.files).forEach(f => fd.append('documents', f));

  try {
    const r    = await fetch('/api/analyze', { method: 'POST', body: fd });
    const data = await r.json();
    if (data.error) throw new Error(data.error);
    applyAIResults(data);
    bannerText.textContent = 'CDF criteria pre-filled from documents — review and adjust as needed';
    loader.style.display = 'none';
    if (data.reasoning) {
      const tips = Object.entries(data.reasoning).map(([k,v]) => `${k}: ${v}`).join(' · ');
      bannerText.textContent += ` · ${tips}`;
    }
  } catch (e) {
    banner.classList.add('hidden');
    showError(`AI analysis failed: ${e.message}`);
  }
});

// ── File inputs ───────────────────────────────────────────────────────────────

let fileCount = 0;

function bindFile(inputId, statusId, onLoad) {
  const inp = document.getElementById(inputId);
  const sta = document.getElementById(statusId);
  inp?.addEventListener('change', () => {
    const f = inp.files[0];
    if (f) {
      sta.textContent = '✓ ' + f.name.replace(/\.[^.]+$/, '').substring(0, 20);
      if (onLoad) onLoad(f);
      fileCount++;
    }
    updateRunButton();
    updateAIButton();
    updateFilesBadge();
  });
}

function updateFilesBadge() {
  const badge = document.getElementById('acc-badge-files');
  if (!badge) return;
  const statuses = ['status-contacts','status-activities','status-ownership','status-fund','status-mining','status-docs'];
  const loaded = statuses.filter(id => {
    const el = document.getElementById(id);
    return el && el.textContent.trim().startsWith('✓');
  }).length;
  badge.textContent = `${loaded} / 6`;
}

bindFile('input-contacts',  'status-contacts');
bindFile('input-ownership', 'status-ownership');
bindFile('input-fund',      'status-fund');

document.getElementById('input-mining')?.addEventListener('change', function () {
  const n   = this.files.length;
  const sta = document.getElementById('status-mining');
  if (sta) sta.textContent = n ? `✓ ${n} file${n > 1 ? 's' : ''}` : '';
  updateRunButton();
  updateFilesBadge();
});

document.getElementById('input-docs')?.addEventListener('change', function () {
  const n = this.files.length;
  document.getElementById('status-docs').textContent = n ? `✓ ${n} file${n>1?'s':''}` : '';
  updateAIButton();
  updateFilesBadge();
});

bindFile('input-activities', 'status-activities', async (file) => {
  const fd = new FormData();
  fd.append('activities', file);
  try {
    const r    = await fetch('/api/detect-symbols', { method: 'POST', body: fd });
    const data = await r.json();
    populateSymbolRows(data.symbols || []);
  } catch {}
});

function populateSymbolRows(symbols) {
  const symbolRow = document.getElementById('symbol-row');
  const otherRow  = document.getElementById('other-symbols-row');
  const subjGrid  = document.getElementById('subject-symbols-grid');
  const otherGrid = document.getElementById('other-symbols-grid');

  if (!symbols.length) {
    if (symbolRow) symbolRow.style.display = 'none';
    if (otherRow)  otherRow.style.display  = 'none';
    return;
  }

  // Subject tickers checkboxes
  subjGrid.innerHTML = '';
  symbols.forEach(s => {
    const label = document.createElement('label');
    label.className = 'city-opt';
    label.dataset.symbol = s;
    label.innerHTML = `<input type="checkbox" name="subject_symbol" value="${s}" class="subj-sym-check"><span>${s}</span>`;
    subjGrid.appendChild(label);
  });
  if (symbols.length === 1) subjGrid.querySelector('input').checked = true;
  symbolRow.style.display = 'block';

  if (symbols.length > 1) {
    otherGrid.innerHTML = '';
    symbols.forEach(s => {
      const label = document.createElement('label');
      label.className = 'city-opt';
      label.dataset.symbol = s;
      label.innerHTML = `<input type="checkbox" name="other_symbols" value="${s}" class="other-sym-check"><span>${s}</span>`;
      otherGrid.appendChild(label);
    });
    subjGrid.querySelectorAll('input.subj-sym-check').forEach(cb => {
      cb.addEventListener('change', syncOtherSymbols);
    });
    syncOtherSymbols();
    otherRow.style.display = 'block';
  } else {
    otherRow.style.display = 'none';
  }
}

function syncOtherSymbols() {
  const subjGrid  = document.getElementById('subject-symbols-grid');
  const otherGrid = document.getElementById('other-symbols-grid');
  if (!subjGrid || !otherGrid) return;
  const checked = new Set(
    Array.from(subjGrid.querySelectorAll('input.subj-sym-check:checked')).map(cb => cb.value)
  );
  otherGrid.querySelectorAll('label[data-symbol]').forEach(lbl => {
    const isSubj = checked.has(lbl.dataset.symbol);
    lbl.style.display = isSubj ? 'none' : '';
    if (isSubj) lbl.querySelector('input').checked = false;
  });
}

function updateAIButton() {
  const btn   = document.getElementById('ai-btn');
  const input = document.getElementById('input-docs');
  if (btn) btn.disabled = !(input?.files?.length > 0);
}

function updateRunButton() {
  const btn     = document.getElementById('run-btn');
  const hasFile = !!document.getElementById('input-contacts')?.files?.length;
  if (btn) btn.disabled = !hasFile;
}

// ── Run filter ────────────────────────────────────────────────────────────────

document.getElementById('run-btn')?.addEventListener('click', runFilter);

async function runFilter() {
  const btn = document.getElementById('run-btn');
  btn.disabled = true;
  btn.querySelector('.run-label').textContent = 'Running…';
  hideError();
  document.getElementById('results-panel')?.classList.add('hidden');

  const fd = new FormData();

  const contactsFile = document.getElementById('input-contacts')?.files[0];
  if (!contactsFile) { showError('Please upload a contacts file.'); resetBtn(); return; }
  fd.append('contacts', contactsFile);

  const ownershipFile = document.getElementById('input-ownership')?.files[0];
  if (ownershipFile) fd.append('ownership', ownershipFile);

  const fundFile = document.getElementById('input-fund')?.files[0];
  if (fundFile) fd.append('fund_ownership', fundFile);

  const actsFile = document.getElementById('input-activities')?.files[0];
  if (actsFile) fd.append('activities', actsFile);

  const miningFiles = document.getElementById('input-mining')?.files;
  if (miningFiles) Array.from(miningFiles).forEach(f => fd.append('mining', f));

  fd.append('company_name', document.getElementById('company-name')?.value || 'Company');

  // Subject tickers
  document.querySelectorAll('input[name="subject_symbol"]:checked').forEach(inp => fd.append('subject_symbol', inp.value));
  // Other symbols
  document.querySelectorAll('input[name="other_symbols"]:checked').forEach(inp => fd.append('other_symbols', inp.value));

  fd.append('hf_treatment',           document.querySelector('input[name="hf_treatment"]:checked')?.value || 'separate');
  fd.append('meeting_exclusion',       document.querySelector('input[name="meeting_exclusion"]:checked')?.value || 'include_all');
  fd.append('shareholder_exclusion',   document.querySelector('input[name="shareholder_exclusion"]:checked')?.value || 'include_all');

  const eaumMin = document.getElementById('eaum-min')?.value?.trim();
  if (eaumMin) fd.append('eaum_min', eaumMin);

  // Routing — read from hidden inputs set by setRoutingMode/setScope
  const cityMode    = document.getElementById('city_mode_input')?.value    || 'virtual';
  const virtualScope = document.getElementById('virtual_scope_input')?.value || 'both';
  fd.append('city_mode',     cityMode);
  fd.append('virtual_scope', virtualScope);

  if (cityMode === 'investment_center') {
    document.querySelectorAll('input[name="selected_ics"]:checked').forEach(inp => fd.append('selected_ics', inp.value));
  } else if (cityMode === 'cities') {
    document.querySelectorAll('input[name="selected_cities"]:checked').forEach(inp => fd.append('selected_cities', inp.value.trim()));
  } else if (cityMode === 'state') {
    document.querySelectorAll('input[name="selected_states"]:checked').forEach(inp => fd.append('selected_states', inp.value));
  }

  selections.industry.forEach(v => fd.append('industry', v));
  selections.style.forEach(v    => fd.append('style', v));
  selections.mcap.forEach(v     => fd.append('mcap', v));
  selections.geo.forEach(v      => fd.append('geo', v));

  try {
    const r    = await fetch('/api/run', { method: 'POST', body: fd });
    const data = await r.json();
    if (!r.ok || data.error) throw new Error(data.error || 'Unknown error');
    renderResults(data);
  } catch (e) {
    showError(e.message);
  } finally {
    resetBtn();
  }
}

function resetBtn() {
  const btn = document.getElementById('run-btn');
  if (btn) { btn.disabled = false; btn.querySelector('.run-label').textContent = 'Run filter'; }
}

// ── Results rendering ─────────────────────────────────────────────────────────

function renderResults(data) {
  const panel = document.getElementById('results-panel');
  panel.classList.remove('hidden');

  const grid = document.getElementById('stats-grid');
  grid.innerHTML = '';

  const stats = [{ label: 'Source contacts', value: data.total_source }];

  if (data.city_counts && Object.keys(data.city_counts).length > 0) {
    Object.entries(data.city_counts).forEach(([city, n]) => {
      stats.push({ label: city, value: n, highlight: true });
    });
  } else {
    stats.push({ label: 'Contacts', value: data.main_count, highlight: true });
  }

  const subs = [
    ['HFs',          data.hf_count],
    ['DNC',          data.dnc_count],
    ['Check',        data.check_count],
    ['Quant',        data.quant_count],
    ['Activist',     data.activist_count],
    ['Fixed Income', data.fi_count],
    ['Too Small',    data.too_small_count],
    ['Excluded',     data.excluded_count],
  ];
  subs.forEach(([label, val]) => { if (val > 0) stats.push({ label, value: val }); });

  stats.forEach(s => {
    const card = document.createElement('div');
    card.className = 'stat-card' + (s.highlight ? ' highlight' : '');
    card.innerHTML = `<div class="stat-label">${s.label}</div><div class="stat-value">${(s.value||0).toLocaleString()}</div>`;
    grid.appendChild(card);
  });

  const bars     = document.getElementById('breakdown-bars');
  bars.innerHTML = '';
  const breakdown = data.match_breakdown || {};
  const maxVal    = Math.max(...Object.values(breakdown), 1);

  [4, 3, 2, 1].forEach(n => {
    const val = breakdown[n] || 0;
    if (!Object.keys(breakdown).length && n > 1) return;
    const row = document.createElement('div');
    row.className = 'breakdown-row';
    const pct = Math.round((val / maxVal) * 100);
    row.innerHTML = `
      <div class="breakdown-key">${n} criteri${n===1?'on':'a'}</div>
      <div class="breakdown-bar-wrap"><div class="breakdown-bar" style="width:${pct}%"></div></div>
      <div class="breakdown-val">${val.toLocaleString()}</div>`;
    bars.appendChild(row);
  });

  const spBtn = document.getElementById('sharepoint-btn');
  if (data.sharepoint_url && spBtn) {
    spBtn.href = data.sharepoint_url;
    spBtn.classList.remove('hidden');
  }

  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ── Error helpers ─────────────────────────────────────────────────────────────

function showError(msg) {
  const p = document.getElementById('error-panel');
  const t = document.getElementById('error-text');
  if (p && t) { t.textContent = msg; p.classList.remove('hidden'); }
}
function hideError() {
  document.getElementById('error-panel')?.classList.add('hidden');
}

// ── Init ──────────────────────────────────────────────────────────────────────

Object.keys(DIM_MAP).forEach(renderCDFList);
renderICList();
renderStateList();
updateRunButton();
updateAIButton();
