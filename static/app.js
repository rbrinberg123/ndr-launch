/* ── NDR Launch app.js ─────────────────────────────────────────────────────── */

const TAX = window.TAXONOMY || {};

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

  // Group Industry Focus by top-level prefix
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

  const cb   = document.createElement('div'); cb.className = 'cdf-checkbox';
  const txt  = document.createElement('div'); txt.className = 'cdf-item-text'; txt.textContent = val;
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

// ── Search filtering ──────────────────────────────────────────────────────────

document.querySelectorAll('.search-input').forEach(inp => {
  inp.addEventListener('input', () => {
    const dim   = inp.dataset.target;
    const query = inp.value.toLowerCase().trim();
    const list  = document.getElementById(`list-${dim}`);
    if (!list) return;
    list.querySelectorAll('.cdf-item').forEach(item => {
      const match = item.dataset.value.toLowerCase().includes(query);
      item.classList.toggle('hidden-item', !match);
    });
    // Hide group headers if all children hidden
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

// ── AI pre-fill ───────────────────────────────────────────────────────────────

function applyAIResults(data) {
  const map = { industry: data.industry||[], style: data.style||[], mcap: data.mcap||[], geo: data.geo||[] };
  Object.entries(map).forEach(([dim, vals]) => {
    selections[dim].clear();
    vals.forEach(v => selections[dim].add(v));
    // Re-render checkboxes
    const list = document.getElementById(`list-${dim}`);
    if (list) {
      list.querySelectorAll('.cdf-item').forEach(item => {
        const sel = selections[dim].has(item.dataset.value);
        item.classList.toggle('selected', sel);
      });
    }
    updateCount(dim, document.getElementById(`count-${dim}`));
  });
  updateRunButton();
}

document.getElementById('ai-btn')?.addEventListener('click', async () => {
  const input = document.getElementById('input-docs');
  if (!input?.files?.length) {
    showError('Please upload company documents first (10-K, deck, etc.)');
    return;
  }
  const banner = document.getElementById('ai-banner');
  const bannerText = document.getElementById('ai-banner-text');
  const loader = document.getElementById('ai-loader');
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

    // Show reasoning if present
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

function bindFile(inputId, statusId, onLoad) {
  const inp = document.getElementById(inputId);
  const sta = document.getElementById(statusId);
  inp?.addEventListener('change', () => {
    const f = inp.files[0];
    if (f) {
      sta.textContent = '✓ ' + f.name.replace(/\.[^.]+$/, '').substring(0, 22);
      if (onLoad) onLoad(f);
    }
    updateRunButton();
    updateAIButton();
  });
}

bindFile('input-contacts',   'status-contacts');
bindFile('input-ownership',  'status-ownership');
bindFile('input-fund',       'status-fund');

bindFile('input-activities', 'status-activities', async (file) => {
  const fd = new FormData();
  fd.append('activities', file);
  try {
    const r    = await fetch('/api/detect-symbols', { method: 'POST', body: fd });
    const data = await r.json();
    const row  = document.getElementById('symbol-row');
    const grid = document.getElementById('symbol-grid');
    if (data.symbols?.length) {
      grid.innerHTML = '';
      data.symbols.forEach(s => {
        const label = document.createElement('label');
        label.className = 'city-opt';
        const cb = document.createElement('input');
        cb.type = 'checkbox'; cb.name = 'subject_symbols'; cb.value = s; cb.className = 'city-check';
        if (data.symbols.length === 1) cb.checked = true;
        const span = document.createElement('span');
        span.textContent = s;
        label.appendChild(cb); label.appendChild(span);
        grid.appendChild(label);
      });
      row.style.display = 'block';
    } else {
      row.style.display = 'none';
    }
  } catch {}
});

// Multi-file docs
document.getElementById('input-docs')?.addEventListener('change', function() {
  const n = this.files.length;
  document.getElementById('status-docs').textContent = n ? `✓ ${n} file${n>1?'s':''}` : '';
  updateAIButton();
});

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

// ── City routing toggle ───────────────────────────────────────────────────────

document.querySelectorAll('input[name="city_mode"]').forEach(r => {
  r.addEventListener('change', () => {
    const show = r.value === 'cities' && r.checked;
    const el = document.getElementById('city-inputs');
    if (el) el.style.display = show ? 'block' : 'none';
  });
});

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

  fd.append('company_name',      document.getElementById('company-name')?.value || 'Company');
  document.querySelectorAll('input[name="subject_symbols"]:checked').forEach(inp => {
    fd.append('subject_symbols', inp.value);
  });
  fd.append('hf_treatment',      document.querySelector('input[name="hf_treatment"]:checked')?.value || 'separate');
  fd.append('meeting_exclusion', document.querySelector('input[name="meeting_exclusion"]:checked')?.value || 'include_all');

  const cityMode = document.querySelector('input[name="city_mode"]:checked')?.value || 'virtual';
  fd.append('city_mode', cityMode);
  if (cityMode === 'cities') {
    document.querySelectorAll('input[name="selected_cities"]:checked').forEach(inp => {
      fd.append('selected_cities', inp.value.trim());
    });
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

  // Stats cards
  const grid = document.getElementById('stats-grid');
  grid.innerHTML = '';

  const stats = [
    { label: 'Source contacts', value: data.total_source },
  ];

  if (data.city_counts && Object.keys(data.city_counts).length > 0) {
    Object.entries(data.city_counts).forEach(([city, n]) => {
      stats.push({ label: city, value: n, highlight: true });
    });
  } else {
    stats.push({ label: 'Contacts', value: data.main_count, highlight: true });
  }

  const subs = [
    ['HFs', data.hf_count],
    ['DNC', data.dnc_count],
    ['Check', data.check_count],
    ['Quant', data.quant_count],
    ['Activist', data.activist_count],
    ['Excluded', data.excluded_count],
  ];
  subs.forEach(([label, val]) => { if (val > 0) stats.push({ label, value: val }); });

  stats.forEach(s => {
    const card = document.createElement('div');
    card.className = 'stat-card' + (s.highlight ? ' highlight' : '');
    card.innerHTML = `<div class="stat-label">${s.label}</div><div class="stat-value">${(s.value||0).toLocaleString()}</div>`;
    grid.appendChild(card);
  });

  // Breakdown bars
  const bars = document.getElementById('breakdown-bars');
  bars.innerHTML = '';
  const breakdown = data.match_breakdown || {};
  const maxVal = Math.max(...Object.values(breakdown), 1);

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

  // SharePoint link
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
updateRunButton();
updateAIButton();
