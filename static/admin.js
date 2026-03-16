/* ── Meeting history upload ─────────────────────────────────────────────── */
let meetingsFile = null;

const meetingsZone  = document.getElementById('meetings-zone');
const meetingsInput = document.getElementById('meetings-file');
const previewBox    = document.getElementById('meetings-preview');
const confirmBtn    = document.getElementById('confirm-meetings-btn');
const meetingsResult = document.getElementById('meetings-result');

if (meetingsZone) {
  meetingsZone.addEventListener('click', () => meetingsInput.click());
  meetingsZone.addEventListener('dragover', e => { e.preventDefault(); meetingsZone.classList.add('drag-over'); });
  meetingsZone.addEventListener('dragleave', () => meetingsZone.classList.remove('drag-over'));
  meetingsZone.addEventListener('drop', e => {
    e.preventDefault();
    meetingsZone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) handleMeetingsFile(f);
  });

  meetingsInput.addEventListener('change', () => {
    if (meetingsInput.files[0]) handleMeetingsFile(meetingsInput.files[0]);
  });
}

async function handleMeetingsFile(f) {
  meetingsFile = f;
  meetingsZone.querySelector('.upload-hint').textContent = `Selected: ${f.name}`;
  previewBox.classList.add('hidden');
  confirmBtn.disabled = true;
  meetingsResult.classList.add('hidden');

  const uploadType = document.querySelector('input[name=upload_type]:checked')?.value || 'incremental';
  if (uploadType === 'incremental') {
    const form = new FormData();
    form.append('meetings_file', f);
    try {
      const resp = await fetch('/api/admin/preview-meetings', { method: 'POST', body: form });
      const data = await resp.json();
      if (data.error) throw new Error(data.error);
      previewBox.classList.remove('hidden');
      previewBox.textContent =
        `Preview: ${data.added.toLocaleString()} new rows, ` +
        `${data.updated.toLocaleString()} updates, ` +
        `${data.skipped.toLocaleString()} skipped. ` +
        `Click "Confirm upload" to proceed.`;
    } catch (err) {
      previewBox.classList.remove('hidden');
      previewBox.style.background = 'var(--danger-bg)';
      previewBox.style.color = 'var(--danger)';
      previewBox.textContent = `Preview error: ${err.message}`;
    }
  }
  confirmBtn.disabled = false;
}

if (confirmBtn) {
  confirmBtn.addEventListener('click', async () => {
    if (!meetingsFile) return;
    confirmBtn.disabled = true;
    confirmBtn.textContent = 'Uploading…';

    const uploadType = document.querySelector('input[name=upload_type]:checked')?.value || 'incremental';
    const form = new FormData();
    form.append('meetings_file', meetingsFile);
    form.append('upload_type', uploadType);

    try {
      const resp = await fetch('/api/admin/upload-meetings', { method: 'POST', body: form });
      const data = await resp.json();
      if (data.error) throw new Error(data.error);

      meetingsResult.className = 'result-msg success';
      meetingsResult.textContent =
        `✓ Upload complete — ${data.added.toLocaleString()} added, ` +
        `${data.updated.toLocaleString()} updated. ` +
        `Total rows: ${data.total.toLocaleString()}`;
      meetingsResult.classList.remove('hidden');
      previewBox.classList.add('hidden');
    } catch (err) {
      meetingsResult.className = 'result-msg error';
      meetingsResult.textContent = `Upload failed: ${err.message}`;
      meetingsResult.classList.remove('hidden');
    } finally {
      confirmBtn.textContent = 'Confirm upload';
      confirmBtn.disabled = false;
    }
  });
}

/* ── Taxonomy upload ────────────────────────────────────────────────────── */
let taxonomyFile = null;

const taxonomyZone   = document.getElementById('taxonomy-zone');
const taxonomyInput  = document.getElementById('taxonomy-file');
const taxonomyBtn    = document.getElementById('upload-taxonomy-btn');
const taxonomyResult = document.getElementById('taxonomy-result');

if (taxonomyZone) {
  taxonomyZone.addEventListener('click', () => taxonomyInput.click());
  taxonomyZone.addEventListener('dragover', e => { e.preventDefault(); taxonomyZone.classList.add('drag-over'); });
  taxonomyZone.addEventListener('dragleave', () => taxonomyZone.classList.remove('drag-over'));
  taxonomyZone.addEventListener('drop', e => {
    e.preventDefault();
    taxonomyZone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) { taxonomyFile = f; taxonomyZone.querySelector('.upload-hint').textContent = `Selected: ${f.name}`; taxonomyBtn.disabled = false; }
  });

  taxonomyInput.addEventListener('change', () => {
    if (taxonomyInput.files[0]) {
      taxonomyFile = taxonomyInput.files[0];
      taxonomyZone.querySelector('.upload-hint').textContent = `Selected: ${taxonomyFile.name}`;
      taxonomyBtn.disabled = false;
    }
  });
}

if (taxonomyBtn) {
  taxonomyBtn.addEventListener('click', async () => {
    if (!taxonomyFile) return;
    taxonomyBtn.disabled = true;
    taxonomyBtn.textContent = 'Uploading…';

    const form = new FormData();
    form.append('taxonomy_file', taxonomyFile);

    try {
      const resp = await fetch('/api/admin/upload-taxonomy', { method: 'POST', body: form });
      const data = await resp.json();
      if (data.error) throw new Error(data.error);

      taxonomyResult.className = 'result-msg success';
      taxonomyResult.textContent = `✓ Taxonomy updated — ${data.rows} values loaded. Reload the filter page to see changes.`;
      taxonomyResult.classList.remove('hidden');
    } catch (err) {
      taxonomyResult.className = 'result-msg error';
      taxonomyResult.textContent = `Upload failed: ${err.message}`;
      taxonomyResult.classList.remove('hidden');
    } finally {
      taxonomyBtn.textContent = 'Upload taxonomy';
      taxonomyBtn.disabled = false;
    }
  });
}

/* ── City map upload ───────────────────────────────────────────────────── */
let citymapFile = null;

const citymapZone   = document.getElementById('citymap-zone');
const citymapInput  = document.getElementById('citymap-file');
const citymapBtn    = document.getElementById('upload-citymap-btn');
const citymapResult = document.getElementById('citymap-result');

if (citymapZone) {
  citymapZone.addEventListener('click', () => citymapInput.click());
  citymapZone.addEventListener('dragover', e => { e.preventDefault(); citymapZone.classList.add('drag-over'); });
  citymapZone.addEventListener('dragleave', () => citymapZone.classList.remove('drag-over'));
  citymapZone.addEventListener('drop', e => {
    e.preventDefault();
    citymapZone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) { citymapFile = f; citymapZone.querySelector('.upload-hint').textContent = `Selected: ${f.name}`; citymapBtn.disabled = false; }
  });

  citymapInput.addEventListener('change', () => {
    if (citymapInput.files[0]) {
      citymapFile = citymapInput.files[0];
      citymapZone.querySelector('.upload-hint').textContent = `Selected: ${citymapFile.name}`;
      citymapBtn.disabled = false;
    }
  });
}

if (citymapBtn) {
  citymapBtn.addEventListener('click', async () => {
    if (!citymapFile) return;
    citymapBtn.disabled = true;
    citymapBtn.textContent = 'Uploading…';

    const form = new FormData();
    form.append('city_map_file', citymapFile);

    try {
      const resp = await fetch('/api/admin/upload-city-map', { method: 'POST', body: form });
      const data = await resp.json();
      if (data.error) throw new Error(data.error);

      citymapResult.className = 'result-msg success';
      citymapResult.textContent = `✓ City map updated — ${data.rows} rows, ${data.investment_centers} investment centers. Reload the filter page to see changes.`;
      citymapResult.classList.remove('hidden');
    } catch (err) {
      citymapResult.className = 'result-msg error';
      citymapResult.textContent = `Upload failed: ${err.message}`;
      citymapResult.classList.remove('hidden');
    } finally {
      citymapBtn.textContent = 'Upload city map';
      citymapBtn.disabled = false;
    }
  });
}
