/* ============================================================
   SheetGenie — app.js
   Capture -> POST /api/improve -> preview -> POST /api/generate -> download.
   Vanilla JS. No frameworks. DOM built with createElement / textContent only
   (server-supplied strings are never injected as HTML).
   ============================================================ */
'use strict';

(function () {
  /* ---------- tiny helpers ---------- */
  const $ = (id) => document.getElementById(id);

  const el = (tag, opts = {}, children = []) => {
    const node = document.createElement(tag);
    if (opts.class) node.className = opts.class;
    if (opts.text != null) node.textContent = String(opts.text);
    if (opts.attrs) for (const [k, v] of Object.entries(opts.attrs)) node.setAttribute(k, v);
    for (const c of children) if (c) node.appendChild(c);
    return node;
  };

  // fetch with a hard timeout so a slow/hung request can never freeze the UI.
  function fetchWithTimeout(url, opts, ms) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), ms);
    return fetch(url, Object.assign({}, opts, { signal: ctrl.signal }))
      .finally(() => clearTimeout(timer));
  }

  /* ---------- element refs ---------- */
  const promptEl = $('prompt');
  const micBtn = $('mic');
  const micHint = $('micHint');
  const dataToggle = $('dataToggle');
  const dataPanel = $('dataPanel');
  const dataInput = $('data');
  const dataBadge = $('dataBadge');
  const uploadBtn = $('uploadBtn');
  const cameraBtn = $('cameraBtn');
  const fileInput = $('fileInput');
  const cameraInput = $('cameraInput');
  const attachList = $('attachList');
  const improveBtn = $('improve');
  const downloadBtn = $('download');
  const toastEl = $('toast');
  const resultsEl = $('results');
  const themeToggle = $('themeToggle');
  const saveLibBtn = $('saveLib');
  const libraryToggle = $('libraryToggle');
  const libCount = $('libCount');
  const librarySection = $('library');
  const libraryClose = $('libraryClose');
  const libList = $('libList');
  const libEmpty = $('libEmpty');
  const editBanner = $('editBanner');
  const editTitle = $('editTitle');
  const newSheetBtn = $('newSheetBtn');

  /* ---------- app state ---------- */
  let currentSpec = null;     // last valid SpreadsheetSpec from /api/improve
  let busy = false;           // a network call is in flight
  let toastTimer = null;
  let attachments = [];       // {name, kind:'image'|'pdf', mediaType, data(base64), thumb}
  const MAX_ATTACH = 6;
  const MAX_TOTAL_B64 = 3500000;  // ~2.6 MB raw; keeps the request under the serverless limit
  let baseSpec = null;        // when set, /api/improve EDITS this spec (edit mode)
  let currentLibId = null;    // library record id for the working spec (null = unsaved)
  const PROMPT_PLACEHOLDER = promptEl.placeholder;

  /* ============================================================
     Theme: manual override stored in localStorage, else OS preference.
     ============================================================ */
  (function initTheme() {
    let saved = null;
    try { saved = localStorage.getItem('sg-theme'); } catch (_) {}
    // Keep the browser/installed-app UI chrome color in sync with a manual theme
    // override (the media-scoped <meta>s only follow OS preference).
    function applyThemeColor() {
      const theme = document.documentElement.getAttribute('data-theme');
      let meta = document.getElementById('tc-manual');
      if (theme !== 'light' && theme !== 'dark') { if (meta) meta.remove(); return; }
      if (!meta) {
        meta = document.createElement('meta');
        meta.id = 'tc-manual';
        meta.name = 'theme-color';
        document.head.appendChild(meta);
      }
      meta.setAttribute('content', theme === 'dark' ? '#0f0e17' : '#f6f5f2');
    }

    if (saved === 'light' || saved === 'dark') {
      document.documentElement.setAttribute('data-theme', saved);
    } else {
      document.documentElement.removeAttribute('data-theme');
    }
    applyThemeColor();
    themeToggle.addEventListener('click', () => {
      const isDark = matchMedia('(prefers-color-scheme: dark)').matches;
      const current = document.documentElement.getAttribute('data-theme') || (isDark ? 'dark' : 'light');
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      try { localStorage.setItem('sg-theme', next); } catch (_) {}
      applyThemeColor();
    });
  })();

  /* ============================================================
     Autosizing textarea for the prompt.
     ============================================================ */
  function autosize() {
    promptEl.style.height = 'auto';
    promptEl.style.height = Math.min(promptEl.scrollHeight, 360) + 'px';
  }
  promptEl.addEventListener('input', autosize);
  autosize();

  /* ============================================================
     Collapsible "I have data to fill in".
     ============================================================ */
  dataToggle.addEventListener('click', () => {
    const open = dataToggle.getAttribute('aria-expanded') === 'true';
    dataToggle.setAttribute('aria-expanded', String(!open));
    dataPanel.hidden = open;
    if (!open) dataInput.focus();
  });

  function hasData() {
    return dataInput.value.trim().length > 0 || attachments.length > 0;
  }
  function updateDataBadge() { dataBadge.hidden = !hasData(); }

  dataInput.addEventListener('input', updateDataBadge);

  /* ---------- attachments: upload / camera / text files ---------- */
  uploadBtn.addEventListener('click', () => fileInput.click());
  cameraBtn.addEventListener('click', () => cameraInput.click());
  fileInput.addEventListener('change', (e) => { handleFiles(e.target.files); fileInput.value = ''; });
  cameraInput.addEventListener('change', (e) => { handleFiles(e.target.files); cameraInput.value = ''; });

  async function handleFiles(fileList) {
    const files = Array.from(fileList || []);
    for (const file of files) {
      if (attachments.length >= MAX_ATTACH) {
        showError('You can attach up to ' + MAX_ATTACH + ' files.');
        break;
      }
      try {
        if (file.type && file.type.indexOf('image/') === 0) {
          const b64 = await downscaleImageToJpegBase64(file);
          if (!addAttachment({ name: file.name || 'photo.jpg', kind: 'image',
            mediaType: 'image/jpeg', data: b64, thumb: 'data:image/jpeg;base64,' + b64 })) break;
        } else if (file.type === 'application/pdf' || /\.pdf$/i.test(file.name || '')) {
          const b64 = await readAsBase64(file);
          if (!addAttachment({ name: file.name || 'document.pdf', kind: 'pdf',
            mediaType: 'application/pdf', data: b64, thumb: null })) break;
        } else if (isTextLike(file)) {
          appendData(await readAsText(file));
        } else {
          showError('Unsupported file: ' + (file.name || file.type || 'unknown') + '. Use an image, PDF, or text/CSV.');
        }
      } catch (_) {
        showError('Could not read "' + (file.name || 'that file') + '". Try a different file.');
      }
    }
    renderAttachments();
    updateDataBadge();
  }

  function isTextLike(file) {
    return (file.type && (file.type.indexOf('text/') === 0 || file.type === 'text/csv'))
      || /\.(csv|tsv|txt)$/i.test(file.name || '');
  }

  function appendData(text) {
    const cur = dataInput.value;
    dataInput.value = (cur && cur.trim()) ? (cur.replace(/\s+$/, '') + '\n' + text) : text;
  }

  function totalB64() { return attachments.reduce((n, a) => n + a.data.length, 0); }

  function addAttachment(att) {
    if (totalB64() + att.data.length > MAX_TOTAL_B64) {
      showError('That would make the upload too large. Use fewer or smaller files (or a clearer photo).');
      return false;
    }
    attachments.push(att);
    return true;
  }

  function readAsBase64(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => { const s = String(r.result); const i = s.indexOf(','); resolve(i >= 0 ? s.slice(i + 1) : s); };
      r.onerror = () => reject(r.error || new Error('read failed'));
      r.readAsDataURL(file);
    });
  }

  function readAsText(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result || ''));
      r.onerror = () => reject(r.error || new Error('read failed'));
      r.readAsText(file);
    });
  }

  // Load an image, downscale to <= 2000px on the long edge, re-encode as JPEG, and
  // return the base64 (no data: prefix). Keeps photos small enough to upload.
  function downscaleImageToJpegBase64(file) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        URL.revokeObjectURL(url);
        const MAX = 2000;
        let w = img.naturalWidth || img.width;
        let h = img.naturalHeight || img.height;
        // SVG/undecodable images can report no intrinsic size -> a 0x0 canvas would
        // produce a blank JPEG. Reject so the caller shows a friendly error instead.
        if (!w || !h) { reject(new Error('image has no dimensions')); return; }
        if (Math.max(w, h) > MAX) { const s = MAX / Math.max(w, h); w = Math.round(w * s); h = Math.round(h * s); }
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        const durl = canvas.toDataURL('image/jpeg', 0.82);
        const i = durl.indexOf(',');
        resolve(i >= 0 ? durl.slice(i + 1) : durl);
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('image decode failed')); };
      img.src = url;
    });
  }

  function renderAttachments() {
    attachList.replaceChildren();
    if (!attachments.length) { attachList.hidden = true; return; }
    attachList.hidden = false;
    attachments.forEach((att, idx) => {
      const li = el('li', { class: 'attach-item' });
      if (att.thumb) li.appendChild(el('img', { class: 'attach-thumb', attrs: { src: att.thumb, alt: '' } }));
      else li.appendChild(el('span', { class: 'attach-ic', attrs: { 'aria-hidden': 'true' }, text: '📄' }));
      li.appendChild(el('span', { class: 'attach-name', text: att.name }));
      const rm = el('button', { class: 'attach-remove', text: '✕',
        attrs: { type: 'button', 'aria-label': 'Remove ' + att.name } });
      rm.addEventListener('click', () => { attachments.splice(idx, 1); renderAttachments(); updateDataBadge(); });
      li.appendChild(rm);
      attachList.appendChild(li);
    });
  }

  /* ============================================================
     Voice capture via Web Speech API.
     Only enabled if SpeechRecognition exists; otherwise hide mic + hint.
     ============================================================ */
  (function initVoice() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) {
      micBtn.hidden = true;
      micHint.hidden = false;
      return;
    }

    let recog = null;
    let recording = false;
    let baseText = '';   // text already in the textarea when recording began

    function setRecording(on) {
      recording = on;
      micBtn.setAttribute('aria-pressed', String(on));
      micBtn.setAttribute('aria-label', on ? 'Stop dictation' : 'Dictate with your voice');
    }

    function stop() {
      if (recog) { try { recog.stop(); } catch (_) {} }
    }

    micBtn.addEventListener('click', () => {
      if (recording) { stop(); return; }

      recog = new SR();
      recog.lang = navigator.language || 'en-US';
      recog.continuous = true;
      recog.interimResults = true;

      baseText = promptEl.value;
      const sep = baseText && !/\s$/.test(baseText) ? ' ' : '';

      recog.onresult = (event) => {
        let finalChunk = '';
        let interimChunk = '';
        for (let i = event.resultIndex; i < event.results.length; i++) {
          const res = event.results[i];
          if (res.isFinal) finalChunk += res[0].transcript;
          else interimChunk += res[0].transcript;
        }
        // Commit finals into baseText so they survive across result events.
        if (finalChunk) {
          baseText = baseText + sep + finalChunk.trim();
        }
        const liveSep = baseText && interimChunk && !/\s$/.test(baseText) ? ' ' : '';
        promptEl.value = baseText + (interimChunk ? liveSep + interimChunk : '');
        autosize();
      };

      recog.onerror = (event) => {
        setRecording(false);
        const code = event && event.error;
        if (code === 'not-allowed' || code === 'service-not-allowed') {
          showError('Microphone access was blocked. You can type instead.');
        } else if (code && code !== 'aborted' && code !== 'no-speech') {
          showError('Voice input stopped. You can type instead.');
        }
      };

      recog.onend = () => { setRecording(false); };

      try {
        recog.start();
        setRecording(true);
      } catch (_) {
        setRecording(false);
        showError('Could not start voice input. You can type instead.');
      }
    });
  })();

  /* ============================================================
     Toast / inline error — non-blocking, never crashes the UI.
     ============================================================ */
  function showError(message) {
    toastEl.textContent = String(message || 'Something went wrong.');
    toastEl.classList.remove('toast-ok');
    toastEl.hidden = false;
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.hidden = true; }, 7000);
  }
  function showOk(message) {
    toastEl.textContent = String(message || 'Done.');
    toastEl.classList.add('toast-ok');
    toastEl.hidden = false;
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.hidden = true; toastEl.classList.remove('toast-ok'); }, 4000);
  }
  function clearError() {
    toastEl.hidden = true;
    toastEl.textContent = '';
    toastEl.classList.remove('toast-ok');
    if (toastTimer) clearTimeout(toastTimer);
  }

  /* ============================================================
     Loading state for the primary action.
     ============================================================ */
  function refreshPrimaryLabel() {
    const label = improveBtn.querySelector('.btn-label');
    if (label) label.textContent = busy ? 'Working...' : (baseSpec ? 'Apply changes' : 'Improve & preview');
  }
  function setBusy(on) {
    busy = on;
    improveBtn.classList.toggle('is-loading', on);
    improveBtn.disabled = on;
    // Download / Save stay disabled while busy or until we have a spec.
    downloadBtn.disabled = on || !currentSpec;
    saveLibBtn.disabled = on || !currentSpec;
    refreshPrimaryLabel();
  }

  /* ============================================================
     Primary action: /api/improve
     ============================================================ */
  improveBtn.addEventListener('click', onImprove);

  // Remember the original request so a clarification round can re-send it.
  let pendingPrompt = '';
  let pendingData = null;
  let pendingFiles = [];
  let pendingBaseSpec = null;

  function onImprove() {
    if (busy) return;
    let prompt = promptEl.value.trim();
    const haveData = dataInput.value.trim().length > 0 || attachments.length > 0;
    if (!prompt && haveData) {
      prompt = baseSpec ? 'Update this spreadsheet with the new data.' : 'Create a spreadsheet from the attached data.';
    }
    if (!prompt) {
      showError(baseSpec ? 'Tell me what to change, or add some data.' : 'Tell me what spreadsheet you need first.');
      promptEl.focus();
      return;
    }
    // A fresh, non-edit generation starts a NEW library entry. Drop any stale
    // currentLibId so saving/downloading the new spec can't overwrite the record
    // of a previously-saved spreadsheet.
    if (!baseSpec) currentLibId = null;
    pendingPrompt = prompt;
    pendingData = dataInput.value.trim() ? dataInput.value : null;
    pendingFiles = attachments.map((a) => ({ type: a.kind, media_type: a.mediaType, data: a.data, name: a.name }));
    pendingBaseSpec = baseSpec;
    runImprove(null);
  }

  // Calls /api/improve. `clarifications` is null on the first pass, or an array of
  // {question, answer} once the user has answered the model's questions.
  async function runImprove(clarifications) {
    if (busy) return;
    clearError();
    setBusy(true);
    // A fresh request invalidates any prior spec. An EDIT keeps the base spec so a
    // failed "Apply changes" leaves the loaded file intact and still downloadable.
    if (!pendingBaseSpec) {
      currentSpec = null;
      downloadBtn.disabled = true;
    }

    const body = {
      prompt: pendingPrompt,
      hasData: !!(pendingData && pendingData.trim()) || pendingFiles.length > 0,
      data: pendingData,
      baseSpec: pendingBaseSpec || null,
      files: pendingFiles.length ? pendingFiles : null,
      clarifications: clarifications || null,
    };

    // Pre-flight size guard: baseSpec + data + attachments must fit the serverless
    // body limit, else the server 413s only after a wasted upload.
    if (JSON.stringify(body).length > 3900000) {
      showError(pendingBaseSpec
        ? 'This spreadsheet plus the new data is too large to edit at once — remove an attachment or split the change.'
        : 'That request is too large — remove an attachment or shorten the data.');
      setBusy(false);
      return;
    }

    try {
      const res = await fetchWithTimeout('/api/improve', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
        body: JSON.stringify(body),
      }, 90000);

      let payload = null;
      try { payload = await res.json(); } catch (_) { payload = null; }

      if (!res.ok || !payload || payload.error) {
        showError((payload && payload.error) ? payload.error
          : 'The improve step failed (' + res.status + '). Please try again.');
        return;
      }

      handleImproveResult(payload);
    } catch (err) {
      if (err && err.name === 'AbortError') {
        showError('That took too long. Try a simpler request or less data, then retry.');
      } else {
        showError('Could not reach the server. Check your connection and try again.');
      }
    } finally {
      setBusy(false);
    }
  }

  // Branch on the two-mode response (docs/SPEC.md §1): clarifying questions, or a
  // ready-to-build spec.
  function handleImproveResult(payload) {
    if (payload.status === 'needs_input' &&
        Array.isArray(payload.questions) && payload.questions.length) {
      renderQuestions(payload);
      resultsEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    // "ready" (or a legacy response without status): expect a spec.
    if (!payload.spec || typeof payload.spec !== 'object' || !Array.isArray(payload.spec.sheets)) {
      showError('The assistant returned an unexpected result. Please try again.');
      return;
    }
    currentSpec = payload.spec;
    downloadBtn.disabled = false;
    saveLibBtn.disabled = false;
    renderResults(payload);
    // If this was an edit, the updated spec becomes the new base for further edits,
    // and we clear the instruction inputs so the next change starts fresh.
    if (baseSpec) {
      baseSpec = currentSpec;
      promptEl.value = ''; autosize();
      dataInput.value = ''; attachments = []; renderAttachments(); updateDataBadge();
      if (currentLibId) saveCurrentToLibrary(false);  // keep a saved file's edits persisted
    }
    resultsEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // Render the model's clarifying questions as a short form; Continue re-runs
  // improve with the answers folded in.
  function renderQuestions(payload) {
    resultsEl.replaceChildren();

    const card = el('div', { class: 'questions-card' });
    card.appendChild(el('p', { class: 'eyebrow', text: 'A couple of quick questions' }));
    card.appendChild(el('p', {
      class: 'notes',
      text: payload.notes || 'A few details will help me build the right spreadsheet.',
    }));

    const form = el('div', { class: 'q-form' });
    const fields = [];
    payload.questions.slice(0, 4).forEach((q, i) => {
      const qText = (q && q.question) ? q.question : ('Question ' + (i + 1));
      const id = 'q-' + i;
      const input = el('input', {
        class: 'q-input',
        attrs: { id: id, type: 'text', placeholder: (q && q.hint) ? q.hint : 'Your answer (optional)' },
      });
      fields.push({ question: qText, input });
      form.appendChild(el('div', { class: 'q-item' }, [
        el('label', { class: 'q-label', text: qText, attrs: { for: id } }),
        input,
      ]));
    });
    card.appendChild(form);

    const cont = el('button', { class: 'btn btn-primary', attrs: { type: 'button' }, text: 'Continue' });
    cont.addEventListener('click', async () => {
      const clar = fields.map((f) => ({ question: f.question, answer: (f.input.value || '').trim() }));
      const label = cont.textContent;
      cont.disabled = true;
      cont.textContent = 'Working...';
      await runImprove(clar);
      if (cont.isConnected) { cont.disabled = false; cont.textContent = label; }
    });
    fields.forEach((f) => f.input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); cont.click(); }
    }));

    card.appendChild(el('div', { class: 'q-actions' }, [cont]));
    resultsEl.appendChild(card);
    if (fields[0]) fields[0].input.focus();
  }

  /* ============================================================
     Render preview — all via createElement/textContent (XSS-safe).
     ============================================================ */
  function renderResults(payload) {
    resultsEl.replaceChildren();

    // Improved prompt + notes (AI results), or just a note (loaded library files).
    if (payload.improvedPrompt) {
      const improvedCard = el('div', { class: 'improved-card' });
      improvedCard.appendChild(el('p', { class: 'eyebrow', text: 'Improved prompt' }));
      improvedCard.appendChild(el('p', { class: 'improved-text', text: payload.improvedPrompt }));
      if (payload.notes) improvedCard.appendChild(el('p', { class: 'notes', text: payload.notes }));
      resultsEl.appendChild(improvedCard);
    } else if (payload.notes) {
      const noteCard = el('div', { class: 'improved-card' });
      noteCard.appendChild(el('p', { class: 'notes', text: payload.notes }));
      resultsEl.appendChild(noteCard);
    }

    // One card per sheet.
    const sheets = Array.isArray(payload.spec && payload.spec.sheets) ? payload.spec.sheets : [];
    sheets.forEach((sheet) => resultsEl.appendChild(renderSheetCard(sheet)));

    // Edit affordance — enter edit mode on the current spec.
    const editRow = el('div', { class: 'results-actions' });
    const editBtn = el('button', { class: 'btn btn-secondary', attrs: { type: 'button' }, text: '✏️ Edit / add data' });
    editBtn.addEventListener('click', () => enterEditMode(currentSpec, currentSpec && currentSpec.title));
    editRow.appendChild(editBtn);
    resultsEl.appendChild(editRow);
  }

  function renderSheetCard(sheet) {
    const card = el('div', { class: 'sheet-card' });
    const columns = Array.isArray(sheet.columns) ? sheet.columns : [];
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    const charts = Array.isArray(sheet.charts) ? sheet.charts : [];

    // Head: name + row/col counts.
    const head = el('div', { class: 'sheet-head' });
    head.appendChild(el('h3', { class: 'sheet-name', text: sheet.name || 'Sheet' }));
    const metaBits = [];
    metaBits.push(columns.length + (columns.length === 1 ? ' column' : ' columns'));
    metaBits.push(rows.length + (rows.length === 1 ? ' row' : ' rows'));
    head.appendChild(el('span', { class: 'sheet-meta', text: metaBits.join(' · ') }));
    card.appendChild(head);

    if (sheet.description) {
      card.appendChild(el('p', { class: 'sheet-desc', text: sheet.description }));
    }

    // Badges: formula columns + charts.
    const badges = el('div', { class: 'badges' });
    columns.forEach((col) => {
      if (col && col.type === 'formula') {
        const b = el('span', { class: 'badge badge-fx' });
        b.appendChild(el('span', { class: 'badge-key', text: 'fx' }));
        b.appendChild(el('span', { text: ' ' + (col.header || 'Formula') }));
        badges.appendChild(b);
      }
    });
    charts.forEach((chart) => {
      if (!chart) return;
      const kind = String(chart.type || 'chart').toUpperCase();
      const b = el('span', { class: 'badge badge-chart' });
      b.appendChild(el('span', { class: 'badge-key', text: kind + ':' }));
      b.appendChild(el('span', { text: ' ' + (chart.title || 'Chart') }));
      badges.appendChild(b);
    });
    if (badges.childNodes.length) card.appendChild(badges);

    // Preview table: headers + up to 8 sample rows.
    if (columns.length) {
      card.appendChild(renderTable(columns, rows));
    }

    return card;
  }

  function renderTable(columns, rows) {
    const scroll = el('div', { class: 'table-scroll' });
    const table = el('table', { class: 'preview' });

    // Header.
    const thead = el('thead');
    const htr = el('tr');
    columns.forEach((col) => {
      htr.appendChild(el('th', { attrs: { scope: 'col' }, text: (col && col.header) ? col.header : '' }));
    });
    thead.appendChild(htr);
    table.appendChild(thead);

    // Body — up to 8 rows.
    const tbody = el('tbody');
    const MAX = 8;
    const shown = rows.slice(0, MAX);

    if (!shown.length) {
      const tr = el('tr');
      const td = el('td', {
        class: 'empty',
        text: 'Empty template — no sample rows',
        attrs: { colspan: String(columns.length) },
      });
      tr.appendChild(td);
      tbody.appendChild(tr);
    } else {
      shown.forEach((row, rIdx) => {
        const tr = el('tr');
        for (let c = 0; c < columns.length; c++) {
          const col = columns[c];
          // Formula columns: the generator computes them; show a derived hint.
          if (col && col.type === 'formula') {
            tr.appendChild(el('td', { class: 'empty', text: 'fx' }));
            continue;
          }
          const val = Array.isArray(row) ? row[c] : undefined;
          tr.appendChild(formatCell(val, col));
        }
        tbody.appendChild(tr);
      });

      if (rows.length > MAX) {
        const tr = el('tr', { class: 'more-row' });
        tr.appendChild(el('td', {
          text: '+ ' + (rows.length - MAX) + ' more rows',
          attrs: { colspan: String(columns.length) },
        }));
        tbody.appendChild(tr);
      }
    }

    table.appendChild(tbody);
    scroll.appendChild(table);
    return scroll;
  }

  function formatCell(val, col) {
    if (val === null || val === undefined || val === '') {
      return el('td', { class: 'empty', text: '—' });
    }
    if (typeof val === 'boolean') {
      return el('td', { text: val ? 'TRUE' : 'FALSE' });
    }
    if (typeof val === 'number' && col && col.type === 'percent') {
      // spec: percent values are fractions (0.25 -> 25.0%)
      return el('td', { text: (val * 100).toFixed(1) + '%' });
    }
    return el('td', { text: String(val) });
  }

  /* ============================================================
     Secondary action: /api/generate -> download .xlsx (binary).
     ============================================================ */
  downloadBtn.addEventListener('click', onDownload);

  async function onDownload() {
    if (busy || !currentSpec) return;
    const ok = await generateAndDownload(currentSpec, downloadBtn);
    if (ok) saveCurrentToLibrary(false);  // a downloaded file is remembered in the library
  }

  // Build + download an .xlsx from any spec. `btn` (optional) shows a busy label.
  async function generateAndDownload(spec, btn) {
    if (!spec) return false;
    clearError();
    const original = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = 'Building...'; }
    const filename = slugify((spec && spec.title) || 'spreadsheet');
    try {
      const res = await fetchWithTimeout('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ spec, filename }),
      }, 60000);

      // Errors come back as JSON { error }.
      const ctype = res.headers.get('Content-Type') || '';
      if (!res.ok || ctype.includes('application/json')) {
        let msg = 'Could not generate the file (' + res.status + ').';
        try { const p = await res.json(); if (p && p.error) msg = p.error; } catch (_) {}
        showError(msg);
        return false;
      }

      const blob = await res.blob();
      const name = filenameFromDisposition(res.headers.get('Content-Disposition')) || (filename + '.xlsx');
      triggerDownload(blob, name);
      return true;
    } catch (err) {
      if (err && err.name === 'AbortError') showError('Building the file took too long. Please try again.');
      else showError('Could not reach the server to build your file. Try again.');
      return false;
    } finally {
      // Main download button follows the spec gate; per-row library buttons just re-enable.
      if (btn) { btn.textContent = original; btn.disabled = (btn === downloadBtn) ? !currentSpec : false; }
    }
  }

  function triggerDownload(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = el('a', { attrs: { href: url, download: name } });
    document.body.appendChild(a);
    a.click();
    a.remove();
    // Revoke after the browser has had a chance to start the download.
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  function filenameFromDisposition(header) {
    if (!header) return null;
    // filename*=UTF-8''... takes precedence over plain filename.
    const star = /filename\*\s*=\s*[^']*''([^;]+)/i.exec(header);
    if (star && star[1]) {
      try { return sanitizeName(decodeURIComponent(star[1])); } catch (_) {}
    }
    const plain = /filename\s*=\s*"?([^";]+)"?/i.exec(header);
    if (plain && plain[1]) return sanitizeName(plain[1].trim());
    return null;
  }

  function sanitizeName(name) {
    // Strip any path components a server might (mis)send; keep it a bare filename.
    // Allow letters, digits, dash, underscore and dot; collapse the rest to '_'.
    let base = String(name).split(/[\\/]/).pop() || '';
    base = base.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^[._]+/, '').trim();
    if (!base) return 'spreadsheet.xlsx';
    if (!/\.xlsx$/i.test(base)) base += '.xlsx';
    return base;
  }

  function slugify(s) {
    const out = String(s)
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 60);
    return out || 'spreadsheet';
  }

  /* ============================================================
     Library — past spreadsheets saved in the browser (IndexedDB).
     Each record: { id, title, spec, updatedAt }.
     ============================================================ */
  const DB_NAME = 'sheetgenie';
  const STORE = 'sheets';

  function idb() {
    return new Promise((resolve, reject) => {
      const r = indexedDB.open(DB_NAME, 1);
      r.onupgradeneeded = () => { r.result.createObjectStore(STORE, { keyPath: 'id' }); };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
  }
  function idbDo(mode, fn) {
    return idb().then((db) => new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const rq = fn(tx.objectStore(STORE));
      let out;
      if (rq) rq.onsuccess = () => { out = rq.result; };
      tx.oncomplete = () => resolve(out);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    }));
  }
  const libPut = (rec) => idbDo('readwrite', (s) => s.put(rec));
  const libAll = () => idbDo('readonly', (s) => s.getAll());
  const libGet = (id) => idbDo('readonly', (s) => s.get(id));
  const libDel = (id) => idbDo('readwrite', (s) => s.delete(id));

  function genId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }

  async function saveCurrentToLibrary(announce) {
    if (!currentSpec) return;
    try {
      if (!currentLibId) currentLibId = genId();
      await libPut({ id: currentLibId, title: currentSpec.title || 'Spreadsheet', spec: currentSpec, updatedAt: Date.now() });
      await refreshLibrary();
      if (announce) showOk('Saved to your library.');
    } catch (_) {
      showError('Could not save to the library on this device.');
    }
  }

  async function refreshLibrary() {
    let items = [];
    try { items = await libAll(); } catch (_) { items = []; }
    items.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    libCount.textContent = String(items.length);
    libCount.hidden = items.length === 0;
    renderLibList(items);
  }

  function renderLibList(items) {
    libList.replaceChildren();
    libEmpty.hidden = items.length > 0;
    items.forEach((rec) => {
      const info = el('div', { class: 'lib-info' }, [
        el('span', { class: 'lib-name', text: rec.title || 'Spreadsheet' }),
        el('span', { class: 'lib-meta', text: libMeta(rec) }),
      ]);
      const openB = el('button', { class: 'lib-act', attrs: { type: 'button' }, text: 'Open' });
      openB.addEventListener('click', () => openFromLibrary(rec.id));
      const dlB = el('button', { class: 'lib-act', attrs: { type: 'button' }, text: 'Download' });
      dlB.addEventListener('click', () => generateAndDownload(rec.spec, dlB));
      const delB = el('button', { class: 'lib-act lib-del', attrs: { type: 'button', 'aria-label': 'Delete ' + (rec.title || 'spreadsheet') }, text: 'Delete' });
      delB.addEventListener('click', async () => {
        try { await libDel(rec.id); } catch (_) {}
        if (currentLibId === rec.id) currentLibId = null;
        refreshLibrary();
      });
      const acts = el('div', { class: 'lib-acts' }, [openB, dlB, delB]);
      libList.appendChild(el('li', { class: 'lib-item' }, [info, acts]));
    });
  }

  function libMeta(rec) {
    const n = Array.isArray(rec.spec && rec.spec.sheets) ? rec.spec.sheets.length : 0;
    const when = rec.updatedAt ? new Date(rec.updatedAt).toLocaleDateString() : '';
    return (n + (n === 1 ? ' sheet' : ' sheets')) + (when ? ' · ' + when : '');
  }

  async function openFromLibrary(id) {
    let rec = null;
    try { rec = await libGet(id); } catch (_) {}
    if (!rec || !rec.spec) { showError('Could not open that spreadsheet.'); return; }
    currentSpec = rec.spec;
    currentLibId = rec.id;
    downloadBtn.disabled = false;
    saveLibBtn.disabled = false;
    enterEditMode(rec.spec, rec.title);
    renderResults({
      improvedPrompt: null,
      notes: 'Loaded “' + (rec.title || 'spreadsheet') + '”. Describe a change or add data above and tap Apply changes — or just Download.',
      spec: rec.spec,
    });
    setLibraryOpen(false);
    const cap = document.querySelector('.capture');
    if (cap) cap.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function enterEditMode(spec, title) {
    if (!spec) return;
    baseSpec = spec;
    editTitle.textContent = title || 'this spreadsheet';
    editBanner.hidden = false;
    promptEl.placeholder = 'What should I change? e.g. add a Tax column, sort by date, append the new data';
    refreshPrimaryLabel();  // primary button becomes "Apply changes"
    promptEl.focus();
  }

  function newSpreadsheet() {
    baseSpec = null;
    currentLibId = null;
    currentSpec = null;
    editBanner.hidden = true;
    promptEl.placeholder = PROMPT_PLACEHOLDER;
    promptEl.value = '';
    dataInput.value = '';
    attachments = [];
    renderAttachments();
    updateDataBadge();
    resultsEl.replaceChildren();
    downloadBtn.disabled = true;
    saveLibBtn.disabled = true;
    setBusy(false);
    autosize();
    promptEl.focus();
  }

  function setLibraryOpen(open) {
    librarySection.hidden = !open;
    libraryToggle.setAttribute('aria-expanded', String(open));
  }
  libraryToggle.addEventListener('click', async () => {
    const show = librarySection.hidden;
    if (show) await refreshLibrary();
    setLibraryOpen(show);
    if (show) librarySection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
  libraryClose.addEventListener('click', () => setLibraryOpen(false));
  saveLibBtn.addEventListener('click', () => saveCurrentToLibrary(true));
  newSheetBtn.addEventListener('click', newSpreadsheet);

  refreshLibrary();

  /* ============================================================
     Register the service worker (PWA / offline shell).
     ============================================================ */
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch(() => { /* offline-shell is best-effort */ });
    });
  }
})();
