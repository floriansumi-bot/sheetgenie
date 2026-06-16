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
  const improveBtn = $('improve');
  const downloadBtn = $('download');
  const toastEl = $('toast');
  const resultsEl = $('results');
  const themeToggle = $('themeToggle');

  /* ---------- app state ---------- */
  let currentSpec = null;     // last valid SpreadsheetSpec from /api/improve
  let busy = false;           // a network call is in flight
  let toastTimer = null;

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

  function hasData() { return dataInput.value.trim().length > 0; }

  dataInput.addEventListener('input', () => {
    dataBadge.hidden = !hasData();
  });

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
    toastEl.hidden = false;
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.hidden = true; }, 7000);
  }
  function clearError() {
    toastEl.hidden = true;
    toastEl.textContent = '';
    if (toastTimer) clearTimeout(toastTimer);
  }

  /* ============================================================
     Loading state for the primary action.
     ============================================================ */
  function setBusy(on) {
    busy = on;
    improveBtn.classList.toggle('is-loading', on);
    improveBtn.disabled = on;
    // Download stays disabled while busy or until we have a spec.
    downloadBtn.disabled = on || !currentSpec;
    const label = improveBtn.querySelector('.btn-label');
    if (label) label.textContent = on ? 'Working...' : 'Improve & preview';
  }

  /* ============================================================
     Primary action: /api/improve
     ============================================================ */
  improveBtn.addEventListener('click', onImprove);

  async function onImprove() {
    if (busy) return;
    const prompt = promptEl.value.trim();
    if (!prompt) {
      showError('Tell me what spreadsheet you need first.');
      promptEl.focus();
      return;
    }

    clearError();
    setBusy(true);

    const body = {
      prompt,
      hasData: hasData(),
      data: hasData() ? dataInput.value : null,
    };

    try {
      const res = await fetchWithTimeout('/api/improve', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
        body: JSON.stringify(body),
      }, 90000);

      let payload = null;
      try { payload = await res.json(); } catch (_) { payload = null; }

      if (!res.ok || !payload || payload.error) {
        const msg = (payload && payload.error) ? payload.error
          : 'The improve step failed (' + res.status + '). Please try again.';
        showError(msg);
        return;
      }

      if (!payload.spec || typeof payload.spec !== 'object' || !Array.isArray(payload.spec.sheets)) {
        showError('The assistant returned an unexpected result. Please try again.');
        return;
      }

      currentSpec = payload.spec;
      downloadBtn.disabled = false;
      renderResults(payload);
      resultsEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
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

  /* ============================================================
     Render preview — all via createElement/textContent (XSS-safe).
     ============================================================ */
  function renderResults(payload) {
    resultsEl.replaceChildren();

    // Improved prompt + notes card.
    const improvedCard = el('div', { class: 'improved-card' });
    improvedCard.appendChild(el('p', { class: 'eyebrow', text: 'Improved prompt' }));
    improvedCard.appendChild(el('p', {
      class: 'improved-text',
      text: payload.improvedPrompt || '(no improved prompt returned)',
    }));
    if (payload.notes) {
      improvedCard.appendChild(el('p', { class: 'notes', text: payload.notes }));
    }
    resultsEl.appendChild(improvedCard);

    // One card per sheet.
    const sheets = Array.isArray(payload.spec.sheets) ? payload.spec.sheets : [];
    sheets.forEach((sheet) => resultsEl.appendChild(renderSheetCard(sheet)));
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

    clearError();
    const original = downloadBtn.textContent;
    downloadBtn.disabled = true;
    downloadBtn.textContent = 'Building...';

    const filename = slugify(currentSpec.title || 'spreadsheet');
    const body = { spec: currentSpec, filename };

    try {
      const res = await fetchWithTimeout('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      }, 60000);

      // Errors come back as JSON { error }.
      const ctype = res.headers.get('Content-Type') || '';
      if (!res.ok || ctype.includes('application/json')) {
        let msg = 'Could not generate the file (' + res.status + ').';
        try {
          const payload = await res.json();
          if (payload && payload.error) msg = payload.error;
        } catch (_) {}
        showError(msg);
        return;
      }

      const blob = await res.blob();
      const name = filenameFromDisposition(res.headers.get('Content-Disposition'))
        || (slugify(currentSpec.title || 'spreadsheet') + '.xlsx');

      triggerDownload(blob, name);
    } catch (err) {
      if (err && err.name === 'AbortError') {
        showError('Building the file took too long. Please try again.');
      } else {
        showError('Could not reach the server to build your file. Try again.');
      }
    } finally {
      downloadBtn.textContent = original;
      downloadBtn.disabled = !currentSpec;
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
     Register the service worker (PWA / offline shell).
     ============================================================ */
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('sw.js').catch(() => { /* offline-shell is best-effort */ });
    });
  }
})();
