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
  const libNoMatch = $('libNoMatch');
  const libSearch = $('libSearch');
  const libExportBtn = $('libExport');
  const libImportBtn = $('libImportBtn');
  const libImportInput = $('libImportInput');
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
    const recIndicator = $('recIndicator');
    const recCanvas = $('recCanvas');
    if (!SR) {
      micBtn.hidden = true;
      micHint.hidden = false;
      return;
    }

    // On touch devices the Web Audio visualizer's getUserMedia stream competes with
    // SpeechRecognition for the single available microphone, which silently breaks
    // transcription on Android. So skip the waveform there and let recognition own the mic.
    const isTouch = matchMedia('(pointer: coarse)').matches
      || /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent || '');
    if (isTouch && recCanvas) recCanvas.hidden = true;

    let recog = null;
    let recording = false;     // user intent: the mic is on — OFF until the user taps the mic
    let manualStop = false;    // user tapped stop (don't auto-restart)
    let restarts = 0;
    let baseText = '';
    let sep = '';
    // Web Audio visualizer (the "you're being recorded" cue + permission gate).
    let audioStream = null, audioCtx = null, analyser = null, rafId = 0;

    function showRecUI(on) {
      if (recIndicator) recIndicator.hidden = !on;
      micBtn.setAttribute('aria-pressed', String(on));
      micBtn.setAttribute('aria-label', on ? 'Stop dictation' : 'Dictate with your voice');
    }

    async function startVisualizer() {
      try {
        if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return;
        audioStream = await navigator.mediaDevices.getUserMedia({ audio: true });
        const AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) return;
        audioCtx = new AC();
        const srcNode = audioCtx.createMediaStreamSource(audioStream);
        analyser = audioCtx.createAnalyser();
        analyser.fftSize = 256;
        analyser.smoothingTimeConstant = 0.7;
        srcNode.connect(analyser);
        drawWave();
      } catch (_) {
        // Visualizer is best-effort; transcription can still run without it.
      }
    }

    function drawWave() {
      if (!analyser || !recCanvas || !recCanvas.getContext) return;
      const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
      const ctx = recCanvas.getContext('2d');
      const W = recCanvas.width, H = recCanvas.height;
      const bins = analyser.frequencyBinCount;
      const data = new Uint8Array(bins);
      const BARS = 32;
      const accent = (getComputedStyle(document.documentElement)
        .getPropertyValue('--accent-1') || '#8b5cf6').trim();
      function frame() {
        rafId = requestAnimationFrame(frame);
        analyser.getByteFrequencyData(data);
        ctx.clearRect(0, 0, W, H);
        const bw = W / BARS;
        for (let i = 0; i < BARS; i++) {
          const v = data[Math.floor((i / BARS) * bins)] / 255;       // 0..1
          const h = Math.max(2, v * H);
          ctx.fillStyle = accent;
          ctx.globalAlpha = 0.3 + 0.7 * v;
          ctx.fillRect(i * bw + 1, (H - h) / 2, Math.max(1, bw - 2), h);
        }
        ctx.globalAlpha = 1;
      }
      if (reduce) {          // draw one static frame, don't animate
        analyser.getByteFrequencyData(data);
        ctx.clearRect(0, 0, W, H);
        ctx.fillStyle = accent;
        ctx.fillRect(0, H / 2 - 1, W, 2);
        return;
      }
      frame();
    }

    function stopVisualizer() {
      if (rafId) { cancelAnimationFrame(rafId); rafId = 0; }
      if (recCanvas && recCanvas.getContext) {
        recCanvas.getContext('2d').clearRect(0, 0, recCanvas.width, recCanvas.height);
      }
      if (audioStream) { audioStream.getTracks().forEach((t) => { try { t.stop(); } catch (_) {} }); audioStream = null; }
      if (audioCtx) { try { audioCtx.close(); } catch (_) {} audioCtx = null; }
      analyser = null;
    }

    function startRecog() {
      recog = new SR();
      try { recog.lang = navigator.language || 'en-US'; } catch (_) {}
      // Android/iOS don't support continuous recognition — it ends instantly and never
      // delivers a result. Use single-utterance there and let onend restart it so the
      // user can keep dictating across pauses; desktop keeps true continuous mode.
      recog.continuous = !isTouch;
      recog.interimResults = true;

      recog.onresult = (event) => {
        let finalChunk = '', interimChunk = '';
        for (let i = event.resultIndex; i < event.results.length; i++) {
          const res = event.results[i];
          if (res.isFinal) finalChunk += res[0].transcript;
          else interimChunk += res[0].transcript;
        }
        if (finalChunk || interimChunk) restarts = 0;   // real audio came through — reset the silence guard
        if (finalChunk) baseText = baseText + sep + finalChunk.trim();
        if (finalChunk && !/\s$/.test(baseText)) sep = ' ';
        const liveSep = baseText && interimChunk && !/\s$/.test(baseText) ? ' ' : '';
        promptEl.value = baseText + (interimChunk ? liveSep + interimChunk : '');
        autosize();
      };

      recog.onerror = (event) => {
        const code = event && event.error;
        if (code === 'no-speech' || code === 'aborted') return;  // transient — onend restarts
        if (code === 'not-allowed' || code === 'service-not-allowed') {
          showError('Microphone access is blocked — allow it in your browser settings, or type instead.');
        } else if (code === 'audio-capture') {
          showError('No microphone was found. You can type instead.');
        } else if (code === 'network') {
          showError('Voice needs an internet connection right now. You can type instead.');
        } else {
          showError('Voice had a hiccup — tap the mic to try again, or type.');
        }
        endRecording();
      };

      recog.onend = () => {
        // Recognition ends itself on every pause (and instantly on mobile). Restart it
        // while the user still wants to dictate — bounded so sustained silence can't loop
        // forever (onresult resets `restarts`, so real dictation is never cut off).
        if (recording && !manualStop && restarts < 60) {
          restarts++;
          setTimeout(() => {
            if (recording && !manualStop) { try { recog.start(); } catch (_) {} }
          }, 250);
          return;
        }
        // User tapped stop, or we gave up after sustained silence: ALWAYS clean up so the
        // "Listening…" indicator can never get stuck on with nothing actually recording.
        recording = false;
        stopVisualizer();
        showRecUI(false);
      };

      try { recog.start(); } catch (_) { /* start() can throw right after stop(); onend retries */ }
    }

    function beginRecording() {
      recording = true; manualStop = false; restarts = 0;
      baseText = promptEl.value;
      sep = baseText && !/\s$/.test(baseText) ? ' ' : '';
      showRecUI(true);
      if (!isTouch) startVisualizer();   // desktop only — see isTouch note above
      startRecog();
    }

    function endRecording() {
      manualStop = true;
      recording = false;
      if (recog) { try { recog.stop(); } catch (_) {} }
      stopVisualizer();
      showRecUI(false);
    }

    micBtn.addEventListener('click', () => {
      if (recording) endRecording(); else beginRecording();
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
  let pendingClarifications = null;   // answers in play, so a layout pick can resend them
  let lastChosenLayout = null;        // remembered so a queued (backup) job can reuse it

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
  // `chosenLayout` is set when the user picks one of the proposed layouts.
  // `directBuild` skips the layout-proposal step (used by Regenerate, where the
  // user has already refined the prompt and wants the spec rebuilt straight away).
  async function runImprove(clarifications, chosenLayout, directBuild) {
    if (busy) return;
    clearError();
    setBusy(true);
    pendingClarifications = clarifications || null;
    lastChosenLayout = chosenLayout || null;
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
      chosenLayout: chosenLayout || null,
      directBuild: !!directBuild,
      locale: (navigator.language || ''),
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
        // All cloud providers busy + a backup worker is configured → offer the
        // "email me the finished file" path instead of a dead-end error.
        if (payload && payload.canQueue) { renderQueueOffer(payload.error); return; }
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
    // "layouts": the model proposed a few structures — let the user pick one to build.
    if (payload.status === 'layouts' &&
        Array.isArray(payload.layouts) && payload.layouts.length) {
      renderLayouts(payload);
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

  // Render the 2-3 proposed layouts as choosable cards. Picking one re-runs improve
  // with that layout (and any answers already given) to build the full spec.
  function renderLayouts(payload) {
    resultsEl.replaceChildren();

    const wrap = el('div', { class: 'layouts-card' });
    wrap.appendChild(el('p', { class: 'eyebrow', text: 'Pick a layout' }));
    wrap.appendChild(el('p', {
      class: 'notes',
      text: payload.notes || 'A few ways to organise this — pick the one you like and I’ll build it.',
    }));

    const grid = el('div', { class: 'layouts-grid' });
    payload.layouts.slice(0, 3).forEach((lay, i) => {
      const card = el('div', { class: 'layout-card' });
      card.appendChild(el('h3', { class: 'layout-title', text: (lay && lay.title) ? lay.title : ('Layout ' + (i + 1)) }));
      if (lay && lay.summary) card.appendChild(el('p', { class: 'layout-summary', text: lay.summary }));

      const sheets = Array.isArray(lay && lay.sheets) ? lay.sheets : [];
      sheets.forEach((sh) => {
        const box = el('div', { class: 'layout-sheet' });
        box.appendChild(el('span', { class: 'layout-sheet-name', text: (sh && sh.name) ? sh.name : 'Sheet' }));
        const cols = el('div', { class: 'layout-cols' });
        (Array.isArray(sh && sh.columns) ? sh.columns : []).forEach((c) => {
          cols.appendChild(el('span', { class: 'layout-col', text: String(c) }));
        });
        box.appendChild(cols);
        card.appendChild(box);
      });

      const pick = el('button', { class: 'btn btn-primary layout-pick', attrs: { type: 'button' }, text: 'Build this layout →' });
      pick.addEventListener('click', async () => {
        const allPicks = resultsEl.querySelectorAll('.layout-pick');
        allPicks.forEach((b) => { b.disabled = true; });
        card.classList.add('chosen');
        const label = pick.textContent;
        pick.textContent = 'Building…';
        await runImprove(pendingClarifications, lay);
        if (pick.isConnected) {
          allPicks.forEach((b) => { b.disabled = false; });
          card.classList.remove('chosen');
          pick.textContent = label;
        }
      });
      card.appendChild(pick);
      grid.appendChild(card);
    });

    wrap.appendChild(grid);
    resultsEl.appendChild(wrap);
  }

  // Every cloud provider was busy and a self-hosted backup is configured. Collect an
  // email and queue the job (POST /api/queue): the backup server builds the file and
  // emails it in a few minutes, so we don't block on the platform's request limit.
  function renderQueueOffer(message) {
    resultsEl.replaceChildren();
    const card = el('div', { class: 'queue-card' });
    card.appendChild(el('p', { class: 'eyebrow', text: 'Fast AI is busy' }));
    card.appendChild(el('p', { class: 'notes',
      text: message || 'Our fast AI is busy right now. Our backup server can still build this — it takes a few minutes.' }));

    const row = el('div', { class: 'queue-row' });
    const input = el('input', { class: 'q-input', attrs: {
      type: 'email', inputmode: 'email', autocomplete: 'email', enterkeyhint: 'send',
      placeholder: 'you@example.com', 'aria-label': 'Your email address' } });
    const btn = el('button', { class: 'btn btn-primary', attrs: { type: 'button' }, text: 'Email it to me' });
    row.appendChild(input);
    row.appendChild(btn);
    card.appendChild(row);
    card.appendChild(el('p', { class: 'attach-hint',
      text: 'We’ll email the finished spreadsheet here in a few minutes. (Photos/PDFs aren’t read on the backup path.)' }));

    const submit = async () => {
      const email = input.value.trim();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
        showError('Please enter a valid email address.'); input.focus(); return;
      }
      btn.disabled = true; input.disabled = true;
      const lbl = btn.textContent; btn.textContent = 'Queueing…';
      try {
        const res = await fetchWithTimeout('/api/queue', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
          body: JSON.stringify({
            prompt: pendingPrompt, data: pendingData,
            clarifications: pendingClarifications, chosenLayout: lastChosenLayout,
            locale: (navigator.language || ''), email: email,
          }),
        }, 30000);
        let p = null; try { p = await res.json(); } catch (_) { p = null; }
        if (res.ok && p && p.queued) {
          resultsEl.replaceChildren();
          const done = el('div', { class: 'queue-card' });
          done.appendChild(el('p', { class: 'eyebrow', text: 'On its way ✓' }));
          done.appendChild(el('p', { class: 'notes',
            text: 'Got it! The backup server is building your spreadsheet and will email it to '
                  + email + ' in a few minutes. You can safely close this page.' }));
          resultsEl.appendChild(done);
        } else {
          btn.disabled = false; input.disabled = false; btn.textContent = lbl;
          showError((p && p.error) ? p.error : 'Could not queue the job. Please try again later.');
        }
      } catch (_) {
        btn.disabled = false; input.disabled = false; btn.textContent = lbl;
        showError('Could not reach the server. Please try again later.');
      }
    };
    btn.addEventListener('click', submit);
    input.addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); });

    resultsEl.appendChild(card);
    resultsEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
    setTimeout(() => { try { input.focus(); } catch (_) {} }, 50);
  }

  /* ============================================================
     Render preview — all via createElement/textContent (XSS-safe).
     ============================================================ */
  function renderResults(payload) {
    resultsEl.replaceChildren();

    // Improved prompt — editable, with a Regenerate button — plus notes. (Loaded
    // library files have no improvedPrompt, so they show just the note.)
    if (payload.improvedPrompt) {
      const improvedCard = el('div', { class: 'improved-card' });
      improvedCard.appendChild(el('p', { class: 'eyebrow', text: 'Improved prompt — tweak it, then regenerate' }));
      const ta = el('textarea', { class: 'improved-edit', attrs: { rows: '3', 'aria-label': 'Improved prompt (editable)' } });
      ta.value = payload.improvedPrompt;
      improvedCard.appendChild(ta);
      if (payload.notes) improvedCard.appendChild(el('p', { class: 'notes', text: payload.notes }));
      const regenBtn = el('button', { class: 'btn btn-secondary improved-regen', attrs: { type: 'button' }, text: '↻ Regenerate from this' });
      regenBtn.addEventListener('click', () => {
        const refined = ta.value.trim();
        if (!refined) { showError('The improved prompt is empty — type what you want.'); return; }
        pendingPrompt = refined;
        pendingBaseSpec = null;   // a refined full prompt -> a fresh spec (keeps any data/files added)
        runImprove(null, null, true);   // build directly — don't re-propose layouts
      });
      improvedCard.appendChild(regenBtn);
      resultsEl.appendChild(improvedCard);
    } else if (payload.notes) {
      const noteCard = el('div', { class: 'improved-card' });
      noteCard.appendChild(el('p', { class: 'notes', text: payload.notes }));
      resultsEl.appendChild(noteCard);
    }

    // One card per sheet.
    const sheets = Array.isArray(payload.spec && payload.spec.sheets) ? payload.spec.sheets : [];
    sheets.forEach((sheet) => resultsEl.appendChild(renderSheetCard(sheet)));

    // Keep refining the finished spreadsheet with another prompt and/or new data.
    const editRow = el('div', { class: 'results-actions' });
    const editBtn = el('button', { class: 'btn btn-secondary', attrs: { type: 'button' }, text: '✏️ Edit with another prompt or data' });
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
      // Preserve a user-assigned (renamed) title: only fall back to the spec title
      // for a brand-new record, so auto-saves (download / edit) never clobber a rename.
      let title = currentSpec.title || 'Spreadsheet';
      if (currentLibId) {
        let existing = null;
        try { existing = await libGet(currentLibId); } catch (_) {}
        if (existing && existing.title) title = existing.title;
      } else {
        currentLibId = genId();
      }
      await libPut({ id: currentLibId, title: title, spec: currentSpec, updatedAt: Date.now() });
      await refreshLibrary();
      if (announce) showOk('Saved to your library.');
    } catch (_) {
      showError('Could not save to the library on this device.');
    }
  }

  // The full, sorted record set is held in memory so the search box can filter it
  // live without re-querying IndexedDB on every keystroke.
  let libItems = [];

  async function refreshLibrary() {
    let items = [];
    try { items = await libAll(); } catch (_) { items = []; }
    items.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    libItems = items;
    libCount.textContent = String(items.length);
    libCount.hidden = items.length === 0;
    renderLibList();
  }

  // Render the (optionally filtered) list. Empty-library and no-match get distinct
  // messages so a search that hides everything doesn't look like an empty library.
  function renderLibList() {
    // Don't tear down the list while a rename input is open (replaceChildren would
    // remove it and fire its blur->commit). Skip the rebuild until rename finishes.
    if (libList.querySelector('.lib-rename')) return;
    const q = (libSearch && libSearch.value ? libSearch.value : '').trim().toLowerCase();
    const shown = q
      ? libItems.filter((rec) => String(rec.title || '').toLowerCase().includes(q))
      : libItems;

    libList.replaceChildren();
    libEmpty.hidden = libItems.length > 0;
    libNoMatch.hidden = !(libItems.length > 0 && shown.length === 0);

    shown.forEach((rec) => {
      const name = el('span', { class: 'lib-name', text: rec.title || 'Spreadsheet' });
      const info = el('div', { class: 'lib-info' }, [
        name,
        el('span', { class: 'lib-meta', text: libMeta(rec) }),
      ]);
      const openB = el('button', { class: 'lib-act', attrs: { type: 'button' }, text: 'Open' });
      openB.addEventListener('click', () => openFromLibrary(rec.id));
      const dlB = el('button', { class: 'lib-act', attrs: { type: 'button' }, text: 'Download' });
      dlB.addEventListener('click', () => generateAndDownload(rec.spec, dlB));
      const renameB = el('button', { class: 'lib-act', attrs: { type: 'button', 'aria-label': 'Rename ' + (rec.title || 'spreadsheet') }, text: 'Rename' });
      renameB.addEventListener('click', () => beginRename(rec, info, name, renameB));
      const dupB = el('button', { class: 'lib-act', attrs: { type: 'button', 'aria-label': 'Duplicate ' + (rec.title || 'spreadsheet') }, text: 'Duplicate' });
      dupB.addEventListener('click', () => duplicateRecord(rec, dupB));
      const delB = el('button', { class: 'lib-act lib-del', attrs: { type: 'button', 'aria-label': 'Delete ' + (rec.title || 'spreadsheet') }, text: 'Delete' });
      delB.addEventListener('click', async () => {
        try { await libDel(rec.id); } catch (_) {}
        if (currentLibId === rec.id) currentLibId = null;
        refreshLibrary();
      });
      const acts = el('div', { class: 'lib-acts' }, [openB, dlB, renameB, dupB, delB]);
      libList.appendChild(el('li', { class: 'lib-item' }, [info, acts]));
    });
  }

  // RENAME — swap the title label for an inline input; commit on Enter/blur, cancel
  // on Escape. Persists via libPut (title + fresh updatedAt) and keeps state in sync.
  function beginRename(rec, info, nameEl, triggerBtn) {
    if (info.querySelector('.lib-rename')) return;  // already editing this row
    const input = el('input', {
      class: 'lib-rename',
      attrs: { type: 'text', 'aria-label': 'New title', value: '' },
    });
    input.value = rec.title || 'Spreadsheet';
    nameEl.replaceWith(input);
    input.focus();
    input.select();

    let done = false;
    const cancel = () => {
      if (done) return;
      done = true;
      input.replaceWith(nameEl);
      if (triggerBtn && triggerBtn.isConnected) triggerBtn.focus();
    };
    const commit = async () => {
      if (done) return;
      const next = input.value.trim();
      if (!next || next === (rec.title || 'Spreadsheet')) { cancel(); return; }
      done = true;
      // Remove the editing input BEFORE refreshing so renderLibList's
      // "rename in progress" guard doesn't skip this rebuild. (That guard is
      // only meant to protect against a SEARCH re-render, which keeps the input
      // mounted, from tearing down an in-progress rename.)
      if (input.isConnected) input.replaceWith(nameEl);
      try {
        await libPut({ id: rec.id, title: next, spec: rec.spec, updatedAt: Date.now() });
        showOk('Renamed.');
      } catch (_) {
        showError('Could not rename on this device.');
      }
      await refreshLibrary();
    };

    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); }
      else if (e.key === 'Escape') { e.preventDefault(); cancel(); }
    });
    input.addEventListener('blur', commit);
  }

  // DUPLICATE — write a copy under a new id with " (copy)" appended to the title.
  async function duplicateRecord(rec, btn) {
    if (btn) btn.disabled = true;
    try {
      await libPut({
        id: genId(),
        title: ((rec.title || 'Spreadsheet') + ' (copy)').slice(0, 200),
        spec: rec.spec,
        updatedAt: Date.now(),
      });
      await refreshLibrary();
      showOk('Duplicated.');
    } catch (_) {
      showError('Could not duplicate on this device.');
    } finally {
      if (btn && btn.isConnected) btn.disabled = false;
    }
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

  /* ---------- SEARCH — live, case-insensitive title filter ---------- */
  if (libSearch) {
    libSearch.addEventListener('input', renderLibList);
    // Native clear (the "x" on type=search) fires 'search'.
    libSearch.addEventListener('search', renderLibList);
  }

  /* ---------- EXPORT — download a JSON backup of every record ---------- */
  if (libExportBtn) libExportBtn.addEventListener('click', exportLibrary);
  async function exportLibrary() {
    let items = [];
    try { items = await libAll(); } catch (_) { items = []; }
    if (!items.length) { showError('Your library is empty — nothing to export yet.'); return; }
    const backup = {
      app: 'sheetgenie',
      kind: 'library-backup',
      version: 1,
      exportedAt: new Date().toISOString(),
      records: items.map((r) => ({
        id: r.id, title: r.title, spec: r.spec, updatedAt: r.updatedAt,
      })),
    };
    let json;
    try { json = JSON.stringify(backup, null, 2); }
    catch (_) { showError('Could not prepare the backup file.'); return; }
    const blob = new Blob([json], { type: 'application/json' });
    const stamp = new Date().toISOString().slice(0, 10);
    triggerDownload(blob, 'sheetgenie-library-' + stamp + '.json');
    showOk('Exported ' + items.length + (items.length === 1 ? ' spreadsheet.' : ' spreadsheets.'));
  }

  /* ---------- IMPORT — merge records from a JSON backup ---------- */
  if (libImportBtn) libImportBtn.addEventListener('click', () => libImportInput && libImportInput.click());
  if (libImportInput) libImportInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    libImportInput.value = '';  // allow re-importing the same file
    if (file) importLibrary(file);
  });

  // A record imports only if it has a usable spec (object with a sheets array). We
  // assign a fresh id when the incoming id is missing or already present, so an
  // import never silently overwrites an existing entry.
  function isImportableSpec(spec) {
    return spec && typeof spec === 'object' && Array.isArray(spec.sheets);
  }

  async function importLibrary(file) {
    if (file && file.size > 20 * 1024 * 1024) {
      showError('That backup file is too large to import.');
      return;
    }
    let text;
    try { text = await readAsText(file); }
    catch (_) { showError('Could not read that file.'); return; }

    let parsed;
    try { parsed = JSON.parse(text); }
    catch (_) { showError('That file is not valid JSON.'); return; }

    // Accept either our backup envelope ({records:[...]}) or a bare array of records.
    const records = Array.isArray(parsed) ? parsed
      : (parsed && Array.isArray(parsed.records)) ? parsed.records
      : null;
    if (!records) { showError('That does not look like a SheetGenie library backup.'); return; }

    let existingIds;
    try { existingIds = new Set((await libAll()).map((r) => r.id)); }
    catch (_) { existingIds = new Set(); }

    let added = 0;
    let skipped = 0;
    for (const rec of records) {
      if (!rec || !isImportableSpec(rec.spec)) { skipped++; continue; }
      let id = (typeof rec.id === 'string' && rec.id) ? rec.id : genId();
      if (existingIds.has(id)) id = genId();  // never clobber an existing record
      existingIds.add(id);
      const title = (typeof rec.title === 'string' && rec.title.trim())
        ? rec.title.trim().slice(0, 200) : 'Spreadsheet';
      const updatedAt = (typeof rec.updatedAt === 'number' && isFinite(rec.updatedAt))
        ? rec.updatedAt : Date.now();
      try { await libPut({ id, title, spec: rec.spec, updatedAt }); added++; }
      catch (_) { skipped++; }
    }

    await refreshLibrary();
    if (added) {
      showOk('Imported ' + added + (added === 1 ? ' spreadsheet' : ' spreadsheets')
        + (skipped ? ' (' + skipped + ' skipped).' : '.'));
    } else {
      showError('No spreadsheets could be imported from that file.');
    }
  }

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
