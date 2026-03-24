  const GEMINI_MODEL = 'gemini-2.5-flash';

  let currentMd = '';
  let currentFilename = 'copilot-chat.md';
  let uploadedFileContent = null;

  // ── Drag & drop ──────────────────────────────────────────────────────────────
  const dropZone = document.getElementById('dropZone');
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFileSelect(e.dataTransfer.files);
  });

  function handleFileSelect(files) {
    const file = files[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['html','htm','mhtml','mht'].includes(ext)) {
      showStatus('Please upload an .html, .htm, or .mhtml file.', 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = e => {
      uploadedFileContent = e.target.result;
      document.getElementById('fileName').textContent = file.name;
      document.getElementById('fileSize').textContent = formatBytes(file.size);
      document.getElementById('fileInfo').style.display = 'flex';
      currentFilename = file.name.replace(/\.(html?|mhtml?)$/i, '') + '-' + new Date().toISOString().slice(0,10) + '.md';
      hideStatus();
    };
    reader.onerror = () => showStatus('Failed to read file.', 'error');
    reader.readAsText(file);
  }

  function clearFile() {
    uploadedFileContent = null;
    document.getElementById('fileInput').value = '';
    document.getElementById('fileInfo').style.display = 'none';
    hideStatus();
  }

  function formatBytes(b) {
    if (b < 1024) return b + ' B';
    if (b < 1048576) return (b/1024).toFixed(1) + ' KB';
    return (b/1048576).toFixed(1) + ' MB';
  }

  // ── HTML-based extraction (runs in browser, no AI involved) ──────────────────
  //
  // Copilot's HTML uses these stable class fragments:
  //   User turns:     elements whose class contains "group/user-message"
  //   Copilot turns:  elements whose class contains "group/ai-message-item"
  //
  // We extract text from each in DOM order, handling <pre> code blocks specially,
  // then hand the already-labelled plain text to Gemini for formatting only.

  function getNodeText(node) {
    // Walk the node tree and build a plain-text representation.
    // <pre> blocks get fenced with ``` so Gemini knows to keep them verbatim.
    let out = '';
    for (const child of node.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        out += child.textContent;
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const tag = child.tagName.toLowerCase();
        if (tag === 'script' || tag === 'style' || tag === 'noscript') continue;
        if (tag === 'pre') {
          // Detect language hint from a sibling button or data attribute if present
          const langBtn = child.closest('[class*="rounded-xl"]') &&
                          child.closest('[class*="rounded-xl"]').querySelector('[class*="max-w-24"]');
          const lang = langBtn ? langBtn.textContent.trim() : '';
          out += '\n```' + lang + '\n' + child.textContent + '\n```\n';
        } else if (tag === 'br') {
          out += '\n';
        } else {
          out += getNodeText(child);
        }
      }
    }
    return out;
  }

  function extractTurnsFromHtml(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');

    // Remove screen-reader-only text (e.g. "You said" labels)
    doc.querySelectorAll('[class*="sr-only"]').forEach(el => el.remove());

    // Collect user and AI elements together, then sort by DOM position
    const userEls  = Array.from(doc.querySelectorAll('[class*="group/user-message"]'));
    const aiEls    = Array.from(doc.querySelectorAll('[class*="group/ai-message-item"]'));

    const all = [
      ...userEls.map(el => ({ el, role: 'USER' })),
      ...aiEls.map(el => ({ el, role: 'COPILOT' }))
    ];

    // Sort by DOM order using compareDocumentPosition
    all.sort((a, b) => {
      const rel = a.el.compareDocumentPosition(b.el);
      if (rel & Node.DOCUMENT_POSITION_FOLLOWING) return -1;
      if (rel & Node.DOCUMENT_POSITION_PRECEDING)  return  1;
      return 0;
    });

    const turns = [];
    for (const { el, role } of all) {
      const text = getNodeText(el).trim();
      if (text) turns.push({ role, text });
    }
    return turns;
  }

  // ── UI helpers ───────────────────────────────────────────────────────────────
  function setMode(mode, btn) {
    document.querySelectorAll('.mode-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.input-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('panel-' + mode).classList.add('active');
    hideStatus();
  }

  function setStep(n) {
    for (let i = 1; i <= 4; i++) {
      document.getElementById('s' + i).className =
        'step' + (i < n ? ' done' : i === n ? ' active' : '');
    }
  }

  function showStatus(msg, type) {
    const bar = document.getElementById('statusBar');
    bar.className = 'status-bar visible' + (type === 'error' ? ' error' : '');
    document.getElementById('statusMsg').textContent = msg;
    document.getElementById('statusLbl').textContent = type === 'error' ? 'Error' : 'Processing';
  }

  function hideStatus() {
    document.getElementById('statusBar').className = 'status-bar';
  }

  // ── Main conversion ──────────────────────────────────────────────────────────
  async function startConversion(mode) {
    let rawInput;
    if (mode === 'upload') {
      if (!uploadedFileContent) { showStatus('No file loaded - please upload an HTML file first.', 'error'); return; }
      rawInput = uploadedFileContent;
    } else {
      rawInput = mode === 'html'
        ? document.getElementById('htmlInput').value.trim()
        : document.getElementById('textInput').value.trim();
    }

    if (!rawInput) { showStatus('Nothing to convert - please paste some content first.', 'error'); return; }
    if (rawInput.length < 100) { showStatus('Content looks too short. Make sure you have the full page.', 'error'); return; }

    const apiKey = document.getElementById('apiKeyInput').value.trim();
    if (!apiKey) { showStatus('Please enter your Gemini API key above.', 'error'); return; }

    ['convertBtnHtml','convertBtnText','convertBtnUpload'].forEach(id => {
      const el = document.getElementById(id); if (el) el.disabled = true;
    });
    document.getElementById('outputSection').className = 'output-section';
    setStep(1);
    showStatus('Parsing HTML...');

    const today = new Date().toISOString().slice(0,10);
    let prompt;

    if (mode === 'html' || mode === 'upload') {
      // ── Step 1-2: extract in-browser ─────────────────────────────────────────
      setStep(2);
      showStatus('Extracting conversation turns...');

      let turns;
      try {
        turns = extractTurnsFromHtml(rawInput);
      } catch(e) {
        turns = [];
      }

      if (!turns.length) {
        // Fallback: couldn't find Copilot-specific classes — send raw HTML
        // with a strict verbatim instruction
        prompt =
          'You will receive HTML from a Microsoft Copilot conversation page.\n\n' +
          'CRITICAL RULES — you MUST follow these exactly:\n' +
          '1. Extract every word of visible text. DO NOT summarise, paraphrase, or omit anything.\n' +
          '2. Copy all text VERBATIM — every sentence, every word, unchanged.\n' +
          '3. Only apply Markdown formatting structure (headings, bold, code fences, lists).\n' +
          '4. Remove only pure UI chrome: button labels (Copy, Like, Regenerate, Share), nav links, cookie banners.\n\n' +
          'Output format:\n' +
          '---\nsource: Microsoft Copilot Shared Conversation\nexported: ' + today + '\n---\n\n' +
          'Use **User:** and **Copilot:** prefixes. Separate turns with ---.\n' +
          'Wrap code blocks in fenced ``` blocks with the language tag.\n' +
          'Output ONLY the markdown. No preamble.\n\n' +
          'HTML:\n\n' + rawInput.slice(0, 120000);
      } else {
        // ── Step 2: build pre-labelled plain text from extracted turns ──────────
        // Gemini only sees clean labelled text — it cannot summarise what it cannot see.
        const extracted = turns
          .map(t => '[' + t.role + ']\n' + t.text)
          .join('\n\n---\n\n');

        prompt =
          'Below is a Microsoft Copilot conversation that has already been extracted from HTML.\n' +
          'Each turn is labelled [USER] or [COPILOT] and the text is VERBATIM from the page.\n\n' +
          'Your ONLY job is to apply clean Markdown formatting. Rules:\n' +
          '1. COPY ALL TEXT EXACTLY AS GIVEN — do not change, summarise, or omit a single word.\n' +
          '2. Replace [USER] with **User:** and [COPILOT] with **Copilot:**\n' +
          '3. Separate each turn with --- (horizontal rule).\n' +
          '4. Detect and wrap code blocks in fenced ``` blocks with the correct language tag.\n' +
          '5. Apply heading levels (# ## ###) where the text already uses heading-like lines.\n' +
          '6. Apply bold/italic where the text uses ** or * markers.\n' +
          '7. Format bullet/numbered lists where they appear in the text.\n' +
          '8. Start with this front matter:\n' +
          '---\nsource: Microsoft Copilot Shared Conversation\nexported: ' + today + '\n---\n\n' +
          'Output ONLY the final markdown. No explanation, no preamble, nothing else.\n\n' +
          'CONVERSATION:\n\n' + extracted.slice(0, 120000);
      }
    } else {
      // Plain text mode — emphasise verbatim strongly
      prompt =
        'Below is plain text copied from a Microsoft Copilot conversation page.\n\n' +
        'Your ONLY job is to apply clean Markdown formatting. Rules:\n' +
        '1. COPY ALL TEXT EXACTLY AS GIVEN — do not change, summarise, or omit a single word.\n' +
        '2. Identify speaker turns and label them **User:** and **Copilot:**\n' +
        '3. Separate each turn with --- (horizontal rule).\n' +
        '4. Detect and wrap code in fenced ``` blocks with language tags.\n' +
        '5. Apply heading levels, bold, lists where visible in the text.\n' +
        '6. Remove only UI chrome: button text like Copy / Regenerate / Like, navigation.\n' +
        '7. Start with:\n' +
        '---\nsource: Microsoft Copilot Shared Conversation\nexported: ' + today + '\n---\n\n' +
        'Output ONLY the markdown. No explanation.\n\n' +
        'TEXT:\n\n' + rawInput.slice(0, 120000);
    }

    // ── Step 3: send to Gemini for formatting ─────────────────────────────────
    setStep(3);
    showStatus('Sending to Gemini for formatting...');

    const endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/' +
      GEMINI_MODEL + ':generateContent?key=' + apiKey;
    const body = JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 8192, temperature: 0.0 }
    });

    const MAX_RETRIES = 4;
    const BASE_MS = 8000;

    try {
      let data;
      for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body
        });

        if (res.status === 429) {
          if (attempt === MAX_RETRIES) throw new Error('Rate limit persists after retries. Wait ~60s then try again.');
          const retryAfter = parseInt(res.headers.get('Retry-After') || '0', 10);
          const waitMs = retryAfter > 0 ? retryAfter * 1000 : BASE_MS * Math.pow(2, attempt - 1);
          showStatus('Rate limit hit - retrying in ' + Math.round(waitMs/1000) + 's... (attempt ' + attempt + '/' + MAX_RETRIES + ')');
          await new Promise(r => setTimeout(r, waitMs));
          showStatus('Sending to Gemini for formatting...');
          continue;
        }

        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          const msg = (err.error && err.error.message) || ('API error ' + res.status);
          throw new Error(res.status === 403 ? 'Invalid API key - check your Gemini key and try again.' : msg);
        }

        data = await res.json();
        break;
      }

      const text = (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts)
        ? data.candidates[0].content.parts.map(p => p.text || '').join('\n').trim()
        : '';

      if (!text) throw new Error('Gemini returned an empty response. Please try again.');

      setStep(4);
      currentMd = text;
      currentFilename = 'copilot-chat-' + today + '.md';
      renderOutput(text);
      hideStatus();

    } catch (err) {
      showStatus('Failed: ' + err.message, 'error');
      console.error(err);
    } finally {
      ['convertBtnHtml','convertBtnText','convertBtnUpload'].forEach(id => {
        const el = document.getElementById(id); if (el) el.disabled = false;
      });
    }
  }

  // ── Output rendering ─────────────────────────────────────────────────────────
  function renderOutput(md) {
    document.getElementById('mdOut').value = md;
    updateCharLine();
    if (typeof marked !== 'undefined') {
      marked.setOptions({ breaks: true, gfm: true });
      document.getElementById('mdPreview').innerHTML = marked.parse(md);
    }
    document.getElementById('charBadge').textContent = md.length.toLocaleString() + ' chars';
    document.getElementById('outputSection').className = 'output-section visible';
    document.getElementById('outputSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function updateCharLine() {
    const v = document.getElementById('mdOut').value;
    document.getElementById('charLine').textContent =
      v.length.toLocaleString() + ' characters \u00b7 ' + v.split('\n').length.toLocaleString() + ' lines';
  }

  document.getElementById('mdOut').addEventListener('input', function() {
    currentMd = this.value;
    updateCharLine();
    if (document.getElementById('panePreview').classList.contains('active') && typeof marked !== 'undefined')
      document.getElementById('mdPreview').innerHTML = marked.parse(currentMd);
  });

  function switchOutTab(tab, btn) {
    document.querySelectorAll('.out-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.out-pane').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(tab === 'raw' ? 'paneRaw' : 'panePreview').classList.add('active');
    if (tab === 'preview' && typeof marked !== 'undefined') {
      marked.setOptions({ breaks: true, gfm: true });
      document.getElementById('mdPreview').innerHTML = marked.parse(currentMd);
    }
  }

  function downloadMd() {
    const blob = new Blob([document.getElementById('mdOut').value], { type: 'text/markdown;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = currentFilename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  async function copyMd() {
    try {
      await navigator.clipboard.writeText(document.getElementById('mdOut').value);
      const btn = event.target, o = btn.textContent;
      btn.textContent = 'Copied!';
      setTimeout(() => btn.textContent = o, 1500);
    } catch { document.getElementById('mdOut').select(); document.execCommand('copy'); }
  }

  function resetAll() {
    document.getElementById('htmlInput').value = '';
    document.getElementById('textInput').value = '';
    clearFile();
    document.getElementById('outputSection').className = 'output-section';
    hideStatus();
    currentMd = '';
  }
