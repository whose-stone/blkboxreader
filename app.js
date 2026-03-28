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
  // then format the extracted turns directly into markdown locally.

  function getNodeText(node) {
    // Walk the node tree and build a plain-text representation.
    // <pre> blocks get fenced with ``` so code is preserved in markdown.
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
    const userEls  = Array.from(doc.querySelectorAll('[class*="group/user-message"], [data-content="user-message"]'));
    const aiEls    = Array.from(doc.querySelectorAll('[class*="group/ai-message-item"], [data-content="ai-message-item"]'));

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

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function ensureHtmlDocument(content, sourceUrl) {
    const trimmed = (content || '').trim();
    if (/<!doctype html/i.test(trimmed) || /<html[\s>]/i.test(trimmed)) {
      return content;
    }
    return '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8"/>\n<title>Copilot Share Snapshot</title>\n</head>\n<body>\n<main>\n<h1>Copilot Share Snapshot</h1>\n<p><strong>Source URL:</strong> <a href="' + escapeHtml(sourceUrl) + '">' + escapeHtml(sourceUrl) + '</a></p>\n<pre>' + escapeHtml(content) + '</pre>\n</main>\n</body>\n</html>\n';
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

  async function fetchShareHtmlFromUrl(url) {
    const cleaned = (url || '').trim();
    if (!/^https:\/\/copilot\.microsoft\.com\/shares\//.test(cleaned)) {
      throw new Error('Please enter a valid Copilot share URL.');
    }

    // Try direct fetch first (works in some environments)
    try {
      const direct = await fetch(cleaned, { method: 'GET' });
      if (direct.ok) {
        const html = await direct.text();
        if (html && html.length > 500) return { html, source: 'direct' };
      }
    } catch (_) {}

    // Fallback to a read proxy to avoid CORS blocks in static hosting.
    const proxyUrl = 'https://r.jina.ai/http://' + cleaned.replace(/^https?:\/\//, '');
    const proxyRes = await fetch(proxyUrl, { method: 'GET' });
    if (!proxyRes.ok) {
      throw new Error('Could not fetch this URL from the browser. Try Upload HTML File mode with the CLI helper.');
    }
    const proxyText = await proxyRes.text();
    if (!proxyText || proxyText.length < 200) {
      throw new Error('Fetched content was empty. Try Upload HTML File mode with the CLI helper.');
    }
    return { html: proxyText, source: 'proxy' };
  }

  function buildShareFilename(url) {
    const tail = (url.split('/').pop() || 'copilot-share').split('?')[0];
    const slug = tail.replace(/[^a-zA-Z0-9_-]/g, '') || 'copilot-share';
    return slug + '.html';
  }

  function buildMarkdownFromTurns(turns, today) {
    const body = turns
      .map(t => `**${t.role === 'USER' ? 'User' : 'Copilot'}:**\n\n${t.text}`)
      .join('\n\n---\n\n');
    return `---\nsource: Microsoft Copilot Shared Conversation\nexported: ${today}\n---\n\n${body}\n`;
  }

  function buildMarkdownFromPlainText(text, today) {
    return `---\nsource: Microsoft Copilot Shared Conversation\nexported: ${today}\n---\n\n${text.trim()}\n`;
  }

  async function downloadHtmlFromUrl() {
    const btn = document.getElementById('downloadHtmlBtn');
    const url = document.getElementById('urlInput').value.trim();
    if (!url) { showStatus('Please paste a Copilot share URL first.', 'error'); return; }

    const original = btn ? btn.textContent : '';
    if (btn) {
      btn.disabled = true;
      btn.textContent = 'Fetching...';
    }
    showStatus('Fetching URL and preparing HTML download...');

    try {
      const fetched = await fetchShareHtmlFromUrl(url);
      const html = ensureHtmlDocument(fetched.html, url);
      const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = buildShareFilename(url);
      a.click();
      URL.revokeObjectURL(a.href);
      const note = fetched.source === 'proxy'
        ? 'Downloaded HTML snapshot from proxy content (may not include full Copilot DOM).'
        : 'Downloaded HTML snapshot from the shared URL.';
      showStatus(note + ' You can now upload it in Upload HTML File mode.');
    } catch (err) {
      showStatus('Failed to fetch URL: ' + err.message, 'error');
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.textContent = original;
      }
    }
  }

  // ── Main conversion ──────────────────────────────────────────────────────────
  async function startConversion(mode) {
    let rawInput;
    if (mode === 'upload') {
      if (!uploadedFileContent) { showStatus('No file loaded - please upload an HTML file first.', 'error'); return; }
      rawInput = uploadedFileContent;
    } else if (mode === 'url') {
      const url = document.getElementById('urlInput').value.trim();
      if (!url) { showStatus('Please paste a Copilot share URL first.', 'error'); return; }
      showStatus('Fetching share URL...');
      setStep(1);
      try {
        rawInput = (await fetchShareHtmlFromUrl(url)).html;
      } catch (err) {
        showStatus(err.message, 'error');
        return;
      }
    } else {
      rawInput = mode === 'html'
        ? document.getElementById('htmlInput').value.trim()
        : document.getElementById('textInput').value.trim();
    }

    if (!rawInput) { showStatus('Nothing to convert - please paste some content first.', 'error'); return; }
    if (rawInput.length < 100) { showStatus('Content looks too short. Make sure you have the full page.', 'error'); return; }

    ['convertBtnUrl','convertBtnHtml','convertBtnText','convertBtnUpload','downloadHtmlBtn'].forEach(id => {
      const el = document.getElementById(id); if (el) el.disabled = true;
    });
    document.getElementById('outputSection').className = 'output-section';
    setStep(1);
    showStatus('Parsing HTML...');

    const today = new Date().toISOString().slice(0,10);
    let extractedTurns = [];
    if (mode === 'html' || mode === 'upload' || mode === 'url') {
      // ── Step 1-2: extract in-browser ─────────────────────────────────────────
      setStep(2);
      showStatus('Extracting conversation turns...');

      try {
        extractedTurns = extractTurnsFromHtml(rawInput);
      } catch(e) {
        extractedTurns = [];
      }

      if (!extractedTurns.length) {
        if (mode === 'url') {
          // URL fetches can return reader/plain text from proxy fallbacks.
          // In that case, continue with plain-text local formatting.
          showStatus('Turn classes not found; using plain-text URL fallback...');
        } else {
          showStatus('Could not find Copilot conversation turns in this HTML. Please use a saved Copilot share page.', 'error');
          return;
        }
      }
    }

    // ── Step 3: build markdown locally (no LLM dependency) ───────────────────
    setStep(3);
    showStatus('Formatting markdown locally...');

    try {
      const text = extractedTurns.length
        ? buildMarkdownFromTurns(extractedTurns, today)
        : buildMarkdownFromPlainText(rawInput, today);

      setStep(4);
      currentMd = text;
      currentFilename = 'copilot-chat-' + today + '.md';
      renderOutput(text);
      showStatus('Done. Converted locally (no LLM used).');

    } catch (err) {
      showStatus('Failed: ' + err.message, 'error');
      console.error(err);
    } finally {
      ['convertBtnUrl','convertBtnHtml','convertBtnText','convertBtnUpload','downloadHtmlBtn'].forEach(id => {
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

  async function copyMd(evt) {
    try {
      await navigator.clipboard.writeText(document.getElementById('mdOut').value);
      const btn = evt && evt.currentTarget;
      if (!btn) return;
      const o = btn.textContent;
      btn.textContent = 'Copied!';
      setTimeout(() => btn.textContent = o, 1500);
    } catch { document.getElementById('mdOut').select(); document.execCommand('copy'); }
  }

  function resetAll() {
    document.getElementById('urlInput').value = '';
    document.getElementById('htmlInput').value = '';
    document.getElementById('textInput').value = '';
    clearFile();
    document.getElementById('outputSection').className = 'output-section';
    hideStatus();
    currentMd = '';
  }

  function getBookmarkletPayload() {
    return "(function(){try{const U='[class*=\\\"group/user-message\\\"],[data-content=\\\"user-message\\\"]';const A='[class*=\\\"group/ai-message-item\\\"],[data-content=\\\"ai-message-item\\\"]';const nodeText=(n)=>{let o='';for(const c of n.childNodes){if(c.nodeType===Node.TEXT_NODE){o+=c.textContent||'';continue;}if(c.nodeType!==Node.ELEMENT_NODE)continue;const t=c.tagName.toLowerCase();if(t==='script'||t==='style'||t==='noscript')continue;if(t==='pre'){o+='\\n```\\n'+(c.textContent||'')+'\\n```\\n';continue;}if(t==='br'){o+='\\n';continue;}o+=nodeText(c);}return o;};const users=Array.from(document.querySelectorAll(U)).map(el=>({el,role:'User'}));const ais=Array.from(document.querySelectorAll(A)).map(el=>({el,role:'Copilot'}));const all=users.concat(ais).sort((a,b)=>{const r=a.el.compareDocumentPosition(b.el);if(r&Node.DOCUMENT_POSITION_FOLLOWING)return-1;if(r&Node.DOCUMENT_POSITION_PRECEDING)return 1;return 0;});if(!all.length){alert('No Copilot turns found on this page. Open a shared conversation first.');return;}const today=new Date().toISOString().slice(0,10);const turns=all.map(({role,el})=>({role,text:nodeText(el).replace(/\\n{3,}/g,'\\n\\n').trim()})).filter(t=>t.text);if(!turns.length){alert('Found message containers, but no text content.');return;}const body=turns.map(t=>'**'+t.role+':**\\n\\n'+t.text).join('\\n\\n---\\n\\n');const md='---\\nsource: Microsoft Copilot Shared Conversation\\nexported: '+today+'\\n---\\n\\n'+body+'\\n';const slug=(location.pathname.split('/').pop()||'copilot-chat').replace(/[^a-zA-Z0-9_-]/g,'')||'copilot-chat';const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([md],{type:'text/markdown;charset=utf-8'}));a.download=slug+'-'+today+'.md';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000);}catch(e){alert('Copilot-to-Markdown bookmarklet failed: '+(e&&e.message?e.message:e));}})();";
  }

  function initBookmarkletUi() {
    const link = document.getElementById('bookmarkletLink');
    const copyBtn = document.getElementById('copyBookmarkletBtn');
    if (!link) return;

    const payload = getBookmarkletPayload();
    link.setAttribute('href', 'javascript:' + payload);

    if (copyBtn) {
      copyBtn.addEventListener('click', async () => {
        const txt = 'javascript:' + payload;
        try {
          await navigator.clipboard.writeText(txt);
          const old = copyBtn.textContent;
          copyBtn.textContent = 'Copied!';
          setTimeout(() => (copyBtn.textContent = old), 1500);
        } catch {
          showStatus('Could not copy bookmarklet automatically. Drag the bookmark link manually.', 'error');
        }
      });
    }
  }

  initBookmarkletUi();
