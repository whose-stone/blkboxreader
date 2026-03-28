  let currentMd = '';
  let currentFilename = 'copilot-chat.md';
  let uploadedFileContent = null;

  // ── Drag & drop + file picker ────────────────────────────────────────────────
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFileSelect(e.dataTransfer.files);
  });
  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) handleFileSelect(fileInput.files);
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

  function getNodeText(node, listDepth) {
    // Walk the DOM and produce markdown that mirrors Copilot's visual structure.
    if (listDepth === undefined) listDepth = 0;
    let out = '';
    for (const child of node.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        out += child.textContent;
        continue;
      }
      if (child.nodeType !== Node.ELEMENT_NODE) continue;
      const tag = child.tagName.toLowerCase();
      if (tag === 'script' || tag === 'style' || tag === 'noscript') continue;

      // Code blocks
      if (tag === 'pre') {
        const codeEl = child.querySelector('code');
        const langClass = codeEl ? (codeEl.className || '') : '';
        const langMatch = langClass.match(/language-(\S+)/);
        const lang = langMatch ? langMatch[1] : '';
        out += '\n\n```' + lang + '\n' + (child.textContent || '') + '\n```\n\n';
        continue;
      }

      // Headings
      const hMatch = tag.match(/^h([1-6])$/);
      if (hMatch) {
        const level = '#'.repeat(parseInt(hMatch[1]));
        out += '\n\n' + level + ' ' + getNodeText(child, listDepth).trim() + '\n\n';
        continue;
      }

      // Paragraphs and divs that act as paragraphs
      if (tag === 'p') {
        out += '\n\n' + getNodeText(child, listDepth).trim() + '\n\n';
        continue;
      }

      // Lists
      if (tag === 'ul' || tag === 'ol') {
        let idx = 0;
        out += '\n';
        for (const li of child.children) {
          if (li.tagName.toLowerCase() !== 'li') continue;
          idx++;
          const indent = '  '.repeat(listDepth);
          const bullet = tag === 'ol' ? (idx + '. ') : '- ';
          const content = getNodeText(li, listDepth + 1).trim();
          out += indent + bullet + content + '\n';
        }
        out += '\n';
        continue;
      }
      if (tag === 'li') {
        out += getNodeText(child, listDepth);
        continue;
      }

      // Inline formatting
      if (tag === 'strong' || tag === 'b') {
        out += '**' + getNodeText(child, listDepth).trim() + '**';
        continue;
      }
      if (tag === 'em' || tag === 'i') {
        out += '*' + getNodeText(child, listDepth).trim() + '*';
        continue;
      }
      if (tag === 'code') {
        out += '`' + (child.textContent || '') + '`';
        continue;
      }
      if (tag === 'a') {
        const href = child.getAttribute('href') || '';
        const text = getNodeText(child, listDepth).trim();
        out += href ? '[' + text + '](' + href + ')' : text;
        continue;
      }

      // Line breaks
      if (tag === 'br') {
        out += '\n';
        continue;
      }

      // Horizontal rules
      if (tag === 'hr') {
        out += '\n\n---\n\n';
        continue;
      }

      // Blockquotes
      if (tag === 'blockquote') {
        const inner = getNodeText(child, listDepth).trim().split('\n').map(l => '> ' + l).join('\n');
        out += '\n\n' + inner + '\n\n';
        continue;
      }

      // Tables
      if (tag === 'table') {
        out += '\n\n' + tableToMarkdown(child) + '\n\n';
        continue;
      }

      // Everything else — recurse
      out += getNodeText(child, listDepth);
    }
    return out;
  }

  function tableToMarkdown(table) {
    const rows = [];
    for (const tr of table.querySelectorAll('tr')) {
      const cells = Array.from(tr.querySelectorAll('th, td')).map(c => c.textContent.trim());
      rows.push(cells);
    }
    if (!rows.length) return '';
    const colCount = Math.max(...rows.map(r => r.length));
    const lines = [];
    rows.forEach((cells, i) => {
      while (cells.length < colCount) cells.push('');
      lines.push('| ' + cells.join(' | ') + ' |');
      if (i === 0) {
        lines.push('| ' + cells.map(() => '---').join(' | ') + ' |');
      }
    });
    return lines.join('\n');
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
  function setStep(n) {
    for (let i = 1; i <= 4; i++) {
      const el = document.getElementById('step' + i);
      if (el) el.style.background = i <= n ? 'var(--accent)' : 'var(--border)';
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
      .map(t => {
        // Clean up excessive blank lines but preserve the markdown structure
        const cleaned = t.text.replace(/\n{3,}/g, '\n\n').trim();
        return `**${t.role === 'USER' ? 'User' : 'Copilot'}:**\n\n${cleaned}`;
      })
      .join('\n\n---\n\n');
    return `---\nsource: Microsoft Copilot Shared Conversation\nexported: ${today}\n---\n\n${body}\n`;
  }

  function buildMarkdownFromPlainText(text, today) {
    return `---\nsource: Microsoft Copilot Shared Conversation\nexported: ${today}\n---\n\n${prettifyBlobText(text)}\n`;
  }

  function prettifyBlobText(input) {
    const source = String(input || '').trim();
    if (!source) return '';
    if (source.includes('```')) return source;

    let text = source.replace(/\s+/g, ' ').trim();

    const sectionPatterns = [
      /Bold summary:/gi,
      /Quick decision table[^:]*:/gi,
      /Recommended hardware[^:]*:/gi,
      /Software stack[^:]*:/gi,
      /Network & router setup[^:]*:/gi,
      /Admin features[^:]*:/gi,
      /Risks, tradeoffs & next steps:?/gi,
      /Next steps[^:]*:/gi
    ];

    for (const re of sectionPatterns) {
      text = text.replace(re, (m) => `\n\n## ${m.replace(/:$/, '').trim()}\n\n`);
    }

    text = text.replace(/\n{3,}/g, '\n\n').trim();

    const blocks = text.split(/\n\n+/).map(b => b.trim()).filter(Boolean);
    const out = [];

    for (const block of blocks) {
      if (block.startsWith('## ')) {
        out.push(block);
        continue;
      }

      const sentences = block.split(/(?<=[.!?])\s+(?=[A-Z0-9])/).filter(Boolean);
      if (sentences.length <= 2) {
        out.push(sentences.join(' '));
      } else {
        for (let i = 0; i < sentences.length; i += 2) {
          out.push(sentences.slice(i, i + 2).join(' '));
        }
      }
    }

    return out.join('\n\n');
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

    ['convertBtnUpload'].forEach(id => {
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
      ['convertBtnUpload'].forEach(id => {
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
    const urlEl = document.getElementById('urlInput');
    const htmlEl = document.getElementById('htmlInput');
    const textEl = document.getElementById('textInput');
    if (urlEl) urlEl.value = '';
    if (htmlEl) htmlEl.value = '';
    if (textEl) textEl.value = '';
    clearFile();
    document.getElementById('outputSection').className = 'output-section';
    hideStatus();
    currentMd = '';
  }

  function getBookmarkletPayload() {
    return "(function(){try{var U='[class*=\\\"group/user-message\\\"],[data-content=\\\"user-message\\\"]';var A='[class*=\\\"group/ai-message-item\\\"],[data-content=\\\"ai-message-item\\\"]';var tbl=function(t){var rows=[];t.querySelectorAll('tr').forEach(function(tr){var cells=[];tr.querySelectorAll('th,td').forEach(function(c){cells.push(c.textContent.trim())});rows.push(cells)});if(!rows.length)return'';var cc=Math.max.apply(null,rows.map(function(r){return r.length}));var lines=[];rows.forEach(function(cells,i){while(cells.length<cc)cells.push('');lines.push('| '+cells.join(' | ')+' |');if(i===0)lines.push('| '+cells.map(function(){return'---'}).join(' | ')+' |')});return lines.join('\\n')};var nodeText=function(n,d){if(d===undefined)d=0;var o='';for(var i=0;i<n.childNodes.length;i++){var c=n.childNodes[i];if(c.nodeType===Node.TEXT_NODE){o+=c.textContent||'';continue}if(c.nodeType!==Node.ELEMENT_NODE)continue;var t=c.tagName.toLowerCase();if(t==='script'||t==='style'||t==='noscript')continue;if(t==='pre'){var ce=c.querySelector('code');var lc=ce?(ce.className||''):'';var lm=lc.match(/language-(\\S+)/);var lang=lm?lm[1]:'';o+='\\n\\n```'+lang+'\\n'+(c.textContent||'')+'\\n```\\n\\n';continue}var hm=t.match(/^h([1-6])$/);if(hm){var lv='';for(var x=0;x<parseInt(hm[1]);x++)lv+='#';o+='\\n\\n'+lv+' '+nodeText(c,d).trim()+'\\n\\n';continue}if(t==='p'){o+='\\n\\n'+nodeText(c,d).trim()+'\\n\\n';continue}if(t==='ul'||t==='ol'){var idx=0;o+='\\n';for(var j=0;j<c.children.length;j++){var li=c.children[j];if(li.tagName.toLowerCase()!=='li')continue;idx++;var ind='';for(var k=0;k<d;k++)ind+='  ';var bul=t==='ol'?(idx+'. '):'- ';o+=ind+bul+nodeText(li,d+1).trim()+'\\n'}o+='\\n';continue}if(t==='li'){o+=nodeText(c,d);continue}if(t==='strong'||t==='b'){o+='**'+nodeText(c,d).trim()+'**';continue}if(t==='em'||t==='i'){o+='*'+nodeText(c,d).trim()+'*';continue}if(t==='code'){o+='`'+(c.textContent||'')+'`';continue}if(t==='a'){var hr=c.getAttribute('href')||'';var tx=nodeText(c,d).trim();o+=hr?'['+tx+']('+hr+')':tx;continue}if(t==='br'){o+='\\n';continue}if(t==='hr'){o+='\\n\\n---\\n\\n';continue}if(t==='blockquote'){o+='\\n\\n> '+nodeText(c,d).trim().split('\\n').join('\\n> ')+'\\n\\n';continue}if(t==='table'){o+='\\n\\n'+tbl(c)+'\\n\\n';continue}o+=nodeText(c,d)}return o};var users=Array.from(document.querySelectorAll(U)).map(function(el){return{el:el,role:'User'}});var ais=Array.from(document.querySelectorAll(A)).map(function(el){return{el:el,role:'Copilot'}});var all=users.concat(ais).sort(function(a,b){var r=a.el.compareDocumentPosition(b.el);if(r&Node.DOCUMENT_POSITION_FOLLOWING)return-1;if(r&Node.DOCUMENT_POSITION_PRECEDING)return 1;return 0});if(!all.length){alert('No Copilot turns found on this page. Open a shared conversation first.');return}var today=new Date().toISOString().slice(0,10);var turns=all.map(function(item){return{role:item.role,text:nodeText(item.el).replace(/\\n{3,}/g,'\\n\\n').trim()}}).filter(function(t){return t.text});if(!turns.length){alert('Found message containers, but no text content.');return}var body=turns.map(function(t){return'**'+t.role+':**\\n\\n'+t.text}).join('\\n\\n---\\n\\n');var md='---\\nsource: Microsoft Copilot Shared Conversation\\nexported: '+today+'\\n---\\n\\n'+body+'\\n';var slug=(location.pathname.split('/').pop()||'copilot-chat').replace(/[^a-zA-Z0-9_-]/g,'')||'copilot-chat';var a=document.createElement('a');a.href=URL.createObjectURL(new Blob([md],{type:'text/markdown;charset=utf-8'}));a.download=slug+'-'+today+'.md';document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(a.href)},1000)}catch(e){alert('Copilot-to-Markdown bookmarklet failed: '+(e&&e.message?e.message:e))}})();";
  }

  function initBookmarkletUi() {
    const installBtn = document.getElementById('installBookmarkletBtn');
    if (!installBtn) return;

    const payload = getBookmarkletPayload();
    installBtn.href = 'javascript:' + payload;
    installBtn.removeAttribute('onclick');
  }

  initBookmarkletUi();
