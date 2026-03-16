/**
 * M Exporter — UI controller.
 * Wires DOM elements to the client-side extraction engine (extractor.js).
 */
(function() {
  const $ = id => document.getElementById(id);

  const dropzone       = $('dropzone');
  const fileInput      = $('fileInput');
  const fileChip       = $('fileChip');
  const fileName       = $('fileName');
  const fileSize       = $('fileSize');
  const fileRemove     = $('fileRemove');
  const modeSelect     = $('modeSelect');
  const metadataToggle = $('metadataToggle');
  const runBtn         = $('runBtn');
  const errorBox       = $('errorBox');
  const outputTabs     = $('outputTabs');
  const outputActions  = $('outputActions');
  const copyBtn        = $('copyBtn');
  const downloadBtn    = $('downloadBtn');
  const emptyState     = $('emptyState');
  const renderedOutput = $('renderedOutput');
  const rawOutput      = $('rawOutput');
  const statsBar       = $('statsBar');
  const graphContainer = $('graphContainer');
  const graphSvg       = $('graphSvg');
  const graphTooltip   = $('graphTooltip');

  let selectedFile = null;
  let lastMarkdown = '';
  let lastFileName = '';
  let lastGraphData = null;

  // ---- File handling ----
  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
  }

  function setFile(file) {
    if (!file) return;
    if (!file.name.match(/\.(pbix|pbit)$/i)) {
      showError('Please select a .pbix or .pbit file');
      return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    fileChip.style.display = 'flex';
    runBtn.disabled = false;
    hideError();
  }

  function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    fileChip.style.display = 'none';
    runBtn.disabled = true;
  }

  fileInput.addEventListener('change', e => setFile(e.target.files[0]));
  fileRemove.addEventListener('click', clearFile);
  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag-over'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
  });

  // ---- Tabs ----
  document.querySelectorAll('.output-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.output-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.output-view').forEach(v => v.classList.remove('active'));
      tab.classList.add('active');
      $('view-' + tab.dataset.tab).classList.add('active');
    });
  });

  // ---- Error ----
  function showError(msg) { errorBox.textContent = msg; errorBox.style.display = 'block'; }
  function hideError() { errorBox.style.display = 'none'; }

  // ---- Markdown renderer ----
  function highlightM(code) {
    code = code.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    code = code.replace(/(\/\/.*$)/gm, '<span class="cmt">$1</span>');
    code = code.replace(/("(?:[^"\\]|\\.)*")/g, '<span class="str">$1</span>');
    const kws = ['let','in','if','then','else','each','true','false','null','not','and','or','meta','type','shared','section','try','otherwise','error','as','is'];
    code = code.replace(new RegExp('\\b(' + kws.join('|') + ')\\b', 'g'), '<span class="kw">$1</span>');
    code = code.replace(/\b([A-Z][a-zA-Z]+\.[A-Z][a-zA-Z]+)\b/g, '<span class="fn">$1</span>');
    code = code.replace(/\b(\d+(?:\.\d+)?)\b/g, '<span class="num">$1</span>');
    return code;
  }

  function renderMd(md) {
    const lines = md.split('\n');
    let html = '', inCode = false, codeBlock = '', inList = false;
    for (const line of lines) {
      if (line.startsWith('```')) {
        if (inCode) { html += '<pre><code>' + highlightM(codeBlock) + '</code></pre>'; codeBlock = ''; inCode = false; }
        else { if (inList) { html += '</ul>'; inList = false; } inCode = true; }
        continue;
      }
      if (inCode) { codeBlock += line + '\n'; continue; }
      if (line.startsWith('### ')) { if (inList) { html += '</ul>'; inList = false; } html += '<h3>' + line.slice(4) + '</h3>'; }
      else if (line.startsWith('## ')) { if (inList) { html += '</ul>'; inList = false; } html += '<h2>' + line.slice(3) + '</h2>'; }
      else if (line.startsWith('# ')) { if (inList) { html += '</ul>'; inList = false; } html += '<h1>' + line.slice(2) + '</h1>'; }
      else if (line.startsWith('> ')) { if (inList) { html += '</ul>'; inList = false; } html += '<blockquote>' + line.slice(2).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>') + '</blockquote>'; }
      else if (line.match(/^\s*- /)) {
        if (!inList) { html += '<ul>'; inList = true; }
        html += '<li>' + line.replace(/^\s*- /, '').replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>').replace(/`([^`]+)`/g, '<code>$1</code>') + '</li>';
      }
      else if (line.trim() === '') { if (inList) { html += '</ul>'; inList = false; } }
      else { html += '<p>' + line.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>') + '</p>'; }
    }
    if (inList) html += '</ul>';
    return html;
  }

  // =====================================================================
  // GRAPH RENDERER (reusable for inline + fullscreen)
  // =====================================================================
  const NODE_W = 160, NODE_H = 36, LAYER_GAP = 100, NODE_GAP = 30, PAD_X = 60, PAD_Y = 50;
  const KIND_COLORS = { Parameter: '#4a9eff', Query: '#f0a500', Function: '#d4a0ff' };

  function buildGraphInto(graphData, targetSvg, targetContainer, targetTooltip, zoomInBtn, zoomOutBtn, zoomFitBtn) {
    const { nodes, edges, layers, dataSources: ds } = graphData;
    // Normalize dataSources key (JS engine uses camelCase)
    const dataSources = ds || graphData.data_sources || {};
    if (!nodes.length) return null;

    const nodeMap = {};
    nodes.forEach(n => { nodeMap[n.name] = n; });
    const depsOf = {}, usedBy = {};
    edges.forEach(e => {
      (depsOf[e.from] = depsOf[e.from] || []).push(e.to);
      (usedBy[e.to] = usedBy[e.to] || []).push(e.from);
    });

    const positions = {};
    let maxLW = 0;
    layers.forEach(l => { maxLW = Math.max(maxLW, l.length); });
    const svgW = Math.max(maxLW * (NODE_W + NODE_GAP) + PAD_X * 2, 600);
    const svgH = layers.length * (NODE_H + LAYER_GAP) + PAD_Y * 2;

    layers.forEach((layer, li) => {
      const totalW = layer.length * NODE_W + (layer.length - 1) * NODE_GAP;
      const startX = (svgW - totalW) / 2;
      const y = PAD_Y + li * (NODE_H + LAYER_GAP);
      layer.forEach((name, ni) => { positions[name] = { x: startX + ni * (NODE_W + NODE_GAP), y }; });
    });

    const ns = 'http://www.w3.org/2000/svg';
    targetSvg.innerHTML = '';
    targetSvg.setAttribute('width', svgW);
    targetSvg.setAttribute('height', svgH);
    targetSvg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);

    const markerId = 'ah-' + Math.random().toString(36).slice(2, 6);
    const defs = document.createElementNS(ns, 'defs');
    const marker = document.createElementNS(ns, 'marker');
    marker.setAttribute('id', markerId);
    marker.setAttribute('markerWidth', '8'); marker.setAttribute('markerHeight', '6');
    marker.setAttribute('refX', '8'); marker.setAttribute('refY', '3'); marker.setAttribute('orient', 'auto');
    const ap = document.createElementNS(ns, 'path');
    ap.setAttribute('d', 'M0,0 L0,6 L8,3 z'); ap.setAttribute('fill', '#3e3e4a');
    marker.appendChild(ap); defs.appendChild(marker); targetSvg.appendChild(defs);

    const rootG = document.createElementNS(ns, 'g');
    targetSvg.appendChild(rootG);
    const edgeGroup = document.createElementNS(ns, 'g');
    rootG.appendChild(edgeGroup);

    const edgeEls = [];
    edges.forEach(e => {
      const fp = positions[e.from], tp = positions[e.to];
      if (!fp || !tp) return;
      const x1 = tp.x + NODE_W/2, y1 = tp.y + NODE_H, x2 = fp.x + NODE_W/2, y2 = fp.y;
      const cp = Math.max(Math.abs(y2 - y1) * 0.4, 30);
      const path = document.createElementNS(ns, 'path');
      path.setAttribute('d', `M${x1},${y1} C${x1},${y1+cp} ${x2},${y2-cp} ${x2},${y2}`);
      path.setAttribute('class', 'graph-edge'); path.setAttribute('stroke', '#3e3e4a');
      path.setAttribute('marker-end', `url(#${markerId})`);
      path.dataset.from = e.from; path.dataset.to = e.to;
      edgeGroup.appendChild(path); edgeEls.push(path);
    });

    const nodeGroup = document.createElementNS(ns, 'g');
    rootG.appendChild(nodeGroup);
    const nodeEls = {};

    nodes.forEach(n => {
      const pos = positions[n.name];
      if (!pos) return;
      const g = document.createElementNS(ns, 'g');
      g.setAttribute('class', 'graph-node');
      g.setAttribute('transform', `translate(${pos.x},${pos.y})`);
      g.dataset.name = n.name;

      const rect = document.createElementNS(ns, 'rect');
      rect.setAttribute('width', NODE_W); rect.setAttribute('height', NODE_H); rect.setAttribute('rx', '4');
      rect.setAttribute('fill', '#1c1c22'); rect.setAttribute('stroke', KIND_COLORS[n.kind] || '#3e3e4a');
      rect.setAttribute('stroke-width', '1.5');
      g.appendChild(rect);

      if (dataSources[n.name]) {
        const badge = document.createElementNS(ns, 'circle');
        badge.setAttribute('cx', NODE_W - 10); badge.setAttribute('cy', 10);
        badge.setAttribute('r', '4'); badge.setAttribute('fill', '#3cb56c');
        g.appendChild(badge);
      }

      const text = document.createElementNS(ns, 'text');
      text.setAttribute('x', NODE_W/2); text.setAttribute('y', NODE_H/2 + 1);
      text.setAttribute('text-anchor', 'middle'); text.setAttribute('dominant-baseline', 'middle');
      text.setAttribute('fill', '#e8e6e3'); text.setAttribute('font-size', '11');
      let label = n.name;
      if (label.length > 18) label = label.slice(0, 16) + '..';
      text.textContent = label;
      g.appendChild(text);

      nodeGroup.appendChild(g);
      nodeEls[n.name] = g;
    });

    // Hover
    function highlight(name) {
      targetContainer.classList.add('has-hover');
      const connected = new Set([name]);
      edges.forEach(e => { if (e.from === name || e.to === name) { connected.add(e.from); connected.add(e.to); } });
      Object.entries(nodeEls).forEach(([n, el]) => { el.classList.toggle('dimmed', !connected.has(n)); el.classList.toggle('highlighted', n === name); });
      edgeEls.forEach(el => {
        const c = el.dataset.from === name || el.dataset.to === name;
        el.classList.toggle('dimmed', !c);
        if (c) el.setAttribute('stroke', KIND_COLORS[nodeMap[name]?.kind] || '#f0a500');
      });
    }
    function clearHL() {
      targetContainer.classList.remove('has-hover');
      Object.values(nodeEls).forEach(el => el.classList.remove('dimmed', 'highlighted'));
      edgeEls.forEach(el => { el.classList.remove('dimmed'); el.setAttribute('stroke', '#3e3e4a'); });
      targetTooltip.classList.remove('visible');
    }

    Object.entries(nodeEls).forEach(([name, el]) => {
      el.addEventListener('mouseenter', () => {
        highlight(name);
        const nd = nodeMap[name];
        let h = `<strong>${name}</strong><span class="tt-kind">${nd.kind}</span>`;
        if (depsOf[name]?.length) h += `<br><span class="tt-dep">depends on: ${depsOf[name].join(', ')}</span>`;
        if (usedBy[name]?.length) h += `<br><span class="tt-dep">used by: ${usedBy[name].join(', ')}</span>`;
        const src = dataSources[name];
        if (src?.length) h += `<br><span class="tt-src">source: ${src.join(', ')}</span>`;
        targetTooltip.innerHTML = h;
        targetTooltip.classList.add('visible');
      });
      el.addEventListener('mouseleave', clearHL);
    });

    targetContainer.addEventListener('mousemove', ev => {
      const r = targetContainer.getBoundingClientRect();
      let x = ev.clientX - r.left + 15, y = ev.clientY - r.top + 15;
      if (x + 250 > r.width) x = ev.clientX - r.left - 260;
      if (y + 100 > r.height) y = ev.clientY - r.top - 80;
      targetTooltip.style.left = x + 'px'; targetTooltip.style.top = y + 'px';
    });

    // Pan / Zoom
    const st = { scale: 1, tx: 0, ty: 0, dragging: false, lx: 0, ly: 0 };
    function apply() { rootG.setAttribute('transform', `translate(${st.tx},${st.ty}) scale(${st.scale})`); }

    function fitToView() {
      const cr = targetContainer.getBoundingClientRect();
      const cw = cr.width || 800, ch = cr.height || 450;
      st.scale = Math.min(cw / svgW, ch / svgH, 1) * 0.92;
      st.tx = (cw - svgW * st.scale) / 2;
      st.ty = (ch - svgH * st.scale) / 2;
      apply();
    }

    if (zoomInBtn) zoomInBtn.onclick = () => { st.scale = Math.min(st.scale * 1.25, 3); apply(); };
    if (zoomOutBtn) zoomOutBtn.onclick = () => { st.scale = Math.max(st.scale * 0.8, 0.15); apply(); };
    if (zoomFitBtn) zoomFitBtn.onclick = fitToView;

    targetContainer.addEventListener('wheel', ev => {
      ev.preventDefault();
      const d = ev.deltaY > 0 ? 0.9 : 1.1;
      const ns2 = Math.max(0.15, Math.min(3, st.scale * d));
      const r = targetContainer.getBoundingClientRect();
      const mx = ev.clientX - r.left, my = ev.clientY - r.top;
      st.tx = mx - (mx - st.tx) * (ns2 / st.scale);
      st.ty = my - (my - st.ty) * (ns2 / st.scale);
      st.scale = ns2; apply();
    }, { passive: false });

    targetContainer.addEventListener('mousedown', ev => {
      if (ev.target.closest('.graph-node')) return;
      st.dragging = true; st.lx = ev.clientX; st.ly = ev.clientY;
    });
    window.addEventListener('mousemove', ev => {
      if (!st.dragging) return;
      st.tx += ev.clientX - st.lx; st.ty += ev.clientY - st.ly;
      st.lx = ev.clientX; st.ly = ev.clientY; apply();
    });
    window.addEventListener('mouseup', () => { st.dragging = false; });

    return { fitToView };
  }

  function renderGraph(graphData) {
    lastGraphData = graphData;
    const ctrl = buildGraphInto(graphData, graphSvg, graphContainer, graphTooltip,
      $('zoomIn'), $('zoomOut'), $('zoomFit'));
    if (!ctrl) return;
    graphContainer.style.display = 'block';
    emptyState.style.display = 'none';
    requestAnimationFrame(ctrl.fitToView);
  }

  // ---- Fullscreen pop-out ----
  const fsOverlay = $('fsOverlay'), fsBody = $('fsBody'), fsSvg = $('fsSvg'), fsTooltip = $('fsTooltip');

  $('popoutBtn').addEventListener('click', () => {
    if (!lastGraphData) return;
    fsOverlay.classList.add('active');
    document.body.style.overflow = 'hidden';
    requestAnimationFrame(() => {
      const ctrl = buildGraphInto(lastGraphData, fsSvg, fsBody, fsTooltip, $('fsZoomIn'), $('fsZoomOut'), $('fsZoomFit'));
      if (ctrl) requestAnimationFrame(ctrl.fitToView);
    });
  });

  function closeFS() { fsOverlay.classList.remove('active'); document.body.style.overflow = ''; fsSvg.innerHTML = ''; }
  $('fsClose').addEventListener('click', closeFS);
  document.addEventListener('keydown', ev => { if (ev.key === 'Escape' && fsOverlay.classList.contains('active')) closeFS(); });

  // =====================================================================
  // EXTRACT — calls client-side engine
  // =====================================================================
  runBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    hideError();
    runBtn.classList.add('processing');
    runBtn.textContent = 'EXTRACTING...';
    runBtn.disabled = true;

    try {
      const arrayBuffer = await selectedFile.arrayBuffer();
      const data = await window.MExporter.extractFile(
        arrayBuffer,
        selectedFile.name,
        modeSelect.value,
        metadataToggle.checked,
      );

      lastMarkdown = data.markdown;
      lastFileName = selectedFile.name.replace(/\.(pbix|pbit)$/i, '') + '.md';

      renderedOutput.style.display = 'block';
      renderedOutput.innerHTML = renderMd(data.markdown);
      rawOutput.value = data.markdown;
      outputTabs.style.display = 'flex';
      outputActions.style.display = 'flex';
      statsBar.style.display = 'flex';

      $('statParams').textContent = data.stats.parameters;
      $('statQueries').textContent = data.stats.queries;
      $('statFuncs').textContent = data.stats.functions;
      $('statDeps').textContent = data.graph.edges.length;

      if (data.graph && data.graph.nodes.length > 0) {
        renderGraph(data.graph);
      }

    } catch (err) {
      showError(err.message || 'Extraction failed');
    } finally {
      runBtn.classList.remove('processing');
      runBtn.textContent = 'EXTRACT';
      runBtn.disabled = !selectedFile;
    }
  });

  // ---- Copy / Download ----
  copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(lastMarkdown).then(() => {
      copyBtn.textContent = 'Copied!';
      setTimeout(() => { copyBtn.textContent = 'Copy'; }, 1500);
    });
  });

  downloadBtn.addEventListener('click', () => {
    const blob = new Blob([lastMarkdown], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = lastFileName; a.click();
    URL.revokeObjectURL(url);
  });

})();
