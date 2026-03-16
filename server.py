#!/usr/bin/env python3
"""
Web frontend for pbix_m_export — stdlib-only HTTP server.
Serves the UI and handles file upload + extraction.
"""

from __future__ import annotations

import cgi
import io
import json
import os
import sys
import tempfile
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from dataclasses import asdict

# Import the extraction engine
from pbix_m_export import (
    MItem,
    MItemKind,
    analyze_dependencies,
    dependency_graph_to_dict,
    detect_mode,
    extract_from_datamashup,
    extract_from_pbit_json,
    render_markdown,
)

PORT = int(os.environ.get("PORT", 8070))
HOST = "0.0.0.0"

# ---------------------------------------------------------------------------
# HTML — embedded as a constant
# ---------------------------------------------------------------------------

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>M Exporter</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500;600;700&family=DM+Sans:wght@400;500;600;700&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg-0: #0c0c0e;
    --bg-1: #141418;
    --bg-2: #1c1c22;
    --bg-3: #26262e;
    --border: #2e2e38;
    --border-hover: #3e3e4a;
    --text-0: #e8e6e3;
    --text-1: #a09e9a;
    --text-2: #6e6c68;
    --amber: #f0a500;
    --amber-dim: #b37e00;
    --amber-glow: rgba(240, 165, 0, 0.12);
    --red: #e04848;
    --green: #3cb56c;
    --blue: #4a9eff;
    --purple: #d4a0ff;
    --mono: 'JetBrains Mono', 'Consolas', 'Courier New', monospace;
    --sans: 'DM Sans', -apple-system, BlinkMacSystemFont, sans-serif;
  }

  html { font-size: 15px; }
  body {
    font-family: var(--sans);
    background: var(--bg-0);
    color: var(--text-0);
    min-height: 100vh;
    line-height: 1.5;
  }

  /* ---- Layout ---- */
  .shell {
    max-width: 1200px;
    margin: 0 auto;
    padding: 2rem 2rem 4rem;
  }

  header {
    display: flex;
    align-items: baseline;
    gap: 1rem;
    padding-bottom: 1.5rem;
    border-bottom: 1px solid var(--border);
    margin-bottom: 2rem;
  }
  header h1 {
    font-family: var(--mono);
    font-weight: 700;
    font-size: 1.4rem;
    letter-spacing: -0.03em;
    color: var(--amber);
  }
  header h1 span { color: var(--text-2); font-weight: 400; }
  header .tag {
    font-family: var(--mono);
    font-size: 0.7rem;
    color: var(--text-2);
    background: var(--bg-2);
    padding: 0.2em 0.6em;
    border-radius: 3px;
    border: 1px solid var(--border);
    letter-spacing: 0.05em;
    text-transform: uppercase;
  }

  .workspace {
    display: grid;
    grid-template-columns: 340px 1fr;
    gap: 1.5rem;
    align-items: start;
  }

  @media (max-width: 800px) {
    .workspace { grid-template-columns: 1fr; }
  }

  /* ---- Panel ---- */
  .panel {
    background: var(--bg-1);
    border: 1px solid var(--border);
    border-radius: 6px;
    overflow: hidden;
  }
  .panel-head {
    font-family: var(--mono);
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--text-2);
    padding: 0.8rem 1.2rem;
    border-bottom: 1px solid var(--border);
    background: var(--bg-2);
  }
  .panel-body { padding: 1.2rem; }

  /* ---- Drop zone ---- */
  .dropzone {
    border: 2px dashed var(--border);
    border-radius: 6px;
    padding: 2rem 1.2rem;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    position: relative;
  }
  .dropzone:hover,
  .dropzone.drag-over {
    border-color: var(--amber);
    background: var(--amber-glow);
  }
  .dropzone input[type="file"] {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
  }
  .dropzone-icon {
    font-size: 2rem;
    margin-bottom: 0.5rem;
    color: var(--text-2);
    transition: color 0.2s;
  }
  .dropzone:hover .dropzone-icon,
  .dropzone.drag-over .dropzone-icon { color: var(--amber); }
  .dropzone-label {
    font-size: 0.85rem;
    color: var(--text-1);
  }
  .dropzone-label strong { color: var(--amber); }
  .dropzone-hint {
    font-family: var(--mono);
    font-size: 0.68rem;
    color: var(--text-2);
    margin-top: 0.4rem;
  }

  .file-chip {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    background: var(--bg-2);
    border: 1px solid var(--border);
    border-radius: 5px;
    padding: 0.6rem 1rem;
    margin-top: 1rem;
  }
  .file-chip-icon { color: var(--amber); font-size: 1.1rem; }
  .file-chip-name {
    font-family: var(--mono);
    font-size: 0.78rem;
    color: var(--text-0);
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .file-chip-size {
    font-family: var(--mono);
    font-size: 0.68rem;
    color: var(--text-2);
  }
  .file-chip-remove {
    background: none;
    border: none;
    color: var(--text-2);
    cursor: pointer;
    font-size: 1rem;
    padding: 0;
    line-height: 1;
  }
  .file-chip-remove:hover { color: var(--red); }

  /* ---- Form controls ---- */
  .control-group {
    margin-top: 1.2rem;
  }
  .control-group label {
    display: block;
    font-family: var(--mono);
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: var(--text-2);
    margin-bottom: 0.4rem;
  }

  .select-wrap {
    position: relative;
  }
  .select-wrap select {
    width: 100%;
    appearance: none;
    font-family: var(--mono);
    font-size: 0.82rem;
    color: var(--text-0);
    background: var(--bg-2);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 0.55rem 2rem 0.55rem 0.8rem;
    cursor: pointer;
    transition: border-color 0.15s;
  }
  .select-wrap select:hover { border-color: var(--border-hover); }
  .select-wrap select:focus { border-color: var(--amber); outline: none; }
  .select-wrap::after {
    content: "\25BE";
    position: absolute;
    right: 0.8rem;
    top: 50%;
    transform: translateY(-50%);
    color: var(--text-2);
    pointer-events: none;
    font-size: 0.8rem;
  }

  .toggle-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-top: 1.2rem;
    padding: 0.6rem 0;
  }
  .toggle-row span {
    font-family: var(--mono);
    font-size: 0.78rem;
    color: var(--text-1);
  }

  .toggle {
    position: relative;
    width: 40px;
    height: 22px;
  }
  .toggle input {
    opacity: 0;
    width: 0;
    height: 0;
  }
  .toggle-track {
    position: absolute;
    inset: 0;
    background: var(--bg-3);
    border-radius: 11px;
    cursor: pointer;
    transition: background 0.2s;
    border: 1px solid var(--border);
  }
  .toggle-track::after {
    content: "";
    position: absolute;
    width: 16px;
    height: 16px;
    left: 2px;
    top: 2px;
    background: var(--text-2);
    border-radius: 50%;
    transition: transform 0.2s, background 0.2s;
  }
  .toggle input:checked + .toggle-track {
    background: var(--amber-dim);
    border-color: var(--amber);
  }
  .toggle input:checked + .toggle-track::after {
    transform: translateX(18px);
    background: var(--amber);
  }

  /* ---- Buttons ---- */
  .btn-run {
    width: 100%;
    margin-top: 1.5rem;
    font-family: var(--mono);
    font-size: 0.85rem;
    font-weight: 600;
    letter-spacing: 0.04em;
    color: var(--bg-0);
    background: var(--amber);
    border: none;
    border-radius: 5px;
    padding: 0.75rem 1.5rem;
    cursor: pointer;
    transition: background 0.15s, transform 0.1s;
    position: relative;
    overflow: hidden;
  }
  .btn-run:hover { background: #ffb515; }
  .btn-run:active { transform: scale(0.98); }
  .btn-run:disabled {
    background: var(--bg-3);
    color: var(--text-2);
    cursor: not-allowed;
    transform: none;
  }
  .btn-run.processing {
    background: var(--amber-dim);
    pointer-events: none;
  }
  .btn-run.processing::after {
    content: "";
    position: absolute;
    left: 0;
    top: 0;
    height: 100%;
    width: 30%;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent);
    animation: shimmer 1.2s infinite;
  }
  @keyframes shimmer { 0%{left:-30%} 100%{left:100%} }

  /* ---- Output panel ---- */
  .output-panel {
    min-height: 400px;
    display: flex;
    flex-direction: column;
  }
  .output-panel .panel-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .output-actions {
    display: flex;
    gap: 0.5rem;
  }
  .btn-sm {
    font-family: var(--mono);
    font-size: 0.68rem;
    font-weight: 500;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    color: var(--text-1);
    background: var(--bg-1);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0.3em 0.7em;
    cursor: pointer;
    transition: border-color 0.15s, color 0.15s;
  }
  .btn-sm:hover {
    border-color: var(--amber);
    color: var(--amber);
  }

  .output-tabs {
    display: flex;
    border-bottom: 1px solid var(--border);
  }
  .output-tab {
    font-family: var(--mono);
    font-size: 0.72rem;
    font-weight: 500;
    color: var(--text-2);
    background: none;
    border: none;
    border-bottom: 2px solid transparent;
    padding: 0.6rem 1.2rem;
    cursor: pointer;
    transition: color 0.15s, border-color 0.15s;
  }
  .output-tab:hover { color: var(--text-1); }
  .output-tab.active {
    color: var(--amber);
    border-bottom-color: var(--amber);
  }

  .output-content {
    flex: 1;
    position: relative;
  }

  .output-view {
    display: none;
    height: 100%;
  }
  .output-view.active { display: block; }

  .output-raw {
    font-family: var(--mono);
    font-size: 0.78rem;
    line-height: 1.65;
    color: var(--text-0);
    background: var(--bg-0);
    border: none;
    padding: 1.2rem;
    width: 100%;
    min-height: 400px;
    resize: vertical;
    overflow: auto;
  }
  .output-raw:focus { outline: none; }

  .output-rendered {
    padding: 1.5rem;
    overflow: auto;
    min-height: 400px;
    max-height: 75vh;
  }

  /* ---- Rendered markdown styles ---- */
  .output-rendered h1 {
    font-family: var(--mono);
    font-size: 1.3rem;
    font-weight: 700;
    color: var(--amber);
    margin-bottom: 0.8rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid var(--border);
  }
  .output-rendered h2 {
    font-family: var(--mono);
    font-size: 1rem;
    font-weight: 600;
    color: var(--text-0);
    margin-top: 2rem;
    margin-bottom: 0.6rem;
    padding-bottom: 0.3rem;
    border-bottom: 1px solid var(--border);
  }
  .output-rendered h3 {
    font-family: var(--mono);
    font-size: 0.88rem;
    font-weight: 600;
    color: var(--amber);
    margin-top: 1.5rem;
    margin-bottom: 0.5rem;
  }
  .output-rendered blockquote {
    font-size: 0.82rem;
    color: var(--text-1);
    border-left: 3px solid var(--amber-dim);
    padding: 0.4rem 1rem;
    margin: 0.6rem 0;
    background: var(--amber-glow);
    border-radius: 0 4px 4px 0;
  }
  .output-rendered ul {
    margin: 0.3rem 0 0.6rem 1.5rem;
    font-size: 0.82rem;
    color: var(--text-1);
  }
  .output-rendered li { margin-bottom: 0.15rem; }
  .output-rendered pre {
    background: var(--bg-0);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 1rem 1.2rem;
    overflow-x: auto;
    margin: 0.4rem 0 1rem;
  }
  .output-rendered code {
    font-family: var(--mono);
    font-size: 0.78rem;
    line-height: 1.6;
    color: var(--text-0);
  }
  .output-rendered strong { color: var(--text-0); }

  /* M syntax highlighting */
  .output-rendered .kw { color: var(--blue); }
  .output-rendered .str { color: var(--green); }
  .output-rendered .cmt { color: var(--text-2); font-style: italic; }
  .output-rendered .fn { color: #d4a0ff; }
  .output-rendered .num { color: #e0905a; }
  .output-rendered .op { color: var(--text-1); }

  /* ---- Graph view ---- */
  .graph-container {
    position: relative;
    min-height: 450px;
    background: var(--bg-0);
    overflow: hidden;
    cursor: grab;
  }
  .graph-container:active { cursor: grabbing; }
  .graph-container svg { display: block; }

  .graph-node rect {
    transition: opacity 0.2s;
  }
  .graph-node text {
    font-family: var(--mono);
    pointer-events: none;
    transition: opacity 0.2s;
  }
  .graph-edge {
    fill: none;
    stroke-width: 1.5;
    transition: opacity 0.2s, stroke 0.2s;
  }
  .graph-connector-badge {
    font-family: var(--mono);
    pointer-events: none;
  }

  /* Dimming on hover */
  .graph-container.has-hover .graph-node.dimmed rect { opacity: 0.15; }
  .graph-container.has-hover .graph-node.dimmed text { opacity: 0.15; }
  .graph-container.has-hover .graph-edge.dimmed { opacity: 0.06; }

  .graph-node.highlighted rect { stroke-width: 2; }

  /* Tooltip */
  .graph-tooltip {
    position: absolute;
    background: var(--bg-2);
    border: 1px solid var(--border);
    border-radius: 5px;
    padding: 0.7rem 1rem;
    font-family: var(--mono);
    font-size: 0.72rem;
    color: var(--text-1);
    pointer-events: none;
    opacity: 0;
    transition: opacity 0.15s;
    max-width: 300px;
    z-index: 10;
    line-height: 1.6;
  }
  .graph-tooltip.visible { opacity: 1; }
  .graph-tooltip strong { color: var(--text-0); display: block; margin-bottom: 0.2rem; }
  .graph-tooltip .tt-kind { color: var(--amber); }
  .graph-tooltip .tt-dep { color: var(--blue); }
  .graph-tooltip .tt-src { color: var(--green); }

  /* Graph legend */
  .graph-legend {
    position: absolute;
    bottom: 0.8rem;
    left: 0.8rem;
    display: flex;
    gap: 1rem;
    font-family: var(--mono);
    font-size: 0.65rem;
    color: var(--text-2);
    background: rgba(12, 12, 14, 0.85);
    padding: 0.4rem 0.8rem;
    border-radius: 4px;
    border: 1px solid var(--border);
  }
  .graph-legend-item {
    display: flex;
    align-items: center;
    gap: 0.3rem;
  }
  .graph-legend-dot {
    width: 8px;
    height: 8px;
    border-radius: 2px;
  }

  .graph-controls {
    position: absolute;
    top: 0.8rem;
    right: 0.8rem;
    display: flex;
    gap: 0.3rem;
  }
  .graph-ctrl-btn {
    width: 28px;
    height: 28px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: var(--bg-2);
    border: 1px solid var(--border);
    border-radius: 4px;
    color: var(--text-2);
    font-family: var(--mono);
    font-size: 0.82rem;
    cursor: pointer;
    transition: border-color 0.15s, color 0.15s;
  }
  .graph-ctrl-btn:hover {
    border-color: var(--amber);
    color: var(--amber);
  }

  /* ---- Graph fullscreen overlay ---- */
  .graph-fullscreen-overlay {
    display: none;
    position: fixed;
    inset: 0;
    z-index: 1000;
    background: var(--bg-0);
    flex-direction: column;
  }
  .graph-fullscreen-overlay.active {
    display: flex;
  }
  .graph-fullscreen-overlay .fs-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0.8rem 1.5rem;
    border-bottom: 1px solid var(--border);
    background: var(--bg-1);
    flex-shrink: 0;
  }
  .graph-fullscreen-overlay .fs-title {
    font-family: var(--mono);
    font-size: 0.82rem;
    font-weight: 600;
    color: var(--amber);
  }
  .graph-fullscreen-overlay .fs-close {
    font-family: var(--mono);
    font-size: 0.75rem;
    color: var(--text-1);
    background: var(--bg-2);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 0.35em 0.8em;
    cursor: pointer;
    transition: border-color 0.15s, color 0.15s;
  }
  .graph-fullscreen-overlay .fs-close:hover {
    border-color: var(--amber);
    color: var(--amber);
  }
  .graph-fullscreen-overlay .fs-body {
    flex: 1;
    position: relative;
    overflow: hidden;
    cursor: grab;
  }
  .graph-fullscreen-overlay .fs-body:active { cursor: grabbing; }
  .graph-fullscreen-overlay .fs-body svg { display: block; }
  .graph-fullscreen-overlay .graph-legend {
    position: absolute;
    bottom: 0.8rem;
    left: 0.8rem;
  }
  .graph-fullscreen-overlay .graph-controls {
    position: absolute;
    top: 0.8rem;
    right: 0.8rem;
  }

  /* ---- Stats bar ---- */
  .stats-bar {
    display: flex;
    gap: 0.5rem;
    flex-wrap: wrap;
    padding: 0.6rem 1.2rem;
    border-top: 1px solid var(--border);
    background: var(--bg-2);
    font-family: var(--mono);
    font-size: 0.7rem;
    color: var(--text-2);
  }
  .stat-chip {
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    background: var(--bg-1);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0.2em 0.6em;
  }
  .stat-chip .dot {
    width: 6px;
    height: 6px;
    border-radius: 50%;
  }
  .dot-param { background: var(--blue); }
  .dot-query { background: var(--amber); }
  .dot-func  { background: #d4a0ff; }
  .dot-dep   { background: var(--text-2); }

  /* ---- Empty / error states ---- */
  .empty-state {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 400px;
    color: var(--text-2);
    text-align: center;
    padding: 2rem;
  }
  .empty-state-icon {
    font-size: 3rem;
    margin-bottom: 1rem;
    opacity: 0.3;
  }
  .empty-state-text {
    font-family: var(--mono);
    font-size: 0.82rem;
  }

  .error-banner {
    background: rgba(224, 72, 72, 0.1);
    border: 1px solid rgba(224, 72, 72, 0.3);
    border-radius: 4px;
    padding: 0.7rem 1rem;
    margin-top: 1rem;
    font-family: var(--mono);
    font-size: 0.78rem;
    color: var(--red);
  }

  /* ---- Scrollbar ---- */
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--bg-3); border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--border-hover); }
</style>
</head>
<body>
<div class="shell">
  <header>
    <h1>M<span>_</span>EXPORTER</h1>
    <span class="tag">Power Query</span>
  </header>

  <div class="workspace">
    <!-- Left: controls -->
    <div>
      <div class="panel">
        <div class="panel-head">Input</div>
        <div class="panel-body">
          <div class="dropzone" id="dropzone">
            <input type="file" id="fileInput" accept=".pbix,.pbit" />
            <div class="dropzone-icon">&#9776;</div>
            <div class="dropzone-label">Drop file or <strong>browse</strong></div>
            <div class="dropzone-hint">.pbix &middot; .pbit</div>
          </div>
          <div id="fileChip" style="display:none" class="file-chip">
            <span class="file-chip-icon">&#9830;</span>
            <span class="file-chip-name" id="fileName"></span>
            <span class="file-chip-size" id="fileSize"></span>
            <button class="file-chip-remove" id="fileRemove">&times;</button>
          </div>

          <div class="control-group">
            <label>Extraction Mode</label>
            <div class="select-wrap">
              <select id="modeSelect">
                <option value="auto">Auto-detect</option>
                <option value="pbix-datamashup">PBIX / DataMashup</option>
                <option value="pbit-json">PBIT / JSON Schema</option>
              </select>
            </div>
          </div>

          <div class="toggle-row">
            <span>Include metadata</span>
            <label class="toggle">
              <input type="checkbox" id="metadataToggle" />
              <span class="toggle-track"></span>
            </label>
          </div>

          <div id="errorBox" class="error-banner" style="display:none"></div>

          <button class="btn-run" id="runBtn" disabled>EXTRACT</button>
        </div>
      </div>
    </div>

    <!-- Right: output -->
    <div class="panel output-panel" id="outputPanel">
      <div class="panel-head">
        <span>Output</span>
        <div class="output-actions" id="outputActions" style="display:none">
          <button class="btn-sm" id="copyBtn">Copy</button>
          <button class="btn-sm" id="downloadBtn">Download .md</button>
        </div>
      </div>
      <div class="output-tabs" id="outputTabs" style="display:none">
        <button class="output-tab active" data-tab="graph">Graph</button>
        <button class="output-tab" data-tab="rendered">Rendered</button>
        <button class="output-tab" data-tab="raw">Markdown</button>
      </div>
      <div class="output-content">
        <div class="output-view active" id="view-graph">
          <div class="empty-state" id="emptyState">
            <div class="empty-state-icon">&#9998;</div>
            <div class="empty-state-text">Upload a .pbix or .pbit file to extract M code</div>
          </div>
          <div class="graph-container" id="graphContainer" style="display:none">
            <svg id="graphSvg"></svg>
            <div class="graph-tooltip" id="graphTooltip"></div>
            <div class="graph-legend">
              <div class="graph-legend-item"><div class="graph-legend-dot" style="background:var(--blue)"></div> Parameter</div>
              <div class="graph-legend-item"><div class="graph-legend-dot" style="background:var(--amber)"></div> Query</div>
              <div class="graph-legend-item"><div class="graph-legend-dot" style="background:var(--purple)"></div> Function</div>
            </div>
            <div class="graph-controls">
              <button class="graph-ctrl-btn" id="popoutBtn" title="Pop out fullscreen">&#11036;</button>
              <button class="graph-ctrl-btn" id="zoomIn" title="Zoom in">+</button>
              <button class="graph-ctrl-btn" id="zoomOut" title="Zoom out">&minus;</button>
              <button class="graph-ctrl-btn" id="zoomFit" title="Fit to view">&#8690;</button>
            </div>
          </div>
        </div>
        <div class="output-view" id="view-rendered">
          <div class="output-rendered" id="renderedOutput" style="display:none"></div>
        </div>
        <div class="output-view" id="view-raw">
          <textarea class="output-raw" id="rawOutput" readonly></textarea>
        </div>
      </div>
      <div class="stats-bar" id="statsBar" style="display:none">
        <span class="stat-chip"><span class="dot dot-param"></span> <span id="statParams">0</span> params</span>
        <span class="stat-chip"><span class="dot dot-query"></span> <span id="statQueries">0</span> queries</span>
        <span class="stat-chip"><span class="dot dot-func"></span> <span id="statFuncs">0</span> functions</span>
        <span class="stat-chip"><span class="dot dot-dep"></span> <span id="statDeps">0</span> deps</span>
      </div>
    </div>
  </div>
</div>

<!-- Fullscreen graph overlay -->
<div class="graph-fullscreen-overlay" id="fsOverlay">
  <div class="fs-header">
    <span class="fs-title">Dependency Graph</span>
    <button class="fs-close" id="fsClose">ESC &times; Close</button>
  </div>
  <div class="fs-body" id="fsBody">
    <svg id="fsSvg"></svg>
    <div class="graph-tooltip" id="fsTooltip"></div>
    <div class="graph-legend">
      <div class="graph-legend-item"><div class="graph-legend-dot" style="background:var(--blue)"></div> Parameter</div>
      <div class="graph-legend-item"><div class="graph-legend-dot" style="background:var(--amber)"></div> Query</div>
      <div class="graph-legend-item"><div class="graph-legend-dot" style="background:var(--purple)"></div> Function</div>
    </div>
    <div class="graph-controls">
      <button class="graph-ctrl-btn" id="fsZoomIn" title="Zoom in">+</button>
      <button class="graph-ctrl-btn" id="fsZoomOut" title="Zoom out">&minus;</button>
      <button class="graph-ctrl-btn" id="fsZoomFit" title="Fit to view">&#8690;</button>
    </div>
  </div>
</div>

<script>
(function() {
  const dropzone     = document.getElementById('dropzone');
  const fileInput    = document.getElementById('fileInput');
  const fileChip     = document.getElementById('fileChip');
  const fileName     = document.getElementById('fileName');
  const fileSize     = document.getElementById('fileSize');
  const fileRemove   = document.getElementById('fileRemove');
  const modeSelect   = document.getElementById('modeSelect');
  const metadataToggle = document.getElementById('metadataToggle');
  const runBtn       = document.getElementById('runBtn');
  const errorBox     = document.getElementById('errorBox');
  const outputTabs   = document.getElementById('outputTabs');
  const outputActions = document.getElementById('outputActions');
  const copyBtn      = document.getElementById('copyBtn');
  const downloadBtn  = document.getElementById('downloadBtn');
  const emptyState   = document.getElementById('emptyState');
  const renderedOutput = document.getElementById('renderedOutput');
  const rawOutput    = document.getElementById('rawOutput');
  const statsBar     = document.getElementById('statsBar');
  const graphContainer = document.getElementById('graphContainer');
  const graphSvg     = document.getElementById('graphSvg');
  const graphTooltip = document.getElementById('graphTooltip');

  let selectedFile = null;
  let lastMarkdown = '';
  let lastFileName = '';

  // --- File handling ---
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

  // --- Tabs ---
  document.querySelectorAll('.output-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.output-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.output-view').forEach(v => v.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById('view-' + tab.dataset.tab).classList.add('active');
    });
  });

  // --- Error ---
  function showError(msg) {
    errorBox.textContent = msg;
    errorBox.style.display = 'block';
  }
  function hideError() { errorBox.style.display = 'none'; }

  // --- Minimal markdown-to-HTML renderer ---
  function highlightM(code) {
    code = code.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    code = code.replace(/(\/\/.*$)/gm, '<span class="cmt">$1</span>');
    code = code.replace(/("(?:[^"\\]|\\.)*")/g, '<span class="str">$1</span>');
    const kws = ['let','in','if','then','else','each','true','false','null','not','and','or','meta','type','shared','section','try','otherwise','error','as','is'];
    const kwRe = new RegExp('\\b(' + kws.join('|') + ')\\b', 'g');
    code = code.replace(kwRe, '<span class="kw">$1</span>');
    code = code.replace(/\b([A-Z][a-zA-Z]+\.[A-Z][a-zA-Z]+)\b/g, '<span class="fn">$1</span>');
    code = code.replace(/\b(\d+(?:\.\d+)?)\b/g, '<span class="num">$1</span>');
    return code;
  }

  function renderMd(md) {
    const lines = md.split('\n');
    let html = '';
    let inCode = false;
    let codeBlock = '';
    let inList = false;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line.startsWith('```')) {
        if (inCode) {
          html += '<pre><code>' + highlightM(codeBlock) + '</code></pre>';
          codeBlock = '';
          inCode = false;
        } else {
          if (inList) { html += '</ul>'; inList = false; }
          inCode = true;
        }
        continue;
      }
      if (inCode) { codeBlock += line + '\n'; continue; }
      if (line.startsWith('### ')) {
        if (inList) { html += '</ul>'; inList = false; }
        html += '<h3>' + line.slice(4) + '</h3>';
      } else if (line.startsWith('## ')) {
        if (inList) { html += '</ul>'; inList = false; }
        html += '<h2>' + line.slice(3) + '</h2>';
      } else if (line.startsWith('# ')) {
        if (inList) { html += '</ul>'; inList = false; }
        html += '<h1>' + line.slice(2) + '</h1>';
      } else if (line.startsWith('> ')) {
        if (inList) { html += '</ul>'; inList = false; }
        html += '<blockquote>' + line.slice(2).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>') + '</blockquote>';
      } else if (line.match(/^\s*- /)) {
        if (!inList) { html += '<ul>'; inList = true; }
        let content = line.replace(/^\s*- /, '').replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
        content = content.replace(/`([^`]+)`/g, '<code>$1</code>');
        html += '<li>' + content + '</li>';
      } else if (line.trim() === '') {
        if (inList) { html += '</ul>'; inList = false; }
      } else {
        html += '<p>' + line.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>') + '</p>';
      }
    }
    if (inList) html += '</ul>';
    return html;
  }

  // =====================================================================
  // GRAPH RENDERER
  // =====================================================================
  const NODE_W = 160;
  const NODE_H = 36;
  const LAYER_GAP = 100;
  const NODE_GAP = 30;
  const PAD_X = 60;
  const PAD_Y = 50;

  const KIND_COLORS = {
    'Parameter': '#4a9eff',
    'Query':     '#f0a500',
    'Function':  '#d4a0ff',
  };

  // Store last graph data for pop-out
  let lastGraphData = null;

  /**
   * Build an interactive graph into a target container.
   * Returns { fitToView } for external control.
   */
  function buildGraphInto(graphData, targetSvg, targetContainer, targetTooltip, zoomInBtn, zoomOutBtn, zoomFitBtn) {
    const { nodes, edges, layers, data_sources } = graphData;
    if (!nodes.length) return null;

    const nodeMap = {};
    nodes.forEach(n => { nodeMap[n.name] = n; });

    const depsOf = {};
    const usedBy = {};
    edges.forEach(e => {
      if (!depsOf[e.from]) depsOf[e.from] = [];
      depsOf[e.from].push(e.to);
      if (!usedBy[e.to]) usedBy[e.to] = [];
      usedBy[e.to].push(e.from);
    });

    const positions = {};
    let maxLayerWidth = 0;
    layers.forEach(layer => { maxLayerWidth = Math.max(maxLayerWidth, layer.length); });

    const svgW = Math.max(maxLayerWidth * (NODE_W + NODE_GAP) + PAD_X * 2, 600);
    const svgH = layers.length * (NODE_H + LAYER_GAP) + PAD_Y * 2;

    layers.forEach((layer, li) => {
      const totalW = layer.length * NODE_W + (layer.length - 1) * NODE_GAP;
      const startX = (svgW - totalW) / 2;
      const y = PAD_Y + li * (NODE_H + LAYER_GAP);
      layer.forEach((name, ni) => {
        positions[name] = { x: startX + ni * (NODE_W + NODE_GAP), y };
      });
    });

    const ns = 'http://www.w3.org/2000/svg';
    targetSvg.innerHTML = '';
    targetSvg.setAttribute('width', svgW);
    targetSvg.setAttribute('height', svgH);
    targetSvg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);

    // Unique marker ID per instance
    const markerId = 'ah-' + Math.random().toString(36).slice(2, 6);
    const defs = document.createElementNS(ns, 'defs');
    const marker = document.createElementNS(ns, 'marker');
    marker.setAttribute('id', markerId);
    marker.setAttribute('markerWidth', '8');
    marker.setAttribute('markerHeight', '6');
    marker.setAttribute('refX', '8');
    marker.setAttribute('refY', '3');
    marker.setAttribute('orient', 'auto');
    const arrowPath = document.createElementNS(ns, 'path');
    arrowPath.setAttribute('d', 'M0,0 L0,6 L8,3 z');
    arrowPath.setAttribute('fill', '#3e3e4a');
    marker.appendChild(arrowPath);
    defs.appendChild(marker);
    targetSvg.appendChild(defs);

    const rootG = document.createElementNS(ns, 'g');
    targetSvg.appendChild(rootG);

    const edgeGroup = document.createElementNS(ns, 'g');
    rootG.appendChild(edgeGroup);

    const edgeEls = [];
    edges.forEach(e => {
      const fromPos = positions[e.from];
      const toPos = positions[e.to];
      if (!fromPos || !toPos) return;
      const x1 = toPos.x + NODE_W / 2, y1 = toPos.y + NODE_H;
      const x2 = fromPos.x + NODE_W / 2, y2 = fromPos.y;
      const cp = Math.max(Math.abs(y2 - y1) * 0.4, 30);
      const path = document.createElementNS(ns, 'path');
      path.setAttribute('d', `M${x1},${y1} C${x1},${y1+cp} ${x2},${y2-cp} ${x2},${y2}`);
      path.setAttribute('class', 'graph-edge');
      path.setAttribute('stroke', '#3e3e4a');
      path.setAttribute('marker-end', `url(#${markerId})`);
      path.dataset.from = e.from;
      path.dataset.to = e.to;
      edgeGroup.appendChild(path);
      edgeEls.push(path);
    });

    const nodeGroup = document.createElementNS(ns, 'g');
    rootG.appendChild(nodeGroup);

    const nodeEls = {};
    nodes.forEach(n => {
      const pos = positions[n.name];
      if (!pos) return;
      const g = document.createElementNS(ns, 'g');
      g.setAttribute('class', 'graph-node');
      g.setAttribute('transform', `translate(${pos.x}, ${pos.y})`);
      g.dataset.name = n.name;

      const rect = document.createElementNS(ns, 'rect');
      rect.setAttribute('width', NODE_W);
      rect.setAttribute('height', NODE_H);
      rect.setAttribute('rx', '4');
      rect.setAttribute('fill', '#1c1c22');
      rect.setAttribute('stroke', KIND_COLORS[n.kind] || '#3e3e4a');
      rect.setAttribute('stroke-width', '1.5');
      g.appendChild(rect);

      if (data_sources[n.name]) {
        const badge = document.createElementNS(ns, 'circle');
        badge.setAttribute('cx', NODE_W - 10);
        badge.setAttribute('cy', 10);
        badge.setAttribute('r', '4');
        badge.setAttribute('fill', '#3cb56c');
        g.appendChild(badge);
      }

      const text = document.createElementNS(ns, 'text');
      text.setAttribute('x', NODE_W / 2);
      text.setAttribute('y', NODE_H / 2 + 1);
      text.setAttribute('text-anchor', 'middle');
      text.setAttribute('dominant-baseline', 'middle');
      text.setAttribute('fill', '#e8e6e3');
      text.setAttribute('font-size', '11');
      let label = n.name;
      if (label.length > 18) label = label.slice(0, 16) + '..';
      text.textContent = label;
      g.appendChild(text);

      nodeGroup.appendChild(g);
      nodeEls[n.name] = g;
    });

    // Hover highlighting
    function highlightNode(name) {
      targetContainer.classList.add('has-hover');
      const connected = new Set([name]);
      edges.forEach(e => {
        if (e.from === name || e.to === name) { connected.add(e.from); connected.add(e.to); }
      });
      Object.entries(nodeEls).forEach(([n, el]) => {
        el.classList.toggle('dimmed', !connected.has(n));
        el.classList.toggle('highlighted', n === name);
      });
      edgeEls.forEach(el => {
        const isConn = el.dataset.from === name || el.dataset.to === name;
        el.classList.toggle('dimmed', !isConn);
        if (isConn) el.setAttribute('stroke', KIND_COLORS[nodeMap[name]?.kind] || '#f0a500');
      });
    }

    function clearHighlight() {
      targetContainer.classList.remove('has-hover');
      Object.values(nodeEls).forEach(el => el.classList.remove('dimmed', 'highlighted'));
      edgeEls.forEach(el => { el.classList.remove('dimmed'); el.setAttribute('stroke', '#3e3e4a'); });
      targetTooltip.classList.remove('visible');
    }

    Object.entries(nodeEls).forEach(([name, el]) => {
      el.addEventListener('mouseenter', () => {
        highlightNode(name);
        const node = nodeMap[name];
        let html = `<strong>${name}</strong><span class="tt-kind">${node.kind}</span>`;
        const deps = depsOf[name];
        if (deps && deps.length) html += `<br><span class="tt-dep">depends on: ${deps.join(', ')}</span>`;
        const users = usedBy[name];
        if (users && users.length) html += `<br><span class="tt-dep">used by: ${users.join(', ')}</span>`;
        const srcs = data_sources[name];
        if (srcs && srcs.length) html += `<br><span class="tt-src">source: ${srcs.join(', ')}</span>`;
        targetTooltip.innerHTML = html;
        targetTooltip.classList.add('visible');
      });
      el.addEventListener('mouseleave', clearHighlight);
    });

    targetContainer.addEventListener('mousemove', (ev) => {
      const r = targetContainer.getBoundingClientRect();
      let x = ev.clientX - r.left + 15, y = ev.clientY - r.top + 15;
      if (x + 250 > r.width) x = ev.clientX - r.left - 260;
      if (y + 100 > r.height) y = ev.clientY - r.top - 80;
      targetTooltip.style.left = x + 'px';
      targetTooltip.style.top = y + 'px';
    });

    // Pan / Zoom state
    const state = { scale: 1, tx: 0, ty: 0, dragging: false, lastX: 0, lastY: 0 };

    function applyTransform() {
      rootG.setAttribute('transform', `translate(${state.tx},${state.ty}) scale(${state.scale})`);
    }

    function fitToView() {
      const cr = targetContainer.getBoundingClientRect();
      const cw = cr.width || 800, ch = cr.height || 450;
      const sx = cw / svgW, sy = ch / svgH;
      state.scale = Math.min(sx, sy, 1) * 0.92;
      state.tx = (cw - svgW * state.scale) / 2;
      state.ty = (ch - svgH * state.scale) / 2;
      applyTransform();
    }

    if (zoomInBtn) zoomInBtn.onclick = () => { state.scale = Math.min(state.scale * 1.25, 3); applyTransform(); };
    if (zoomOutBtn) zoomOutBtn.onclick = () => { state.scale = Math.max(state.scale * 0.8, 0.15); applyTransform(); };
    if (zoomFitBtn) zoomFitBtn.onclick = fitToView;

    targetContainer.addEventListener('wheel', (ev) => {
      ev.preventDefault();
      const delta = ev.deltaY > 0 ? 0.9 : 1.1;
      const ns2 = Math.max(0.15, Math.min(3, state.scale * delta));
      const r = targetContainer.getBoundingClientRect();
      const mx = ev.clientX - r.left, my = ev.clientY - r.top;
      state.tx = mx - (mx - state.tx) * (ns2 / state.scale);
      state.ty = my - (my - state.ty) * (ns2 / state.scale);
      state.scale = ns2;
      applyTransform();
    }, { passive: false });

    targetContainer.addEventListener('mousedown', (ev) => {
      if (ev.target.closest('.graph-node')) return;
      state.dragging = true;
      state.lastX = ev.clientX;
      state.lastY = ev.clientY;
    });
    const onMove = (ev) => {
      if (!state.dragging) return;
      state.tx += ev.clientX - state.lastX;
      state.ty += ev.clientY - state.lastY;
      state.lastX = ev.clientX;
      state.lastY = ev.clientY;
      applyTransform();
    };
    const onUp = () => { state.dragging = false; };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);

    return { fitToView };
  }

  /**
   * Render graph into the inline panel.
   */
  function renderGraph(graphData) {
    lastGraphData = graphData;
    const ctrl = buildGraphInto(
      graphData, graphSvg, graphContainer, graphTooltip,
      document.getElementById('zoomIn'),
      document.getElementById('zoomOut'),
      document.getElementById('zoomFit')
    );
    if (!ctrl) return;
    graphContainer.style.display = 'block';
    emptyState.style.display = 'none';
    requestAnimationFrame(ctrl.fitToView);
  }

  // =====================================================================
  // FULLSCREEN POP-OUT
  // =====================================================================
  const fsOverlay  = document.getElementById('fsOverlay');
  const fsBody     = document.getElementById('fsBody');
  const fsSvg      = document.getElementById('fsSvg');
  const fsTooltip  = document.getElementById('fsTooltip');
  const fsClose    = document.getElementById('fsClose');

  document.getElementById('popoutBtn').addEventListener('click', () => {
    if (!lastGraphData) return;
    fsOverlay.classList.add('active');
    document.body.style.overflow = 'hidden';
    requestAnimationFrame(() => {
      const ctrl = buildGraphInto(
        lastGraphData, fsSvg, fsBody, fsTooltip,
        document.getElementById('fsZoomIn'),
        document.getElementById('fsZoomOut'),
        document.getElementById('fsZoomFit')
      );
      if (ctrl) requestAnimationFrame(ctrl.fitToView);
    });
  });

  function closeFullscreen() {
    fsOverlay.classList.remove('active');
    document.body.style.overflow = '';
    fsSvg.innerHTML = '';
  }

  fsClose.addEventListener('click', closeFullscreen);
  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape' && fsOverlay.classList.contains('active')) closeFullscreen();
  });

  // --- Extract ---
  runBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    hideError();
    runBtn.classList.add('processing');
    runBtn.textContent = 'EXTRACTING...';
    runBtn.disabled = true;

    const form = new FormData();
    form.append('file', selectedFile);
    form.append('mode', modeSelect.value);
    form.append('metadata', metadataToggle.checked ? '1' : '0');

    try {
      const resp = await fetch('/api/extract', { method: 'POST', body: form });
      const data = await resp.json();

      if (data.error) {
        showError(data.error);
        return;
      }

      lastMarkdown = data.markdown;
      lastFileName = selectedFile.name.replace(/\.(pbix|pbit)$/i, '') + '.md';

      // Show rendered + raw
      renderedOutput.style.display = 'block';
      renderedOutput.innerHTML = renderMd(data.markdown);
      rawOutput.value = data.markdown;
      outputTabs.style.display = 'flex';
      outputActions.style.display = 'flex';
      statsBar.style.display = 'flex';

      document.getElementById('statParams').textContent  = data.stats.parameters;
      document.getElementById('statQueries').textContent  = data.stats.queries;
      document.getElementById('statFuncs').textContent    = data.stats.functions;
      document.getElementById('statDeps').textContent     = data.graph.edges.length;

      // Render graph
      if (data.graph && data.graph.nodes.length > 0) {
        renderGraph(data.graph);
      }

    } catch (err) {
      showError('Request failed: ' + err.message);
    } finally {
      runBtn.classList.remove('processing');
      runBtn.textContent = 'EXTRACT';
      runBtn.disabled = !selectedFile;
    }
  });

  // --- Copy / Download ---
  copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(lastMarkdown).then(() => {
      copyBtn.textContent = 'Copied!';
      setTimeout(() => copyBtn.textContent = 'Copy', 1500);
    });
  });

  downloadBtn.addEventListener('click', () => {
    const blob = new Blob([lastMarkdown], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = lastFileName;
    a.click();
    URL.revokeObjectURL(url);
  });

})();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        # Quieter logging: one-line format
        sys.stderr.write(f"[server] {args[0]}\n")

    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(INDEX_HTML.encode("utf-8"))
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path != "/api/extract":
            self.send_error(404)
            return

        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            self._json_response(400, {"error": "Expected multipart/form-data"})
            return

        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": content_type,
                },
            )

            file_field = form["file"]
            if not file_field.file:
                self._json_response(400, {"error": "No file uploaded"})
                return

            file_data = file_field.file.read()
            file_name = file_field.filename or "upload.pbix"
            mode = form.getfirst("mode", "auto")
            include_metadata = form.getfirst("metadata", "0") == "1"

            # Write to temp file (zipfile needs seekable file)
            suffix = Path(file_name).suffix.lower()
            with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                tmp.write(file_data)
                tmp_path = Path(tmp.name)

            try:
                # Determine mode
                if mode == "auto":
                    mode = detect_mode(tmp_path)

                if mode == "pbit-json":
                    items = extract_from_pbit_json(tmp_path)
                else:
                    items = extract_from_datamashup(tmp_path)

                # Analyze dependencies
                dep_graph = analyze_dependencies(items)

                md = render_markdown(items, file_name, include_metadata, dep_graph)

                stats = {
                    "parameters": sum(1 for i in items if i.kind == MItemKind.PARAMETER),
                    "queries": sum(1 for i in items if i.kind == MItemKind.QUERY),
                    "functions": sum(1 for i in items if i.kind == MItemKind.FUNCTION),
                    "total": len(items),
                }

                self._json_response(200, {
                    "markdown": md,
                    "stats": stats,
                    "graph": dependency_graph_to_dict(dep_graph),
                })
            finally:
                tmp_path.unlink(missing_ok=True)

        except KeyError as exc:
            self._json_response(400, {"error": f"Missing entry in archive: {exc}"})
        except Exception as exc:
            self._json_response(500, {"error": str(exc)})

    def _json_response(self, code: int, obj: dict):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    server = HTTPServer((HOST, PORT), Handler)
    print(f"M Exporter running at http://{HOST}:{PORT}")
    print(f"  -> Open http://100.90.11.37:{PORT} in your browser")
    print("  -> Ctrl+C to stop")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
        server.server_close()


if __name__ == "__main__":
    main()
