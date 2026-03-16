/**
 * Power BI M Exporter — client-side extraction engine.
 * Ports the Python pbix_m_export logic to browser JS.
 * Requires JSZip (loaded via CDN in index.html).
 */

// ---------------------------------------------------------------------------
// Data model
// ---------------------------------------------------------------------------

const MItemKind = {
  QUERY: 'Query',
  PARAMETER: 'Parameter',
  FUNCTION: 'Function',
};

class MItem {
  constructor(name, kind, mCode, metadata = {}) {
    this.name = name;
    this.kind = kind;
    this.mCode = mCode;
    this.metadata = metadata;
  }
}

class DependencyGraph {
  constructor(edges, dataSources, nodes, layers) {
    this.edges = edges;         // [{from, to}, ...]
    this.dataSources = dataSources; // {queryName: [connector, ...]}
    this.nodes = nodes;         // [{name, kind}, ...]
    this.layers = layers;       // [[name, ...], ...]
  }
}

// ---------------------------------------------------------------------------
// Known connectors
// ---------------------------------------------------------------------------

const KNOWN_CONNECTORS = new Set([
  'Sql.Database', 'Sql.Databases', 'Oracle.Database',
  'Odbc.DataSource', 'Odbc.Query', 'OleDb.DataSource', 'OleDb.Query',
  'PostgreSQL.Database', 'MySQL.Database', 'Snowflake.Databases',
  'Excel.Workbook', 'Excel.CurrentWorkbook',
  'Csv.Document', 'Json.Document', 'Xml.Document', 'Xml.Tables',
  'File.Contents', 'Folder.Files', 'Folder.Contents',
  'Web.Contents', 'Web.Page', 'Web.BrowserContents',
  'OData.Feed',
  'Salesforce.Data', 'Salesforce.Reports',
  'GoogleAnalytics.Accounts',
  'SharePoint.Files', 'SharePoint.Contents', 'SharePoint.Tables',
  'AzureStorage.Blobs', 'AzureStorage.Tables', 'AzureStorage.DataLake',
  'AzureEnterprise.Contents',
  'AnalysisServices.Database', 'AnalysisServices.Databases',
  'Cube.AttributeMemberset',
  'Exchange.Contents', 'ActiveDirectory.Domains',
  'Facebook.Graph', 'AdobeAnalytics.Cubes',
  'Access.Database', 'Pdf.Tables',
  'Hadoop.Containers', 'HdInsight.Containers',
  'Databricks.Catalogs',
]);

const CONNECTOR_RE = new RegExp(
  '\\b(' + Array.from(KNOWN_CONNECTORS).map(c => c.replace('.', '\\.')).join('|') + ')\\b',
  'g'
);

const BARE_BLOCKLIST = new Set([
  'Source', 'Result', 'Table', 'Output', 'Data', 'Custom',
  'List', 'Record', 'Type', 'Text', 'Number', 'Date', 'Time',
  'Duration', 'Logical', 'Binary', 'Error', 'Action', 'Function',
  'None', 'Step', 'Value', 'Row', 'Rows', 'Column', 'Columns',
  'Query', 'Name', 'Index', 'Count', 'Sum', 'Min', 'Max',
]);

// ---------------------------------------------------------------------------
// Section1.m parser
// ---------------------------------------------------------------------------

const SECTION_HEADER_RE = /^\s*section\s+Section1\s*;/i;
const SHARED_DECL_RE = /(?:^|\n)\s*shared\s+(#".*?"|\w+)\s*=\s*/gs;

function splitSectionM(text) {
  text = text.replace(SECTION_HEADER_RE, '').trim();
  const matches = [];
  let m;
  // Reset lastIndex
  SHARED_DECL_RE.lastIndex = 0;
  while ((m = SHARED_DECL_RE.exec(text)) !== null) {
    matches.push({ name: m[1].trim(), start: m.index, end: m.index + m[0].length });
  }
  if (!matches.length) return [];

  const items = [];
  for (let i = 0; i < matches.length; i++) {
    let name = matches[i].name;
    if (name.startsWith('#"') && name.endsWith('"')) {
      name = name.slice(2, -1);
    }
    const bodyStart = matches[i].end;
    const bodyEnd = i + 1 < matches.length ? matches[i + 1].start : text.length;
    let body = text.slice(bodyStart, bodyEnd).trim();
    if (body.endsWith(';')) body = body.slice(0, -1).trimEnd();
    items.push({ name, body });
  }
  return items;
}

function classifyM(name, body) {
  if (/IsParameterQuery\s*=\s*true/i.test(body)) return MItemKind.PARAMETER;
  const stripped = body.trimStart();
  if (stripped.startsWith('(') && body.includes('=>')) return MItemKind.FUNCTION;
  if (/\(\s*\w+.*?\)\s*=>/.test(body)) return MItemKind.FUNCTION;
  return MItemKind.QUERY;
}

// ---------------------------------------------------------------------------
// DataMashup (MS-QDEFF) binary parser
// ---------------------------------------------------------------------------

function parseDataMashup(buffer) {
  const view = new DataView(buffer);
  let offset = 0;

  const version = view.getUint32(offset, true);
  offset += 4;
  if (version !== 0) throw new Error(`Unexpected DataMashup version: ${version}`);

  const pkgLen = view.getUint32(offset, true);
  offset += 4;
  const pkgData = buffer.slice(offset, offset + pkgLen);
  offset += pkgLen;

  const permLen = view.getUint32(offset, true);
  offset += 4;
  offset += permLen;

  const metaLen = view.getUint32(offset, true);
  offset += 4;
  const metaData = buffer.slice(offset, offset + metaLen);
  offset += metaLen;

  const metadataXml = new TextDecoder('utf-8').decode(metaData);
  return { pkgData, metadataXml };
}

function parseMetadataXml(xmlText) {
  if (!xmlText.trim()) return {};
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, 'text/xml');
  const NS = 'http://schemas.microsoft.com/DataMashup';
  const result = {};
  const queryGroups = {};

  const items = doc.getElementsByTagNameNS(NS, 'Item');
  for (const item of items) {
    const loc = item.getElementsByTagNameNS(NS, 'ItemLocation')[0];
    if (!loc) continue;
    const itemTypeEl = loc.getElementsByTagNameNS(NS, 'ItemType')[0];
    const itemPathEl = loc.getElementsByTagNameNS(NS, 'ItemPath')[0];
    if (!itemTypeEl || !itemPathEl) continue;

    const itemType = itemTypeEl.textContent || '';
    const itemPath = itemPathEl.textContent || '';
    const entries = {};

    const entryEls = item.getElementsByTagNameNS(NS, 'Entry');
    for (const entry of entryEls) {
      const etype = entry.getAttribute('Type') || '';
      let evalue = entry.getAttribute('Value') || '';
      if (evalue.length > 1 && 'slfcd'.includes(evalue[0])) {
        evalue = evalue.slice(1);
      }
      entries[etype] = evalue;
    }

    if (itemType === 'AllFormulas' && entries.QueryGroups) {
      try {
        const groups = JSON.parse(entries.QueryGroups);
        for (const g of groups) {
          if (g.Id && g.Name) queryGroups[g.Id] = g.Name;
        }
      } catch (e) { /* ignore */ }
    } else if (itemPath) {
      const cleanPath = itemPath.includes('/') ? itemPath.split('/').slice(1).join('/') : itemPath;
      result[cleanPath] = entries;
    }
  }

  // Resolve group names
  for (const [path, entries] of Object.entries(result)) {
    const gid = entries.QueryGroupID || '';
    if (gid && queryGroups[gid]) {
      entries.QueryGroup = queryGroups[gid];
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// Extraction routes
// ---------------------------------------------------------------------------

async function extractFromDataMashup(arrayBuffer) {
  const outerZip = await JSZip.loadAsync(arrayBuffer);
  const mashupFile = outerZip.file('DataMashup');
  if (!mashupFile) throw new Error('No DataMashup found in archive');
  const mashupBuf = await mashupFile.async('arraybuffer');

  const { pkgData, metadataXml } = parseDataMashup(mashupBuf);
  const innerZip = await JSZip.loadAsync(pkgData);

  let section1m = '';
  for (const [name, file] of Object.entries(innerZip.files)) {
    if (name.toLowerCase().endsWith('.m')) {
      section1m = await file.async('string');
      break;
    }
  }
  if (!section1m) throw new Error('No M code found in DataMashup');

  const metaMap = parseMetadataXml(metadataXml);
  const rawItems = splitSectionM(section1m);

  return rawItems.map(({ name, body }) => {
    const kind = classifyM(name, body);
    const metadata = {};
    const itemMeta = metaMap[name] || {};
    if (itemMeta.QueryGroup) metadata.group = itemMeta.QueryGroup;
    if (itemMeta.AddedToDataModel === '1') metadata.load_enabled = 'true';
    else if ('AddedToDataModel' in itemMeta) metadata.load_enabled = 'false';
    if (itemMeta.IsPrivate === '1') metadata.is_private = 'true';
    return new MItem(name, kind, body, metadata);
  });
}

async function extractFromPbitJson(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const schemaFile = zip.file('DataModelSchema');
  if (!schemaFile) throw new Error('No DataModelSchema found in archive');

  // DataModelSchema is UTF-16 LE with BOM
  const rawBuf = await schemaFile.async('arraybuffer');
  let text;
  try {
    text = new TextDecoder('utf-16le').decode(rawBuf);
  } catch (e) {
    text = new TextDecoder('utf-8').decode(rawBuf);
  }
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);

  const schema = JSON.parse(text);
  const model = schema.model || {};
  const items = [];

  // Shared expressions
  for (const exprObj of (model.expressions || [])) {
    const name = exprObj.name || 'Unknown';
    let expr = exprObj.expression || '';
    if (Array.isArray(expr)) expr = expr.join('\n');
    const kind = classifyM(name, expr);
    const metadata = {};
    if (exprObj.description) metadata.description = exprObj.description;
    items.push(new MItem(name, kind, expr, metadata));
  }

  // Table partitions
  for (const table of (model.tables || [])) {
    for (const partition of (table.partitions || [])) {
      const source = partition.source || {};
      if (source.type !== 'm') continue;
      let expr = source.expression || '';
      if (Array.isArray(expr)) expr = expr.join('\n');
      if (!expr.trim()) continue;
      const kind = classifyM(table.name, expr);
      const metadata = { load_enabled: 'true' };
      if (table.description) metadata.description = table.description;
      items.push(new MItem(table.name, kind, expr, metadata));
    }
  }

  return items;
}

// ---------------------------------------------------------------------------
// Dependency analysis
// ---------------------------------------------------------------------------

function analyzeDependencies(items) {
  const names = new Set(items.map(i => i.name));
  const nameToKind = {};
  items.forEach(i => { nameToKind[i.name] = i.kind; });

  // Pre-compile patterns
  const patterns = {};
  for (const name of names) {
    const quoted = '#"' + name + '"';
    let bareRe = null;
    if (/^[A-Za-z_]\w+$/.test(name) && name.length >= 3 && !BARE_BLOCKLIST.has(name)) {
      bareRe = new RegExp('\\b' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b');
    }
    patterns[name] = { quoted, bareRe };
  }

  const edges = [];
  const dataSources = {};

  for (const item of items) {
    const code = item.mCode;
    const connectors = [];

    for (const otherName of names) {
      if (otherName === item.name) continue;
      const { quoted, bareRe } = patterns[otherName];
      if (code.includes(quoted)) {
        edges.push({ from: item.name, to: otherName });
      } else if (bareRe && bareRe.test(code)) {
        edges.push({ from: item.name, to: otherName });
      }
    }

    CONNECTOR_RE.lastIndex = 0;
    let cm;
    while ((cm = CONNECTOR_RE.exec(code)) !== null) {
      connectors.push(cm[1]);
    }
    if (connectors.length) {
      dataSources[item.name] = [...new Set(connectors)].sort();
    }
  }

  // Deduplicate edges
  const edgeSet = new Set(edges.map(e => e.from + '\0' + e.to));
  const uniqueEdges = [...edgeSet].map(s => {
    const [from, to] = s.split('\0');
    return { from, to };
  }).sort((a, b) => a.from.localeCompare(b.from) || a.to.localeCompare(b.to));

  const nodes = items.map(i => ({ name: i.name, kind: i.kind }));
  const layers = computeLayers(names, uniqueEdges);

  return new DependencyGraph(uniqueEdges, dataSources, nodes, layers);
}

function computeLayers(names, edges) {
  const dependents = {};  // who depends on this node
  const dependencies = {}; // what this node depends on

  for (const { from, to } of edges) {
    if (!dependencies[from]) dependencies[from] = new Set();
    dependencies[from].add(to);
    if (!dependents[to]) dependents[to] = new Set();
    dependents[to].add(from);
  }

  const inDegree = {};
  for (const n of names) {
    inDegree[n] = dependencies[n] ? dependencies[n].size : 0;
  }

  let queue = [...names].filter(n => inDegree[n] === 0).sort();
  const layers = [];
  const visited = new Set();

  while (queue.length) {
    // Barycenter ordering
    let layer;
    if (layers.length > 0) {
      const prevPositions = {};
      layers[layers.length - 1].forEach((n, i) => { prevPositions[n] = i; });
      layer = queue.slice().sort((a, b) => {
        const aUps = dependencies[a] || new Set();
        const bUps = dependencies[b] || new Set();
        const aPositions = [...aUps].filter(u => u in prevPositions).map(u => prevPositions[u]);
        const bPositions = [...bUps].filter(u => u in prevPositions).map(u => prevPositions[u]);
        const aAvg = aPositions.length ? aPositions.reduce((s, v) => s + v, 0) / aPositions.length : 0;
        const bAvg = bPositions.length ? bPositions.reduce((s, v) => s + v, 0) / bPositions.length : 0;
        return aAvg - bAvg || a.localeCompare(b);
      });
    } else {
      layer = queue.slice().sort();
    }

    layers.push(layer);
    const nextQueue = [];
    for (const node of layer) {
      visited.add(node);
      for (const dep of (dependents[node] || new Set())) {
        inDegree[dep]--;
        if (inDegree[dep] === 0) nextQueue.push(dep);
      }
    }
    queue = nextQueue;
  }

  // Remaining (cycles)
  const remaining = [...names].filter(n => !visited.has(n)).sort();
  if (remaining.length) layers.push(remaining);

  return layers;
}

// ---------------------------------------------------------------------------
// Markdown renderer
// ---------------------------------------------------------------------------

function renderMarkdown(items, sourceFile, includeMetadata, depGraph) {
  const groups = {
    [MItemKind.PARAMETER]: [],
    [MItemKind.FUNCTION]: [],
    [MItemKind.QUERY]: [],
  };
  items.forEach(item => groups[item.kind].push(item));
  Object.values(groups).forEach(lst => lst.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase())));

  // Dependency lookups
  const depsOf = {};
  const usedBy = {};
  if (depGraph) {
    depGraph.edges.forEach(({ from, to }) => {
      if (!depsOf[from]) depsOf[from] = [];
      depsOf[from].push(to);
      if (!usedBy[to]) usedBy[to] = [];
      usedBy[to].push(from);
    });
  }

  const lines = [];
  lines.push(`# Power Query M \u2014 ${sourceFile}`);
  lines.push('');

  const total = items.length;
  lines.push(`> **${total}** items extracted (${groups[MItemKind.PARAMETER].length} parameters, ${groups[MItemKind.QUERY].length} queries, ${groups[MItemKind.FUNCTION].length} functions)`);
  lines.push('');

  const sectionOrder = [
    [MItemKind.PARAMETER, 'Parameters'],
    [MItemKind.QUERY, 'Queries'],
    [MItemKind.FUNCTION, 'Functions'],
  ];

  // TOC
  lines.push('## Table of Contents');
  lines.push('');
  if (depGraph) lines.push('- [Dependency Summary](#dependency-summary)');
  for (const [kind, heading] of sectionOrder) {
    if (!groups[kind].length) continue;
    lines.push(`- [${heading}](#${heading.toLowerCase()})`);
    for (const item of groups[kind]) {
      const anchor = item.name.toLowerCase().replace(/[^a-z0-9\- ]/g, '').replace(/ /g, '-');
      lines.push(`  - [${item.name}](#${anchor})`);
    }
  }
  lines.push('');

  // Dependency summary
  if (depGraph) {
    lines.push('## Dependency Summary');
    lines.push('');
    lines.push(`> **${depGraph.edges.length}** dependencies across **${depGraph.layers.length}** layers`);
    lines.push('');

    if (Object.keys(depGraph.dataSources).length) {
      lines.push('### Data Sources');
      lines.push('');
      for (const qname of Object.keys(depGraph.dataSources).sort()) {
        const connectors = depGraph.dataSources[qname].map(c => '`' + c + '`').join(', ');
        lines.push(`- **${qname}**: ${connectors}`);
      }
      lines.push('');
    }

    lines.push('### Query Layers');
    lines.push('');
    depGraph.layers.forEach((layer, i) => {
      const label = i === 0 ? 'Sources' : `Layer ${i}`;
      lines.push(`- **${label}**: ${layer.join(', ')}`);
    });
    lines.push('');
  }

  // Sections
  for (const [kind, heading] of sectionOrder) {
    if (!groups[kind].length) continue;
    lines.push(`## ${heading}`);
    lines.push('');

    for (const item of groups[kind]) {
      lines.push(`### ${item.name}`);
      lines.push('');

      if (includeMetadata && Object.keys(item.metadata).length) {
        for (const [k, v] of Object.entries(item.metadata).sort()) {
          lines.push(`- **${k}**: ${v}`);
        }
        lines.push('');
      }

      if (depGraph) {
        const depLines = [];
        if (depsOf[item.name]?.length) depLines.push('- **depends on**: ' + depsOf[item.name].sort().join(', '));
        if (usedBy[item.name]?.length) depLines.push('- **used by**: ' + usedBy[item.name].sort().join(', '));
        if (depGraph.dataSources[item.name]) {
          depLines.push('- **data source**: ' + depGraph.dataSources[item.name].map(c => '`' + c + '`').join(', '));
        }
        if (depLines.length) {
          lines.push(...depLines);
          lines.push('');
        }
      }

      lines.push('```m');
      lines.push(item.mCode);
      lines.push('```');
      lines.push('');
    }
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

async function extractFile(arrayBuffer, fileName, mode, includeMetadata) {
  if (mode === 'auto') {
    mode = fileName.toLowerCase().endsWith('.pbit') ? 'pbit-json' : 'pbix-datamashup';
  }

  let items;
  if (mode === 'pbit-json') {
    items = await extractFromPbitJson(arrayBuffer);
  } else {
    items = await extractFromDataMashup(arrayBuffer);
  }

  const depGraph = analyzeDependencies(items);
  const markdown = renderMarkdown(items, fileName, includeMetadata, depGraph);

  const stats = {
    parameters: items.filter(i => i.kind === MItemKind.PARAMETER).length,
    queries: items.filter(i => i.kind === MItemKind.QUERY).length,
    functions: items.filter(i => i.kind === MItemKind.FUNCTION).length,
    total: items.length,
  };

  return { items, markdown, stats, graph: depGraph };
}

// Make available globally
window.MExporter = { extractFile, MItemKind };
