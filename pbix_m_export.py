#!/usr/bin/env python3
"""
Power BI M Exporter — extracts Power Query M code from .pbix / .pbit files
and outputs a single Markdown document.

Requires only Python standard library (3.7+).
"""

from __future__ import annotations

import argparse
import io
import json
import re
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict, deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class MItemKind:
    QUERY = "Query"
    PARAMETER = "Parameter"
    FUNCTION = "Function"


@dataclass
class MItem:
    name: str
    kind: str  # one of MItemKind
    m_code: str
    metadata: Dict[str, str] = field(default_factory=dict)


@dataclass
class DependencyGraph:
    edges: List[Tuple[str, str]]          # (from_name, to_name) — "from depends on to"
    data_sources: Dict[str, List[str]]    # {query_name: [connector, ...]}
    nodes: List[Dict[str, str]]           # [{name, kind}, ...] for rendering
    layers: List[List[str]]               # topological layers for layout


# ---------------------------------------------------------------------------
# Dependency analysis
# ---------------------------------------------------------------------------

KNOWN_CONNECTORS: Set[str] = {
    "Sql.Database", "Sql.Databases", "Oracle.Database",
    "Odbc.DataSource", "Odbc.Query", "OleDb.DataSource", "OleDb.Query",
    "PostgreSQL.Database", "MySQL.Database", "Snowflake.Databases",
    "Excel.Workbook", "Excel.CurrentWorkbook",
    "Csv.Document", "Json.Document", "Xml.Document", "Xml.Tables",
    "File.Contents", "Folder.Files", "Folder.Contents",
    "Web.Contents", "Web.Page", "Web.BrowserContents",
    "OData.Feed",
    "Salesforce.Data", "Salesforce.Reports",
    "GoogleAnalytics.Accounts",
    "SharePoint.Files", "SharePoint.Contents", "SharePoint.Tables",
    "AzureStorage.Blobs", "AzureStorage.Tables", "AzureStorage.DataLake",
    "AzureEnterprise.Contents",
    "AnalysisServices.Database", "AnalysisServices.Databases",
    "Cube.AttributeMemberset",
    "Exchange.Contents", "ActiveDirectory.Domains",
    "Facebook.Graph", "AdobeAnalytics.Cubes",
    "Access.Database", "Pdf.Tables",
    "Hadoop.Containers", "HdInsight.Containers",
    "Databricks.Catalogs",
}

_CONNECTOR_RE = re.compile(
    r"\b(" + "|".join(re.escape(c) for c in sorted(KNOWN_CONNECTORS)) + r")\b"
)

# Common M variable names that cause false-positive matches
_BARE_BLOCKLIST: Set[str] = {
    "Source", "Result", "Table", "Output", "Data", "Custom",
    "List", "Record", "Type", "Text", "Number", "Date", "Time",
    "Duration", "Logical", "Binary", "Error", "Action", "Function",
    "None", "Step", "Value", "Row", "Rows", "Column", "Columns",
    "Query", "Name", "Index", "Count", "Sum", "Min", "Max",
}


def analyze_dependencies(items: List[MItem]) -> DependencyGraph:
    """Analyze inter-query references and data source connectors."""
    names = {item.name for item in items}
    name_to_kind = {item.name: item.kind for item in items}

    # Pre-compile match patterns per query name
    patterns: Dict[str, Tuple[str, Optional[re.Pattern]]] = {}
    for name in names:
        quoted = '#"' + name + '"'
        # Only do bare-identifier matching for safe names
        if (
            re.match(r"^[A-Za-z_]\w+$", name)
            and len(name) >= 3
            and name not in _BARE_BLOCKLIST
        ):
            bare_re = re.compile(r"\b" + re.escape(name) + r"\b")
        else:
            bare_re = None
        patterns[name] = (quoted, bare_re)

    edges: List[Tuple[str, str]] = []
    data_sources: Dict[str, List[str]] = {}

    for item in items:
        code = item.m_code
        connectors_found: List[str] = []

        # Find references to other queries
        for other_name in names:
            if other_name == item.name:
                continue
            quoted, bare_re = patterns[other_name]
            if quoted in code:
                edges.append((item.name, other_name))
            elif bare_re and bare_re.search(code):
                edges.append((item.name, other_name))

        # Find data source connectors
        for match in _CONNECTOR_RE.findall(code):
            connectors_found.append(match)
        if connectors_found:
            data_sources[item.name] = sorted(set(connectors_found))

    # Deduplicate edges
    edges = sorted(set(edges))

    nodes = [{"name": item.name, "kind": item.kind} for item in items]
    layers = _compute_layers(names, edges)

    return DependencyGraph(
        edges=edges,
        data_sources=data_sources,
        nodes=nodes,
        layers=layers,
    )


def _compute_layers(
    names: Set[str], edges: List[Tuple[str, str]]
) -> List[List[str]]:
    """Assign nodes to layers via topological sort (Kahn's algorithm).

    Layer 0 = root sources (no dependencies), last layer = final consumers.
    """
    dependents: Dict[str, Set[str]] = defaultdict(set)   # who depends on this node
    dependencies: Dict[str, Set[str]] = defaultdict(set)  # what this node depends on

    for a, b in edges:
        dependencies[a].add(b)
        dependents[b].add(a)

    in_degree = {n: len(dependencies.get(n, set())) for n in names}
    queue = deque(sorted(n for n in names if in_degree[n] == 0))

    layers: List[List[str]] = []
    visited: Set[str] = set()

    while queue:
        # Barycenter ordering: sort by average position of upstream neighbors
        # in the previous layer (falls back to alphabetical for the first layer)
        if layers:
            prev_positions = {
                name: i for i, name in enumerate(layers[-1])
            }

            def _bary(node: str) -> float:
                ups = dependencies.get(node, set())
                positions = [prev_positions[u] for u in ups if u in prev_positions]
                return sum(positions) / len(positions) if positions else 0.0

            layer = sorted(queue, key=lambda n: (_bary(n), n))
        else:
            layer = sorted(queue)

        layers.append(layer)
        next_queue: deque = deque()
        for node in layer:
            visited.add(node)
            for dep in dependents.get(node, set()):
                in_degree[dep] -= 1
                if in_degree[dep] == 0:
                    next_queue.append(dep)
        queue = next_queue

    # Remaining nodes (cycles) go in a final layer
    remaining = sorted(names - visited)
    if remaining:
        layers.append(remaining)

    return layers


def dependency_graph_to_dict(graph: DependencyGraph) -> dict:
    """Serialize a DependencyGraph to a JSON-compatible dict."""
    return {
        "edges": [{"from": a, "to": b} for a, b in graph.edges],
        "data_sources": graph.data_sources,
        "nodes": graph.nodes,
        "layers": graph.layers,
    }


# ---------------------------------------------------------------------------
# Section1.m parser
# ---------------------------------------------------------------------------

_SECTION_HEADER = re.compile(r"^\s*section\s+Section1\s*;", re.IGNORECASE)

# Match: shared <Name> = <body>
# We need to handle nested let/in, strings, etc.  The safest stdlib-only
# approach is to split on `shared <Name> =` boundaries.
_SHARED_DECL = re.compile(r'(?:^|\n)\s*shared\s+(#".*?"|\w+)\s*=\s*', re.DOTALL)


def _split_section_m(text: str) -> List[Tuple[str, str]]:
    """Split a Section1.m document into (name, body) pairs."""
    # Strip the section header
    text = _SECTION_HEADER.sub("", text, count=1).strip()

    # Find all `shared <Name> = ` positions
    matches = list(_SHARED_DECL.finditer(text))
    if not matches:
        return []

    items: List[Tuple[str, str]] = []
    for i, m in enumerate(matches):
        name = m.group(1).strip()
        # Remove surrounding quotes from #"quoted names"
        if name.startswith('#"') and name.endswith('"'):
            name = name[2:-1]
        body_start = m.end()
        if i + 1 < len(matches):
            body_end = matches[i + 1].start()
        else:
            body_end = len(text)
        body = text[body_start:body_end].strip()
        # Remove trailing semicolon that terminates the member
        if body.endswith(";"):
            body = body[:-1].rstrip()
        items.append((name, body))
    return items


def _classify_m(name: str, body: str) -> str:
    """Classify an M expression as Parameter, Function, or Query."""
    # Parameters have meta [IsParameterQuery=true, ...]
    if re.search(r"IsParameterQuery\s*=\s*true", body, re.IGNORECASE):
        return MItemKind.PARAMETER
    # Functions: body starts with (args) => or let ... (args) => pattern
    # Simplification: look for a top-level `=>`
    stripped = body.lstrip()
    if stripped.startswith("(") and "=>" in body:
        return MItemKind.FUNCTION
    # Also match let ... Source = (param) => ... patterns often used for functions
    if re.search(r"\(\s*\w+.*?\)\s*=>", body):
        return MItemKind.FUNCTION
    return MItemKind.QUERY


# ---------------------------------------------------------------------------
# DataMashup (MS-QDEFF) binary parser
# ---------------------------------------------------------------------------

def _parse_datamashup(raw: bytes) -> Tuple[str, str]:
    """Parse the MS-QDEFF binary stream.

    Returns (section1_m_text, metadata_xml_text).
    """
    offset = 0

    # Version (uint32, must be 0)
    version = struct.unpack_from("<I", raw, offset)[0]
    offset += 4
    if version != 0:
        raise ValueError(f"Unexpected DataMashup version: {version}")

    # Package Parts (OPC ZIP)
    pkg_len = struct.unpack_from("<I", raw, offset)[0]
    offset += 4
    pkg_data = raw[offset : offset + pkg_len]
    offset += pkg_len

    # Permissions
    perm_len = struct.unpack_from("<I", raw, offset)[0]
    offset += 4
    offset += perm_len

    # Metadata XML
    meta_len = struct.unpack_from("<I", raw, offset)[0]
    offset += 4
    meta_data = raw[offset : offset + meta_len]
    offset += meta_len

    # Extract Section1.m from embedded ZIP
    section1_m = ""
    with zipfile.ZipFile(io.BytesIO(pkg_data)) as pkg:
        for name in pkg.namelist():
            if name.lower().endswith(".m"):
                section1_m = pkg.read(name).decode("utf-8")
                break

    metadata_xml = meta_data.decode("utf-8") if meta_data else ""
    return section1_m, metadata_xml


def _parse_metadata_xml(xml_text: str) -> Dict[str, Dict[str, str]]:
    """Parse the DataMashup metadata XML.

    Returns {item_path: {entry_type: value, ...}}.
    """
    if not xml_text.strip():
        return {}

    ns = {"dm": "http://schemas.microsoft.com/DataMashup"}
    result: Dict[str, Dict[str, str]] = {}
    # Query groups mapping
    query_groups: Dict[str, str] = {}

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return {}

    for item_el in root.findall(".//dm:Item", ns):
        loc = item_el.find("dm:ItemLocation", ns)
        if loc is None:
            continue
        item_type_el = loc.find("dm:ItemType", ns)
        item_path_el = loc.find("dm:ItemPath", ns)
        if item_type_el is None or item_path_el is None:
            continue

        item_path = item_path_el.text or ""
        entries: Dict[str, str] = {}
        for entry in item_el.findall(".//dm:Entry", ns):
            etype = entry.get("Type", "")
            evalue = entry.get("Value", "")
            # Strip type prefix (s, l, f, c, d)
            if evalue and len(evalue) > 1 and evalue[0] in "slfcd":
                evalue = evalue[1:]
            entries[etype] = evalue

        if item_type_el.text == "AllFormulas" and "QueryGroups" in entries:
            # Parse query groups JSON
            try:
                groups_data = json.loads(entries["QueryGroups"])
                for group in groups_data:
                    gid = group.get("Id", "")
                    gname = group.get("Name", "")
                    if gid and gname:
                        query_groups[gid] = gname
            except (json.JSONDecodeError, TypeError):
                pass
        elif item_path:
            # Strip Section1/ prefix
            clean_path = item_path
            if "/" in clean_path:
                clean_path = clean_path.split("/", 1)[1]
            result[clean_path] = entries

    # Resolve group names
    for path, entries in result.items():
        gid = entries.get("QueryGroupID", "")
        if gid and gid in query_groups:
            entries["QueryGroup"] = query_groups[gid]

    return result


# ---------------------------------------------------------------------------
# Extraction routes
# ---------------------------------------------------------------------------

def extract_from_datamashup(pbix_path: Path) -> List[MItem]:
    """Extract M items from a .pbix or .pbit file via its DataMashup blob."""
    with zipfile.ZipFile(pbix_path, "r") as z:
        raw = z.read("DataMashup")

    section1_m, metadata_xml = _parse_datamashup(raw)
    if not section1_m:
        raise ValueError("No M code found in DataMashup")

    meta_map = _parse_metadata_xml(metadata_xml)
    raw_items = _split_section_m(section1_m)

    items: List[MItem] = []
    for name, body in raw_items:
        kind = _classify_m(name, body)
        metadata: Dict[str, str] = {}

        item_meta = meta_map.get(name, {})
        if item_meta.get("QueryGroup"):
            metadata["group"] = item_meta["QueryGroup"]
        if item_meta.get("AddedToDataModel") == "1":
            metadata["load_enabled"] = "true"
        elif "AddedToDataModel" in item_meta:
            metadata["load_enabled"] = "false"
        if item_meta.get("IsPrivate") == "1":
            metadata["is_private"] = "true"

        items.append(MItem(name=name, kind=kind, m_code=body, metadata=metadata))

    return items


def extract_from_pbit_json(pbit_path: Path) -> List[MItem]:
    """Extract M items from a .pbit file via its DataModelSchema JSON."""
    with zipfile.ZipFile(pbit_path, "r") as z:
        raw = z.read("DataModelSchema")

    # DataModelSchema is UTF-16 LE with BOM
    try:
        text = raw.decode("utf-16-le")
    except UnicodeDecodeError:
        text = raw.decode("utf-8")

    # Strip BOM if present
    if text and text[0] == "\ufeff":
        text = text[1:]

    schema = json.loads(text)
    model = schema.get("model", {})

    items: List[MItem] = []

    # 1. Shared expressions (parameters, functions, connection-only tables)
    for expr_obj in model.get("expressions", []):
        name = expr_obj.get("name", "Unknown")
        expr = expr_obj.get("expression", "")
        if isinstance(expr, list):
            expr = "\n".join(expr)

        kind = _classify_m(name, expr)
        metadata: Dict[str, str] = {}

        desc = expr_obj.get("description")
        if desc:
            metadata["description"] = desc

        items.append(MItem(name=name, kind=kind, m_code=expr, metadata=metadata))

    # 2. Table partitions with M source
    for table in model.get("tables", []):
        for partition in table.get("partitions", []):
            source = partition.get("source", {})
            if source.get("type") != "m":
                continue
            expr = source.get("expression", "")
            if isinstance(expr, list):
                expr = "\n".join(expr)
            if not expr.strip():
                continue

            name = table["name"]
            kind = _classify_m(name, expr)
            metadata = {"load_enabled": "true"}

            desc = table.get("description")
            if desc:
                metadata["description"] = desc

            items.append(MItem(name=name, kind=kind, m_code=expr, metadata=metadata))

    return items


# ---------------------------------------------------------------------------
# Markdown renderer
# ---------------------------------------------------------------------------

def render_markdown(
    items: List[MItem],
    source_file: str,
    include_metadata: bool,
    dep_graph: Optional[DependencyGraph] = None,
) -> str:
    """Render a list of MItems to a Markdown string."""
    groups = {
        MItemKind.PARAMETER: [],
        MItemKind.FUNCTION: [],
        MItemKind.QUERY: [],
    }
    for item in items:
        groups[item.kind].append(item)

    # Sort each group alphabetically
    for lst in groups.values():
        lst.sort(key=lambda x: x.name.lower())

    # Pre-compute dependency lookups
    deps_of: Dict[str, List[str]] = defaultdict(list)   # what X depends on
    used_by: Dict[str, List[str]] = defaultdict(list)    # what uses X
    if dep_graph:
        for a, b in dep_graph.edges:
            deps_of[a].append(b)
            used_by[b].append(a)

    lines: List[str] = []
    lines.append(f"# Power Query M — {source_file}")
    lines.append("")

    # Table of contents
    total = len(items)
    lines.append(f"> **{total}** items extracted "
                 f"({len(groups[MItemKind.PARAMETER])} parameters, "
                 f"{len(groups[MItemKind.QUERY])} queries, "
                 f"{len(groups[MItemKind.FUNCTION])} functions)")
    lines.append("")

    section_order = [
        (MItemKind.PARAMETER, "Parameters"),
        (MItemKind.QUERY, "Queries"),
        (MItemKind.FUNCTION, "Functions"),
    ]

    # TOC
    lines.append("## Table of Contents")
    lines.append("")
    if dep_graph:
        lines.append("- [Dependency Summary](#dependency-summary)")
    for kind, heading in section_order:
        if not groups[kind]:
            continue
        lines.append(f"- [{heading}](#{heading.lower()})")
        for item in groups[kind]:
            anchor = re.sub(r"[^a-z0-9\- ]", "", item.name.lower()).replace(" ", "-")
            lines.append(f"  - [{item.name}](#{anchor})")
    lines.append("")

    # Dependency summary section
    if dep_graph:
        lines.append("## Dependency Summary")
        lines.append("")
        lines.append(f"> **{len(dep_graph.edges)}** dependencies across "
                     f"**{len(dep_graph.layers)}** layers")
        lines.append("")

        # Data sources
        if dep_graph.data_sources:
            lines.append("### Data Sources")
            lines.append("")
            for qname in sorted(dep_graph.data_sources):
                connectors = ", ".join(f"`{c}`" for c in dep_graph.data_sources[qname])
                lines.append(f"- **{qname}**: {connectors}")
            lines.append("")

        # Layer overview
        lines.append("### Query Layers")
        lines.append("")
        for i, layer in enumerate(dep_graph.layers):
            label = "Sources" if i == 0 else f"Layer {i}"
            lines.append(f"- **{label}**: {', '.join(layer)}")
        lines.append("")

    # Sections
    for kind, heading in section_order:
        if not groups[kind]:
            continue
        lines.append(f"## {heading}")
        lines.append("")

        for item in groups[kind]:
            lines.append(f"### {item.name}")
            lines.append("")

            if include_metadata and item.metadata:
                for k, v in sorted(item.metadata.items()):
                    lines.append(f"- **{k}**: {v}")
                lines.append("")

            # Dependency info
            if dep_graph:
                dep_lines = []
                if deps_of[item.name]:
                    dep_lines.append("- **depends on**: " + ", ".join(sorted(deps_of[item.name])))
                if used_by[item.name]:
                    dep_lines.append("- **used by**: " + ", ".join(sorted(used_by[item.name])))
                if item.name in dep_graph.data_sources:
                    connectors = ", ".join(f"`{c}`" for c in dep_graph.data_sources[item.name])
                    dep_lines.append(f"- **data source**: {connectors}")
                if dep_lines:
                    lines.extend(dep_lines)
                    lines.append("")

            lines.append("```m")
            lines.append(item.m_code)
            lines.append("```")
            lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def detect_mode(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".pbit":
        return "pbit-json"
    return "pbix-datamashup"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract Power Query M code from .pbix/.pbit files to Markdown.",
    )
    parser.add_argument("input", type=Path, help="Path to .pbix or .pbit file")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output Markdown file path (default: <input>.md)",
    )
    parser.add_argument(
        "--mode",
        choices=["auto", "pbix-datamashup", "pbit-json"],
        default="auto",
        help="Extraction mode (default: auto-detect from file extension)",
    )
    parser.add_argument(
        "--metadata", action="store_true", default=False,
        help="Include metadata (group, load status, etc.) in output",
    )

    args = parser.parse_args(argv)
    input_path: Path = args.input
    output_path: Path = args.output or input_path.with_suffix(".md")
    mode: str = args.mode
    include_metadata: bool = args.metadata

    if not input_path.exists():
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        return 1

    if mode == "auto":
        mode = detect_mode(input_path)

    try:
        if mode == "pbit-json":
            items = extract_from_pbit_json(input_path)
        else:
            items = extract_from_datamashup(input_path)
    except KeyError as exc:
        print(f"Error: expected entry not found in archive: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error extracting M code: {exc}", file=sys.stderr)
        return 1

    if not items:
        print("Warning: no M items found.", file=sys.stderr)
        return 0

    dep_graph = analyze_dependencies(items)
    md = render_markdown(items, input_path.name, include_metadata, dep_graph)
    output_path.write_text(md, encoding="utf-8")
    print(f"Wrote {len(items)} items to {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
