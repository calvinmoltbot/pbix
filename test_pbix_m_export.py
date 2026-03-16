#!/usr/bin/env python3
"""Tests for pbix_m_export using synthetic .pbix and .pbit files."""

import io
import json
import struct
import tempfile
import zipfile
from pathlib import Path

from pbix_m_export import (
    MItem,
    MItemKind,
    DependencyGraph,
    _classify_m,
    _split_section_m,
    _parse_datamashup,
    _parse_metadata_xml,
    _compute_layers,
    analyze_dependencies,
    dependency_graph_to_dict,
    extract_from_datamashup,
    extract_from_pbit_json,
    render_markdown,
    main,
)


# ---------------------------------------------------------------------------
# Helpers to build synthetic files
# ---------------------------------------------------------------------------

SAMPLE_SECTION1_M = '''\
section Section1;

shared ServerName = "myserver.database.windows.net" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared #"My Query" = let
    Source = Sql.Database(ServerName, "mydb"),
    dbo_Sales = Source{[Schema="dbo",Item="Sales"]}[Data]
in
    dbo_Sales;

shared fnTransform = (tbl as table) =>
    let
        Result = Table.AddColumn(tbl, "New", each [A] + [B])
    in
        Result;
'''

SAMPLE_METADATA_XML = '''\
<?xml version="1.0" encoding="utf-8"?>
<LocalPackageMetadataFile xmlns="http://schemas.microsoft.com/DataMashup">
  <Items>
    <Item>
      <ItemLocation>
        <ItemType>Formula</ItemType>
        <ItemPath>Section1/ServerName</ItemPath>
      </ItemLocation>
      <StableEntries>
        <Entry Type="IsPrivate" Value="l0"/>
        <Entry Type="ResultType" Value="sText"/>
      </StableEntries>
    </Item>
    <Item>
      <ItemLocation>
        <ItemType>Formula</ItemType>
        <ItemPath>Section1/My Query</ItemPath>
      </ItemLocation>
      <StableEntries>
        <Entry Type="AddedToDataModel" Value="l1"/>
        <Entry Type="QueryGroupID" Value="sgroup-1"/>
      </StableEntries>
    </Item>
    <Item>
      <ItemLocation>
        <ItemType>AllFormulas</ItemType>
        <ItemPath></ItemPath>
      </ItemLocation>
      <StableEntries>
        <Entry Type="QueryGroups" Value="s[{&quot;Id&quot;:&quot;group-1&quot;,&quot;Name&quot;:&quot;Fact Tables&quot;}]"/>
      </StableEntries>
    </Item>
  </Items>
</LocalPackageMetadataFile>
'''


def _build_datamashup(section1_m: str, metadata_xml: str = "") -> bytes:
    """Build a minimal MS-QDEFF binary blob."""
    # Build the inner OPC ZIP (Package Parts)
    pkg_buf = io.BytesIO()
    with zipfile.ZipFile(pkg_buf, "w") as pkg:
        pkg.writestr("Formulas/Section1.m", section1_m)
        pkg.writestr("Config/Package.xml",
                     '<Package xmlns="http://schemas.microsoft.com/DataMashup">'
                     "<Version>2.0</Version></Package>")
    pkg_bytes = pkg_buf.getvalue()

    meta_bytes = metadata_xml.encode("utf-8")
    perm_bytes = b""
    bind_bytes = b""

    out = io.BytesIO()
    out.write(struct.pack("<I", 0))  # version
    out.write(struct.pack("<I", len(pkg_bytes)))
    out.write(pkg_bytes)
    out.write(struct.pack("<I", len(perm_bytes)))
    out.write(perm_bytes)
    out.write(struct.pack("<I", len(meta_bytes)))
    out.write(meta_bytes)
    out.write(struct.pack("<I", len(bind_bytes)))
    out.write(bind_bytes)
    return out.getvalue()


def _build_pbix(section1_m: str, metadata_xml: str = "") -> bytes:
    """Build a minimal .pbix ZIP containing a DataMashup blob."""
    mashup = _build_datamashup(section1_m, metadata_xml)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("DataMashup", mashup)
    return buf.getvalue()


def _build_pbit() -> bytes:
    """Build a minimal .pbit ZIP with DataModelSchema."""
    schema = {
        "name": "SemanticModel",
        "compatibilityLevel": 1550,
        "model": {
            "tables": [
                {
                    "name": "Sales",
                    "description": "Sales fact table",
                    "partitions": [
                        {
                            "name": "Sales",
                            "source": {
                                "type": "m",
                                "expression": [
                                    "let",
                                    '    Source = Sql.Database("server", "db"),',
                                    '    dbo_Sales = Source{[Schema="dbo",Item="Sales"]}[Data]',
                                    "in",
                                    "    dbo_Sales",
                                ],
                            },
                        }
                    ],
                }
            ],
            "expressions": [
                {
                    "name": "ServerName",
                    "kind": "m",
                    "expression": '"myserver" meta [IsParameterQuery=true, Type="Text"]',
                },
                {
                    "name": "fnClean",
                    "kind": "m",
                    "expression": [
                        "(tbl as table) =>",
                        "    Table.SelectRows(tbl, each [Active] = true)",
                    ],
                },
            ],
        },
    }
    schema_bytes = ("\ufeff" + json.dumps(schema)).encode("utf-16-le")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("DataModelSchema", schema_bytes)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_split_section_m():
    items = _split_section_m(SAMPLE_SECTION1_M)
    names = [n for n, _ in items]
    assert "ServerName" in names
    assert "My Query" in names
    assert "fnTransform" in names
    print("  PASS: test_split_section_m")


def test_classify():
    assert _classify_m("X", '"val" meta [IsParameterQuery=true, Type="Text"]') == MItemKind.PARAMETER
    assert _classify_m("fn", "(x as number) => x + 1") == MItemKind.FUNCTION
    assert _classify_m("Q", "let Source = 1 in Source") == MItemKind.QUERY
    print("  PASS: test_classify")


def test_parse_metadata_xml():
    meta = _parse_metadata_xml(SAMPLE_METADATA_XML)
    assert "My Query" in meta
    assert meta["My Query"].get("QueryGroup") == "Fact Tables"
    assert meta["My Query"].get("AddedToDataModel") == "1"
    print("  PASS: test_parse_metadata_xml")


def test_datamashup_roundtrip():
    section1_m, meta_xml = _parse_datamashup(
        _build_datamashup(SAMPLE_SECTION1_M, SAMPLE_METADATA_XML)
    )
    assert "shared" in section1_m or "ServerName" in section1_m
    assert "DataMashup" in meta_xml or "LocalPackageMetadata" in meta_xml
    print("  PASS: test_datamashup_roundtrip")


def test_extract_pbix():
    with tempfile.NamedTemporaryFile(suffix=".pbix", delete=False) as f:
        f.write(_build_pbix(SAMPLE_SECTION1_M, SAMPLE_METADATA_XML))
        f.flush()
        items = extract_from_datamashup(Path(f.name))
    assert len(items) == 3
    kinds = {i.kind for i in items}
    assert MItemKind.PARAMETER in kinds
    assert MItemKind.QUERY in kinds
    assert MItemKind.FUNCTION in kinds
    # Check metadata was populated
    query_item = next(i for i in items if i.name == "My Query")
    assert query_item.metadata.get("group") == "Fact Tables"
    assert query_item.metadata.get("load_enabled") == "true"
    print("  PASS: test_extract_pbix")


def test_extract_pbit():
    with tempfile.NamedTemporaryFile(suffix=".pbit", delete=False) as f:
        f.write(_build_pbit())
        f.flush()
        items = extract_from_pbit_json(Path(f.name))
    assert len(items) == 3
    names = {i.name for i in items}
    assert "ServerName" in names
    assert "fnClean" in names
    assert "Sales" in names
    param = next(i for i in items if i.name == "ServerName")
    assert param.kind == MItemKind.PARAMETER
    func = next(i for i in items if i.name == "fnClean")
    assert func.kind == MItemKind.FUNCTION
    print("  PASS: test_extract_pbit")


def test_render_markdown():
    items = [
        MItem("Param1", MItemKind.PARAMETER, '"val" meta [IsParameterQuery=true]', {"group": "Config"}),
        MItem("MyQuery", MItemKind.QUERY, "let Source = 1 in Source", {"load_enabled": "true"}),
        MItem("fnHelper", MItemKind.FUNCTION, "(x) => x + 1", {}),
    ]
    md = render_markdown(items, "test.pbix", include_metadata=True)
    assert "## Parameters" in md
    assert "## Queries" in md
    assert "## Functions" in md
    assert "### Param1" in md
    assert "```m" in md
    assert "**group**: Config" in md
    print("  PASS: test_render_markdown")


def test_cli_pbix():
    with tempfile.NamedTemporaryFile(suffix=".pbix", delete=False) as f:
        f.write(_build_pbix(SAMPLE_SECTION1_M, SAMPLE_METADATA_XML))
        f.flush()
        out = Path(f.name).with_suffix(".md")
        rc = main([f.name, "-o", str(out), "--metadata"])
    assert rc == 0
    assert out.exists()
    md = out.read_text()
    assert "## Parameters" in md
    assert "ServerName" in md
    out.unlink()
    print("  PASS: test_cli_pbix")


def test_cli_pbit():
    with tempfile.NamedTemporaryFile(suffix=".pbit", delete=False) as f:
        f.write(_build_pbit())
        f.flush()
        out = Path(f.name).with_suffix(".md")
        rc = main([f.name, "-o", str(out)])
    assert rc == 0
    assert out.exists()
    md = out.read_text()
    assert "Sales" in md
    out.unlink()
    print("  PASS: test_cli_pbit")


def test_analyze_dependencies_basic():
    """Test dependency detection on sample M code."""
    items = [
        MItem("ServerName", MItemKind.PARAMETER,
              '"myserver" meta [IsParameterQuery=true, Type="Text"]'),
        MItem("My Query", MItemKind.QUERY,
              'let\n    Source = Sql.Database(ServerName, "mydb"),\n'
              '    dbo_Sales = Source{[Schema="dbo",Item="Sales"]}[Data]\nin\n    dbo_Sales'),
        MItem("fnTransform", MItemKind.FUNCTION,
              "(tbl as table) =>\n    Table.AddColumn(tbl, \"New\", each [A] + [B])"),
    ]
    graph = analyze_dependencies(items)

    # My Query depends on ServerName (bare identifier match)
    assert ("My Query", "ServerName") in graph.edges
    # fnTransform has no dependencies
    assert not any(a == "fnTransform" for a, _ in graph.edges)
    # My Query uses Sql.Database
    assert "My Query" in graph.data_sources
    assert "Sql.Database" in graph.data_sources["My Query"]
    # Layers: ServerName + fnTransform first, My Query second
    assert len(graph.layers) == 2
    assert "ServerName" in graph.layers[0]
    assert "fnTransform" in graph.layers[0]
    assert "My Query" in graph.layers[1]
    print("  PASS: test_analyze_dependencies_basic")


def test_analyze_quoted_references():
    """Test that #"Quoted Name" references are detected."""
    items = [
        MItem("Raw Data", MItemKind.QUERY, 'let Source = Excel.Workbook(File.Contents("x.xlsx")) in Source'),
        MItem("Clean", MItemKind.QUERY, 'let Source = #"Raw Data", Filtered = Table.Skip(Source, 1) in Filtered'),
    ]
    graph = analyze_dependencies(items)
    assert ("Clean", "Raw Data") in graph.edges
    assert "Raw Data" in graph.data_sources
    assert "Excel.Workbook" in graph.data_sources["Raw Data"]
    print("  PASS: test_analyze_quoted_references")


def test_compute_layers_cycle():
    """Test graceful handling of circular dependencies."""
    names = {"A", "B", "C"}
    edges = [("A", "B"), ("B", "C"), ("C", "A")]
    layers = _compute_layers(names, edges)
    # All nodes should still appear somewhere
    all_nodes = [n for layer in layers for n in layer]
    assert set(all_nodes) == names
    print("  PASS: test_compute_layers_cycle")


def test_connectors_detected():
    """Test that various connectors are identified."""
    items = [
        MItem("Q1", MItemKind.QUERY, 'let Source = Web.Contents("https://api.example.com") in Source'),
        MItem("Q2", MItemKind.QUERY, 'let Source = OData.Feed("https://odata.example.com") in Source'),
        MItem("Q3", MItemKind.QUERY, 'let Source = Csv.Document(File.Contents("data.csv")) in Source'),
    ]
    graph = analyze_dependencies(items)
    assert "Web.Contents" in graph.data_sources.get("Q1", [])
    assert "OData.Feed" in graph.data_sources.get("Q2", [])
    assert "Csv.Document" in graph.data_sources.get("Q3", [])
    assert "File.Contents" in graph.data_sources.get("Q3", [])
    print("  PASS: test_connectors_detected")


def test_dependency_graph_to_dict():
    """Test serialization produces expected structure."""
    items = [
        MItem("A", MItemKind.QUERY, "let Source = 1 in Source"),
        MItem("B", MItemKind.QUERY, "let Source = A in Source"),
    ]
    graph = analyze_dependencies(items)
    d = dependency_graph_to_dict(graph)
    assert "edges" in d
    assert "data_sources" in d
    assert "nodes" in d
    assert "layers" in d
    assert isinstance(d["edges"], list)
    assert all("from" in e and "to" in e for e in d["edges"])
    print("  PASS: test_dependency_graph_to_dict")


def test_render_markdown_with_deps():
    """Test that markdown includes dependency info when graph is provided."""
    items = [
        MItem("ServerParam", MItemKind.PARAMETER, '"x" meta [IsParameterQuery=true]'),
        MItem("Sales", MItemKind.QUERY, 'let Source = Sql.Database(ServerParam, "db") in Source'),
    ]
    graph = analyze_dependencies(items)
    md = render_markdown(items, "test.pbix", include_metadata=True, dep_graph=graph)
    assert "## Dependency Summary" in md
    assert "depends on" in md
    assert "used by" in md
    assert "Sql.Database" in md
    print("  PASS: test_render_markdown_with_deps")


if __name__ == "__main__":
    print("Running tests...")
    test_split_section_m()
    test_classify()
    test_parse_metadata_xml()
    test_datamashup_roundtrip()
    test_extract_pbix()
    test_extract_pbit()
    test_render_markdown()
    test_cli_pbix()
    test_cli_pbit()
    test_analyze_dependencies_basic()
    test_analyze_quoted_references()
    test_compute_layers_cycle()
    test_connectors_detected()
    test_dependency_graph_to_dict()
    test_render_markdown_with_deps()
    print("\nAll tests passed!")
