"""Tests for TableData.to_html() method."""

import pytest

from gslides_api.agnostic.element import TableData


class TestTableDataToHtml:
    def test_basic_table(self):
        td = TableData(headers=["Name", "Value"], rows=[["Alice", "100"], ["Bob", "200"]])
        result = td.to_html()
        assert "<table>" in result
        assert "<thead>" in result
        assert "<tbody>" in result
        assert "<th>Name</th>" in result
        assert "<th>Value</th>" in result
        assert "<td>Alice</td>" in result
        assert "<td>100</td>" in result
        assert "<td>Bob</td>" in result
        assert "<td>200</td>" in result

    def test_with_css_class(self):
        td = TableData(headers=["A"], rows=[["1"]])
        result = td.to_html(css_class="dtbl")
        assert '<table class="dtbl">' in result

    def test_without_css_class(self):
        td = TableData(headers=["A"], rows=[["1"]])
        result = td.to_html()
        assert "<table>" in result
        assert "class=" not in result

    def test_empty_headers(self):
        td = TableData(headers=[], rows=[])
        assert td.to_html() == ""

    def test_html_escaping(self):
        td = TableData(
            headers=["<script>", "A&B"],
            rows=[["x < y", '"quoted"']],
        )
        result = td.to_html()
        assert "&lt;script&gt;" in result
        assert "A&amp;B" in result
        assert "x &lt; y" in result
        assert "&quot;quoted&quot;" in result
        assert "<script>" not in result  # must be escaped

    def test_css_class_escaping(self):
        td = TableData(headers=["A"], rows=[["1"]])
        result = td.to_html(css_class='x" onclick="alert(1)')
        # The double quote must be escaped so it can't break out of the attribute
        assert '&quot;' in result
        assert 'class="x&quot; onclick=&quot;alert(1)"' in result

    def test_short_row_pads_with_empty(self):
        td = TableData(headers=["A", "B", "C"], rows=[["only_one"]])
        result = td.to_html()
        assert "<td>only_one</td>" in result
        assert result.count("<td></td>") == 2

    def test_structure_order(self):
        td = TableData(headers=["H"], rows=[["R"]])
        result = td.to_html()
        # Verify structural ordering
        assert result.index("<thead>") < result.index("</thead>")
        assert result.index("</thead>") < result.index("<tbody>")
        assert result.index("<tbody>") < result.index("</tbody>")
        assert result.index("</tbody>") < result.index("</table>")
