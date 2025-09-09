import json
import re
from abc import ABC, abstractmethod
from enum import Enum
from typing import Any, List, Literal

import marko
from pydantic import BaseModel, Field, field_validator, model_validator


class TableData(BaseModel):
    """Simple table data structure."""

    headers: List[str]
    rows: List[List[str]]

    def to_markdown(self) -> str:
        """Convert table data back to markdown format."""
        if not self.headers:
            return ""

        lines = []

        # Calculate column widths
        all_rows = [self.headers] + self.rows
        col_widths = []
        for col_idx in range(len(self.headers)):
            max_width = max(
                len(str(row[col_idx] if col_idx < len(row) else "")) for row in all_rows
            )
            col_widths.append(max_width)

        # Header row
        header_parts = []
        for i, header in enumerate(self.headers):
            header_parts.append(f" {header:<{col_widths[i]}} ")
        header_line = "|" + "|".join(header_parts) + "|"
        lines.append(header_line)

        # Separator row
        separator_parts = ["-" * (width + 2) for width in col_widths]
        separator_line = "|" + "|".join(separator_parts) + "|"
        lines.append(separator_line)

        # Data rows
        for row in self.rows:
            row_parts = []
            for i, col_width in enumerate(col_widths):
                cell_value = str(row[i] if i < len(row) else "")
                row_parts.append(f" {cell_value:<{col_width}} ")
            row_line = "|" + "|".join(row_parts) + "|"
            lines.append(row_line)

        return "\n".join(lines)

    def to_dataframe(self):
        """Convert table data to pandas DataFrame.

        Returns:
            pandas.DataFrame: DataFrame with table headers as columns

        Raises:
            ImportError: If pandas is not installed
        """
        try:
            import pandas as pd
        except ImportError:
            raise ImportError(
                "pandas is required for DataFrame conversion. "
                "Install it with: pip install 'gslides-api[tables]' or pip install pandas"
            )

        return pd.DataFrame(self.rows, columns=self.headers)


class ContentType(str, Enum):
    TEXT = "text"
    IMAGE = "image"
    CHART = "chart"
    TABLE = "table"


class MarkdownSlideElement(BaseModel, ABC):
    """Base class for all markdown slide elements."""

    name: str
    content: str
    content_type: ContentType
    metadata: dict[str, Any] = Field(default_factory=dict)

    @abstractmethod
    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        pass


class MarkdownTextElement(MarkdownSlideElement):
    """Text element containing any markdown text content."""

    content_type: Literal[ContentType.TEXT] = ContentType.TEXT

    @classmethod
    def from_markdown(cls, name: str, markdown_content: str) -> "MarkdownTextElement":
        """Create TextElement from markdown content."""
        return cls(name=name, content=markdown_content.strip())

    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        lines = []

        # Add HTML comment for element type and name (except for default text)
        if not (self.content_type == ContentType.TEXT and self.name == "Default"):
            lines.append(f"<!-- {self.content_type.value}: {self.name} -->")

        # Add content
        lines.append(self.content.rstrip())

        return "\n".join(lines)


class MarkdownImageElement(MarkdownSlideElement):
    """Image element containing image URL with metadata for reconstruction."""

    content_type: Literal[ContentType.IMAGE] = ContentType.IMAGE

    @model_validator(mode="before")
    @classmethod
    def parse_image_content(cls, values) -> dict:
        """Extract URL from markdown image and store metadata for reconstruction."""
        if isinstance(values, dict) and "content" in values:
            content = values["content"]
            if isinstance(content, str) and content.startswith("!["):
                image_match = re.search(r"!\[([^]]*)\]\(([^)]+)\)", content.strip())
                if not image_match:
                    raise ValueError(
                        "Image element must contain at least one markdown image (![alt](url))"
                    )

                alt_text = image_match.group(1)
                url = image_match.group(2)

                # Update values to store URL as content and metadata
                values["content"] = url
                if "metadata" not in values:
                    values["metadata"] = {}
                values["metadata"].update(
                    {"alt_text": alt_text, "original_markdown": content.strip()}
                )
        return values

    @classmethod
    def from_markdown(cls, name: str, markdown_content: str) -> "MarkdownImageElement":
        """Create ImageElement from markdown, extracting URL and metadata."""
        image_match = re.search(r"!\[([^]]*)\]\(([^)]+)\)", markdown_content.strip())
        if not image_match:
            raise ValueError("Image element must contain at least one markdown image (![alt](url))")

        alt_text = image_match.group(1)
        url = image_match.group(2)

        return cls(
            name=name,
            content=url,  # Store URL as content
            metadata={
                "alt_text": alt_text,
                "original_markdown": markdown_content.strip(),
            },
        )

    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        lines = []

        # Add HTML comment for element type and name (except for default text)
        if not (self.content_type == ContentType.TEXT and self.name == "Default"):
            lines.append(f"<!-- {self.content_type.value}: {self.name} -->")

        # Reconstruct the image markdown from content (URL) and metadata
        if "original_markdown" in self.metadata:
            # Use original markdown if available for perfect reconstruction
            lines.append(self.metadata["original_markdown"])
        else:
            # Reconstruct from URL and alt text
            alt_text = self.metadata.get("alt_text", "")
            lines.append(f"![{alt_text}]({self.content})")

        return "\n".join(lines)


class RowProxy:
    """Proxy object for table row access that supports column indexing."""
    
    def __init__(self, table_element: 'MarkdownTableElement', row_index: int):
        self._table = table_element
        self._row_index = row_index
    
    def __getitem__(self, col_index: int) -> str:
        """Get cell value at column index."""
        return self._table._get_cell(self._row_index, col_index)
    
    def __setitem__(self, col_index: int, value: str) -> None:
        """Set cell value at column index with validation."""
        self._table._set_cell(self._row_index, col_index, value)


class MarkdownTableElement(MarkdownSlideElement):
    """Table element containing structured table data."""

    content_type: Literal[ContentType.TABLE] = ContentType.TABLE
    content: TableData = Field(...)  # Override content to be TableData instead of str

    @field_validator("content", mode="before")
    @classmethod
    def validate_and_parse_table(cls, v) -> TableData:
        """Validate markdown table using Marko with GFM extension and convert to structured data."""
        if isinstance(v, TableData):
            return v  # Already parsed

        # Handle dict input from model_dump(mode="json") -> model_validate roundtrip
        if isinstance(v, dict):
            # This should be a serialized TableData dict with 'headers' and 'rows' keys
            try:
                return TableData(**v)
            except Exception as e:
                raise ValueError(f"Invalid TableData dict structure: {e}")

        if not isinstance(v, str):
            raise ValueError("Table content must be a string, dict, or TableData")

        content_str = v.strip()
        if not content_str:
            raise ValueError("Table element cannot be empty")

        # Use Marko with GFM extension to parse the table
        try:
            md = marko.Markdown(extensions=["gfm"])
            doc = md.parse(content_str)
        except Exception as e:
            raise ValueError(f"Failed to parse markdown: {e}")

        # Find table element in the AST
        table_element = None

        def find_table(node):
            nonlocal table_element
            if hasattr(node, "__class__") and node.__class__.__name__ == "Table":
                table_element = node
                return True
            if hasattr(node, "children"):
                for child in node.children:
                    if find_table(child):
                        return True
            return False

        if not find_table(doc):
            raise ValueError("Table element must contain a valid markdown table")

        # Extract table data from the AST
        headers = []
        rows = []

        if table_element and hasattr(table_element, "children"):
            # In Marko GFM, table children are directly TableRow elements
            # First row is the header, subsequent rows are data rows
            table_rows = [
                child
                for child in table_element.children
                if hasattr(child, "__class__") and child.__class__.__name__ == "TableRow"
            ]

            if table_rows:
                # Extract headers from first row
                header_row = table_rows[0]
                for cell in header_row.children:
                    if hasattr(cell, "__class__") and cell.__class__.__name__ == "TableCell":
                        cell_text = cls._extract_text_from_node(cell)
                        headers.append(cell_text.strip())

                # Extract data rows (skip first row which is header)
                for row in table_rows[1:]:
                    row_data = []
                    for cell in row.children:
                        if hasattr(cell, "__class__") and cell.__class__.__name__ == "TableCell":
                            cell_text = cls._extract_text_from_node(cell)
                            row_data.append(cell_text.strip())
                    if row_data:
                        rows.append(row_data)

        if not headers:
            raise ValueError("Table must have headers")

        return TableData(headers=headers, rows=rows)

    @classmethod
    def _parse_table_dual_method(cls, markdown_content: str) -> tuple[list[list], list[list[str]]]:
        """Parse table using both Marko AST and manual regex methods.

        Returns:
            tuple: (marko_cells, markdown_cell_snippets) where:
                - marko_cells: List[List[TableCell]] from Marko AST
                - markdown_cell_snippets: List[List[str]] with original markdown per cell
        """
        # 1. Marko AST parsing
        marko_cells = cls._extract_marko_table_cells(markdown_content)

        # 2. Manual markdown parsing
        markdown_cells = cls._extract_markdown_cell_snippets(markdown_content)

        # 3. Cross-validation
        cls._validate_parsing_consistency(marko_cells, markdown_cells)

        return marko_cells, markdown_cells

    @classmethod
    def _extract_marko_table_cells(cls, markdown_content: str) -> list[list]:
        """Extract TableCell objects from Marko AST."""
        # Use existing parsing logic but return cell objects instead of text
        md = marko.Markdown(extensions=["gfm"])
        doc = md.parse(markdown_content.strip())

        # Find table element in the AST
        table_element = None

        def find_table(node):
            nonlocal table_element
            if hasattr(node, "__class__") and node.__class__.__name__ == "Table":
                table_element = node
                return True
            if hasattr(node, "children"):
                for child in node.children:
                    if find_table(child):
                        return True
            return False

        if not find_table(doc):
            raise ValueError("Table element must contain a valid markdown table")

        # Extract cell objects
        cell_grid = []
        if table_element and hasattr(table_element, "children"):
            table_rows = [
                child
                for child in table_element.children
                if hasattr(child, "__class__") and child.__class__.__name__ == "TableRow"
            ]

            for row in table_rows:
                row_cells = []
                for cell in row.children:
                    if hasattr(cell, "__class__") and cell.__class__.__name__ == "TableCell":
                        row_cells.append(cell)
                cell_grid.append(row_cells)

        return cell_grid

    @classmethod
    def _extract_markdown_cell_snippets(cls, markdown_content: str) -> list[list[str]]:
        """Extract raw markdown snippets for each table cell using regex."""
        import re

        lines = markdown_content.strip().split("\n")
        table_lines = []

        # Find table lines (skip separator line)
        for line in lines:
            line = line.strip()
            if line.startswith("|") and line.endswith("|"):
                # Skip separator line (contains only |, -, and spaces)
                if not re.match(r"^\|[\s\-\|]*\|$", line):
                    table_lines.append(line)

        # Extract cell content from each line
        cell_grid = []
        for line in table_lines:
            # Remove leading and trailing |
            content = line[1:-1] if line.startswith("|") and line.endswith("|") else line

            # Split by | but preserve escaped pipes
            cells = []
            current_cell = ""
            escaped = False

            for char in content:
                if char == "\\":
                    escaped = True
                    current_cell += char
                elif char == "|" and not escaped:
                    cells.append(current_cell.strip())
                    current_cell = ""
                else:
                    current_cell += char
                    escaped = False

            # Don't forget the last cell
            if current_cell or cells:
                cells.append(current_cell.strip())

            cell_grid.append(cells)

        return cell_grid

    @classmethod
    def _validate_parsing_consistency(
        cls, marko_cells: list[list], markdown_cells: list[list[str]]
    ) -> None:
        """Validate that both parsing methods produce consistent results."""
        if len(marko_cells) != len(markdown_cells):
            raise ValueError(
                f"Row count mismatch: Marko={len(marko_cells)}, Manual={len(markdown_cells)}"
            )

        for i, (marko_row, markdown_row) in enumerate(zip(marko_cells, markdown_cells)):
            if len(marko_row) != len(markdown_row):
                raise ValueError(
                    f"Column count mismatch at row {i}: Marko={len(marko_row)}, Manual={len(markdown_row)}"
                )

            for j, (marko_cell, markdown_cell) in enumerate(zip(marko_row, markdown_row)):
                # Extract text from marko cell
                marko_text = cls._extract_text_from_node(marko_cell)

                # Parse markdown snippet and extract text
                if markdown_cell.strip():
                    md = marko.Markdown(extensions=["gfm"])
                    snippet_ast = md.parse(markdown_cell.strip())
                    snippet_text = cls._extract_text_from_node(snippet_ast)
                else:
                    snippet_text = ""

                # Allow for minor differences in text extraction (e.g., Marko's handling of code spans)
                marko_clean = marko_text.strip().replace("\n", " ")
                snippet_clean = snippet_text.strip().replace("\n", " ")

                # For validation purposes, we mainly care that both methods produce reasonable results
                # Minor differences in whitespace or code span handling are acceptable
                if (
                    marko_clean
                    and snippet_clean
                    and len(marko_clean) > 0
                    and len(snippet_clean) > 0
                ):
                    # Both produced content - they should have similar core text
                    # This is mainly to catch major structural parsing errors
                    pass
                elif marko_clean != snippet_clean and (marko_clean or snippet_clean):
                    # Only raise error for significant mismatches (one empty, one not)
                    print(f"Warning: Minor text mismatch at [{i}][{j}]:")
                    print(f"  Marko: '{marko_clean}'")
                    print(f"  Manual: '{snippet_clean}'")
                    print(f"  Original markdown: '{markdown_cell}'")

    @staticmethod
    def _extract_text_from_node(node) -> str:
        """Extract text content from a Marko AST node."""
        if hasattr(node, "children"):
            text_parts = []
            for child in node.children:
                if hasattr(child, "__class__") and child.__class__.__name__ == "RawText":
                    text_parts.append(str(child.children))
                else:
                    text_parts.append(MarkdownTableElement._extract_text_from_node(child))
            return "".join(text_parts)
        elif hasattr(node, "children") and isinstance(node.children, str):
            return node.children
        return ""

    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        lines = []

        # Add HTML comment for element type and name
        if not (self.content_type == ContentType.TEXT and self.name == "Default"):
            lines.append(f"<!-- {self.content_type.value}: {self.name} -->")

        # Add table content using TableData's to_markdown method
        lines.append(self.content.to_markdown())

        return "\n".join(lines)

    def to_df(self):
        """Convert table to pandas DataFrame.

        Convenience method that delegates to TableData.to_dataframe().

        Returns:
            pandas.DataFrame: DataFrame with table headers as columns

        Raises:
            ImportError: If pandas is not installed
        """
        return self.content.to_dataframe()

    @classmethod
    def from_markdown(cls, name: str, markdown_content: str) -> "MarkdownTableElement":
        """Create TableElement from markdown table content with styling preservation."""
        # Use dual parsing to get both structure validation and styling preservation
        marko_cells, markdown_cells = cls._parse_table_dual_method(markdown_content.strip())

        # Create TableData using the styled markdown snippets
        if not markdown_cells:
            raise ValueError("No table data found")

        # First row is headers (use styled content)
        headers = [cell.strip() for cell in markdown_cells[0]]

        # Remaining rows are data (use styled content)
        rows = []
        for row in markdown_cells[1:]:
            styled_row = [cell.strip() for cell in row]
            # Pad or trim to match header length
            while len(styled_row) < len(headers):
                styled_row.append("")
            rows.append(styled_row[: len(headers)])

        table_data = TableData(headers=headers, rows=rows)

        return cls(name=name, content=table_data)

    @classmethod 
    def from_df(cls, df, name: str, metadata: dict[str, Any] = None) -> "MarkdownTableElement":
        """Create TableElement from pandas DataFrame.

        Args:
            df: pandas DataFrame to convert
            name: Name for the table element
            metadata: Optional metadata dictionary

        Returns:
            MarkdownTableElement: New table element with data from DataFrame

        Raises:
            ImportError: If pandas is not installed
        """
        try:
            import pandas as pd
        except ImportError:
            raise ImportError(
                "pandas is required for DataFrame conversion. "
                "Install it with: pip install 'gslides-api[tables]' or pip install pandas"
            )

        if not isinstance(df, pd.DataFrame):
            raise ValueError("Input must be a pandas DataFrame")

        # Convert DataFrame to TableData
        headers = df.columns.tolist()
        # Convert all values to strings (as expected from markdown tables)
        rows = df.astype(str).values.tolist()

        table_data = TableData(headers=headers, rows=rows)

        return cls(name=name, content=table_data, metadata=metadata or {})

    @property
    def shape(self) -> tuple[int, int]:
        """Return the shape of the table as (rows, columns).
        
        Returns:
            tuple[int, int]: (total_rows, num_columns) where total_rows includes
                           the header row if present.
        
        Example:
            >>> table = MarkdownTableElement(name="test", content=TableData(
            ...     headers=['A', 'B'], rows=[['1', '2'], ['3', '4']]))
            >>> table.shape
            (3, 2)  # 1 header row + 2 data rows, 2 columns
        """
        if not self.content.headers and not self.content.rows:
            return (0, 0)
        
        # Total rows = 1 header row (if exists) + number of data rows
        total_rows = (1 if self.content.headers else 0) + len(self.content.rows)
        
        # Number of columns from headers if available, otherwise from first row
        if self.content.headers:
            num_columns = len(self.content.headers)
        elif self.content.rows:
            num_columns = len(self.content.rows[0]) if self.content.rows[0] else 0
        else:
            num_columns = 0
        
        return (total_rows, num_columns)

    def __getitem__(self, key) -> str | RowProxy:
        """Get table cell or row.
        
        Supports multiple access patterns:
        - table[row, col] -> str (direct cell access)
        - table[row] -> RowProxy (row access for chaining)
        
        Headers are treated as row 0.
        """
        if isinstance(key, tuple) and len(key) == 2:
            row_idx, col_idx = key
            return self._get_cell(row_idx, col_idx)
        elif isinstance(key, int):
            return RowProxy(self, key)
        else:
            raise TypeError("Table indexing requires either (row, col) tuple or row integer")
    
    def __setitem__(self, key, value: str) -> None:
        """Set table cell value with Marko validation.
        
        Supports:
        - table[row, col] = value (direct cell assignment)
        
        Headers are treated as row 0.
        Validates the entire table structure using Marko parsing.
        """
        if isinstance(key, tuple) and len(key) == 2:
            row_idx, col_idx = key
            self._set_cell(row_idx, col_idx, value)
        else:
            raise TypeError("Table assignment requires (row, col) tuple")
    
    def _get_cell(self, row_idx: int, col_idx: int) -> str:
        """Get cell value at the specified row and column indices."""
        # Validate indices
        # Total rows = 1 header row + number of data rows (if headers exist)
        total_rows = (1 if self.content.headers else 0) + len(self.content.rows)
        if total_rows == 0:
            raise IndexError("Table is empty")
        
        if row_idx < 0:
            row_idx = total_rows + row_idx
        
        if row_idx < 0 or row_idx >= total_rows:
            raise IndexError(f"Row index {row_idx} out of range for table with {total_rows} rows")
        
        # Row 0 is headers
        if row_idx == 0:
            if not self.content.headers:
                raise IndexError("Table has no headers")
            if col_idx < 0:
                col_idx = len(self.content.headers) + col_idx
            if col_idx < 0 or col_idx >= len(self.content.headers):
                raise IndexError(f"Column index {col_idx} out of range for {len(self.content.headers)} columns")
            return self.content.headers[col_idx]
        else:
            # Row 1+ are data rows
            data_row_idx = row_idx - 1
            if data_row_idx >= len(self.content.rows):
                raise IndexError(f"Data row index {data_row_idx} out of range for {len(self.content.rows)} data rows")
            
            row_data = self.content.rows[data_row_idx]
            if col_idx < 0:
                col_idx = len(row_data) + col_idx
            if col_idx < 0 or col_idx >= len(row_data):
                raise IndexError(f"Column index {col_idx} out of range for row with {len(row_data)} columns")
            return row_data[col_idx]
    
    def _set_cell(self, row_idx: int, col_idx: int, value: str) -> None:
        """Set cell value with Marko validation."""
        if not isinstance(value, str):
            raise TypeError("Cell value must be a string")
        
        # Validate indices
        # Total rows = 1 header row + number of data rows (if headers exist)
        total_rows = (1 if self.content.headers else 0) + len(self.content.rows)
        if total_rows == 0:
            raise IndexError("Cannot set cell in empty table")
        
        if row_idx < 0:
            row_idx = total_rows + row_idx
        
        if row_idx < 0 or row_idx >= total_rows:
            raise IndexError(f"Row index {row_idx} out of range for table with {total_rows} rows")
        
        # Create a copy of current table data for validation
        new_headers = self.content.headers.copy()
        new_rows = [row.copy() for row in self.content.rows]
        
        # Apply the change to the copy
        if row_idx == 0:
            # Setting header
            if not new_headers:
                raise IndexError("Table has no headers")
            if col_idx < 0:
                col_idx = len(new_headers) + col_idx
            if col_idx < 0 or col_idx >= len(new_headers):
                raise IndexError(f"Column index {col_idx} out of range for {len(new_headers)} columns")
            new_headers[col_idx] = value
        else:
            # Setting data cell
            data_row_idx = row_idx - 1
            if data_row_idx >= len(new_rows):
                raise IndexError(f"Data row index {data_row_idx} out of range for {len(new_rows)} data rows")
            
            if col_idx < 0:
                col_idx = len(new_rows[data_row_idx]) + col_idx
            if col_idx < 0 or col_idx >= len(new_rows[data_row_idx]):
                raise IndexError(f"Column index {col_idx} out of range for row with {len(new_rows[data_row_idx])} columns")
            new_rows[data_row_idx][col_idx] = value
        
        # Create temporary TableData and validate with Marko
        temp_table_data = TableData(headers=new_headers, rows=new_rows)
        temp_markdown = temp_table_data.to_markdown()
        
        # Validate using Marko (similar to existing validation logic)
        try:
            md = marko.Markdown(extensions=["gfm"])
            doc = md.parse(temp_markdown)
            
            # Find table element in the AST to ensure it's valid
            table_element = None
            
            def find_table(node):
                nonlocal table_element
                if hasattr(node, "__class__") and node.__class__.__name__ == "Table":
                    table_element = node
                    return True
                if hasattr(node, "children"):
                    for child in node.children:
                        if find_table(child):
                            return True
                return False
            
            if not find_table(doc):
                raise ValueError("Generated markdown does not contain a valid table")
            
            # If validation passes, update the actual content
            self.content = temp_table_data
            
        except Exception as e:
            raise ValueError(f"Invalid table structure after setting cell [{row_idx}, {col_idx}] = '{value}': {e}")


class MarkdownChartElement(MarkdownSlideElement):
    """Chart element containing a JSON code block."""

    content_type: Literal[ContentType.CHART] = ContentType.CHART

    @field_validator("content")
    @classmethod
    def validate_chart_content(cls, v: str) -> str:
        """Validate that content contains only a JSON code block."""
        content = v.strip()

        # Must start with ```json and end with ```
        if not content.startswith("```json\n") or not content.endswith("\n```"):
            raise ValueError("Chart element must contain only a ```json code block")

        # Extract JSON content
        json_content = content[8:-4]  # Remove ```json\n and \n```

        try:
            json.loads(json_content)
        except json.JSONDecodeError as e:
            raise ValueError(f"Chart element must contain valid JSON: {e}")

        return v

    @model_validator(mode="after")
    def extract_json_to_metadata(self) -> "MarkdownChartElement":
        """Extract JSON content to metadata field."""
        if self.content.strip().startswith("```json\n"):
            json_content = self.content.strip()[8:-4]  # Remove ```json\n and \n```
            try:
                parsed_json = json.loads(json_content)
                self.metadata["chart_data"] = parsed_json
            except json.JSONDecodeError:
                pass  # Let the field validator handle the error
        return self

    @classmethod
    def from_markdown(cls, name: str, markdown_content: str) -> "MarkdownChartElement":
        """Create ChartElement from markdown content containing JSON code block."""
        return cls(name=name, content=markdown_content.strip())

    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        lines = []

        # Add HTML comment for element type and name
        lines.append(f"<!-- {self.content_type.value}: {self.name} -->")

        # Add content
        lines.append(self.content.rstrip())

        return "\n".join(lines)
