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


class TextElement(MarkdownSlideElement):
    """Text element containing any markdown text content."""

    content_type: Literal[ContentType.TEXT] = ContentType.TEXT

    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        lines = []

        # Add HTML comment for element type and name (except for default text)
        if not (self.content_type == ContentType.TEXT and self.name == "Default"):
            lines.append(f"<!-- {self.content_type.value}: {self.name} -->")

        # Add content
        lines.append(self.content.rstrip())

        return "\n".join(lines)


class ImageElement(MarkdownSlideElement):
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
    def from_markdown(cls, name: str, markdown_content: str) -> "ImageElement":
        """Create ImageElement from markdown, extracting URL and metadata."""
        image_match = re.search(r"!\[([^]]*)\]\(([^)]+)\)", markdown_content.strip())
        if not image_match:
            raise ValueError("Image element must contain at least one markdown image (![alt](url))")

        alt_text = image_match.group(1)
        url = image_match.group(2)

        return cls(
            name=name,
            content=url,  # Store URL as content
            metadata={"alt_text": alt_text, "original_markdown": markdown_content.strip()},
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


class TableElement(MarkdownSlideElement):
    """Table element containing structured table data."""

    content_type: Literal[ContentType.TABLE] = ContentType.TABLE
    content: TableData = Field(...)  # Override content to be TableData instead of str

    @field_validator("content", mode="before")
    @classmethod
    def validate_and_parse_table(cls, v) -> TableData:
        """Validate markdown table using Marko with GFM extension and convert to structured data."""
        if isinstance(v, TableData):
            return v  # Already parsed

        if not isinstance(v, str):
            raise ValueError("Table content must be a string or TableData")

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

    @staticmethod
    def _extract_text_from_node(node) -> str:
        """Extract text content from a Marko AST node."""
        if hasattr(node, "children"):
            text_parts = []
            for child in node.children:
                if hasattr(child, "__class__") and child.__class__.__name__ == "RawText":
                    text_parts.append(str(child.children))
                else:
                    text_parts.append(TableElement._extract_text_from_node(child))
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
    def from_df(cls, df, name: str, metadata: dict[str, Any] = None) -> "TableElement":
        """Create TableElement from pandas DataFrame.

        Args:
            df: pandas DataFrame to convert
            name: Name for the table element
            metadata: Optional metadata dictionary

        Returns:
            TableElement: New table element with data from DataFrame

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


class ChartElement(MarkdownSlideElement):
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
    def extract_json_to_metadata(self) -> "ChartElement":
        """Extract JSON content to metadata field."""
        if self.content.strip().startswith("```json\n"):
            json_content = self.content.strip()[8:-4]  # Remove ```json\n and \n```
            try:
                parsed_json = json.loads(json_content)
                self.metadata["chart_data"] = parsed_json
            except json.JSONDecodeError:
                pass  # Let the field validator handle the error
        return self

    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        lines = []

        # Add HTML comment for element type and name
        lines.append(f"<!-- {self.content_type.value}: {self.name} -->")

        # Add content
        lines.append(self.content.rstrip())

        return "\n".join(lines)
