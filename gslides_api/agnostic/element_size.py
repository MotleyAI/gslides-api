from pydantic import BaseModel, computed_field


class ElementSizeMeta(BaseModel):
    """Presentation element size metadata, captured during layout ingestion."""

    box_width_inches: float
    box_height_inches: float
    font_size_pt: float

    @computed_field
    @property
    def approx_char_capacity(self) -> int:
        """Estimate how many characters fit in this textbox."""
        char_width_in = self.font_size_pt * 0.5 / 72
        line_height_in = self.font_size_pt * 1.2 / 72
        if char_width_in <= 0 or line_height_in <= 0:
            return 0
        chars_per_line = self.box_width_inches / char_width_in
        num_lines = self.box_height_inches / line_height_in
        return int(chars_per_line * num_lines * 0.85)
