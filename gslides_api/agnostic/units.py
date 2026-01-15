"""Output unit enum and EMU conversion functions.

This module is separate from converters.py to avoid circular imports,
since domain/domain.py needs to import these for unit conversions.
"""

from enum import Enum


# EMU (English Metric Unit) conversion constants
EMU_PER_PT = 12700      # 1 point = 12,700 EMUs
EMU_PER_INCH = 914400   # 1 inch = 914,400 EMUs
EMU_PER_CM = 360000     # 1 cm = 360,000 EMUs


class OutputUnit(Enum):
    """Output unit for dimension conversions."""

    EMU = "EMU"  # English Metric Units (native Google Slides unit)
    IN = "in"    # Inches
    CM = "cm"    # Centimeters
    PT = "pt"    # Points


def from_emu(value_emu: float, target_unit: OutputUnit) -> float:
    """Convert EMU value to target unit.

    Args:
        value_emu: Value in EMUs to convert.
        target_unit: Target unit for conversion.

    Returns:
        Converted value in target unit.

    Raises:
        ValueError: If target_unit is not a valid OutputUnit.
    """
    if target_unit == OutputUnit.EMU:
        return value_emu
    elif target_unit == OutputUnit.IN:
        return value_emu / EMU_PER_INCH
    elif target_unit == OutputUnit.CM:
        return value_emu / EMU_PER_CM
    elif target_unit == OutputUnit.PT:
        return value_emu / EMU_PER_PT
    else:
        raise ValueError(f"Unknown unit: {target_unit}")


def to_emu(value: float, source_unit: OutputUnit) -> float:
    """Convert value from source unit to EMU.

    Args:
        value: Value to convert.
        source_unit: Unit of the input value.

    Returns:
        Value converted to EMUs.

    Raises:
        ValueError: If source_unit is not a valid OutputUnit.
    """
    if source_unit == OutputUnit.EMU:
        return value
    elif source_unit == OutputUnit.IN:
        return value * EMU_PER_INCH
    elif source_unit == OutputUnit.CM:
        return value * EMU_PER_CM
    elif source_unit == OutputUnit.PT:
        return value * EMU_PER_PT
    else:
        raise ValueError(f"Unknown unit: {source_unit}")
