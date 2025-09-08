"""
Test that compares the output of Color.to_api_format() and
Color.model_dump(exclude_none=True, mode="json") using the json_diff function.
"""

import json

import pytest

from gslides_api.domain_old import Color, RgbColor, ThemeColorType
from gslides_api.json_diff import json_diff


def test_color_format_comparison():
    """
    Test that compares the output of Color.to_api_format() and
    Color.model_dump(exclude_none=True, mode="json") using the json_diff function.
    """
    # Create a variety of Color objects to test
    test_cases = [
        # Empty Color
        ("Empty Color", Color()),
        # Color with only rgbColor
        ("RGB Color", Color(rgbColor=RgbColor(red=0.5, green=0.7, blue=0.9))),
        # Color with only themeColor
        ("Theme Color", Color(themeColor=ThemeColorType.ACCENT1)),
        # Color with both rgbColor and themeColor
        (
            "RGB and Theme Color",
            Color(
                rgbColor=RgbColor(red=0.1, green=0.2, blue=0.3),
                themeColor=ThemeColorType.DARK1,
            ),
        ),
        # Color with partial rgbColor (some values None)
        ("Partial RGB Color", Color(rgbColor=RgbColor(red=0.5, green=None, blue=0.9))),
    ]

    # Test each case
    for name, color in test_cases:
        # Get the outputs from both methods
        api_format = color.to_api_format()
        model_dump = color.model_dump(exclude_none=True, mode="json")

        # Print the outputs for debugging
        print(f"\n{name}:")
        print(f"to_api_format(): {json.dumps(api_format, indent=2)}")
        print(f"model_dump(): {json.dumps(model_dump, indent=2)}")

        # Use json_diff to find differences
        differences = json_diff(api_format, model_dump)

        if differences:
            print(f"Differences: {differences}")
        else:
            print("No differences found")

        # Assert that there are no differences
        assert not differences, f"Found differences in {name}: {differences}"

    print(
        "\nAll tests passed! Color.to_api_format() and Color.model_dump(exclude_none=True, mode='json') produce identical results."
    )


if __name__ == "__main__":
    # Run the test directly if this file is executed
    test_color_format_comparison()
