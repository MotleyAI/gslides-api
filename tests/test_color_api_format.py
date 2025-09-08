import json

import pytest

from gslides_api.domain_old import Color, RgbColor, ThemeColorType
from gslides_api.json_diff import json_diff


def test_color_to_api_format_vs_model_dump():
    """
    Test that compares the output of Color.to_api_format() and
    Color.model_dump(exclude_none=True, mode="json") using the json_diff function.
    """
    # Test cases with different Color configurations
    test_cases = [
        # Empty Color
        Color(),
        # Color with only rgbColor
        Color(rgbColor=RgbColor(red=0.5, green=0.7, blue=0.9)),
        # Color with only themeColor
        Color(themeColor=ThemeColorType.ACCENT1),
        # Color with both rgbColor and themeColor
        Color(
            rgbColor=RgbColor(red=0.1, green=0.2, blue=0.3),
            themeColor=ThemeColorType.DARK1,
        ),
        # Color with partial rgbColor (some values None)
        Color(rgbColor=RgbColor(red=0.5, green=None, blue=0.9)),
    ]

    # Test each case
    for i, color in enumerate(test_cases):
        # Get the outputs from both methods
        api_format = color.to_api_format()
        model_dump = color.model_dump(exclude_none=True, mode="json")

        # Use json_diff to find differences
        differences = json_diff(api_format, model_dump)

        # Always print the outputs for debugging
        print(f"\nTest case {i+1}:")
        print(f"to_api_format(): {json.dumps(api_format, indent=2)}")
        print(f"model_dump(): {json.dumps(model_dump, indent=2)}")
        if differences:
            print(f"Differences: {differences}")

        # Assert that there are no differences
        assert not differences, f"Found differences in test case {i+1}: {differences}"


def test_color_with_enum_handling():
    """
    Test specifically focusing on how enums are handled in to_api_format() vs model_dump().
    """
    # Create a Color with a theme color
    color = Color(themeColor=ThemeColorType.ACCENT1)

    # Get the outputs from both methods
    api_format = color.to_api_format()
    model_dump = color.model_dump(exclude_none=True, mode="json")

    # Check how the enum is represented in each output
    assert "themeColor" in api_format
    assert "themeColor" in model_dump

    # Print values for debugging
    print(f"\nEnum handling test:")
    print(f"API format: {json.dumps(api_format, indent=2)}")
    print(f"Model dump: {json.dumps(model_dump, indent=2)}")
    print(f"API format themeColor: {api_format['themeColor']}")
    print(f"Model dump themeColor: {model_dump['themeColor']}")

    # Check if there are differences in how the enum is handled
    differences = json_diff(api_format, model_dump)

    # If there are differences, they're likely in how the enum is serialized
    if differences:
        print(f"Differences in enum handling: {differences}")

        # The difference might be that to_api_format() uses .value while model_dump() might use the name
        # Let's check if that's the case
        if api_format["themeColor"] == "ACCENT1" and model_dump["themeColor"] != "ACCENT1":
            print(
                "to_api_format() uses enum.value while model_dump() uses a different representation"
            )
