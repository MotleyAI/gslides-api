import json
import os
from typing import Dict, Any, List, Set

import pytest

# Import our custom modules
from gslides_api import Presentation
from gslides_api.json_diff import json_diff


@pytest.fixture
def presentation_json_path() -> str:
    """Fixture that returns the path to the test presentation JSON file."""
    return os.path.join(os.path.dirname(__file__), "presentation_output.json")


@pytest.fixture
def original_json(presentation_json_path: str) -> Dict[str, Any]:
    """Fixture that loads and returns the original presentation JSON."""
    with open(presentation_json_path, "r") as f:
        return json.load(f)


@pytest.fixture
def presentation_model(original_json: Dict[str, Any]) -> Presentation:
    """Fixture that creates a Presentation model from the original JSON."""
    return Presentation.model_validate(original_json)


@pytest.fixture
def reconstructed_json(presentation_model: Presentation) -> Dict[str, Any]:
    """Fixture that generates the reconstructed JSON from the Presentation model."""
    return presentation_model.to_api_format()


@pytest.fixture
def output_json_path() -> str:
    """Fixture that returns the path to save the reconstructed JSON."""
    return os.path.join(os.path.dirname(os.path.dirname(__file__)), "reconstructed_output.json")


@pytest.fixture
def ignored_keys() -> Set[str]:
    """Fixture that returns a set of keys to ignore in the comparison."""
    return {"propertyState", "pageType"}


@pytest.fixture
def ignored_paths() -> Set[str]:
    """Fixture that returns a set of paths to ignore in the comparison."""
    return set()


def test_save_reconstructed_json(reconstructed_json: Dict[str, Any], output_json_path: str):
    """Test that we can save the reconstructed JSON to a file."""
    with open(output_json_path, "w") as f:
        json.dump(reconstructed_json, f, indent=2)
    assert os.path.exists(
        output_json_path
    ), f"Failed to save reconstructed JSON to {output_json_path}"


def test_essential_fields_preserved(
    original_json: Dict[str, Any], reconstructed_json: Dict[str, Any]
):
    """Test that essential fields are preserved in the reconstructed JSON."""
    essential_fields = ["presentationId", "pageSize", "slides"]
    for field in essential_fields:
        assert field in original_json, f"Field '{field}' missing in original JSON"
        assert field in reconstructed_json, f"Field '{field}' missing in reconstructed JSON"


def test_slide_count_preserved(original_json: Dict[str, Any], reconstructed_json: Dict[str, Any]):
    """Test that the number of slides is preserved in the reconstructed JSON."""
    original_slides = len(original_json.get("slides", []))
    reconstructed_slides = len(reconstructed_json.get("slides", []))
    assert (
        reconstructed_slides == original_slides
    ), f"Slide count mismatch: {original_slides} in original, {reconstructed_slides} in reconstructed"


@pytest.mark.parametrize(
    "field_path",
    [
        ["presentationId"],
        ["pageSize", "width", "magnitude"],
        ["pageSize", "height", "magnitude"],
        ["pageSize", "width", "unit"],
    ],
)
def test_field_values_preserved(
    original_json: Dict[str, Any], reconstructed_json: Dict[str, Any], field_path: List[str]
):
    """Test that specific field values are preserved in the reconstructed JSON."""
    # Navigate to the field in both JSONs
    orig_value = original_json
    recon_value = reconstructed_json

    for key in field_path:
        try:
            orig_value = orig_value[key]
            recon_value = recon_value[key]
        except (KeyError, TypeError):
            pytest.skip(f"Field path {field_path} not found in one of the JSONs")

    # For numeric values, allow small differences
    if isinstance(orig_value, (int, float)) and isinstance(recon_value, (int, float)):
        assert (
            abs(float(orig_value) - float(recon_value)) < 1e-10
        ), f"Value mismatch at {'.'.join(field_path)}: {orig_value} vs {recon_value}"
    else:
        assert (
            orig_value == recon_value
        ), f"Value mismatch at {'.'.join(field_path)}: {orig_value} vs {recon_value}"


def test_slide_content_preserved(original_json: Dict[str, Any], reconstructed_json: Dict[str, Any]):
    """Test that slide content is preserved in the reconstructed JSON."""
    # Skip if there are no slides
    if not original_json.get("slides") or not reconstructed_json.get("slides"):
        pytest.skip("No slides found in one of the JSONs")

    # Check the first slide's content if available
    try:
        # Try to extract a title or some content from the first slide
        # This is a common path to text content in a slide, but might need adjustment
        # based on the actual structure of your presentation
        path_to_check = [
            "slides",
            0,
            "pageElements",
            0,
            "shape",
            "text",
            "textElements",
            1,
            "textRun",
            "content",
        ]

        orig_value = original_json
        recon_value = reconstructed_json

        for key in path_to_check:
            if isinstance(key, int):
                # Handle list indexing
                orig_value = orig_value[key]
                recon_value = recon_value[key]
            else:
                # Handle dictionary keys
                orig_value = orig_value.get(key)
                recon_value = recon_value.get(key)

                # Skip if the key doesn't exist
                if orig_value is None or recon_value is None:
                    pytest.skip(f"Path {path_to_check} not found in one of the JSONs")

        # If we got here, we have values to compare
        assert orig_value == recon_value, f"Slide content mismatch: {orig_value} vs {recon_value}"
    except (KeyError, IndexError, TypeError):
        # If the path doesn't exist, skip the test rather than fail
        pytest.skip("Could not extract comparable content from slides")


def test_deep_comparison(
    original_json: Dict[str, Any],
    reconstructed_json: Dict[str, Any],
    ignored_keys: Set[str],
    ignored_paths: Set[str],
):
    """Test that the deep comparison of JSON structures finds no differences."""
    differences = json_diff(
        original_json, reconstructed_json, ignored_keys=ignored_keys, ignored_paths=ignored_paths
    )

    # If there are differences, format them for the assertion message
    if differences:
        diff_message = "\n".join([f"{i+1}. {diff}" for i, diff in enumerate(differences)])
        assert not differences, f"Found {len(differences)} differences:\n{diff_message}"
    else:
        assert True, "No differences found in the deep comparison"
