from typing import Any, Collection, List


def json_diff(
    original: Any,
    reconstructed: Any,
    path: str = "",
    ignored_keys: Collection[str] = None,
    ignored_paths: Collection[str] = None,
) -> List[str]:
    """
    Recursively compare two JSON structures and return a list of differences.

    Args:
        original: The original JSON structure
        reconstructed: The reconstructed JSON structure
        path: The current path in the JSON structure (for error reporting)
        ignored_keys: Set of keys to ignore in the comparison
        ignored_paths: Set of paths to ignore in the comparison

    Returns:
        A list of differences between the two structures
    """
    if ignored_keys is None:
        ignored_keys = set()
    else:
        ignored_keys = set(ignored_keys)

    if ignored_paths is None:
        ignored_paths = set()
    else:
        ignored_paths = set(ignored_paths)

    # Check if the current path should be ignored
    if any(
        path.startswith(ignored_path) for ignored_path in ignored_paths if ignored_path
    ):
        return []

    differences = []

    # If types are different, handle special cases
    if type(original) != type(reconstructed):
        # Allow int/float conversions for numeric values
        if isinstance(original, (int, float)) and isinstance(
            reconstructed, (int, float)
        ):
            # For numeric values, compare the actual values with a small tolerance
            if abs(float(original) - float(reconstructed)) > 1e-10:
                differences.append(
                    f"Value mismatch at {path}: {original} vs {reconstructed}"
                )
        else:
            differences.append(
                f"Type mismatch at {path}: {type(original)} vs {type(reconstructed)}"
            )
            return differences

    # Handle dictionaries
    elif isinstance(original, dict):
        # Check for missing keys
        original_keys = set(original.keys()) - ignored_keys
        reconstructed_keys = set(reconstructed.keys()) - ignored_keys

        # Keys in original but not in reconstructed
        for key in original_keys - reconstructed_keys:
            # Skip default style properties that might be missing in the original
            differences.append(
                f"Key '{key}' at {path} exists in original but not in reconstructed"
            )

        # Keys in reconstructed but not in original
        for key in reconstructed_keys - original_keys:
            differences.append(
                f"Key '{key}' at {path} exists in reconstructed but not in original"
            )

        # Recursively compare values for keys that exist in both
        for key in original_keys & reconstructed_keys:
            new_path = f"{path}.{key}" if path else key
            differences.extend(
                json_diff(
                    original[key],
                    reconstructed[key],
                    new_path,
                    ignored_keys,
                    ignored_paths,
                )
            )

    # Handle lists
    elif isinstance(original, list):
        if len(original) != len(reconstructed):
            differences.append(
                f"List length mismatch at {path}: {len(original)} vs {len(reconstructed)}"
            )
        else:
            # Recursively compare each item in the list
            for i, (orig_item, recon_item) in enumerate(zip(original, reconstructed)):
                new_path = f"{path}[{i}]"
                differences.extend(
                    json_diff(
                        orig_item, recon_item, new_path, ignored_keys, ignored_paths
                    )
                )

    # Handle primitive values (strings, numbers, booleans, None)
    elif original != reconstructed:
        # For floating point values, allow small differences
        if isinstance(original, float) and isinstance(reconstructed, float):
            if abs(original - reconstructed) > 1e-10:
                differences.append(
                    f"Value mismatch at {path}: {original} vs {reconstructed}"
                )
        else:
            differences.append(
                f"Value mismatch at {path}: {original} vs {reconstructed}"
            )

    return differences
