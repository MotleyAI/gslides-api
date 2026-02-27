"""Utilities for saving and loading presentation snapshots.

Snapshots are JSON files captured from real Google Slides API responses.
They can be loaded into a MockGoogleAPIClient for reproducible testing.

Usage:
    # Capture a snapshot from a real presentation:
    from gslides_api import GoogleAPIClient
    from gslides_api.mock.snapshots import dump_presentation_snapshot

    client = GoogleAPIClient()
    client.initialize_credentials("/path/to/creds")
    dump_presentation_snapshot(client, "real-presentation-id", "tests/snapshots/my_pres.json")

    # Load into a mock for testing:
    from gslides_api.mock import MockGoogleAPIClient, load_presentation_snapshot

    mock = MockGoogleAPIClient()
    load_presentation_snapshot(mock, "tests/snapshots/my_pres.json")
"""

import json
from typing import Optional

from gslides_api.client import GoogleAPIClient
from gslides_api.mock.client import MockGoogleAPIClient


def dump_presentation_snapshot(
    client: GoogleAPIClient,
    presentation_id: str,
    output_path: str,
) -> None:
    """Fetch a presentation from Google and save its JSON to a file.

    Args:
        client: An initialized GoogleAPIClient (real, with credentials).
        presentation_id: The ID of the presentation to snapshot.
        output_path: File path to write the JSON to.
    """
    data = client.get_presentation_json(presentation_id)
    with open(output_path, "w") as f:
        json.dump(data, f, indent=2)


def load_presentation_snapshot(
    mock_client: MockGoogleAPIClient,
    snapshot_path: str,
    presentation_id: Optional[str] = None,
) -> str:
    """Load a JSON snapshot file into a MockGoogleAPIClient.

    Args:
        mock_client: The mock client to load the snapshot into.
        snapshot_path: Path to the JSON snapshot file.
        presentation_id: Override the presentation ID. If None, uses the
            presentationId from the snapshot JSON.

    Returns:
        The presentation ID used for storage.
    """
    with open(snapshot_path) as f:
        data = json.load(f)

    pres_id = presentation_id or data.get("presentationId")
    if pres_id is None:
        raise ValueError(
            "No presentation_id provided and snapshot JSON has no 'presentationId' field"
        )

    mock_client.seed_presentation(pres_id, data)
    return pres_id
