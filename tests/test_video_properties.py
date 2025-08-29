import pytest

from gslides_api.domain import Outline, Video, VideoProperties, VideoSourceType


def test_video_properties_creation():
    """Test creating VideoProperties directly."""
    props = VideoProperties(
        autoPlay=True,
        start=10,
        end=60,
        mute=False,
    )

    assert props.autoPlay is True
    assert props.start == 10
    assert props.end == 60
    assert props.mute is False
    assert props.outline is None


def test_video_with_properties():
    """Test creating a Video with VideoProperties."""
    video = Video(
        url="https://www.youtube.com/watch?v=dQw4w9WgXcQ",
        source=VideoSourceType.YOUTUBE,
        id="dQw4w9WgXcQ",
        videoProperties=VideoProperties(
            autoPlay=True,
            start=10,
            end=60,
            mute=True,
        ),
    )

    assert isinstance(video.videoProperties, VideoProperties)
    assert video.videoProperties.autoPlay is True
    assert video.videoProperties.start == 10
    assert video.videoProperties.end == 60
    assert video.videoProperties.mute is True


def test_video_to_api_format():
    """Test converting a Video with VideoProperties to API format."""
    video = Video(
        url="https://www.youtube.com/watch?v=dQw4w9WgXcQ",
        source=VideoSourceType.YOUTUBE,
        id="dQw4w9WgXcQ",
        videoProperties=VideoProperties(
            autoPlay=True,
            start=10,
            end=60,
            mute=True,
        ),
    )

    api_format = video.to_api_format()
    assert "videoProperties" in api_format
    assert api_format["videoProperties"]["autoPlay"] is True
    assert api_format["videoProperties"]["start"] == 10
    assert api_format["videoProperties"]["end"] == 60
    assert api_format["videoProperties"]["mute"] is True
