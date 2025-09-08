import pytest

from gslides_api.domain_old import (
    Color,
    ColorStop,
    CropProperties,
    Image,
    ImageProperties,
    Outline,
    Recolor,
    RecolorName,
    RgbColor,
    Shadow,
)


def test_image_properties_creation():
    """Test creating ImageProperties directly."""
    props = ImageProperties(
        transparency=0.5,
        brightness=0.2,
        contrast=-0.1,
        cropProperties=CropProperties(
            leftOffset=0.1,
            rightOffset=0.1,
            topOffset=0.1,
            bottomOffset=0.1,
            angle=0.0,
        ),
    )

    assert props.transparency == 0.5
    assert props.brightness == 0.2
    assert props.contrast == -0.1
    assert props.cropProperties.leftOffset == 0.1


def test_image_with_properties():
    """Test creating an Image with ImageProperties."""
    image = Image(
        contentUrl="https://example.com/image.jpg",
        imageProperties=ImageProperties(
            transparency=0.5,
            brightness=0.2,
        ),
    )

    assert isinstance(image.imageProperties, ImageProperties)
    assert image.imageProperties.transparency == 0.5
    assert image.imageProperties.brightness == 0.2


def test_image_with_dict_properties():
    """Test creating an Image with dict properties that get converted."""
    image = Image(
        contentUrl="https://example.com/image.jpg",
        imageProperties={
            "transparency": 0.5,
            "brightness": 0.2,
            "cropProperties": {
                "leftOffset": 0.1,
                "rightOffset": 0.1,
            },
        },
    )

    assert isinstance(image.imageProperties, ImageProperties)
    assert image.imageProperties.transparency == 0.5
    assert image.imageProperties.brightness == 0.2
    assert image.imageProperties.cropProperties.leftOffset == 0.1


def test_image_to_api_format():
    """Test converting an Image with ImageProperties to API format."""
    image = Image(
        contentUrl="https://example.com/image.jpg",
        imageProperties=ImageProperties(
            transparency=0.5,
            brightness=0.2,
            recolor=Recolor(
                name=RecolorName.GRAYSCALE,
                recolorStops=[
                    ColorStop(
                        color=Color(rgbColor=RgbColor(red=0.5, green=0.5, blue=0.5)),
                        position=0.5,
                    )
                ],
            ),
        ),
    )

    api_format = image.to_api_format()
    assert "imageProperties" in api_format
    assert api_format["imageProperties"]["transparency"] == 0.5
    assert api_format["imageProperties"]["brightness"] == 0.2
    assert api_format["imageProperties"]["recolor"]["name"] == "GRAYSCALE"
