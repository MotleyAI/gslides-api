import unittest

from gslides_api.domain import (Color, Dimension, PageBackgroundFill,
                                PropertyState, RgbColor, Size, SolidFill,
                                StretchedPictureFill)


class TestPageBackgroundFill(unittest.TestCase):
    def test_solid_fill_background(self):
        # Create a solid fill background
        rgb_color = RgbColor(red=0.5, green=0.7, blue=0.9)
        color = Color(rgbColor=rgb_color)
        solid_fill = SolidFill(color=color, alpha=0.8)

        background_fill = PageBackgroundFill(
            propertyState=PropertyState.RENDERED, solidFill=solid_fill
        )

        # Convert to API format
        api_format = background_fill.to_api_format()

        # Verify the result
        self.assertEqual(api_format["propertyState"], "RENDERED")
        self.assertIn("solidFill", api_format)
        self.assertEqual(api_format["solidFill"]["alpha"], 0.8)
        self.assertEqual(api_format["solidFill"]["color"]["rgbColor"]["red"], 0.5)
        self.assertEqual(api_format["solidFill"]["color"]["rgbColor"]["green"], 0.7)
        self.assertEqual(api_format["solidFill"]["color"]["rgbColor"]["blue"], 0.9)

    def test_stretched_picture_fill_background(self):
        # Create a stretched picture fill background
        size = Size(
            width=Dimension(magnitude=100, unit="PT"),
            height=Dimension(magnitude=200, unit="PT"),
        )
        stretched_picture = StretchedPictureFill(
            contentUrl="https://example.com/image.jpg", size=size
        )

        background_fill = PageBackgroundFill(
            propertyState=PropertyState.RENDERED, stretchedPictureFill=stretched_picture
        )

        # Convert to API format
        api_format = background_fill.to_api_format()

        # Verify the result
        self.assertEqual(api_format["propertyState"], "RENDERED")
        self.assertIn("stretchedPictureFill", api_format)
        self.assertEqual(
            api_format["stretchedPictureFill"]["contentUrl"],
            "https://example.com/image.jpg",
        )
        self.assertEqual(
            api_format["stretchedPictureFill"]["size"]["width"]["magnitude"], 100
        )
        self.assertEqual(
            api_format["stretchedPictureFill"]["size"]["width"]["unit"], "PT"
        )
        self.assertEqual(
            api_format["stretchedPictureFill"]["size"]["height"]["magnitude"], 200
        )
        self.assertEqual(
            api_format["stretchedPictureFill"]["size"]["height"]["unit"], "PT"
        )


if __name__ == "__main__":
    unittest.main()
