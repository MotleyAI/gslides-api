from typing import Any, Dict, List, Optional

from pydantic import Field

from gslides_api.domain import (
    BulletGlyphPreset,
    GSlidesBaseModel,
    TextStyle,
    ShapeType,
    LayoutReference,
    ShapeProperties,
    ImageProperties,
    Size,
    Transform,
)
from gslides_api.request.domain import (
    Range,
    TableCellLocation,
    ElementProperties,
    PlaceholderIdMapping,
    ObjectIdMapping,
)


class GslidesAPIRequest(GSlidesBaseModel):
    """Base class for all requests to the Google Slides API."""

    def to_request(self) -> List[Dict[str, Any]]:
        """Convert to the format expected by the Google Slides API."""
        request_name = self.__class__.__name__.replace("Request", "")
        # make first letter lowercase
        request_name = request_name[0].lower() + request_name[1:]

        return [{request_name: self.to_api_format()}]


class CreateParagraphBulletsRequest(GslidesAPIRequest):
    """Creates bullets for paragraphs in a shape or table cell.

    This request converts plain paragraphs into bulleted lists using a specified
    bullet preset pattern. The bullets are applied to all paragraphs that overlap
    with the given text range.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createparagraphbulletsrequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table containing the text to add bullets to"
    )
    textRange: Range = Field(
        description="The range of text to add bullets to, based on TextElement indexes"
    )
    bulletPreset: Optional[BulletGlyphPreset] = Field(
        default=None, description="The kinds of bullet glyphs to be used"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The optional table cell location if the text to be modified is in a table cell. If present, the objectId must refer to a table.",
    )


class InsertTextRequest(GslidesAPIRequest):
    """Inserts text into a shape or table cell.

    This request inserts text at the specified insertion index within a shape or table cell.
    The text is inserted at the given index, and all existing text at and after that index
    is shifted to accommodate the new text.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#inserttextrequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table containing the text to insert into"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The optional table cell location if the text is to be inserted into a table cell. If present, the objectId must refer to a table.",
    )
    text: str = Field(description="The text to insert")
    insertionIndex: Optional[int] = Field(
        description="The index where the text will be inserted, in Unicode code units. Text is inserted before the character currently at this index. An insertion index of 0 will insert the text at the beginning of the text."
    )


class UpdateTextStyleRequest(GslidesAPIRequest):
    """Updates the styling of text within a Shape or Table.

    This request updates the text style for the specified range of text within a shape or table cell.
    The style changes are applied to all text elements that overlap with the given text range.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatetextstylerequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table with the text to be styled"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The location of the cell in the table containing the text to style. If objectId refers to a table, cellLocation must have a value. Otherwise, it must not.",
    )
    style: TextStyle = Field(
        description="The style(s) to set on the text. If the value for a particular style matches that of the parent, that style will be set to inherit."
    )
    textRange: Range = Field(
        description="The range of text to style. The range may be extended to include adjacent newlines. If the range fully contains a paragraph belonging to a list, the paragraph's bullet is also updated with the matching text style."
    )
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'style' is implied and should not be specified. A single '*' can be used as short-hand for listing every field. For example, to update the text style to bold, set fields to 'bold'."
    )


class DeleteTextRequest(GslidesAPIRequest):
    """Deletes text from a shape or table cell.

    This request deletes text from the specified range within a shape or table cell.
    The text range can be specified to delete all text or a specific range.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#deletetextrequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table containing the text to delete"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The optional table cell location if the text to be deleted is in a table cell. If present, the objectId must refer to a table.",
    )
    textRange: Range = Field(
        description="The range of text to delete, based on TextElement indexes"
    )


class CreateShapeRequest(GslidesAPIRequest):
    """Creates a new shape.

    This request creates a new shape on the specified page. The shape can be of various types
    like text box, rectangle, ellipse, etc.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createshaperequest
    """

    objectId: Optional[str] = Field(
        default=None,
        description="A user-supplied object ID. If specified, the ID must be unique among all pages and page elements in the presentation.",
    )
    elementProperties: ElementProperties = Field(description="The element properties for the shape")
    shapeType: ShapeType = Field(description="The shape type")


class UpdateShapePropertiesRequest(GslidesAPIRequest):
    """Updates the properties of a Shape.

    This request updates the shape properties for the specified shape.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateshapepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the shape to update")
    shapeProperties: ShapeProperties = Field(description="The shape properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'shapeProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class ReplaceImageRequest(GslidesAPIRequest):
    """Replaces an existing image with a new image.

    This request replaces the image at the specified object ID with a new image from the provided URL.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#replaceimagerequest
    """

    imageObjectId: str = Field(description="The ID of the existing image that will be replaced")
    url: str = Field(
        description="The image URL. The image is fetched once at insertion time and a copy is stored for display inside the presentation. Images must be less than 50MB in size, cannot exceed 25 megapixels, and must be in one of PNG, JPEG, or GIF format."
    )
    imageReplaceMethod: Optional[str] = Field(
        default="CENTER_INSIDE",
        description="The image replace method. This field is optional and defaults to CENTER_INSIDE.",
    )


class CreateSlideRequest(GslidesAPIRequest):
    """Creates a new slide.

    This request creates a new slide in the presentation. The slide can be created with a specific
    layout or as a blank slide.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createsliderequest
    """

    objectId: Optional[str] = Field(
        default=None,
        description="A user-supplied object ID. If specified, the ID must be unique among all pages and page elements in the presentation.",
    )
    insertionIndex: Optional[int] = Field(
        default=None,
        description="The optional zero-based index indicating where to insert the slides. If you don't specify an index, the new slide is created at the end.",
    )
    slideLayoutReference: Optional[LayoutReference] = Field(
        default=None,
        description="Layout reference of the slide to be inserted, based on the current master, which is one of the following: - The master of the previous slide index. - The master of the first slide, if the insertion_index is zero. - The first master in the presentation, if there are no slides.",
    )
    placeholderIdMappings: Optional[List[PlaceholderIdMapping]] = Field(
        default=None,
        description="An optional list of object ID mappings from the placeholder(s) on the layout to the placeholder(s) that will be created on the new slide from that specified layout. Can only be used when slideLayoutReference is specified.",
    )


class UpdateSlidePropertiesRequest(GslidesAPIRequest):
    """Updates the properties of a Slide.

    This request updates the slide properties for the specified slide.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateslidepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the slide to update")
    slideProperties: Dict[str, Any] = Field(description="The slide properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'slideProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class UpdateSlidesPositionRequest(GslidesAPIRequest):
    """Updates the position of slides in the presentation.

    This request moves slides to a new position in the presentation.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateslidespositionrequest
    """

    slideObjectIds: List[str] = Field(
        description="The IDs of the slides in the presentation that should be moved. The slides in this list must be in the same order as they appear in the presentation."
    )
    insertionIndex: int = Field(
        description="The index where the slides should be inserted, based on the slide arrangement before the move takes place. Must be between zero and the number of slides in the presentation, inclusive."
    )


class UpdatePagePropertiesRequest(GslidesAPIRequest):
    """Updates the properties of a Page.

    This request updates the page properties for the specified page.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatepagepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the page to update")
    pageProperties: Dict[str, Any] = Field(description="The page properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'pageProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class DeleteObjectRequest(GslidesAPIRequest):
    """Deletes an object, either a page or page element, from the presentation.

    This request deletes the specified object from the presentation. If the object is a page,
    its page elements are also deleted. If the object is a page element, it is removed from its page.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#deleteobjectrequest
    """

    objectId: str = Field(description="The object ID of the page or page element to delete")


class DuplicateObjectRequest(GslidesAPIRequest):
    """Duplicates a slide or page element.

    This request duplicates the specified slide or page element. When duplicating a slide,
    the duplicate slide will be created immediately following the specified slide.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#duplicateobjectrequest
    """

    objectId: str = Field(description="The ID of the object to duplicate")
    objectIds: Optional[Dict[str, str]] = Field(
        default=None,
        description="The object being duplicated may contain other objects, for example when duplicating a slide or a group page element. This map defines how the IDs of duplicated objects are generated: the keys are the IDs of the original objects and its values are the IDs that will be assigned to the corresponding duplicate object.",
    )


class UpdateImagePropertiesRequest(GslidesAPIRequest):
    """Updates the properties of an Image.

    This request updates the image properties for the specified image.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateimagepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the image to update")
    imageProperties: ImageProperties = Field(description="The image properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'imageProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class UpdatePageElementAltTextRequest(GslidesAPIRequest):
    """Updates the alt text title and/or description of a page element.

    This request updates the alternative text (alt text) for accessibility purposes
    on page elements like images, shapes, and other visual elements. The alt text
    is exposed to screen readers and other accessibility interfaces.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#UpdatePageElementAltTextRequest
    """

    objectId: str = Field(
        description="The object ID of the page element the updates are applied to"
    )
    title: Optional[str] = Field(
        default=None,
        description="The updated alt text title of the page element. If unset the existing value will be maintained. The title is exposed to screen readers and other accessibility interfaces. Only use human readable values related to the content of the page element.",
    )
    description: Optional[str] = Field(
        default=None,
        description="The updated alt text description of the page element. If unset the existing value will be maintained. The description is exposed to screen readers and other accessibility interfaces. Only use human readable values related to the content of the page element.",
    )


class UpdateVideoPropertiesRequest(GslidesAPIRequest):
    """Updates the properties of a Video.

    This request updates the video properties for the specified video.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatevideopropertiesrequest
    """

    objectId: str = Field(description="The object ID of the video to update")
    videoProperties: Dict[str, Any] = Field(description="The video properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'videoProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class UpdateLinePropertiesRequest(GslidesAPIRequest):
    """Updates the properties of a Line.

    This request updates the line properties for the specified line.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatelinepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the line to update")
    lineProperties: Dict[str, Any] = Field(description="The line properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'lineProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class UpdateSheetsChartPropertiesRequest(GslidesAPIRequest):
    """Updates the properties of a SheetsChart.

    This request updates the sheets chart properties for the specified chart.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatesheetschartpropertiesrequest
    """

    objectId: str = Field(description="The object ID of the sheets chart to update")
    sheetsChartProperties: Dict[str, Any] = Field(description="The sheets chart properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'sheetsChartProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )
