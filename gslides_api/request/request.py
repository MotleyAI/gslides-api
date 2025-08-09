from typing import Any, Dict, List, Optional, TYPE_CHECKING

from pydantic import Field

from gslides_api.domain import (
    BulletGlyphPreset,
    GSlidesBaseModel,
    LayoutReference,
    ImageProperties,
    PageElementProperties,
    VideoProperties,
    LineProperties,
    AffineTransform,
)
from gslides_api import ShapeProperties


from gslides_api.text import Type, TextStyle, ParagraphStyle
from gslides_api.request.domain import (
    Range,
    SubstringMatchCriteria,
    TableCellLocation,
    PlaceholderIdMapping,
)

if TYPE_CHECKING:
    from gslides_api.page.slide_properties import SlideProperties
    from gslides_api.page.base import PageProperties


class GSlidesAPIRequest(GSlidesBaseModel):
    """Base class for all requests to the Google Slides API."""

    def to_request(self) -> List[Dict[str, Any]]:
        """Convert to the format expected by the Google Slides API."""
        request_name = self.__class__.__name__.replace("Request", "")
        # make first letter lowercase
        request_name = request_name[0].lower() + request_name[1:]

        return [{request_name: self.to_api_format()}]


class CreateParagraphBulletsRequest(GSlidesAPIRequest):
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


class InsertTextRequest(GSlidesAPIRequest):
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


class UpdateTextStyleRequest(GSlidesAPIRequest):
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


class DeleteTextRequest(GSlidesAPIRequest):
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


class CreateShapeRequest(GSlidesAPIRequest):
    """Creates a new shape.

    This request creates a new shape on the specified page. The shape can be of various types
    like text box, rectangle, ellipse, etc.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createshaperequest
    """

    objectId: Optional[str] = Field(
        default=None,
        description="A user-supplied object ID. If specified, the ID must be unique among all pages and page elements in the presentation.",
    )
    elementProperties: PageElementProperties = Field(
        description="The element properties for the shape"
    )
    shapeType: Type = Field(description="The shape type")


class UpdateShapePropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Shape.

    This request updates the shape properties for the specified shape.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateshapepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the shape to update")
    shapeProperties: ShapeProperties = Field(description="The shape properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'shapeProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class ReplaceImageRequest(GSlidesAPIRequest):
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


class CreateSlideRequest(GSlidesAPIRequest):
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


class UpdateSlidePropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Slide.

    This request updates the slide properties for the specified slide.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateslidepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the slide to update")
    slideProperties: "SlideProperties" = Field(description="The slide properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'slideProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class UpdateSlidesPositionRequest(GSlidesAPIRequest):
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


class UpdatePagePropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Page.

    This request updates the page properties for the specified page.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatepagepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the page to update")
    pageProperties: "PageProperties" = Field(description="The page properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'pageProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class DeleteObjectRequest(GSlidesAPIRequest):
    """Deletes an object, either a page or page element, from the presentation.

    This request deletes the specified object from the presentation. If the object is a page,
    its page elements are also deleted. If the object is a page element, it is removed from its page.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#deleteobjectrequest
    """

    objectId: str = Field(description="The object ID of the page or page element to delete")


class DuplicateObjectRequest(GSlidesAPIRequest):
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


class UpdateImagePropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of an Image.

    This request updates the image properties for the specified image.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateimagepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the image to update")
    imageProperties: ImageProperties = Field(description="The image properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'imageProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class UpdatePageElementAltTextRequest(GSlidesAPIRequest):
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


class UpdateVideoPropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Video.

    This request updates the video properties for the specified video.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatevideopropertiesrequest
    """

    objectId: str = Field(description="The object ID of the video to update")
    videoProperties: VideoProperties = Field(description="The video properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'videoProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


class UpdateLinePropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Line.

    This request updates the line properties for the specified line.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatelinepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the line to update")
    lineProperties: LineProperties = Field(description="The line properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'lineProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )


#  This one appears to have been hallucinated, no trace of it in the actual docs
# class UpdateSheetsChartPropertiesRequest(GSlidesAPIRequest):
#     """Updates the properties of a SheetsChart.
#
#     This request updates the sheets chart properties for the specified chart.
#
#     Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatesheetschartpropertiesrequest
#     """
#
#     objectId: str = Field(description="The object ID of the sheets chart to update")
#     sheetsChartProperties: Dict[str, Any] = Field(
#         description="The sheets chart properties to update"
#     )
#     fields: str = Field(
#         description="The fields that should be updated. At least one field must be specified. The root 'sheetsChartProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
#     )
#


class CreateImageRequest(GSlidesAPIRequest):
    """Creates a new image.

    This request creates a new image on the specified page from a URL.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createimagerequest
    """

    objectId: Optional[str] = Field(
        default=None,
        description="A user-supplied object ID. If specified, the ID must be unique among all pages and page elements in the presentation.",
    )
    elementProperties: PageElementProperties = Field(
        description="The element properties for the image"
    )
    url: str = Field(description="The image URL")


class CreateVideoRequest(GSlidesAPIRequest):
    """Creates a new video.

    This request creates a new video on the specified page.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createvideorequest
    """

    elementProperties: PageElementProperties = Field(
        description="The element properties for the video"
    )
    source: str = Field(description="The video source type (e.g., 'YOUTUBE')")
    id: str = Field(description="The video ID")


class CreateLineRequest(GSlidesAPIRequest):
    """Creates a new line.

    This request creates a new line on the specified page.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createlinerequest
    """

    elementProperties: PageElementProperties = Field(
        description="The element properties for the line"
    )
    lineCategory: str = Field(description="The line category (e.g., 'STRAIGHT')")


# This seems to have been hallucinated
# class CreateWordArtRequest(GSlidesAPIRequest):
#     """Creates a new word art.
#
#     This request creates a new word art on the specified page.
#
#     Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createwordartrequest
#     """
#
#     elementProperties: Dict[str, Any] = Field(description="The element properties for the word art")
#     renderedText: str = Field(description="The text to render as word art")


class CreateSheetsChartRequest(GSlidesAPIRequest):
    """Creates a new sheets chart.

    This request creates a new sheets chart on the specified page.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createsheetschartrequest
    """

    elementProperties: PageElementProperties = Field(
        description="The element properties for the sheets chart"
    )
    spreadsheetId: str = Field(
        description="The ID of the Google Sheets spreadsheet that contains the chart"
    )
    chartId: int = Field(description="The ID of the specific chart in the spreadsheet")


class ReplaceAllTextRequest(GSlidesAPIRequest):
    """Replaces all instances of text matching some criteria with replace text.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#replacealltextrequest
    """

    containsText: SubstringMatchCriteria = Field(
        description="Finds all instances of text matching this substring"
    )
    replaceText: str = Field(description="The text that will replace the matched text")
    pageObjectIds: Optional[List[str]] = Field(
        default=None,
        description="If non-empty, limits the matches to page elements only on the given pages",
    )


class UpdatePageElementTransformRequest(GSlidesAPIRequest):
    """Updates the transform of a page element.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatepageelementtransformrequest
    """

    objectId: str = Field(description="The object ID of the page element to update")
    transform: AffineTransform = Field(
        description="The input transform matrix used to update the page element"
    )
    applyMode: str = Field(description="The apply mode of the transform update")


class RefreshSheetsChartRequest(GSlidesAPIRequest):
    """Refreshes an embedded Google Sheets chart by replacing it with the latest version of the chart from Google Sheets.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#refreshsheetschartrequest
    """

    objectId: str = Field(description="The object ID of the chart to refresh")


class ReplaceAllShapesWithImageRequest(GSlidesAPIRequest):
    """Replaces all shapes that match the given criteria with the provided image.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#replaceallshapeswithimagerequest
    """

    containsText: SubstringMatchCriteria = Field(
        description="If set, this request will replace all of the shapes that contain the given text"
    )
    imageUrl: Optional[str] = Field(default=None, description="The image URL")
    imageReplaceMethod: Optional[str] = Field(default=None, description="The image replace method")
    pageObjectIds: Optional[List[str]] = Field(
        default=None,
        description="If non-empty, limits the matches to page elements only on the given pages",
    )


class ReplaceAllShapesWithSheetsChartRequest(GSlidesAPIRequest):
    """Replaces all shapes that match the given criteria with the provided Google Sheets chart.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#replaceallshapeswithsheetschartrequest
    """

    containsText: SubstringMatchCriteria = Field(
        description="The criteria that the shapes must match in order to be replaced"
    )
    spreadsheetId: str = Field(
        description="The ID of the Google Sheets spreadsheet that contains the chart"
    )
    chartId: int = Field(
        description="The ID of the specific chart in the Google Sheets spreadsheet"
    )
    linkingMode: Optional[str] = Field(
        default=None,
        description="The mode with which the chart is linked to the source spreadsheet",
    )
    pageObjectIds: Optional[List[str]] = Field(
        default=None,
        description="If non-empty, limits the matches to page elements only on the given pages",
    )


class DeleteParagraphBulletsRequest(GSlidesAPIRequest):
    """Deletes bullets from all of the paragraphs that overlap with the given text index range.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#deleteparagraphbulletsrequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table containing the text to delete bullets from"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The optional table cell location if the text to be modified is in a table cell",
    )
    textRange: Range = Field(description="The range of text to delete bullets from")


class UpdateParagraphStyleRequest(GSlidesAPIRequest):
    """Updates the styling for all of the paragraphs within a Shape or Table that overlap with the given text index range.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateparagraphstylerequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table with the text to be styled"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The location of the cell in the table containing the paragraph(s) to style",
    )
    style: ParagraphStyle = Field(description="The paragraph's style")
    textRange: Range = Field(description="The range of text containing the paragraph(s) to style")
    fields: str = Field(description="The fields that should be updated")


class GroupObjectsRequest(GSlidesAPIRequest):
    """Groups objects to create an object group.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#groupobjectsrequest
    """

    groupObjectId: Optional[str] = Field(
        default=None, description="A user-supplied object ID for the group to be created"
    )
    childrenObjectIds: List[str] = Field(description="The object IDs of the objects to group")


class UngroupObjectsRequest(GSlidesAPIRequest):
    """Ungroups objects, such as groups.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#ungroupobjectsrequest
    """

    objectIds: List[str] = Field(description="The object IDs of the objects to ungroup")


class UpdatePageElementsZOrderRequest(GSlidesAPIRequest):
    """Updates the Z-order of page elements.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatepageelementsZorderrequest
    """

    pageElementObjectIds: List[str] = Field(
        description="The object IDs of the page elements to update"
    )
    operation: str = Field(description="The Z-order operation to apply on the page elements")


class UpdateLineCategoryRequest(GSlidesAPIRequest):
    """Updates the category of a line.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatelinecategoryrequest
    """

    objectId: str = Field(description="The object ID of the line the update is applied to")
    lineCategory: str = Field(description="The line category to update to")


class RerouteLineRequest(GSlidesAPIRequest):
    """Reroutes a line such that it's connected at the two closest connection sites on the connected page elements.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#reroutelinerequest
    """

    objectId: str = Field(description="The object ID of the line to reroute")
