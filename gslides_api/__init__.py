from ._version import __version__, __version_info__
from .client import GoogleAPIClient, initialize_credentials
from .domain import (AffineTransform, BlurRadius, Color, ColorScheme,
                     ColorStop, CropProperties, DashStyle, Dimension, Group,
                     ImageProperties, Line, LineProperties, Outline,
                     OutlineFill, PageBackgroundFill, PredefinedLayout,
                     PropertyState, Recolor, RecolorName, RectanglePosition,
                     RgbColor, Shadow, ShadowTransform, ShadowType,
                     ShapeBackgroundFill, SheetsChart, SheetsChartProperties,
                     Size, SolidFill, SpeakerSpotlight,
                     SpeakerSpotlightProperties, StretchedPictureFill,
                     ThemeColorPair, ThemeColorType, Unit, Video,
                     VideoProperties, VideoSourceType, Weight, WordArt)
from .element.base import ElementKind
from .page.base import BasePage, PageProperties, PageType
from .page.notes import Notes, NotesProperties
from .page.page import (Layout, LayoutProperties, Master, MasterProperties,
                        NotesMaster, Page)
from .page.slide import Slide
from .page.slide_properties import SlideProperties
from .presentation import Presentation
from .request.domain import Range, RangeType, TableCellLocation
from .request.request import (CreateParagraphBulletsRequest,
                              CreateShapeRequest, CreateSlideRequest,
                              DeleteObjectRequest, DeleteTextRequest,
                              DuplicateObjectRequest, InsertTextRequest,
                              ReplaceImageRequest,
                              UpdateImagePropertiesRequest,
                              UpdatePagePropertiesRequest,
                              UpdateShapePropertiesRequest,
                              UpdateSlidePropertiesRequest,
                              UpdateSlidesPositionRequest,
                              UpdateTextStyleRequest)
from .text import AutoText, AutoTextType, ShapeProperties, TextElement

__all__ = [
    # Version info
    "__version__",
    "__version_info__",
    "GoogleAPIClient",
    # Domain objects
    "Size",
    "Dimension",
    "TextElement",
    "Video",
    "VideoProperties",
    "VideoSourceType",
    "RgbColor",
    "Color",
    "ThemeColorType",
    "SolidFill",
    "ShapeBackgroundFill",
    "OutlineFill",
    "Weight",
    "Outline",
    "DashStyle",
    "ShadowTransform",
    "BlurRadius",
    "Shadow",
    "ShadowType",
    "RectanglePosition",
    "ShapeProperties",
    "CropProperties",
    "ColorStop",
    "RecolorName",
    "Recolor",
    "ImageProperties",
    "PropertyState",
    "StretchedPictureFill",
    "PageBackgroundFill",
    "AutoText",
    "AutoTextType",
    "PredefinedLayout",
    "ColorScheme",
    "ThemeColorPair",
    "Line",
    "LineProperties",
    "WordArt",
    "SheetsChart",
    "SheetsChartProperties",
    "SpeakerSpotlight",
    "SpeakerSpotlightProperties",
    "Group",
    "Unit",
    "AffineTransform",
    # Presentation
    "Presentation",
    # Pages
    "Layout",
    "LayoutProperties",
    "Master",
    "MasterProperties",
    "NotesMaster",
    "Page",
    "Notes",
    "NotesProperties",
    "Slide",
    "SlideProperties",
    "BasePage",
    "PageProperties",
    "PageType",
    # Elements
    "ElementKind",
    # Client
    "initialize_credentials",
    # Requests
    "CreateParagraphBulletsRequest",
    "InsertTextRequest",
    "UpdateTextStyleRequest",
    "DeleteTextRequest",
    "CreateShapeRequest",
    "UpdateShapePropertiesRequest",
    "UpdateImagePropertiesRequest",
    "ReplaceImageRequest",
    "CreateSlideRequest",
    "UpdateSlidePropertiesRequest",
    "UpdateSlidesPositionRequest",
    "UpdatePagePropertiesRequest",
    "DeleteObjectRequest",
    "DuplicateObjectRequest",
    "Range",
    "RangeType",
    "TableCellLocation",
]

# Rebuild models to resolve forward references after all imports
UpdateSlidePropertiesRequest.model_rebuild()
UpdatePagePropertiesRequest.model_rebuild()
UpdateShapePropertiesRequest.model_rebuild()
UpdateImagePropertiesRequest.model_rebuild()
