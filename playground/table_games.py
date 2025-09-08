import json
import os
import logging

from langchain_core.language_models import BaseLanguageModel

from gslides_api import GoogleAPIClient, Presentation, initialize_credentials
from gslides_api.client import api_client
from motleycrew.common import LLMFramework
from motleycrew.common.llms import init_llm
from storyline.common.cached_call import cached_call
from storyline.common.logging import configure_logging
from storyline.config import LoggingParams
from storyline.domain.chart_image_to_config import image_to_config
from storyline.domain.layout import ChartBlock

from storyline.slides.ingest_presentation import name_slides, delete_alt_titles, ingest_presentation
from storyline.slides.slide_deck import SlideDeck

logger = logging.getLogger(__name__)

configure_logging(logging_params=LoggingParams(level="INFO"))

llm = init_llm(llm_framework=LLMFramework.LANGCHAIN, llm_name="gpt-5")

api_client.auto_flush = True

credential_location = os.getenv("GSLIDES_CREDENTIALS_PATH")
initialize_credentials(credential_location)

copy = False
# # This is a copy of the Noosa deck
# presentation_id = "1Y5RNVGxTJh0ocUbGId4tX6_u7CTPP2NpigH8hTsjaYY"
# This is a test of table ingestion
presentation_id = "1U_8TmamiNcb43eCC4teagrzs93h-oYiCXV4l5WyULiY"

p = Presentation.from_id(presentation_id, api_client=api_client)

slide = p.get_slide_by_name("Slide 4")
table = slide.get_elements_by_alt_title("Table_1")[0]
cell = table[0, 0]
print(cell.read_text())
table.write_text_to_cell(text="Hello world!", location=(0, 0))
api_client.flush_batch_update()

table.resize(n_rows=5, n_columns=4)
table.resize(n_rows=3, n_columns=2)
table.resize(n_rows=4, n_columns=4)
print("yay!")
