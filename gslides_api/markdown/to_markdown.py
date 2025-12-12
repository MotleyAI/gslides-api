# This file previously contained text_elements_to_markdown() and related functions.
# These have been replaced by the platform-agnostic IR system:
#
# - text_elements_to_ir() in gslides_api/agnostic/converters.py
# - ir_to_markdown() in gslides_api/agnostic/ir_to_markdown.py
#
# The new system provides better run consolidation (prevents "doubled asterisks" bugs)
# and cleaner separation of concerns via the intermediate representation (IR).
