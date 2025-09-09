import os

from storyline.common.cached_call import cached_call
from storyline.slides.template import GoogleSlidesSettings, SlideDeck

from gslides_api import Presentation, initialize_credentials
from gslides_api.client import api_client

here = os.path.dirname(os.path.abspath(__file__))
credential_location = "/home/james/Dropbox/PyCharmProjects/gslides-playground/"
initialize_credentials(credential_location)

# Setup, choose presentation
presentation_id = "1YzTBZD8mR8N09Ti1JAdzodXfK_I8UaHt-3bZtQYDeJA"
json = cached_call(
    lambda: api_client.get_presentation_json(presentation_id), "gigs_presentation.json"
)

# Let's examine the structure of group objects in the JSON
import json as json_module

print("Let's look at a specific failing group object:")
# Look at slide 2, pageElement 2 which was mentioned in the error
if "slides" in json and len(json["slides"]) > 2:
    slide2 = json["slides"][2]
    if "pageElements" in slide2 and len(slide2["pageElements"]) > 2:
        element = slide2["pageElements"][2]
        print("Full element structure:")
        print(json_module.dumps(element, indent=2))
        print("\n" + "=" * 50 + "\n")
        if "group" in element:
            print("Group object structure:")
            print(json_module.dumps(element["group"], indent=2))

print("\nNow attempting to create Presentation object...")
try:
    p = Presentation.from_json(json)
    print("Success!")
except Exception as e:
    print(f"Failed with error: {e}")

    # Let's try to validate just a single group element to see the exact issue
    print("\nTrying to validate a single group element...")
    if "slides" in json and len(json["slides"]) > 2:
        slide2 = json["slides"][2]
        if "pageElements" in slide2 and len(slide2["pageElements"]) > 2:
            element = slide2["pageElements"][2]
            print("Attempting to validate this element:")
            print(json_module.dumps(element, indent=2))

            # Try to validate just this element
            from pydantic import TypeAdapter

            from gslides_api.element.element import GroupElement, PageElement

            try:
                page_element_adapter = TypeAdapter(PageElement)
                validated_element = page_element_adapter.validate_python(element)
                print("Single element validation succeeded!")
            except Exception as single_error:
                print(f"Single element validation failed: {single_error}")

                # Let's also check what the discriminator returns
                from gslides_api.element.element import element_discriminator

                discriminator_result = element_discriminator(element)
                print(f"Discriminator returned: {discriminator_result}")

                # Let's try to validate directly as GroupElement
                print("\nTrying to validate directly as GroupElement:")
                try:
                    group_element = GroupElement.model_validate(element)
                    print("Direct GroupElement validation succeeded!")
                except Exception as group_element_error:
                    print(f"Direct GroupElement validation failed: {group_element_error}")

                # Let's check the Group model specifically
                if "elementGroup" in element:
                    print("\nTrying to validate just the elementGroup:")
                    from gslides_api.domain.domain import Group

                    try:
                        group = Group.model_validate(element["elementGroup"])
                        print("Group validation succeeded!")
                    except Exception as group_error:
                        print(f"Group validation failed: {group_error}")

                # Let's also try adding a size field to see if that fixes it
                print("\nTrying with a dummy size field added:")
                element_with_size = element.copy()
                element_with_size["size"] = {"width": 100, "height": 100}
                try:
                    validated_with_size = page_element_adapter.validate_python(element_with_size)
                    print("Validation with dummy size succeeded!")
                except Exception as size_error:
                    print(f"Validation with dummy size failed: {size_error}")
