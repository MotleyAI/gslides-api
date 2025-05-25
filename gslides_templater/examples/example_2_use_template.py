"""
Example 2: Creating presentation from ready template with data filling

IMPROVED VERSION with better image handling:
- More reliable image URLs
- Better error handling for images
- Fallback strategies for failed images
- Enhanced diagnostics
"""

import sys
import os
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gslides_templater import (
    create_templater,
    AuthConfig,
    SlidesAPIError,
    AuthenticationError,
    TemplateValidationError,
    MarkdownProcessingError
)


def get_reliable_image_urls(count: int) -> list:
    """Get a list of reliable, publicly accessible image URLs"""

    # Using placeholder.com for reliable test images
    reliable_images = []

    # Different colors and sizes for variety
    colors = ['4285f4', 'ea4335', '34a853', 'fbbc04', '9aa0a6', 'f28b82', 'aecbfa', 'fdd663']
    sizes = ['800x600', '600x400', '400x300', '500x500']

    for i in range(count):
        color = colors[i % len(colors)]
        size = sizes[i % len(sizes)]
        text = f"Image+{i + 1}"
        url = f"https://via.placeholder.com/{size}/{color}/ffffff?text={text}"
        reliable_images.append(url)

    return reliable_images


def main():
    """Create presentation from ready template with NEW data"""

    print("🎯 Creating presentation from template with NEW data")
    print("=" * 55)

    try:
        print("🔑 Setting up authentication...")

        credentials_file = "credentials.json"
        if not os.path.exists(credentials_file):
            credentials_file = "credentials.json"

        if not os.path.exists(credentials_file):
            print("❌ File credentials.json not found!")
            print("   Place it next to the script or in parent folder")
            return

        # Create auth config using new structure
        auth_config = AuthConfig(
            credentials_path=credentials_file,
            token_path="token.json"
        )

        templater = create_templater(auth_config=auth_config)
        print("   ✓ Authentication successful")

        template_filename = "presentation_template.json"
        print(f"\n📁 Loading template from file: {template_filename}")

        if not os.path.exists(template_filename):
            print(f"❌ Template file not found: {template_filename}")
            print("   First run example_1_create_template.py")
            return

        try:
            template_config = templater.load_template(template_filename)
        except Exception as e:
            print(f"❌ Error loading template: {e}")
            print("   Template file may be corrupted")
            return

        print(f"   ✓ Template loaded: {template_config['name']}")
        print(f"   ✓ Source presentation: {template_config['source_presentation_id']}")
        print(f"   ✓ Placeholders in template: {len(template_config.get('placeholders', {}))}")

        # Show template structure info
        slide_size = template_config.get('slide_size', {})
        if slide_size:
            print(f"   ✓ Slide size: {slide_size.get('width', 720)}x{slide_size.get('height', 540)} points")

        layout_config = template_config.get('layout_config', {})
        if layout_config:
            print(f"   ✓ Layout margins: {layout_config.get('margin_x', 50)}x{layout_config.get('margin_y', 50)}")

        placeholders = template_config.get('placeholders', {})
        if not placeholders:
            print("❌ No placeholders in template!")
            return

        text_placeholders = {name: info for name, info in placeholders.items()
                             if info['type'] == 'text'}
        image_placeholders = {name: info for name, info in placeholders.items()
                              if info['type'] == 'image'}
        other_placeholders = {name: info for name, info in placeholders.items()
                              if info['type'] not in ['text', 'image']}

        print(f"\n📝 Found placeholders:")
        print(f"   📝 Text: {len(text_placeholders)}")
        print(f"   🖼️ Images: {len(image_placeholders)}")
        print(f"   📦 Other: {len(other_placeholders)}")

        print(f"\n   Placeholder examples:")
        for i, (name, info) in enumerate(list(placeholders.items())[:5]):
            position = info.get('position', {})
            if position:
                print(f"   • {name} ({info['type']}) - slide {info.get('slide_index', 0) + 1} "
                      f"[{position.get('x', 0):.0f},{position.get('y', 0):.0f}] "
                      f"layer:{info.get('layer', 0)}")
            else:
                print(f"   • {name} ({info['type']}) - slide {info.get('slide_index', 0) + 1}")
        if len(placeholders) > 5:
            print(f"   ... and {len(placeholders) - 5} more")

        print(f"\n📊 Creating NEW data for filling...")

        new_data = {}

        # Enhanced text data with better Markdown formatting
        text_data_examples = {
            "name_here": "# **Ekaterina Volkova** 👑\n\n*Chief Innovation Officer*",
            "baby_album": "### 🎈 *Family Moments 2024*\n\n> Creating memories one day at a time",
            "bath_time": """**Fun bath time!** 🛁\n\n- Splashing around\n- Playing with toys\n- *Pure joy*""",
            "look_at": """#### 👀 **Look at us!**\n\n***Adventures await***""",
            "slide_2_text_15": """🍕 **First cafe in life!** \n\n***Indescribable taste!***\n\n`#FirstTimeCafe`""",
            "boys_love": """#### 🚂 **Boys and their hobbies!** 🔬\n\n- Building trains\n- Science experiments\n- *Endless curiosity*""",
            "slide_5_text_36": """##### ⭐ **Our little commander!** \n\n> "Leadership from diapers" 😎\n\n**Skills:**\n- Strategic thinking\n- Team coordination\n- *Natural charisma*""",
            "slide_3_text": """**Games and entertainment** 🎮\n\n- Video games\n- Board games\n- Outdoor activities\n\n*Fun for everyone!*""",
            "slide_4_text": """### Walk in the park 🌳 \n\n`Healthy sleep` guaranteed! 😴\n\n**Benefits:**\n- Fresh air\n- Exercise\n- ~~Tired kids~~ Happy kids"""
        }

        for placeholder_name in text_placeholders.keys():
            if placeholder_name in text_data_examples:
                new_data[placeholder_name] = text_data_examples[placeholder_name]
                print(f"   📝 {placeholder_name}: personalized content")
            else:
                slide_num = text_placeholders[placeholder_name].get('slide_index', 0) + 1
                new_data[
                    placeholder_name] = f"""**New content for slide {slide_num}**\n\n- Enhanced with **bold** text\n- *Italic* formatting\n- `Code style` elements\n\n> This is a quote for better presentation"""
                print(f"   📝 {placeholder_name}: generated Markdown (slide {slide_num})")

        # Get reliable images
        image_count = len(image_placeholders)
        reliable_images = get_reliable_image_urls(image_count)

        print(f"\n🖼️ Using reliable placeholder images...")

        image_counter = 0
        for placeholder_name in image_placeholders.keys():
            if image_counter < len(reliable_images):
                new_data[placeholder_name] = reliable_images[image_counter]
                print(f"   🖼️ {placeholder_name}: reliable image #{image_counter + 1}")
            else:
                # Fallback to a simple placeholder
                new_data[
                    placeholder_name] = f"https://via.placeholder.com/600x400/cccccc/000000?text=Image+{image_counter + 1}"
                print(f"   🖼️ {placeholder_name}: fallback image")
            image_counter += 1

        for placeholder_name in other_placeholders.keys():
            placeholder_type = other_placeholders[placeholder_name]['type']
            new_data[placeholder_name] = f"New data for {placeholder_type} element"
            print(f"   📦 {placeholder_name}: data for {placeholder_type}")

        print(f"\n   ✅ Total data prepared: {len(new_data)}")
        print(f"   📝 Text elements: {len(text_placeholders)}")
        print(f"   🖼️ Image elements: {len(image_placeholders)}")
        print(f"   📦 Other elements: {len(other_placeholders)}")

        print(f"\n📋 Examples of new data:")
        for i, (name, value) in enumerate(list(new_data.items())[:3]):
            if isinstance(value, str):
                if value.startswith("http"):
                    print(f"   🖼️ {name}: {value}")
                else:
                    preview = value[:80] + '...' if len(value) > 80 else value
                    preview = preview.replace('\n', ' ↵ ')
                    print(f"   📝 {name}: {preview}")
        if len(new_data) > 3:
            print(f"   ... and {len(new_data) - 3} more elements")

        # Validate template data before applying
        print(f"\n🔍 Validating template data...")
        try:
            validation_result = templater.validate_template_data(template_config, new_data)
            if not validation_result['valid']:
                print(f"❌ Template validation failed:")
                if validation_result['missing_placeholders']:
                    print(f"   Missing: {validation_result['missing_placeholders']}")
                if validation_result['invalid_types']:
                    print(f"   Invalid types: {validation_result['invalid_types']}")
                return
            else:
                print(f"   ✅ All data is valid")
                if validation_result['warnings']:
                    print(f"   ⚠️ Warnings: {len(validation_result['warnings'])}")
        except Exception as e:
            print(f"⚠️ Validation error: {e}")

        print(f"\n🔄 Creating new presentation with data...")

        start_time = time.time()
        try:
            new_presentation_id = templater.apply_template(
                template=template_config,
                data=new_data,
                title=f"🎨 New presentation from template - {time.strftime('%d.%m.%Y %H:%M')}"
            )
        except TemplateValidationError as e:
            print(f"❌ Template validation error: {e}")
            return
        except Exception as e:
            print(f"❌ Error applying template: {e}")
            return

        end_time = time.time()

        print(f"   ✅ Presentation created in {end_time - start_time:.1f} seconds!")
        print(f"   ✅ ID: {new_presentation_id}")

        # Get info about created presentation
        try:
            new_presentation = templater.get_presentation(new_presentation_id)
            presentation_info = templater.get_presentation_info(new_presentation_id)
        except Exception as e:
            print(f"⚠️ Could not get presentation info: {e}")
            new_presentation = {"title": "Unknown"}
            presentation_info = {}

        print(f"\n📋 Information about created presentation:")
        print(f"   📄 Title: {new_presentation.get('title')}")
        print(f"   🆔 ID: {new_presentation_id}")
        print(f"   📊 Number of slides: {len(new_presentation.get('slides', []))}")
        print(f"   📦 Total elements: {presentation_info.get('total_elements', 'Unknown')}")

        # Element type breakdown
        element_types = presentation_info.get('element_types', {})
        if element_types:
            print(f"   📋 Element types:")
            for elem_type, count in element_types.items():
                print(f"      {elem_type}: {count}")

        print(f"\n🔄 Statistics of applied replacements:")
        replacement_stats = {
            'text_replaced': 0,
            'images_replaced': 0,
            'other_replaced': 0,
            'total_characters': 0,
            'markdown_elements': 0
        }

        for placeholder_name, value in new_data.items():
            placeholder_info = placeholders.get(placeholder_name, {})
            placeholder_type = placeholder_info.get('type', 'unknown')

            if placeholder_type == 'text':
                replacement_stats['text_replaced'] += 1
                replacement_stats['total_characters'] += len(str(value))
                # Count markdown elements
                if any(marker in str(value) for marker in ['**', '*', '#', '`', '>', '-', '~~']):
                    replacement_stats['markdown_elements'] += 1
            elif placeholder_type == 'image':
                replacement_stats['images_replaced'] += 1
            else:
                replacement_stats['other_replaced'] += 1

        print(f"   📝 Text elements replaced: {replacement_stats['text_replaced']}")
        print(f"   🖼️ Images replaced: {replacement_stats['images_replaced']}")
        print(f"   📦 Other elements replaced: {replacement_stats['other_replaced']}")
        print(f"   📊 Markdown characters processed: {replacement_stats['total_characters']:,}")
        print(f"   🎨 Elements with Markdown formatting: {replacement_stats['markdown_elements']}")
        print(f"   📄 Slides processed: {len(template_config.get('slides', []))}")

        presentation_url = templater.get_presentation_url(new_presentation_id)
        print(f"\n🌐 Link to new presentation:")
        print(f"   {presentation_url}")

        print(f"\n✅ PRESENTATION SUCCESSFULLY CREATED WITH NEW DATA!")
        print(f"🎨 All Markdown elements converted to Google Slides")
        print(f"📝 All styles preserved: bold, italic, headers, lists, quotes")
        print(f"🖼️ All images replaced with reliable placeholder images")
        print(f"📏 Layout and positioning information preserved")
        print(f"🚀 Ready for viewing and editing!")

        print(f"\n💡 What to do next:")
        print(f"   1. 🔗 Open the link above in browser")
        print(f"   2. 🎨 Check quality of replacements and formatting")
        print(f"   3. 🖼️ Replace placeholder images with your own if needed")
        print(f"   4. ✏️ Edit manually if needed")
        print(f"   5. 📤 Share presentation with colleagues")
        print(f"   6. 💾 Save as PDF if needed")

        print(f"\n🔧 Technical information:")
        print(f"   📅 Template creation date: {template_config.get('created_at', 'Unknown')}")
        print(f"   🆔 Source presentation ID: {template_config['source_presentation_id']}")
        print(f"   📝 Template name: {template_config['name']}")
        print(f"   ⏱️ Processing time: {end_time - start_time:.2f} seconds")

        # Show layout information
        if slide_size:
            print(f"   📐 Slide dimensions: {slide_size.get('width', 720)}x{slide_size.get('height', 540)} points")
        if layout_config:
            print(
                f"   📏 Layout margins: {layout_config.get('margin_x', 50)}x{layout_config.get('margin_y', 50)} points")

    except FileNotFoundError as e:
        file_name = str(e).split("'")[1] if "'" in str(e) else "file"
        print(f"\n❌ ERROR: File not found")

        if "presentation_template.json" in str(e):
            print(f"📁 Template file not found: presentation_template.json")
            print(f"🔧 Solution:")
            print(f"   1. Run example_1_create_template.py")
            print(f"   2. Create template from existing presentation")
            print(f"   3. Make sure presentation_template.json file is created")
        elif "credentials.json" in str(e):
            print(f"🔑 File credentials.json not found")
            print(f"🔧 Solution:")
            print(f"   1. Download OAuth credentials from Google Cloud Console")
            print(f"   2. Save as credentials.json")
            print(f"   3. Place next to script")
        else:
            print(f"📄 File not found: {file_name}")
            print(f"🔧 Check file paths")

    except AuthenticationError as e:
        print(f"\n❌ AUTHENTICATION ERROR: {e}")
        print(f"\n🔧 Authentication solutions:")
        print(f"   • Delete token.json and re-authorize")
        print(f"   • Check credentials.json file")
        print(f"   • Make sure OAuth Client ID is configured correctly")
        print(f"   • Verify Google Slides API is enabled")

    except TemplateValidationError as e:
        print(f"\n❌ TEMPLATE VALIDATION ERROR: {e}")
        print(f"\n🔧 Template solutions:")
        print(f"   • Check template file integrity")
        print(f"   • Recreate template with example_1_create_template.py")
        print(f"   • Verify data types match template requirements")

    except MarkdownProcessingError as e:
        print(f"\n❌ MARKDOWN PROCESSING ERROR: {e}")
        print(f"\n🔧 Markdown solutions:")
        print(f"   • Check Markdown syntax in your content")
        print(f"   • Simplify complex formatting")
        print(f"   • Verify unclosed tags (**, *, `, ~~)")

    except SlidesAPIError as e:
        error_msg = str(e)
        print(f"\n❌ SLIDES API ERROR: {error_msg}")

        if "image was not found" in error_msg.lower():
            print(f"\n🔧 DIAGNOSIS: Image access issue")
            print(f"   🎯 Reasons:")
            print(f"   • Image URL is not publicly accessible")
            print(f"   • Image server is down or blocking requests")
            print(f"   • Image URL format is incorrect")
            print(f"   🔧 Solutions:")
            print(f"   • Use reliable image hosting services")
            print(f"   • Check image URLs are publicly accessible")
            print(f"   • Try different image URLs")

        elif "404" in error_msg or "not found" in error_msg.lower():
            print(f"\n🔧 DIAGNOSIS: Resource not found")
            print(f"   🎯 Reasons:")
            print(f"   • Source presentation was deleted")
            print(f"   • Wrong presentation ID in template")
            print(f"   • No access to presentation")

        elif "403" in error_msg or "permission" in error_msg.lower():
            print(f"\n🔧 DIAGNOSIS: No access rights")
            print(f"   🎯 Reasons:")
            print(f"   • No rights to source presentation")
            print(f"   • Google Slides API not enabled")
            print(f"   • OAuth credentials issues")

        elif "Rate limit" in error_msg or "429" in error_msg:
            print(f"\n🔧 DIAGNOSIS: Rate limit exceeded")
            print(f"   🎯 Solutions:")
            print(f"   • Wait a few minutes and try again")
            print(f"   • Check API quotas in Google Cloud Console")

    except Exception as e:
        error_msg = str(e)
        print(f"\n❌ UNEXPECTED ERROR: {error_msg}")
        print(f"\n🔧 GENERAL DIAGNOSIS:")
        print(f"   🌐 Check internet connection")
        print(f"   🔄 Try again in a few minutes")
        print(f"   📊 May have exceeded API request limit")
        print(f"   🆕 Make sure you're using latest code version")


if __name__ == "__main__":
    main()