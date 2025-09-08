"""Test ItemList flattening functionality."""

import pytest

from gslides_api.text import TextElement
from gslides_api.markdown.from_markdown import BulletPointGroup, ItemList, NumberedListGroup
from gslides_api.text import TextRun, TextStyle


def create_text_element(content: str, start_index: int = 0) -> TextElement:
    """Helper function to create a TextElement."""
    return TextElement(
        startIndex=start_index,
        endIndex=start_index + len(content),
        textRun=TextRun(content=content, style=TextStyle()),
    )


def test_itemlist_with_only_text_elements():
    """Test ItemList with only TextElement objects (no flattening needed)."""
    text1 = create_text_element("Hello", 0)
    text2 = create_text_element("World", 5)

    item_list = ItemList(children=[text1, text2])

    assert len(item_list.children) == 2
    assert item_list.children[0] == text1
    assert item_list.children[1] == text2


def test_itemlist_with_nested_itemlists():
    """Test ItemList flattening when nested ItemLists are provided."""
    # Create some text elements
    text1 = create_text_element("First", 0)
    text2 = create_text_element("Second", 5)
    text3 = create_text_element("Third", 11)
    text4 = create_text_element("Fourth", 16)

    # Create nested ItemLists
    nested_list1 = ItemList(children=[text1, text2])
    nested_list2 = ItemList(children=[text3, text4])

    # Create parent ItemList with mixed content
    parent_list = ItemList(children=[nested_list1, nested_list2])

    # Should be flattened to contain all 4 text elements
    assert len(parent_list.children) == 4
    assert parent_list.children[0] == text1
    assert parent_list.children[1] == text2
    assert parent_list.children[2] == text3
    assert parent_list.children[3] == text4


def test_itemlist_with_mixed_content():
    """Test ItemList with mixed TextElement and ItemList objects."""
    text1 = create_text_element("Direct", 0)
    text2 = create_text_element("Nested1", 6)
    text3 = create_text_element("Nested2", 13)
    text4 = create_text_element("Direct2", 20)

    # Create a nested ItemList
    nested_list = ItemList(children=[text2, text3])

    # Create parent with mixed content
    parent_list = ItemList(children=[text1, nested_list, text4])

    # Should be flattened
    assert len(parent_list.children) == 4
    assert parent_list.children[0] == text1
    assert parent_list.children[1] == text2
    assert parent_list.children[2] == text3
    assert parent_list.children[3] == text4


def test_bulletpointgroup_inherits_flattening():
    """Test that BulletPointGroup inherits the flattening behavior."""
    text1 = create_text_element("Bullet1", 0)
    text2 = create_text_element("Bullet2", 7)
    text3 = create_text_element("Bullet3", 14)

    nested_list = ItemList(children=[text2, text3])
    bullet_group = BulletPointGroup(children=[text1, nested_list])

    assert len(bullet_group.children) == 3
    assert bullet_group.children[0] == text1
    assert bullet_group.children[1] == text2
    assert bullet_group.children[2] == text3


def test_numberedlistgroup_inherits_flattening():
    """Test that NumberedListGroup inherits the flattening behavior."""
    text1 = create_text_element("Item1", 0)
    text2 = create_text_element("Item2", 5)
    text3 = create_text_element("Item3", 10)

    nested_list = ItemList(children=[text2, text3])
    numbered_group = NumberedListGroup(children=[text1, nested_list])

    assert len(numbered_group.children) == 3
    assert numbered_group.children[0] == text1
    assert numbered_group.children[1] == text2
    assert numbered_group.children[2] == text3


def test_empty_itemlist():
    """Test ItemList with empty children list."""
    item_list = ItemList(children=[])
    assert len(item_list.children) == 0


def test_deeply_nested_itemlists():
    """Test deeply nested ItemLists are properly flattened."""
    text1 = create_text_element("Level1", 0)
    text2 = create_text_element("Level2", 6)
    text3 = create_text_element("Level3", 12)

    # Create deeply nested structure
    level3_list = ItemList(children=[text3])
    level2_list = ItemList(children=[text2, level3_list])
    level1_list = ItemList(children=[text1, level2_list])

    # Should flatten to all text elements
    assert len(level1_list.children) == 3
    assert level1_list.children[0] == text1
    assert level1_list.children[1] == text2
    assert level1_list.children[2] == text3


def test_itemlist_properties_work_after_flattening():
    """Test that start_index and end_index properties work correctly after flattening."""
    text1 = create_text_element("Hello", 0)
    text2 = create_text_element("World", 10)

    nested_list = ItemList(children=[text2])
    parent_list = ItemList(children=[text1, nested_list])

    assert parent_list.start_index == 0
    assert parent_list.end_index == 15  # "World" ends at 10 + 5 = 15
