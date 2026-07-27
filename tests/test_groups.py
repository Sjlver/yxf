"""Unit tests for recognizing the rows that open and close groups.

XLSForm accepts several spellings for these rows, and the metadata type "end"
looks deceptively similar to the end of a group. Getting this wrong silently
corrupts the nesting depth for the rest of a sheet, so it is tested closely.
"""

import pytest

from yxf.xlsform import (
    is_group_begin,
    is_group_end,
    nesting_levels,
    parse_group_marker,
)


ALL_SPELLINGS = [
    ("begin group", ("begin", "group")),
    ("begin_group", ("begin", "group")),
    ("end group", ("end", "group")),
    ("end_group", ("end", "group")),
    ("begin repeat", ("begin", "repeat")),
    ("begin_repeat", ("begin", "repeat")),
    ("end repeat", ("end", "repeat")),
    ("end_repeat", ("end", "repeat")),
]

# Types that must never be mistaken for a group marker. "start" and "end" are
# metadata questions that record submission times.
NOT_GROUP_MARKERS = [
    "end",
    "start",
    "today",
    "text",
    "note",
    "integer",
    "calculate",
    "deviceid",
    "begin",
    "group",
    "repeat",
    "endgroup",
    "begingroup",
    "end_of_list",
    "beginner",
    "select_one begin_group",
    "text audit",
    "",
    None,
]


class TestParseGroupMarker:
    """Tests for parse_group_marker."""

    @pytest.mark.parametrize("type_value,expected", ALL_SPELLINGS)
    def test_all_spellings_are_recognized(self, type_value, expected):
        assert parse_group_marker(type_value) == expected

    @pytest.mark.parametrize("type_value", NOT_GROUP_MARKERS)
    def test_other_types_are_not_group_markers(self, type_value):
        assert parse_group_marker(type_value) is None

    def test_metadata_end_is_not_a_group_end(self):
        """The "end" metadata question must not close a group."""
        assert parse_group_marker("end") is None
        assert not is_group_end("end")

    @pytest.mark.parametrize(
        "type_value",
        [
            "BEGIN GROUP",
            "Begin_Group",
            "begin  group",
            "  begin group  ",
            "begin\tgroup",
        ],
    )
    def test_case_and_whitespace_are_ignored(self, type_value):
        assert parse_group_marker(type_value) == ("begin", "group")

    def test_non_string_values(self):
        assert parse_group_marker(None) is None
        assert parse_group_marker(42) is None


class TestGroupPredicates:
    """Tests for is_group_begin and is_group_end."""

    @pytest.mark.parametrize("type_value,expected", ALL_SPELLINGS)
    def test_predicates_agree_with_parse(self, type_value, expected):
        assert is_group_begin(type_value) == (expected[0] == "begin")
        assert is_group_end(type_value) == (expected[0] == "end")

    @pytest.mark.parametrize("type_value", NOT_GROUP_MARKERS)
    def test_other_types_are_neither(self, type_value):
        assert not is_group_begin(type_value)
        assert not is_group_end(type_value)


class TestNestingLevels:
    """Tests for nesting_levels."""

    def test_flat_sheet(self):
        assert nesting_levels(["text", "integer", "note"]) == [0, 0, 0]

    def test_single_group_includes_its_delimiters(self):
        """Both the begin and the end row belong to the group they delimit."""
        types = ["text", "begin_group", "text", "end_group", "text"]
        assert nesting_levels(types) == [0, 1, 1, 1, 0]

    def test_nested_groups(self):
        types = [
            "begin_group",
            "text",
            "begin_repeat",
            "text",
            "begin group",
            "text",
            "end group",
            "end_repeat",
            "end_group",
        ]
        assert nesting_levels(types) == [1, 1, 2, 2, 3, 3, 3, 2, 1]

    def test_metadata_end_does_not_close_a_group(self):
        """Regression test: "type: end" used to close the enclosing group.

        In a real form the metadata rows sit inside the first group, which made
        every later group look like a top-level one.
        """
        types = ["begin_group", "start", "end", "text", "end_group", "text"]
        assert nesting_levels(types) == [1, 1, 1, 1, 1, 0]

    def test_metadata_end_at_top_level(self):
        """A top-level "end" used to drive the depth negative, permanently."""
        types = ["start", "end", "begin_group", "text", "end_group"]
        assert nesting_levels(types) == [0, 0, 1, 1, 1]

    def test_adjacent_groups(self):
        types = ["begin_group", "end_group", "begin_group", "end_group"]
        assert nesting_levels(types) == [1, 1, 1, 1]

    def test_mixed_spellings_pair_up(self):
        """A group may be opened and closed with different spellings."""
        types = ["begin_group", "begin repeat", "end_repeat", "end group"]
        assert nesting_levels(types) == [1, 2, 2, 1]

    def test_unclosed_group_does_not_raise(self):
        assert nesting_levels(["begin_group", "text"]) == [1, 1]

    def test_surplus_end_does_not_go_negative(self):
        types = ["end_group", "text", "begin_group", "text", "end_group"]
        assert nesting_levels(types) == [0, 0, 1, 1, 1]

    def test_empty_input(self):
        assert nesting_levels([]) == []

    def test_blank_and_missing_types(self):
        """Blank rows keep the level of the group that contains them."""
        types = ["begin_group", None, "text", "", "end_group"]
        assert nesting_levels(types) == [1, 1, 1, 1, 1]
