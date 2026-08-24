import pytest

from outlook_mcp_server.backend.config import email_defaults_config
from outlook_mcp_server.backend.email_composition import (
    _format_recipient,
    _read_cached_recipient,
    merge_default_cc,
)

ENV_VAR = email_defaults_config.DEFAULT_CC_ENV_VAR


@pytest.fixture(autouse=True)
def clear_default_cc(monkeypatch):
    """Keep a stray real environment variable from leaking into the tests."""
    monkeypatch.delenv(ENV_VAR, raising=False)


class TestDefaultCcConfig:
    """Test suite for parsing the DEFAULT_CC environment variable."""

    def test_unset_gives_no_addresses(self):
        assert email_defaults_config.DEFAULT_CC == []

    def test_blank_gives_no_addresses(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "   ")
        assert email_defaults_config.DEFAULT_CC == []

    def test_single_address(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "boss@example.com")
        assert email_defaults_config.DEFAULT_CC == ["boss@example.com"]

    def test_semicolon_and_comma_separators(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "a@example.com; b@example.com, c@example.com")
        assert email_defaults_config.DEFAULT_CC == [
            "a@example.com",
            "b@example.com",
            "c@example.com",
        ]

    def test_empty_entries_are_dropped(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, ";a@example.com;;")
        assert email_defaults_config.DEFAULT_CC == ["a@example.com"]

    def test_display_name_form_is_preserved(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "Boss <boss@example.com>")
        assert email_defaults_config.DEFAULT_CC == ["Boss <boss@example.com>"]


class TestMergeDefaultCc:
    """Test suite for merging defaults into a CC list."""

    def test_no_defaults_leaves_list_untouched(self):
        assert merge_default_cc(["a@example.com"]) == ["a@example.com"]

    def test_default_is_appended(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "boss@example.com")
        assert merge_default_cc(["a@example.com"]) == ["a@example.com", "boss@example.com"]

    def test_default_applies_to_an_empty_cc_list(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "boss@example.com")
        assert merge_default_cc([]) == ["boss@example.com"]

    def test_exact_duplicate_is_not_added_twice(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "boss@example.com")
        assert merge_default_cc(["boss@example.com"]) == ["boss@example.com"]

    def test_duplicate_is_detected_through_display_name_form(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "boss@example.com")
        existing = ["The Boss <boss@example.com>"]
        assert merge_default_cc(existing) == existing

    def test_duplicate_is_detected_case_insensitively(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "BOSS@Example.COM")
        existing = ["boss@example.com"]
        assert merge_default_cc(existing) == existing

    def test_sender_is_never_cced_back(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "boss@example.com")

        def is_sender(address):
            return "boss@example.com" in address.lower()

        assert merge_default_cc(["a@example.com"], is_sender) == ["a@example.com"]

    def test_several_defaults_keep_their_order(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "one@example.com;two@example.com")
        assert merge_default_cc(["a@example.com"]) == [
            "a@example.com",
            "one@example.com",
            "two@example.com",
        ]

    def test_input_list_is_not_mutated(self, monkeypatch):
        monkeypatch.setenv(ENV_VAR, "boss@example.com")
        original = ["a@example.com"]
        merge_default_cc(original)
        assert original == ["a@example.com"]


class TestCachedRecipientReading:
    """Test suite for reading recipient entries out of the cache.

    Regression guard: the extractors emit {"address", "name"}, but the reply path
    used to read {"email", "display_name"} only, which silently dropped every CC.
    """

    def test_extractor_shape_is_read(self):
        entry = {"address": "boss@example.com", "name": "The Boss"}
        assert _read_cached_recipient(entry) == ("boss@example.com", "The Boss")

    def test_legacy_shape_is_still_read(self):
        entry = {"email": "boss@example.com", "display_name": "The Boss"}
        assert _read_cached_recipient(entry) == ("boss@example.com", "The Boss")

    def test_display_name_only_entry(self):
        entry = {"address": "Sales Team", "name": "Sales Team"}
        assert _read_cached_recipient(entry) == ("Sales Team", "Sales Team")

    def test_missing_keys_give_empty_strings(self):
        assert _read_cached_recipient({}) == ("", "")

    def test_none_values_do_not_crash(self):
        assert _read_cached_recipient({"address": None, "name": None}) == ("", "")

    def test_values_are_stripped(self):
        entry = {"address": "  boss@example.com  ", "name": "  The Boss  "}
        assert _read_cached_recipient(entry) == ("boss@example.com", "The Boss")


class TestRecipientFormatting:
    """Test suite for rendering a recipient into an Outlook address field."""

    def test_name_and_address_are_combined(self):
        assert _format_recipient("boss@example.com", "The Boss") == "The Boss <boss@example.com>"

    def test_address_only(self):
        assert _format_recipient("boss@example.com", "") == "boss@example.com"

    def test_identical_name_and_address_are_not_doubled(self):
        # Exchange entries often repeat the display name in both fields;
        # "Name <Name>" is not something Outlook can resolve.
        assert _format_recipient("Sales Team", "Sales Team") == "Sales Team"

    def test_display_name_without_at_sign_is_used_bare(self):
        assert _format_recipient("Sales Team", "Sales Dept") == "Sales Team"

    def test_empty_address_falls_back_to_display_name(self):
        assert _format_recipient("", "The Boss") == "The Boss"
