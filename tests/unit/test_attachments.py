import os
from unittest.mock import MagicMock

import pytest

from outlook_mcp_server.backend.attachments import (
    ResolvedAttachment,
    attach_files,
    describe_attachments,
    resolve_attachments,
)
from outlook_mcp_server.backend.config import attachment_config
from outlook_mcp_server.backend.validation import ValidationError


@pytest.fixture
def sample_file(tmp_path):
    """A small real file on disk to attach."""
    path = tmp_path / "report.xlsx"
    path.write_bytes(b"x" * 1024)
    return str(path)


class TestResolveAttachments:
    """Test suite for attachment path resolution."""

    def test_none_returns_empty_list(self):
        assert resolve_attachments(None) == []

    def test_single_string_is_accepted(self, sample_file):
        resolved = resolve_attachments(sample_file)
        assert len(resolved) == 1
        assert resolved[0].name == "report.xlsx"
        assert resolved[0].size == 1024
        assert os.path.isabs(resolved[0].path)

    def test_list_preserves_order(self, tmp_path):
        first = tmp_path / "a.txt"
        second = tmp_path / "b.txt"
        first.write_text("a")
        second.write_text("bb")

        resolved = resolve_attachments([str(first), str(second)])
        assert [item.name for item in resolved] == ["a.txt", "b.txt"]
        assert [item.size for item in resolved] == [1, 2]

    def test_user_home_is_expanded(self, monkeypatch, tmp_path):
        home_file = tmp_path / "in_home.txt"
        home_file.write_text("data")
        monkeypatch.setenv("USERPROFILE", str(tmp_path))
        monkeypatch.setenv("HOME", str(tmp_path))

        resolved = resolve_attachments("~/in_home.txt")
        assert resolved[0].name == "in_home.txt"

    def test_environment_variable_is_expanded(self, monkeypatch, tmp_path):
        target = tmp_path / "from_env.txt"
        target.write_text("data")
        monkeypatch.setenv("ATTACH_TEST_DIR", str(tmp_path))

        resolved = resolve_attachments(os.path.join("$ATTACH_TEST_DIR", "from_env.txt"))
        assert resolved[0].name == "from_env.txt"

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(ValidationError, match="not found"):
            resolve_attachments(str(tmp_path / "nope.xlsx"))

    def test_directory_raises(self, tmp_path):
        with pytest.raises(ValidationError, match="directory"):
            resolve_attachments(str(tmp_path))

    def test_empty_file_raises(self, tmp_path):
        empty = tmp_path / "empty.txt"
        empty.write_bytes(b"")
        with pytest.raises(ValidationError, match="empty file"):
            resolve_attachments(str(empty))

    def test_empty_string_raises(self):
        with pytest.raises(ValidationError, match="non-empty file path"):
            resolve_attachments(["   "])

    def test_non_string_entry_raises(self, sample_file):
        with pytest.raises(ValidationError, match="non-empty file path"):
            resolve_attachments([sample_file, 42])

    def test_wrong_type_raises(self):
        with pytest.raises(ValidationError, match="file path string"):
            resolve_attachments({"path": "a.txt"})

    def test_duplicate_path_raises(self, sample_file):
        with pytest.raises(ValidationError, match="duplicate"):
            resolve_attachments([sample_file, sample_file])

    def test_too_many_files_raises(self, tmp_path, monkeypatch):
        monkeypatch.setattr(attachment_config, "MAX_ATTACHMENT_COUNT", 2)
        paths = []
        for index in range(3):
            path = tmp_path / f"file{index}.txt"
            path.write_text("data")
            paths.append(str(path))

        with pytest.raises(ValidationError, match="Too many attachments"):
            resolve_attachments(paths)

    def test_total_size_limit_raises(self, tmp_path, monkeypatch):
        monkeypatch.setattr(attachment_config, "MAX_TOTAL_ATTACHMENT_BYTES", 1000)
        big = tmp_path / "big.bin"
        big.write_bytes(b"x" * 2000)

        with pytest.raises(ValidationError, match="exceeds"):
            resolve_attachments(str(big))

    def test_size_limit_applies_to_the_total_not_each_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(attachment_config, "MAX_TOTAL_ATTACHMENT_BYTES", 1000)
        first = tmp_path / "one.bin"
        second = tmp_path / "two.bin"
        first.write_bytes(b"x" * 600)
        second.write_bytes(b"x" * 600)

        with pytest.raises(ValidationError, match="exceeds"):
            resolve_attachments([str(first), str(second)])


class TestAttachFiles:
    """Test suite for handing resolved files to Outlook."""

    def test_each_file_is_added(self):
        mail = MagicMock()
        resolved = [
            ResolvedAttachment(path=r"C:\a.txt", name="a.txt", size=1),
            ResolvedAttachment(path=r"C:\b.txt", name="b.txt", size=2),
        ]

        attach_files(mail, resolved)

        assert mail.Attachments.Add.call_count == 2
        mail.Attachments.Add.assert_any_call(r"C:\a.txt")
        mail.Attachments.Add.assert_any_call(r"C:\b.txt")

    def test_outlook_failure_is_raised_with_file_name(self):
        mail = MagicMock()
        mail.Attachments.Add.side_effect = Exception("COM error")
        resolved = [ResolvedAttachment(path=r"C:\a.txt", name="a.txt", size=1)]

        with pytest.raises(RuntimeError, match="a.txt"):
            attach_files(mail, resolved)


class TestDescribeAttachments:
    """Test suite for the human-readable summary."""

    def test_empty_list_produces_no_text(self):
        assert describe_attachments([]) == ""

    def test_names_and_sizes_are_listed(self):
        resolved = [ResolvedAttachment(path=r"C:\a.xlsx", name="a.xlsx", size=1024)]
        summary = describe_attachments(resolved)
        assert "1 file(s)" in summary
        assert "a.xlsx" in summary
        assert "KB" in summary
