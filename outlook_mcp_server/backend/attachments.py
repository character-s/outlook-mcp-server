"""Attachment handling for outgoing mail.

Composed messages are always saved as drafts, never sent, so attachments are
resolved and validated up front: a bad path should fail before a half-built
message reaches the Drafts folder.
"""

# Standard library imports
import os
from typing import List, NamedTuple, Optional, Union

# Local application imports
from .config import attachment_config
from .email_utils import format_file_size
from .logging_config import get_logger
from .validation import ValidationError

logger = get_logger(__name__)


class ResolvedAttachment(NamedTuple):
    """A local file that has been checked and is ready to hand to Outlook."""

    path: str
    name: str
    size: int


def resolve_attachments(
    attachments: Optional[Union[str, List[str]]],
) -> List[ResolvedAttachment]:
    """Validate attachment paths before any Outlook item is created.

    Args:
        attachments: A single path, a list of paths, or None. Paths may use
            ~ and environment variables; they are expanded and made absolute.

    Returns:
        list: Resolved attachments in the order given (empty when None).

    Raises:
        ValidationError: If a path is missing, is not a file, or the batch
            exceeds the configured count/size limits.
    """
    if attachments is None:
        return []

    if isinstance(attachments, str):
        attachments = [attachments]

    if not isinstance(attachments, list):
        raise ValidationError("Attachments must be a file path string or a list of file path strings")

    resolved: List[ResolvedAttachment] = []
    seen_paths = set()

    for index, raw_path in enumerate(attachments, start=1):
        if not isinstance(raw_path, str) or not raw_path.strip():
            raise ValidationError(f"Attachment #{index} must be a non-empty file path string")

        expanded = os.path.abspath(
            os.path.expanduser(os.path.expandvars(raw_path.strip()))
        )

        if os.path.isdir(expanded):
            raise ValidationError(
                f"Attachment #{index} is a directory, not a file: {expanded}. "
                "Attach individual files, or zip the folder first."
            )
        if not os.path.isfile(expanded):
            raise ValidationError(
                f"Attachment #{index} not found: {expanded}. "
                "Check the path - it must be a local file reachable from this machine."
            )

        # Outlook happily attaches the same file twice; the user almost never
        # means that, and a silent duplicate is easy to miss in a draft.
        lookup_key = os.path.normcase(expanded)
        if lookup_key in seen_paths:
            raise ValidationError(f"Attachment #{index} is a duplicate of an earlier attachment: {expanded}")
        seen_paths.add(lookup_key)

        size = os.path.getsize(expanded)
        if size == 0:
            raise ValidationError(f"Attachment #{index} is an empty file (0 bytes): {expanded}")

        resolved.append(ResolvedAttachment(path=expanded, name=os.path.basename(expanded), size=size))

    if len(resolved) > attachment_config.MAX_ATTACHMENT_COUNT:
        raise ValidationError(
            f"Too many attachments: {len(resolved)} "
            f"(maximum {attachment_config.MAX_ATTACHMENT_COUNT})"
        )

    total_size = sum(item.size for item in resolved)
    if total_size > attachment_config.MAX_TOTAL_ATTACHMENT_BYTES:
        raise ValidationError(
            f"Attachments total {format_file_size(total_size)}, which exceeds the "
            f"{format_file_size(attachment_config.MAX_TOTAL_ATTACHMENT_BYTES)} limit. "
            "Most mail servers reject messages this large - share a link instead, "
            "or raise MAX_TOTAL_ATTACHMENT_BYTES in backend/config.py if your server allows it."
        )

    return resolved


def attach_files(mail, resolved: List[ResolvedAttachment]) -> None:
    """Add resolved files to an unsaved Outlook mail item.

    Args:
        mail: The Outlook MailItem COM object to attach to.
        resolved: Attachments from resolve_attachments().

    Raises:
        RuntimeError: If Outlook rejects a file. The caller is expected to
            discard the draft so a partially attached message is never saved.
    """
    for item in resolved:
        try:
            mail.Attachments.Add(item.path)
            logger.debug(f"Attached {item.name} ({format_file_size(item.size)})")
        except Exception as e:
            logger.error(f"Outlook rejected attachment {item.path}: {e}")
            raise RuntimeError(f"Outlook could not attach '{item.name}': {e}")


def describe_attachments(resolved: List[ResolvedAttachment]) -> str:
    """Build a human-readable summary of what was attached, for tool output."""
    if not resolved:
        return ""

    listed = ", ".join(f"{item.name} ({format_file_size(item.size)})" for item in resolved)
    return f" Attached {len(resolved)} file(s): {listed}."
