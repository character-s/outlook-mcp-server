"""Email composition and reply functions with improved encoding handling"""

# Type imports
from typing import Any, Callable, Dict, List, Optional, Union

# Local application imports
from .attachments import attach_files, describe_attachments, resolve_attachments
from .config import email_defaults_config
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager
from .shared import email_cache, email_cache_order
from .utils import safe_encode_text, normalize_email_address
from .validation import (
    DisplayConstants,
    OutlookConstants,
    ValidationError,
    validate_cache_available,
    validate_email_number
)
from .validators import EmailComposeParams, EmailReplyParams

logger = get_logger(__name__)


def _read_cached_recipient(recipient_info: Dict[str, Any]) -> tuple:
    """Pull the address and display name out of a cached recipient entry.

    The extractors emit {"address", "name"}. Older cache files (and any caller
    building entries by hand) use {"email", "display_name"}, so both shapes are
    accepted - reading only the latter silently dropped every CC recipient.

    Args:
        recipient_info: One entry from a cached email's to/cc recipient list.

    Returns:
        tuple: (address, display_name), each stripped, either possibly empty.
    """
    address = recipient_info.get("address") or recipient_info.get("email") or ""
    display_name = recipient_info.get("name") or recipient_info.get("display_name") or ""
    return address.strip(), display_name.strip()


def _format_recipient(address: str, display_name: str) -> str:
    """Render a recipient for an Outlook To/CC field.

    Uses "Display Name <address>" only when that is actually meaningful. Exchange
    entries often carry the display name in both fields, and "Name <Name>" is not
    an address Outlook can resolve.

    Args:
        address: The recipient address (may itself be a display name).
        display_name: The recipient's display name, if known.

    Returns:
        str: The recipient string to put in the field.
    """
    if display_name and "@" in address and display_name != address:
        return f"{display_name} <{address}>"
    return address or display_name


def merge_default_cc(
    cc_entries: List[str],
    is_sender: Optional[Callable[[str], bool]] = None,
) -> List[str]:
    """Add the configured default CC addresses to a CC list.

    Args:
        cc_entries: CC entries already on the message, either bare addresses
            or "Display Name <address>" strings.
        is_sender: Optional predicate identifying the person being replied to.
            Matching defaults are skipped so a reply never CCs them back.

    Returns:
        list: cc_entries followed by the default addresses not already present.
            Comparison is on the normalized address, so "Name <a@b.com>" and
            "a@b.com" count as the same recipient.
    """
    merged = list(cc_entries)
    defaults = email_defaults_config.DEFAULT_CC
    if not defaults:
        return merged

    seen = {normalize_email_address(entry) for entry in merged}

    for address in defaults:
        normalized = normalize_email_address(address)
        if normalized in seen:
            logger.debug(f"Default CC already present, skipping: {address}")
            continue
        if is_sender is not None and is_sender(address):
            logger.debug(f"Default CC matches the original sender, skipping: {address}")
            continue

        merged.append(address)
        seen.add(normalized)
        logger.debug(f"Added default CC: {address}")

    return merged


def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[Union[str, List[str]]] = None,
    cc_recipients: Optional[Union[str, List[str]]] = None,
    attachments: Optional[Union[str, List[str]]] = None,
) -> str:
    """
    Reply to an email with custom recipients if provided.

    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
        attachments: Local file path, or list of file paths, to attach to the draft

    Returns:
        str: Success or error message
    """
    # Resolve attachments before touching Outlook so a bad path fails without
    # leaving a half-built draft behind.
    resolved_attachments = resolve_attachments(attachments)

    # Validate inputs using Pydantic
    try:
        params = EmailReplyParams(
            email_number=email_number,
            reply_text=reply_text,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
        )
    except Exception as e:
        logger.error(f"Validation error in reply_to_email_by_number: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Convert to list if needed (validator already did this)
    to_recipients = params.to_recipients
    cc_recipients = params.cc_recipients
    reply_text = params.reply_text

    try:
        validate_cache_available(len(email_cache_order))
        validate_email_number(email_number, len(email_cache_order))
    except ValidationError as e:
        logger.error(f"Validation error in reply_to_email_by_number: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Get the entry_id from the cache order
    entry_id = email_cache_order[email_number - 1]
    if not entry_id:
        raise ValueError(f"Email #{email_number} has no entry ID")

    # Get the cached email data
    cached_email = email_cache.get(entry_id)
    if not cached_email:
        raise ValueError(f"Email #{email_number} data not found in cache")

    with OutlookSessionManager() as session:
        try:
            # Get the email ID, handling different key names that might be used
            email_id = cached_email.get("id") or cached_email.get("entry_id")
            if not email_id:
                raise ValueError(f"Email ID not found in cached data. Available keys: {list(cached_email.keys())}")
            
            email = session.namespace.GetItemFromID(email_id)
            if not email:
                raise RuntimeError("Could not retrieve the email from Outlook.")

            # Create a new email message to have full control over formatting
            new_mail = session.outlook.CreateItem(OutlookConstants.OL_MAIL_ITEM)

            # Extract sender email early for use in CC filtering
            sender_email = safe_encode_text(
                getattr(email, "SenderEmailAddress", "unknown@example.com"), "to_address"
            )
            normalized_sender_email = normalize_email_address(sender_email)

            # Additional sender extraction for robustness
            sender_name = getattr(email, "SenderName", "")
            sender_address = getattr(email, "SenderEmailAddress", "")

            # Log comprehensive sender information for debugging
            logger.debug(f"=== SENDER EXTRACTION DEBUG ===")
            logger.debug(f"SenderEmailAddress: {sender_email}")
            logger.debug(f"SenderName: {sender_name}")
            logger.debug(f"Combined sender info: {sender_name} <{sender_address}>")
            logger.debug(f"Normalized sender email: {normalized_sender_email}")
            logger.debug(f"=== END SENDER EXTRACTION DEBUG ===")

            # Also check if sender appears in original email fields
            original_to = safe_encode_text(getattr(email, "To", ""), "original_to")
            original_cc = safe_encode_text(getattr(email, "CC", ""), "original_cc")
            logger.debug(f"Original TO field: {original_to}")
            logger.debug(f"Original CC field: {original_cc}")

            # Create a comprehensive list of sender variations to filter against
            sender_variations = set()
            sender_variations.add(normalized_sender_email)

            # Add display name variations
            if sender_name and sender_address:
                # "Name <email@domain.com>" format
                display_format = f"{sender_name} <{sender_address}>".strip()
                sender_variations.add(normalize_email_address(display_format))

                # Also check individual components
                sender_variations.add(normalize_email_address(sender_name))

            # Note whether the sender also appears in the original To/CC fields.
            # These addresses are deliberately NOT registered as sender variations:
            # doing so marked every original recipient as "the sender", so ReplyAll
            # silently dropped the entire CC list.
            for field_name, field_value in (("TO", original_to), ("CC", original_cc)):
                if not field_value:
                    continue
                for address in field_value.split(";"):
                    address = address.strip()
                    if address and normalize_email_address(address) == normalized_sender_email:
                        logger.debug(f"Found sender in original {field_name} field: {address}")

            logger.debug(f"Sender variations to filter against: {sorted(sender_variations)}")

            # Create a comprehensive filtering function
            def is_sender_email(email_address: str) -> bool:
                """Check if an email address matches any sender variation"""
                normalized = normalize_email_address(email_address)
                return normalized in sender_variations

            # Determine recipients based on parameters
            if to_recipients is None and cc_recipients is None:
                # ReplyAll behavior - get all original recipients
                new_mail.To = sender_email

                # Use cached recipient data to avoid Outlook name resolution issues
                cc_recipients_set = set()

                # Get CC recipients from cache using both display names and email addresses
                cc_recipients_data = cached_email.get("cc_recipients", [])
                logger.debug(f"Processing {len(cc_recipients_data)} CC recipients from cache")

                for i, recipient_info in enumerate(cc_recipients_data):
                    if isinstance(recipient_info, dict):
                        recipient_email, recipient_display_name = _read_cached_recipient(
                            recipient_info
                        )
                        normalized_recipient_email = normalize_email_address(recipient_email)

                        logger.debug(f"CC recipient {i+1}: {recipient_info}")
                        logger.debug(f"  Extracted email: '{recipient_email}'")
                        logger.debug(f"  Extracted display name: '{recipient_display_name}'")
                        logger.debug(f"  Normalized email: '{normalized_recipient_email}'")
                        logger.debug(f"  Sender normalized: '{normalized_sender_email}'")
                        logger.debug(f"  Is sender: {is_sender_email(recipient_email)}")

                        if recipient_email:
                            if not is_sender_email(recipient_email):
                                recipient_string = _format_recipient(
                                    recipient_email, recipient_display_name
                                )
                                cc_recipients_set.add(recipient_string)
                                logger.debug(f"  -> ADDED to CC: {recipient_string}")
                            else:
                                logger.debug(
                                    f"  -> FILTERED OUT (matches sender): {recipient_email}"
                                )
                        else:
                            logger.debug(f"  -> SKIPPED (empty email)")
                    else:
                        logger.debug(f"CC recipient {i+1}: Non-dict format: {recipient_info}")

                logger.debug(f"Total CC recipients after filtering: {len(cc_recipients_set)}")
                if cc_recipients_set:
                    logger.debug(f"CC recipients list: {sorted(cc_recipients_set)}")

                # Set CC field with filtered CC recipients plus any configured defaults
                final_cc = merge_default_cc(sorted(cc_recipients_set), is_sender_email)
                if final_cc:
                    logger.debug(f"Setting CC to (ReplyAll): {final_cc}")
                    new_mail.CC = "; ".join(final_cc)
                else:
                    # Explicitly clear CC field if no valid recipients remain
                    logger.debug("No CC recipients after filtering - clearing CC field")
                    new_mail.CC = ""
            else:
                # Use custom recipients, but ensure original sender is not in CC
                if to_recipients is not None:
                    new_mail.To = "; ".join(to_recipients)

                filtered_cc = []
                if cc_recipients is not None:
                    # Filter out the original sender from CC recipients
                    for recipient in cc_recipients:
                        # Use comprehensive sender filtering
                        if not is_sender_email(recipient):
                            filtered_cc.append(recipient)
                            logger.debug(f"CC recipient kept: {recipient}")
                        else:
                            logger.info(f"Filtered out original sender from CC: {recipient}")

                # Configured defaults apply even when the caller named no CC at all.
                final_cc = merge_default_cc(filtered_cc, is_sender_email)
                if final_cc:
                    logger.debug(f"Setting CC to: {final_cc}")
                    new_mail.CC = "; ".join(final_cc)
                else:
                    # Explicitly clear CC field if no valid recipients remain
                    logger.debug("No CC recipients after filtering - clearing CC field")
                    new_mail.CC = ""

            # Set subject with RE: prefix
            subject = safe_encode_text(getattr(email, "Subject", "No Subject"), "subject")
            new_mail.Subject = f"RE: {subject}"

            # Build the email body with proper formatting and encoding
            reply_text_safe = safe_encode_text(reply_text, "reply_text")
            sender_name = safe_encode_text(
                getattr(email, "SenderName", "Unknown Sender"), "sender_name"
            )
            sent_on = safe_encode_text(str(getattr(email, "SentOn", "Unknown")), "sent_on")
            to_field = safe_encode_text(getattr(email, "To", "Unknown"), "to_field")

            # Build body content
            body_lines = [
                reply_text_safe,
                "",
                "_" * DisplayConstants.SEPARATOR_LINE_LENGTH,
                f"From: {sender_name}",
                f"Sent: {sent_on}",
                f"To: {to_field}",
            ]

            # Add CC if present
            original_cc = safe_encode_text(getattr(email, "CC", ""), "original_cc")
            if original_cc and original_cc.strip():
                body_lines.append(f"Cc: {original_cc}")

            body_lines.extend([f"Subject: {subject}", ""])

            # Add the original email content
            original_body = safe_encode_text(getattr(email, "Body", ""), "original_body")
            body_lines.append(original_body)

            # Join with proper line endings
            body_content = "\n".join(body_lines)

            # Set the body of the new email
            try:
                new_mail.Body = body_content
            except Exception as e:
                logger.warning(f"Failed to set email body, using simplified version: {e}")
                # Fallback to simple body
                new_mail.Body = (
                    f"{reply_text_safe}\n\n{'_' * DisplayConstants.SEPARATOR_LINE_LENGTH}\n[Original email content unavailable]"
                )

            if resolved_attachments:
                try:
                    attach_files(new_mail, resolved_attachments)
                except Exception:
                    # Discard rather than save a draft that is missing files the
                    # user asked for and might not notice are absent.
                    try:
                        new_mail.Close(OutlookConstants.OL_DISCARD)
                    except Exception as close_error:
                        logger.warning(f"Failed to discard incomplete draft: {close_error}")
                    raise

            # Save as a draft instead of sending. Nothing leaves the mailbox
            # without the user opening Drafts in Outlook and clicking Send.
            new_mail.Save()
            logger.info(
                f"Saved draft reply for email #{email_number} "
                f"with {len(resolved_attachments)} attachment(s)"
            )
            return (
                f"Draft reply to email #{email_number} saved to the Drafts folder."
                f"{describe_attachments(resolved_attachments)} "
                "This server never sends mail; review and send it yourself in Outlook."
            )

        except Exception as e:
            logger.error(f"Error replying to email #{email_number}: {e}")
            return f"Error replying to email: {str(e)}"


def compose_email(
    to_recipients: List[str],
    subject: str,
    body: str,
    cc_recipients: Optional[List[str]] = None,
    html: bool = False,
    attachments: Optional[Union[str, List[str]]] = None,
) -> str:
    """
    Compose a new email using Outlook COM API and save it as a draft.

    Args:
        to_recipients: List of recipient email addresses
        subject: Email subject line
        body: Email body content
        cc_recipients: Optional list of CC email addresses
        html: If True, body is treated as HTML (default: False)
        attachments: Local file path, or list of file paths, to attach to the draft

    Returns:
        str: Success/error message
    """
    # Resolve attachments before touching Outlook so a bad path fails without
    # leaving a half-built draft behind.
    resolved_attachments = resolve_attachments(attachments)

    # Validate inputs using Pydantic
    try:
        params = EmailComposeParams(
            recipient_email=to_recipients[0] if to_recipients else "",
            subject=subject,
            body=body,
            cc_email=cc_recipients[0] if cc_recipients else None,
        )
    except Exception as e:
        logger.error(f"Validation error in compose_email: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Additional validation for list
    if not to_recipients or not isinstance(to_recipients, list):
        raise ValueError("To recipients must be a non-empty list")

    if not all(isinstance(email, str) and email.strip() for email in to_recipients):
        raise ValueError("All recipient email addresses must be non-empty strings")

    if cc_recipients is not None:
        if not isinstance(cc_recipients, list):
            raise ValueError("CC recipients must be a list or None")
        if not all(isinstance(email, str) and email.strip() for email in cc_recipients):
            raise ValueError("All CC email addresses must be non-empty strings")

    with OutlookSessionManager() as session:
        try:
            # Encode all components safely
            encoded_to = [
                safe_encode_text(recipient, "to_recipient").strip() for recipient in to_recipients
            ]
            subject_safe = safe_encode_text(subject, "subject")
            body_safe = safe_encode_text(body, "body")

            encoded_cc = []
            if cc_recipients:
                encoded_cc = [
                    safe_encode_text(recipient, "cc_recipient").strip()
                    for recipient in cc_recipients
                ]

            # Create and send the email
            mail = session.outlook.CreateItem(OutlookConstants.OL_MAIL_ITEM)
            mail.To = "; ".join(encoded_to)
            mail.Subject = subject_safe

            final_cc = merge_default_cc(encoded_cc)
            if final_cc:
                mail.CC = "; ".join(final_cc)

            try:
                if html:
                    mail.HTMLBody = body_safe
                else:
                    mail.Body = body_safe
            except Exception as e:
                logger.warning(f"Failed to set email body format, using plain text: {e}")
                mail.Body = body_safe

            if resolved_attachments:
                try:
                    attach_files(mail, resolved_attachments)
                except Exception:
                    # Discard rather than save a draft that is missing files the
                    # user asked for and might not notice are absent.
                    try:
                        mail.Close(OutlookConstants.OL_DISCARD)
                    except Exception as close_error:
                        logger.warning(f"Failed to discard incomplete draft: {close_error}")
                    raise

            # Save as a draft instead of sending. See reply_to_email_by_number.
            mail.Save()
            logger.info(
                f"Saved draft email for {len(to_recipients)} recipients "
                f"with {len(resolved_attachments)} attachment(s)"
            )
            return (
                f"Draft email to {len(to_recipients)} recipient(s) saved to the Drafts folder."
                f"{describe_attachments(resolved_attachments)} "
                "This server never sends mail; review and send it yourself in Outlook."
            )

        except Exception as e:
            logger.error(f"Error composing email: {e}")
            return f"Error composing email: {str(e)}"
