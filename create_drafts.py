import argparse
import csv
import html
import logging
import os
import re
import uuid
import winreg
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlsplit


BASE_DIR = Path(__file__).resolve().parent
LOGGER = logging.getLogger("outlook_draft_mailer")


@dataclass(frozen=True)
class OutlookSignature:
    markup: str
    directory: Path


def _column_name(index: int) -> str:
    name = ""
    while index >= 0:
        index, remainder = divmod(index, 26)
        name = chr(ord("A") + remainder) + name
        index -= 1
    return name


def configure_logging(log_path: Path) -> logging.Logger:
    logger = LOGGER
    logger.setLevel(logging.INFO)
    logger.propagate = False

    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
        handler.close()

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    )
    logger.addHandler(file_handler)
    return logger


def load_contacts(
    input_path: Path
) -> tuple[list[str], list[dict[str, str]]]:
    suffix = input_path.suffix.lower()
    if suffix == ".xlsx":
        try:
            from openpyxl import load_workbook
        except ImportError as error:
            raise RuntimeError(
                "Excel support is not installed. Run "
                "'python -m pip install -r requirements.txt'."
            ) from error

        workbook = load_workbook(input_path, read_only=True, data_only=True)
        try:
            worksheet = workbook.active
            all_rows = [
                list(row)
                for row in worksheet.iter_rows(values_only=True)
                if any(value is not None and str(value).strip() for value in row)
            ]
            if not all_rows:
                raise ValueError(f"{input_path.name} contains no rows.")

            headers = [
                _column_name(index)
                for index in range(max(len(row) for row in all_rows))
            ]
            first_data_row_number = 1
            if len(headers) != len(set(headers)):
                raise ValueError(f"{input_path.name} contains duplicate column names.")
            LOGGER.info("Headers loaded from %s: %s", input_path.name, headers)
            contacts = []
            for row_number, row in enumerate(all_rows, start=first_data_row_number):
                if len(row) > len(headers) and any(
                    value is not None for value in row[len(headers):]
                ):
                    raise ValueError(
                        f"{input_path.name}:{row_number} has more values than headers."
                    )
                contacts.append(
                    {
                        header: str(row[index]).strip()
                        if index < len(row) and row[index] is not None
                        else ""
                        for index, header in enumerate(headers)
                    }
                )
        finally:
            workbook.close()
    else:
        raise ValueError("Choose a Excel workbook (.xlsx) file.")

    if not contacts:
        raise ValueError(f"{input_path.name} contains no contact rows.")
    return headers, contacts


def render_messages(
    contacts: list[dict[str, str]],
    subject_template: str,
    body_template: str,
    recipient_column: str = "email",
) -> list[dict[str, Any]]:
    if not contacts:
        raise ValueError("The contact file contains no contact rows.")
    if recipient_column not in contacts[0]:
        raise ValueError(
            f"Choose a recipient column from the available columns: "
            f"{', '.join(contacts[0])}."
        )

    messages = []
    for row_number, contact in enumerate(contacts, start=2):
        email = contact.get(recipient_column, "").strip()
        if not email:
            raise ValueError(
                f"Contact row {row_number} has no address in the selected "
                f"recipient column {recipient_column!r}."
            )

        try:
            subject = subject_template.format_map(contact).strip()
            body = body_template.format_map(contact).strip()
        except KeyError as error:
            missing_field = error.args[0]
            raise ValueError(
                f"Contact row {row_number} is missing the template "
                f"column {missing_field!r}. Available columns: "
                f"{', '.join(contact)}."
            ) from error

        if not subject or not body:
            raise ValueError(
                f"Contact row {row_number} rendered an empty subject or body."
            )
        messages.append(
            {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [
                    {"emailAddress": {"address": email}},
                ],
            }
        )
    return messages


def _read_signature_file(path: Path) -> str:
    content = path.read_bytes()
    if content.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = content.decode("utf-16")
    else:
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = content.decode("cp1252")
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def get_default_outlook_signature() -> OutlookSignature:
    signature_name = ""
    for version in ("16.0", "15.0", "14.0", "12.0"):
        key_path = rf"Software\Microsoft\Office\{version}\Common\MailSettings"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                signature_name, _ = winreg.QueryValueEx(key, "NewSignature")
                if signature_name:
                    break
        except FileNotFoundError:
            continue

    signature_dir = Path(os.environ["APPDATA"]) / "Microsoft" / "Signatures"
    if not signature_name:
        available_signatures = {
            path.stem
            for pattern in ("*.txt", "*.htm")
            for path in signature_dir.glob(pattern)
        }
        if len(available_signatures) == 1:
            signature_name = available_signatures.pop()
        elif available_signatures:
            raise RuntimeError(
                "No default Outlook signature is configured, and multiple "
                "signatures were found. Set a default signature for new messages "
                "in classic Outlook, then try again."
            )
        else:
            raise RuntimeError(
                "No default Outlook signature is configured and no local "
                "signature files were found. Set a default signature for new "
                "messages in classic Outlook, then try again."
            )

    html_path = signature_dir / f"{signature_name}.htm"
    if html_path.is_file():
        return OutlookSignature(_read_signature_file(html_path), signature_dir)

    text_path = signature_dir / f"{signature_name}.txt"
    if text_path.is_file():
        text = _read_signature_file(text_path)
        markup = f"<div>{html.escape(text).replace(chr(10), '<br>')}</div>"
        return OutlookSignature(markup, signature_dir)

    raise RuntimeError(
        f"Could not find the Outlook signature file for {signature_name!r} "
        f"in {signature_dir}."
    )


def _signature_resource_path(reference: str, signature_dir: Path) -> Path | None:
    reference = html.unescape(unquote(reference.strip())).replace("\\", "/")
    if re.match(r"^[A-Za-z]:/", reference):
        path = Path(reference)
        resource_name = ""
    else:
        parsed = urlsplit(reference)
        if parsed.scheme and parsed.scheme.lower() != "file":
            return None
        resource_name = parsed.path if parsed.scheme.lower() == "file" else reference
        path = Path(resource_name)

    if not resource_name or resource_name.startswith("//"):
        if not path.is_absolute():
            return None

    if not path.is_absolute():
        path = signature_dir / path
    try:
        resolved = path.resolve(strict=True)
        resolved.relative_to(signature_dir.resolve())
    except (OSError, ValueError):
        return None
    return resolved if resolved.is_file() else None


def prepare_outlook_html_body(
    body: str,
    signature: OutlookSignature,
) -> tuple[str, list[tuple[Path, str]]]:
    body_markup = (
        '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">'
        f"{html.escape(body).replace(chr(10), '<br>')}"
        "</div><br>"
    )
    signature_markup = signature.markup
    resources: dict[Path, str] = {}

    def content_id_for(reference: str) -> str | None:
        path = _signature_resource_path(reference, signature.directory)
        if path is None:
            return None
        if path not in resources:
            resources[path] = f"{uuid.uuid4().hex}@outlook-draft-mailer"
        return resources[path]

    def replace_attribute(match: re.Match[str]) -> str:
        content_id = content_id_for(match.group(3))
        if content_id is None:
            return match.group(0)
        return f'{match.group(1)}={match.group(2)}cid:{content_id}{match.group(2)}'

    signature_markup = re.sub(
        r"""(?is)\b(src|background|poster)\s*=\s*(["'])(.*?)\2""",
        replace_attribute,
        signature_markup,
    )

    def replace_css_url(match: re.Match[str]) -> str:
        content_id = content_id_for(match.group(2))
        if content_id is None:
            return match.group(0)
        return f'url("cid:{content_id}")'

    signature_markup = re.sub(
        r"""(?is)url\(\s*(["']?)([^"')]+)\1\s*\)""",
        replace_css_url,
        signature_markup,
    )

    body_opening = re.search(r"<body\b[^>]*>", signature_markup, flags=re.I)
    if body_opening:
        insert_at = body_opening.end()
        html_body = (
            signature_markup[:insert_at]
            + body_markup
            + signature_markup[insert_at:]
        )
    else:
        html_opening = re.search(r"<html\b[^>]*>", signature_markup, flags=re.I)
        if html_opening:
            insert_at = html_opening.end()
            html_body = (
                signature_markup[:insert_at]
                + f"<body>{body_markup}"
                + signature_markup[insert_at:]
            )
        else:
            html_body = f"<html><body>{body_markup}{signature_markup}</body></html>"

    return html_body, [(path, content_id) for path, content_id in resources.items()]


def select_contacts(
    contacts: list[dict[str, str]],
    selected_indices: set[int],
) -> list[dict[str, str]]:
    return [
        contact
        for index, contact in enumerate(contacts)
        if index in selected_indices
    ]


def render_drafts(
    csv_path: Path,
    subject_template_path: Path,
    body_template_path: Path,
) -> list[dict[str, Any]]:
    _, contacts = load_contacts(csv_path)
    subject_template = subject_template_path.read_text(encoding="utf-8")
    body_template = body_template_path.read_text(encoding="utf-8")
    return render_messages(contacts, subject_template, body_template, "email")


def create_draft(message: dict[str, object]) -> str:
    try:
        import pywintypes
        from win32com.client import Dispatch
    except ImportError as error:
        raise RuntimeError(
            "The Outlook desktop integration is not installed. Run "
            "'python -m pip install -r requirements.txt'."
        ) from error

    try:
        outlook = Dispatch("Outlook.Application")
        outlook_message = outlook.CreateItem(0)
        outlook_message.To = message["toRecipients"][0]["emailAddress"]["address"]
        outlook_message.Subject = message["subject"]
        signature = message.get("outlookSignature")
        if isinstance(signature, OutlookSignature):
            html_body, resources = prepare_outlook_html_body(
                message["body"]["content"],
                signature,
            )
            outlook_message.HTMLBody = html_body
            for path, content_id in resources:
                attachment = outlook_message.Attachments.Add(str(path), 1)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    content_id,
                )
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
                    True,
                )
        else:
            outlook_message.Body = message["body"]["content"]
        outlook_message.Save()
        draft_id = outlook_message.EntryID
    except pywintypes.com_error as error:
        raise RuntimeError(
            "Could not create an Outlook draft. Make sure classic Outlook is "
            "installed and configured, then try again."
        ) from error

    if not draft_id:
        raise RuntimeError("Outlook saved the draft without returning an entry ID.")
    return draft_id


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create Outlook drafts from a Excel workbook (.xlsx) file; this script never sends mail."
    )
    parser.add_argument("--csv", type=Path, default=BASE_DIR / "contacts.csv")
    parser.add_argument(
        "--subject-template",
        type=Path,
        default=BASE_DIR / "subject_template.txt",
    )
    parser.add_argument(
        "--body-template",
        type=Path,
        default=BASE_DIR / "body_template.txt",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    drafts = render_drafts(args.csv, args.subject_template, args.body_template)

    created = 0
    for message in drafts:
        draft_id = create_draft(message)
        created += 1
        print(f"Created draft {created}/{len(drafts)} (ID: {draft_id})")

    print(f"Done. {created} draft(s) are in Outlook Drafts; none were sent.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, RuntimeError) as error:
        raise SystemExit(f"Error: {error}") from error
