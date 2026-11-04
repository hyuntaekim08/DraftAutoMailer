import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from create_drafts import (
    configure_logging,
    create_draft,
    get_default_outlook_signature,
    load_contacts,
    OutlookSignature,
    prepare_outlook_html_body,
    render_drafts,
    render_messages,
    select_contacts,
)


class OutlookSignatureTests(unittest.TestCase):
    @patch("create_drafts.os.environ", {"APPDATA": "C:\\Users\\test\\AppData\\Roaming"})
    @patch("create_drafts.winreg.QueryValueEx", return_value=("Work", 1))
    @patch("create_drafts.winreg.OpenKey")
    def test_reads_default_signature_text_file(
        self,
        open_key: Mock,
        query_value: Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            signature_dir = Path(directory) / "Microsoft" / "Signatures"
            signature_dir.mkdir(parents=True)
            (signature_dir / "Work.txt").write_text(
                "Best,\nAlex", encoding="utf-8"
            )
            with patch.dict(
                "create_drafts.os.environ",
                {"APPDATA": str(Path(directory))},
            ):
                signature = get_default_outlook_signature()

        self.assertIn("Best,", signature.markup)
        self.assertEqual(signature.directory, Path(directory) / "Microsoft" / "Signatures")
        query_value.assert_called_once()

    @patch("create_drafts.winreg.QueryValueEx", side_effect=FileNotFoundError)
    @patch("create_drafts.winreg.OpenKey")
    def test_uses_only_local_signature_when_no_default_is_configured(
        self,
        open_key: Mock,
        query_value: Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            signature_dir = Path(directory) / "Microsoft" / "Signatures"
            signature_dir.mkdir(parents=True)
            (signature_dir / "Work.txt").write_text(
                "Best,\nAlex", encoding="utf-8"
            )
            with patch.dict(
                "create_drafts.os.environ",
                {"APPDATA": str(Path(directory))},
            ):
                signature = get_default_outlook_signature()

        self.assertIn("Best,", signature.markup)
        query_value.assert_called()

    @patch("create_drafts.winreg.QueryValueEx", return_value=("Work", 1))
    @patch("create_drafts.winreg.OpenKey")
    def test_reads_utf16_outlook_signature_file(
        self,
        open_key: Mock,
        query_value: Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            signature_dir = Path(directory) / "Microsoft" / "Signatures"
            signature_dir.mkdir(parents=True)
            (signature_dir / "Work.txt").write_text(
                "Best,\nAlex", encoding="utf-16"
            )
            with patch.dict(
                "create_drafts.os.environ",
                {"APPDATA": str(Path(directory))},
            ):
                signature = get_default_outlook_signature()

        self.assertIn("Best,", signature.markup)

    @patch("create_drafts.winreg.QueryValueEx", return_value=("Work", 1))
    @patch("create_drafts.winreg.OpenKey")
    def test_reads_legacy_windows_encoded_html_signature(
        self,
        open_key: Mock,
        query_value: Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            signature_dir = Path(directory) / "Microsoft" / "Signatures"
            signature_dir.mkdir(parents=True)
            (signature_dir / "Work.htm").write_bytes(
                '<html><body>André</body></html>'.encode("cp1252")
            )
            with patch.dict(
                "create_drafts.os.environ",
                {"APPDATA": str(Path(directory))},
            ):
                signature = get_default_outlook_signature()

        self.assertIn("André", signature.markup)

    def test_prepares_html_signature_and_embeds_local_images(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            signature_dir = Path(directory)
            image_dir = signature_dir / "Work_files"
            image_dir.mkdir()
            image_path = image_dir / "logo.png"
            image_path.write_bytes(b"image")
            signature = OutlookSignature(
                '<html><head><style>.name { color: #123456; }</style></head>'
                '<body><span class="name">Alex</span>'
                '<img src="Work_files/logo.png"></body></html>',
                signature_dir,
            )

            body, resources = prepare_outlook_html_body("Hello Alex.", signature)

        self.assertIn("color: #123456", body)
        self.assertIn("<span class=\"name\">Alex</span>", body)
        self.assertIn("Hello Alex.", body)
        self.assertEqual(len(resources), 1)
        self.assertEqual(resources[0][0], image_path.resolve())
        self.assertIn(f"cid:{resources[0][1]}", body)

    @patch("create_drafts.winreg.QueryValueEx", side_effect=FileNotFoundError)
    @patch("create_drafts.winreg.OpenKey")
    def test_requires_default_when_multiple_local_signatures_exist(
        self,
        open_key: Mock,
        query_value: Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            signature_dir = Path(directory) / "Microsoft" / "Signatures"
            signature_dir.mkdir(parents=True)
            (signature_dir / "Work.txt").write_text("Work", encoding="utf-8")
            (signature_dir / "Personal.htm").write_text(
                "<p>Personal</p>", encoding="utf-8"
            )
            with patch.dict(
                "create_drafts.os.environ",
                {"APPDATA": str(Path(directory))},
            ):
                with self.assertRaisesRegex(RuntimeError, "multiple signatures"):
                    get_default_outlook_signature()


class RenderDraftsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp_dir.name)
        self.csv_path = self.directory / "contacts.csv"
        self.subject_path = self.directory / "subject.txt"
        self.body_path = self.directory / "body.txt"

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_writes_activity_to_log_file(self) -> None:
        log_path = self.directory / "mailer.log"
        logger = configure_logging(log_path)
        try:
            logger.info("Loaded 3 contacts.")
            for handler in logger.handlers:
                handler.flush()
        finally:
            for handler in logger.handlers[:]:
                logger.removeHandler(handler)
                handler.close()

        log_text = log_path.read_text(encoding="utf-8")
        self.assertIn("INFO", log_text)
        self.assertIn("Loaded 3 contacts.", log_text)

    def test_renders_personalized_message(self) -> None:
        self.csv_path.write_text(
            "email,first_name,company\nalex@example.com,Alex,Example Co\n",
            encoding="utf-8",
        )
        self.subject_path.write_text("Hello {first_name}", encoding="utf-8")
        self.body_path.write_text("Welcome to {company}.", encoding="utf-8")

        drafts = render_drafts(self.csv_path, self.subject_path, self.body_path)

        self.assertEqual(drafts[0]["subject"], "Hello Alex")
        self.assertEqual(
            drafts[0]["body"],
            {"contentType": "Text", "content": "Welcome to Example Co."},
        )
        self.assertEqual(
            drafts[0]["toRecipients"],
            [{"emailAddress": {"address": "alex@example.com"}}],
        )

    def test_loads_contacts_without_a_fixed_email_column(self) -> None:
        self.csv_path.write_text("first_name\nAlex\n", encoding="utf-8")

        columns, contacts = load_contacts(self.csv_path)

        self.assertEqual(columns, ["first_name"])
        self.assertEqual(contacts, [{"first_name": "Alex"}])


    def test_headerless_xlsx_uses_first_row_as_data(self) -> None:
        from openpyxl import Workbook

        workbook_path = self.directory / "headerless.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["alex@example.com", "Alex"])
        worksheet.append(["sam@example.com", "Sam"])
        workbook.save(workbook_path)
        workbook.close()

        columns, contacts = load_contacts(workbook_path)

        self.assertEqual(columns, ["A", "B"])
        self.assertEqual(
            contacts[0],
            {"A": "alex@example.com", "B": "Alex"},
        )

    def test_rejects_template_field_not_in_csv(self) -> None:
        self.csv_path.write_text("email\nalex@example.com\n", encoding="utf-8")
        self.subject_path.write_text("Hello {first_name}", encoding="utf-8")
        self.body_path.write_text("Message", encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "first_name"):
            render_drafts(self.csv_path, self.subject_path, self.body_path)

    def test_rejects_short_csv_row(self) -> None:
        self.csv_path.write_text(
            "email,first_name\nalex@example.com\n",
            encoding="utf-8",
        )
        self.subject_path.write_text("Hello", encoding="utf-8")
        self.body_path.write_text("Message", encoding="utf-8")

        with self.assertRaisesRegex(ValueError, "first_name"):
            render_drafts(self.csv_path, self.subject_path, self.body_path)

    def test_loads_xlsx_columns_and_contacts(self) -> None:
        from openpyxl import Workbook

        workbook_path = self.directory / "contacts.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["email", "first_name"])
        worksheet.append(["alex@example.com", "Alex"])
        workbook.save(workbook_path)
        workbook.close()

        columns, contacts = load_contacts(workbook_path)

        self.assertEqual(columns, ["email", "first_name"])
        self.assertEqual(
            contacts,
            [{"email": "alex@example.com", "first_name": "Alex"}],
        )

    def test_preserves_excel_header_text_for_placeholders(self) -> None:
        from openpyxl import Workbook

        workbook_path = self.directory / "contacts.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append([" Mail ", "First Name"])
        worksheet.append(["alex@example.com", "Alex"])
        workbook.save(workbook_path)
        workbook.close()

        columns, contacts = load_contacts(workbook_path)
        messages = render_messages(
            contacts,
            "Hello {First Name}",
            "Contact: { Mail }",
            " Mail ",
        )

        self.assertEqual(columns, [" Mail ", "First Name"])
        self.assertEqual(messages[0]["subject"], "Hello Alex")
        self.assertEqual(messages[0]["body"]["content"], "Contact: alex@example.com")

    def test_renders_messages_from_gui_templates(self) -> None:
        messages = render_messages(
            [{"contact": "alex@example.com", "first_name": "Alex"}],
            "Hello {first_name}",
            "Welcome, {first_name}.",
            "contact",
        )

        self.assertEqual(messages[0]["subject"], "Hello Alex")
        self.assertEqual(messages[0]["body"]["content"], "Welcome, Alex.")

    def test_uses_selected_column_as_recipient(self) -> None:
        messages = render_messages(
            [{"contact": "alex@example.com", "name": "Alex"}],
            "Hello {name}",
            "Welcome, {name}.",
            "contact",
        )

        self.assertEqual(
            messages[0]["toRecipients"],
            [{"emailAddress": {"address": "alex@example.com"}}],
        )

    def test_selects_only_checked_contacts_in_file_order(self) -> None:
        contacts = [
            {"mail": "first@example.com"},
            {"mail": "second@example.com"},
            {"mail": "third@example.com"},
        ]

        selected = select_contacts(contacts, {0, 2})

        self.assertEqual(
            selected,
            [{"mail": "first@example.com"}, {"mail": "third@example.com"}],
        )

    @patch("win32com.client.Dispatch")
    def test_saves_outlook_draft_without_sending(self, dispatch: Mock) -> None:
        outlook_message = dispatch.return_value.CreateItem.return_value
        outlook_message.EntryID = "draft-id"
        message = {
            "subject": "Hello Alex",
            "body": {"contentType": "Text", "content": "Welcome."},
            "toRecipients": [
                {"emailAddress": {"address": "alex@example.com"}},
            ],
        }

        draft_id = create_draft(message)

        self.assertEqual(draft_id, "draft-id")
        dispatch.assert_called_once_with("Outlook.Application")
        dispatch.return_value.CreateItem.assert_called_once_with(0)
        self.assertEqual(outlook_message.To, "alex@example.com")
        self.assertEqual(outlook_message.Subject, "Hello Alex")
        self.assertEqual(outlook_message.Body, "Welcome.")
        outlook_message.Save.assert_called_once_with()
        outlook_message.Send.assert_not_called()

    @patch("win32com.client.Dispatch")
    def test_saves_signature_as_formatted_html_with_inline_image(
        self,
        dispatch: Mock,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            signature_dir = Path(directory)
            image_path = signature_dir / "logo.png"
            image_path.write_bytes(b"image")
            outlook_message = dispatch.return_value.CreateItem.return_value
            outlook_message.EntryID = "draft-id"
            message = {
                "subject": "Hello Alex",
                "body": {"contentType": "Text", "content": "Welcome."},
                "toRecipients": [
                    {"emailAddress": {"address": "alex@example.com"}},
                ],
                "outlookSignature": OutlookSignature(
                    '<html><body><span style="color: red">Best, Alex</span>'
                    '<img src="logo.png"></body></html>',
                    signature_dir,
                ),
            }

            draft_id = create_draft(message)

        self.assertEqual(draft_id, "draft-id")
        self.assertIn("color: red", outlook_message.HTMLBody)
        self.assertIn("Welcome.", outlook_message.HTMLBody)
        outlook_message.Attachments.Add.assert_called_once_with(str(image_path.resolve()), 1)
        attachment = outlook_message.Attachments.Add.return_value
        self.assertEqual(attachment.PropertyAccessor.SetProperty.call_count, 2)
        outlook_message.Save.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
