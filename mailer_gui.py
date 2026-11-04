import os
import logging
import re
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from create_drafts import (
    BASE_DIR,
    OutlookSignature,
    create_draft,
    configure_logging,
    get_default_outlook_signature,
    load_contacts,
    render_messages,
    select_contacts,
)


class TkTextLogHandler(logging.Handler):
    def __init__(self, text_widget: tk.Text) -> None:
        super().__init__()
        self.text_widget = text_widget
        self.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s: %(message)s", "%H:%M:%S")
        )

    def emit(self, record: logging.LogRecord) -> None:
        message = self.format(record)
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", message + "\n")
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")
        self.text_widget.update_idletasks()


class DraftMailerWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Outlook Draft Mailer")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = min(1100, screen_width - 80)
        window_height = min(760, screen_height - 100)
        window_x = max(0, (screen_width - window_width) // 2)
        window_y = max(0, (screen_height - window_height) // 2)
        self.root.geometry(
            f"{window_width}x{window_height}+{window_x}+{window_y}"
        )
        self.root.minsize(min(800, window_width), min(600, window_height))
        self.input_path = tk.StringVar()
        self.recipient_column = tk.StringVar()
        self.contacts: list[dict[str, str]] = []
        self.columns: list[str] = []
        self.checked_rows: set[int] = set()
        self.active_editor: ttk.Entry | tk.Text | None = None
        self.outlook_signature: OutlookSignature | None = None
        self.log_visible = False
        self.geometry_before_log: tuple[int, int, int, int] | None = None

        self._build_window()
        if getattr(sys, "frozen", False):
            log_dir = Path(os.environ.get("LOCALAPPDATA", Path.home()))
            log_dir /= "Outlook Draft Mailer"
            log_dir.mkdir(parents=True, exist_ok=True)
        else:
            log_dir = BASE_DIR
        self.logger = configure_logging(log_dir / "mailer.log")
        self.logger.addHandler(TkTextLogHandler(self.log_text))
        self.logger.info("Outlook Draft Mailer started.")

    def _build_window(self) -> None:
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=2, minsize=120)
        main.rowconfigure(9, weight=1)
        main.rowconfigure(11, weight=0)

        ttk.Label(
            main,
            text="Choose a Excel (.xlsx) file. Its first row supplies the "
            "column names you can use.",
            wraplength=700,
        ).grid(row=0, column=0, sticky="w")

        file_row = ttk.Frame(main)
        file_row.grid(row=1, column=0, sticky="ew", pady=(8, 12))
        file_row.columnconfigure(0, weight=1)
        ttk.Entry(file_row, textvariable=self.input_path).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(file_row, text="Browse...", command=self._browse).grid(
            row=0, column=1, padx=(8, 0)
        )
        ttk.Button(file_row, text="Load columns", command=self._load_file).grid(
            row=0, column=2, padx=(8, 0)
        )

        recipient_row = ttk.Frame(main)
        recipient_row.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(recipient_row, text="Column containing recipient addresses:").pack(
            side="left"
        )
        self.recipient_dropdown = ttk.Combobox(
            recipient_row,
            textvariable=self.recipient_column,
            state="readonly",
            width=24,
        )
        self.recipient_dropdown.pack(side="left", padx=(8, 0))

        ttk.Label(
            main,
            text="Contact table — click the checkbox to include a row as a draft recipient:",
        ).grid(row=3, column=0, sticky="w")
        table_frame = ttk.Frame(main)
        table_frame.grid(row=4, column=0, sticky="nsew", pady=(4, 8))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.contact_table = ttk.Treeview(
            table_frame,
            show="headings",
            selectmode="browse",
            height=8,
        )
        self.contact_table.grid(row=0, column=0, sticky="nsew")
        self.contact_table.bind("<Button-1>", self._toggle_row)
        vertical_scrollbar = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.contact_table.yview
        )
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.contact_table.xview
        )
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.contact_table.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set,
        )

        selection_row = ttk.Frame(main)
        selection_row.grid(row=5, column=0, sticky="ew", pady=(0, 8))
        self.selection_label = ttk.Label(selection_row, text="No rows selected.")
        self.selection_label.pack(side="left", fill="x", expand=True)
        ttk.Button(
            selection_row,
            text="Select all",
            command=self._select_all,
        ).pack(side="right", padx=(8, 0))
        ttk.Button(
            selection_row,
            text="Clear selection",
            command=self._clear_selection,
        ).pack(side="right")

        ttk.Label(main, text="Available columns (double-click or select and insert):").grid(
            row=6, column=0, sticky="w"
        )
        columns_row = ttk.Frame(main)
        columns_row.grid(row=7, column=0, sticky="ew", pady=(4, 12))
        columns_row.columnconfigure(0, weight=1)
        self.column_list = tk.Listbox(columns_row, height=3, exportselection=False)
        self.column_list.grid(row=0, column=0, sticky="ew")
        self.column_list.bind("<Double-Button-1>", self._insert_placeholder)
        ttk.Button(
            columns_row,
            text="Insert {column}",
            command=self._insert_placeholder,
        ).grid(row=0, column=1, padx=(8, 0), sticky="n")

        subject_frame = ttk.LabelFrame(main, text="Subject", padding=8)
        subject_frame.grid(row=8, column=0, sticky="ew", pady=(0, 10))
        subject_frame.columnconfigure(0, weight=1)
        self.subject_entry = ttk.Entry(subject_frame)
        self.subject_entry.grid(row=0, column=0, sticky="ew")
        self.subject_entry.bind("<FocusIn>", self._remember_editor)

        body_frame = ttk.LabelFrame(main, text="Message body", padding=8)
        body_frame.grid(row=9, column=0, sticky="nsew")
        body_frame.rowconfigure(0, weight=1)
        body_frame.columnconfigure(0, weight=1)
        self.body_text = tk.Text(body_frame, height=5, wrap="word")
        self.body_text.grid(row=0, column=0, sticky="nsew")
        self.body_text.bind("<FocusIn>", self._remember_editor)
        self.body_text.bind("<KeyRelease>", self._highlight_placeholders)
        self.body_text.tag_configure(
            "placeholder",
            foreground="#0759A5",
            background="#E5F1FF",
        )
        self.active_editor = self.body_text
        ttk.Button(
            body_frame,
            text="Add Outlook signature",
            command=self._add_outlook_signature,
        ).grid(row=1, column=0, sticky="e", pady=(6, 0))

        action_row = ttk.Frame(main)
        action_row.grid(row=10, column=0, sticky="ew", pady=(12, 0))
        self.status_label = ttk.Label(action_row, text="Choose a contact file to begin.")
        self.status_label.pack(side="left", fill="x", expand=True)
        self.log_toggle_button = ttk.Button(
            action_row,
            text="Show log",
            command=self._toggle_log_panel,
        )
        self.log_toggle_button.pack(side="right", padx=(8, 0))
        ttk.Button(action_row, text="Preview first email", command=self._preview).pack(
            side="right", padx=(8, 0)
        )
        ttk.Button(action_row, text="Create Outlook drafts", command=self._create_drafts).pack(
            side="right"
        )

        log_frame = ttk.LabelFrame(main, text="Activity log", padding=8)
        self.log_frame = log_frame
        log_frame.grid(row=11, column=0, sticky="nsew", pady=(12, 0))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=4, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scrollbar = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        ttk.Button(
            log_frame,
            text="Clear display",
            command=self._clear_log_display,
        ).grid(row=1, column=0, sticky="e", pady=(6, 0))
        log_frame.grid_remove()

    def _toggle_log_panel(self) -> None:
        self.log_visible = not self.log_visible
        if self.log_visible:
            self.root.update_idletasks()
            self.geometry_before_log = (
                self.root.winfo_width(),
                self.root.winfo_height(),
                self.root.winfo_x(),
                self.root.winfo_y(),
            )
            self.log_frame.grid()
            self.log_toggle_button.config(text="Hide log")
            self.log_frame.master.rowconfigure(11, weight=1)
            screen_height = self.root.winfo_screenheight()
            expanded_height = min(
                self.geometry_before_log[1] + 220,
                screen_height - 80,
            )
            expanded_y = max(
                0,
                min(
                    self.geometry_before_log[3],
                    screen_height - expanded_height - 80,
                ),
            )
            self.root.geometry(
                f"{self.geometry_before_log[0]}x{expanded_height}"
                f"+{self.geometry_before_log[2]}+{expanded_y}"
            )
        else:
            self.log_frame.grid_remove()
            self.log_toggle_button.config(text="Show log")
            self.log_frame.master.rowconfigure(11, weight=0)
            if self.geometry_before_log is not None:
                width, height, x, y = self.geometry_before_log
                self.root.geometry(f"{width}x{height}+{x}+{y}")
                self.geometry_before_log = None

    def _clear_log_display(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _browse(self) -> None:
        selected = filedialog.askopenfilename(
            title="Choose contacts file",
            filetypes=[
                ("Excel workbooks", "*.xlsx"),
            ],
        )
        if selected:
            self.input_path.set(selected)
            self._load_file()

    def _load_file(self) -> None:
        path_text = self.input_path.get().strip()
        if not path_text:
            messagebox.showerror("Choose a file", "Select a Excel file first.")
            return
        try:
            columns, contacts = load_contacts(
                Path(path_text)
            )
        except (OSError, ValueError, RuntimeError) as error:
            self.logger.exception("Could not load contacts from %s.", path_text)
            messagebox.showerror("Could not load contacts", str(error))
            return

        self.columns = columns
        self.contacts = contacts
        self.checked_rows.clear()
        self.recipient_dropdown["values"] = columns
        if self.recipient_column.get() not in columns:
            self.recipient_column.set("")
        self.column_list.delete(0, tk.END)
        for column in columns:
            self.column_list.insert(tk.END, column)
        self._refresh_contact_table()
        self.status_label.config(
            text=(
                f"Loaded {len(contacts)} contact(s). "
                "Select the rows to make drafts for."
            )
        )
        self._update_selection_label()
        self.logger.info(
            "Loaded %d contact(s) with %d column(s) from %s.",
            len(contacts),
            len(columns),
            Path(path_text).name,
        )
        self.logger.info("Headers: %s", columns)

    def _refresh_contact_table(self) -> None:
        table_columns = ("include", *self.columns)
        self.contact_table.configure(columns=table_columns)
        self.contact_table.heading("include", text="Draft?")
        self.contact_table.column("include", width=60, minwidth=60, stretch=False, anchor="center")
        for column in self.columns:
            self.contact_table.heading(column, text=column.title())
            self.contact_table.column(column, width=150, minwidth=80, anchor="w")
        for item in self.contact_table.get_children():
            self.contact_table.delete(item)
        for index, contact in enumerate(self.contacts):
            values = ["☐", *(contact[column] for column in self.columns)]
            self.contact_table.insert("", "end", iid=str(index), values=values)

    def _toggle_row(self, event: tk.Event) -> str | None:
        if self.contact_table.identify_column(event.x) != "#1":
            return None
        item = self.contact_table.identify_row(event.y)
        if not item:
            return None
        index = int(item)
        if index in self.checked_rows:
            self.checked_rows.remove(index)
        else:
            self.checked_rows.add(index)
        self._refresh_row_checkbox(index)
        self._update_selection_label()
        return "break"

    def _refresh_row_checkbox(self, index: int) -> None:
        item = str(index)
        values = list(self.contact_table.item(item, "values"))
        values[0] = "☑" if index in self.checked_rows else "☐"
        self.contact_table.item(item, values=values)

    def _select_all(self) -> None:
        self.checked_rows = set(range(len(self.contacts)))
        for index in self.checked_rows:
            self._refresh_row_checkbox(index)
        self._update_selection_label()

    def _clear_selection(self) -> None:
        self.checked_rows.clear()
        for index in range(len(self.contacts)):
            self._refresh_row_checkbox(index)
        self._update_selection_label()

    def _update_selection_label(self) -> None:
        count = len(self.checked_rows)
        self.selection_label.config(
            text=f"{count} of {len(self.contacts)} contact(s) selected for drafts."
        )

    def _insert_placeholder(self, _event: tk.Event | None = None) -> None:
        selection = self.column_list.curselection()
        if not selection:
            messagebox.showinfo("Select a column", "Choose a column name first.")
            return

        placeholder = "{" + self.columns[selection[0]] + "}"
        if self.active_editor == self.subject_entry:
            position = self.subject_entry.index(tk.INSERT)
            self.subject_entry.insert(position, placeholder)
            self.subject_entry.focus_set()
        else:
            self.body_text.insert(tk.INSERT, placeholder)
            self._highlight_placeholders()
            self.body_text.focus_set()

    def _highlight_placeholders(self, _event: tk.Event | None = None) -> None:
        self.body_text.tag_remove("placeholder", "1.0", "end")
        body = self.body_text.get("1.0", "end-1c")
        for match in re.finditer(r"\{[^{}\r\n]+\}", body):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            self.body_text.tag_add("placeholder", start, end)

    def _add_outlook_signature(self) -> None:
        try:
            self.outlook_signature = get_default_outlook_signature()
        except (OSError, RuntimeError) as error:
            self.logger.exception("Could not load the default Outlook signature.")
            messagebox.showerror("Could not load Outlook signature", str(error))
            return

        self.status_label.config(
            text="The formatted Outlook signature will be included in created drafts."
        )
        self.logger.info(
            "Loaded the Outlook signature; original formatting and images will be retained."
        )

    def _remember_editor(self, event: tk.Event) -> None:
        self.active_editor = event.widget

    def _get_messages(self) -> list[dict[str, object]]:
        if not self.contacts:
            raise ValueError("Load a contact file before previewing or creating drafts.")
        if not self.checked_rows:
            raise ValueError("Check at least one contact row in the table.")
        recipient_column = self.recipient_column.get()
        if not recipient_column:
            raise ValueError("Choose which column contains recipient email addresses.")
        subject = self.subject_entry.get()
        body = self.body_text.get("1.0", "end-1c")
        selected_contacts = select_contacts(self.contacts, self.checked_rows)
        messages = render_messages(
            selected_contacts,
            subject,
            body,
            recipient_column,
        )
        if self.outlook_signature is not None:
            for message in messages:
                message["outlookSignature"] = self.outlook_signature
        return messages

    def _preview(self) -> None:
        try:
            messages = self._get_messages()
        except ValueError as error:
            self.logger.warning("Preview could not be created: %s", error)
            messagebox.showerror("Cannot preview", str(error))
            return

        first = messages[0]
        recipient = first["toRecipients"][0]["emailAddress"]["address"]
        self.logger.info("Previewed the first selected contact (%d total).", len(messages))
        messagebox.showinfo(
            "First email preview",
            f"To: {recipient}\n\n"
            f"Subject: {first['subject']}\n\n"
            f"{first['body']['content']}\n\n"
            f"Total contacts: {len(messages)}",
        )

    def _create_drafts(self) -> None:
        try:
            messages = self._get_messages()
        except ValueError as error:
            self.logger.warning("Draft creation validation failed: %s", error)
            messagebox.showerror("Cannot create drafts", str(error))
            return

        confirmed = messagebox.askyesno(
            "Create Outlook drafts?",
            f"This will save {len(messages)} email(s) in your Outlook Drafts folder. "
            "They will not be sent. Continue?",
        )
        if not confirmed:
            self.logger.info("Draft creation cancelled by the user.")
            return

        created = 0
        self.logger.info("Starting creation of %d Outlook draft(s).", len(messages))
        try:
            for message in messages:
                create_draft(message)
                created += 1
                self.logger.info("Saved draft %d of %d.", created, len(messages))
        except (OSError, RuntimeError) as error:
            self.logger.exception(
                "Draft creation stopped after %d of %d draft(s).",
                created,
                len(messages),
            )
            messagebox.showerror(
                "Draft creation stopped",
                f"Created {created} of {len(messages)} drafts before the error:\n\n{error}",
            )
            self.status_label.config(text=f"Stopped after creating {created} draft(s).")
            return

        self.status_label.config(text=f"Created {created} draft(s). None were sent.")
        self.logger.info("Finished: created %d draft(s); none were sent.", created)
        messagebox.showinfo(
            "Drafts created",
            f"Created {created} draft(s) in Outlook. None were sent.",
        )


def main() -> None:
    root = tk.Tk()
    DraftMailerWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
