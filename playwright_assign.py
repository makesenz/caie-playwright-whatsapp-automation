import json
import random
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

CONTACTS_FILE = BASE_DIR / "contacts.xlsx"
PROFILE_DIR = BASE_DIR / ".whatsapp_profile"
SCREENSHOT_DIR = BASE_DIR / "screenshots"
DEBUG_DIR = BASE_DIR / "debug"

DEFAULT_MESSAGE = "Hello {name}, hope you are doing well."

LOGIN_TIMEOUT_MS = 5 * 60 * 1000

MIN_DELAY_SECONDS = 2
MAX_DELAY_SECONDS = 5


# ============================================================
# UTILITY
# ============================================================

def pause(page, minimum=MIN_DELAY_SECONDS, maximum=MAX_DELAY_SECONDS):
    """Human-like random delay."""
    page.wait_for_timeout(
        random.randint(minimum * 1000, maximum * 1000)
    )


def clean_phone_number(phone_value):
    """
    Convert Excel/Numbers phone values into a clean WhatsApp number.

    Examples:
        919042512304.0 -> 919042512304
        +91 9042512304 -> 919042512304
        919042512304   -> 919042512304
    """

    if phone_value is None:
        return ""

    # Handle Excel numeric values such as 919042512304.0
    if isinstance(phone_value, float):
        if phone_value.is_integer():
            phone = str(int(phone_value))
        else:
            phone = str(phone_value)
    elif isinstance(phone_value, int):
        phone = str(phone_value)
    else:
        phone = str(phone_value).strip()

    # Remove everything except digits
    phone = re.sub(r"\D", "", phone)

    return phone


def safe_filename(value):
    """Create a safe filename."""
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", value)
    cleaned = cleaned.strip("_")

    return cleaned or "contact"


# ============================================================
# LOAD CONTACTS
# ============================================================

def load_contacts():

    if not CONTACTS_FILE.exists():
        raise FileNotFoundError(
            f"Contacts workbook not found: {CONTACTS_FILE}"
        )

    workbook = load_workbook(
        CONTACTS_FILE,
        data_only=True,
        read_only=True
    )

    try:

        sheet = workbook.active

        rows = sheet.iter_rows(values_only=True)

        headers = next(rows, None)

        if not headers:
            raise ValueError("contacts.xlsx is empty")

        # Case-insensitive column mapping
        columns = {
            str(value).strip().lower(): index
            for index, value in enumerate(headers)
            if value is not None
        }

        required = {"name", "phone", "message"}

        missing = required - columns.keys()

        if missing:
            raise ValueError(
                f"Missing required columns: {', '.join(sorted(missing))}"
            )

        contacts = []

        for row_number, row in enumerate(rows, start=2):

            # Skip completely empty rows
            if not any(value is not None for value in row):
                continue

            name_value = (
                row[columns["name"]]
                if columns["name"] < len(row)
                else None
            )

            phone_value = (
                row[columns["phone"]]
                if columns["phone"] < len(row)
                else None
            )

            message_value = (
                row[columns["message"]]
                if columns["message"] < len(row)
                else None
            )

            name = (
                str(name_value).strip()
                if name_value is not None
                else ""
            )

            phone = clean_phone_number(phone_value)

            template = (
                str(message_value).strip()
                if message_value is not None
                else ""
            )

            # Personalize message
            message = (
                template or DEFAULT_MESSAGE
            ).replace("{name}", name)

            # Ignore completely empty rows
            if not name and not phone:
                continue

            contacts.append(
                {
                    "name": name,
                    "phone": phone,
                    "message": message,
                    "source_row": row_number,
                }
            )

        return contacts

    finally:

        workbook.close()


# ============================================================
# PLAYWRIGHT HELPERS
# ============================================================

def first_visible(locator):

    for index in range(locator.count()):

        candidate = locator.nth(index)

        try:
            if candidate.is_visible():
                return candidate
        except Exception:
            continue

    return None


def dismiss_whats_new_dialog(page):

    dialogs = page.get_by_role("dialog")

    for index in range(dialogs.count()):

        dialog = dialogs.nth(index)

        try:
            if not dialog.is_visible():
                continue
        except Exception:
            continue

        label_id = dialog.get_attribute("aria-labelledby")

        heading = ""

        if label_id:

            try:
                heading = (
                    page.locator(f"#{label_id}")
                    .inner_text()
                    .strip()
                )
            except Exception:
                pass

        if not (
            heading.casefold().startswith("what")
            and "new on whatsapp web" in heading.casefold()
        ):
            continue

        button = first_visible(
            dialog.get_by_role(
                "button",
                name="Continue",
                exact=True
            )
        )

        if button is None:

            button = first_visible(
                dialog.get_by_role(
                    "button",
                    name="Close",
                    exact=True
                )
            )

        if button:

            button.click()

            pause(page, 1, 2)


# ============================================================
# SEARCH
# ============================================================

def search_box(page):

    selectors = [

        # Current / common WhatsApp selectors
        'div[contenteditable="true"][data-tab="3"]',

        'input[data-tab="3"]',

        '[role="textbox"][aria-label*="Search" i]',

        'input[placeholder*="Search" i]',

        '[aria-label*="Search or start a new chat" i]',

    ]

    for selector in selectors:

        try:

            found = first_visible(
                page.locator(selector)
            )

            if found:
                return found

        except Exception:
            continue

    return None


def open_contact(page, contact):

    name = contact["name"].strip()

    if not name:
        raise ValueError(
            f"Row {contact['source_row']}: contact name is required"
        )

    # Click search button if available
    search_button = first_visible(
        page.get_by_role(
            "button",
            name=re.compile("search", re.IGNORECASE)
        )
    )

    if search_button:

        try:
            search_button.click()
            pause(page, 1, 2)
        except Exception:
            pass

    box = search_box(page)

    if box is None:
        raise RuntimeError(
            "WhatsApp search field was not found"
        )

    # Clear previous search
    try:
        box.fill("")
    except Exception:
        pass

    box.fill(name)

    pause(page, 2, 3)

    # WhatsApp chat search results
    matches = page.locator(
        '#pane-side '
        '[data-testid="cell-frame-container"] '
        '[data-testid="cell-frame-title"]'
    )

    visible_matches = []

    for index in range(matches.count()):

        candidate = matches.nth(index)

        try:

            if not candidate.is_visible():
                continue

            text = candidate.inner_text().strip()

            if text == name:
                visible_matches.append(candidate)

        except Exception:
            continue

    if not visible_matches:

        raise LookupError(
            f"No matching chat found for contact name: {name}"
        )

    if len(visible_matches) > 1:

        raise LookupError(
            f"Multiple chats match contact name "
            f"{name}; refusing to guess"
        )

    visible_matches[0].click()

    # Wait for chat composer
    page.wait_for_selector(
        '[contenteditable="true"][data-tab="10"], '
        '[contenteditable="true"][aria-label*="message" i], '
        '[role="textbox"][title*="message" i], '
        'footer [contenteditable="true"]',
        state="visible",
        timeout=15000,
    )


# ============================================================
# MESSAGE COMPOSER
# ============================================================

def message_box(page):

    selectors = [

        '[contenteditable="true"][data-tab="10"]',

        '[contenteditable="true"][aria-label*="message" i]',

        '[role="textbox"][title*="message" i]',

        'footer [contenteditable="true"]',

        'footer div[contenteditable="true"]',

    ]

    for selector in selectors:

        try:

            found = first_visible(
                page.locator(selector)
            )

            if found:
                return found

        except Exception:
            continue

    raise RuntimeError(
        "WhatsApp message composer was not found"
    )


# ============================================================
# MESSAGE DETECTION
# ============================================================

def find_message_text(page, message):

    """
    Find exact message text anywhere in the active chat.

    We intentionally do NOT depend on div.message-out
    because WhatsApp changes its DOM structure frequently.
    """

    selectors = [

        # WhatsApp selectable message text
        'span.selectable-text',

        # Generic text inside message containers
        '[data-pre-plain-text] span.selectable-text',

        # Generic message rows
        '[role="row"] span.selectable-text',

        # Fallback
        'div.copyable-text span.selectable-text',

    ]

    for selector in selectors:

        try:

            locator = page.locator(selector).filter(
                has_text=message
            )

            for index in range(locator.count()):

                candidate = locator.nth(index)

                if candidate.is_visible():

                    try:

                        text = candidate.inner_text().strip()

                        if text == message:

                            return candidate

                    except Exception:
                        continue

        except Exception:
            continue

    return None


def wait_for_sent_message(page, message, timeout=20000):

    """
    Wait until the exact sent message appears.

    We check several possible WhatsApp DOM structures
    instead of relying on one fragile class.
    """

    deadline = datetime.now().timestamp() + (
        timeout / 1000
    )

    while datetime.now().timestamp() < deadline:

        found = find_message_text(
            page,
            message
        )

        if found:

            return found

        page.wait_for_timeout(500)

    raise PlaywrightTimeoutError(
        f"Could not confirm sent message in WhatsApp UI: "
        f"{message}"
    )


# ============================================================
# LAST 3 INCOMING MESSAGES
# ============================================================

def extract_last_messages(page):

    """
    Extract the last 3 messages received from the contact.

    We primarily use message-in.
    """

    selectors = [

        'div.message-in span.selectable-text',

        '[data-pre-plain-text] span.selectable-text',

    ]

    extracted = []

    for selector in selectors:

        try:

            messages = page.locator(selector)

            for index in range(messages.count()):

                candidate = messages.nth(index)

                try:

                    if not candidate.is_visible():
                        continue

                    text = candidate.inner_text().strip()

                    if text and text not in extracted:

                        extracted.append(text)

                except Exception:
                    continue

        except Exception:
            continue

        # If we already found enough
        if len(extracted) >= 3:
            break

    return extracted[-3:]


# ============================================================
# DEBUG
# ============================================================

def save_debug_screenshot(page, contact_name):

    try:

        DEBUG_DIR.mkdir(
            parents=True,
            exist_ok=True
        )

        path = DEBUG_DIR / (
            f"{safe_filename(contact_name)}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )

        page.screenshot(
            path=str(path),
            full_page=False
        )

        return str(
            path.relative_to(BASE_DIR)
        )

    except Exception:
        return None


# ============================================================
# AUTOMATE ONE CONTACT
# ============================================================

def automate_contact(page, contact):

    record = {

        "name": contact["name"],

        "phone": contact["phone"],

        "sent_message": contact["message"],

        "status": "failed",

        "sent_at": None,

        "screenshot": None,

        "last_3_messages": [],

        "error": None,
    }

    try:

        # ----------------------------------------------------
        # Validate contact
        # ----------------------------------------------------

        if not contact["name"]:

            raise ValueError(
                f"Row {contact['source_row']}: "
                f"contact name is required"
            )

        if not contact["phone"]:

            raise ValueError(
                f"Row {contact['source_row']}: "
                f"phone number is required"
            )

        if not 8 <= len(contact["phone"]) <= 15:

            raise ValueError(
                f"Invalid phone number for "
                f"{contact['name']}: "
                f"{contact['phone']}"
            )

        # ----------------------------------------------------
        # Open WhatsApp chat
        # ----------------------------------------------------

        print(
            f"\nProcessing: {contact['name']}"
        )

        print(
            f"Phone: {contact['phone']}"
        )

        open_contact(
            page,
            contact
        )

        pause(page)

        # ----------------------------------------------------
        # Find composer
        # ----------------------------------------------------

        composer = message_box(page)

        # ----------------------------------------------------
        # Type message
        # ----------------------------------------------------

        composer.fill(
            contact["message"]
        )

        pause(page)

        print(
            "Sending message..."
        )

        # ----------------------------------------------------
        # Send
        # ----------------------------------------------------

        composer.press("Enter")

        # Small wait for WhatsApp to update DOM
        page.wait_for_timeout(1000)

        # ----------------------------------------------------
        # Confirm sent message
        # ----------------------------------------------------

        try:

            sent_bubble = wait_for_sent_message(
                page,
                contact["message"],
                timeout=20000
            )

        except PlaywrightTimeoutError:

            # Important:
            # Message may already have been sent.
            # We DO NOT retry automatically.

            debug_path = save_debug_screenshot(
                page,
                contact["name"]
            )

            record["status"] = "unconfirmed"

            record["error"] = (
                "Message may have been sent, "
                "but WhatsApp UI confirmation "
                "could not be detected."
            )

            if debug_path:

                record["debug_screenshot"] = debug_path

            print(
                "WARNING: Message send could not be confirmed."
            )

            return record

        # ----------------------------------------------------
        # Mark successful
        # ----------------------------------------------------

        record["status"] = "sent"

        record["sent_at"] = (
            datetime.now()
            .astimezone()
            .isoformat(timespec="seconds")
        )

        print(
            "Message sent successfully."
        )

        # ----------------------------------------------------
        # Screenshot
        # ----------------------------------------------------

        SCREENSHOT_DIR.mkdir(
            parents=True,
            exist_ok=True
        )

        screenshot_path = (
            SCREENSHOT_DIR
            / f"{safe_filename(contact['name'])}_"
              f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )

        # Screenshot the message if possible
        try:

            sent_bubble.screenshot(
                path=str(screenshot_path)
            )

        except Exception:

            page.screenshot(
                path=str(screenshot_path),
                full_page=False
            )

        record["screenshot"] = str(
            screenshot_path.relative_to(BASE_DIR)
        )

        # ----------------------------------------------------
        # Wait before extracting
        # ----------------------------------------------------

        pause(page)

        # ----------------------------------------------------
        # Last 3 incoming messages
        # ----------------------------------------------------

        record["last_3_messages"] = (
            extract_last_messages(page)
        )

        print(
            "Last 3 messages extracted."
        )

    except (
        PlaywrightTimeoutError,
        LookupError,
        RuntimeError,
        ValueError,
    ) as error:

        record["error"] = str(error)

        print(
            f"FAILED: {contact['name']}"
        )

        print(
            f"Error: {error}"
        )

        debug_path = save_debug_screenshot(
            page,
            contact["name"]
        )

        if debug_path:

            record["debug_screenshot"] = debug_path

    except Exception as error:

        record["error"] = (
            f"{type(error).__name__}: {error}"
        )

        print(
            f"FAILED: {contact['name']}"
        )

        print(
            f"Error: {error}"
        )

        debug_path = save_debug_screenshot(
            page,
            contact["name"]
        )

        if debug_path:

            record["debug_screenshot"] = debug_path

    return record


# ============================================================
# SAVE REPORTS
# ============================================================

def save_reports(records, report_date):

    json_path = (
        BASE_DIR
        / f"whatsapp_report_{report_date}.json"
    )

    excel_path = (
        BASE_DIR
        / f"whatsapp_report_{report_date}.xlsx"
    )

    # --------------------------------------------------------
    # JSON
    # --------------------------------------------------------

    with json_path.open(
        "w",
        encoding="utf-8"
    ) as report_file:

        json.dump(
            records,
            report_file,
            ensure_ascii=False,
            indent=2
        )

    # --------------------------------------------------------
    # Excel
    # --------------------------------------------------------

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "WhatsApp Report"

    sheet.append(
        [
            "Name",
            "Phone",
            "Status",
            "Sent At",
            "Sent Message",
            "Last 3 Messages",
            "Screenshot",
            "Error",
        ]
    )

    for record in records:

        sheet.append(
            [
                record.get("name"),
                record.get("phone"),
                record.get("status"),
                record.get("sent_at"),
                record.get("sent_message"),
                " | ".join(
                    record.get(
                        "last_3_messages",
                        []
                    )
                ),
                record.get("screenshot"),
                record.get("error"),
            ]
        )

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = sheet.dimensions

    # Column widths
    widths = {
        "A": 28,
        "B": 18,
        "C": 15,
        "D": 25,
        "E": 50,
        "F": 60,
        "G": 45,
        "H": 60,
    }

    for column, width in widths.items():

        sheet.column_dimensions[column].width = width

    workbook.save(excel_path)

    print(
        f"\nJSON report: {json_path}"
    )

    print(
        f"Excel report: {excel_path}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    # --------------------------------------------------------
    # Load contacts
    # --------------------------------------------------------

    try:

        contacts = load_contacts()

    except Exception as error:

        print(
            f"Cannot read contacts: {error}",
            file=sys.stderr
        )

        return 1

    if not contacts:

        print(
            "No contacts found in contacts.xlsx",
            file=sys.stderr
        )

        return 1

    print(
        f"\nLoaded {len(contacts)} contact(s)."
    )

    for contact in contacts:

        print(
            f"  - {contact['name']} "
            f"| {contact['phone']}"
        )

    # --------------------------------------------------------
    # Report
    # --------------------------------------------------------

    report_date = datetime.now().strftime(
        "%Y-%m-%d"
    )

    records = []

    login_error = None

    # --------------------------------------------------------
    # Start Playwright
    # --------------------------------------------------------

    try:

        with sync_playwright() as playwright:

            context = (
                playwright.chromium
                .launch_persistent_context(
                    user_data_dir=str(
                        PROFILE_DIR
                    ),
                    headless=False,
                )
            )

            try:

                page = (
                    context.pages[0]
                    if context.pages
                    else context.new_page()
                )

                page.goto(
                    "https://web.whatsapp.com",
                    wait_until="domcontentloaded"
                )

                print(
                    "\nWhatsApp Web opened."
                )

                print(
                    "If this is the first run, "
                    "scan the WhatsApp Web QR code."
                )

                # ------------------------------------------------
                # Wait for WhatsApp to be ready
                # ------------------------------------------------

                print(
                    "Waiting for WhatsApp Web..."
                )

                page.wait_for_selector(
                    "#pane-side",
                    state="visible",
                    timeout=LOGIN_TIMEOUT_MS
                )

                print(
                    "WhatsApp Web is ready."
                )

                dismiss_whats_new_dialog(
                    page
                )

                # ------------------------------------------------
                # Process contacts
                # ------------------------------------------------

                for index, contact in enumerate(
                    contacts,
                    start=1
                ):

                    print(
                        "\n"
                        + "=" * 60
                    )

                    print(
                        f"Contact {index}/"
                        f"{len(contacts)}"
                    )

                    print(
                        f"Name: {contact['name']}"
                    )

                    print(
                        f"Phone: {contact['phone']}"
                    )

                    print(
                        "=" * 60
                    )

                    record = automate_contact(
                        page,
                        contact
                    )

                    records.append(
                        record
                    )

                    print(
                        f"\nResult: "
                        f"{record['status']}"
                    )

                    # Delay between contacts
                    if index < len(contacts):

                        pause(
                            page,
                            3,
                            6
                        )

            except Exception as error:

                login_error = (
                    f"{type(error).__name__}: "
                    f"{error}"
                )

            finally:

                context.close()

    except Exception as error:

        login_error = (
            f"{type(error).__name__}: "
            f"{error}"
        )

    # --------------------------------------------------------
    # Handle login/browser-level errors
    # --------------------------------------------------------

    if login_error:

        completed_rows = {
            record.get("phone")
            for record in records
        }

        for contact in contacts:

            if contact["phone"] not in completed_rows:

                records.append(
                    {
                        "name": contact["name"],
                        "phone": contact["phone"],
                        "sent_message": contact["message"],
                        "status": "failed",
                        "sent_at": None,
                        "screenshot": None,
                        "last_3_messages": [],
                        "error": login_error,
                    }
                )

    # --------------------------------------------------------
    # Save reports
    # --------------------------------------------------------

    save_reports(
        records,
        report_date
    )

    # --------------------------------------------------------
    # Final summary
    # --------------------------------------------------------

    sent_count = sum(
        1
        for record in records
        if record["status"] == "sent"
    )

    failed_count = sum(
        1
        for record in records
        if record["status"] == "failed"
    )

    unconfirmed_count = sum(
        1
        for record in records
        if record["status"] == "unconfirmed"
    )

    print(
        "\n"
        + "=" * 60
    )

    print(
        "FINAL SUMMARY"
    )

    print(
        "=" * 60
    )

    print(
        f"Total      : {len(records)}"
    )

    print(
        f"Sent       : {sent_count}"
    )

    print(
        f"Failed     : {failed_count}"
    )

    print(
        f"Unconfirmed: {unconfirmed_count}"
    )

    print(
        "=" * 60
    )

    # Return success only when everything is confirmed sent
    return (
        0
        if len(records) > 0
        and all(
            record["status"] == "sent"
            for record in records
        )
        else 1
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    raise SystemExit(
        main()
    )