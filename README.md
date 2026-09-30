# CAIE Playwright WhatsApp Automation

A Python Playwright automation project that sends personalized WhatsApp messages from an Excel contact list and generates JSON and Excel reports.

## Features

- Read contacts from `contacts.xlsx`
- Search WhatsApp contacts by name
- Send personalized messages
- Replace `{name}` in message templates
- Use a persistent WhatsApp Web browser profile
- Wait for WhatsApp Web to load before processing
- Add randomized delays between actions
- Confirm sent messages through the WhatsApp Web UI
- Capture screenshots of sent messages
- Extract the last 3 incoming messages
- Generate JSON and Excel reports
- Handle individual contact failures without stopping the entire process

## Tech Stack

- Python
- Playwright
- OpenPyXL
- WhatsApp Web

## Project Structure

```text
caie-playwright-whatsapp-automation/
│
├── playwright_assign.py
├── contacts.xlsx
├── requirements.txt
├── README.md
├── .gitignore
│
├── .whatsapp_profile/
├── screenshots/
└── debug/