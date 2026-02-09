# Camping Reservation Management System (GAS)

This project is a Google Apps Script (GAS) based reservation management system for campsites. It supports automatic reservation extraction from Rakuten Travel and Nap (なっぷ) emails.

## Features
- **Automatic Email Parsing**: Extracts reservation details from Gmail and saves them to Google Sheets.
- **Calendar Integration**: Synchronizes reservations with Google Calendar.
- **Automated Emails**: Sends reminder and info emails to customers.
- **Web Management Interface**: A management dashboard built with GAS Web Apps.

## Supported Platforms
- Rakuten Travel (楽天トラベルキャンプ)
- Nap (なっぷ)

## Project Structure
- `config.js`: Application configuration and platform settings.
- `gmail_to_sheets.js`: Email extraction logic.
- `sync_calendar.js`: Calendar synchronization logic.
- `mail_to_client.js`: Automated email sending logic.
- `webapp.js`: Backend for the Web App.
- `index.html`: Frontend for the Web App.
- `debug_tool.js`: Utilities for testing and debugging.

## Setup
1. Clone this repository.
2. Use [clasp](https://github.com/google/clasp) to push the code to your GAS project.
3. Configure `config.js` with your Spreadsheet ID and Calendar ID.
4. Set up triggers for `mainSequence` and `sendAllReminders`.
