# ALICE Lite

### Automated Calling Assistant

ALICE Lite is a Python-based outbound calling automation system designed for:

* BizPhone (LDPlayer)
* Zentrex

It automates:

* outbound dialing
* lead management
* call disposition logging
* follow-up tracking
* remarks entry

using Excel as a lightweight CRM.

---

# Features

## Core Features

* Automated outbound dialing
* Excel lead management
* Auto status logging
* Follow-up remarks tracking
* Date logging
* Redial previous contact
* Pause / Resume support
* Multi-platform dialer support

---

# Supported Dialers

| Dialer   | Support |
| -------- | ------- |
| BizPhone | ✅       |
| Zentrex  | ✅       |

---

# Project Structure

```text
ALICE/
│
├── bizphone.py
├── zentrex.py
├── Calling.xlsx
└── README.md
```

---

# How It Works

## Workflow

```text
Press .
↓
Switch to Dialer
↓
Dial Number
↓
Switch Back to Terminal
↓
Enter Call Status
↓
Update Excel
```

---

# Excel Format

| Column | Description  |
| ------ | ------------ |
| A      | Phone Number |
| B      | Call Status  |
| C      | Date         |
| D      | Remarks      |

Example:

| Phone Number | Status    | Date       | Remarks               |
| ------------ | --------- | ---------- | --------------------- |
| 91234567     | Follow Up | 24/05/2026 | Call back next Monday |
| 98765432     | DNC       | 24/05/2026 |                       |

---

# Installation

## 1. Install Python

Download Python:

* https://www.python.org/downloads/

IMPORTANT:
Enable:

```text
Add Python to PATH
```

during installation.

---

# Install Required Libraries

```bash
pip install pyautogui openpyxl pygetwindow pynput psutil
```

Optional:

```bash
pip install pyperclip
```

---

# Running ALICE

```bash
python main.py
```

---

# Dialer Selection

When ALICE starts:

```text
1. BizPhone
2. Zentrex
```

Select the dialer you want to use.

---

# Controls

| Key | Function         |
| --- | ---------------- |
| .   | Call next number |
| \   | End current call |
| Esc | Pause / Resume   |

---

# Call Status Commands

| Input | Saved As               |
| ----- | ---------------------- |
| DNC   | DNC                    |
| NPU   | NPU                    |
| NI    | Not Interested         |
| FU    | Follow Up              |
| NIU   | NIU                    |
| Appt  | Appointment            |
| prev  | Redial previous number |
| dc    | End current call       |

---

# Zentrex Automation

ALICE Z uses:

* keyboard automation
* textbox detection
* auto window switching

Workflow:

1. Focus Zentrex
2. Focus phone textbox
3. Clear previous number
4. Type new number
5. Press Enter to call

---

# BizPhone Automation

ALICE Lite for BizPhone:

* works through LDPlayer
* uses keypad automation
* supports auto dialing through emulator

---

# Packaging Into EXE

Install PyInstaller:

```bash
pip install pyinstaller
```

Build executable:

```bash
pyinstaller --onefile main.py
```

Executable output:

```text
dist/
```

---

# Current Limitations

## Windows Focused

Current version is optimized for Windows.

### Fully Supported

* Windows 10
* Windows 11

### Partial Support

* macOS
* Linux

---

# Future Improvements

* GUI launcher
* SQLite database
* CRM integration
* OCR call detection
* Voice transcription
* AI summarization
* Call analytics dashboard
* Browser automation
* Cloud sync

---

# Disclaimer

This software is intended for workflow automation and productivity purposes only.

Users are responsible for complying with:

* local telemarketing laws
* DNC regulations
* company compliance policies

---

# Author

Created by Josh Soh

ALICE Lite © 2026
