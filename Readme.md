# ALICE Lite - Automated Dialer

ALICE Lite is a Python-based automated dialing assistant designed for BizPhone running on LDPlayer.
It helps automate outbound calling workflows by reading phone numbers from Excel, dialing automatically, updating statuses, logging dates, and recording remarks.

---

# Features

* Automated dialing through BizPhone (LDPlayer)
* Reads phone numbers directly from Excel
* Automatically skips completed leads
* Redial previous number using `prev`
* End current call instantly using `\`
* Quick status shortcuts:

  * `NI` → Not Interested
  * `FU` → Follow Up
* Automatically logs:

  * Call status
  * Date
  * Remarks
* Auto-switches between LDPlayer and Python terminal
* Supports follow-up remarks entry
* Minimalist console startup logo

---

# Excel Format

Your Excel file should follow this structure:

| Column | Description  |
| ------ | ------------ |
| A      | Phone Number |
| B      | Call Status  |
| C      | Date         |
| D      | Remarks      |

Example:

| Phone Number | Status    | Date       | Remarks             |
| ------------ | --------- | ---------- | ------------------- |
| 91234567     | Follow Up | 24/05/2026 | Call back next week |
| 98765432     | DNC       | 24/05/2026 |                     |

---

# Requirements

Install the required Python libraries:

```bash
pip install pyautogui openpyxl pygetwindow pynput psutil
```

---

# How To Run

## 1. Open LDPlayer

Launch LDPlayer and open BizPhone.

## 2. Prepare Excel File

Ensure your Excel file exists at:

```python
FILE_PATH = r"C:\Users\Lenovo\OneDrive\Prudential\Leads\LeadsAuto.xlsx"
```

Update the path if necessary.

---

## 3. Run The Script

```bash
python PyCall.py
```

---

# Controls

| Key   | Function         |
| ----- | ---------------- |
| `.`   | Call next number |
| `\\`  | End current call |
| `Esc` | Pause / Resume   |

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

# Workflow

1. Press `.` to call the next number
2. ALICE automatically:

   * Ends previous call
   * Clears numpad
   * Dials next lead
3. Enter call status
4. If status is:

   * `FU`
   * `Appt`

   ALICE will request remarks
5. Excel updates automatically

---

# Example

```text
Calling: 91234567
Switched to BizPhone.
Dialed: 91234567

Enter call status:
FU

Enter remarks:
Call back next Monday

Row updated successfully.
```

---

# Packaging Into EXE

Install PyInstaller:

```bash
pip install pyinstaller
```

Build executable:

```bash
pyinstaller --onefile PyCall.py
```

Your executable will appear in:

```text
dist/
```

---

# Future Improvements

* Voice detection
* Auto voicemail detection
* CRM integration
* GUI version
* Statistics dashboard
* Auto scheduling

---

# Disclaimer

This project is intended for personal productivity and workflow automation purposes only.
Please ensure compliance with your local calling regulations and company policies.

---

# Author

Created by Josh Soh

ALICE Lite © 2025
