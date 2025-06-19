# Mail Manager - Outlook Add-in

**Mail Manager** is a custom Microsoft Outlook add-in that enhances your email-sending experience with advanced file handling, rule enforcement, and customizable security modes.

---

## Features

### Automatic Local Backup
- A JSON file with email metadata (To, Subject, Body, Timestamp, etc.) is saved locally each time you send an email.
- All attachments are also saved to your local device for backup and auditing.

---

## Email Rule Enforcement

Before an email is sent, the following rules are validated (depending on selected mail mode):

1. **Suspicious Attachment Name Check**  
   Blocks files with names like `malware.txt`, `virus.js`, etc.

2. **Attachment Size Limit**  
   Ensures no single attachment exceeds **5MB**.

3. **Confidential Content Scan**  
   Scans `.txt` files for sensitive keywords such as `"secret"` or `"confidential"`.

---

## Mail Modes

Users can select from **three different modes** depending on their use-case. Each mode determines which validation rules apply:

| Mode         | Attachment Name | Attachment Size | Attachment Content |
|--------------|------------------|------------------|----------------------|
| **Private**   | Checked         | Checked         | Checked             |
| **Protected** | Checked         | Checked         | Not Checked         |
| **Public**    | Checked         | Not Checked     | Not Checked         |

> Private mode is the most secure; Public mode is the most lenient.

---

## How to Install

### Prerequisites
- Outlook (Web/Desktop)
- Manifest file (`manifest.xml`)
- Python executable for local file storage

---

### First-Time Setup

1. **Download** the `manifest.xml` file.
2. Open **Microsoft Outlook**.
3. Click on the **Apps (grid)** icon → **Get Add-ins**.

   ![Get Add-ins](https://github.com/user-attachments/assets/83a309a4-df68-4b7c-b54b-b9e7309c6d47)
   ![image](https://github.com/user-attachments/assets/271d1845-8484-4d9c-98e9-0da56d26e1c1)


5. Go to **My Add-ins** → Scroll down → Click **+ Add a custom add-in** → Select `manifest.xml`.

   ![Upload Manifest](https://github.com/user-attachments/assets/9c3318e0-eead-4b23-b724-2d676d3941cf)

6. You're all set! The add-in is now active.

---

### Already Installed?

If the add-in is already added to your account, just run the provided **Python executable** to enable local saving functionality.

---

## Tech Stack

- **Frontend:** HTML, CSS, JavaScript (Hosted on Render)
- **Backend:** Python (Flask, packaged as `.exe` for local execution)
- **APIs:** Office.js (Outlook Add-in framework)

---

