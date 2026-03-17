# NetDocuments Link Converter – Deployment Guide

## How It Works

The macro creates `.html` redirect files instead of `.url` shortcut files.

**Why .html?**
- `.html` files are **not blocked** by Outlook's attachment security policy
- No registry changes, no GPO security overrides, no Exchange transport rules
- Zero security policy modifications needed

**Flow:** User double-clicks `.html` attachment → browser opens → instant
redirect to NetDocuments URL → ndOffice intercepts → document opens in the
native app (Word/Excel/etc.) using the **existing instance** if already running.

---

## Table of Contents
1. [VBA Macro Installation](#1-vba-macro-installation)
2. [Testing](#2-testing)
3. [Organisation-Wide Deployment](#3-organisation-wide-deployment)
4. [Verification & Troubleshooting](#4-verification--troubleshooting)

---

## 1. VBA Macro Installation

### Option A – Import the .bas file
1. Open **Outlook** → press **Alt + F11** to open the VBA Editor
2. In the **Project Explorer**, right-click **Project1** → **Import File…**
3. Select `ConvertNetDocsLinksToAttachments.bas` → click **Open**
4. Close the VBA Editor

### Option B – Copy/paste
1. **Alt + F11** → **Insert** → **Module**
2. Paste the entire contents of the `.bas` file
3. Close the VBA Editor

### Running the Macro
1. Open an email in its own window (double-click it)
2. **Alt + F8** → select `ConvertNetDocsLinksToAttachments` → **Run**

### Optional: Ribbon Button
**File** → **Options** → **Customize Ribbon** → under **Choose commands from**
select **Macros** → add `ConvertNetDocsLinksToAttachments` to a custom group.

---

## 2. Testing

### 2A. Macro Security (one-time setup)

1. **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. **Macro Settings** → select **Notifications for all macros**
3. Click **OK** → restart Outlook

### 2B. Test Procedure

1. **Create a test email** (new or open a draft)
2. **Paste NetDocuments links** into the body:
   ```
   Check out this document: https://vault.netvoyage.com/neWeb2/docCent.aspx?id=123456789

   Also see the contract here:
   https://eu.netdocuments.com/nddocview/nav/12345-6789/v1
   ```
   And/or insert a hyperlink with display text like `Contract_v5.docx` pointing
   to a NetDocuments URL.

3. **Run the macro:** Alt + F8 → `ConvertNetDocsLinksToAttachments` → Run

4. **Verify:**
   - [ ] `.html` files appear in the attachment bar
   - [ ] Filenames match hyperlink text or contain the document ID
   - [ ] All NetDocuments links removed from body
   - [ ] "Documents attached via NetDocuments" summary line present
   - [ ] All other email content and formatting preserved
   - [ ] Non-NetDocuments links untouched

5. **Re-run the macro** on the same email:
   - [ ] No duplicate attachments created

6. **Test attachment opening:**
   - [ ] Double-click a `.html` attachment
   - [ ] Browser opens momentarily, then ndOffice takes over
   - [ ] Document opens in the native app (Word/Excel)
   - [ ] If Word was already open, the document opens in that existing instance
         (no new Word window spawned)

### 2C. Cleanup After Testing

```cmd
RMDIR /S /Q "%TEMP%\NetDocsLinks"
```

---

## 3. Organisation-Wide Deployment

Since `.html` attachments are not blocked by Outlook, deployment is
straightforward — **no security policy changes are required**. You only need
to distribute the macro and set macro security.

### 3A. Macro Deployment Options

#### Option 1 – COM Add-in (Recommended for large orgs)
Package the VBA code as a signed COM add-in (`.dll`) and deploy via:
- GPO Software Installation
- SCCM / Intune

This avoids macro security prompts entirely.

#### Option 2 – OTM File via GPO
1. Export from one configured machine:
   `%APPDATA%\Microsoft\Outlook\VbaProject.OTM`
2. Deploy via GPO Preferences:
   ```
   User Configuration → Preferences → Windows Settings → Files
   Source: \\server\share\VbaProject.OTM
   Destination: %APPDATA%\Microsoft\Outlook\VbaProject.OTM
   Action: Replace
   ```
   > **Caution:** Overwrites existing user macros. Best for fresh deployments.

#### Option 3 – Manual Distribution
Distribute the `.bas` file with instructions from Section 1 to users or
their IT support.

### 3B. Macro Security via GPO

To allow the macro to run without per-user prompts:

1. Open **Group Policy Management** (`gpmc.msc`)
2. Create or edit a GPO linked to the target OU
3. Navigate to:
   ```
   User Configuration
     → Policies
       → Administrative Templates
         → Microsoft Outlook 2016
           → Security
             → Trust Center
               → Security Setting for Macros
   ```
4. Set to: **Notifications for digitally signed macros, all other macros disabled**
5. Sign the VBA project with a code-signing certificate trusted across the
   domain via GPO

> Prerequisite: Office ADMX templates loaded into your Central Store.
> Download: https://www.microsoft.com/en-us/download/details.aspx?id=49030

### 3C. Verify Deployment

On a target machine:

```cmd
gpupdate /force
```

Then: open Outlook → open an email with NetDocuments links → run the macro →
confirm `.html` attachments appear and open correctly.

---

## 4. Verification & Troubleshooting

### Expected Behaviour

| Action | Result |
|---|---|
| Run macro on email with 5 ND links | 5 `.html` attachments created |
| Run macro again on same email | No new attachments (duplicates skipped) |
| Double-click `.html` attachment | Browser opens → redirects → ndOffice opens in native app |
| Word already open when clicking | Document opens in existing Word instance |
| Email body after macro | ND links removed, summary line present |
| Non-ND links in body | Untouched |

### Common Issues

| Symptom | Cause | Fix |
|---|---|---|
| `.html` opens in browser but stays there | ndOffice not installed or not handling the ND URL | Check ndOffice installation — outside macro scope |
| New Word instance opens instead of reusing | ndOffice configuration issue | Check ndOffice settings for "reuse existing instance" |
| Macro not visible in Alt+F8 | Module not imported | Re-import `.bas` file (Section 1) |
| `Run-time error 91` | No email open in Inspector | Open the email in its own window first |
| Macro security prompt every time | VBA project not signed | Sign with code-signing certificate (Section 3B) |

### Security Notes

- `.html` files are not on Outlook's blocked-attachment list — no policy
  changes are required
- Each `.html` file contains only a standard redirect and a fallback link —
  no scripts, no executable code, no ActiveX
- The redirect points exclusively to `netdocuments.com` URLs extracted from
  the original email
- ndOffice controls native app behaviour (instance reuse, protocol handling) —
  the macro does not attempt to manage Office application instances
