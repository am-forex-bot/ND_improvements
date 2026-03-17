# NetDocuments Link Converter – Deployment Guide

## Security Architecture

Outlook's `Level1Remove` is **extension-based only** — it cannot be scoped to a
specific folder. To achieve path-scoped security, this solution uses three
complementary layers:

| Layer | Purpose | Scoped? |
|---|---|---|
| **1. Outlook** (`Level1Remove`) | Permits `.url` attachments to be opened by users | Extension-only (required) |
| **2. Exchange Transport Rule** | Blocks ALL inbound `.url` attachments from external senders | Blocks the external attack vector entirely |
| **3. Software Restriction Policy (SRP)** | Restricts `.url` execution to `%TEMP%\NetDocsLinks\` and Outlook's secure cache only | **Path-scoped at OS level** |

**Net effect:** External `.url` attacks are stopped at the mail server. Internally
created `.url` files only execute from approved paths. Users cannot be tricked
into running a malicious `.url` received via email from outside the organisation.

---

## Table of Contents
1. [VBA Macro Installation](#1-vba-macro-installation)
2. [Layer 1 – Outlook: Allow .url Extension](#2-layer-1--outlook-allow-url-extension)
3. [Layer 2 – Exchange: Block Inbound .url Attachments](#3-layer-2--exchange-block-inbound-url-attachments)
4. [Layer 3 – SRP: Path-Restrict .url Execution](#4-layer-3--srp-path-restrict-url-execution)
5. [Single-Machine Test Procedure](#5-single-machine-test-procedure)
6. [GPO Domain Deployment](#6-gpo-domain-deployment)
7. [Verification & Troubleshooting](#7-verification--troubleshooting)

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

## 2. Layer 1 – Outlook: Allow .url Extension

> **What this does:** Removes `.url` from Outlook's Level 1 blocked-attachment
> list so users can open `.url` attachments. This is extension-wide (not
> path-scoped), which is why Layers 2 and 3 are essential.

### 2A. Single-Machine Registry Test

1. Press **Win + R** → `regedit` → **Enter**
2. Navigate to your Outlook version's key:

   | Outlook version | Path |
   |---|---|
   | 2016 / 2019 / 365 | `HKCU\Software\Microsoft\Office\16.0\Outlook\Security` |
   | 2013 | `HKCU\Software\Microsoft\Office\15.0\Outlook\Security` |

3. If the `Security` subkey does not exist: right-click `Outlook` → **New** →
   **Key** → name it `Security`
4. In the right pane, look for `Level1Remove`:
   - **Exists:** double-click and append `.url;` to the value
   - **Does not exist:** right-click → **New** → **String Value** → name it
     `Level1Remove` → set value to `.url`
5. **Restart Outlook**

#### Command-line equivalent (elevated prompt)

```cmd
REG ADD "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v Level1Remove /t REG_SZ /d ".url" /f
```

### 2B. GPO Deployment (via ADMX)

> Prerequisite: Office ADMX templates loaded into your Central Store.
> Download: https://www.microsoft.com/en-us/download/details.aspx?id=49030

1. Open **Group Policy Management** (`gpmc.msc`)
2. Create or edit a GPO linked to the target OU
3. Navigate to:
   ```
   User Configuration
     → Policies
       → Administrative Templates
         → Microsoft Outlook 2016
           → Security
             → Security Form Settings
               → Attachment Security
   ```
4. Open **"Remove file extensions blocked as Level 1"**
5. Set to **Enabled**
6. In the extensions text box, enter: `.url`
7. Click **OK**

### 2C. GPO Deployment (Registry Preference fallback)

If ADMX templates are not available:

1. In the GPO editor, navigate to:
   ```
   User Configuration
     → Preferences
       → Windows Settings
         → Registry
   ```
2. Right-click → **New** → **Registry Item**
3. Configure:

   | Field | Value |
   |---|---|
   | Action | **Update** |
   | Hive | `HKEY_CURRENT_USER` |
   | Key Path | `Software\Microsoft\Office\16.0\Outlook\Security` |
   | Value name | `Level1Remove` |
   | Value type | `REG_SZ` |
   | Value data | `.url` |

4. Click **OK**

---

## 3. Layer 2 – Exchange: Block Inbound .url Attachments

> **What this does:** Prevents externally-sent `.url` files from reaching user
> mailboxes. This is the primary defence against the attack vector that
> Level1Remove opens. Internal mail (macro-generated attachments) is unaffected.

### 3A. Exchange Online (Microsoft 365)

1. Open **Exchange Admin Center** → https://admin.exchange.microsoft.com
2. Navigate to **Mail flow** → **Rules**
3. Click **+ Add a rule** → **Create a new rule**
4. Configure:

   | Field | Value |
   |---|---|
   | Name | `Block inbound .url attachments` |
   | Apply this rule if… | **The sender is located…** → **Outside the organization** |
   | And… | **Any attachment's file extension includes…** → `url` |
   | Do the following… | **Reject the message with the explanation…** → `.url attachments are not permitted` |

   Alternative "Do the following" (less disruptive):
   - **Remove the attachment and notify** → `A .url attachment was removed for security`

5. Set **Priority** appropriately (lower number = higher priority)
6. Click **Save**

### 3B. Exchange On-Premises (2016/2019)

Open **Exchange Management Shell** and run:

```powershell
New-TransportRule -Name "Block inbound .url attachments" `
    -FromScope "NotInOrganization" `
    -AttachmentExtensionMatchesWords "url" `
    -RejectMessageReasonText ".url attachments from external senders are not permitted" `
    -Enabled $true
```

To strip the attachment instead of rejecting:

```powershell
New-TransportRule -Name "Strip inbound .url attachments" `
    -FromScope "NotInOrganization" `
    -AttachmentExtensionMatchesWords "url" `
    -StripClassificationFromHeader $false `
    -RemoveHeader "X-Placeholder" `
    -Enabled $true
```

> **Note:** Exchange transport rules do not affect mail sent between internal
> users, so `.url` attachments created by the macro and sent internally will
> pass through normally.

### 3C. Verify the Transport Rule

Send a test email **from a personal/external account** to an internal mailbox
with a `.url` file attached. It should be rejected or stripped.

---

## 4. Layer 3 – SRP: Path-Restrict .url Execution

> **What this does:** Even though Outlook now allows opening `.url` files,
> Windows will only execute them from approved folders. A `.url` file saved
> to the Desktop or downloaded from the web will be blocked at the OS level.

### Important: Execution Paths

When a user double-clicks a `.url` attachment in Outlook, the file is extracted
to Outlook's secure temp cache before execution:

```
%LOCALAPPDATA%\Microsoft\Windows\INetCache\Content.Outlook\<8-char-hash>\
```

You must allow BOTH paths:
1. `%TEMP%\NetDocsLinks\` — where the macro creates the files
2. Outlook's `Content.Outlook` cache — where attachments execute from

### 4A. Single-Machine Test (Local Security Policy)

1. Press **Win + R** → `secpol.msc` → **Enter**
2. Navigate to **Software Restriction Policies**
   - If "No Software Restriction Policies Defined" appears:
     right-click → **New Software Restriction Policies**
3. Open **Designated File Types** (in the right pane)
4. Check if `URL` is listed. If not:
   - In the **File extension** box, type `URL`
   - Click **Add**
5. Go to **Additional Rules** (in the left pane)
6. Right-click → **New Path Rule…** for each path below:

   **Rule 1 – Macro creation folder:**

   | Field | Value |
   |---|---|
   | Path | `%TEMP%\NetDocsLinks\*.url` |
   | Security level | **Unrestricted** |

   **Rule 2 – Outlook attachment cache:**

   | Field | Value |
   |---|---|
   | Path | `%LOCALAPPDATA%\Microsoft\Windows\INetCache\Content.Outlook\*.url` |
   | Security level | **Unrestricted** |

7. Set the **Default Security Level** to confirm it is **Unrestricted** for
   general file types (SRP default). The point here is that `.url` is now a
   designated file type, and only the two path rules above permit execution.

> **Important:** If your organisation already uses SRP with a default of
> **Disallowed**, adding `.url` to Designated File Types and creating these
> two path rules is all you need. If SRP is not currently in use and default
> is Unrestricted, the path rules above serve as documentation and
> preparation — the real blocking comes from having a Disallowed default or
> from using WDAC (see 4C below).

### 4B. GPO Deployment (SRP)

1. Open **Group Policy Management** (`gpmc.msc`)
2. Create or edit a GPO linked to the target OU
3. Navigate to:
   ```
   Computer Configuration
     → Policies
       → Windows Settings
         → Security Settings
           → Software Restriction Policies
   ```
4. If no policies exist: right-click → **New Software Restriction Policies**
5. Open **Designated File Types** → add `URL` if not present
6. Under **Additional Rules**, create the two path rules from 4A above
7. Link the GPO and run `gpupdate /force` on target machines

### 4C. Alternative – Windows Defender Application Control (WDAC)

For organisations using WDAC instead of SRP:

1. Create a WDAC policy that includes a file rule for `.url` with a path
   condition restricting execution to:
   - `%TEMP%\NetDocsLinks\`
   - `%LOCALAPPDATA%\Microsoft\Windows\INetCache\Content.Outlook\`
2. Deploy via GPO, Intune, or SCCM

> WDAC is the modern replacement for SRP and is recommended for Windows 10/11
> Enterprise and Server 2016+. SRP is simpler to configure but is considered
> legacy.

---

## 5. Single-Machine Test Procedure

### Pre-requisites
- [ ] Layer 1 applied: `Level1Remove` registry key set (Section 2A)
- [ ] Outlook restarted after registry change
- [ ] Macro imported (Section 1)
- [ ] Macro security set to "Notifications for all macros" or lower

### Test Steps

1. **Create a test email** (new email or open a draft)
2. **Paste NetDocuments links** into the body:
   ```
   Check out this document: https://vault.netvoyage.com/neWeb2/docCent.aspx?id=123456789

   Also see the contract here:
   https://eu.netdocuments.com/nddocview/nav/12345-6789/v1
   ```
   And/or insert a hyperlink with display text like `Contract_v5.docx` pointing
   to a NetDocuments URL.

3. **Run the macro:** Alt + F8 → `ConvertNetDocsLinksToAttachments` → Run

4. **Verify results:**
   - [ ] `.url` files appear in the attachment bar
   - [ ] Filenames match hyperlink text or contain the document ID
   - [ ] All NetDocuments links removed from body
   - [ ] "Documents attached via NetDocuments" summary line present
   - [ ] All other email content and formatting preserved
   - [ ] Non-NetDocuments links untouched

5. **Re-run the macro** on the same email:
   - [ ] No duplicate attachments created
   - [ ] Message box still confirms completion

6. **Test attachment opening:**
   - [ ] Double-click a `.url` attachment
   - [ ] It opens via NetDocuments integration (ndOffice → native app)

7. **Test Layer 2** (if Exchange rule configured):
   - [ ] Send a `.url` file from an external email address
   - [ ] Confirm it is rejected or stripped

### Cleanup After Testing

```cmd
REM Remove Level1Remove override
REG DELETE "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v Level1Remove /f

REM Remove SRP path rules (via secpol.msc → delete the two rules)

REM Delete temp files
RMDIR /S /Q "%TEMP%\NetDocsLinks"
```

Restart Outlook after cleanup.

---

## 6. GPO Domain Deployment

### Deployment Checklist

| Step | Action | Section |
|---|---|---|
| 1 | Load Office ADMX templates into Central Store (if not done) | 2B |
| 2 | Create GPO: "Outlook – NetDocs URL Converter" | — |
| 3 | Configure Level1Remove for `.url` | 2B or 2C |
| 4 | Create Exchange transport rule to block external `.url` | 3A or 3B |
| 5 | Configure SRP path rules for `.url` | 4B |
| 6 | Deploy the VBA macro (choose one method below) | 6A–6C |
| 7 | Configure macro security + sign the VBA project | 6D |
| 8 | Test on pilot group | 5 |
| 9 | Roll out to production | — |

### 6A. Macro Deployment – COM Add-in (Recommended)

Package the VBA code as a signed COM add-in (`.dll`) and deploy via:
- GPO Software Installation, or
- SCCM / Intune

This avoids macro security prompts entirely and gives the cleanest user
experience.

### 6B. Macro Deployment – OTM File

1. Export from one machine: `%APPDATA%\Microsoft\Outlook\VbaProject.OTM`
2. Deploy via GPO Preferences:
   ```
   User Configuration → Preferences → Windows Settings → Files
   Source: \\server\share\VbaProject.OTM
   Destination: %APPDATA%\Microsoft\Outlook\VbaProject.OTM
   Action: Replace
   ```
   > **Caution:** Overwrites existing user macros.

### 6C. Macro Deployment – Manual

Distribute the `.bas` file with instructions from Section 1.

### 6D. Macro Security via GPO

```
User Configuration
  → Policies
    → Administrative Templates
      → Microsoft Outlook 2016
        → Security
          → Trust Center
            → Security Setting for Macros
```

Set to: **Notifications for digitally signed macros, all other macros disabled**

Then sign the VBA project with a code-signing certificate trusted across the
domain via GPO.

### Verify GPO Deployment

On a target machine:

```cmd
REM Force policy update
gpupdate /force

REM Verify Level1Remove
REG QUERY "HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security" /v Level1Remove

REM Verify SRP rules
gpresult /H gpresult.html
REM Open gpresult.html and check Software Restriction Policies section
```

---

## 7. Verification & Troubleshooting

### Expected Behaviour

| Action | Result |
|---|---|
| Run macro on email with 5 ND links | 5 `.url` attachments created |
| Run macro again on same email | No new attachments (duplicates skipped) |
| Double-click `.url` attachment | Opens via NetDocuments / ndOffice |
| Email body after macro | ND links removed, summary line present |
| Non-ND links in body | Untouched |
| External sender sends `.url` attachment | Blocked by Exchange transport rule |

### Common Issues

| Symptom | Cause | Fix |
|---|---|---|
| "Outlook blocked access to unsafe attachments" | `Level1Remove` not applied | Verify registry key; restart Outlook |
| `.url` opens in browser instead of native app | ndOffice not installed or not registered as handler | Check ndOffice — outside macro scope |
| Macro not visible in Alt+F8 | Module not imported | Re-import `.bas` file (Section 1) |
| `Run-time error 91` | No email open in Inspector | Open email in its own window first |
| Duplicate attachments appearing | Filename changed between runs | Check for HTML entity changes in body |
| SRP blocks `.url` from wrong path | Path rule missing or wrong | Verify both path rules (Section 4A) |
| Macro security prompt every time | VBA project not signed | Sign with code-signing certificate |

### Security Notes

- **Layer 1** (Level1Remove) is extension-wide by design — Layers 2 and 3
  compensate for this
- **Layer 2** (Exchange rule) is the most critical control — it blocks the
  primary external attack vector
- **Layer 3** (SRP) provides defence-in-depth at the OS level
- `.url` files contain only `[InternetShortcut]` and `URL=` — no executable code
- The macro only creates `.url` files pointing to `netdocuments.com` domains
- Consider deploying to a pilot group before full rollout
- All three layers should be deployed together for the intended security posture
