# NetDocuments Link Converter – Deployment Guide

## Table of Contents
1. [VBA Macro Installation](#1-vba-macro-installation)
2. [Single-Machine Registry Test](#2-single-machine-registry-test)
3. [GPO Domain Deployment](#3-gpo-domain-deployment)
4. [Verification & Troubleshooting](#4-verification--troubleshooting)

---

## 1. VBA Macro Installation

### Option A – Import the .bas file
1. Open **Outlook**
2. Press **Alt + F11** to open the VBA Editor
3. In the **Project Explorer** (left pane), right-click **Project1** → **Import File…**
4. Browse to `ConvertNetDocsLinksToAttachments.bas` and click **Open**
5. Close the VBA Editor

### Option B – Copy/paste
1. Open **Outlook** → **Alt + F11**
2. **Insert** → **Module**
3. Paste the entire contents of `ConvertNetDocsLinksToAttachments.bas`
4. Close the VBA Editor

### Running the Macro
1. Open (double-click) an email that contains NetDocuments links
2. Press **Alt + F8**
3. Select `ConvertNetDocsLinksToAttachments` → click **Run**
4. The macro will:
   - Create `.url` attachment files
   - Attach them to the email
   - Remove all NetDocuments links from the body
   - Show a confirmation message

### Optional: Add a Ribbon Button
1. **File** → **Options** → **Customize Ribbon**
2. Under **Choose commands from**, select **Macros**
3. Select `ConvertNetDocsLinksToAttachments`
4. Click **Add >>** to place it on a custom ribbon group

---

## 2. Single-Machine Registry Test

### Why this is needed
Many organisations block `.url` files via Outlook attachment security (the
`Level1Remove` registry value). The macro creates `.url` files in
`%TEMP%\NetDocsLinks\`; Outlook must be told to allow opening these.

### 2A. Allow .url File Extension in Outlook

> **Important:** These changes affect which attachment types Outlook allows
> the user to open. The `.url` extension is being added to the *Level1Remove*
> key, which tells Outlook to **not** block it.

#### Step-by-step (Registry Editor)

1. Press **Win + R** → type `regedit` → press **Enter**
2. Navigate to the key for your Outlook version:

   | Outlook version | Registry path |
   |---|---|
   | Outlook 2016 / 2019 / 365 | `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Security` |
   | Outlook 2013 | `HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Security` |

3. If the `Security` subkey does not exist, right-click `Outlook` → **New** → **Key** → name it `Security`
4. In the right pane, look for a String value named `Level1Remove`
   - If it exists, double-click it and **append** `.url;` to the existing value
     (e.g. if it currently reads `.mdb;` change it to `.mdb;.url;`)
   - If it does not exist:
     - Right-click → **New** → **String Value** → name it `Level1Remove`
     - Double-click → set value to `.url`
5. Close Registry Editor
6. **Restart Outlook** for the change to take effect

#### One-liner (elevated Command Prompt)

```cmd
REG ADD "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v Level1Remove /t REG_SZ /d ".url" /f
```

> Adjust `16.0` to `15.0` for Outlook 2013.

### 2B. Allow the Macro to Run (Macro Security)

1. In Outlook: **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. **Macro Settings** → select **Notifications for all macros** (or lower)
3. Click **OK** → restart Outlook
4. When prompted "A program is trying to access…", click **Allow**

### 2C. Test Procedure

1. Compose a new email or open a draft
2. Paste one or more NetDocuments links into the body, for example:
   ```
   https://vault.netvoyage.com/neWeb2/docCent.aspx?id=123456789
   ```
   Or insert a hyperlink with display text like `Contract_v5.docx` pointing to
   a NetDocuments URL.
3. Run the macro (Alt + F8 → `ConvertNetDocsLinksToAttachments`)
4. Verify:
   - [ ] `.url` files appear in the attachment bar
   - [ ] Filenames match the hyperlink text / document ID
   - [ ] Links have been removed from the body
   - [ ] "Documents attached via NetDocuments" line is present
   - [ ] Double-clicking a `.url` attachment opens the document via
         NetDocuments integration (ndOffice / browser, depending on config)
5. Run the macro again on the same email → verify no duplicate attachments

### 2D. Cleanup After Testing

```cmd
REM Remove the registry override
REG DELETE "HKCU\Software\Microsoft\Office\16.0\Outlook\Security" /v Level1Remove /f

REM Delete temp files
RMDIR /S /Q "%TEMP%\NetDocsLinks"
```

Restart Outlook after removing the registry key.

---

## 3. GPO Domain Deployment

Once the single-machine test succeeds, deploy via Group Policy.

### 3A. Create (or Edit) the GPO

1. Open **Group Policy Management Console** (`gpmc.msc`)
2. Right-click the target OU → **Create a GPO in this domain, and Link it here…**
   - Name: `Outlook – Allow .url Attachments for NetDocuments`
3. Right-click the new GPO → **Edit**

### 3B. Configure Outlook Attachment Security via ADMX

> **Prerequisite:** Office ADMX templates must be loaded into your Central Store.
> Download from Microsoft:
> [Office ADMX/ADML](https://www.microsoft.com/en-us/download/details.aspx?id=49030)

#### Path in Group Policy Editor

```
User Configuration
  → Policies
    → Administrative Templates
      → Microsoft Outlook 2016
        → Security
          → Security Form Settings
            → Attachment Security
```

#### Setting to modify

| Setting | Value |
|---|---|
| **Remove file extensions blocked as Level 1** | **Enabled** |
| Extension list | `.url` |

1. Double-click **"Remove file extensions blocked as Level 1"**
2. Select **Enabled**
3. In the text box labelled **"Remove these extensions"**, enter: `.url`
4. Click **OK**

> This is the GPO equivalent of the `Level1Remove` registry value tested in
> Section 2A. It writes to `HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security\Level1Remove`.

### 3C. Alternative – Registry Preference (if ADMX not available)

If Office ADMX templates are not loaded:

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

### 3D. Deploy the VBA Macro via GPO (Optional)

For organisation-wide macro deployment, you have several options:

#### Option 1 – COM Add-in (Recommended for large orgs)
Package the VBA code as a COM add-in (.dll) and deploy via Software Installation
GPO or SCCM/Intune. This avoids macro security prompts entirely.

#### Option 2 – OTM File Deployment
1. Export the macro from one machine:
   - The VBA project is stored in `%APPDATA%\Microsoft\Outlook\VbaProject.OTM`
2. Deploy the OTM file via logon script or GPO Preferences file copy:
   ```
   User Configuration → Preferences → Windows Settings → Files
   Source: \\server\share\VbaProject.OTM
   Destination: %APPDATA%\Microsoft\Outlook\VbaProject.OTM
   Action: Replace
   ```
   > **Warning:** This overwrites any existing user macros. Best for fresh deployments.

#### Option 3 – Manual Distribution
Distribute the `.bas` file and instructions (Section 1) to users or their IT support.

### 3E. Set Macro Security via GPO

To allow the macro to run without per-user prompts:

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

Then sign the VBA project with a code-signing certificate trusted via GPO.

### 3F. Verify GPO Deployment

1. On a target machine, run:
   ```cmd
   gpupdate /force
   ```
2. Verify the registry key was applied:
   ```cmd
   REG QUERY "HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security" /v Level1Remove
   ```
   Expected output: `.url`
3. Open Outlook → open an email with NetDocuments links → run the macro
4. Confirm `.url` attachments can be opened

---

## 4. Verification & Troubleshooting

### Expected Behaviour

| Action | Result |
|---|---|
| Run macro on email with 5 ND links | 5 `.url` attachments created |
| Run macro again on same email | No new attachments (duplicates skipped) |
| Double-click `.url` attachment | Opens via NetDocuments integration |
| Email body after macro | ND links removed, summary line present |
| Non-ND links in body | Untouched |

### Common Issues

| Symptom | Cause | Fix |
|---|---|---|
| "Outlook blocked access to the following potentially unsafe attachments" | `Level1Remove` not applied | Verify registry key; restart Outlook |
| `.url` opens in browser instead of native app | ndOffice not installed or not handling the protocol | Check ndOffice installation; this is outside macro scope |
| Macro security prompt every time | VBA project not signed | Sign with code-signing cert or lower macro security |
| Macro not visible in Alt+F8 | Module not imported | Re-import `.bas` file per Section 1 |
| `Run-time error 91: Object variable not set` | No email open in Inspector | Open the email in its own window first |
| Duplicate attachments appearing | Filename changed between runs | Ensure filenames are deterministic; check for HTML entity changes |

### Security Notes

- The `.url` extension unblock applies **only** to Outlook attachment handling
- It does NOT affect Windows SmartScreen, antivirus, or web proxy policies
- `.url` files created by the macro point exclusively to `netdocuments.com` domains
- No executable code is embedded in `.url` files – they contain only a URL string
- Consider restricting the unblock to pilot groups before full domain rollout
