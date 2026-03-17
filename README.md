# NetDocuments Link-to-Attachment Converter for Outlook

Converts NetDocuments links in emails into `.url` attachment files that open via existing NetDocuments integration (ndOffice).

## Files

| File | Purpose |
|---|---|
| `ConvertNetDocsLinksToAttachments.bas` | Outlook VBA macro – import directly into the VBA Editor |
| `DEPLOYMENT_GUIDE.md` | Full instructions: installation, registry test, GPO deployment, troubleshooting |

## Quick Start

1. Open Outlook → **Alt+F11** → **Import File…** → select the `.bas` file
2. Open an email with NetDocuments links
3. **Alt+F8** → `ConvertNetDocsLinksToAttachments` → **Run**

See [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md) for registry/GPO setup to allow `.url` attachments.
