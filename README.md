# Veeva Vault - content factory Export PoC

## 1. Purpose

This proof of concept supports the Veeva litigation export process. It can query Veeva for matching document records and download source content versions through the Vault API. Audit trails are obtained through Vault's bulk Document Export workflow, and PDF annotations are exported separately for each major and minor version.

```text
Veeva query and source-version download
        ↓
Manual bulk export of audit trails and per-version annotations
        ↓
Organize all artifacts by document number
        ↓
Rename annotation files
        ↓
Create tracker and exceptions report
        ↓
Create one delivery ZIP
```

The API script does not modify Vault records or export annotations. The local organizer does not upload to Teams/SharePoint.

---

## 2. Business process supported

The source process requires the following activities:

1. Create a Veeva report for requested documents.
2. Export source documents for all required major and minor versions.
3. Export audit trails.
4. Export annotations for each document version where available.
5. Organize exported files into folders named for the document number.
6. Rename annotation files using the document number and version number.
7. Prepare the content for delivery to the customer’s designated Teams channel.

The API downloader supports finding documents and downloading their source versions. Vault bulk export and annotation export remain user-performed steps, followed by local organization and packaging.

---

## 3. Scope

### Included

- Authenticate to the configured Vault through the configured proxy.
- Read matching document versions with major/minor version metadata and download source files into document-number folders.
- Write a manifest of downloaded source versions.
- Read manual export files from the input folder, including subfolders.
- Detect document numbers such as `DOC123456` and hyphenated Vault document numbers.
- Classify files as **Source**, **Audit_Trail**, or **Annotations**.
- Create one folder per document number.
- Rename annotation files.
- Create a collection tracker CSV.
- Create an exceptions report CSV for files without a document number.
- Create one ZIP package for the processed request.

### Not included

- Automatic Veeva bulk export of audit trails.
- Automatic PDF annotation export.
- Automatic creation of Veeva reports.
- PDF page splitting, OCR, or interpretation of document content.
- Validation that the Veeva export is complete.
- Checksum manifest generation.
- Automatic Microsoft Teams / SharePoint upload.
- Processing files that do not contain an identifiable document number unless manually renamed before execution.

---

## 4. Prerequisites

### Required software

- Windows 10 or Windows 11
- Python 3.10 or later

Check your Python version in PowerShell:

```powershell
python --version
```

The current downloader has been tested with:

```text
Python 3.10.x
```

### Required permissions

- Read access to the manually downloaded Veeva files.
- Write access to your local project folder.
- A Veeva account with API access and permission to retrieve the requested documents.
- Corporate network/proxy access and the corporate proxy root certificate trusted by Windows.

### Python dependencies

The API downloader requires `requests`. The local organizer uses only the Python standard library.

---

## 5. Project files

After extracting this package, keep the following files together:

```text
veeva-simple-poc\
├── zip_creator.py
├── veeva_download.py
├── .env
├── README.md
├── sample_files_note.txt
└── .gitignore
```

| File | Purpose |
|---|---|
| `zip_creator.py` | Main script that organizes files and creates the delivery ZIP. |
| `veeva_download.py` | Authenticates, queries all matching major/minor source versions, and downloads files. |
| `.env` | Local credentials, Vault/proxy settings, search text, and paths. It is Git-ignored. |
| `sharepoint_upload.py` | Separate Microsoft Graph uploader for an approved delivery ZIP. |
| `README.md` | This operating guide. |
| `sample_files_note.txt` | Short filename examples. |
| `.gitignore` | Prevents generated `output` content from being added to Git. |

---

## 6. Prepare the local folders

### 6.1 Project folder

The project folder is:

```text
C:\Users\EFDDQ\Documents\GIT\automation\veeva-simple-poc
```

### 6.2 Input folder

The API downloader writes source files into the configured input folder. Place the audit trail export and annotation PDFs in the same folder before organizing:

```text
C:\Users\EFDDQ\Documents\GIT\automation\veeva-simple-poc\input
```

You may use another input-folder path by changing `VEEVA_INPUT_FOLDER` in `.env`.

### 6.3 Output folder

The script creates the output folder automatically. The recommended local output location is:

```text
C:\Users\EFDDQ\Documents\GIT\automation\veeva-simple-poc\output
```

---

## 7. Manual Veeva preparation

### API source download

Create a local `.env` file. Keep credentials and proxy secrets out of source control:

```text
VEEVA_INPUT_FOLDER=C:\\Users\\<user>\\Documents\\veeva\\input
VEEVA_REQUEST_ID=LIT-001
VEEVA_OUTPUT_FOLDER=output
VEEVA_USERNAME=<Veeva username>
VEEVA_PASSWORD=<Veeva password>
VEEVA_DOCUMENT_NAME_CONTAINS=<document name search text>
BASE_URL=https://<vault-host>
VEEVA_API_VERSION=v26.2
VEEVA_PROXY=http://<username>:<password>@<proxy-host>:<port>
```

Install the API dependency and run the downloader. A search phrase can be passed on the command line to override the `.env` value:

```powershell
python -m pip install requests
python veeva_download.py "document name"
```

The script authenticates, checks product metadata, queries `allversions documents`, downloads the source file for every matching major/minor version, and writes `veeva_download_manifest.csv` in the input folder. It uses Windows' trusted root and CA certificates for TLS verification; do not disable certificate verification.

The Vault query returns no document rows when the search text does not match. Choose the search text supplied by the business team and confirm the query results before proceeding.

### Vault API reference

The downloader uses Vault API `v26.2` and sends requests through `VEEVA_PROXY` when configured.

#### 1. Authenticate

```http
POST https://sbbayer-uat.veevavault.com/api/v26.2/auth
Accept: application/json
Content-Type: application/x-www-form-urlencoded

username=<Veeva username>&password=<Veeva password>
```

The response contains `sessionId`, `vaultId`, and `vaultIds[].url`. The downloader does not display or persist the session ID. It selects the API URL for the returned `vaultId` and uses that URL for subsequent requests. This is important because a session authenticated through a discovery or Stage URL may be bound to the Vault API host returned in the authentication response.

#### 2. Validate metadata

```http
GET {vault_api_url}/v26.2/metadata/vobjects/product__v
Authorization: <sessionId>
Accept: application/json
```

The request must return `responseStatus=SUCCESS` before document retrieval continues.

#### 3. Query all document versions

```http
POST {vault_api_url}/v26.2/query
Authorization: <sessionId>
Accept: application/json
Content-Type: application/x-www-form-urlencoded

q=SELECT id, name__v, filename__v, document_number__v,
major_version_number__v, minor_version_number__v
FROM allversions documents
WHERE name__v CONTAINS ('<search text>')
```

The downloader records these fields:

| Field | Use |
|---|---|
| `id` | Document resource identifier for file retrieval. |
| `name__v` | Document name returned by the search. |
| `filename__v` | Original Vault filename used as the local filename fallback. |
| `document_number__v` | Folder and delivery document number. |
| `major_version_number__v` | Major version path component. |
| `minor_version_number__v` | Minor version path component. |

Vault can return `responseStatus=WARNING` with valid data, for example when a duplicate query is detected. The downloader accepts both `SUCCESS` and `WARNING` when data is present.

#### 4. Download a specific version

```http
GET {vault_api_url}/v26.2/objects/documents/{id}/versions/{major}/{minor}/file
Authorization: <sessionId>
```

Binary content is written to:

```text
input/<document_number>/Source/<document_number>_v<major>.<minor>_<filename>
```

JSON responses from this endpoint are treated as Vault errors and recorded in `veeva_download_manifest.csv`. Common causes include missing content files and insufficient `Download Source` permission for the document lifecycle/status.

#### API security and connectivity

- Use a local `.env` file for credentials and proxy settings; never commit it.
- Keep TLS certificate verification enabled.
- The Windows trusted `ROOT` and `CA` certificate stores are loaded for corporate proxy TLS inspection.
- Use a Vault API access token or another approved production authentication method instead of embedding passwords in scheduled jobs.
- Do not include session IDs, passwords, proxy credentials, or API tokens in logs, screenshots, manifests, or support tickets.

### Manual audit trail and annotation export

Before running `zip_creator.py`, complete these Vault steps:

1. Create or open the requested document report and enter each requested document name individually.
2. Confirm the requested document numbers.
3. Select all requested documents with the bulk action and choose **Document Export**.
4. Choose **Source Documents** for all major and minor versions and **Audit Trails**, then select the requested document naming rule.
5. Finish the export and download the ZIP from the Vault notification.
6. For every document and every major/minor version, open the version and export its PDF annotations individually.
7. Rename each annotation `<Document Number>_v<Version Number>_Annotation.pdf`.
8. Extract the bulk export and place it, plus annotation PDFs, under the input folder.

The PoC expects individual files in the input folder. If Veeva gives you a ZIP file, manually extract that ZIP into:

```text
C:\Users\EFDDQ\Downloads\Veeva_Export
```

The script scans all subfolders inside the input folder.

---

## 8. Filename requirements

### 8.1 Document number

Each file should contain its document number in the filename using this pattern:

```text
DOC followed by digits
```

Examples:

```text
DOC123456_v1.0_source.pdf
DOC123456_v1.0_annotation.pdf
DOC123456_v1.0_audit.csv
DOC234567_v2.1_source.msg
```

The organizer recognizes `DOC123456` and hyphenated Veeva document numbers such as `PP-QLA-CR-0004`.

### 8.2 Version number

For an annotation, include the version where possible:

```text
DOC123456_v2.1_annotation.pdf
```

The script looks for a version in this form:

```text
v1.0
v2.1
v3_0
```

An underscore in a version is converted to a decimal point. For example:

```text
v2_1 → v2.1
```

If no version is found in an annotation filename, the script uses `UNKNOWN`:

```text
DOC123456_vUNKNOWN_Annotation.pdf
```

Review such files manually before delivery.

---

## 9. File classification rules

The script classifies each file using its filename.

| Filename contains | Target folder | Example |
|---|---|---|
| `annotation` | `Annotations` | `DOC123456_v1.0_annotation.pdf` |
| `audit` or `trail` | `Audit_Trail` | `DOC123456_v1.0_audit.csv` |
| Any other filename with a document number | `Source` | `DOC123456_v1.0_source.pdf` |

Classification is case-insensitive. For example, `ANNOTATION`, `Annotation`, and `annotation` are treated the same way.

---

## 10. Annotation naming rule

The source process requires this naming convention:

```text
<Document Number>_v<Version Number>_Annotation.<extension>
```

Example input:

```text
DOC123456_v2.1_annotation.pdf
```

Example output:

```text
DOC123456_v2.1_Annotation.pdf
```

The original file extension is retained. This allows the script to work with PDF, TXT, or another manually exported annotation format.

---

## 11. Run the PoC

Open **PowerShell** and run the local organizer after the source download and manual artifact exports are complete:

```powershell
cd C:\Users\EFDDQ\Documents\GIT\automation\veeva-simple-poc
```

The `.env` file is already configured locally. Set `VEEVA_REQUEST_ID` there (it is currently commented out) before running the organizer:

```text
VEEVA_REQUEST_ID=LIT-001
```

Run the script using `.env`:

```powershell
python zip_creator.py
```

The command-line form is also supported and overrides `.env` values:

```powershell
python zip_creator.py `
  "C:\Users\EFDDQ\Downloads\Veeva_Export" `
  "LIT-001" `
  "output"
```

### Upload the ZIP to Teams-connected SharePoint

The upload is a separate, explicit step. Add these settings to the local `.env` file:

```text
AZURE_TENANT_ID=<Microsoft Entra tenant ID>
AZURE_CLIENT_ID=<app registration client ID>
AZURE_CLIENT_SECRET=<app registration client secret>
SHAREPOINT_SITE_ID=<SharePoint site ID>
SHAREPOINT_DRIVE_ID=<document library drive ID>
SHAREPOINT_FOLDER_PATH=Approved/Veeva Deliveries
```

The app registration must have approved Microsoft Graph application permission `Files.ReadWrite.All` (or a more restricted approved permission) with admin consent. Upload the ZIP generated by the `.env` configuration:

```powershell
python sharepoint_upload.py
```

Or upload a specific ZIP:

```powershell
python sharepoint_upload.py "output\LIT-001\LIT-001_delivery.zip"
```

The uploader uses the SharePoint document library associated with the configured site and drive. It does not upload to a Teams chat directly; files in the channel's SharePoint folder appear in the Teams channel.

### Command parameters

| Parameter | Example | Meaning |
|---|---|---|
| Input folder | `C:\Users\EFDDQ\Documents\GIT\automation\veeva-simple-poc\input` | API source downloads plus manually exported audit and annotation files. |
| Request ID | `LIT-001` | Identifier for this litigation collection or delivery run. |
| Output folder | `output` | Local root folder where the script creates results. |

Use a unique request ID for each collection, for example:

```text
LIT-001
LIT-2026-09-21
DR-SCHOLLS-REQUEST-01
```

---

## 12. Expected output

For request ID `LIT-001`, the result is created here:

```text
C:\Users\EFDDQ\Documents\GIT\automation\veeva-simple-poc\output\LIT-001
```

Example structure:

```text
output\LIT-001\
├── veeva_download_manifest.csv (in the input folder before organization)
├── DOC123456\
│   ├── Source\
│   │   └── DOC123456_v1.0_source.pdf
│   ├── Audit_Trail\
│   │   └── DOC123456_v1.0_audit.csv
│   └── Annotations\
│       └── DOC123456_v1.0_Annotation.pdf
├── DOC234567\
│   └── Source\
│       └── DOC234567_v2.1_source.msg
├── Unmatched\
│   └── file_without_document_number.pdf
├── collection_tracker.csv
├── exceptions_report.csv
└── LIT-001_delivery.zip
```

---

## 13. Output files explained

### 13.1 Document folders

Each identified document number gets one folder:

```text
DOC123456\
```

The script creates the applicable subfolders as it finds files:

```text
Source\
Audit_Trail\
Annotations\
```

### 13.2 `collection_tracker.csv`

This file records successfully classified files.

Columns:

| Column | Meaning |
|---|---|
| `document_number` | Document number detected in the filename. |
| `version` | Version detected in the filename or `UNKNOWN`. |
| `category` | `Source`, `Audit_Trail`, or `Annotations`. |
| `source_file` | Original local file path. |
| `output_file` | New organized local file path. |

Open it in Excel to review the collection result.

`veeva_download_manifest.csv` is created by `veeva_download.py` in the input folder. It records each downloaded document number, version, Vault document ID, and source-file path.

### 13.3 `exceptions_report.csv`

This file records files that could not be matched to a document number.

Example reason:

```text
Document number not found in filename
```

Correct those filenames manually, then rerun the script using a new request ID or remove the prior output folder first.

### 13.4 `Unmatched` folder

Files without a document number are copied to this folder. They are not deleted or ignored.

### 13.5 Delivery ZIP

The final ZIP is created here:

```text
output\LIT-001\LIT-001_delivery.zip
```

It contains document folders, the tracker, and the exceptions report. Review the output before manually uploading the ZIP or its contents to the required customer Teams/SharePoint location.

---

## 14. Example end-to-end run

### Input files

```text
C:\Users\EFDDQ\Downloads\Veeva_Export\DOC123456_v1.0_source.pdf
C:\Users\EFDDQ\Downloads\Veeva_Export\DOC123456_v1.0_annotation.pdf
C:\Users\EFDDQ\Downloads\Veeva_Export\DOC123456_v1.0_audit.csv
C:\Users\EFDDQ\Downloads\Veeva_Export\DOC234567_v2.1_source.pdf
C:\Users\EFDDQ\Downloads\Veeva_Export\notes.txt
```

### Command

```powershell
python zip_creator.py `
  "C:\Users\EFDDQ\Downloads\Veeva_Export" `
  "LIT-DEMO" `
  "output"
```

### Result

```text
output\LIT-DEMO\DOC123456\Source\DOC123456_v1.0_source.pdf
output\LIT-DEMO\DOC123456\Audit_Trail\DOC123456_v1.0_audit.csv
output\LIT-DEMO\DOC123456\Annotations\DOC123456_v1.0_Annotation.pdf
output\LIT-DEMO\DOC234567\Source\DOC234567_v2.1_source.pdf
output\LIT-DEMO\Unmatched\notes.txt
output\LIT-DEMO\collection_tracker.csv
output\LIT-DEMO\exceptions_report.csv
output\LIT-DEMO\LIT-DEMO_delivery.zip
```

---

## 15. Troubleshooting

### Error: `python is not recognized`

Python is not installed or is not available in the Windows PATH.

Actions:

1. Install Python 3.10 or later.
2. During installation, select **Add Python to PATH**.
3. Close and reopen PowerShell.
4. Run:

```powershell
python --version
```

### Error: `Input folder does not exist`

The input-folder path passed to the script is not correct.

Actions:

1. Confirm the folder exists in File Explorer.
2. Copy the path from File Explorer.
3. Put the path in double quotes in PowerShell.

### File is in `Unmatched`

The filename does not contain a recognized `DOC` plus digits or hyphenated Veeva document number.

Action: Rename the source file before running the script, for example:

```text
old_file_name.pdf
```

to:

```text
DOC123456_v1.0_source.pdf
```

### Annotation is named `vUNKNOWN`

The filename did not contain a detectable version.

Action: Rename the original file before running the script, for example:

```text
DOC123456_annotation.pdf
```

to:

```text
DOC123456_v2.1_annotation.pdf
```

### Existing output must be replaced

The simple PoC copies files again if the same request ID is reused. To create a clean delivery, either:

- use a new request ID, or
- delete the existing request folder under `output` before running again.

Example:

```powershell
Remove-Item -Recurse -Force ".\output\LIT-001"
```

Use this delete command only after confirming that the existing output is no longer needed.

---

## 16. Validation before delivery

Before manually uploading the final package to Teams/SharePoint, complete this checklist:

- [ ] Requested document numbers are present.
- [ ] Required source documents and versions are present.
- [ ] Audit trails are present where required.
- [ ] Required annotations are present.
- [ ] Annotation names use the required document/version format.
- [ ] `exceptions_report.csv` is reviewed and resolved or accepted.
- [ ] `collection_tracker.csv` is reviewed.
- [ ] The delivery ZIP opens successfully.
- [ ] The output is uploaded only to the approved customer destination.

---

## 17. Security and handling guidance

- Do not store Veeva passwords, session cookies, or API tokens in this project.
- Treat manually exported litigation content according to applicable Bayer access, retention, and information-handling requirements.
- Keep the input folder and output folder in approved local/network locations.
- Review the delivery package before transferring it to Teams/SharePoint.
- Do not share the output ZIP through unapproved channels.

---

## 18. Future enhancements, not part of this PoC

After this simple PoC is accepted, possible small enhancements include:

1. Add a mapping CSV for files with inconsistent names.
2. Add duplicate detection and checksum manifests.
3. Process a Veeva-export ZIP directly instead of manually extracting it.
4. Add a review screen or lightweight local user interface.
5. Add approved Microsoft Graph upload to the Teams-connected SharePoint library.
6. Add formal audit logging and restart support.

---

## 19. Summary

The PoC automates source-version retrieval and repetitive local organization while leaving the Vault bulk audit export, per-version annotation export, and final customer delivery under operator control:

```text
Query Veeva and download source versions
        ↓
Bulk export audit trails and export annotations per version
        ↓
Place artifacts in the configured input folder and run zip_creator.py
        ↓
Review tracker and exceptions
        ↓
Use the delivery ZIP for approved manual Teams/SharePoint delivery
```
