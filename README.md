# office-tools

A collection of PowerShell cmdlets for handling Microsoft Office documents.

*This project was started based on some old scripts I still had around. The main focus was simply having fun while converting them to more usable PowerShell modules and extending their original capabilities. Besides that, I used it as a purpose for engaging and gaining first experience with AI-assisted coding.*

## Included Scripts

### 1. Convert-LegacyDocumentToXMLDocument
**Purpose:** Batch converts legacy Microsoft Office documents to modern XML-based formats.

**Supported Conversions:**
- `.doc` → `.docx`
- `.xls` → `.xlsx`

**Requirements:**
- Word and Excel (any recent version)
- PowerShell 7.4 LTS or later
- Windows OS

**Use Case:**
- Convert legacy Microsoft Office documents to modern formats
- Ensure compatibility with current Microsoft Office versions

**Features:**
- Recursive directory scanning
- Automatic CSV logging
- Duplicate file protection
- Preserves original files

### 2. Convert-OfficeDocumentToOpenDocument
**Purpose:** Batch converts Microsoft Office documents to Open Document Format using LibreOffice.

**Supported Conversions:**
- `.doc`, `.docx` → `.odt`
- `.xls`, `.xlsx` → `.ods`
- `.ppt`, `.pptx` → `.odp`

**Requirements:**
- LibreOffice (any recent version)
- PowerShell 7.4 LTS or later
- Windows OS

**Use Case:**
- Convert to vendor-independent file formats
- Migrate documents between different office suites

**Features:**
- Recursive directory scanning
- Automatic CSV logging
- Duplicate file protection
- Preserves original files

## Installation Instructions
1. Clone the repository
2. Copy the `*.psm1` files to your PowerShell modules directory
3. Import the modules in PowerShell

## Examples

### Example 1: Convert Legacy Documents to Modern Format
Convert all .doc and .xls files in a directory and its subdirectories to modern XML format:

```powershell
Convert-LegacyDocumentToXMLDocument -Path "C:\Documents\Archive" -Recursive
```

### Example 2: Convert to Open Document Format
Convert all Microsoft Office (Word, Excel and PowerPoint) files in a specific directory to ODF format:

```powershell
Convert-OfficeDocumentToOpenDocument -Path "D:\Projects\Reports" -Recursive
```

## Disclaimer
The code files in this repository were generated with assistance from artificial intelligence (AI). While the functionality has been tested and refined, please review the code carefully before using it in production environments. AI-assisted code may contain limitations or edge cases not covered by testing. Always maintain backups of your important documents before running batch conversion operations.

## License
This project is licensed under the MIT License - see the LICENSE.md file for details.