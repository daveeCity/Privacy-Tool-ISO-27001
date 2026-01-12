# Privacy-Tool-ISO-27001

> A standalone tool for the secure **pseudonymization** and **anonymization** of Excel files, designed to ensure GDPR and ISO 27001 compliance within HR processes.

## Overview

This tool was developed to enable the Human Resources department to share files containing sensitive data (PII - Personally Identifiable Information) securely.
Unlike standard scripts, this software implements **enterprise-grade security and robustness logic**: it handles complex cells (hyperlinks, emails), maintains the integrity of original files, and records every operation in a central audit database.

### Key Features

* ** Hybrid Pseudonymization:** Replaces sensitive data with unique codes (e.g., `USER_001`). The decoding key is saved in a `mappatura_segreta.csv` file created locally in the working folder, ensuring data segregation per project.
* ** Irreversible Anonymization:** Applies **SHA-256** hashing to permanently remove the original data.
* ** Centralized Audit Trail:** Every operation is logged to an external MySQL database (Who, What, When).
* ** "Anti-Crash" Robustness:** Proprietary `Get-SafeString` algorithm to handle "dirty" Excel cells (formulas, hyperlinks, merged cells).
* ** Data Integrity First:** **Copy-on-Write** strategy: the tool never modifies the original file, working exclusively on a cloned copy.

---

## System Requirements

* **OS:** Windows 10 / 11 / Server.
* **Runtime:** PowerShell 5.1 (Pre-installed on Windows) and .NET Framework 4.8+.
* **Software:** Microsoft Excel installed (required for COM Interop libraries).
* **Driver:** `MySql.Data.dll` (See Installation).

---

## Installation and Configuration

For security reasons (ISO 27001 - Control A.8.24), **database credentials are not included in the source code**.
Follow these steps to configure the environment.

### 1. Clone the Repository
Download the files to a local folder

### 2. Add the MySQL Driver
The `MySql.Data.dll` file is not included in the repository to keep it lightweight.
1.  Download the official package from [NuGet MySql.Data](https://www.nuget.org/packages/MySql.Data).
2.  Rename the downloaded `.nupkg` file to `.zip`.
3.  Extract the `MySql.Data.dll` file from the `lib/net48` folder.
4.  Place it in the root folder of the project.

### 3. Create the Config File
Create a file named **`config.json`** in the same folder as the `.ps1` script.
Copy and paste this template, inserting your real credentials:

```json
{
    "DbServer": "192.168.X.X",
    "DbName": "DATABASE_NAME",
    "DbUser": "RESTRICTED_LOG_USER",
    "DbPassword": "YOUR_SECURE_PASSWORD"
}

```
Usage

    Run the script HR_Privacy_Tool.ps1 (Right-click -> Run with PowerShell).

    Step 1: Click on "Select Excel File..." and choose the file to process.

    Step 2: The tool parses the header. Select the columns you want to hide from the list (e.g., Salary, Name, IBAN).

    Step 3: Choose the mode:

        Pseudonymization: Creates a _PSEUDO.xlsx file and updates the local CSV mapping.

        Anonymization: Creates an _ANON.xlsx file with encrypted data (irreversible).

    Click on "START PROCESS".

Project Structure

/HR_Privacy_Tool
│
├── HR_Privacy_Tool.ps1      # Core Application
├── config.json              # Credentials (TO BE CREATED - IGNORED BY GIT)
├── MySql.Data.dll           # Database Driver (TO BE ADDED)
├── README.md                # Documentation
└── .gitignore               # Git exclusion rules

Security & Compliance Notes

This software implements security measures compliant with the ISO/IEC 27001 standard:

    Separation of Secrets: No passwords are hard-coded.

    Least Privilege: The DB user should have limited permissions (INSERT only) on the log table.

    Input Sanitization: Prevention of errors caused by unexpected data types.

    Local Mapping: The reconciliation table (CSV) resides locally (or on a protected network share) and never leaves the corporate perimeter.
