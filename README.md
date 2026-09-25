# ALCO-web-Downloader

PowerShell script for automatically downloading documents from an ALCO-web portal.

## Features

- Downloads the current accounting data.
- Stores all documents for each account in separate subdirectories.
- Saves all account transactions to a `kontoauszug.csv` file.
- Downloads all available PDF documents and prevents duplicate downloads.

## Requirements

* Windows 10/11
* Windows PowerShell 5.1
* Access to an ALCO-web portal

## Usage

1. Set up the login credentials using `credentials.ps1`. This encrypts the credentials and stores them in `credentials.xml`.

2. Set the URL of the ALCO-web portal in `download.ps1` by changing the `BaseUrl` variable.

3. Start the download using `download.ps1`. The downloaded documents are organized in the following directory structure:

```
Accounting period/
└── Account name/
    ├── PDF documents
    └── kontoauszug.csv
```
