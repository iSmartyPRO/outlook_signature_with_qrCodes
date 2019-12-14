# Outlook Signature from Active Directory with QR Codes vCard

This script is generate Outlook Signatures using VB Script.

In Outlook signature added vCard QR Code with full detail for employee.

All data is getting from Active Directory.

# Generate QR Codes

For generate QR Codes required to install Powershell Module

```
Install-Module -Name QRCodeGenerator -Scope CurrentUser -Force
```

After installation of Powershell Script need to replace improved powershell script from my repository: https://github.com/ilianapro/QRCodeGenerator_PowerShell

Especially replace only this file: https://github.com/ilianapro/QRCodeGenerator_PowerShell/blob/master/2.2/New-PSOneQRCodeVCard.ps1

to the following path: C:\Program Files\WindowsPowerShell\Modules\QRCodeGenerator

## How to generate QR Codes for users in specific OU?

* copy config_sample.json to config.json
* edit config.json file with your Data

Run following command via PowerShell Console:
```
.\generate_qr_codes.ps1
```

# Additional requirements

To get Active Directory Users need to install RSAT from following URL:

https://www.microsoft.com/ru-RU/download/details.aspx?id=45520