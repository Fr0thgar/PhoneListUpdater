# PowerShell Script for Updating SharePoint Phone List

## Overview
This PowerShell script automates the process of updating a SharePoint list (**Telefonliste**) with data from a CSV file. It ensures that:
- The correct version of **PnP.PowerShell** is installed and updated if necessary.
- The script connects to SharePoint securely.
- It checks if the list and necessary fields exist before performing updates.
- Data from the CSV file is synchronized with the SharePoint list.

---

## Prerequisites
Before running the script, ensure that:
1. You have **PowerShell 7+** installed.
2. You have **SharePoint Admin** permissions or sufficient rights to modify lists.
3. You have installed or have access to **PnP.PowerShell**.
4. Your SharePoint list "Telefonliste" exists and has the correct fields.
5. Your CSV file is formatted correctly.

---

## Setup Instructions
### 1️⃣ Install PowerShell 7+ (if not installed)
Download and install PowerShell from:
[https://github.com/PowerShell/PowerShell/releases](https://github.com/PowerShell/PowerShell/releases)

To check your current version, run:
```powershell
$PSVersionTable.PSVersion
```

---

### 2️⃣ Install or Update PnP.PowerShell
The script automatically checks and updates **PnP.PowerShell** if needed. However, you can manually install it using:
```powershell
Install-Module -Name PnP.PowerShell -Force
```

To verify the installed version:
```powershell
Get-Module -Name PnP.PowerShell -ListAvailable
```

To update to the latest version manually:
```powershell
Update-Module -Name PnP.PowerShell -Force
```

---

### 3️⃣ Configure SharePoint Credentials
Ensure you have the necessary permissions in SharePoint.
- **Application Permissions:** `Sites.FullControl.All`
- **Delegated Permissions:** `AllSites.FullControl`, `AllSites.Read`

When running the script, you will be prompted to authenticate interactively.

---

### 4️⃣ Prepare the CSV File
Ensure your CSV file is correctly formatted with the following headers:
```csv
Fuldenavn,Direkte,Mobil,Mail,Stilling,Afdeling,Firma,Lokation
```
Save the file at the specified path in the script.

---

### 5️⃣ Run the Script
Execute the script in PowerShell by navigating to the directory where it's saved and running:
```powershell
.\YourScript.ps1
```
If necessary, allow script execution:
```powershell
Set-ExecutionPolicy Unrestricted -Scope Process
```

---

## Troubleshooting
### 🔹 Common Issues & Fixes
#### 1. **"Access is denied" (HRESULT: 0x80070005)**
- Ensure you have SharePoint admin permissions.
- Run PowerShell as an administrator.
- Check permissions for your SharePoint list.
- Close Terminal and run it again.

#### 2. **"PnP.PowerShell not recognized"**
- Ensure the module is installed: `Install-Module -Name PnP.PowerShell`
- Restart PowerShell after installation.

#### 3. **"Column does not exist"**
- Verify SharePoint internal field names using:
  ```powershell
  Get-PnPField -List "Telefonliste" | Select Title, InternalName
  ```
- Update the script to use the correct internal names.

#### 4. **CSV Data Not Updating in SharePoint**
- Ensure the CSV is saved properly and accessible.
- Check that column names match the internal SharePoint field names.

---

## Notes
- This script is designed to be **safe** and will not make unnecessary changes if data is already up to date.
- If needed, modify the field mappings in the script to align with your SharePoint list structure.

---

## ✅ Summary
1. Install PowerShell 7+ ✅
2. Install/Update PnP.PowerShell ✅
3. Ensure SharePoint permissions ✅
4. Format CSV correctly ✅
5. Run the script ✅

---

For further support, consult the [PnP.PowerShell Documentation](https://pnp.github.io/powershell/) or your SharePoint administrator.

