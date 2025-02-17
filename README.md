## PDF Page Count in SharePoint Document Library through PowerShell

## Description
This PowerShell script automates the process of counting pages in PDF files stored in a SharePoint Online Document Library. It utilizes the **PnP PowerShell module** to connect to SharePoint and the **PdfSharp** library to extract the page count from each PDF. The script updates a custom SharePoint column `PDFPageCount` with the number of pages in each PDF file.

## Features ✨
✅ Connects to SharePoint Online securely using PnP PowerShell.  
✅ Retrieves all PDF files from the specified SharePoint document library.  
✅ Downloads each PDF file temporarily to a local directory.  
✅ Uses PdfSharp to extract the total page count.  
✅ Updates the `PDFPageCount` column in the SharePoint document library.  
✅ Cleans up downloaded files after processing.  

## Requirements ⚙️
📌 **PnP PowerShell module** (Install with `Install-Module PnP.PowerShell -Force -Scope CurrentUser`)  
📌 **PdfSharp.dll** (Download and place in a known directory, e.g., `C:\Users\YourUser\Documents\Office_Files\PdfSharp.dll`)  
📌 **SharePoint Online access** with necessary permissions to read and update list items.  
📌 **PowerShell execution policy** allowing script execution (`Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`).  

## Installation & Setup 🛠️
1. **Install required PowerShell modules:**  
   ```powershell
   Install-Module PnP.PowerShell -Force -Scope CurrentUser
   ```

2. **Download and place PdfSharp.dll** at a known directory.

3. **Update script variables** with your SharePoint site URL and library name.

4. **Run the script in PowerShell:**  
   ```powershell
   .\PDFPageCount_SharePoint.ps1
   ```

## Script Usage 🚀

### Load PnP PowerShell Module
```powershell
Import-Module PnP.PowerShell
```

### Connect to SharePoint Online
```powershell
$SiteURL = "https://yourtenant.sharepoint.com/sites/yoursite"
Connect-PnPOnline -Url $SiteURL -UseWebLogin
```

### Load PdfSharp DLL
```powershell
Add-Type -Path "C:\Users\YourUser\Documents\Office_Files\PdfSharp.dll"
```

### Function to Get PDF Page Count
```powershell
Function Get-PDFPageCount {
    param ($FilePath)
    $PDF = [PdfSharp.Pdf.IO.PdfReader]::Open($FilePath, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::ReadOnly)
    $PageCount = $PDF.PageCount
    $PDF.Close()
    return $PageCount
}
```

### Define SharePoint Library
```powershell
$LibraryName = "YourDocumentLibrary"
```

### Retrieve PDF Files and Process
```powershell
$PDFFiles = Get-PnPListItem -List $LibraryName | Where-Object { $_["File_x0020_Type"] -eq "pdf" }

foreach ($File in $PDFFiles) {
    $FilePath = $File["FileRef"]
    $LocalFile = "C:\Temp\" + $File["FileLeafRef"]

    Get-PnPFile -Url $FilePath -Path "C:\Temp\" -FileName $File["FileLeafRef"] -AsFile -Force
    
    $PageCount = Get-PDFPageCount -FilePath $LocalFile
    
    Set-PnPListItem -List $LibraryName -Identity $File.Id -Values @{"PDFPageCount" = $PageCount}
    
    Write-Host "Updated $($File['FileLeafRef']) with $PageCount pages"
}
```

### Disconnect SharePoint Session
```powershell
Disconnect-PnPOnline
```


