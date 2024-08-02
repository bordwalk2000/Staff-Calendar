# StaffCalendar PowerShell Module

A module with different tools around creating staff calendars for a specified year, either from a list of users with the same work hours or from a CSV file.

## 🚀 Quick start

Quickly create quick Excel file.

```powershell
New-StaffCalendar -year 1997 -users "Jack O", "Sam C", "Daniel J"
````

## 💿 Installation

The StaffCalendar PowerShell Module is published to the PowerShell Gallery.  You can install it into your user profile by running the following command.

```powershell
Install-Module -Name StaffCalendar
```

## 💽 Developer Instructions

If you want to run this module from source it can found at [GitHub](https://github.com/bordwalk2000/StaffCalendar).  The can be built with the ModuleBuilder module and then running the following command.

```powershell
Start-ModuleBuild.ps1
```

This will package all code into files located in .\Output\StaffCalendar.  That folder is now ready to be installed, copy to any path listed in you PSModulePath environment variable and you are good to go!

### Source Files Folder structure

- All building files must in Source folders:
  - In the root, place the module manifest
    - In Public, place functions accessible by users
    - In Private, place functions that inaccessible by users, e.g. helper functions
    - Place one function per file, and file name must match the name of the function
- In the root of the repository we have:
  - Start-ModuleBuild.ps1, builds the module with files from Source folder and puts them into Output folder
