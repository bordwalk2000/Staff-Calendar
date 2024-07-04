#Requires -Module ModuleBuilder
param(
    [version]$Version = (Import-PowerShellDataFile "$PSScriptRoot\Source\StaffCalendar.psd1").ModuleVersion
)

$params = @{
    SourcePath                 = "$PSScriptRoot\Source\StaffCalendar.psd1"
    CopyPaths                  = @("$PSScriptRoot\README")
    Version                    = $Version
    UnversionedOutputDirectory = $false
}
Build-Module @params