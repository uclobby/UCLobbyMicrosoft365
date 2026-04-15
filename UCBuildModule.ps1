param(
    [version]$Version = "1.1.0"
)

#Getting module name from the current script path
$ModuleName = Split-Path $PSScriptRoot -Leaf

#Check if Output path exists, if so deletes previous files.
if (!(Test-Path -Path "$PSScriptRoot\Output\$ModuleName" -PathType Container)) {
    New-Item -Path "$PSScriptRoot\Output\$ModuleName" -ItemType Directory | Out-Null
}
else {
    Remove-Item "$PSScriptRoot\Output\$ModuleName\*.*" -Recurse -Force
}

#region Copy Readme and Format file if exists.
if (Test-Path -Path "$PSScriptRoot\LICENSE" -PathType Leaf) {
    Copy-Item -Path "$PSScriptRoot\LICENSE" -Destination "$PSScriptRoot\Output\$ModuleName\LICENSE"
}

if (Test-Path -Path "$PSScriptRoot\README.md" -PathType Leaf) {
    Copy-Item -Path "$PSScriptRoot\README.md" -Destination "$PSScriptRoot\Output\$ModuleName\README.md"
}

if (Test-Path -Path "$PSScriptRoot\Source\$ModuleName.format.ps1xml" -PathType Leaf) {
    Copy-Item -Path "$PSScriptRoot\Source\$ModuleName.format.ps1xml" -Destination "$PSScriptRoot\Output\$ModuleName\$ModuleName.format.ps1xml"
}
#endregion

#region Create Module 
$outModuleContent = ""
$FunctionsToExport = "@("

Get-ChildItem -Path ("$PSScriptRoot\Source\Private") -Filter *.ps1 | ForEach-Object { $outModuleContent += [System.IO.File]::ReadAllText($_.FullName) }
Get-ChildItem -Path ("$PSScriptRoot\Source\Public") -Filter *.ps1 | ForEach-Object { $outModuleContent += [System.IO.File]::ReadAllText($_.FullName); $FunctionsToExport += "'" + $_.BaseName + "'," }

[System.IO.File]::WriteAllText("$PSScriptRoot\Output\$ModuleName\$ModuleName.psm1", ($outModuleContent -join "`n`n"), [System.Text.Encoding]::UTF8)
$FunctionsToExport = $FunctionsToExport.Substring(0, $FunctionsToExport.Length - 1) + ")"
#endregion

#region Update and copy Module Manifest
if (Test-Path -Path "$PSScriptRoot\Source\$ModuleName.psd1" -PathType Leaf) {
    $ModuleManifest = [System.IO.File]::ReadAllText("$PSScriptRoot\Source\$ModuleName.psd1")
    $ModuleManifest = $ModuleManifest -replace '#%ModuleVersion%', ("""$Version""")
    $ModuleManifest = $ModuleManifest -replace '#%FunctionsToExport%', $FunctionsToExport
    [System.IO.File]::WriteAllText("$PSScriptRoot\Output\$ModuleName\$ModuleName.psd1", ($ModuleManifest -join "`n`n"), [System.Text.Encoding]::UTF8)
}
else {
    Write-Warning "Missing Manifest file: $ModuleName.psd1"
}
#endregion

"$PSScriptRoot\Output\$ModuleName"
Import-Module -Name ("$PSScriptRoot\Output\$ModuleName") -Force
