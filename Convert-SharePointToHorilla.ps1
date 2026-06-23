#Requires -Version 5.1
<#
.SYNOPSIS
    Converts a SharePoint staff list Excel export into a Horilla employee import template.

.DESCRIPTION
    Reads a SharePoint list export (.xlsx), maps columns to Horilla work-info import fields,
    and writes a filled workbook that matches the Horilla template headers exactly.

    Requires the ImportExcel module. The script installs it automatically on
    first run if it is missing.

.PARAMETER SharePointExportPath
    Path to the SharePoint Excel export. When omitted, the script looks for
    sharepoint_list.xlsx in the script folder or current directory.

.PARAMETER TemplatePath
    Path to the empty Horilla import template downloaded from Horilla (Actions > Import).
    Defaults to work_info_template.xlsx in the script folder or current directory.

.PARAMETER OutputPath
    Path for the generated Horilla import file. Defaults to horilla_import.xlsx
    next to the script.

.PARAMETER ConfigPath
    Optional JSON config for column mappings and default values.

.PARAMETER WorksheetName
    SharePoint worksheet name. When omitted, the first worksheet is used.

.EXAMPLE
    .\Convert-SharePointToHorilla.ps1

.EXAMPLE
    .\Convert-SharePointToHorilla.ps1 -ConfigPath .\sharepoint-horilla.config.json
#>
[CmdletBinding()]
param(
    [string]$SharePointExportPath,

    [string]$TemplatePath,

    [string]$OutputPath,

    [string]$ConfigPath,

    [string]$WorksheetName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ScriptDirectory {
    param(
        [string]$InvocationPath
    )

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        return $PSScriptRoot
    }

    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        return (Split-Path -Parent $PSCommandPath)
    }

    if (-not [string]::IsNullOrWhiteSpace($InvocationPath)) {
        return (Split-Path -Parent $InvocationPath)
    }

    return (Get-Location).Path
}

$Script:ScriptRoot = Get-ScriptDirectory -InvocationPath $MyInvocation.MyCommand.Path

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $Script:ScriptRoot 'sharepoint-horilla.config.json'
}

Set-Location -LiteralPath $Script:ScriptRoot

function Ensure-ImportExcelModule {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return
    }

    Write-Host 'ImportExcel module not found. Installing for current user (one-time setup)...'

    try {
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        }

        $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($gallery -and $gallery.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }

        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        Write-Host 'ImportExcel installed successfully.'
    }
    catch {
        throw @"
The ImportExcel module is required and automatic installation failed.

Run these commands manually in PowerShell, then try again:

    Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
    Install-PackageProvider -Name NuGet -Force
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-Module ImportExcel -Scope CurrentUser -Force

Then run:

    cd $Script:ScriptRoot
    .\Convert-SharePointToHorilla.ps1

Original error: $($_.Exception.Message)
"@
    }
}

function Resolve-InputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $resolved) {
        throw "$Label not found: $Path"
    }

    return $resolved.Path
}

function Find-NamedFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $searchRoots = @(
        $Script:ScriptRoot
        (Get-Location).Path
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($root in $searchRoots) {
        $candidate = Join-Path $root $FileName
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    foreach ($root in $searchRoots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        $found = Get-ChildItem -LiteralPath $root -Filter $FileName -File -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if ($found) {
            return $found.FullName
        }
    }

    throw "Could not find '$FileName'. Place it in the script folder or current directory."
}

function Resolve-DefaultPaths {
    param(
        [string]$SharePointExportPath,
        [string]$TemplatePath,
        [string]$OutputPath
    )

    if ([string]::IsNullOrWhiteSpace($SharePointExportPath)) {
        $SharePointExportPath = Find-NamedFile -FileName 'sharepoint_list.xlsx'
    }

    if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
        $TemplatePath = Find-NamedFile -FileName 'work_info_template.xlsx'
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $basePath = if (-not [string]::IsNullOrWhiteSpace($Script:ScriptRoot)) {
            $Script:ScriptRoot
        }
        else {
            Split-Path -Parent $SharePointExportPath
        }

        $OutputPath = Join-Path $basePath 'horilla_import.xlsx'
    }

    return [pscustomobject]@{
        SharePointExportPath = $SharePointExportPath
        TemplatePath         = $TemplatePath
        OutputPath           = $OutputPath
    }
}

function Get-Config {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config file not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    return $raw
}

function Normalize-ColumnName {
    param([string]$Name)

    if ($null -eq $Name) {
        return ''
    }

    return ($Name -replace '\s+$', '').Trim()
}

function Get-RowValue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Row,
        [string]$ColumnName
    )

    if ([string]::IsNullOrWhiteSpace($ColumnName)) {
        return $null
    }

    $normalizedTarget = Normalize-ColumnName $ColumnName
    foreach ($key in $Row.Keys) {
        if ((Normalize-ColumnName $key) -eq $normalizedTarget) {
            return $Row[$key]
        }
    }

    return $null
}

function ConvertTo-PlainText {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [datetime]) {
        return $Value.ToString('yyyy-MM-dd')
    }

    $text = [string]$Value
    if ($text -match '^\s*(\d+)\.0+\s*$') {
        return $Matches[1]
    }

    return $text.Trim()
}

function ConvertTo-HorillaDate {
    param(
        $Value,
        [string]$Format = 'yyyy-MM-dd'
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return ''
    }

    if ($Value -is [datetime]) {
        return $Value.ToString($Format)
    }

    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [int]) {
        try {
            return ([datetime]::FromOADate([double]$Value)).ToString($Format)
        }
        catch {
            return ''
        }
    }

    $parsed = $null
    if ([datetime]::TryParse([string]$Value, [ref]$parsed)) {
        return $parsed.ToString($Format)
    }

    return ConvertTo-PlainText $Value
}

function Split-FullName {
    param(
        [string]$FullName
    )

    $fullName = (ConvertTo-PlainText $FullName)
    if ([string]::IsNullOrWhiteSpace($fullName)) {
        return @{
            FirstName = ''
            LastName  = ''
        }
    }

    $parts = $fullName -split '\s+'
    if ($parts.Count -eq 1) {
        return @{
            FirstName = $parts[0]
            LastName  = ''
        }
    }

    return @{
        FirstName = $parts[0]
        LastName  = ($parts[1..($parts.Count - 1)] -join ' ')
    }
}

function Get-FullNameFromRow {
    param(
        [hashtable]$Row,
        [string[]]$NameSourceColumns
    )

    foreach ($column in $NameSourceColumns) {
        $value = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $column)
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return ''
}

function Resolve-MappedValue {
    param(
        [hashtable]$Map,
        [string]$Value
    )

    $text = ConvertTo-PlainText $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    foreach ($entry in $Map.PSObject.Properties) {
        if ($entry.Name -eq $text) {
            return [string]$entry.Value
        }
    }

    return $text
}

function New-PlaceholderEmail {
    param(
        [string]$BadgeId,
        [string]$FirstName,
        [string]$LastName,
        [string]$Domain
    )

    $slug = if (-not [string]::IsNullOrWhiteSpace($BadgeId)) {
        ($BadgeId -replace '[^A-Za-z0-9]+', '.').Trim('.').ToLowerInvariant()
    }
    else {
        $nameSlug = ("$FirstName.$LastName" -replace '[^A-Za-z0-9.]+', '.').Trim('.').ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($nameSlug)) {
            'employee'
        }
        else {
            $nameSlug
        }
    }

    return "$slug@$Domain"
}

function Test-ShouldSkipRow {
    param(
        [hashtable]$Row,
        $FilterConfig
    )

    $itemType = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName 'Item Type')
    if (-not [string]::IsNullOrWhiteSpace($itemType) -and $itemType -ne 'Item') {
        return $true
    }

    $title = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName 'Title')
    if ([string]::IsNullOrWhiteSpace($title) -or $title -eq 'Title') {
        return $true
    }

    if ($FilterConfig.SkipInactive) {
        $status = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName 'Statusi')
        foreach ($inactiveValue in $FilterConfig.InactiveStatusValues) {
            if ($status -eq $inactiveValue) {
                return $true
            }
        }
    }

    return $false
}

function Get-HorillaHeaders {
    param(
        [string]$Path
    )

    $package = Open-ExcelPackage -Path $Path
    try {
        $worksheet = $package.Workbook.Worksheets[1]
        $headers = @()
        $column = 1

        while ($true) {
            $cellValue = $worksheet.Cells[1, $column].Text
            if ([string]::IsNullOrWhiteSpace($cellValue)) {
                break
            }

            $headers += $cellValue
            $column++
        }

        if ($headers.Count -eq 0) {
            throw "No headers found in Horilla template: $Path"
        }

        return ,$headers
    }
    finally {
        Close-ExcelPackage $package -NoSave
    }
}

function Convert-SharePointRowToHorilla {
    param(
        [hashtable]$Row,
        [object]$Config,
        [string[]]$HorillaHeaders
    )

    $defaults = $Config.defaults
    $columnMap = $Config.columnMap
    $fullName = Get-FullNameFromRow -Row $Row -NameSourceColumns @($Config.nameSourceColumns)
    $nameParts = Split-FullName -FullName $fullName

    $badgeId = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.'Badge ID')
    $email = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.Email)
    if ([string]::IsNullOrWhiteSpace($email)) {
        $email = New-PlaceholderEmail `
            -BadgeId $badgeId `
            -FirstName $nameParts.FirstName `
            -LastName $nameParts.LastName `
            -Domain $defaults.PlaceholderEmailDomain
    }

    $rawGender = Get-RowValue -Row $Row -ColumnName $columnMap.Gender
    $gender = Resolve-MappedValue -Map $Config.genderMap -Value $rawGender

    $rawWorkType = Get-RowValue -Row $Row -ColumnName $columnMap.'Work Type'
    $workType = Resolve-MappedValue -Map $Config.workTypeMap -Value $rawWorkType
    if ([string]::IsNullOrWhiteSpace($workType)) {
        $workType = $defaults.WorkType
    }

    $employeeTypeSource = Get-RowValue -Row $Row -ColumnName $columnMap.'Work Type'
    $employeeType = Resolve-MappedValue -Map $Config.employeeTypeMap -Value $employeeTypeSource
    if ([string]::IsNullOrWhiteSpace($employeeType)) {
        $employeeType = $defaults.EmployeeType
    }

    $valuesByHeader = [ordered]@{
        'Badge ID'          = $badgeId
        'First Name'        = $nameParts.FirstName
        'Last Name'         = $nameParts.LastName
        'Email'             = $email
        'Phone'             = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.Phone)
        'Gender'            = $gender
        'Department'        = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.Department)
        'Job Position'      = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.'Job Position')
        'Job Role'          = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.'Job Role')
        'Shift'             = $defaults.Shift
        'Work Type'         = $workType
        'Reporting Manager' = $defaults.ReportingManager
        'Employee Type'     = $employeeType
        'Location'          = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.Location)
        'Date Joining'      = ConvertTo-HorillaDate -Value (Get-RowValue -Row $Row -ColumnName $columnMap.'Date Joining') -Format $defaults.DateFormat
        'Basic Salary'      = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.'Basic Salary')
        'Salary Hour'       = ConvertTo-PlainText (Get-RowValue -Row $Row -ColumnName $columnMap.'Salary Hour')
        'Contract End Date' = ConvertTo-HorillaDate -Value (Get-RowValue -Row $Row -ColumnName $columnMap.'Contract End Date') -Format $defaults.DateFormat
        'Company'           = $defaults.Company
    }

    $orderedRow = [ordered]@{}
    foreach ($header in $HorillaHeaders) {
        if ($valuesByHeader.Contains($header)) {
            $orderedRow[$header] = $valuesByHeader[$header]
        }
        else {
            $orderedRow[$header] = ''
        }
    }

    return [pscustomobject]$orderedRow
}

function Export-HorillaWorkbook {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,
        [Parameter(Mandatory = $true)]
        [string[]]$HorillaHeaders
    )

    $outputDirectory = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force

    if ($Rows.Count -eq 0) {
        return
    }

    $exportRows = foreach ($row in $Rows) {
        $line = [ordered]@{}
        foreach ($header in $HorillaHeaders) {
            $line[$header] = $row.$header
        }
        [pscustomobject]$line
    }

    $exportRows | Export-Excel -Path $OutputPath -WorksheetName 'Sheet1' -StartRow 2 -ClearSheet:$false -AutoSize
}

Ensure-ImportExcelModule

$resolvedPaths = Resolve-DefaultPaths `
    -SharePointExportPath $SharePointExportPath `
    -TemplatePath $TemplatePath `
    -OutputPath $OutputPath

$sharePointPath = Resolve-InputPath -Path $resolvedPaths.SharePointExportPath -Label 'SharePoint export'
$templatePath = Resolve-InputPath -Path $resolvedPaths.TemplatePath -Label 'Horilla template'
$OutputPath = $resolvedPaths.OutputPath
$config = Get-Config -Path $ConfigPath
$horillaHeaders = Get-HorillaHeaders -Path $templatePath

if ([string]::IsNullOrWhiteSpace($WorksheetName)) {
    $sheetInfo = Get-ExcelSheetInfo -Path $sharePointPath
    if (-not $sheetInfo -or $sheetInfo.Count -eq 0) {
        throw "No worksheets found in SharePoint export: $sharePointPath"
    }
    $WorksheetName = $sheetInfo[0].Name
    Write-Verbose "Using SharePoint worksheet: $WorksheetName"
}

$importedRows = Import-Excel -Path $sharePointPath -WorksheetName $WorksheetName
if (-not $importedRows) {
    throw "No rows found in SharePoint worksheet '$WorksheetName'."
}

$convertedRows = New-Object System.Collections.Generic.List[object]
$skippedCount = 0
$placeholderEmailCount = 0

foreach ($importedRow in $importedRows) {
  $rowHash = @{}
  foreach ($property in $importedRow.PSObject.Properties) {
    $rowHash[$property.Name] = $property.Value
  }

  if (Test-ShouldSkipRow -Row $rowHash -FilterConfig $config.filters) {
    $skippedCount++
    continue
  }

  $converted = Convert-SharePointRowToHorilla -Row $rowHash -Config $config -HorillaHeaders $horillaHeaders

  $originalEmail = ConvertTo-PlainText (Get-RowValue -Row $rowHash -ColumnName $config.columnMap.Email)
  if ([string]::IsNullOrWhiteSpace($originalEmail)) {
    $placeholderEmailCount++
  }

  $convertedRows.Add($converted) | Out-Null
}

Export-HorillaWorkbook -TemplatePath $templatePath -OutputPath $OutputPath -Rows $convertedRows -HorillaHeaders $horillaHeaders

Write-Host "Conversion complete."
Write-Host "  SharePoint file      : $sharePointPath"
Write-Host "  SharePoint worksheet : $WorksheetName"
Write-Host "  Rows converted       : $($convertedRows.Count)"
Write-Host "  Rows skipped         : $skippedCount"
Write-Host "  Placeholder emails   : $placeholderEmailCount"
Write-Host "  Output file          : $OutputPath"

if ($placeholderEmailCount -gt 0) {
  Write-Warning "Some employees had no email in SharePoint. Placeholder emails were generated using domain '$($config.defaults.PlaceholderEmailDomain)'. Update them in Horilla after import if needed."
}
