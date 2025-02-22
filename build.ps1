<#

. ".\build.ps1" -Verbose
. "C:\Git\SqlQueryEditor\build.ps1" -Verbose


#>
[CmdletBinding()]
Param ()

$ErrorActionPreference = 'Stop'

# Get Project Configuration Data
$data = Get-MTProjectInfo
try {
    Write-Verbose 'Running dist folder reset'
    if (Test-Path $data.OutputDir) {
        Remove-Item -Path $data.OutputDir -Recurse -Force
    }
    # Setup Folders
    New-Item -Path $data.OutputDir -ItemType Directory -Force | Out-Null # Dist folder
    New-Item -Path $data.OutputModuleDir -Type Directory -Force | Out-Null # Module Folder
} catch {
    Write-Error 'Failed to reset Dist folder'
}

Write-Verbose ($data | Out-String)

function Get-FunctionNameFromFile {
    param($filePath)
    try {
        $moduleContent = Get-Content -Path $filePath -Raw
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($moduleContent, [ref]$null, [ref]$null)
        $functionName = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) | ForEach-Object { $_.Name } 
        return $functionName
    }
    catch { return '' }
}

# Copy Project Resources to distribution folder for testing and publishing 
Write-Verbose 'Copying Project Resources to distribution folder'
$resFolder = [System.IO.Path]::Join($data.ProjectRoot, 'src', 'resources')
if (Test-Path $resFolder) {
    $items = Get-ChildItem -Path $resFolder -ErrorAction SilentlyContinue
    if ($items) {
        Write-Verbose 'Files found in resource folder, copying resource folder content'
        foreach ($item in $items) {
            Copy-Item -Path $item.FullName -Destination ($data.OutputModuleDir) -Recurse -Force -ErrorAction Stop
        }
    }
}

Write-Verbose 'Building Application files'
# Test-ProjectSchema -Schema Build | Out-Null
$sb = [System.Text.StringBuilder]::new()
# Public Folder
$files = Get-ChildItem -Path $data.PublicDir -Filter *.ps1
$files | ForEach-Object {
    $sb.AppendLine([IO.File]::ReadAllText($_.FullName)) | Out-Null
}

# Private Folder
$files = Get-ChildItem -Path $data.PrivateDir -Filter *.ps1 -ErrorAction SilentlyContinue
foreach ($file in $files) {
    $sb.AppendLine([IO.File]::ReadAllText($file.FullName)) | Out-Null
}
try {
    Set-Content -Path $data.ModuleFilePSM1 -Value $sb.ToString() -Encoding 'UTF8' -ErrorAction Stop # psm1 file
} catch {
    Write-Error 'Failed to create psm1 file' -ErrorAction Stop
}

Write-Verbose 'Building psd1 data file Manifest'

    ## TODO - DO schema check

    $PubFunctionFiles = Get-ChildItem -Path $data.PublicDir -Filter *.ps1
    $functionToExport = @()
    $PubFunctionFiles | ForEach-Object {
        $functionToExport += Get-FunctionNameFromFile -filePath $_.FullName
    }

    $ManifestAllowedParams = (Get-Command New-ModuleManifest).Parameters.Keys

    $ParmsManifest = @{
        Path              = $data.ManifestFilePSD1
        Description       = $data.Description
        FunctionsToExport = $functionToExport
        RootModule        = "$($data.ProjectName).psm1"
        ModuleVersion     = $data.Version
    }

    # Accept only valid Manifest Parameters
    $data.Manifest.Keys | ForEach-Object {
        if ( $ManifestAllowedParams -contains $_) {
            if ($data.Manifest.$_) {
                $ParmsManifest.add($_, $data.Manifest.$_ )
            }
        } else {
            Write-Warning "Unknown parameter $_ in Manifest"
        }
    }

    try {
        New-ModuleManifest @ParmsManifest -ErrorAction Stop
    } catch {
        'Failed to create Manifest: {0}' -f $_.Exception.Message | Write-Error -ErrorAction Stop
    }

