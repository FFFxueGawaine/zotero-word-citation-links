param(
    [string]$TemplatePath = "",
    [string]$BasPath = "",
    [string]$BackupDir = "",
    [switch]$AllowWordRunning
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $TemplatePath) {
    if ($env:ZWL_TEMPLATE_PATH) {
        $TemplatePath = $env:ZWL_TEMPLATE_PATH
    }
    else {
        $TemplatePath = Join-Path $env:APPDATA "Microsoft\Word\STARTUP\Zotero.dotm"
    }
}

if (-not $BasPath) {
    $BasPath = Join-Path $PSScriptRoot "ZoteroWordHyperlinks.bas"
}

if (-not $BackupDir) {
    $BackupDir = Join-Path $PSScriptRoot "backup"
}

$backupName = "Zotero.backup.before-linking.dotm"
$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Security"
$separatorId = "ZoteroCitationLinksSeparator"
$createId = "ZoteroCreateCitationLinksButton"
$removeId = "ZoteroRemoveCitationLinksButton"
$setColorId = "ZoteroSetLinkColorButton"
$refreshId = "RefreshZotero"
$unlinkId = "ZoteroRemoveCodes"

function Assert-FileExists {
    param([string]$PathValue, [string]$Label)
    if (-not (Test-Path -LiteralPath $PathValue)) {
        throw "$Label not found: $PathValue"
    }
}

function Get-AccessVbomState {
    if (-not (Test-Path -LiteralPath $securityKey)) {
        New-Item -Path $securityKey -Force | Out-Null
    }

    $item = Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return @{
            Exists = $false
            Value = 0
        }
    }

    return @{
        Exists = $true
        Value = [int]$item.AccessVBOM
    }
}

function Set-AccessVbomState {
    param([bool]$Exists, [int]$Value)
    if (-not (Test-Path -LiteralPath $securityKey)) {
        New-Item -Path $securityKey -Force | Out-Null
    }

    if ($Exists) {
        New-ItemProperty -Path $securityKey -Name AccessVBOM -PropertyType DWord -Value $Value -Force | Out-Null
    }
    else {
        Remove-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    }
}

function Ensure-AccessVbomEnabled {
    if (-not (Test-Path -LiteralPath $securityKey)) {
        New-Item -Path $securityKey -Force | Out-Null
    }
    New-ItemProperty -Path $securityKey -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null
}

function Assert-WordNotRunning {
    if ($AllowWordRunning -or $env:ZWL_ALLOW_WORD_RUNNING -eq "1") {
        return
    }

    $wordProcesses = Get-Process WINWORD -ErrorAction SilentlyContinue
    if ($null -ne $wordProcesses) {
        throw "Word is running. Please close Microsoft Word and run install.bat again."
    }
}

function Backup-Template {
    param([string]$SourcePath, [string]$BackupFolder)

    New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
    $backupPath = Join-Path $BackupFolder $backupName
    Copy-Item -LiteralPath $SourcePath -Destination $backupPath -Force
    return $backupPath
}

function Get-CustomUiXmlBytes {
    param([string]$DotmPath)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = [System.IO.Compression.ZipFile]::OpenRead($DotmPath)
    try {
        $entry = $archive.GetEntry("customUI/customUI.xml")
        if ($null -eq $entry) {
            throw "customUI/customUI.xml was not found in $DotmPath"
        }

        $stream = $entry.Open()
        try {
            $memory = New-Object System.IO.MemoryStream
            $stream.CopyTo($memory)
            return $memory.ToArray()
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Update-CustomUiXml {
    param([byte[]]$XmlBytes)

    $xmlText = [System.Text.Encoding]::UTF8.GetString($XmlBytes)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($xmlText)

    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("ui", "http://schemas.microsoft.com/office/2006/01/customui")

    $group = $doc.SelectSingleNode("//ui:group[@id='ZoteroGroup']", $ns)
    if ($null -eq $group) {
        throw "ZoteroGroup was not found in customUI.xml"
    }

    foreach ($id in @($separatorId, $createId, $removeId, $setColorId)) {
        $node = $doc.SelectSingleNode("//ui:*[@id='$id']", $ns)
        if ($null -ne $node -and $null -ne $node.ParentNode) {
            [void]$node.ParentNode.RemoveChild($node)
        }
    }

    $refreshButton = $doc.SelectSingleNode("//ui:*[@id='$refreshId']", $ns)
    $unlinkButton = $doc.SelectSingleNode("//ui:*[@id='$unlinkId']", $ns)
    if ($null -eq $refreshButton) {
        throw "RefreshZotero button was not found in customUI.xml"
    }
    if ($null -eq $unlinkButton) {
        throw "ZoteroRemoveCodes button was not found in customUI.xml"
    }

    [void]$group.RemoveChild($unlinkButton)
    if ($null -ne $refreshButton.NextSibling) {
        [void]$group.InsertBefore($unlinkButton, $refreshButton.NextSibling)
    }
    else {
        [void]$group.AppendChild($unlinkButton)
    }

    $separator = $doc.CreateElement("separator", $ns.LookupNamespace("ui"))
    [void]$separator.SetAttribute("id", $separatorId)

    $createButton = $doc.CreateElement("button", $ns.LookupNamespace("ui"))
    [void]$createButton.SetAttribute("id", $createId)
    [void]$createButton.SetAttribute("label", "Create Citation Links")
    [void]$createButton.SetAttribute("imageMso", "HyperlinkInsert")
    [void]$createButton.SetAttribute("onAction", "ZoteroWordHyperlinks.ZoteroCreateCitationLinks")
    [void]$createButton.SetAttribute("supertip", "Create clickable links from Zotero citations to bibliography entries")
    [void]$createButton.SetAttribute("keytip", "K")

    $removeButton = $doc.CreateElement("button", $ns.LookupNamespace("ui"))
    [void]$removeButton.SetAttribute("id", $removeId)
    [void]$removeButton.SetAttribute("label", "Remove Citation Links")
    [void]$removeButton.SetAttribute("imageMso", "TableUnlinkExternalData")
    [void]$removeButton.SetAttribute("onAction", "ZoteroWordHyperlinks.ZoteroRemoveCitationLinks")
    [void]$removeButton.SetAttribute("supertip", "Remove citation links and bibliography bookmarks created by the hyperlink helper")
    [void]$removeButton.SetAttribute("keytip", "L")

    if ($null -ne $unlinkButton.NextSibling) {
        [void]$group.InsertBefore($separator, $unlinkButton.NextSibling)
    }
    else {
        [void]$group.AppendChild($separator)
    }

    if ($null -ne $separator.NextSibling) {
        [void]$group.InsertBefore($createButton, $separator.NextSibling)
    }
    else {
        [void]$group.AppendChild($createButton)
    }

    if ($null -ne $createButton.NextSibling) {
        [void]$group.InsertBefore($removeButton, $createButton.NextSibling)
    }
    else {
        [void]$group.AppendChild($removeButton)
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $true
    $settings.OmitXmlDeclaration = $false

    $memory = New-Object System.IO.MemoryStream
    $writer = [System.Xml.XmlWriter]::Create($memory, $settings)
    try {
        $doc.Save($writer)
    }
    finally {
        $writer.Dispose()
    }
    return $memory.ToArray()
}

function Replace-CustomUiInDotm {
    param([string]$DotmPath, [byte[]]$UpdatedCustomUiBytes)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tempPath = Join-Path ([System.IO.Path]::GetDirectoryName($DotmPath)) ([System.IO.Path]::GetRandomFileName() + ".dotm")

    $sourceArchive = [System.IO.Compression.ZipFile]::OpenRead($DotmPath)
    try {
        $targetArchive = [System.IO.Compression.ZipFile]::Open($tempPath, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($entry in $sourceArchive.Entries) {
                $newEntry = $targetArchive.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                $outStream = $newEntry.Open()
                try {
                    if ($entry.FullName -eq "customUI/customUI.xml") {
                        $outStream.Write($UpdatedCustomUiBytes, 0, $UpdatedCustomUiBytes.Length)
                    }
                    else {
                        $inStream = $entry.Open()
                        try {
                            $inStream.CopyTo($outStream)
                        }
                        finally {
                            $inStream.Dispose()
                        }
                    }
                }
                finally {
                    $outStream.Dispose()
                }
            }
        }
        finally {
            $targetArchive.Dispose()
        }
    }
    finally {
        $sourceArchive.Dispose()
    }

    Move-Item -LiteralPath $tempPath -Destination $DotmPath -Force
}

function Import-BasModule {
    param([string]$DotmPath, [string]$BasModulePath)

    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $document = $word.Documents.Open($DotmPath, $false, $false, $false)
        $project = $document.VBProject
        $components = $project.VBComponents

        for ($i = $components.Count; $i -ge 1; $i--) {
            $component = $components.Item($i)
            if ($component.Name -eq "ZoteroWordHyperlinks") {
                $components.Remove($component)
            }
        }

        [void]$components.Import($BasModulePath)
        [void]$document.Save()
    }
    finally {
        if ($null -ne $document) {
            $document.Saved = $true
            [void]$document.Close()
        }
        if ($null -ne $word) {
            [void]$word.Quit()
        }
    }
}

Assert-WordNotRunning
Assert-FileExists -PathValue $TemplatePath -Label "Template"
Assert-FileExists -PathValue $BasPath -Label "BAS module"

$originalAccessVbom = Get-AccessVbomState
try {
    Ensure-AccessVbomEnabled
    $backupPath = Backup-Template -SourcePath $TemplatePath -BackupFolder $BackupDir
    $customUiBytes = Get-CustomUiXmlBytes -DotmPath $TemplatePath
    $updatedCustomUiBytes = Update-CustomUiXml -XmlBytes $customUiBytes
    Replace-CustomUiInDotm -DotmPath $TemplatePath -UpdatedCustomUiBytes $updatedCustomUiBytes
    Import-BasModule -DotmPath $TemplatePath -BasModulePath $BasPath

    Write-Host "Backup created: $backupPath"
    Write-Host "Template updated: $TemplatePath"
    Write-Host "Install finished."
}
finally {
    Set-AccessVbomState -Exists $originalAccessVbom.Exists -Value $originalAccessVbom.Value
}
