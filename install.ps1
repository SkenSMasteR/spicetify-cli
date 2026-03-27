param (
  [string]$v
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region Variables
$spicetifyFolderPath = "$env:LOCALAPPDATA\spicetify"
$spicetifyOldFolderPath = "$HOME\spicetify-cli"
$script:ui = $null
#endregion Variables

#region UI Functions
function Set-CursorPosition {
  param ([int]$x, [int]$y)
  $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates($x, $y)
}

function Write-At {
  param (
    [int]$x,
    [int]$y,
    [string]$text,
    [string]$color = 'Gray'
  )
  if (-not $script:ui) { return }
  if ($y -lt 0 -or $y -ge $script:ui.height) { return }
  Set-CursorPosition -x $x -y $y
  Write-Host -NoNewline -ForegroundColor $color -Object $text
}

function Clear-RectLine {
  param ([int]$x, [int]$y, [int]$width)
  Write-At -x $x -y $y -text (' ' * $width)
}

function Initialize-InstallTui {
  $window = $Host.UI.RawUI.WindowSize
  $width = $window.Width
  $height = $window.Height
  if ($width -lt 60 -or $height -lt 20) {
    throw "Console too small for TUI ($width x $height). Minimum supported size is 60 x 20."
  }
  $leftWidth = [Math]::Floor($width / 2)
  $rightStart = $leftWidth + 1

  $script:ui = @{
    width = $width
    height = $height
    leftWidth = $leftWidth
    rightStart = $rightStart
    statusRow = [Math]::Floor($height / 2) - 3
    progressRow = [Math]::Floor($height / 2)
    logRow = 2
    logTop = 2
    logBottom = $height - 3
  }

  Clear-Host
  for ($y = 0; $y -lt $script:ui.height; $y++) {
    Write-At -x $script:ui.leftWidth -y $y -text '|'
  }

  $leftTitle = 'STATUS'
  $rightTitle = 'LOGS'
  $leftX = [Math]::Max(0, [Math]::Floor(($script:ui.leftWidth - $leftTitle.Length) / 2))
  Write-At -x $leftX -y 0 -text $leftTitle -color 'Cyan'
  Write-At -x ($script:ui.rightStart + 2) -y 0 -text $rightTitle -color 'Cyan'
}

function Write-LeftCentered {
  param (
    [string]$text,
    [int]$row,
    [string]$color = 'Gray'
  )
  $maxLen = [Math]::Max(1, $script:ui.leftWidth - 2)
  if ($text.Length -gt $maxLen) {
    $text = $text.Substring(0, $maxLen)
  }
  Clear-RectLine -x 0 -y $row -width $script:ui.leftWidth
  $x = [Math]::Max(0, [Math]::Floor(($script:ui.leftWidth - $text.Length) / 2))
  Write-At -x $x -y $row -text $text -color $color
}

function Write-Log {
  param (
    [string]$message,
    [string]$color = 'Gray'
  )
  if (-not $script:ui) {
    Write-Host $message
    return
  }

  $maxLen = [Math]::Max(10, $script:ui.width - $script:ui.rightStart - 3)
  if ($message.Length -gt $maxLen) {
    $message = $message.Substring(0, $maxLen - 3) + '...'
  }

  if ($script:ui.logRow -gt $script:ui.logBottom) {
    for ($row = $script:ui.logTop; $row -le $script:ui.logBottom; $row++) {
      Clear-RectLine -x ($script:ui.rightStart + 1) -y $row -width ($script:ui.width - $script:ui.rightStart - 1)
    }
    $script:ui.logRow = $script:ui.logTop
  }

  Clear-RectLine -x ($script:ui.rightStart + 1) -y $script:ui.logRow -width ($script:ui.width - $script:ui.rightStart - 1)
  Write-At -x ($script:ui.rightStart + 1) -y $script:ui.logRow -text $message -color $color
  $script:ui.logRow++
}

function Update-Progress {
  param ([int]$percent)
  $p = [Math]::Max(0, [Math]::Min(100, $percent))
  $barWidth = [Math]::Max(10, [Math]::Min(36, $script:ui.leftWidth - 10))
  $filled = [Math]::Floor(($p / 100) * $barWidth)
  $bar = ('#' * $filled) + ('-' * ($barWidth - $filled))
  $text = "[{0}] {1,3}%" -f $bar, $p
  Write-LeftCentered -text $text -row $script:ui.progressRow -color 'Cyan'
}

function Set-Status {
  param (
    [string]$text,
    [int]$percent,
    [string]$color = 'Gray'
  )
  if (-not $script:ui) {
    Write-Host $text -ForegroundColor $color
    return
  }
  Write-LeftCentered -text $text -row $script:ui.statusRow -color $color
  Update-Progress -percent $percent
}

function Move-CursorToBottom {
  if (-not $script:ui) { return }
  Set-CursorPosition -x 0 -y ($script:ui.height - 1)
  Write-Host ''
}

function Enter-StandardOutput {
  Move-CursorToBottom
  Write-Host ''
  $script:ui = $null
}
#endregion UI Functions

#region Core Functions
function Write-Success {
  Write-Log -message ' > OK' -color 'Green'
}

function Write-Unsuccess {
  Write-Log -message ' > ERROR' -color 'Red'
}

function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
  -not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PowerShellVersion {
  $PSMinVersion = [version]'5.1'
  $PSVersionTable.PSVersion -ge $PSMinVersion
}

function Move-OldSpicetifyFolder {
  if (Test-Path -Path $spicetifyOldFolderPath) {
    Write-Log -message 'Moving old spicetify folder...'
    Copy-Item -Path "$spicetifyOldFolderPath\*" -Destination $spicetifyFolderPath -Recurse -Force
    Remove-Item -Path $spicetifyOldFolderPath -Recurse -Force
    Write-Log -message 'Old folder moved.' -color 'Green'
  }
  else {
    Write-Log -message 'No old spicetify folder found.'
  }
}

function Get-Spicetify {
  if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
    $architecture = 'x64'
  }
  elseif ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
    $architecture = 'arm64'
  }
  else {
    $architecture = 'x32'
  }

  if ($v) {
    if ($v -match '^v?(\d+\.\d+\.\d+)$') {
      $targetVersion = $Matches[1]
    }
    else {
      throw "Invalid spicetify version: $v (expected x.y.z)"
    }
  }
  else {
    Write-Log -message 'Fetching latest spicetify version...'
    $previousProgressPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
      $latestRelease = Invoke-RestMethod -Uri 'https://api.github.com/repos/spicetify/cli/releases/latest'
    }
    finally {
      $ProgressPreference = $previousProgressPreference
    }
    $targetVersion = $latestRelease.tag_name -replace 'v', ''
    Write-Log -message "Latest version: v$targetVersion" -color 'Green'
  }

  $archivePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'spicetify.zip')
  Write-Log -message "Downloading spicetify v$targetVersion..."

  $downloadUri = "https://github.com/spicetify/cli/releases/download/v$targetVersion/spicetify-$targetVersion-windows-$architecture.zip"
  Download-FileWithProgress -Uri $downloadUri -OutFile $archivePath -StageStartPercent 35 -StageEndPercent 70

  Write-Log -message 'Download complete.' -color 'Green'
  @{
    ArchivePath = $archivePath
    Version = $targetVersion
  }
}

function Download-FileWithProgress {
  param (
    [string]$Uri,
    [string]$OutFile,
    [int]$StageStartPercent = 35,
    [int]$StageEndPercent = 70
  )

  $previousProgressPreference = $ProgressPreference
  $ProgressPreference = 'SilentlyContinue'
  $request = $null
  $response = $null
  $responseStream = $null
  $fileStream = $null

  try {
    $request = [System.Net.HttpWebRequest]::Create($Uri)
    $request.Method = 'GET'
    $request.AllowAutoRedirect = $true
    $request.Timeout = 1800000
    $request.ReadWriteTimeout = 1800000

    $response = $request.GetResponse()
    $totalBytes = $response.ContentLength
    $responseStream = $response.GetResponseStream()
    $fileStream = [System.IO.File]::Open($OutFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)

    $buffer = New-Object byte[] 65536
    $bytesReadTotal = 0L
    $lastDownloadPercent = -1

    while (($bytesRead = $responseStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
      $fileStream.Write($buffer, 0, $bytesRead)
      $bytesReadTotal += $bytesRead

      if ($totalBytes -gt 0) {
        $downloadPercent = [int](($bytesReadTotal * 100) / $totalBytes)
        if ($downloadPercent -ne $lastDownloadPercent) {
          $mappedPercent = $StageStartPercent + [int](($downloadPercent * ($StageEndPercent - $StageStartPercent)) / 100)
          Set-Status -text "Downloading package... $downloadPercent%" -percent $mappedPercent -color 'Yellow'
          $lastDownloadPercent = $downloadPercent
        }
      }
    }

    Set-Status -text 'Downloading package... 100%' -percent $StageEndPercent -color 'Yellow'
  }
  finally {
    if ($fileStream) { $fileStream.Dispose() }
    if ($responseStream) { $responseStream.Dispose() }
    if ($response) { $response.Dispose() }
    $ProgressPreference = $previousProgressPreference
  }
}

function Add-SpicetifyToPath {
  Write-Log -message 'Updating user PATH...'
  $user = [EnvironmentVariableTarget]::User
  $userPath = [Environment]::GetEnvironmentVariable('PATH', $user)
  $userPath = $userPath -replace "$([regex]::Escape($spicetifyOldFolderPath))\\*;*", ''
  if ($userPath -notlike "*$spicetifyFolderPath*") {
    $userPath = "$userPath;$spicetifyFolderPath"
  }

  [Environment]::SetEnvironmentVariable('PATH', $userPath, $user)

  $processPath = $env:PATH
  $processPath = $processPath -replace "$([regex]::Escape($spicetifyOldFolderPath))\\*;*", ''
  if ($processPath -notlike "*$spicetifyFolderPath*") {
    $env:PATH = "$processPath;$spicetifyFolderPath"
  }
  else {
    $env:PATH = $processPath
  }

  Write-Log -message 'PATH updated.' -color 'Green'
}

function Install-Spicetify {
  Set-Status -text 'Preparing install...' -percent 10 -color 'Yellow'
  Move-OldSpicetifyFolder

  Set-Status -text 'Downloading package...' -percent 35 -color 'Yellow'
  $download = Get-Spicetify

  Set-Status -text 'Extracting files...' -percent 70 -color 'Yellow'
  Expand-Archive -Path $download.ArchivePath -DestinationPath $spicetifyFolderPath -Force
  Write-Log -message "Extracted spicetify v$($download.Version)." -color 'Green'

  Set-Status -text 'Configuring PATH...' -percent 85 -color 'Yellow'
  Add-SpicetifyToPath

  Set-Status -text 'Finalizing...' -percent 95 -color 'Yellow'
  Remove-Item -Path $download.ArchivePath -Force -ErrorAction SilentlyContinue
  Write-Log -message 'Temporary archive removed.'

  Set-Status -text 'SUCCESS' -percent 100 -color 'Green'
  Write-Log -message 'spicetify installation completed.' -color 'Green'
  Write-Log -message 'Run "spicetify -h" to get started.' -color 'Cyan'
}

function Install-Marketplace {
  Write-Host 'Starting Marketplace installation...'
  $marketplaceInstallerCommit = 'e6b09f31e84e039ac4753216cd5aedc748ccd88f'
  $marketplaceInstallerSha256 = '0962C57F8E36936228429B3A3C5CADABFCC9BBD9C3688C884180AF917DC33F02'
  $marketplaceInstallerUrl = "https://raw.githubusercontent.com/spicetify/spicetify-marketplace/$marketplaceInstallerCommit/resources/install.ps1"
  $tempScript = Join-Path $env:TEMP 'spicetify-marketplace-install.ps1'
  $downloadParams = @{
    Uri = $marketplaceInstallerUrl
    UseBasicParsing = $true
    OutFile = $tempScript
  }

  $previousProgressPreference = $ProgressPreference
  $ProgressPreference = 'SilentlyContinue'
  try {
    Invoke-WebRequest @downloadParams
    $actualSha256 = (Get-FileHash -Path $tempScript -Algorithm SHA256).Hash
    if ($actualSha256 -ne $marketplaceInstallerSha256) {
      throw "Marketplace installer checksum mismatch. Expected $marketplaceInstallerSha256, got $actualSha256."
    }
    $powershellExe = Join-Path $PSHOME 'powershell.exe'
    if (-not (Test-Path -Path $powershellExe)) {
      $powershellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    }

    if (Test-Path -Path $powershellExe) {
      & $powershellExe -NoProfile -ExecutionPolicy Bypass -File $tempScript
      if ($LASTEXITCODE -ne 0) {
        throw "Marketplace installer exited with code $LASTEXITCODE"
      }
    }
    else {
      $powershellFromPath = Get-Command powershell.exe -ErrorAction SilentlyContinue
      if ($powershellFromPath) {
        & $powershellFromPath.Source -NoProfile -ExecutionPolicy Bypass -File $tempScript
        if ($LASTEXITCODE -ne 0) {
          throw "Marketplace installer exited with code $LASTEXITCODE"
        }
      }
      else {
        & $tempScript
        if ($LASTEXITCODE -ne 0) {
          throw "Marketplace installer exited with code $LASTEXITCODE"
        }
      }
    }
  }
  finally {
    $ProgressPreference = $previousProgressPreference
    Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue
  }

  Write-Host 'Marketplace installation completed.'
}

function Fail-AndExit {
  param ([string]$Message)
  if ($script:ui) {
    Set-Status -text 'FAILED' -percent 100 -color 'Red'
    Write-Log -message $Message -color 'Red'
    Move-CursorToBottom
  }
  else {
    Write-Host "FAILED: $Message" -ForegroundColor Red
  }
  Pause
  exit 1
}
#endregion Core Functions

#region Main
try {
  try {
    Initialize-InstallTui
  }
  catch {
    $script:ui = $null
    Write-Host "TUI disabled: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host 'Falling back to standard output mode.' -ForegroundColor Yellow
  }

  Set-Status -text 'Running checks...' -percent 5 -color 'Yellow'
  if (-not (Test-PowerShellVersion)) {
    throw "PowerShell 5.1+ required. Current: $($PSVersionTable.PSVersion)"
  }
  Write-Log -message "PowerShell version check passed: $($PSVersionTable.PSVersion)" -color 'Green'

  if (-not (Test-Admin)) {
    Write-Log -message 'Warning: running as Administrator may cause issues.' -color 'Yellow'
    Move-CursorToBottom
    $Host.UI.RawUI.Flushinputbuffer()
    $choices = [System.Management.Automation.Host.ChoiceDescription[]] @(
      (New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Abort installation.'),
      (New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Continue installation.')
    )
    $choice = $Host.UI.PromptForChoice('', 'Do you want to abort the installation process?', $choices, 0)
    if ($choice -eq 0) {
      throw 'spicetify installation aborted by user.'
    }
    Write-Log -message 'Continuing installation in administrator mode.' -color 'Yellow'
  }
  else {
    Write-Log -message 'Admin check passed.'
  }

  Install-Spicetify

  Move-CursorToBottom
  $Host.UI.RawUI.Flushinputbuffer()
  $choices = [System.Management.Automation.Host.ChoiceDescription[]] @(
    (New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Install Spicetify Marketplace.'),
    (New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Do not install Spicetify Marketplace.')
  )
  $choice = $Host.UI.PromptForChoice('', 'Do you also want to install Spicetify Marketplace?', $choices, 0)
  if ($choice -eq 0) {
    Enter-StandardOutput
    Install-Marketplace
  }
  else {
    Write-Log -message 'Marketplace installation skipped.' -color 'Yellow'
    Move-CursorToBottom
  }
}
catch {
  Fail-AndExit -Message $_.Exception.Message
}
#endregion Main
