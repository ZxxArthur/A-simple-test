param(
  [switch]$Install,
  [ValidateSet('winget', 'standalone', 'pip')]
  [string]$Method = 'winget'
)

$ErrorActionPreference = 'Stop'

function Get-UvCommand {
  $uvCommand = Get-Command uv -ErrorAction SilentlyContinue
  if ($uvCommand) {
    return $uvCommand
  }

  $candidatePaths = @(
    (Join-Path $env:USERPROFILE '.local\bin\uv.exe'),
    (Join-Path $env:APPDATA 'Python\Python310\Scripts\uv.exe'),
    (Join-Path $env:APPDATA 'Python\Python311\Scripts\uv.exe'),
    (Join-Path $env:APPDATA 'Python\Python312\Scripts\uv.exe')
  )

  $storePythonRoot = Join-Path $env:LOCALAPPDATA 'Packages'
  if (Test-Path $storePythonRoot) {
    $candidatePaths += Get-ChildItem $storePythonRoot -Directory -Filter 'PythonSoftwareFoundation.Python.*' -ErrorAction SilentlyContinue |
      ForEach-Object { Join-Path $_.FullName 'LocalCache\local-packages\Python310\Scripts\uv.exe' }
  }

  foreach ($candidatePath in $candidatePaths) {
    if (Test-Path $candidatePath) {
      return Get-Item $candidatePath
    }
  }
}

function Get-CommandPath($Command) {
  if ($Command.Source) {
    return $Command.Source
  }
  return $Command.FullName
}

$uvCommand = Get-UvCommand
if ($uvCommand) {
  $uvPath = Get-CommandPath $uvCommand
  Write-Host "uv found: $uvPath"
  & $uvPath --version
  exit 0
}

if (-not $Install) {
  Write-Host 'uv was not found.'
  Write-Host 'To install it for this machine, run:'
  Write-Host '  powershell -ExecutionPolicy Bypass -File .\setup-uv.ps1 -Install'
  exit 1
}

if ($Method -eq 'winget') {
  $wingetCommand = Get-Command winget -ErrorAction SilentlyContinue
  if (-not $wingetCommand) {
    throw 'winget was not found. Re-run with -Method standalone or install uv manually from https://docs.astral.sh/uv/getting-started/installation/'
  }

  & $wingetCommand.Source install --id astral-sh.uv -e
} elseif ($Method -eq 'standalone') {
  powershell -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex"
} else {
  $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
  if (-not $pythonCommand) {
    $pythonCommand = Get-Command py -ErrorAction SilentlyContinue
  }
  if (-not $pythonCommand) {
    throw 'Python was not found. Install Python first or re-run with -Method winget / -Method standalone.'
  }

  & $pythonCommand.Source -m pip install --user uv --index-url https://pypi.org/simple --timeout 20 --retries 1
}

$uvCommand = Get-UvCommand
if (-not $uvCommand) {
  Write-Host 'uv was installed, but it is not visible in the current PATH yet. Open a new terminal and run: uv --version'
  exit 0
}

$uvPath = Get-CommandPath $uvCommand
Write-Host "uv installed: $uvPath"
& $uvPath --version