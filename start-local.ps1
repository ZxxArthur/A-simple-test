param(
  [string]$ExcelPath = '',
  [int]$Port = 3765
)

$ErrorActionPreference = 'Stop'
$nodePath = Join-Path $env:ProgramFiles 'nodejs\node.exe'
if (-not (Test-Path $nodePath)) {
  $nodeCommand = Get-Command node -ErrorAction SilentlyContinue
  if (-not $nodeCommand) {
    throw 'Node.js was not found.'
  }
  $nodePath = $nodeCommand.Source
}

$serverArgs = @((Join-Path $PSScriptRoot 'server-local.js'), '--port', $Port, '--open')
if ($ExcelPath) {
  $serverArgs += @('--excel', $ExcelPath)
}

& $nodePath @serverArgs