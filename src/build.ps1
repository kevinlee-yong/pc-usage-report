# Build PCUsageReport.exe (.NET Framework 4.x built-in compiler, nothing to install)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$fw   = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$csc  = Join-Path $fw 'csc.exe'
$out  = Join-Path $here 'PCUsageReport.exe'
& $csc /nologo /target:exe /platform:anycpu /optimize+ /codepage:65001 `
  "/out:$out" "/win32manifest:$(Join-Path $here 'app.manifest')" `
  "/resource:$(Join-Path $here 'template.html'),template.html" `
  "/reference:$(Join-Path $fw 'System.Web.Extensions.dll')" `
  /reference:System.Core.dll `
  /reference:System.Drawing.dll `
  /reference:System.Windows.Forms.dll `
  (Join-Path $here 'Program.cs')
if ($LASTEXITCODE -ne 0) { throw "build failed" }
"built: $out"
