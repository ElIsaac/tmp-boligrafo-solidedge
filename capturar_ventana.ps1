# Captura la ventana COMPLETA de Solid Edge (barra de titulo, cinta, arbol y vista) a un PNG.
# Uso: powershell -ExecutionPolicy Bypass -File capturar_ventana.ps1 -Salida "C:\...\D3_clic\capturas\d3_leva.png"
param([Parameter(Mandatory=$true)][string]$Salida)

Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  public struct RECT { public int Left, Top, Right, Bottom; }
}
"@
[Win]::SetProcessDPIAware() | Out-Null

$proc = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and ($_.MainWindowTitle -like "*Solid Edge*" -or ($_.ProcessName -eq "Edge" -and $_.Path -like "*Solid Edge*")) } | Select-Object -First 1
if (-not $proc) { Write-Error "No encontre la ventana de Solid Edge abierta"; exit 1 }
$h = $proc.MainWindowHandle

[Win]::ShowWindow($h, 3) | Out-Null
[Win]::SetForegroundWindow($h) | Out-Null
Start-Sleep -Milliseconds 1200

$r = New-Object Win+RECT
[Win]::GetWindowRect($h, [ref]$r) | Out-Null
$ancho = $r.Right - $r.Left
$alto = $r.Bottom - $r.Top
$bmp = New-Object System.Drawing.Bitmap $ancho, $alto
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($r.Left, $r.Top, 0, 0, $bmp.Size)
New-Item -ItemType Directory -Force -Path (Split-Path $Salida) | Out-Null
$bmp.Save($Salida, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "OK $Salida ${ancho}x${alto}"
