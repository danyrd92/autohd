# =========================================================================
# Actualizar-HDZStudio.ps1  —  Actualiza HDZ Studio desde GitHub
# =========================================================================
# Descarga SOLO los ficheros que listes en version.json (los scripts pesan KB),
# hace copia de seguridad de los actuales y los reemplaza. Si alguna HERRAMIENTA
# cambió de version, vuelve a ejecutar el instalador para refrescar solo esa.
# NUNCA toca tus ajustes/credenciales (estan en %APPDATA%).
#
# Lo usa tanto el boton "Actualizar" del programa como "Actualizar HDZ Studio.cmd".
# Compatible con Windows PowerShell 5.1 (guardado con BOM).
#
#   -Forzar : aplica aunque la version no sea mas nueva (re-instalar igual).
# =========================================================================
param([switch]$Forzar)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
function Escribe($t, $c = "Gray") { Write-Host $t -ForegroundColor $c }

# --- Carpeta del programa ---
$instalDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $instalDir
if (-not (Test-Path -LiteralPath (Join-Path $appRoot "HDZ-GUI.ps1"))) { $appRoot = $instalDir }

function Descargar($url, $destino) {
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", "HDZ-Studio-Updater")
    $wc.DownloadFile($url, $destino)
}

# Base de descarga. En condiciones normales se construye desde HDZ-update.config.json (GitHub raw).
# $env:HDZ_UPDATE_BASE permite apuntar a otra base (se usa solo en pruebas locales).
$base = "$env:HDZ_UPDATE_BASE"
$repo = ""
if (-not $base) {
    $cfgPath = Join-Path $appRoot "HDZ-update.config.json"
    if (-not (Test-Path -LiteralPath $cfgPath)) { Escribe "No encuentro HDZ-update.config.json." "Red"; return }
    $cfg = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    # Si hay "baseUrl" (servidor propio) se usa esa; si no, se construye la de GitHub raw.
    $baseCfg = "$($cfg.baseUrl)".Trim()
    if ($baseCfg) {
        $base = $baseCfg.TrimEnd("/") + "/"
    } else {
        $repo = "$($cfg.repo)"; $rama = "$($cfg.rama)"; if (-not $rama) { $rama = "main" }
        if ([string]::IsNullOrWhiteSpace($repo) -or $repo -like "*USUARIO/REPO*") {
            Escribe "Aun no has configurado el repositorio en HDZ-update.config.json (sigue en USUARIO/REPO)." "Yellow"
            Escribe "Pon ahi tu repo de GitHub (p.ej. \"repo\": \"tuusuario/hdz-studio\") y vuelve a intentarlo." "Yellow"
            return
        }
        $base = "https://raw.githubusercontent.com/$repo/$rama/"
    }
}
# Codifica cada segmento de la ruta (por si hay espacios en un nombre de fichero).
function UrlArchivo($nombre) {
    $partes = $nombre -split "[\\/]" | ForEach-Object { [System.Uri]::EscapeDataString($_) }
    return $base + ($partes -join "/")
}

Escribe ""
Escribe "  HDZ Studio — Actualizador" "White"
Escribe "  Origen: $base"

# --- Leer version.json (se descarga y se parsea; vale para http(s) y file://) ---
$vjTmp = Join-Path $env:TEMP ("hdz_vj_" + [System.Guid]::NewGuid().ToString("N") + ".json")
try {
    Descargar ($base + "version.json") $vjTmp
    $remoto = Get-Content -LiteralPath $vjTmp -Raw -Encoding UTF8 | ConvertFrom-Json
    Remove-Item -LiteralPath $vjTmp -Force -ErrorAction SilentlyContinue
}
catch { Escribe "No pude leer version.json del repo: $($_.Exception.Message)" "Red"; return }

$verRemota = "$($remoto.version)"
$verLocalFile = Join-Path $appRoot "VERSION.txt"
$verLocal = if (Test-Path -LiteralPath $verLocalFile) { (Get-Content -LiteralPath $verLocalFile -Raw).Trim() } else { "0.0.0" }
Escribe "  Version instalada: $verLocal   |   disponible: $verRemota"

$esNueva = $false
try { $esNueva = ([version]$verRemota) -gt ([version]$verLocal) } catch { $esNueva = ($verRemota -ne $verLocal) }
if (-not $esNueva -and -not $Forzar) { Escribe "  Ya estas en la ultima version." "Green"; return }

# --- Descargar los ficheros a un temporal y validarlos ---
$archivos = @($remoto.archivos)
if ($archivos.Count -eq 0) { Escribe "version.json no lista 'archivos'." "Red"; return }
$tmp = Join-Path $env:TEMP ("hdz_upd_" + [System.Guid]::NewGuid().ToString("N"))
[void][System.IO.Directory]::CreateDirectory($tmp)
Escribe ""
Escribe "  Descargando $($archivos.Count) fichero(s)…"
$bajados = @()
try {
    foreach ($a in $archivos) {
        $destino = Join-Path $tmp $a
        [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $destino))
        Descargar (UrlArchivo $a) $destino
        if ((Get-Item -LiteralPath $destino).Length -le 0) { throw "el fichero '$a' llego vacio" }
        $bajados += [pscustomobject]@{ Rel = $a; Tmp = $destino }
        Escribe "    OK  $a"
    }
} catch {
    Escribe "  [!] Fallo la descarga ($($_.Exception.Message)). No se ha cambiado nada." "Red"
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    return
}

# --- Copia de seguridad de los actuales y reemplazo ---
$backup = Join-Path $instalDir ("copia-seguridad\" + $verLocal + "_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
[void][System.IO.Directory]::CreateDirectory($backup)
Escribe ""
Escribe "  Aplicando (copia de seguridad en: $backup )…"
foreach ($b in $bajados) {
    $actual = Join-Path $appRoot $b.Rel
    if (Test-Path -LiteralPath $actual) {
        $bk = Join-Path $backup $b.Rel
        [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $bk))
        Copy-Item -LiteralPath $actual -Destination $bk -Force
    }
    $destFinal = Join-Path $appRoot $b.Rel
    [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $destFinal))
    Copy-Item -LiteralPath $b.Tmp -Destination $destFinal -Force
    Escribe "    Actualizado  $($b.Rel)" "Green"
}
Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue

# --- Herramientas: si alguna version cambió, re-ejecutar el instalador (refresca solo esa) ---
if ($remoto.herramientas) {
    $rutaVerJson = Join-Path $appRoot "bin\herramientas.json"
    $instaladas = @{}
    if (Test-Path -LiteralPath $rutaVerJson) {
        try { $j = Get-Content -LiteralPath $rutaVerJson -Raw -Encoding UTF8 | ConvertFrom-Json; foreach ($p in $j.PSObject.Properties) { $instaladas[$p.Name] = "$($p.Value)" } } catch {}
    }
    $hayCambioTool = $false
    foreach ($p in $remoto.herramientas.PSObject.Properties) {
        $obj = "$($p.Value)"
        if ($obj -ne "latest" -and "$($instaladas[$p.Name])" -ne $obj) { $hayCambioTool = $true }
    }
    if ($hayCambioTool) {
        Escribe ""
        Escribe "  Hay herramientas con nueva version: ejecutando el instalador para refrescarlas…" "Cyan"
        $inst = Join-Path $instalDir "Instalar-HDZStudio.ps1"
        if (Test-Path -LiteralPath $inst) {
            try { & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $inst | Out-Host } catch { Escribe "    No se pudo lanzar el instalador: $($_.Exception.Message)" "Yellow" }
        }
    }
}

Escribe ""
Escribe "  Actualizado a la version $verRemota." "Green"
if ($remoto.notas) { Escribe "  Novedades: $($remoto.notas)" "Gray" }

# --- Relanzar el programa ---
$vbs = Join-Path $appRoot "HDZ Studio.vbs"
if (Test-Path -LiteralPath $vbs) {
    Escribe "  Reabriendo HDZ Studio…" "Gray"
    Start-Process "wscript.exe" -ArgumentList ('"' + $vbs + '"')
} else {
    Escribe "  Abre HDZ Studio de nuevo para usar la version actualizada." "Gray"
}
