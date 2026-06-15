#requires -Version 7.0
# =========================================================================
# SCRIPT MAESTRO HDZ: LÍNEA DE MONTAJE AUTOMATIZADA (BATCH EDITION)
# Versión refactorizada
# =========================================================================
# NOTA: requiere PowerShell 7+ (usa ProcessStartInfo.ArgumentList, que no existe
# en Windows PowerShell 5.1). El #requires lo hace fallar limpiamente al inicio
# en lugar de reventar a mitad de proceso.

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$rutaCarpeta = (Get-Location).Path
$ErrorActionPreference = "Stop"

# Herramientas portables incluidas (ffmpeg, mkvmerge, mediainfo, dovi_tool…): si existe bin\ junto
# a este script, la anteponemos al PATH para usar SIEMPRE esas versiones probadas (no las del sistema).
$rutaBinHDZ = Join-Path $PSScriptRoot "bin"
if ((Test-Path -LiteralPath $rutaBinHDZ) -and ($env:PATH -notlike "*$rutaBinHDZ*")) { $env:PATH = "$rutaBinHDZ;$env:PATH" }

# --- CONFIGURACIÓN EXTERNA (GUI) ---
# Si la interfaz gráfica (HDZ-GUI.ps1) lanzó este script, deja en $env:HDZ_CONFIG la ruta de
# un JSON con las respuestas pre-elegidas. Cada pregunta consulta primero esa configuración
# (Get-CfgGui) y solo pregunta por consola si la clave no está definida (null = "preguntar").
# Sin la variable de entorno, el script se comporta exactamente igual que siempre.
$global:cfgGui = $null
if ($env:HDZ_CONFIG -and (Test-Path -LiteralPath $env:HDZ_CONFIG)) {
    try {
        $global:cfgGui = Get-Content -LiteralPath $env:HDZ_CONFIG -Raw -Encoding UTF8 | ConvertFrom-Json
        Write-Host "Configuración de la GUI cargada: $env:HDZ_CONFIG" -ForegroundColor DarkGray
    } catch {
        Write-Host "Aviso: no se pudo leer la configuración de la GUI ($($_.Exception.Message)). Se preguntará todo por consola." -ForegroundColor DarkYellow
    }
}
# Devuelve el valor de una clave de la configuración GUI, o $null si no está definida
# (en cuyo caso el script pregunta por consola como siempre).
function Get-CfgGui($nombre) {
    if (-not $global:cfgGui) { return $null }
    if ($global:cfgGui.PSObject.Properties.Name -notcontains $nombre) { return $null }
    $v = $global:cfgGui.$nombre
    if ($null -eq $v -or "$v" -eq "") { return $null }
    return $v
}
# Decisión forzado/completo de un sub único, POR idioma+formato (la GUI manda una sección por sub
# en SubsUnicosPorIdioma = mapa "<cod>|<Text|PGS>" -> "Forzado"/"Completo"). Respaldo: decisión global.
function Get-DecisionSubGui($cod, $fmt) {
    $mapa = Get-CfgGui "SubsUnicosPorIdioma"
    if ($mapa) {
        $clave = "$cod|$fmt"
        if ($mapa.PSObject.Properties.Name -contains $clave) { return "$($mapa.$clave)" }
    }
    return (Get-CfgGui "SubsUnicosDecision")
}

# --- PROGRESO PARA LA GUI ---
# Si la GUI lanzó este script, deja en $env:HDZ_PROGRESS la ruta de un JSON que vamos
# actualizando con el avance del montaje; la interfaz lo lee y mueve su barra de progreso.
# Sin la variable, estas funciones no hacen nada (ejecución directa por consola igual que siempre).
$global:rutaProgreso = $env:HDZ_PROGRESS
# Estado de progreso. La barra PRINCIPAL (pct/fase/archivo) refleja el avance global del lote.
# La barra SECUNDARIA (pct2/fase2) muestra el detalle de la tarea en curso (p.ej. el ensamblado
# del MKV con mkvmerge). pct2 = -1 oculta la segunda barra.
$global:gPct = 0; $global:gFase = ""; $global:gArchivo = ""; $global:gPct2 = -1; $global:gFase2 = ""
function Write-ProgresoGui([bool]$fin = $false) {
    if (-not $global:rutaProgreso) { return }
    try {
        ([ordered]@{ pct = $global:gPct; fase = $global:gFase; archivo = $global:gArchivo;
            pct2 = $global:gPct2; fase2 = $global:gFase2; fin = $fin } |
            ConvertTo-Json -Compress) | Set-Content -LiteralPath $global:rutaProgreso -Encoding UTF8
    } catch {}
}
function Set-ProgresoGui([int]$pct, [string]$fase, [string]$archivo = "") {
    if ($pct -lt 0) { $pct = 0 } elseif ($pct -gt 100) { $pct = 100 }
    $global:gPct = $pct; $global:gFase = $fase; $global:gArchivo = $archivo
    Write-ProgresoGui
}
# Segunda barra (detalle de la tarea actual). $pct2 = -1 para ocultarla.
function Set-ProgresoGui2([int]$pct2, [string]$fase2 = "") {
    if ($pct2 -gt 100) { $pct2 = 100 }
    $global:gPct2 = $pct2; $global:gFase2 = $fase2
    Write-ProgresoGui
}
function Set-ProgresoGuiFin([string]$fase) {
    $global:gPct = 100; $global:gFase = $fase; $global:gArchivo = ""; $global:gPct2 = -1; $global:gFase2 = ""
    Write-ProgresoGui $true
}
$global:idxGrupoGui = 0
$global:totalGruposGui = 0

# Resultados para la GUI: por cada torrent creado, una línea JSON (tipo/torrent/vídeo) que la GUI
# lee para ir creando pestañas de subida automáticamente. Sin la variable, no hace nada.
$global:rutaResultados = $env:HDZ_RESULTS
function Add-ResultadoGui($tipo, $torrent, $video) {
    if (-not $global:rutaResultados) { return }
    try {
        ([ordered]@{ tipo = "$tipo"; torrent = "$torrent"; video = "$video" } | ConvertTo-Json -Compress) |
            Add-Content -LiteralPath $global:rutaResultados -Encoding UTF8
    } catch {}
}

# Decisiones de idioma POR PISTA tomadas en la GUI para pistas 'und' (clave
# "<ruta>|<id>|<Audio|Sub>" -> código de idioma). La GUI lista cada pista sin idioma
# por separado, así que puede haber idiomas distintos para distintas pistas.
$global:undPistasGui = @{}
$vCfgUndPistas = Get-CfgGui "IdiomasUndPistas"
if ($vCfgUndPistas) {
    foreach ($p in @($vCfgUndPistas)) {
        if ($p.Archivo -and $null -ne $p.Id -and $p.Idioma) {
            $global:undPistasGui["$($p.Archivo)|$($p.Id)|$($p.Tipo)"] = "$($p.Idioma)"
        }
    }
}

# --- LOGGING PERSISTENTE (D3) ---
$rutaLog = Join-Path $rutaCarpeta "hdz_montaje_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
# Fichero temporal de diagnóstico: el diagnóstico extensivo se escribe aquí directamente
# (sin pasar por consola), y al terminar el script se anexa al log principal y se borra,
# de forma que el usuario acaba con UN SOLO archivo .log pero sin ver el diagnóstico en pantalla.
$global:rutaDiag = "$rutaLog.diag.tmp"
try { Start-Transcript -Path $rutaLog -Append | Out-Null } catch { Write-Host "Aviso: no se pudo iniciar transcript ($($_.Exception.Message))" -ForegroundColor DarkYellow }

# Registro global de incidencias: se acumulan los problemas encontrados durante toda la
# ejecución (fallos de mkvpropedit, archivos saltados, errores de muxeo, etc.) y se muestran
# en un resumen al final, para que ningún error pase desapercibido.
$global:incidencias = New-Object System.Collections.Generic.List[string]
function Add-Incidencia($archivo, $mensaje) {
    $linea = if ($archivo) { "[$archivo] $mensaje" } else { $mensaje }
    $global:incidencias.Add($linea)
}

# Memoria global de idiomas resueltos para subtítulos 'und' (indeterminado). El usuario elige
# el idioma de una lista y se guarda aquí con clave "<archivo>|<formato>" (formato = Text/PGS).
# Build-PistasBrutas lo aplica para asignar el código correcto a la pista.
$global:decisionesIdiomaUnd = @{}

# Idioma global elegido para TODAS las pistas 'und' del lote (audio y sub por separado).
# Se rellenan en el pre-escaneo inicial. $null = no preguntado / no aplicar.
$global:idiomaUndAudioLote = $null
$global:idiomaUndSubLote = $null
# Idiomas 'und' pre-decididos desde la GUI. El valor "und" también es válido: significa
# "dejar como indeterminado sin preguntar".
$vCfgUnd = Get-CfgGui "IdiomaUndAudio"; if ($vCfgUnd) { $global:idiomaUndAudioLote = "$vCfgUnd" }
$vCfgUnd = Get-CfgGui "IdiomaUndSub";   if ($vCfgUnd) { $global:idiomaUndSubLote   = "$vCfgUnd" }

# =========================================================================
# DICCIONARIOS Y CONSTANTES
# =========================================================================
$dictIdiomas = @{
    "spa"="Castellano"; "cas"="Castellano"; "es"="Castellano"; "es-es"="Castellano";
    "es-419"="Latino"; "es-mx"="Latino"; "es-ar"="Latino"; "es-co"="Latino"; "es-cl"="Latino"; "es-pe"="Latino"; "es-ve"="Latino"; "es-la"="Latino"; "lat"="Latino";
    "eng"="Inglés"; "en"="Inglés";
    "fra"="Francés"; "fre"="Francés"; "fr"="Francés"; "ger"="Alemán"; "deu"="Alemán"; "de"="Alemán";
    "ita"="Italiano"; "it"="Italiano"; "por"="Portugués"; "pt"="Portugués"; "jpn"="Japonés"; "ja"="Japonés";
    "chi"="Chino"; "zho"="Chino"; "zh"="Chino"; "kor"="Coreano"; "ko"="Coreano";
    "cat"="Catalán"; "ca"="Catalán"; "glg"="Gallego"; "gl"="Gallego"; "eus"="Euskera"; "baq"="Euskera"; "eu"="Euskera"
}
$idiomasComunes = @(
    @{Cod="es";     Nom="Castellano"},
    @{Cod="es-419"; Nom="Latino"},
    @{Cod="eng";    Nom="Inglés"},
    @{Cod="cat";    Nom="Catalán"},
    @{Cod="glg";    Nom="Gallego"},
    @{Cod="eus";    Nom="Euskera"},
    @{Cod="fra";    Nom="Francés"},
    @{Cod="ger";    Nom="Alemán"},
    @{Cod="ita";    Nom="Italiano"},
    @{Cod="por";    Nom="Portugués"},
    @{Cod="jpn";    Nom="Japonés"},
    @{Cod="chi";    Nom="Chino"},
    @{Cod="kor";    Nom="Coreano"}
)
$opcPlataformas = @(
    @{Nombre="Netflix";      Valor="NF"},
    @{Nombre="Amazon Prime"; Valor="AMZN"},
    @{Nombre="Disney+";      Valor="DSNP"},
    @{Nombre="Movistar+";    Valor="MVSTP"},
    @{Nombre="SkyShowtime";  Valor="SKST"},
    @{Nombre="Filmin";       Valor="FLMN"},
    @{Nombre="Apple TV+";    Valor="ATVP"},
    @{Nombre="iTunes";       Valor="iT"},
    @{Nombre="HBO Max";      Valor="HMAX"},
    @{Nombre="RTVE Play";    Valor="RTVP"},
    @{Nombre="Rakuten";      Valor="RKTN"}
)
$opcFormatosFisicos = @(
    @{Nombre="UHDFull";    Valor="UHDFull"},
    @{Nombre="BDFull";     Valor="BDFull"},
    @{Nombre="UHDRemux";   Valor="UHDRemux"},
    @{Nombre="BDRemux";    Valor="BDRemux"},
    @{Nombre="UHDRip";     Valor="UHDRip"},
    @{Nombre="BDRip";      Valor="BDRip"},
    @{Nombre="MicroHD";    Valor="MHD"},
    @{Nombre="Remastered"; Valor="Remastered"}
)

# Cache global de mediainfo (D5): clave = path absoluto, valor = objeto JSON parseado
$cacheMediainfo = @{}

# Cache de Resolve-Hibrido: la detección DV/HDR10 ejecuta ffprobe completo sobre cada archivo
# del grupo, y la función se llama varias veces por grupo (pre-escaneos + bucle principal).
# Clave = paths del grupo concatenados; valor = objeto info ya calculado.
$cacheHibrido = @{}

# =========================================================================
# FUNCIONES AUXILIARES (D1)
# =========================================================================
function Get-LanguageName($lang, $title = "") {
    $l = "$lang".ToLower()
    # Caso especial PRIMERO (antes del diccionario, que también contiene es/spa/cas y
    # respondería "Castellano" sin mirar el Title): idioma genérico español pero el
    # Title indica Latino.
    if ($l -in @("es", "spa", "cas") -and $title -match "(?i)latino|latinoam|latam|latin spanish|spanish.*latin|hispanoam") {
        return "Latino"
    }
    if ($dictIdiomas.ContainsKey($l)) { return $dictIdiomas[$l] }
    if ($lang -and $lang -ne "und") { return $lang } else { return "Und" }
}

# Devuelve el código de idioma "canónico" que vamos a escribir en el MKV final.
# Normaliza variantes regionales (es-ES, en-US, pt-BR, etc.) a un único código por idioma,
# preservando la distinción Castellano (es) vs Latino (es-419) que sí es semánticamente relevante.
function Get-LanguageCode($lang, $title = "") {
    $l = "$lang".ToLower()

    # --- ESPAÑOL ---
    # Variantes latinas → es-419 (BCP 47 estándar "Spanish, Latin America and Caribbean")
    if ($l -match "^es-(419|mx|ar|co|cl|pe|ve|la|us)$" -or $l -eq "lat") { return "es-419" }
    # Castellano (España) y código genérico → "es" a secas
    if ($l -in @("es", "spa", "cas", "es-es", "es-latn", "es-latn-es")) {
        if ($title -match "(?i)latino|latinoam|latam|latin spanish|spanish.*latin|hispanoam") {
            return "es-419"
        }
        return "es"
    }

    # --- INGLÉS ---
    if ($l -in @("en", "eng") -or $l -match "^en-(us|gb|uk|ca|au|nz|ie|za|in)$") { return "eng" }

    # --- PORTUGUÉS ---
    # Unificamos Portugal y Brasil bajo "por" (decisión del usuario)
    if ($l -in @("pt", "por") -or $l -match "^pt-(pt|br|ao|mz)$") { return "por" }

    # --- FRANCÉS ---
    if ($l -in @("fr", "fra", "fre") -or $l -match "^fr-(fr|ca|be|ch)$") { return "fre" }

    # --- ALEMÁN ---
    if ($l -in @("de", "ger", "deu") -or $l -match "^de-(de|at|ch)$") { return "ger" }

    # --- ITALIANO ---
    if ($l -in @("it", "ita") -or $l -match "^it-(it|ch)$") { return "ita" }

    # --- JAPONÉS ---
    if ($l -in @("ja", "jpn", "jp") -or $l -match "^ja-jp$") { return "jpn" }

    # --- CHINO ---
    if ($l -in @("zh", "chi", "zho", "cmn") -or $l -match "^zh-(cn|tw|hk|sg|hans|hant)$") { return "chi" }

    # --- COREANO ---
    if ($l -in @("ko", "kor") -or $l -match "^ko-kr$") { return "kor" }

    # --- LENGUAS COOFICIALES DE ESPAÑA ---
    if ($l -in @("ca", "cat") -or $l -match "^ca-(es|ad)$") { return "cat" }
    if ($l -in @("gl", "glg") -or $l -eq "gl-es") { return "glg" }
    if ($l -in @("eu", "eus", "baq") -or $l -eq "eu-es") { return "eus" }

    # Cualquier otro idioma se devuelve tal cual (en su forma original)
    return $l
}

function Get-PesoIdioma($nombre) { 
    if ($nombre -eq "Castellano") { return 1 }
    if ($nombre -match "Latino") { return 2 }
    if ($nombre -eq "Inglés") { return 3 }
    return 4 
}

function Get-PesoFormatoSub($formato) { 
    if ($formato -match "(?i)srt|subrip|utf-8|planos") { return 1 }
    return 2 
}

function Mostrar-Menu($titulo, $opciones) {
    Write-Host "`n>> $titulo" -ForegroundColor Cyan
    for ($i = 0; $i -lt $opciones.Count; $i++) { Write-Host "  [$($i + 1)] $($opciones[$i].Nombre)" }
    Write-Host "  [0] Escribir manualmente" -ForegroundColor DarkGray
    while ($true) {
        $seleccion = Read-Host "Elige una opción (0-$($opciones.Count))"
        if ($seleccion -eq "0") { return Read-Host "-> Escribe el valor" }
        if ($seleccion -match '^\d+$') {
            $n = [int]$seleccion
            if ($n -gt 0 -and $n -le $opciones.Count) { return $opciones[$n - 1].Valor }
        }
        Write-Host "   Entrada no válida. Introduce un número entre 0 y $($opciones.Count)." -ForegroundColor Yellow
    }
}

# Pregunta al usuario el idioma de una pista cuyo idioma es 'und' (indeterminado).
# Muestra una lista numerada de idiomas comunes (incluidas cooficiales de España) y
# devuelve el código elegido. El usuario elige por nombre, no necesita saber los códigos.
function Resolve-IdiomaUnd($descripcion) {
    Write-Host "`n>> $descripcion no tiene idioma definido (und). ¿Qué idioma es?" -ForegroundColor Cyan
    for ($i = 0; $i -lt $idiomasComunes.Count; $i++) {
        Write-Host "  [$($i + 1)] $($idiomasComunes[$i].Nom)"
    }
    Write-Host "  [0] Dejar como indeterminado (und)" -ForegroundColor DarkGray
    while ($true) {
        $sel = Read-Host "Elige una opción (0-$($idiomasComunes.Count))"
        if ($sel -eq "0") { return "und" }
        if ($sel -match '^\d+$') {
            $n = [int]$sel
            if ($n -ge 1 -and $n -le $idiomasComunes.Count) { return $idiomasComunes[$n - 1].Cod }
        }
        Write-Host "   Entrada no válida." -ForegroundColor Yellow
    }
}
# Alias compatible para el subtítulo (mantiene el nombre usado en otros sitios)
function Resolve-IdiomaSubUnd($descripcion) { return Resolve-IdiomaUnd "El subtítulo $descripcion" }

# Pre-escaneo del lote: detecta pistas de AUDIO y de SUBTÍTULO con idioma 'und'. Si las hay,
# pregunta UNA vez (por tipo) a qué idioma corresponden, y guarda la decisión global para que
# Build-PistasBrutas la aplique a todas las pistas 'und' de ese tipo en el lote.
function Resolve-IdiomasUndLote($grupos) {
    # Si la GUI decidió los idiomas PISTA A PISTA, no hay pregunta de lote que hacer:
    # las decisiones ya están en las memorias y lo no decidido se preguntará por pista.
    if ($global:undPistasGui.Count -gt 0) {
        Write-Host "`n[GUI] Idiomas de pistas 'und' definidos por pista en la configuración ($($global:undPistasGui.Count) pista(s))." -ForegroundColor DarkGray
        return
    }
    # Si la GUI ya decidió ambos idiomas (o "und" = dejar como está), no hay nada que
    # preguntar ni que escanear.
    if ($global:idiomaUndAudioLote -and $global:idiomaUndSubLote) {
        Write-Host "`n[GUI] Idiomas para pistas 'und' definidos en la configuración (audio: $($global:idiomaUndAudioLote), subs: $($global:idiomaUndSubLote))." -ForegroundColor DarkGray
        return
    }
    Write-Host "`nComprobando idiomas de pistas en el lote..." -ForegroundColor DarkGray
    $hayAudioUnd = $false
    $haySubUnd = $false
    $archivosAudioUnd = New-Object System.Collections.Generic.List[string]
    $archivosSubUnd = New-Object System.Collections.Generic.List[string]
    $total = ($grupos | ForEach-Object { (Resolve-Hibrido $_).OrigenesPistas.Count } | Measure-Object -Sum).Sum
    $nProc = 0
    foreach ($g in $grupos) {
        $hi = Resolve-Hibrido $g
        foreach ($ruta in $hi.OrigenesPistas) {
            $nProc++
            $nombreArchivo = Split-Path $ruta -Leaf
            Write-Progress -Activity "Analizando pistas und en el lote" -Status "$nProc/$total : $nombreArchivo" -PercentComplete (($nProc/$total)*100)
            $mi = Get-MediainfoJson $ruta
            if (-not $mi) { continue }
            foreach ($p in $mi.media.track) {
                $tipo = "$($p.'@type')"
                if ($tipo -notin @("Audio", "Text")) { continue }
                $lang = "$($p.Language)"
                $cod = if ([string]::IsNullOrWhiteSpace($lang)) { "und" } else { Get-LanguageCode $lang "$($p.Title)" }
                if ($cod -eq "und") {
                    if ($tipo -eq "Audio") {
                        $hayAudioUnd = $true
                        if (-not $archivosAudioUnd.Contains($nombreArchivo)) { $archivosAudioUnd.Add($nombreArchivo) }
                    } else {
                        $haySubUnd = $true
                        if (-not $archivosSubUnd.Contains($nombreArchivo)) { $archivosSubUnd.Add($nombreArchivo) }
                    }
                }
            }
        }
    }
    Write-Progress -Activity "Analizando pistas und en el lote" -Completed

    if ($hayAudioUnd -and -not $global:idiomaUndAudioLote) {
        Write-Host "`n>> Pistas de AUDIO sin idioma definido (und) detectadas en $($archivosAudioUnd.Count) archivo(s):" -ForegroundColor Yellow
        foreach ($a in $archivosAudioUnd) { Write-Host "   - $a" -ForegroundColor Gray }
        $c = Resolve-IdiomaUnd "Hay pistas de AUDIO en el lote que"
        if ($c -ne "und") { $global:idiomaUndAudioLote = $c }
    }
    if ($haySubUnd -and -not $global:idiomaUndSubLote) {
        Write-Host "`n>> SUBTÍTULOS sin idioma definido (und) detectados en $($archivosSubUnd.Count) archivo(s):" -ForegroundColor Yellow
        foreach ($a in $archivosSubUnd) { Write-Host "   - $a" -ForegroundColor Gray }
        $c = Resolve-IdiomaUnd "Hay SUBTÍTULOS en el lote que"
        if ($c -ne "und") { $global:idiomaUndSubLote = $c }
    }
}

function Detectar-Episodio($nombre) {
    # Detecta el episodio (o rango de episodios) del nombre. Soporta episodios múltiples en un
    # mismo archivo (S01E02-E03, S01E02-03, S01E02E03, 1x02-1x03, E02-E03...), normalizando
    # SIEMPRE al formato "S01E02-E03-E04" (con guion y la E repetida) o "E02-E03" si no hay temporada.
    # Episodios y temporadas pueden tener más de 2 dígitos; rellenamos a mínimo 2 sin recortar.

    # Helper: a partir de la posición donde acabó el primer "Eyy", captura episodios adicionales
    # encadenados (rango múltiple). Un eslabón válido SOLO si:
    #   - empieza por guion (con espacios opcionales) seguido de número, con E/EP opcional:
    #     "-E03", "- 03", "-03"
    #   - o es "E"/"Ep" pegado directamente al anterior: "E04"
    # Deliberadamente NO aceptamos " 1080" (espacio + número suelto, que sería resolución) ni la
    # "x" como separador de extras (evita confundir codecs x264/x265 con episodios).
    $capturarExtras = {
        param($resto)
        $extras = @()
        $rx = [regex]"(?i)^(?:\s*-\s*(?:E|EP)?\s*(\d{1,4})|(?:E|EP)(\d{1,4}))"
        $pos = 0
        while ($pos -lt $resto.Length) {
            $m = $rx.Match($resto.Substring($pos))
            if (-not $m.Success) { break }
            $num = if ($m.Groups[1].Value) { $m.Groups[1].Value } else { $m.Groups[2].Value }
            $extras += $num
            $pos += $m.Length
        }
        return $extras
    }

    # 1) Formato SxxEyy (+ extras)
    if ($nombre -match "(?i)S(\d{1,4})[\.\s-]*E(\d{1,4})") {
        $temp = $Matches[1].PadLeft(2,'0')
        $ep1  = $Matches[2].PadLeft(2,'0')
        $idxFin = $nombre.IndexOf($Matches[0]) + $Matches[0].Length
        $resto = $nombre.Substring($idxFin)
        $extras = & $capturarExtras $resto
        $lista = @($ep1) + @($extras | ForEach-Object { $_.PadLeft(2,'0') })
        return "S$temp" + (($lista | ForEach-Object { "E$_" }) -join "-")
    }
    # 2) Formato NxNN (+ extras). Los extras pueden venir como "-NxNN" (repiten temporada) o
    #    como "-NN"/"-ENN". Primero normalizamos los "-NxNN" quitando la temporada repetida.
    if ($nombre -match "(?i)(\d{1,4})x(\d{1,4})") {
        $temp = $Matches[1].PadLeft(2,'0')
        $ep1  = $Matches[2].PadLeft(2,'0')
        $idxFin = $nombre.IndexOf($Matches[0]) + $Matches[0].Length
        $resto = $nombre.Substring($idxFin)
        # Convertir eslabones "-NxMM" en "-MM" para que el helper capture solo el episodio.
        $resto = [regex]::Replace($resto, "(?i)(-\s*)\d{1,4}x(\d{1,4})", '${1}${2}')
        $extras = & $capturarExtras $resto
        $lista = @($ep1) + @($extras | ForEach-Object { $_.PadLeft(2,'0') })
        return "S$temp" + (($lista | ForEach-Object { "E$_" }) -join "-")
    }
    # 3) Formato E/EP/Cap NN (+ extras), sin temporada
    if ($nombre -match "(?i)(?:^|[\s_.-])(?:E|EP|Cap)[\.\s-]*(\d{1,4})") {
        $ep1 = $Matches[1].PadLeft(2,'0')
        $idxFin = $nombre.IndexOf($Matches[0]) + $Matches[0].Length
        $resto = $nombre.Substring($idxFin)
        $extras = & $capturarExtras $resto
        $lista = @($ep1) + @($extras | ForEach-Object { $_.PadLeft(2,'0') })
        return (($lista | ForEach-Object { "E$_" }) -join "-")
    }
    # 4) Respaldo: nombres que son SOLO un número (001.mp4, 243.mp4, 198_v2.mp4, 104 Subtitulat.mkv).
    $base = [System.IO.Path]::GetFileNameWithoutExtension($nombre)
    if ($base -match "^(\d{1,4})(?:[\s_.-]|$)") { return "E$($Matches[1].PadLeft(2,'0'))" }
    return $null
}

# Cache de mediainfo (D5) - una llamada por archivo, se reutiliza
function Get-MediainfoJson($ruta) {
    if (-not $cacheMediainfo.ContainsKey($ruta)) {
        try {
            $cacheMediainfo[$ruta] = mediainfo --Output=JSON $ruta | Out-String | ConvertFrom-Json
        } catch {
            Write-Host "   [!] mediainfo falló sobre $ruta : $($_.Exception.Message)" -ForegroundColor Red
            $cacheMediainfo[$ruta] = $null
        }
    }
    return $cacheMediainfo[$ruta]
}

# Cachea el JSON de mkvmerge -J por path (igual que mediainfo).
# mkvmerge es la fuente fiable de IDs de pista para muxear (mediainfo a veces
# numera con huecos y no coincide). Lo usamos para corregir los IDs.
$cacheMkvmergeJ = @{}
function Get-MkvmergeJson($ruta) {
    if (-not $cacheMkvmergeJ.ContainsKey($ruta)) {
        try {
            $raw = & mkvmerge -J $ruta 2>$null | Out-String
            $cacheMkvmergeJ[$ruta] = $raw | ConvertFrom-Json
        } catch {
            Write-Host "   [!] mkvmerge -J falló sobre $ruta : $($_.Exception.Message)" -ForegroundColor Red
            $cacheMkvmergeJ[$ruta] = $null
        }
    }
    return $cacheMkvmergeJ[$ruta]
}

# Devuelve un hashtable que mapea "tipo+orden" -> ID real de mkvmerge.
# Ej: "Video+1" -> 0, "Audio+1" -> 1, "Audio+2" -> 2, ..., "Text+1" -> 6, "Text+2" -> 7
# Sirve para corregir los IDs que mediainfo reporta con huecos (caso real: Fight Club
# donde mediainfo numera los PGS como 7, 9, 11, 13... y mkvmerge espera 6, 7, 8, 9...).
function Get-MkvmergeTrackIdMap($ruta) {
    $map = @{}
    $json = Get-MkvmergeJson $ruta
    if (-not $json -or -not $json.tracks) { return $map }
    $contadores = @{ video = 0; audio = 0; subtitles = 0 }
    foreach ($t in $json.tracks) {
        $tipo = "$($t.type)".ToLower()
        if (-not $contadores.ContainsKey($tipo)) { continue }
        $contadores[$tipo]++
        $tipoClave = switch ($tipo) {
            "video"     { "Video" }
            "audio"     { "Audio" }
            "subtitles" { "Text" }
            default     { $tipo }
        }
        $map["${tipoClave}+$($contadores[$tipo])"] = [int]$t.id
    }
    return $map
}

# ============================================================================
# DIAGNÓSTICO EXTENSIVO AL LOG
# ============================================================================
# Estas funciones vuelcan información detallada a un fichero de diagnóstico separado
# ($global:rutaDiag). NO pasan por consola ni por Write-Verbose, así que el usuario no
# las ve en pantalla mientras el script trabaja. Al terminar, el contenido se anexa al
# log principal para que todo quede en un único archivo.

function Write-DiagRaw($texto) {
    # Escritura directa al fichero de diagnóstico. Silenciosa (no toca la consola).
    try { Add-Content -LiteralPath $global:rutaDiag -Value $texto -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# Marca de tiempo para medir cuánto tarda cada paso. Se escribe SOLO al fichero de diagnóstico,
# nunca a la consola. Formato: "[HH:mm:ss.fff] PASO: <descripción>".
function Write-DiagPaso($descripcion) {
    try {
        $ts = (Get-Date).ToString("HH:mm:ss.fff")
        Add-Content -LiteralPath $global:rutaDiag -Value "[$ts] PASO: $descripcion" -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

function Write-DiagSeccion($titulo) {
    Write-DiagRaw ""
    Write-DiagRaw ("=" * 78)
    Write-DiagRaw "[DIAG] $titulo"
    Write-DiagRaw ("=" * 78)
}

function Write-DiagLinea($texto) {
    $ts = (Get-Date).ToString("HH:mm:ss.fff")
    Write-DiagRaw "[$ts] [DIAG] $texto"
}

function Write-DiagBloque($texto) {
    # Para bloques multilínea (JSON dumps, listas de pistas, etc.)
    foreach ($l in ($texto -split "`r?`n")) {
        Write-DiagRaw "[DIAG]   $l"
    }
}

# Vuelca al log el entorno completo del sistema (versiones de herramientas, locale, etc.).
# Útil para diagnosticar diferencias entre máquinas (caso típico: mismo script, distintos
# resultados porque la versión de mediainfo difiere entre tu PC y el de un amigo).
function Write-DiagEntorno {
    Write-DiagSeccion "ENTORNO DEL SISTEMA"
    try { Write-DiagLinea "PowerShell: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" } catch {}
    try { Write-DiagLinea "OS: $($PSVersionTable.OS)" } catch {}
    try { Write-DiagLinea "Locale: $((Get-Culture).Name) / Charset: $([Console]::OutputEncoding.WebName)" } catch {}
    try { Write-DiagLinea "Ruta carpeta: $rutaCarpeta" } catch {}
    try {
        $espacio = [System.IO.DriveInfo]::new((Split-Path $rutaCarpeta -Qualifier))
        Write-DiagLinea ("Espacio libre en unidad {0}: {1:N2} GB de {2:N2} GB" -f $espacio.Name, ($espacio.AvailableFreeSpace / 1GB), ($espacio.TotalSize / 1GB))
    } catch {}

    Write-DiagLinea "--- Versiones de herramientas externas ---"
    foreach ($cmd in @("mediainfo", "ffmpeg", "ffprobe", "mkvmerge", "mkvextract", "mkvpropedit", "dovi_tool")) {
        try {
            $cmdInfo = Get-Command $cmd -ErrorAction SilentlyContinue
            if ($cmdInfo) {
                $ruta = $cmdInfo.Source
                $verLine = ""
                try { $verLine = (& $cmd --version 2>&1 | Select-Object -First 1) -as [string] } catch {}
                Write-DiagLinea "  $cmd : $verLine [$ruta]"
            } else {
                Write-DiagLinea "  $cmd : NO ENCONTRADO en PATH"
            }
        } catch {
            Write-DiagLinea "  $cmd : error al consultar versión"
        }
    }
}

# Vuelca al log toda la información disponible sobre un archivo MKV: tamaño, output de
# mkvmerge -i, output de mkvmerge -J (resumido), output de mediainfo JSON completo,
# y output de ffprobe por streams. Es la radiografía completa del archivo de entrada.
function Write-DiagArchivoEntrada($rutaMkv, $etiqueta = "ENTRADA") {
    Write-DiagSeccion "ARCHIVO DE $etiqueta : $(Split-Path $rutaMkv -Leaf)"
    if (-not (Test-Path -LiteralPath $rutaMkv)) {
        Write-DiagLinea "  ERROR: el archivo no existe en disco"
        return
    }
    $fi = Get-Item -LiteralPath $rutaMkv
    Write-DiagLinea "Path: $rutaMkv"
    Write-DiagLinea ("Tamaño: {0:N0} bytes ({1:N2} GB)" -f $fi.Length, ($fi.Length / 1GB))
    Write-DiagLinea "Última modificación: $($fi.LastWriteTime)"

    Write-DiagLinea "--- mkvmerge -i ---"
    try {
        $mi = & mkvmerge -i $rutaMkv 2>&1 | Out-String
        Write-DiagBloque $mi
    } catch { Write-DiagLinea "  ERROR ejecutando mkvmerge -i: $($_.Exception.Message)" }

    Write-DiagLinea "--- mkvmerge -J (resumen de pistas) ---"
    $jMkv = Get-MkvmergeJson $rutaMkv
    if ($jMkv -and $jMkv.tracks) {
        foreach ($t in $jMkv.tracks) {
            $props = $t.properties | ConvertTo-Json -Compress -Depth 4 -ErrorAction SilentlyContinue
            Write-DiagLinea "  Track id=$($t.id) type=$($t.type) codec=$($t.codec)"
            if ($props) { Write-DiagLinea "    props: $props" }
        }
    } else {
        Write-DiagLinea "  (mkvmerge -J no devolvió tracks)"
    }

    Write-DiagLinea "--- Mapping ID mkvmerge por (tipo+orden) ---"
    $idMap = Get-MkvmergeTrackIdMap $rutaMkv
    foreach ($k in ($idMap.Keys | Sort-Object)) {
        Write-DiagLinea "  $k -> $($idMap[$k])"
    }

    Write-DiagLinea "--- mediainfo --Output=JSON (resumen por pista) ---"
    $jMi = Get-MediainfoJson $rutaMkv
    if ($jMi -and $jMi.media -and $jMi.media.track) {
        foreach ($pista in $jMi.media.track) {
            $tipo = "$($pista.'@type')"
            if ($tipo -in @('General','Menu')) { continue }
            $resumen = "Pista mediainfo ID=$($pista.ID) tipo=$tipo Format=$($pista.Format)"
            if ($tipo -eq "Video") {
                $resumen += " Width=$($pista.Width) Height=$($pista.Height) BitDepth=$($pista.BitDepth)"
                $resumen += " colour_primaries=$($pista.colour_primaries) transfer=$($pista.transfer_characteristics)"
                $resumen += " HDR_Format=$($pista.HDR_Format)"
                $resumen += " HDR_Format_Compatibility=$($pista.HDR_Format_Compatibility)"
                $resumen += " HDR_Format_String=$($pista.HDR_Format_String)"
            }
            if ($tipo -in @('Audio','Text')) {
                $resumen += " Lang=$($pista.Language) Title=`"$($pista.Title)`""
                $resumen += " Profile=$($pista.Format_Profile) Comm=`"$($pista.Format_Commercial_IfAny)`""
                $resumen += " Channels=$($pista.Channels) Forced=$($pista.Forced) Default=$($pista.Default)"
                $resumen += " StreamSize=$($pista.StreamSize)"
            }
            Write-DiagLinea $resumen
        }
    } else {
        Write-DiagLinea "  (mediainfo no devolvió tracks)"
    }

    Write-DiagLinea "--- ffprobe -show_streams (audio/sub) ---"
    try {
        $fp = ffprobe -v error -show_entries "stream=index,codec_name,codec_type:stream_tags=language,title" -of csv=p=1 $rutaMkv 2>&1 | Out-String
        Write-DiagBloque $fp
    } catch { Write-DiagLinea "  ERROR ejecutando ffprobe: $($_.Exception.Message)" }
}

# Vuelca al log la configuración global del proyecto seleccionada por el usuario.
function Write-DiagConfiguracion {
    Write-DiagSeccion "CONFIGURACIÓN ELEGIDA"
    try { Write-DiagLinea "modoLote: $modoLote" } catch {}
    try { Write-DiagLinea "borrarOriginales: $borrarOriginales" } catch {}
    try { Write-DiagLinea "modoConversionDTS: $modoConversionDTS" } catch {}
    try { Write-DiagLinea "numCapturas: $numCapturas" } catch {}
    try {
        if ($datosProyectoGlobal) {
            Write-DiagLinea "datosProyectoGlobal (modo homogéneo):"
            Write-DiagBloque ($datosProyectoGlobal | ConvertTo-Json -Depth 3)
        }
    } catch {}
    try {
        if ($decisionesSubUnico -and $decisionesSubUnico.Keys.Count -gt 0) {
            Write-DiagLinea "decisionesSubUnico:"
            foreach ($k in ($decisionesSubUnico.Keys | Sort-Object)) {
                Write-DiagLinea "  $k -> $($decisionesSubUnico[$k])"
            }
        }
    } catch {}
}

# Vuelca al log el contenido completo de la carpeta de trabajo. Útil para detectar
# archivos externos (SRT, AC3, SUP, etc.) que están condicionando el comportamiento.
function Write-DiagCarpeta($etiqueta = "CARPETA") {
    Write-DiagSeccion "${etiqueta}: $rutaCarpeta"
    try {
        $items = Get-ChildItem -LiteralPath $rutaCarpeta -ErrorAction SilentlyContinue | Sort-Object Name
        foreach ($it in $items) {
            $tam = if ($it.PSIsContainer) { "<DIR>" } else { "{0,15:N0}" -f $it.Length }
            Write-DiagLinea "  $tam  $($it.Name)"
        }
    } catch { Write-DiagLinea "  ERROR listando carpeta: $($_.Exception.Message)" }
}

# Vuelca al log el detalle de las pistas que el script va a muxear, después de aplicar
# resolución de idiomas, deduplicación, ordenación y detección de forzados.
function Write-DiagPistasFinal($audiosOrdenados, $subsOrdenados, $idVideo, $altoVideo, $codecVideo, $hdrVideo) {
    Write-DiagSeccion "PISTAS A MUXEAR (después de Build-PistasBrutas + Format-Audios/Subs)"
    Write-DiagLinea "Video: idVideo=$idVideo altoVideo=$altoVideo codecVideo=$codecVideo hdrVideo='$hdrVideo'"
    Write-DiagLinea "Audios ($($audiosOrdenados.Count)):"
    for ($i = 0; $i -lt $audiosOrdenados.Count; $i++) {
        $a = $audiosOrdenados[$i]
        $linea = "  #$i Origen=$($a.Origen) ID=$($a.ID) ArchivoFuente=$(if ($a.ArchivoFuenteMkv) { Split-Path $a.ArchivoFuenteMkv -Leaf } else { $a.ID })"
        $linea += " CodLang=$($a.CodLang) NomLang=$($a.NomLang) Format=$($a.Format) Profile=$($a.Profile)"
        $linea += " Comm=`"$($a.Comm)`" Chan=$($a.Chan) Pts=$($a.Pts) PesoFam=$($a.PesoFamilia) NomFinal=`"$($a.NomFinal)`""
        Write-DiagLinea $linea
    }
    Write-DiagLinea "Subs ($($subsOrdenados.Count)):"
    for ($i = 0; $i -lt $subsOrdenados.Count; $i++) {
        $s = $subsOrdenados[$i]
        $linea = "  #$i Origen=$($s.Origen) ID=$($s.ID) ArchivoFuente=$(if ($s.ArchivoFuenteMkv) { Split-Path $s.ArchivoFuenteMkv -Leaf } else { $s.ID })"
        $linea += " CodLang=$($s.CodLang) NomLang=$($s.NomLang) Format=$($s.Format) Size=$($s.Size)"
        $linea += " IsForced=$($s.IsForced) MediainfoForcedFlag=$($s.MediainfoForcedFlag) NomFinal=`"$($s.NomFinal)`""
        Write-DiagLinea $linea
    }
}

# Vuelca al log el comando completo que se va a pasar a mkvmerge.
function Write-DiagComandoMkvmerge($argsMkv) {
    Write-DiagSeccion "COMANDO MKVMERGE"
    Write-DiagLinea "mkvmerge $($argsMkv -join ' ')"
}

# Vuelca al log el resultado final tras el muxeo: lo que mediainfo dice del archivo
# resultante (para verificar que coincide con lo esperado).
function Write-DiagArchivoSalida($rutaSalida, $nombreFinal) {
    Write-DiagSeccion "ARCHIVO FINAL"
    Write-DiagLinea "Path temporal muxeado: $rutaSalida"
    Write-DiagLinea "Nombre final calculado: $nombreFinal"
    if (-not (Test-Path -LiteralPath $rutaSalida)) {
        Write-DiagLinea "  ATENCIÓN: el archivo final NO EXISTE en disco"
        return
    }
    $fi = Get-Item -LiteralPath $rutaSalida
    Write-DiagLinea ("Tamaño: {0:N0} bytes ({1:N2} GB)" -f $fi.Length, ($fi.Length / 1GB))

    # mediainfo del archivo final (invalidando cache para forzar re-lectura)
    $cacheMediainfo.Remove($rutaSalida) | Out-Null
    $jMi = Get-MediainfoJson $rutaSalida
    if ($jMi -and $jMi.media -and $jMi.media.track) {
        foreach ($pista in $jMi.media.track) {
            $tipo = "$($pista.'@type')"
            if ($tipo -in @('General','Menu')) { continue }
            $resumen = "Pista final ID=$($pista.ID) tipo=$tipo Format=$($pista.Format)"
            if ($tipo -eq "Video") {
                $resumen += " HDR_Format=$($pista.HDR_Format)"
                $resumen += " colour_primaries=$($pista.colour_primaries) transfer=$($pista.transfer_characteristics) BitDepth=$($pista.BitDepth)"
            }
            if ($tipo -in @('Audio','Text')) {
                $resumen += " Lang=$($pista.Language) Title=`"$($pista.Title)`" Forced=$($pista.Forced) Default=$($pista.Default)"
            }
            Write-DiagLinea $resumen
        }
    }
}

function Test-Dependencias($necesarias) {
    $faltan = $false
    foreach ($h in $necesarias) {
        if (-not (Get-Command $h -ErrorAction SilentlyContinue)) { 
            Write-Host "Falta: $h" -ForegroundColor Red
            $faltan = $true 
        }
    }
    return (-not $faltan)
}

# Detecta si un archivo de vídeo trae al menos una pista PGS (para preguntar solo cuando aplica)
function Test-TienePGS($rutaVideo) {
    try {
        $resultado = ffprobe -v error -select_streams s -show_entries stream=codec_name -of csv=p=0 $rutaVideo 2>$null
        if (-not $resultado) { return $false }
        return ($resultado -match "(?i)pgs|hdmv_pgs")
    } catch {
        return $false
    }
}

# Detecta el fps real del vídeo (Bug 5) y devuelve la cadena tipo "24000/1001" o "25/1"
function Get-FpsReal($rutaVideo) {
    try {
        $fps = ffprobe -v error -select_streams v:0 -show_entries stream=r_frame_rate -of default=nw=1:nk=1 $rutaVideo 2>$null
        if ($fps -and $fps -match "^\d+/\d+$") { return $fps.Trim() }
    } catch {}
    return $null
}

# Renombra evitando colisiones (Bug 10): si existe el destino, añade sufijo incremental.
function Rename-EvitarColision($pathOrigen, $nombreDeseado) {
    $dir = Split-Path $pathOrigen -Parent
    $destino = Join-Path $dir $nombreDeseado
    if (-not (Test-Path -LiteralPath $destino)) {
        Rename-Item -LiteralPath $pathOrigen -NewName $nombreDeseado
        return $destino
    }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($nombreDeseado)
    $ext = [System.IO.Path]::GetExtension($nombreDeseado)
    $i = 2
    while ($true) {
        $candidato = "$base ($i)$ext"
        $ruta = Join-Path $dir $candidato
        if (-not (Test-Path -LiteralPath $ruta)) {
            Write-Host "   [!] Ya existe '$nombreDeseado'. Renombrando a '$candidato'." -ForegroundColor Yellow
            Rename-Item -LiteralPath $pathOrigen -NewName $candidato
            return $ruta
        }
        $i++
    }
}

# Wrapper para mkvpropedit que captura errores en lugar de silenciarlos (Bug 9)
function Invoke-Mkvpropedit($argumentos, $contexto) {
    # Reintentos ante bloqueos transitorios del archivo (código 2). Causas habituales:
    # antivirus escaneando el MKV recién creado, indexador de Windows, escáneres de medios
    # (Jellyfin), o latencia al liberar el handle en unidades de red/USB.
    $maxIntentos = 4
    $esperaSeg = 2
    for ($intento = 1; $intento -le $maxIntentos; $intento++) {
        $salida = & mkvpropedit @argumentos 2>&1
        if ($LASTEXITCODE -eq 0) { return $true }

        $textoSalida = "$($salida -join ' | ')"
        # Solo reintentamos si parece un bloqueo (no se pudo abrir para escritura / bloqueado).
        $esBloqueo = $textoSalida -match "(?i)no se pudo abrir|bloquead|locked|could not be opened|access|permiso"
        if ($intento -lt $maxIntentos -and $esBloqueo) {
            Write-Host "   [!] mkvpropedit no pudo escribir ($contexto), intento $intento/$maxIntentos. Reintentando en ${esperaSeg}s..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $esperaSeg
            continue
        }
        # Fallo definitivo (o no es un bloqueo): reportamos y salimos.
        Write-Host "   [!] mkvpropedit falló ($contexto). Código: $LASTEXITCODE" -ForegroundColor Red
        Write-Host "       $textoSalida" -ForegroundColor DarkRed
        Add-Incidencia $global:archivoEnCurso "mkvpropedit falló ($contexto) tras $maxIntentos intento(s). La pista pudo quedar sin renombrar o sin banderas."
        return $false
    }
    return $false
}

# Extrae una pista de subtítulo interna a un fichero temporal, mide su tamaño en disco
# y borra el temporal. Se usa cuando mediainfo no reporta StreamSize y necesitamos
# distinguir forzados de completos comparando pesos reales.
# Devuelve el tamaño en bytes, o 0 si la extracción falla.
function Measure-SubEnDisco($rutaMkv, $idPistaMkvmerge) {
    # Vía rápida y sin tocar el archivo: leer 'tag_number_of_bytes' del JSON de mkvmerge -J,
    # que ya está cacheado (Get-MkvmergeJson). Esto evita extraer la pista entera, que en archivos
    # grandes en red (UHD AV1 de varios GB sobre HDD remoto) tarda varios minutos.
    # IMPORTANTE (unidades): el valor devuelto se compara contra StreamSize de mediainfo (bytes),
    # así que SIEMPRE intentamos devolver bytes (tag o extracción real). num_index_entries NO son
    # bytes y solo se usa como último recurso si todo lo demás falla; mezclarlo con bytes haría
    # absurda la heurística del 35% (forzado vs completo).
    $entradasIndice = 0
    try {
        $mj = Get-MkvmergeJson $rutaMkv
        if ($mj -and $mj.tracks) {
            foreach ($t in $mj.tracks) {
                if ($t.id -eq [int]$idPistaMkvmerge) {
                    if ($t.properties -and $t.properties.tag_number_of_bytes) {
                        return [long]$t.properties.tag_number_of_bytes
                    }
                    if ($t.properties -and $t.properties.num_index_entries) {
                        $entradasIndice = [long]$t.properties.num_index_entries
                    }
                    break
                }
            }
        }
    } catch {}

    # Sin tag de bytes: extracción real a disco (bytes de verdad). Lento en red, pero es la
    # única medida comparable con los StreamSize del resto del grupo.
    $tempPath = Join-Path $env:TEMP "hdz_subsize_$([guid]::NewGuid().ToString('N')).tmp"
    try {
        $salida = & mkvextract tracks $rutaMkv "${idPistaMkvmerge}:$tempPath" 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tempPath)) {
            & ffmpeg -i $rutaMkv -map "0:$idPistaMkvmerge" -c copy $tempPath -y -v error 2>&1 | Out-Null
        }
        if (Test-Path -LiteralPath $tempPath) {
            $tam = (Get-Item -LiteralPath $tempPath).Length
            return [long]$tam
        }
    } catch {
    } finally {
        Remove-Item -LiteralPath $tempPath -ErrorAction SilentlyContinue
    }

    # Último recurso: num_index_entries como proxy (unidad distinta, solo mejor que 0).
    if ($entradasIndice -gt 0) {
        Write-DiagLinea "Measure-SubEnDisco: usando num_index_entries=$entradasIndice como proxy (no bytes) para pista $idPistaMkvmerge de $(Split-Path $rutaMkv -Leaf)"
        return $entradasIndice
    }
    return 0
}

# Determina el tipo de pista de un archivo externo (Bug 7) con detección más fiable.
# Las extensiones de subtítulo y de audio son inequívocas, así que las usamos como fuente de verdad
# y mediainfo solo como confirmación/fallback.
function Get-TipoPistaExterno($archivo, $pistaMediainfo) {
    $ext = $archivo.Extension.ToLower().TrimStart('.')
    $extsAudioInequivoco = @("ac3", "eac3", "dts", "dtshd", "truehd", "flac", "wav", "mka", "m4a", "aac", "opus")
    $extsSubInequivoco   = @("srt", "ass", "ssa", "sup", "pgs", "idx", "sub", "vtt")
    if ($extsAudioInequivoco -contains $ext) { return "Audio" }
    if ($extsSubInequivoco   -contains $ext) { return "Sub" }
    if ($pistaMediainfo) {
        if ($pistaMediainfo.'@type' -eq "Audio") { return "Audio" }
        if ($pistaMediainfo.'@type' -eq "Text")  { return "Sub" }
    }
    return $null   # no clasificable → el caller debe ignorarlo
}

# Pre-escaneo: detecta los grupos (idioma + formato) que tienen UN ÚNICO sub interno
# en un archivo (o pareja híbrida) sin señal positiva de forzado.
# Devuelve una lista de PSCustomObject con: CodLang, NomLang, FmtSimple, Size, RutaArchivoFuente, IdPista.
# El criterio "1 solo sub tras todo lo conocido" replica exactamente la lógica de Build-PistasBrutas:
# - Si hay 2 subs del mismo (idioma+formato) y se distinguen por tamaño (un Forzado y un Completo) → no hay duda.
# - Si hay 2 subs idénticos en tamaño → se reducirán a 1 tras la deduplicación → AHÍ sí preguntamos.
# - Si solo hay 1 sub de entrada → preguntamos.
# La función toma 1 ó 2 paths (lista de archivos del par híbrido).
function Get-SubsUnicosArchivo($rutasArchivosMkv, $incluirExternos = $false, $prefijoArchivo = "") {
    $resultado = @()
    $listaArchivos = @($rutasArchivosMkv)

    # Recolectar todos los subs internos con su tamaño "real" (midiendo en disco si hace falta).
    $todosSubs = @()
    foreach ($ruta in $listaArchivos) {
        $miJson = Get-MediainfoJson $ruta
        if (-not $miJson) { continue }
        $idMapL = Get-MkvmergeTrackIdMap $ruta
        $contTextL = 0
        foreach ($pista in $miJson.media.track) {
            if ($pista.'@type' -ne "Text") { continue }
            $contTextL++
            $claveMap = "Text+$contTextL"
            $idReal = if ($idMapL.ContainsKey($claveMap)) { $idMapL[$claveMap] } else { $pista.ID - 1 }
            $langCode = "$($pista.Language)"
            if ([string]::IsNullOrWhiteSpace($langCode)) { $langCode = "und" }
            $tituloPista = "$($pista.Title)"
            $langCode = Get-LanguageCode $langCode $tituloPista
            # Decisión por pista de la GUI: resuelve el idioma 'und' antes de clasificar
            # el sub (así los grupos y las preguntas usan ya el idioma correcto).
            if ($langCode -eq "und" -and $global:undPistasGui.ContainsKey("$ruta|$idReal|Sub")) {
                $langCode = Get-LanguageCode $global:undPistasGui["$ruta|$idReal|Sub"] ""
            }

            # Señal positiva de forzado: flag mediainfo o título con "forced"
            $tieneSenalForzado = $false
            if ($pista.PSObject.Properties.Name -contains "Forced" -and "$($pista.Forced)" -match "(?i)yes|1|true") { $tieneSenalForzado = $true }
            if ($tituloPista -match "(?i)(forced|forzado|forzados)") { $tieneSenalForzado = $true }

            # Señal positiva de COMPLETO: título con "completo"/"completos"/"complete"/"full".
            # Si el sub viene etiquetado explícitamente como completo, no hay ambigüedad → no preguntar.
            $tieneSenalCompleto = $false
            if ($tituloPista -match "(?i)(completos?|complete|full)") { $tieneSenalCompleto = $true }

            $tam = if ($pista.StreamSize) { [long]$pista.StreamSize } else { 0 }
            $fmtSimple = if ("$($pista.Format)" -match "(?i)srt|subrip|ass|ssa|utf-?8|text") { "Text" } else { "PGS" }

            $todosSubs += [PSCustomObject]@{
                CodLang = $langCode
                NomLang = (Get-LanguageName $langCode $tituloPista)
                FmtSimple = $fmtSimple
                Size = $tam
                IdPista = $idReal
                RutaArchivoFuente = $ruta
                EsExterno = $false
                NombreExterno = ""
                TieneSenalForzado = $tieneSenalForzado
                TieneSenalCompleto = $tieneSenalCompleto
            }
        }
    }

    # Subs EXTERNOS de la carpeta (SRT/ASS/SSA sueltos): se incluyen solo si se pide explícitamente
    # (típicamente justo antes de muxear, cuando ya pueden existir los SRT convertidos o los que
    # el usuario metió a mano). Identificamos el idioma por el nombre del archivo.
    if ($incluirExternos) {
        $subsExt = Get-ChildItem -LiteralPath $rutaCarpeta -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match "(?i)\.(srt|ass|ssa)$" }
        foreach ($ext in $subsExt) {
            # En modo lote (con prefijo), solo consideramos los externos de ESTE archivo.
            if ($prefijoArchivo) {
                $prefEsc = [regex]::Escape($prefijoArchivo)
                # SRT convertidos de PGS de este archivo, o cualquier SRT que empiece por su nombre.
                if ($ext.Name -notmatch "(?i)^$prefEsc") { continue }
            }
            # Idioma: lo buscamos en el nombre del archivo (es, spa, eng, en, fre...).
            $codNombre = "und"
            if ($ext.BaseName -match "(?i)_sub_\d+_([a-z]{2,3})$") { $codNombre = Get-LanguageCode $Matches[1] "" }
            elseif ($ext.BaseName -match "(?i)[\._\s-](es-419|lat|cas|spa|es|eng|en|fre|fr|ger|de|ita|it|por|pt|jpn|ja|chi|zh|kor|ko)(?:[\._\s-]|$)") { $codNombre = Get-LanguageCode $Matches[1] "" }
            $fmtSimple = "Text"
            # Señal de forzado en el nombre del archivo externo
            $senal = ($ext.Name -match "(?i)(forced|forzado|forzados)")
            $senalComp = ($ext.Name -match "(?i)(completos?|complete|full)")
            $todosSubs += [PSCustomObject]@{
                CodLang = $codNombre
                NomLang = (Get-LanguageName $codNombre "")
                FmtSimple = $fmtSimple
                Size = [long]$ext.Length
                IdPista = -1
                RutaArchivoFuente = $ext.FullName
                EsExterno = $true
                NombreExterno = $ext.Name
                TieneSenalForzado = $senal
                TieneSenalCompleto = $senalComp
            }
        }
    }

    if ($todosSubs.Count -eq 0) { return $resultado }

    # IMPORTANTE (rendimiento): la medición de tamaño en disco (Measure-SubEnDisco) puede caer en
    # mkvextract, que es MUY lento en archivos UHD (abre y demuxea varios GB por pista). Por eso NO
    # medimos todos los subs por adelantado: lo hacemos de forma PEREZOSA, solo en los grupos que de
    # verdad necesitan distinguir forzado/completo por tamaño (varios subs del mismo idioma/formato y
    # sin ninguna señal en el título). Los subs únicos o etiquetados ("forzado"/"completo") nunca se
    # miden. (Antes se medían TODOS → 48 extracciones sobre 4K = ~media hora por temporada.)

    # Agrupar por (idioma + formato_simplificado)
    $grupos = @{}
    foreach ($s in $todosSubs) {
        $k = "$($s.CodLang)|$($s.FmtSimple)"
        if (-not $grupos.ContainsKey($k)) { $grupos[$k] = @() }
        $grupos[$k] += $s
    }

    # Para cada grupo:
    # - Si en el grupo hay AL MENOS UN sub con señal positiva de forzado → no hay ambigüedad: nada que preguntar.
    # - Si el grupo tiene un solo sub (caso típico no-híbrido) → CANDIDATO a preguntar.
    # - Si tiene varios y se distinguen por tamaño (max > 0 y el menor ≤ 35% del max) → la heurística los separa → nada que preguntar.
    # - Si tiene varios pero TODOS pesan parecido (el menor > 35% del max, es decir, se considerarán duplicados
    #   tras la deduplicación) → CANDIDATO a preguntar, porque tras dedup queda 1 solo.
    foreach ($k in $grupos.Keys) {
        $lista = $grupos[$k]
        $haySenal = ($lista | Where-Object { $_.TieneSenalForzado }).Count -gt 0
        if ($haySenal) { continue }
        # Si TODOS los subs del grupo están explícitamente marcados como completos en el título,
        # no hay ambigüedad: ninguno es forzado → nada que preguntar.
        $todosCompletos = ($lista.Count -gt 0) -and (($lista | Where-Object { $_.TieneSenalCompleto }).Count -eq $lista.Count)
        if ($todosCompletos) { continue }
        if ($lista.Count -eq 1) {
            $resultado += $lista[0]
            continue
        }
        # Varios subs en el grupo: ver si la heurística de tamaño los distingue.
        # AQUÍ (y solo aquí) medimos en disco los que no traigan tamaño de mediainfo: es el único
        # caso en que el tamaño se usa de verdad, así evitamos extraer subs innecesariamente.
        foreach ($s in $lista) {
            if ($s.Size -le 0 -and -not $s.EsExterno) { $s.Size = Measure-SubEnDisco $s.RutaArchivoFuente $s.IdPista }
        }
        $maxS = ($lista | Measure-Object Size -Maximum).Maximum
        $minS = ($lista | Measure-Object Size -Minimum).Minimum
        if ($maxS -gt 0 -and $minS -le ($maxS * 0.35)) { continue }  # heurística los separa: no preguntar
        # Tamaños similares → tras dedup quedará 1 → CANDIDATO. Devolvemos el primero (representante del grupo).
        $resultado += $lista[0]
    }
    return $resultado
}

# Devuelve la lista de idiomas DISTINTOS de subtítulos presentes en un archivo o par híbrido.
# Cada entrada es un PSCustomObject con CodLang (código canónico) y NomLang (nombre legible).
# Castellano (es) y Latino (es-419) cuentan como distintos.
function Get-IdiomasSubsArchivo($rutasArchivosMkv) {
    $vistos = @{}
    foreach ($ruta in @($rutasArchivosMkv)) {
        $mi = Get-MediainfoJson $ruta
        if (-not $mi) { continue }
        foreach ($p in $mi.media.track) {
            if ($p.'@type' -ne "Text") { continue }
            $lang = "$($p.Language)"; if ([string]::IsNullOrWhiteSpace($lang)) { $lang = "und" }
            $titulo = "$($p.Title)"
            $cod = Get-LanguageCode $lang $titulo
            if (-not $vistos.ContainsKey($cod)) {
                $vistos[$cod] = [PSCustomObject]@{ CodLang = $cod; NomLang = (Get-LanguageName $cod $titulo) }
            }
        }
    }
    return @($vistos.Values)
}

# Pre-escaneo del lote (modo HOMOGÉNEO): si algún archivo tiene MÁS DE 3 idiomas distintos de
# subtítulos, pregunta qué hacer. Devuelve una decisión global que se aplicará al filtrar:
#   - $null  -> no filtrar (mantener todos)
#   - @("__CAST_ENG__")  -> mantener solo castellano (es) e inglés (eng)
#   - @("es","eng","ita",...)  -> lista explícita de códigos a mantener (selección personalizada)
function Resolve-IdiomasSubsLote($grupos) {
    # Decisión pre-elegida en la GUI: "TODOS" (sin filtro), "CAST_ENG" o lista de códigos.
    # Si está definida, se aplica SIEMPRE (la GUI lo eligió explícitamente) y se omite el escaneo.
    $cfgFiltro = Get-CfgGui "IdiomasSubsMantener"
    if ($cfgFiltro) {
        if ("$cfgFiltro" -eq "TODOS")    { Write-Host "`n[GUI] Subtítulos: mantener todos los idiomas (configuración)." -ForegroundColor DarkGray; return $null }
        if ("$cfgFiltro" -eq "CAST_ENG") { Write-Host "`n[GUI] Subtítulos: solo castellano e inglés (configuración)." -ForegroundColor DarkGray; return @("__CAST_ENG__") }
        Write-Host "`n[GUI] Subtítulos: mantener idiomas $(@($cfgFiltro) -join ', ') (configuración)." -ForegroundColor DarkGray
        return @($cfgFiltro)
    }

    # Buscar el conjunto de idiomas más amplio del lote (normalmente todos los capítulos
    # comparten los mismos; usamos el archivo con más idiomas como representante).
    $maxIdiomas = @()
    $total = $grupos.Count
    $nProc = 0
    foreach ($g in $grupos) {
        $nProc++
        $hibridoInfo = Resolve-Hibrido $g
        Write-Progress -Activity "Analizando idiomas de subtítulos en el lote" -Status "$nProc/$total : $($hibridoInfo.ArchivoPrincipal.Name)" -PercentComplete (($nProc/$total)*100)
        $idiomas = Get-IdiomasSubsArchivo $hibridoInfo.OrigenesPistas
        if ($idiomas.Count -gt $maxIdiomas.Count) { $maxIdiomas = $idiomas }
    }
    Write-Progress -Activity "Analizando idiomas de subtítulos en el lote" -Completed

    if ($maxIdiomas.Count -le 3) { return $null }   # 3 o menos: no preguntamos

    Write-Host "`n>> Se han detectado $($maxIdiomas.Count) idiomas distintos de subtítulos:" -ForegroundColor Cyan
    foreach ($idi in $maxIdiomas) { Write-Host "   - $($idi.NomLang) ($($idi.CodLang))" -ForegroundColor Gray }

    $opcion = Mostrar-Menu "¿Qué quieres hacer con los subtítulos?" @(
        @{Nombre="Mantener solo castellano e inglés"; Valor="CAST_ENG"},
        @{Nombre="Selección personalizada de idiomas"; Valor="PERSONALIZADA"},
        @{Nombre="Mantener todos"; Valor="TODOS"}
    )

    switch ($opcion) {
        "TODOS"    { return $null }
        "CAST_ENG" { return @("__CAST_ENG__") }
        "PERSONALIZADA" {
            Write-Host "`nIdiomas de subtítulos detectados:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $maxIdiomas.Count; $i++) {
                Write-Host "   [$($i+1)] $($maxIdiomas[$i].NomLang) ($($maxIdiomas[$i].CodLang))"
            }
            $seleccion = @()
            while ($seleccion.Count -eq 0) {
                $raw = Read-Host "Escribe los números a MANTENER (separados por comas)"
                $indices = $raw -split "[,\s]+" | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
                foreach ($idx in $indices) {
                    if ($idx -ge 1 -and $idx -le $maxIdiomas.Count) {
                        $seleccion += $maxIdiomas[$idx-1].CodLang
                    }
                }
                if ($seleccion.Count -eq 0) { Write-Host "   Selección no válida, inténtalo de nuevo." -ForegroundColor Yellow }
            }
            return @($seleccion | Select-Object -Unique)
        }
    }
    return $null
}

# Versión por archivo (HETEROGÉNEO): pregunta para un único archivo si tiene >3 idiomas de subs.
# Devuelve la misma estructura que Resolve-IdiomasSubsLote.
function Resolve-IdiomasSubsArchivo($rutasArchivosMkv) {
    # Decisión pre-elegida en la GUI (mismas reglas que Resolve-IdiomasSubsLote).
    $cfgFiltro = Get-CfgGui "IdiomasSubsMantener"
    if ($cfgFiltro) {
        if ("$cfgFiltro" -eq "TODOS")    { return $null }
        if ("$cfgFiltro" -eq "CAST_ENG") { return @("__CAST_ENG__") }
        return @($cfgFiltro)
    }

    $idiomas = Get-IdiomasSubsArchivo $rutasArchivosMkv
    if ($idiomas.Count -le 3) { return $null }

    Write-Host "`n>> Este archivo tiene $($idiomas.Count) idiomas distintos de subtítulos:" -ForegroundColor Cyan
    foreach ($idi in $idiomas) { Write-Host "   - $($idi.NomLang) ($($idi.CodLang))" -ForegroundColor Gray }

    $opcion = Mostrar-Menu "¿Qué quieres hacer con los subtítulos?" @(
        @{Nombre="Mantener solo castellano e inglés"; Valor="CAST_ENG"},
        @{Nombre="Selección personalizada de idiomas"; Valor="PERSONALIZADA"},
        @{Nombre="Mantener todos"; Valor="TODOS"}
    )
    switch ($opcion) {
        "TODOS"    { return $null }
        "CAST_ENG" { return @("__CAST_ENG__") }
        "PERSONALIZADA" {
            Write-Host "`nIdiomas de subtítulos detectados:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $idiomas.Count; $i++) {
                Write-Host "   [$($i+1)] $($idiomas[$i].NomLang) ($($idiomas[$i].CodLang))"
            }
            $seleccion = @()
            while ($seleccion.Count -eq 0) {
                $raw = Read-Host "Escribe los números a MANTENER (separados por comas)"
                $indices = $raw -split "[,\s]+" | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
                foreach ($idx in $indices) {
                    if ($idx -ge 1 -and $idx -le $idiomas.Count) { $seleccion += $idiomas[$idx-1].CodLang }
                }
                if ($seleccion.Count -eq 0) { Write-Host "   Selección no válida, inténtalo de nuevo." -ForegroundColor Yellow }
            }
            return @($seleccion | Select-Object -Unique)
        }
    }
    return $null
}

# Pre-escaneo del lote en modo HOMOGÉNEO: detecta los (idioma+formato) que aparecen como sub único
# sin señal en al menos un archivo, y hace UNA pregunta por combinación al usuario.
# Devuelve una hashtable con:
#   - claves "<idioma>|<formato>" para decisiones globales (aplicar a toda la serie)
#   - clave "<idioma>|<formato>|__PREGUNTAR__" cuando el usuario quiere decidir archivo por archivo
# La memoria por archivo concreto se rellenará después.
function Resolve-SubsUnicosLote($grupos, $idiomasSubsMantener = $null) {
    $decisiones = @{}

    # 1. Pre-escanear todos los archivos del lote
    Write-Host "`nPre-escaneando subtítulos del lote..." -ForegroundColor DarkGray
    $combinaciones = @{}   # "idioma|fmt" -> @{ NomLang=...; Archivos=@(path1, path2, ...) }
    $total = $grupos.Count
    $nProc = 0
    foreach ($g in $grupos) {
        $nProc++
        $hibridoInfo = Resolve-Hibrido $g
        Write-Progress -Activity "Pre-escaneando subtítulos del lote" -Status "$nProc/$total : $($hibridoInfo.ArchivoPrincipal.Name)" -PercentComplete (($nProc/$total)*100)
        Set-ProgresoGui ([int](8 + 6 * $nProc / [Math]::Max(1,$total))) "Pre-escaneando subtítulos ($nProc/$total)" $hibridoInfo.ArchivoPrincipal.Name
        $rutas = $hibridoInfo.OrigenesPistas
        $candidatos = Get-SubsUnicosArchivo $rutas
        foreach ($c in $candidatos) {
            # Respetar el filtro de idiomas: no preguntar por subs que se van a descartar.
            if ($idiomasSubsMantener) {
                $seMantiene = if ($idiomasSubsMantener -contains "__CAST_ENG__") { $c.CodLang -in @("es", "eng") } else { $c.CodLang -in $idiomasSubsMantener }
                if (-not $seMantiene) { continue }
            }
            $k = "$($c.CodLang)|$($c.FmtSimple)"
            if (-not $combinaciones.ContainsKey($k)) {
                $combinaciones[$k] = @{ NomLang = $c.NomLang; CodLang = $c.CodLang; FmtSimple = $c.FmtSimple; Archivos = @() }
            }
            # Usamos el ArchivoPrincipal como identificador (es el que verá Build-PistasBrutas)
            $combinaciones[$k].Archivos += $hibridoInfo.ArchivoPrincipal.FullName
        }
    }
    Write-Progress -Activity "Pre-escaneando subtítulos del lote" -Completed

    if ($combinaciones.Keys.Count -eq 0) {
        Write-Host "   -> No se detectaron subtítulos únicos ambiguos en el lote." -ForegroundColor DarkGray
        return $decisiones
    }

    # 2. Mostrar resumen
    Write-Host "`n>> Subtítulos individuales detectados (sin señal de forzado en mediainfo)" -ForegroundColor Cyan
    foreach ($k in $combinaciones.Keys) {
        $info = $combinaciones[$k]
        Write-Host "   - $($info.NomLang) ($($info.FmtSimple)): $($info.Archivos.Count) archivo(s) con un único sub" -ForegroundColor Gray
    }

    # 3. Preguntar por cada combinación
    foreach ($k in $combinaciones.Keys) {
        $info = $combinaciones[$k]

        # Clave bajo la que se GUARDA la decisión. Si el idioma era 'und' y se resuelve, hay que
        # guardar con el código RESUELTO: es el que tendrán las pistas tras Build-PistasBrutas y
        # el que Set-FlagForzados usará al consultar. (Antes se guardaba "...|und|..." y la
        # decisión del usuario se perdía en silencio.)
        $kDecision = $k

        # Si el idioma es indeterminado (und): el pre-escaneo Resolve-IdiomasUndLote ya preguntó
        # el idioma de los subs 'und' del lote. Lo reutilizamos sin volver a preguntar. Solo si
        # por algún motivo no se decidió antes, preguntamos aquí como respaldo.
        if ($info.CodLang -eq "und") {
            $codElegido = if ($global:idiomaUndSubLote) { $global:idiomaUndSubLote } else { Resolve-IdiomaSubUnd "$($info.NomLang) ($($info.FmtSimple))" }
            if ($codElegido -ne "und") {
                foreach ($f in $info.Archivos) { $global:decisionesIdiomaUnd["${f}|$($info.FmtSimple)"] = $codElegido }
                $info.NomLang = (Get-LanguageName $codElegido "")
                $kDecision = "$codElegido|$($info.FmtSimple)"
            }
        }

        # Decisión POR sub (idioma+formato) elegida en la GUI; respaldo: decisión global.
        $cfgSubUnico = Get-DecisionSubGui $info.CodLang $info.FmtSimple
        $resp = switch ("$cfgSubUnico") {
            "Forzado"   { "1" }
            "Completo"  { "2" }
            "PREGUNTAR" { "3" }
            default     { $null }
        }
        if ($resp) {
            $txtResp = switch ($resp) { "1" { "Forzado" } "2" { "Completo" } "3" { "preguntar por archivo" } }
            Write-Host "   [GUI] Sub único $($info.NomLang) ($($info.FmtSimple)) -> $txtResp (configuración)." -ForegroundColor DarkGray
        } else {
            Write-Host "`n   ¿El subtítulo $($info.NomLang) ($($info.FmtSimple)) es...?" -ForegroundColor Cyan
            Write-Host "     [1] Forzado en todos los archivos"
            Write-Host "     [2] Completo en todos los archivos"
            Write-Host "     [3] Preguntar para cada archivo durante el procesamiento"
            while (-not $resp) {
                $r = Read-Host "   Elige una opción (1-3)"
                if ($r -match '^[1-3]$') { $resp = $r }
                else { Write-Host "   Entrada no válida." -ForegroundColor Yellow }
            }
        }
        switch ($resp) {
            "1" { foreach ($f in $info.Archivos) { $decisiones["${f}|$kDecision"] = "Forzado" } }
            "2" { foreach ($f in $info.Archivos) { $decisiones["${f}|$kDecision"] = "Completo" } }
            "3" { foreach ($f in $info.Archivos) { $decisiones["${f}|$kDecision"] = "__PREGUNTAR__" } }
        }
    }
    return $decisiones
}

# Pregunta para UN solo archivo (modo HETEROGÉNEO o cuando la decisión homogénea fue "__PREGUNTAR__").
# Mira los subs únicos del archivo en cuestión y pregunta por cada combinación.
# Actualiza $decisiones in-place.
function Resolve-SubsUnicosArchivo($rutasArchivosMkv, $archivoPrincipal, $decisiones, $idiomasSubsMantener = $null) {
    $candidatos = Get-SubsUnicosArchivo $rutasArchivosMkv
    if (-not $candidatos -or $candidatos.Count -eq 0) { return }
    $pathRef = $archivoPrincipal.FullName

    foreach ($c in $candidatos) {
        # Respetar el filtro de idiomas: no preguntar por subs de idiomas que se van a descartar.
        if ($idiomasSubsMantener) {
            $seMantiene = if ($idiomasSubsMantener -contains "__CAST_ENG__") { $c.CodLang -in @("es", "eng") } else { $c.CodLang -in $idiomasSubsMantener }
            if (-not $seMantiene) { continue }
        }

        # Idioma indeterminado: resolverlo ANTES de construir la clave de decisión, para que esta
        # coincida con el código que tendrán las pistas (y con las decisiones del pre-escaneo del
        # lote, que también se guardan con el código resuelto). Si el lote ya decidió el idioma,
        # se reutiliza sin volver a preguntar.
        $nomMostrar = $c.NomLang
        $codEfectivo = $c.CodLang
        if ($c.CodLang -eq "und") {
            $kUnd = "${pathRef}|$($c.FmtSimple)"
            if ($global:decisionesIdiomaUnd.ContainsKey($kUnd)) {
                $codEfectivo = $global:decisionesIdiomaUnd[$kUnd]
            } elseif ($global:idiomaUndSubLote) {
                $codEfectivo = $global:idiomaUndSubLote
                $global:decisionesIdiomaUnd[$kUnd] = $codEfectivo
            } else {
                $codElegido = Resolve-IdiomaSubUnd "$($c.NomLang) ($($c.FmtSimple))"
                if ($codElegido -ne "und") {
                    $global:decisionesIdiomaUnd[$kUnd] = $codElegido
                    $codEfectivo = $codElegido
                }
            }
            if ($codEfectivo -ne "und") { $nomMostrar = (Get-LanguageName $codEfectivo "") }
        }

        $k = "$codEfectivo|$($c.FmtSimple)"
        $llave = "${pathRef}|$k"
        # Si ya hay decisión global (1/2) la respetamos. Si es __PREGUNTAR__ o no existe, preguntamos aquí.
        $decisionPrevia = $null
        if ($decisiones.ContainsKey($llave)) { $decisionPrevia = $decisiones[$llave] }
        if ($decisionPrevia -eq "Forzado" -or $decisionPrevia -eq "Completo") { continue }

        # Decisión pre-elegida en la GUI (modo heterogéneo): por sub (idioma+formato), con respaldo global.
        $cfgSubUnico = Get-DecisionSubGui $c.CodLang $c.FmtSimple
        if ("$cfgSubUnico" -in @("Forzado", "Completo")) {
            $decisiones[$llave] = "$cfgSubUnico"
            Write-Host "   [GUI] Sub único $nomMostrar ($($c.FmtSimple)) -> $cfgSubUnico (configuración)." -ForegroundColor DarkGray
            continue
        }

        Write-Host "`n   ¿El subtítulo $nomMostrar ($($c.FmtSimple)) es...?" -ForegroundColor Cyan
        Write-Host "     [1] Forzado"
        Write-Host "     [2] Completo"
        $resp = $null
        while (-not $resp) {
            $r = Read-Host "   Elige una opción (1-2)"
            if ($r -match '^[1-2]$') { $resp = $r }
            else { Write-Host "   Entrada no válida." -ForegroundColor Yellow }
        }
        switch ($resp) {
            "1" { $decisiones[$llave] = "Forzado" }
            "2" { $decisiones[$llave] = "Completo" }
        }
    }
}

# Pregunta forzado/completo de los SUBTÍTULOS EXTERNOS únicos (SRT/ASS sueltos en la carpeta,
# o convertidos de PGS). Se llama JUSTO ANTES de muxear cada archivo (tras la pausa de Subtitle
# Edit, si la hubo), porque es cuando esos SRT ya existen en la carpeta.
# - Si el nombre del archivo externo contiene "forced"/"forzado" → se asume forzado, no se pregunta.
# - Si queda único en su grupo (idioma+formato) sin señal → se pregunta.
# La decisión se guarda en $decisiones con clave "EXT|<nombreArchivo>", que Set-FlagForzados aplica.
function Resolve-SubsExternosForzados($rutasArchivosMkv, $prefijoArchivo, $decisiones, $idiomasSubsMantener = $null) {
    $candidatos = Get-SubsUnicosArchivo $rutasArchivosMkv $true $prefijoArchivo
    if (-not $candidatos -or $candidatos.Count -eq 0) { return }

    foreach ($c in $candidatos) {
        # Solo nos interesan los EXTERNOS aquí (los internos ya se preguntaron en su momento).
        if (-not $c.EsExterno) { continue }
        # Respetar el filtro de idiomas.
        if ($idiomasSubsMantener) {
            $seMantiene = if ($idiomasSubsMantener -contains "__CAST_ENG__") { $c.CodLang -in @("es", "eng") } else { $c.CodLang -in $idiomasSubsMantener }
            if (-not $seMantiene) { continue }
        }
        # Si el nombre ya indica forzado, asumir forzado y NO preguntar.
        if ($c.TieneSenalForzado) {
            $decisiones["EXT|$($c.NombreExterno)"] = "Forzado"
            Write-Host "   [subs externos] '$($c.NombreExterno)' detectado como FORZADO por el nombre." -ForegroundColor DarkGray
            continue
        }
        $kExt = "EXT|$($c.NombreExterno)"
        if ($decisiones.ContainsKey($kExt)) { continue }   # ya decidido

        # Decisión pre-elegida en la GUI: por sub (idioma+formato), con respaldo global.
        $cfgSubUnico = Get-DecisionSubGui $c.CodLang $c.FmtSimple
        if ("$cfgSubUnico" -in @("Forzado", "Completo")) {
            $decisiones[$kExt] = "$cfgSubUnico"
            Write-Host "   [GUI] Sub externo '$($c.NombreExterno)' -> $cfgSubUnico (configuración)." -ForegroundColor DarkGray
            continue
        }

        Write-Host "`n   ¿El subtítulo externo $($c.NomLang) '$($c.NombreExterno)' es...?" -ForegroundColor Cyan
        Write-Host "     [1] Forzado"
        Write-Host "     [2] Completo"
        $resp = $null
        while (-not $resp) {
            $r = Read-Host "   Elige una opción (1-2)"
            if ($r -match '^[1-2]$') { $resp = $r }
            else { Write-Host "   Entrada no válida." -ForegroundColor Yellow }
        }
        switch ($resp) {
            "1" { $decisiones[$kExt] = "Forzado" }
            "2" { $decisiones[$kExt] = "Completo" }
        }
    }
}
# $rutaMkvOrigen: si se proporciona, cuando un grupo no tiene tamaños reportados se intentará
# extraer cada sub interno y medir su peso en disco.
function Set-FlagForzados($listaSubs, $rutaMkvOrigen = $null, $decisionesSubUnico = $null) {
    if (-not $listaSubs -or $listaSubs.Count -eq 0) { return }

    # Agrupar por idioma + formato (lo hacemos a mano para no depender de Group-Object,
    # que en algunos casos puede crear copias o trabajar con propiedades calculadas).
    # Importante: usamos el FORMATO SIMPLIFICADO (Text vs PGS) para que variantes como
    # "UTF-8" y "SubRip" caigan en el mismo grupo: son técnicamente el mismo formato de
    # subtítulo (SubRip Text). Si los tratásemos por separado nunca podríamos distinguir
    # forzado vs completo cuando vienen como variantes distintas en el mismo MKV.
    $clavesGrupo = @{}
    for ($i = 0; $i -lt $listaSubs.Count; $i++) {
        $fmtSimple = if ($listaSubs[$i].Format -match "(?i)srt|subrip|ass|ssa|utf-?8|text") { "Text" } else { "PGS" }
        $clave = "$($listaSubs[$i].CodLang)__$fmtSimple"
        if (-not $clavesGrupo.ContainsKey($clave)) { $clavesGrupo[$clave] = @() }
        $clavesGrupo[$clave] += $i   # guardamos el índice, no el objeto
    }

    foreach ($clave in $clavesGrupo.Keys) {
        $indices = $clavesGrupo[$clave]
        Write-Host "   [debug forzados] Grupo '$clave': $($indices.Count) sub(s)" -ForegroundColor DarkGray

        # Paso 1: respetar flag explícito de mediainfo si alguno lo trae
        $algunoConFlag = $false
        foreach ($idx in $indices) {
            if ($listaSubs[$idx].MediainfoForcedFlag -eq $true) {
                $listaSubs[$idx].IsForced = $true
                $algunoConFlag = $true
                Write-Host "       -> Sub idx=$idx marcado FORZADO por flag mediainfo" -ForegroundColor DarkGray
            }
        }

        # Paso 2: si nadie tenía flag y hay más de uno, heurística por tamaño
        if (-not $algunoConFlag -and $indices.Count -gt 1) {
            $maxS = 0
            foreach ($idx in $indices) {
                if ($listaSubs[$idx].Size -gt $maxS) { $maxS = $listaSubs[$idx].Size }
            }

            # Paso 2.5: si nadie reporta tamaño y los subs son INTERNOS, extraerlos a disco
            # y medir. Cada pista interna sabe de qué MKV viene (ArchivoFuenteMkv), así que
            # extraemos del archivo correcto (importante en híbridos con merge de pistas).
            if ($maxS -eq 0) {
                $extraibles = @()
                foreach ($idx in $indices) {
                    if ($listaSubs[$idx].Origen -eq "Interno" -and $listaSubs[$idx].ArchivoFuenteMkv) { $extraibles += $idx }
                }
                if ($extraibles.Count -gt 1) {
                    Write-Host "       -> Tamaños no reportados; extrayendo subs a disco para medir..." -ForegroundColor DarkGray
                    foreach ($idx in $extraibles) {
                        $rutaSubMkv = $listaSubs[$idx].ArchivoFuenteMkv
                        $tam = Measure-SubEnDisco $rutaSubMkv $listaSubs[$idx].ID
                        $listaSubs[$idx].Size = $tam
                        Write-Host "          sub idx=$idx (ID mkv=$($listaSubs[$idx].ID), de $(Split-Path $rutaSubMkv -Leaf)) -> $tam bytes" -ForegroundColor DarkGray
                        if ($tam -gt $maxS) { $maxS = $tam }
                    }
                }
            }

            if ($maxS -gt 0) {
                foreach ($idx in $indices) {
                    if ($listaSubs[$idx].Size -gt 0 -and $listaSubs[$idx].Size -le ($maxS * 0.35)) {
                        $listaSubs[$idx].IsForced = $true
                        Write-Host "       -> Sub idx=$idx marcado FORZADO por heurística tamaño ($($listaSubs[$idx].Size) <= 35% de $maxS)" -ForegroundColor DarkGray
                    }
                }
            } else {
                Write-Host "       -> Heurística no aplicable: ningún sub del grupo reporta tamaño ni se pudo extraer" -ForegroundColor DarkYellow
            }
        }

        # Paso 3: consultar memoria de decisiones del usuario (preguntas pre-escaneo de sub único).
        # Aplica cuando el grupo, después de todo, sigue con un sub sin marca de forzado y el usuario
        # respondió "Forzado" en la pregunta inicial para este (archivo + idioma + formato).
        # Si dijo "Completo", no hacemos nada (IsForced ya es $false por defecto).
        if ($decisionesSubUnico -and $decisionesSubUnico.Count -gt 0) {
            $fmtClave = ($clave -split "__")[1]   # "es__Text" -> "Text"
            $langClave = ($clave -split "__")[0]
            foreach ($idx in $indices) {
                if ($listaSubs[$idx].IsForced -eq $true) { continue }

                if ($listaSubs[$idx].Origen -eq "Interno") {
                    # Construir las claves a probar: por ArchivoFuenteMkv (donde vive físicamente)
                    # y por ArchivoPrincipal (donde se registró la decisión del usuario).
                    $rutaSubMkv = $listaSubs[$idx].ArchivoFuenteMkv
                    $clavesAProbar = @(
                        "${rutaSubMkv}|${langClave}|${fmtClave}"
                    )
                    if ($rutaMkvOrigen -and $rutaMkvOrigen -ne $rutaSubMkv) {
                        $clavesAProbar += "${rutaMkvOrigen}|${langClave}|${fmtClave}"
                    }
                    foreach ($k in $clavesAProbar) {
                        if ($decisionesSubUnico.ContainsKey($k) -and $decisionesSubUnico[$k] -eq "Forzado") {
                            $listaSubs[$idx].IsForced = $true
                            Write-Host "       -> Sub idx=$idx marcado FORZADO por decisión del usuario (pregunta sub único)" -ForegroundColor DarkGray
                            break
                        }
                    }
                } else {
                    # Pista EXTERNA (SRT suelto / convertido de PGS): la decisión se registró con clave
                    # basada en el nombre del archivo externo ("EXT|<nombreArchivo>").
                    $nombreExt = $listaSubs[$idx].NombreExterno
                    if ($nombreExt) {
                        $kExt = "EXT|$nombreExt"
                        if ($decisionesSubUnico.ContainsKey($kExt) -and $decisionesSubUnico[$kExt] -eq "Forzado") {
                            $listaSubs[$idx].IsForced = $true
                            Write-Host "       -> Sub externo idx=$idx ($nombreExt) marcado FORZADO por decisión del usuario" -ForegroundColor DarkGray
                        }
                    }
                }
            }
        }
    }
}

# Recoge los datos de proyecto (título, año, serie, origen, plataforma, etiquetas).
# Se usa una vez en modo homogéneo, o una vez por archivo en modo heterogéneo.
# El parámetro $contexto se imprime como cabecera para que en heterogéneo se vea
# de qué archivo se están preguntando los datos.
function Get-DatosProyecto($contexto) {
    # Datos pre-rellenados desde la GUI. En modo heterogéneo (con $contexto) solo se usan si
    # el usuario marcó "aplicar a todos los archivos"; si no, se pregunta por consola por cada uno.
    $cfgProy = Get-CfgGui "Proyecto"
    if ($cfgProy -and "$($cfgProy.Titulo)" -ne "" -and (-not $contexto -or $cfgProy.AplicarATodos)) {
        if ($contexto) { Write-Host "`n[GUI] Usando datos de proyecto de la configuración para: $contexto" -ForegroundColor DarkGray }
        else           { Write-Host "`n[GUI] Usando datos de proyecto de la configuración." -ForegroundColor DarkGray }
        $origenCfg  = if ("$($cfgProy.TipoOrigen)" -eq "FISICO") { "FISICO" } else { "WEB" }
        $webTipoCfg = if ("$($cfgProy.WebTipo)" -in @("WEB-DL", "WEBRip")) { "$($cfgProy.WebTipo)" } else { "WEB-DL" }
        return [PSCustomObject]@{
            Titulo            = "$($cfgProy.Titulo)"
            Ano               = "$($cfgProy.Ano)"
            EsSerie           = [bool]$cfgProy.EsSerie
            TipoOrigen        = $origenCfg
            WebTipo           = $webTipoCfg
            PlataformaFormato = "$($cfgProy.PlataformaFormato)"
            EtiquetasExtra    = "$($cfgProy.EtiquetasExtra)"
        }
    }

    if ($contexto) {
        Write-Host "`n--- DATOS PARA: $contexto ---" -ForegroundColor Yellow
    } else {
        Write-Host "`n--- CONFIGURACIÓN GLOBAL DEL PROYECTO ---" -ForegroundColor Yellow
    }

    $titulo = Read-Host "1. Título limpio"
    $ano    = Read-Host "2. Año"
    $serie  = ((Read-Host "3. ¿Es una serie? (S/N)") -match "^[sS]")

    $origen = Mostrar-Menu "4. Origen" @(@{Nombre="WEB-DL"; Valor="WEB"}, @{Nombre="Físico (Remux, Full...)"; Valor="FISICO"})
    # Dentro de WEB hay dos variantes que cambian la etiqueta del nombre: WEB-DL y WEBRip.
    $webTipo = "WEB-DL"
    if ($origen -eq "WEB") {
        $webTipo = Mostrar-Menu "Tipo de WEB" @(@{Nombre="WEB-DL"; Valor="WEB-DL"}, @{Nombre="WEBRip"; Valor="WEBRip"})
    }
    $plat   = if ($origen -eq "WEB") { Mostrar-Menu "Plataforma" $opcPlataformas } else { Mostrar-Menu "Formato Físico" $opcFormatosFisicos }
    $etiq   = if ($origen -eq "FISICO") { Read-Host "   -> Etiquetas Extra (opcional)" } else { "" }

    return [PSCustomObject]@{
        Titulo            = $titulo
        Ano               = $ano
        EsSerie           = $serie
        TipoOrigen        = $origen
        WebTipo           = $webTipo    # "WEB-DL" o "WEBRip" (solo relevante si TipoOrigen=WEB)
        PlataformaFormato = $plat
        EtiquetasExtra    = $etiq
    }
}

# =========================================================================
# FUNCIONES DE FASE (D1)
# =========================================================================

# Detecta si un grupo es híbrido DV+HDR10 y devuelve un objeto con la información.
# Cacheado: se llama varias veces por grupo (pre-escaneos de und/idiomas/subs/DTS/PGS y el
# bucle principal) y cada llamada ejecutaba ffprobe completo sobre todos los archivos.
function Resolve-Hibrido($grupo) {
    $claveCache = (@($grupo.Group | ForEach-Object { $_.FullName }) -join "||")
    if ($cacheHibrido.ContainsKey($claveCache)) { return $cacheHibrido[$claveCache] }

    $info = [PSCustomObject]@{
        EsHibrido = $false; ArchivoDV = $null; ArchivoHDR10 = $null
        ArchivoPrincipal = $grupo.Group[0]
        OrigenPistas = $grupo.Group[0].FullName    # compat: 1er archivo
        OrigenesPistas = @($grupo.Group[0].FullName)  # nueva: lista de archivos para mergear pistas
    }

    if ($grupo.Count -lt 2) { $cacheHibrido[$claveCache] = $info; return $info }

    foreach ($f in $grupo.Group) {
        $hasDV = (ffprobe -v quiet -show_streams -select_streams v:0 "$($f.FullName)") -join "`n"
        if ($hasDV -match "dv_profile=") { $info.ArchivoDV = $f } else { $info.ArchivoHDR10 = $f }
    }

    if ($info.ArchivoDV -and $info.ArchivoHDR10) {
        $info.EsHibrido = $true
        # En híbrido leemos pistas de AMBOS archivos. El orden importa: ponemos primero el que
        # tiene mejor metadata, así en caso de duplicado (mismo idioma+códec+canales) prevalece
        # la pista del archivo "bueno". También usamos ese como ArchivoPrincipal para nombre/serie.
        $metaDV  = Get-CalidadMetadata $info.ArchivoDV.FullName
        $metaHDR = Get-CalidadMetadata $info.ArchivoHDR10.FullName
        Write-Host "   [debug híbrido] HDR10 score metadata=$metaHDR; DV score metadata=$metaDV" -ForegroundColor DarkGray

        if ($metaDV -gt $metaHDR) {
            $info.OrigenesPistas = @($info.ArchivoDV.FullName, $info.ArchivoHDR10.FullName)
            $info.OrigenPistas    = $info.ArchivoDV.FullName
            $info.ArchivoPrincipal = $info.ArchivoDV
            Write-Host "   [debug híbrido] -> Lectura de pistas: DV (preferente) + HDR10 (merge)" -ForegroundColor DarkGray
        } else {
            $info.OrigenesPistas = @($info.ArchivoHDR10.FullName, $info.ArchivoDV.FullName)
            $info.OrigenPistas    = $info.ArchivoHDR10.FullName
            $info.ArchivoPrincipal = $info.ArchivoHDR10
            Write-Host "   [debug híbrido] -> Lectura de pistas: HDR10 (preferente) + DV (merge)" -ForegroundColor DarkGray
        }
    }

    $cacheHibrido[$claveCache] = $info
    return $info
}

# Calcula un "score" de calidad de metadata para elegir el mejor archivo cuando hay híbrido.
# Suma puntos por cada pista de audio/texto que tenga:
#   +2 si trae StreamSize numérico (no vacío) — indicador fuerte de metadata curada
#   +3 si trae Forced=Yes — confirma que el archivo distingue forzados de completos
#   +1 si trae Default=Yes/No no vacío — metadata mínima presente
# Un archivo "pelado" (todos los Forced a "No" sin StreamSize) puntúa muy bajo.
# Un archivo curado (con Forced=Yes en alguna pista, y StreamSize rellenos) puntúa alto.
function Get-CalidadMetadata($ruta) {
    $json = Get-MediainfoJson $ruta
    if (-not $json) { return 0 }
    $score = 0
    foreach ($pista in $json.media.track) {
        if ($pista.'@type' -match "Audio|Text") {
            if ($pista.PSObject.Properties.Name -contains "StreamSize" -and -not [string]::IsNullOrWhiteSpace("$($pista.StreamSize)")) {
                $score += 2
            }
            if ($pista.PSObject.Properties.Name -contains "Forced" -and "$($pista.Forced)" -match "(?i)yes|1|true") {
                $score += 3
            }
            if ($pista.PSObject.Properties.Name -contains "Default" -and -not [string]::IsNullOrWhiteSpace("$($pista.Default)")) {
                $score += 1
            }
        }
    }
    return $score
}

# Resuelve subtítulos con idioma indeterminado preguntando al usuario (con memoria).
# $idiomaGlobalUnd: idioma elegido en el pre-escaneo del lote para los subs 'und' internos
# (modo homogéneo). Si viene, se aplica directamente sin volver a preguntar.
function Resolve-IdiomasSubs($origenesPistas, $archivoPrincipal, $memoria, $idiomaGlobalUnd = $null) {
    $subsDesconocidos = @()

    # Aceptamos tanto un único path (string) como un array de paths
    $listaArchivos = @($origenesPistas)

    foreach ($rutaArchivo in $listaArchivos) {
        $miJsonOrigen = Get-MediainfoJson $rutaArchivo
        if (-not $miJsonOrigen) { continue }
        $nombreArchivo = Split-Path $rutaArchivo -Leaf
        $idMapL = Get-MkvmergeTrackIdMap $rutaArchivo
        $contTextL = 0
        foreach ($pista in $miJsonOrigen.media.track) {
            if ($pista.'@type' -eq "Text") {
                $contTextL++
                $claveMap = "Text+$contTextL"
                $idReal = if ($idMapL.ContainsKey($claveMap)) { $idMapL[$claveMap] } else { $pista.ID - 1 }
                $langCode = $pista.Language
                if ([string]::IsNullOrWhiteSpace($langCode) -or $langCode -eq "und") {
                    # D4: clave de memoria incluye path del archivo para evitar colisiones
                    $subsDesconocidos += [PSCustomObject]@{
                        Origen = "Interno"
                        ID = $idReal
                        ArchivoFuente = $rutaArchivo
                        Info = "Subtítulo Interno (ID Pista: $idReal) en $nombreArchivo"
                    }
                }
            }
        }
    }

    $archivosExternosSubs = Get-ChildItem -LiteralPath $rutaCarpeta -File | Where-Object { $_.Extension -match "\.(srt|ass|sup|ssa|pgs|idx|sub)$" }
    # El .sub de una pareja VobSub .idx/.sub se gestiona a través de su .idx (mkvmerge lo toma
    # automáticamente): no preguntamos su idioma por separado.
    $archivosExternosSubs = @($archivosExternosSubs | Where-Object {
        -not ($_.Extension -match "(?i)^\.sub$" -and (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($_.FullName, ".idx"))))
    })
    foreach ($ext in $archivosExternosSubs) {
        $miExt = Get-MediainfoJson $ext.FullName
        $pExt = $null
        if ($miExt) { $pExt = $miExt.media.track | Where-Object { $_.'@type' -eq "Text" } | Select-Object -First 1 }
        if ($pExt -or $ext.Extension -match "srt|ass|ssa|sup|pgs|idx|sub") {
            $langCode = if ($pExt) { $pExt.Language } else { "" }
            if ([string]::IsNullOrWhiteSpace($langCode) -or $langCode -eq "und") {
                if     ($ext.Name -match "(?i)(?<![a-z])(latino|latam|lat)(?![a-z])")           { $memoria["Externo_$($ext.FullName)"] = "es-419" }
                elseif ($ext.Name -match "(?i)(?<![a-z])(spa|es|castellano|cast)(?![a-z])")     { $memoria["Externo_$($ext.FullName)"] = "es" }
                elseif ($ext.Name -match "(?i)(?<![a-z])(eng|en|ingles|english)(?![a-z])")      { $memoria["Externo_$($ext.FullName)"] = "eng" }
                elseif ($ext.Name -match "(?i)(?<![a-z])(fre|fr|frances|french)(?![a-z])")      { $memoria["Externo_$($ext.FullName)"] = "fre" }
                else { 
                    $subsDesconocidos += [PSCustomObject]@{
                        Origen = "Externo"; ID = $ext.FullName; ArchivoFuente = $ext.FullName
                        Info = "Archivo Externo: $($ext.Name)"
                    }
                }
            }
        }
    }

    if ($subsDesconocidos.Count -gt 0) {
        Write-Host "`n[!] RESOLUCIÓN DE CONFLICTOS DE IDIOMA:" -ForegroundColor Yellow
        foreach ($sub in $subsDesconocidos) {
            # D4: clave incluye el path para internos también
            $llave = if ($sub.Origen -eq "Interno") { "Interno_$($sub.ArchivoFuente)_$($sub.ID)" } else { "Externo_$($sub.ID)" }
            if (-not $memoria.ContainsKey($llave)) {
                # Si el pre-escaneo del lote ya decidió el idioma de los subs 'und' internos
                # (modo homogéneo), se aplica directamente SIN volver a preguntar (antes se
                # preguntaba dos veces: una a nivel de lote y otra por cada pista).
                if ($idiomaGlobalUnd -and $sub.Origen -eq "Interno") {
                    $memoria[$llave] = $idiomaGlobalUnd
                    continue
                }
                Write-Host " -> $($sub.Info)" -ForegroundColor Cyan
                for ($i=0; $i -lt $idiomasComunes.Count; $i++) { 
                    Write-Host (" {0,2}. {1}" -f ($i+1), $idiomasComunes[$i].Nom) -NoNewline
                    if (($i + 1) % 4 -eq 0) { Write-Host "" } 
                }
                Write-Host ""
                # Bucle hasta que el usuario introduzca un idioma válido. Sin idioma no podemos
                # decidir bien forzados/completos, así que es obligatorio.
                $codElegido = $null
                while (-not $codElegido) {
                    $seleccion = Read-Host "Seleccione idioma (1-$($idiomasComunes.Count), 0 para introducir código manualmente)"
                    if ($seleccion -eq "0") {
                        $manual = Read-Host "   Escribe el código de idioma (ej: rus, ara, hin, swe...)"
                        if (-not [string]::IsNullOrWhiteSpace($manual)) { $codElegido = $manual.Trim().ToLower() }
                    } elseif ($seleccion -match '^\d+$' -and [int]$seleccion -gt 0 -and [int]$seleccion -le $idiomasComunes.Count) {
                        $codElegido = $idiomasComunes[[int]$seleccion - 1].Cod
                    } else {
                        Write-Host "   Entrada no válida. Debes elegir un idioma." -ForegroundColor Yellow
                    }
                }
                $memoria[$llave] = $codElegido
            }
        }
    }
}

# Resuelve idiomas indeterminados de pistas de audio (internas y externas) preguntando al usuario.
# Comparte la misma estructura que Resolve-IdiomasSubs pero ofrece la opción "U" para dejar como und,
# porque en audio aceptamos und (la conversión a EAC3 es por compatibilidad y no depende del idioma).
function Resolve-IdiomasAudios($origenesPistas, $archivoPrincipal, $memoria, $idiomaGlobalUnd = $null) {
    $audiosDesconocidos = @()

    $listaArchivos = @($origenesPistas)

    foreach ($rutaArchivo in $listaArchivos) {
        $miJsonOrigen = Get-MediainfoJson $rutaArchivo
        if (-not $miJsonOrigen) { continue }
        $nombreArchivo = Split-Path $rutaArchivo -Leaf
        $idMapL = Get-MkvmergeTrackIdMap $rutaArchivo
        $contAudL = 0
        foreach ($pista in $miJsonOrigen.media.track) {
            if ($pista.'@type' -eq "Audio") {
                $contAudL++
                $claveMap = "Audio+$contAudL"
                $idReal = if ($idMapL.ContainsKey($claveMap)) { $idMapL[$claveMap] } else { $pista.ID - 1 }
                $langCode = $pista.Language
                if ([string]::IsNullOrWhiteSpace($langCode) -or $langCode -eq "und") {
                    $audiosDesconocidos += [PSCustomObject]@{
                        Origen = "Interno"
                        ID = $idReal
                        ArchivoFuente = $rutaArchivo
                        Info = "Pista de Audio Interna (ID Pista: $idReal) en $nombreArchivo"
                    }
                }
            }
        }
    }

    $archivosExternosAudio = Get-ChildItem -LiteralPath $rutaCarpeta -File | Where-Object { $_.Extension -match "\.(ac3|eac3|dts|dtshd|truehd|flac|wav|mka|m4a|aac|opus)$" }
    foreach ($ext in $archivosExternosAudio) {
        $miExt = Get-MediainfoJson $ext.FullName
        $pExt = $null
        if ($miExt) { $pExt = $miExt.media.track | Where-Object { $_.'@type' -eq "Audio" } | Select-Object -First 1 }
        $langCode = if ($pExt) { $pExt.Language } else { "" }
        if ([string]::IsNullOrWhiteSpace($langCode) -or $langCode -eq "und") {
            # Detección automática por nombre (idiomas comunes y patrón audio_<lang>_temp/final)
            if     ($ext.Name -match "(?i)audio_([a-z]{2,3}(?:-[a-z0-9]{2,3})?)_(temp|final)") { $memoria["Externo_$($ext.FullName)"] = $Matches[1].ToLower() }
            elseif ($ext.Name -match "(?i)(?<![a-z])(latino|latam|lat)(?![a-z])")             { $memoria["Externo_$($ext.FullName)"] = "es-419" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(spa|es|castellano|cast)(?![a-z])")       { $memoria["Externo_$($ext.FullName)"] = "es" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(eng|en|ingles|english)(?![a-z])")        { $memoria["Externo_$($ext.FullName)"] = "eng" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(fre|fr|frances|french)(?![a-z])")        { $memoria["Externo_$($ext.FullName)"] = "fre" }
            else {
                $audiosDesconocidos += [PSCustomObject]@{
                    Origen = "Externo"; ID = $ext.FullName; ArchivoFuente = $ext.FullName
                    Info = "Archivo de Audio Externo: $($ext.Name)"
                }
            }
        }
    }

    if ($audiosDesconocidos.Count -gt 0) {
        Write-Host "`n[!] RESOLUCIÓN DE IDIOMA EN PISTAS DE AUDIO:" -ForegroundColor Yellow
        foreach ($au in $audiosDesconocidos) {
            $llave = if ($au.Origen -eq "Interno") { "Interno_$($au.ArchivoFuente)_$($au.ID)" } else { "Externo_$($au.ID)" }
            if (-not $memoria.ContainsKey($llave)) {
                # Si hay un idioma global decidido en el pre-escaneo del lote (modo homogéneo),
                # lo aplicamos directamente SIN preguntar (ya se preguntó una vez para todo el lote).
                if ($idiomaGlobalUnd) {
                    $memoria[$llave] = $idiomaGlobalUnd
                    continue
                }
                Write-Host " -> $($au.Info)" -ForegroundColor Cyan
                for ($i=0; $i -lt $idiomasComunes.Count; $i++) {
                    Write-Host (" {0,2}. {1}" -f ($i+1), $idiomasComunes[$i].Nom) -NoNewline
                    if (($i + 1) % 4 -eq 0) { Write-Host "" }
                }
                Write-Host ""
                $codElegido = $null
                while (-not $codElegido) {
                    $seleccion = Read-Host "Seleccione idioma (1-$($idiomasComunes.Count), 0 para introducir código manualmente, U para dejar como und)"
                    if ($seleccion -match '^[uU]$') {
                        $codElegido = "und"
                    } elseif ($seleccion -eq "0") {
                        $manual = Read-Host "   Escribe el código de idioma (ej: rus, ara, hin, swe...)"
                        if (-not [string]::IsNullOrWhiteSpace($manual)) { $codElegido = $manual.Trim().ToLower() }
                    } elseif ($seleccion -match '^\d+$' -and [int]$seleccion -gt 0 -and [int]$seleccion -le $idiomasComunes.Count) {
                        $codElegido = $idiomasComunes[[int]$seleccion - 1].Cod
                    } else {
                        Write-Host "   Entrada no válida." -ForegroundColor Yellow
                    }
                }
                $memoria[$llave] = $codElegido
            }
        }
    }
}

# Ejecuta un proceso externo redirigiendo stdout/stderr SIN riesgo de deadlock: ambos streams
# se drenan siempre (si no se leen y el proceso llena el buffer del pipe, se queda bloqueado
# para siempre), y la lectura de stderr usa esperas con timeout en lugar de ReadLine()
# bloqueante (que podía dejar el script colgado indefinidamente si el proceso enmudecía).
# $alLeerLineaErr: scriptblock invocado por cada línea de stderr (ffmpeg emite ahí el progreso).
# Devuelve el código de salida del proceso.
function Invoke-ProcesoMonitorizado($exe, $listaArgs, $alLeerLineaErr) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe
    foreach ($a in $listaArgs) { $psi.ArgumentList.Add($a) }
    $psi.RedirectStandardError  = $true
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    # stdout no nos interesa, pero hay que drenarlo en segundo plano igualmente
    $drenajeOut = $proc.StandardOutput.ReadToEndAsync()
    $tareaLinea = $proc.StandardError.ReadLineAsync()
    while ($true) {
        $hayLinea = $false
        try { $hayLinea = $tareaLinea.Wait(250) } catch { break }
        if ($hayLinea) {
            $linea = $tareaLinea.Result
            if ($null -eq $linea) { break }   # stream cerrado: el proceso terminó
            if ($alLeerLineaErr) { & $alLeerLineaErr $linea }
            $tareaLinea = $proc.StandardError.ReadLineAsync()
        } elseif ($proc.HasExited) {
            # Proceso terminado: apurar las líneas pendientes (con timeout corto) y salir.
            try {
                while ($tareaLinea.Wait(500)) {
                    $linea = $tareaLinea.Result
                    if ($null -eq $linea) { break }
                    if ($alLeerLineaErr) { & $alLeerLineaErr $linea }
                    $tareaLinea = $proc.StandardError.ReadLineAsync()
                }
            } catch {}
            break
        }
    }
    $proc.WaitForExit()
    return $proc.ExitCode
}

# Fusiona capa Dolby Vision con base HDR10 vía dovi_tool
function Invoke-FusionDolbyVision($archivoDV, $archivoHDR10) {
    Write-Host "`n[>>] Fusionando capa Dolby Vision con HDR10..." -ForegroundColor Cyan

    $actividad = "Fusión Dolby Vision + HDR10"
    Write-Progress -Activity $actividad -Status "Inicializando..." -PercentComplete 0

    # Duración del video DV (en segundos) — la usamos para calcular % cuando ffmpeg da out_time_ms.
    $duracionTotal = 0.0
    try {
        $dur = ffprobe -v error -show_entries format=duration -of csv=p=0 "$($archivoDV.FullName)" 2>$null
        if ($dur) { $duracionTotal = [double]($dur -replace ',', '.') }
    } catch {}

    # ==== FASE 1: extraer HEVC del archivo DV y aislar RPU (0% → 40%) ====
    Write-Progress -Activity $actividad -Status "Paso 1/3: extrayendo capa Dolby Vision (RPU)..." -PercentComplete 1

    # ffmpeg con -progress pipe:2 emite líneas "out_time_ms=..." a stderr; las parseamos para % en tiempo real.
    # El HEVC se vuelca a un fichero temporal y luego dovi_tool extrae el RPU de ahí.
    $tmpRpuPipe = Join-Path $rutaCarpeta "tmp_dv_raw.hevc"
    $argsFf1 = @(
        "-hide_banner", "-loglevel", "error", "-nostats",
        "-progress", "pipe:2",
        "-i", $archivoDV.FullName,
        "-map", "0:v:0", "-c:v", "copy", "-bsf:v", "hevc_mp4toannexb",
        "-f", "hevc", $tmpRpuPipe, "-y"
    )
    $exit1 = Invoke-ProcesoMonitorizado "ffmpeg" $argsFf1 ({
        param($line)
        if ($line -match '^out_time_ms=(\d+)' -and $duracionTotal -gt 0) {
            $secs = [double]$Matches[1] / 1000000.0
            $pct = [Math]::Min(33, [int](($secs / $duracionTotal) * 33))
            Write-Progress -Activity $actividad -Status "Paso 1/3: extrayendo HEVC del archivo DV..." -PercentComplete $pct
        }
    }.GetNewClosure())
    if ($exit1 -ne 0 -or -not (Test-Path -LiteralPath $tmpRpuPipe)) {
        Write-Progress -Activity $actividad -Completed
        Write-Host "   [!] ERROR: ffmpeg falló extrayendo HEVC del archivo DV." -ForegroundColor Red
        Remove-Item -LiteralPath $tmpRpuPipe -ErrorAction SilentlyContinue
        return $null
    }

    # Ahora dovi_tool sobre el HEVC extraído (proceso rápido, 33% → 40%)
    Write-Progress -Activity $actividad -Status "Paso 1/3: aislando metadata Dolby Vision..." -PercentComplete 34
    dovi_tool -m 3 extract-rpu $tmpRpuPipe -o metadata_p8.rpu 2>&1 | Out-Null
    Remove-Item -LiteralPath $tmpRpuPipe -ErrorAction SilentlyContinue

    if (-not (Test-Path "metadata_p8.rpu") -or (Get-Item "metadata_p8.rpu").Length -lt 100) {
        Write-Progress -Activity $actividad -Completed
        Write-Host "   [!] ERROR: dovi_tool no pudo extraer el RPU del archivo Dolby Vision." -ForegroundColor Red
        Remove-Item "metadata_p8.rpu" -ErrorAction SilentlyContinue
        return $null
    }
    Write-Progress -Activity $actividad -Status "Paso 1/3 completado." -PercentComplete 40

    # ==== FASE 2: extraer base HDR10 (40% → 75%) ====
    Write-Progress -Activity $actividad -Status "Paso 2/3: extrayendo base HDR10..." -PercentComplete 41
    $argsFf2 = @(
        "-hide_banner", "-loglevel", "error", "-nostats",
        "-progress", "pipe:2",
        "-i", $archivoHDR10.FullName,
        "-map", "0:v:0", "-c:v", "copy", "-bsf:v", "hevc_mp4toannexb",
        "video_base.hevc", "-y"
    )
    $exit2 = Invoke-ProcesoMonitorizado "ffmpeg" $argsFf2 ({
        param($line)
        if ($line -match '^out_time_ms=(\d+)' -and $duracionTotal -gt 0) {
            $secs = [double]$Matches[1] / 1000000.0
            $pct = 41 + [Math]::Min(34, [int](($secs / $duracionTotal) * 34))
            Write-Progress -Activity $actividad -Status "Paso 2/3: extrayendo base HDR10..." -PercentComplete $pct
        }
    }.GetNewClosure())
    if ($exit2 -ne 0) {
        Write-Progress -Activity $actividad -Completed
        Write-Host "   [!] ERROR: ffmpeg falló al extraer base HDR10." -ForegroundColor Red
        Remove-Item "metadata_p8.rpu","video_base.hevc" -ErrorAction SilentlyContinue
        return $null
    }
    Write-Progress -Activity $actividad -Status "Paso 2/3 completado." -PercentComplete 75

    # ==== FASE 3: inject-rpu (75% → 100%) ====
    # dovi_tool no expone progreso fiable, así que monitorizamos el tamaño del fichero de salida
    # comparándolo con video_base.hevc para estimar el avance.
    $tamRef = (Get-Item "video_base.hevc").Length
    Write-Progress -Activity $actividad -Status "Paso 3/3: inyectando RPU en base HDR10..." -PercentComplete 76

    $psi3 = New-Object System.Diagnostics.ProcessStartInfo
    $psi3.FileName = "dovi_tool"
    foreach ($a in @("inject-rpu","--rpu-in","metadata_p8.rpu","video_base.hevc","-o","video_definitivo.hevc")) { $psi3.ArgumentList.Add($a) }
    $psi3.RedirectStandardError  = $true
    $psi3.RedirectStandardOutput = $true
    $psi3.UseShellExecute = $false
    $psi3.CreateNoWindow  = $true
    $proc3 = [System.Diagnostics.Process]::Start($psi3)
    # Drenaje en segundo plano de stdout/stderr: si no se leen y dovi_tool escribe más de lo
    # que cabe en el buffer del pipe (~4 KB), el proceso se queda bloqueado para siempre.
    $drenOut3 = $proc3.StandardOutput.ReadToEndAsync()
    $drenErr3 = $proc3.StandardError.ReadToEndAsync()
    while (-not $proc3.HasExited) {
        Start-Sleep -Milliseconds 300
        if (Test-Path "video_definitivo.hevc") {
            $tamAct = (Get-Item "video_definitivo.hevc").Length
            if ($tamRef -gt 0) {
                $pct = 76 + [Math]::Min(23, [int](($tamAct / $tamRef) * 23))
                Write-Progress -Activity $actividad -Status "Paso 3/3: inyectando RPU en base HDR10..." -PercentComplete $pct
            }
        }
    }
    $proc3.WaitForExit()

    if (-not (Test-Path "video_definitivo.hevc") -or (Get-Item "video_definitivo.hevc").Length -lt 1MB) {
        Write-Progress -Activity $actividad -Completed
        Write-Host "`n[!] ERROR CRÍTICO: dovi_tool ha fallado en inject-rpu (posible 'Invalid PPS index')." -ForegroundColor Red
        Remove-Item "metadata_p8.rpu","video_base.hevc","video_definitivo.hevc" -ErrorAction SilentlyContinue
        return $null
    }

    Write-Progress -Activity $actividad -Status "Fusión completada." -PercentComplete 100
    Start-Sleep -Milliseconds 200
    Write-Progress -Activity $actividad -Completed

    return @{
        VideoFusionado = (Join-Path $rutaCarpeta "video_definitivo.hevc")
        Borrables = @("metadata_p8.rpu", "video_base.hevc", "video_definitivo.hevc")
    }
}

# Convierte pistas DTS internas a E-AC3 para cualquier idioma que no tenga ya AC3/EAC3 disponible
# (ni interno ni externo). Una conversión por idioma. Aplicado a TODOS los idiomas presentes
# (incluyendo und si lo hay) para garantizar compatibilidad con reproductores sin soporte DTS.
# $modo: "SIEMPRE" (convierte sin preguntar), "NUNCA" (omite siempre), "PREGUNTAR" (una pregunta por archivo)
function Invoke-ConversionDTSMultiidioma($origenesPistas, $modo = "SIEMPRE", $memoriaIdiomas = @{}) {
    Write-Host "`n[>>] Analizando pistas de audio internas..." -ForegroundColor Cyan
    Write-DiagPaso "Conversion DTS: INICIO"

    if ($modo -eq "NUNCA") {
        Write-Host "   -> Conversión DTS desactivada por configuración. Se conserva el audio original." -ForegroundColor DarkGray
        return @()
    }

    $listaArchivos = @($origenesPistas)

    # 1. Listar todas las pistas de audio internas de TODOS los archivos del par.
    # Salida JSON en lugar de CSV: con CSV, si una pista no tenía tag 'language' pero sí
    # 'title', las columnas se desplazaban (el título acababa en Lang); y un título con
    # comas rompía el registro entero.
    $pistasInternas = @()
    foreach ($rutaArchivo in $listaArchivos) {
        $probeAud = $null
        try { $probeAud = ffprobe -v error -select_streams a -show_entries stream=index,codec_name:stream_tags=language,title -of json $rutaArchivo 2>$null | Out-String | ConvertFrom-Json } catch {}
        if (-not $probeAud -or -not $probeAud.streams) { continue }
        $crudas = @($probeAud.streams | ForEach-Object {
            [PSCustomObject]@{
                Index = "$($_.index)"
                Codec = "$($_.codec_name)"
                Lang  = if ($_.tags -and $_.tags.language) { "$($_.tags.language)" } else { "" }
                Title = if ($_.tags -and $_.tags.title)    { "$($_.tags.title)" }    else { "" }
            }
        })
        foreach ($p in $crudas) {
            # Añadimos el path del archivo a cada pista para que la conversión sepa de dónde sacarla
            $p | Add-Member -NotePropertyName "RutaArchivo" -NotePropertyValue $rutaArchivo -Force
            $pistasInternas += $p
        }
    }
    if ($pistasInternas.Count -eq 0) { return @() }

    # Resolver idiomas indeterminados consultando la memoria (Resolve-IdiomasAudios ya las pobló)
    # y normalizar variantes (Castellano vs Latino) usando el Title como pista.
    foreach ($p in $pistasInternas) {
        if ([string]::IsNullOrWhiteSpace($p.Lang) -or $p.Lang -eq "und") {
            $idx = [int]$p.Index   # ojo: aquí "Index" es el índice global de ffmpeg (= ID de mkvmerge)
            $llaveMem = "Interno_$($p.RutaArchivo)_$idx"
            if ($memoriaIdiomas.ContainsKey($llaveMem)) { $p.Lang = $memoriaIdiomas[$llaveMem] }
            elseif ([string]::IsNullOrWhiteSpace($p.Lang)) { $p.Lang = "und" }
        }
        # Normalización: spa+title="Latinoamérica" → es-419
        $p.Lang = Get-LanguageCode $p.Lang $p.Title
    }

    # 2. Agrupar por idioma. Para cada idioma, ver si tiene AC3/EAC3 (interno o externo del mismo idioma).
    $idiomasConAC3 = @{}   # langCode -> $true si ya hay AC3/EAC3
    foreach ($p in $pistasInternas) {
        if ($p.Codec -match "(?i)ac3|eac3") {
            $idiomasConAC3[$p.Lang] = $true
        }
    }
    # Externos AC3/EAC3 con idioma identificado
    $externosAC3 = Get-ChildItem -LiteralPath $rutaCarpeta -File | Where-Object { $_.Extension -match "\.(ac3|eac3)$" }
    foreach ($ext in $externosAC3) {
        $langExt = $null
        $titleExt = ""
        $llaveExt = "Externo_$($ext.FullName)"
        if ($memoriaIdiomas.ContainsKey($llaveExt)) { $langExt = $memoriaIdiomas[$llaveExt] }
        if (-not $langExt) {
            if ($ext.Name -match "(?i)audio_([a-z]{2,3}(?:-[a-z0-9]{2,3})?)_(temp|final)") { $langExt = $Matches[1].ToLower() }
            elseif ($ext.Name -match "(?i)(?<![a-z])(latino|latam|lat)(?![a-z])")           { $langExt = "es-419" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(spa|es|castellano|cast)(?![a-z])")     { $langExt = "es" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(eng|en|ingles|english)(?![a-z])")      { $langExt = "eng" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(fre|fr|frances|french)(?![a-z])")      { $langExt = "fre" }
        }
        $miExt = Get-MediainfoJson $ext.FullName
        if ($miExt) {
            $pAExt = $miExt.media.track | Where-Object { $_.'@type' -eq "Audio" } | Select-Object -First 1
            if ($pAExt) { $titleExt = "$($pAExt.Title)" }
        }
        if ($langExt) { $langExt = Get-LanguageCode $langExt $titleExt }
        if ($langExt) { $idiomasConAC3[$langExt] = $true }
    }

    # 3. Para cada idioma con DTS pero sin AC3/EAC3, elegir la mejor pista DTS y marcarla para conversión.
    # Si el mismo idioma tiene DTS en varios archivos, elegimos la mejor (Master Audio/HD prevalece).
    $aConvertir = @()
    $idiomasDTS = @{}
    foreach ($p in $pistasInternas) {
        if ($p.Codec -match "(?i)dts") {
            $lng = if ([string]::IsNullOrWhiteSpace($p.Lang)) { "und" } else { $p.Lang }
            if (-not $idiomasDTS.ContainsKey($lng)) { $idiomasDTS[$lng] = @() }
            $idiomasDTS[$lng] += $p
        }
    }

    foreach ($lng in $idiomasDTS.Keys) {
        if ($idiomasConAC3.ContainsKey($lng)) {
            Write-Host "   -> Idioma '$lng' ya tiene AC3/EAC3. Se omite conversión DTS." -ForegroundColor DarkGray
            continue
        }
        $mejorDTS = $idiomasDTS[$lng] | Sort-Object { $_.Title -like "*Master*" -or $_.Title -like "*HD*" } -Descending | Select-Object -First 1
        $aConvertir += [PSCustomObject]@{
            Lang  = $lng
            Index = $mejorDTS.Index
            Title = if ([string]::IsNullOrWhiteSpace($mejorDTS.Title)) { "(sin título)" } else { $mejorDTS.Title }
            RutaArchivo = $mejorDTS.RutaArchivo   # de qué MKV concreto extraer
        }
    }

    if ($aConvertir.Count -eq 0) {
        Write-Host "   -> No hay pistas DTS que requieran conversión." -ForegroundColor DarkGray
        return @()
    }

    # 4. Modo PREGUNTAR: una sola pregunta listando todos los idiomas a convertir
    if ($modo -eq "PREGUNTAR") {
        Write-Host "   -> Pistas DTS detectadas para conversión a E-AC3 (compatibilidad):" -ForegroundColor Yellow
        foreach ($conv in $aConvertir) {
            Write-Host "      - $($conv.Lang) (Index $($conv.Index), de $(Split-Path $conv.RutaArchivo -Leaf)) - $($conv.Title)" -ForegroundColor Cyan
        }
        $respDTS = Read-Host "   [?] ¿Convertir todas a E-AC3? (S/N)"
        if ($respDTS -notmatch "^[sS]") {
            Write-Host "   -> Omitidas todas las conversiones por elección del usuario." -ForegroundColor DarkGray
            return @()
        }
    }

    # 5. Convertir cada pista, generando un EAC3 temporal por idioma.
    # IMPORTANTE: preservamos el delay de origen. Algunas pistas DTS no empiezan en 0 (tienen un
    # start_time/retardo para sincronizar con el vídeo, p.ej. 0.971s). Al extraer a un EAC3 crudo
    # ese delay se perdería y el audio quedaría adelantado. Leemos el start_time de la pista y lo
    # reaplicamos en la propia conversión con -af adelay, de modo que el EAC3 ya nace sincronizado.
    $generados = @()
    foreach ($conv in $aConvertir) {
        # Nombre intuitivo: "<nombre del video>_audio_<lang>.eac3" (antes: "audio_<lang>_temp.eac3").
        $baseNombre = [System.IO.Path]::GetFileNameWithoutExtension((Split-Path $conv.RutaArchivo -Leaf))
        $audioEAC3_temp = Join-Path $rutaCarpeta "${baseNombre}_audio_$($conv.Lang).eac3"

        # Leer el delay (start_time) de la pista DTS de origen, en segundos.
        $startTimeRaw = ffprobe -v error -select_streams "$($conv.Index)" -show_entries stream=start_time -of csv=p=0 $conv.RutaArchivo 2>$null
        $delaySeg = 0.0
        [double]::TryParse(("$startTimeRaw").Replace(',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$delaySeg) | Out-Null
        $delayMs = [int][Math]::Round($delaySeg * 1000)

        Write-Host "   -> Extrayendo DTS '$($conv.Lang)' (Index $($conv.Index)) a $($audioEAC3_temp | Split-Path -Leaf)..." -ForegroundColor Gray
        Write-DiagPaso "  ffmpeg DTS->EAC3 INICIO: lang=$($conv.Lang) delay=${delayMs}ms"
        if ($delayMs -gt 0) {
            Write-Host "      (delay de origen: ${delayMs} ms, se preserva en el EAC3)" -ForegroundColor DarkGray
            # adelay inserta silencio al inicio para reproducir el retardo original (todos los canales).
            ffmpeg -i $conv.RutaArchivo -map "0:$($conv.Index)" -af "adelay=${delayMs}:all=1" -c:a eac3 -b:a 1024k $audioEAC3_temp -y -v error
        } else {
            ffmpeg -i $conv.RutaArchivo -map "0:$($conv.Index)" -c:a eac3 -b:a 1024k $audioEAC3_temp -y -v error
        }
        Write-DiagPaso "  ffmpeg DTS->EAC3 FIN: lang=$($conv.Lang) exit=$LASTEXITCODE"
        if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $audioEAC3_temp)) {
            $generados += $audioEAC3_temp
        } else {
            Write-Host "   [!] ffmpeg falló convirtiendo DTS '$($conv.Lang)'." -ForegroundColor Red
        }
    }
    Write-DiagPaso "Conversion DTS: FIN ($($generados.Count) pistas convertidas)"
    return $generados
}

# Extrae pistas PGS para OCR manual. NO hace pausa ni preguntas (se gestiona fuera).
function Invoke-ExtraccionPGS($origenPistas, $idiomasSubsMantener = $null, $prefijoArchivo = "") {
    # Extrae los PGS de un archivo a ficheros .sup. Devuelve el nº de PGS extraídos.
    # Nombre intuitivo: "<nombre del video>_sub_<idx>_<lang>.sup". El nombre del vídeo permite
    # identificar de un vistazo a qué archivo pertenece cada .sup, y evita mezclar los de
    # distintos archivos cuando se extrae todo el lote de golpe (modo serie).
    # Si no se pasa $prefijoArchivo, lo derivamos del propio archivo de origen.
    if (-not $prefijoArchivo) {
        $prefijoArchivo = [System.IO.Path]::GetFileNameWithoutExtension((Split-Path $origenPistas -Leaf))
    }
    Write-Host "   -> Extrayendo PGS de: $(Split-Path $origenPistas -Leaf)" -ForegroundColor Gray
    # Salida JSON en lugar de CSV (mismo motivo que en la conversión DTS: los tags ausentes
    # desplazaban columnas y las comas rompían el parseo).
    $probeSubs = $null
    try { $probeSubs = ffprobe -v error -select_streams s -show_entries stream=index,codec_name:stream_tags=language -of json $origenPistas 2>$null | Out-String | ConvertFrom-Json } catch {}
    $pistasPGS = @()
    if ($probeSubs -and $probeSubs.streams) {
        $pistasPGS = @($probeSubs.streams | Where-Object { "$($_.codec_name)" -match "(?i)pgs" } | ForEach-Object {
            [PSCustomObject]@{
                Index = $_.index
                Codec = "$($_.codec_name)"
                Lang  = if ($_.tags -and $_.tags.language) { "$($_.tags.language)" } else { "" }
            }
        })
    }
    if ($pistasPGS.Count -eq 0) {
        return 0
    }
    $extraidos = 0
    foreach ($sub in $pistasPGS) {
        $lang = if ($sub.Lang) { $sub.Lang } else { "unk" }
        # Si el usuario eligió recortar idiomas, no extraemos los PGS de idiomas descartados.
        if ($idiomasSubsMantener) {
            $codNorm = Get-LanguageCode $lang ""
            $mantener = if ($idiomasSubsMantener -contains "__CAST_ENG__") { $codNorm -in @("es", "eng") } else { $codNorm -in $idiomasSubsMantener }
            if (-not $mantener) {
                Write-Host "      -> Omitido (idioma no seleccionado): $lang" -ForegroundColor DarkGray
                continue
            }
        }
        $supName = "${prefijoArchivo}_sub_$($sub.Index)_$lang.sup"
        ffmpeg -i $origenPistas -map 0:$($sub.Index) -c copy $supName -y -v error
        Write-Host "      -> Extraído: $supName" -ForegroundColor Gray
        $extraidos++
    }
    return $extraidos
}

# Construye la lista bruta de pistas (internas + externas) con todos sus metadatos
function Build-PistasBrutas($origenesPistas, $esHibrido, $memoriaIdiomas, $decisionPGS = "CONSERVAR_PGS", $idiomasSubsMantener = $null, $prefijoSrtPropio = "") {
    $pistasBrutas = @()
    $idVideo = 0; $altoVideo = 0; $codecVideo = "264"; $hdrVideo = ""

    # Aceptamos un único path o un array. Internamente trabajamos siempre con array.
    $listaArchivos = @($origenesPistas)

    # El video se lee SOLO del primer archivo (el preferente / mejor metadata).
    # Las pistas de audio/subs se leen de TODOS los archivos y luego se deduplican.
    $primerArchivo = $listaArchivos[0]

    for ($idxArchivo = 0; $idxArchivo -lt $listaArchivos.Count; $idxArchivo++) {
        $rutaArchivo = $listaArchivos[$idxArchivo]
        $miJson = Get-MediainfoJson $rutaArchivo
        if (-not $miJson) { continue }

        # Mapeo de IDs reales de mkvmerge por (tipo + orden de aparición). mediainfo a veces
        # numera pistas con huecos (visto en Fight Club: PGS internos como 7, 9, 11, 13, 15...
        # cuando mkvmerge espera 6, 7, 8, 9, 10). El pista.ID - 1 funciona en archivos típicos
        # pero falla en estos casos. mkvmerge -J es la fuente fiable.
        $idMap = Get-MkvmergeTrackIdMap $rutaArchivo
        $contIdx = @{ Video = 0; Audio = 0; Text = 0 }

        foreach ($pista in $miJson.media.track) {
            # Video solo del primer archivo
            if ($pista.'@type' -eq "Video" -and $rutaArchivo -eq $primerArchivo) {
                $contIdx.Video++
                $claveMap = "Video+$($contIdx.Video)"
                $idVideo = if ($idMap.ContainsKey($claveMap)) { $idMap[$claveMap] } else { $pista.ID - 1 }
                $altoVideo = [int]$pista.Height
                $codecVideo = if     ($pista.Format -match "(?i)AV1")              { "AV1" }
                              elseif ($pista.Format -match "(?i)HEVC|H\.?265")     { "265" }
                              elseif ($pista.Format -match "(?i)AVC|H\.?264")      { "264" }
                              else {
                                  Write-Host "   [!] Códec de video no reconocido: '$($pista.Format)' — usando H.264 por defecto" -ForegroundColor DarkYellow
                                  "264"
                              }
                $vidStr = $pista | ConvertTo-Json -Depth 2 -Compress
                $arrH = @()
                if ($vidStr -match "Dolby Vision" -or $esHibrido) { $arrH += "DV" }
                if ($vidStr -match "HDR10\+") { $arrH += "HDR10+" } elseif ($vidStr -match "HDR10") { $arrH += "HDR10" }
                # Fallback: detección por BT.2020 + PQ + >=10 bits cuando los campos HDR_Format
                # de mediainfo vienen vacíos (típico en encodes AV1 hechos con fastflix u otros).
                if (-not ($arrH -contains "HDR10") -and -not ($arrH -contains "HDR10+")) {
                    $primarias = "$($pista.colour_primaries)"
                    $transferencia = "$($pista.transfer_characteristics)"
                    $profBits = 0
                    if ($pista.BitDepth) { $profBits = [int]"$($pista.BitDepth)" }
                    if ($primarias -match "(?i)BT\.?2020" -and $transferencia -match "(?i)PQ|SMPTE ST 2084|ST 2084" -and $profBits -ge 10) {
                        $arrH += "HDR10"
                    }
                }
                $hdrVideo = $arrH -join " "
            }

            if ($pista.'@type' -match "Audio|Text") {
                # ID real en mkvmerge: contamos por tipo y consultamos el map.
                $tipoClave = if ($pista.'@type' -eq "Text") { "Text" } else { "Audio" }
                $contIdx[$tipoClave]++
                $claveMap = "$tipoClave+$($contIdx[$tipoClave])"
                $idMkvmerge = if ($idMap.ContainsKey($claveMap)) { $idMap[$claveMap] } else { $pista.ID - 1 }

                $langCode = $pista.Language
                $tituloPista = "$($pista.Title)"
                # D4: clave incluye path. Usamos el ID de mkvmerge para coherencia con el resto del flujo.
                $llaveMemoria = "Interno_${rutaArchivo}_$idMkvmerge"
                if ($memoriaIdiomas.ContainsKey($llaveMemoria)) { $langCode = $memoriaIdiomas[$llaveMemoria] }
                if ([string]::IsNullOrWhiteSpace($langCode)) { $langCode = "und" }
                # Normalizamos: si es español + title indica Latino → "es-419"
                $langCode = Get-LanguageCode $langCode $tituloPista

                # Si la pista quedó como 'und', aplicamos el idioma que el usuario eligió en el
                # pre-escaneo del lote (uno global para audios, otro para subs). Para subs, además,
                # se respeta la decisión por archivo concreto si existe (clave "<archivo>|Text").
                if ($langCode -eq "und") {
                    if ($pista.'@type' -eq "Text") {
                        # La decisión de idioma se guardó con el formato REAL del sub ("Text" o
                        # "PGS", antes se buscaba siempre "|Text" y los PGS nunca la encontraban)
                        # y puede estar registrada con el path de este archivo o con el del
                        # principal del par híbrido (donde se hizo la pregunta).
                        $fmtSimpleUnd = if ("$($pista.Format)" -match "(?i)srt|subrip|ass|ssa|utf-?8|text") { "Text" } else { "PGS" }
                        $clavesUnd = @("${rutaArchivo}|$fmtSimpleUnd")
                        if ($rutaArchivo -ne $primerArchivo) { $clavesUnd += "${primerArchivo}|$fmtSimpleUnd" }
                        $undResuelto = $false
                        foreach ($kUnd in $clavesUnd) {
                            if ($global:decisionesIdiomaUnd.ContainsKey($kUnd)) {
                                $langCode = $global:decisionesIdiomaUnd[$kUnd]
                                $undResuelto = $true
                                break
                            }
                        }
                        if (-not $undResuelto -and $global:idiomaUndSubLote) {
                            $langCode = $global:idiomaUndSubLote
                        }
                    } elseif ($pista.'@type' -eq "Audio") {
                        if ($global:idiomaUndAudioLote) {
                            $langCode = $global:idiomaUndAudioLote
                        }
                    }
                }
                $lNom = Get-LanguageName $langCode $tituloPista

                # Bug 8: capturar flag de mediainfo si existe
                $forcedFlag = $false
                if ($pista.PSObject.Properties.Name -contains "Forced") {
                    if ("$($pista.Forced)" -match "(?i)yes|1|true") { $forcedFlag = $true }
                }
                # Heurística adicional: si el Title de la pista incluye palabras tipo "forced",
                # "forzado", "forzados" o variantes con guiones (p.ej. "es-ES--forced--"),
                # marcamos como forzado. Algunos releases (Disney+, Apple TV) usan Title pero
                # dejan el flag Forced=No, y sin esto se nos escapan.
                if ($tituloPista -match "(?i)(forced|forzado|forzados)") { $forcedFlag = $true }

                $pistasBrutas += [PSCustomObject]@{
                    Origen = "Interno"; ID = $idMkvmerge
                    ArchivoFuenteMkv = $rutaArchivo   # nuevo: de qué MKV viene esta pista
                    Tipo = if ($pista.'@type' -eq "Text") { "Sub" } else { "Audio" }
                    CodLang = $langCode; NomLang = $lNom; PesoLang = Get-PesoIdioma $lNom
                    Format = $pista.Format; Profile = $pista.Format_Profile
                    Comm = $pista.Format_Commercial_IfAny; Chan = $pista.Channels
                    Size = if ($pista.StreamSize) { [long]$pista.StreamSize } else { 0 }
                    IsForced = $false; MediainfoForcedFlag = $forcedFlag; NomFinal = ""
                }
            }
        }
    }

    # ----- DEDUPLICACIÓN entre archivos del par DV+HDR10 -----
    # Diagnóstico previo: listamos TODOS los audios leídos (de todos los archivos) antes de dedup,
    # para detectar si falta algún audio de alguno de los archivos del híbrido.
    foreach ($p in @($pistasBrutas | Where-Object { $_.Tipo -eq "Audio" })) {
        Write-DiagLinea "  [merge audio leido] $($p.NomLang) $($p.Format) $($p.Chan)ch <- $(Split-Path $p.ArchivoFuenteMkv -Leaf)"
    }
    # Audios: clave (idioma + códec_simplificado + canales). Si hay duplicado, conservamos el primero
    # añadido (que viene del archivo preferente, según orden de $listaArchivos).
    # Subs: clave (idioma + formato_simplificado + categoría_tamaño). La categoría distingue subs
    # forzados de completos por tamaño (heurística 35%): así un sub forzado de 3KB y uno completo
    # de 30KB no se consideran duplicados aunque compartan idioma+formato.
    if ($listaArchivos.Count -gt 1) {

        # Pre-paso: medir tamaño real de subs internos sin StreamSize reportado.
        # mediainfo deja vacío StreamSize en muchos SRT internos; sin eso la categorización
        # forzado/completo no funciona. Extraemos a disco solo los que hagan falta.
        $subsAMedir = @($pistasBrutas | Where-Object { $_.Tipo -eq "Sub" -and $_.Origen -eq "Interno" -and $_.Size -le 0 -and $_.ArchivoFuenteMkv })
        if ($subsAMedir.Count -gt 0) {
            Write-Host "   [debug merge] Midiendo tamaño real de $($subsAMedir.Count) sub(s) interno(s) sin StreamSize..." -ForegroundColor DarkGray
            foreach ($s in $subsAMedir) {
                $tam = Measure-SubEnDisco $s.ArchivoFuenteMkv $s.ID
                $s.Size = $tam
            }
        }

        # Calcular categoría forzado/completo para cada sub interno, agrupando por (idioma+formato)
        # para que el "más grande del grupo" sea coherente. Audios no necesitan categoría.
        $gruposSub = @{}   # clave (idioma+fmtSimple) -> lista de subs del grupo
        foreach ($p in $pistasBrutas) {
            if ($p.Tipo -ne "Sub" -or $p.Origen -ne "Interno") { continue }
            $fmtSimple = if ($p.Format -match "(?i)srt|subrip|ass|ssa|utf-?8|text") { "Text" } else { "PGS" }
            $kg = "$($p.CodLang)|$fmtSimple"
            if (-not $gruposSub.ContainsKey($kg)) { $gruposSub[$kg] = @() }
            $gruposSub[$kg] += $p
        }
        # Etiquetamos cada sub con su CategoriaTam: "Forzado" si pesa ≤35% del max del grupo, "Completo" si no
        # (criterio coherente con Set-FlagForzados). Si el grupo tiene un solo sub o todos pesan 0,
        # se etiquetan todos como "Completo" (sin distinción posible).
        foreach ($kg in $gruposSub.Keys) {
            $lista = $gruposSub[$kg]
            $maxS = 0
            foreach ($s in $lista) { if ($s.Size -gt $maxS) { $maxS = $s.Size } }
            foreach ($s in $lista) {
                $cat = if ($maxS -gt 0 -and $s.Size -le ($maxS * 0.35)) { "Forzado" } else { "Completo" }
                $s | Add-Member -NotePropertyName "CategoriaTam" -NotePropertyValue $cat -Force
            }
        }

        $vistasAudio = @{}
        $vistasSub   = @{}
        $deduplicadas = @()
        $descartadas  = 0
        foreach ($p in $pistasBrutas) {
            if ($p.Origen -ne "Interno") { $deduplicadas += $p; continue }
            if ($p.Tipo -eq "Audio") {
                # Códec simplificado: agrupamos las variantes (DTS-HD MA / DTS / etc.) por nombre comercial básico
                $codSimple = "$($p.Format)"
                if ($p.Format -match "(?i)E-?AC-?3" -or $p.Comm -match "(?i)Plus") { $codSimple = "EAC3" }
                elseif ($p.Format -match "(?i)AC-?3")    { $codSimple = "AC3" }
                elseif ($p.Format -match "(?i)TrueHD|MLP") { $codSimple = "TrueHD" }
                elseif ($p.Format -match "(?i)DTS")      { $codSimple = "DTS" }
                $clave = "AUDIO|$($p.CodLang)|$codSimple|$($p.Chan)"
                if ($vistasAudio.ContainsKey($clave)) {
                    Write-Host "   [debug merge] Audio duplicado descartado: $($p.NomLang) $codSimple $($p.Chan)ch (de $(Split-Path $p.ArchivoFuenteMkv -Leaf))" -ForegroundColor DarkGray
                    $descartadas++
                    continue
                }
                $vistasAudio[$clave] = $true
                $deduplicadas += $p
            } elseif ($p.Tipo -eq "Sub") {
                $fmtSimple = if ($p.Format -match "(?i)srt|subrip|ass|ssa|utf-?8|text") { "Text" } else { "PGS" }
                $catTam = if ($p.PSObject.Properties.Name -contains "CategoriaTam") { $p.CategoriaTam } else { "Completo" }
                $clave = "SUB|$($p.CodLang)|$fmtSimple|$catTam"
                if ($vistasSub.ContainsKey($clave)) {
                    Write-Host "   [debug merge] Sub duplicado descartado: $($p.NomLang) $fmtSimple $catTam (de $(Split-Path $p.ArchivoFuenteMkv -Leaf))" -ForegroundColor DarkGray
                    $descartadas++
                    continue
                }
                $vistasSub[$clave] = $true
                $deduplicadas += $p
            } else {
                $deduplicadas += $p
            }
        }
        if ($descartadas -gt 0) {
            Write-Host "   [merge híbrido] Pistas combinadas de $($listaArchivos.Count) archivos; $descartadas duplicadas descartadas." -ForegroundColor Cyan
        } else {
            Write-Host "   [merge híbrido] Pistas combinadas de $($listaArchivos.Count) archivos (sin duplicados)." -ForegroundColor Cyan
        }
        # Diagnóstico: qué audios sobrevivieron a la dedup y de qué archivo vienen.
        foreach ($p in @($deduplicadas | Where-Object { $_.Tipo -eq "Audio" })) {
            Write-DiagLinea "  [merge audio superviviente] $($p.NomLang) $($p.Format) $($p.Chan)ch <- $(Split-Path $p.ArchivoFuenteMkv -Leaf)"
        }
        $pistasBrutas = $deduplicadas
    }

    # Externos: audios sueltos, SRT, y PGS (.sup) que el usuario haya dejado a mano.
    # Incluimos .sup en la lista, pero más abajo distinguimos los EXTRAÍDOS por el script
    # (patrón "<algo>_sub_<N>_<lang>.sup" o "sub_<N>_<lang>.sup") de los que mete el usuario.
    $archivosExternos = Get-ChildItem -LiteralPath $rutaCarpeta -File | Where-Object { $_.Extension -match "\.(ac3|eac3|dts|dtshd|truehd|flac|wav|srt|ass|ssa|sup|idx|sub)$" }

    # Patrón de archivos generados/extraídos por el script desde PGS:
    #   SRT convertidos: "<nombre>_sub_<N>_<lang>.srt"   (o "sub_<N>_<lang>.srt" en versiones viejas)
    #   SUP extraídos:   "<nombre>_sub_<N>_<lang>.sup"   (o "sub_<N>_<lang>.sup")
    $patronSrtPgsGenerico = "(?i)^(.+_)?sub_\d+_[a-z]{2,3}\.srt$"
    $patronSupExtraido    = "(?i)^(.+_)?sub_\d+_[a-z]{2,3}\.sup$"

    # Los .sup EXTRAÍDOS por el script NUNCA se meten como externos: o ya están dentro del MKV
    # como PGS originales (opción mantener), o se van a eliminar (opción solo SRT). Los .sup que
    # el usuario deja a mano (con cualquier OTRO nombre) SÍ se meten en ambas opciones.
    $archivosExternos = @($archivosExternos | Where-Object { $_.Name -notmatch $patronSupExtraido })

    # Parejas VobSub .idx/.sub: al pasarle el .idx, mkvmerge toma su .sub automáticamente.
    # Si metiéramos también el .sub como input separado, la pista saldría duplicada.
    $archivosExternos = @($archivosExternos | Where-Object {
        -not ($_.Extension -match "(?i)^\.sub$" -and (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($_.FullName, ".idx"))))
    })

    # SRT: cada archivo solo recoge SUS propios SRT (los que empiezan por su nombre). Los SRT de
    # otros archivos del lote se descartan aquí para no mezclarlos (modo serie).
    if ($prefijoSrtPropio) {
        $prefijoEsc = [regex]::Escape($prefijoSrtPropio)
        $archivosExternos = @($archivosExternos | Where-Object {
            if ($_.Name -match $patronSrtPgsGenerico) {
                return ($_.Name -match "(?i)^${prefijoEsc}_sub_\d+_[a-z]{2,3}\.srt$")
            }
            return $true
        })
    }
    # NOTA: en AMBAS opciones (CONSERVAR_PGS y ELIMINAR_PGS) los SRT convertidos SÍ entran al MKV.
    # La diferencia entre opciones es si se conservan o no los PGS INTERNOS del MKV, lo cual se
    # gestiona fuera (filtro de pistas internas tipo PGS según $decisionPGS).
    foreach ($ext in $archivosExternos) {
        $miExt = Get-MediainfoJson $ext.FullName
        $pExt = $null
        if ($miExt) { $pExt = $miExt.media.track | Where-Object { $_.'@type' -match "Audio|Text" } | Select-Object -First 1 }

        # Bug 7: detección de tipo más fiable
        $tipoPista = Get-TipoPistaExterno $ext $pExt
        if (-not $tipoPista) { 
            Write-Host "   [!] Externo no clasificable, ignorando: $($ext.Name)" -ForegroundColor DarkYellow
            continue 
        }

        $langCode = if ($pExt) { $pExt.Language } else { "" }
        $llaveMemoria = "Externo_$($ext.FullName)"
        if ($memoriaIdiomas.ContainsKey($llaveMemoria)) { $langCode = $memoriaIdiomas[$llaveMemoria] } 
        elseif ([string]::IsNullOrWhiteSpace($langCode) -or $langCode -eq "und") {
            if     ($ext.Name -match "(?i)(?<![a-z])(latino|latam|lat)(?![a-z])")           { $langCode = "es-419" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(spa|es|castellano|cast)(?![a-z])")     { $langCode = "es" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(eng|en|ingles|english)(?![a-z])")      { $langCode = "eng" }
            elseif ($ext.Name -match "(?i)(?<![a-z])(fre|fr|frances|french)(?![a-z])")      { $langCode = "fre" }
            else { $langCode = "und" }
        }
        # Patrón de los EAC3 generados por el propio script: "<nombre del video>_audio_<lang>.eac3".
        # También soportamos los formatos antiguos por compatibilidad: audio_<lang>_temp/final.eac3.
        # Extraemos el código de idioma (códigos cortos spa/eng o compuestos es-419).
        if ($ext.Name -match "(?i)_audio_([a-z]{2,3}(?:-[a-z0-9]{2,3})?)\.eac3$") { $langCode = $Matches[1].ToLower() }
        elseif ($ext.Name -match "(?i)audio_([a-z]{2,3}(?:-[a-z0-9]{2,3})?)_(temp|final)") { $langCode = $Matches[1].ToLower() }

        # Title del externo (mediainfo). Sirve para detectar Latino si el código es solo "es"
        $tituloExt = if ($pExt) { "$($pExt.Title)" } else { "" }
        $langCode = Get-LanguageCode $langCode $tituloExt

        # Bug 8: flag de mediainfo en externos también
        $forcedFlag = $false
        if ($pExt -and $pExt.PSObject.Properties.Name -contains "Forced") {
            if ("$($pExt.Forced)" -match "(?i)yes|1|true") { $forcedFlag = $true }
        }
        # Heurística adicional: si el nombre del archivo contiene "forced" o "forzado", asumimos forzado real
        if ($ext.Name -match "(?i)\b(forced|forzado|forzados)\b") { $forcedFlag = $true }
        # Y si el Title (de mediainfo) lo indica también
        if ($tituloExt -match "(?i)(forced|forzado|forzados)") { $forcedFlag = $true }

        $pistasBrutas += [PSCustomObject]@{
            Origen = "Externo"; ID = $ext.FullName; Tipo = $tipoPista
            ArchivoFuenteMkv = $null   # los externos no vienen de un MKV
            CodLang = $langCode; NomLang = (Get-LanguageName $langCode $tituloExt); PesoLang = Get-PesoIdioma (Get-LanguageName $langCode $tituloExt)
            Format  = if ($pExt) { $pExt.Format } else { $ext.Extension.Trim('.') }
            Profile = if ($pExt) { $pExt.Format_Profile } else { "" }
            Comm    = if ($pExt) { $pExt.Format_Commercial_IfAny } else { "" }
            Chan    = if ($pExt) { $pExt.Channels } else { "" }
            Size    = if ($pExt -and $pExt.StreamSize) { [long]$pExt.StreamSize } elseif ($ext.Length) { [long]$ext.Length } else { 0 }
            IsForced = $false; MediainfoForcedFlag = $forcedFlag; NomFinal = ""
            NombreExterno = $ext.Name
        }
    }

    # Filtro de idiomas de subtítulos (cuando hay >3 idiomas y el usuario eligió recortar).
    # Se aplica sobre TODAS las pistas de subtítulos (internas + externas), antes de devolver.
    # Quitar un idioma elimina TODO de ese idioma (forzados y completos). Los audios no se tocan.
    if ($idiomasSubsMantener) {
        $antes = @($pistasBrutas | Where-Object { $_.Tipo -eq "Sub" }).Count
        if ($idiomasSubsMantener -contains "__CAST_ENG__") {
            # Mantener solo castellano (es) e inglés (eng). Nota: NO incluye latino (es-419) a propósito.
            $pistasBrutas = @($pistasBrutas | Where-Object {
                $_.Tipo -ne "Sub" -or $_.CodLang -in @("es", "eng")
            })
        } else {
            # Lista explícita de códigos a mantener
            $pistasBrutas = @($pistasBrutas | Where-Object {
                $_.Tipo -ne "Sub" -or $_.CodLang -in $idiomasSubsMantener
            })
        }
        $despues = @($pistasBrutas | Where-Object { $_.Tipo -eq "Sub" }).Count
        if ($antes -ne $despues) {
            Write-Host "   [filtro idiomas subs] Subs: $antes -> $despues (eliminados $($antes - $despues) de idiomas no seleccionados)" -ForegroundColor DarkGray
        }
    }

    return @{
        Pistas = $pistasBrutas
        IdVideo = $idVideo
        AltoVideo = $altoVideo
        CodecVideo = $codecVideo
        HdrVideo = $hdrVideo
    }
}

# Ordena audios por idioma + calidad y prepara su nombre comercial
function Format-Audios($audios) {
    foreach ($a in $audios) {
        $pts = 10
        $esDTS = ($a.Format -match "(?i)DTS")
        $esDolby = ($a.Format -match "(?i)E-?AC-?3|AC-?3|TrueHD|MLP" -or $a.Comm -match "(?i)Dolby|Plus|Atmos")
        # Pts = calidad técnica de la pista (Atmos > TrueHD/MA > DTS > DD+ > DD ...)
        if ($a.Comm -match "(?i)Atmos" -or $a.Profile -match "(?i)X")        { $pts = 100 } 
        elseif ($a.Format -match "(?i)TrueHD" -or $a.Profile -match "(?i)MA|Master Audio") { $pts = 90 } 
        elseif ($esDTS)                                                       { $pts = 70 } 
        elseif ($a.Format -match "(?i)E-?AC-?3" -or $a.Comm -match "(?i)Plus" -or $a.Profile -match "(?i)E-?AC-?3") { $pts = 60 } 
        elseif ($a.Format -match "(?i)AC-?3")                                 { $pts = 50 }
        # PtsCalidad = calidad técnica PURA, antes de la demotion de DTS en castellano/latino.
        # La usa el filtro de híbridos para elegir la mejor pista por idioma: con Pts (demotado)
        # un EAC3 convertido (60) "ganaba" al DTS-HD original castellano (40) y lo descartaba.
        $a | Add-Member -NotePropertyName "PtsCalidad" -NotePropertyValue $pts -Force
        if ($a.PesoLang -le 2 -and $esDTS) { $pts = 40 }
        $a | Add-Member -NotePropertyName "Pts" -NotePropertyValue $pts -Force

        # PesoFamilia = orden por familia de códecs. 1 = Dolby (siempre antes), 2 = DTS, 3 = otros.
        # Esto coloca cualquier audio Dolby por encima de cualquier DTS del mismo idioma,
        # independientemente de Pts. Reglamento del usuario.
        $pesoFamilia = 3
        if     ($esDolby) { $pesoFamilia = 1 }
        elseif ($esDTS)   { $pesoFamilia = 2 }
        $a | Add-Member -NotePropertyName "PesoFamilia" -NotePropertyValue $pesoFamilia -Force

        $cStr = switch ("$($a.Chan)") { "8" {"7.1"} "6" {"5.1"} "2" {"2.0"} "1" {"1.0"} default {"$($a.Chan)"} }
        if (-not $cStr) { $cStr = "2.0" }

        # Nombre largo (para mkvpropedit --set name=): formato bonito con espacios
        # P. ej.: "Castellano DD+ 5.1", "Inglés DTS-HD MA 7.1", "Castellano DD+ 5.1 Atmos"
        $codStr = $a.Format
        if ($a.Format -match "(?i)E-?AC-?3" -or $a.Comm -match "(?i)Plus" -or $a.Profile -match "(?i)E-?AC-?3") { $codStr = "DD+" }
        elseif ($a.Format -match "(?i)AC-?3")    { $codStr = "DD" }
        elseif ($a.Format -match "(?i)TrueHD|MLP") { $codStr = "TrueHD" }
        elseif ($a.Format -match "(?i)DTS") {
            # Solo dos casos: DTS-HD MA (cuando profile o commercial-string contienen MA / Master Audio)
            # o DTS a secas (todo lo demás: Core, X, HR, ES, etc.). Decisión del usuario.
            # Importante: miramos también Comm (Format_Commercial_IfAny) porque mediainfo a veces
            # deja Profile vacío y mete la info en Comm (ej. "DTS-HD MA + DTS:X").
            if ($a.Profile -match "(?i)MA|Master Audio" -or $a.Comm -match "(?i)MA|Master Audio") { $codStr = "DTS-HD MA" }
            else { $codStr = "DTS" }
        }
        elseif ($a.Format -match "(?i)AAC") { $codStr = "AAC" }
        elseif ($a.Format -match "(?i)FLAC") { $codStr = "FLAC" }
        elseif ($a.Format -match "(?i)PCM|LPCM") { $codStr = "LPCM" }
        $atm = if ($a.Comm -match "(?i)Atmos" -or $a.Format -match "(?i)Atmos") { " Atmos" } else { "" }
        $a.NomFinal = "$($a.NomLang) $codStr $cStr$atm"
    }
    # Orden final: 1) idioma (PesoLang), 2) familia Dolby antes que DTS (PesoFamilia), 3) calidad (Pts desc)
    return ,@($audios | Sort-Object PesoLang, PesoFamilia, @{Expression={$_.Pts}; Descending=$true})
}

# Ordena subtítulos detectando forzados con flag de mediainfo + heurística (Bug 8)
function Format-Subs($subs, $rutaMkvOrigen = $null, $decisionesSubUnico = $null) {
    Set-FlagForzados $subs $rutaMkvOrigen $decisionesSubUnico
    $subsF = @()
    foreach ($s in $subs) {
        $tStr = if ($s.IsForced) { "Forzados" } else { "Completos" }
        $fStr = if ($s.Format -match "(?i)srt|subrip|ass|ssa|utf-?8|text") { "Planos" } else { "PGS" }
        $s.NomFinal = "$($s.NomLang) $tStr $fStr"
        $s | Add-Member -NotePropertyName "FmtPeso" -NotePropertyValue (Get-PesoFormatoSub $fStr) -Force
        $subsF += $s
    }
    return ,@($subsF | Sort-Object FmtPeso, PesoLang, @{Expression={$_.IsForced}; Descending=$true})
}

# Ensambla el MKV final con mkvmerge
function Invoke-Ensamblaje($archivoPrincipal, $origenesPistas, $esHibrido, $videoA_Muxear, $idVideo, $audiosOrdenados, $subsOrdenados) {
    $archivoSalidaTemporal = Join-Path $rutaCarpeta "Ensamblaje_$($archivoPrincipal.BaseName).mkv"
    $argsMkv = @("-o", $archivoSalidaTemporal, "--no-global-tags", "--gui-mode")

    $listaArchivos = @($origenesPistas)

    # ID de la pista de vídeo en el archivo que realmente la aporta: en híbrido el input 0 es
    # el HEVC fusionado, cuya ÚNICA pista es siempre la 0 ($idVideo es el ID en el MKV origen
    # y no aplica ahí; solo coincidía cuando el vídeo era casualmente la pista 0 del origen).
    $idVideoSalida = if ($esHibrido) { 0 } else { $idVideo }

    # Pre-cálculo: para cada MKV de origen, qué track-IDs (audio y sub) sobreviven a la deduplicación.
    # mkvmerge incluye TODAS las pistas de cada input por defecto; si no le pasamos -a/-s explícitos,
    # se cuelan pistas duplicadas o no deseadas. Por tanto construimos los filtros aquí.
    $audiosPorMkv = @{}   # path → lista de IDs de audio a incluir
    $subsPorMkv   = @{}   # path → lista de IDs de sub a incluir
    foreach ($mkv in $listaArchivos) {
        $audiosPorMkv[$mkv] = @()
        $subsPorMkv[$mkv]   = @()
    }
    foreach ($a in $audiosOrdenados) {
        if ($a.Origen -eq "Interno" -and $audiosPorMkv.ContainsKey($a.ArchivoFuenteMkv)) {
            $audiosPorMkv[$a.ArchivoFuenteMkv] += $a.ID
        }
    }
    foreach ($s in $subsOrdenados) {
        if ($s.Origen -eq "Interno" -and $subsPorMkv.ContainsKey($s.ArchivoFuenteMkv)) {
            $subsPorMkv[$s.ArchivoFuenteMkv] += $s.ID
        }
    }

    # Mapa: path absoluto → índice de archivo en mkvmerge.
    # En híbrido, el video viene del archivo fusionado (índice 0) y las pistas no-video de los MKV
    # originales (índices 1, 2, ...). En no-híbrido, el primer MKV es índice 0.
    $archMapeados = @{}
    if ($esHibrido) {
        # 0 = video fusionado HEVC
        $archMapeados[$videoA_Muxear] = 0
        $idxFile = 1
        foreach ($mkv in $listaArchivos) {
            if (-not $archMapeados.ContainsKey($mkv)) {
                $archMapeados[$mkv] = $idxFile
                $idxFile++
            }
        }
    } else {
        $archMapeados[$listaArchivos[0]] = 0
        $idxFile = 1
    }

    $trackOrder = @("0:$idVideoSalida")
    $argsMkv += "--default-track-flag", "$($idVideoSalida):1"

    # Construye los argumentos -a/-s/-A/-S para un MKV concreto. Si la lista de IDs está vacía,
    # se usa -A/-S (no incluir ninguna pista de ese tipo). Si tiene IDs, se pasa -a 1,2,3.

    if ($esHibrido) {
        # Bug 5: detectar fps real en lugar de hardcodear 23.976
        # Usamos el primer archivo de la lista (el preferente) como referencia del fps
        $fpsReal = Get-FpsReal $listaArchivos[0]
        if ($fpsReal) {
            $argsMkv += "--default-duration"; $argsMkv += "0:${fpsReal}fps"
            Write-Host "   -> FPS detectado para video fusionado: $fpsReal" -ForegroundColor DarkGray
        } else {
            Write-Host "   [!] No se pudo detectar el fps real; mkvmerge usará el del header." -ForegroundColor DarkYellow
        }
        # Input 0: el video fusionado HEVC (no aporta audios/subs, pero con -A -S por seguridad)
        $argsMkv += "-A"; $argsMkv += "-S"
        $argsMkv += $videoA_Muxear
        # Inputs 1..N: cada MKV de origen sin video, con filtros explícitos de qué audios/subs incluir
        foreach ($mkv in $listaArchivos) {
            $argsMkv += "--no-video"
            $audIds = $audiosPorMkv[$mkv]
            $subIds = $subsPorMkv[$mkv]
            if ($audIds.Count -gt 0) { $argsMkv += "-a"; $argsMkv += ($audIds -join ",") }
            else                     { $argsMkv += "-A" }
            if ($subIds.Count -gt 0) { $argsMkv += "-s"; $argsMkv += ($subIds -join ",") }
            else                     { $argsMkv += "-S" }
            $argsMkv += $mkv
        }
    } else {
        # No híbrido: un solo MKV con todo. También aplicamos filtros por si el script descartó alguna pista.
        $audIds = $audiosPorMkv[$listaArchivos[0]]
        $subIds = $subsPorMkv[$listaArchivos[0]]
        if ($audIds.Count -gt 0) { $argsMkv += "-a"; $argsMkv += ($audIds -join ",") }
        else                     { $argsMkv += "-A" }
        if ($subIds.Count -gt 0) { $argsMkv += "-s"; $argsMkv += ($subIds -join ",") }
        else                     { $argsMkv += "-S" }
        $argsMkv += $listaArchivos[0]
    }

    # Externos (audio/sub) van al final, cada uno como input separado
    $externosAUsar = @($audiosOrdenados) + @($subsOrdenados) | Where-Object { $_.Origen -eq "Externo" }
    foreach ($e in $externosAUsar) {
        if (-not $archMapeados.ContainsKey($e.ID)) { $archMapeados[$e.ID] = $idxFile; $idxFile++ }
        $argsMkv += $e.ID
    }

    # Construir el track-order: para cada pista, su input-index y su track-id dentro del input.
    foreach ($a in $audiosOrdenados) { 
        if ($a.Origen -eq "Interno") {
            $tId = $a.ID
            $llaveA = $a.ArchivoFuenteMkv
        } else {
            $tId = 0
            $llaveA = $a.ID   # path del externo
        }
        $trackOrder += "$($archMapeados[$llaveA]):$tId" 
    }
    foreach ($s in $subsOrdenados) { 
        if ($s.Origen -eq "Interno") {
            $tId = $s.ID
            $llaveS = $s.ArchivoFuenteMkv
        } else {
            $tId = 0
            $llaveS = $s.ID
        }
        $trackOrder += "$($archMapeados[$llaveS]):$tId" 
    }

    $argsMkv += "--track-order"; $argsMkv += ($trackOrder -join ",")

    # Volcado al log del comando completo de mkvmerge ANTES de ejecutarlo.
    # Así, si algo falla a mitad o el archivo queda mal, tenemos la línea de invocación exacta.
    Write-DiagComandoMkvmerge $argsMkv

    Write-Host "   -> Muxeando archivo temporal..." -ForegroundColor Gray
    Write-DiagPaso "mkvmerge: INICIO"
    # Capturamos toda la salida (no solo el progreso) por si mkvmerge falla.
    # Las líneas que coincidan con #GUI#progress se usan para la barra; el resto se acumula
    # como mensajes de mkvmerge para diagnóstico.
    $salidaMkv = New-Object System.Collections.Generic.List[string]
    Set-ProgresoGui2 0 "Ensamblando $($archivoPrincipal.BaseName)"   # arranca la 2ª barra de la GUI
    $ultPctMkv = -1
    & mkvmerge @argsMkv 2>&1 | ForEach-Object {
        $linea = "$_"
        if ($linea -match "#GUI#progress (\d+)%") {
            $porcentaje = [int]$Matches[1]
            Write-Progress -Activity "Ensamblando MKV con mkvmerge" -Status "Completado: $porcentaje%" -PercentComplete $porcentaje
            # 2ª barra de la GUI: solo escribimos cuando cambia el entero (evita E/S de más).
            if ($porcentaje -ne $ultPctMkv) { Set-ProgresoGui2 $porcentaje "Ensamblando $($archivoPrincipal.BaseName)"; $ultPctMkv = $porcentaje }
        } else {
            $salidaMkv.Add($linea)
        }
    }
    $codigoMkv = $LASTEXITCODE
    Write-DiagPaso "mkvmerge: FIN (exit=$codigoMkv)"
    Write-Progress -Activity "Ensamblando MKV con mkvmerge" -Completed
    Set-ProgresoGui2 -1   # oculta la 2ª barra al terminar el ensamblado

    # Volcado al log de la salida completa de mkvmerge (siempre, no solo cuando falla).
    # Incluye warnings que mkvmerge dé aunque haya generado el archivo OK.
    Write-DiagRaw "[mkvmerge] --- Código de salida: $codigoMkv ---"
    Write-DiagRaw "[mkvmerge] --- Salida completa ---"
    foreach ($l in $salidaMkv) {
        if (-not [string]::IsNullOrWhiteSpace($l)) { Write-DiagRaw "[mkvmerge] $l" }
    }

    # Códigos de salida de mkvmerge: 0 = OK, 1 = terminó CON AVISOS (el archivo se genera
    # igualmente y suele ser válido), 2 = error real. Antes el código 1 se trataba como fatal:
    # se descartaba un muxeo correcto y se dejaba el temporal "Ensamblaje_*.mkv" huérfano en
    # la carpeta (que la siguiente ejecución trataría como un vídeo de entrada más).
    if (-not (Test-Path -LiteralPath $archivoSalidaTemporal) -or $codigoMkv -ge 2) {
        Write-Host "`n[!] ERROR CRÍTICO: mkvmerge falló (código $codigoMkv) y no generó el archivo final." -ForegroundColor Red
        Write-Host "   -> Detalles en el log." -ForegroundColor DarkRed
        Remove-Item -LiteralPath $archivoSalidaTemporal -ErrorAction SilentlyContinue
        return $null
    }
    if ($codigoMkv -eq 1) {
        Write-Host "   [!] mkvmerge terminó con avisos (código 1). El archivo se generó; detalles en el log." -ForegroundColor DarkYellow
        Add-Incidencia $global:archivoEnCurso "mkvmerge emitió avisos durante el muxeo (código 1). El archivo se generó igualmente; revisar el log."
    }
    return $archivoSalidaTemporal
}

# Aplica banderas de pista (default/forced/idioma/nombre) al MKV ya muxeado.
# $defaultPreferidoAudio: "DOLBY" o "DTS". Cuando el idioma prioritario tiene ambos,
# el flag default va a la pista de la familia elegida (sin cambiar el orden de pistas).
function Set-Banderas($archivoMkv, $audiosOrdenados, $subsOrdenados, $defaultPreferidoAudio = "DOLBY") {
    # Bug 9: Invoke-Mkvpropedit captura errores
    Invoke-Mkvpropedit @($archivoMkv, "--tags", "all:", "--delete", "title") "borrar título global" | Out-Null
    Invoke-Mkvpropedit @($archivoMkv, "--edit", "track:v1", "--set", "flag-default=1", "--set", "flag-forced=0") "video por defecto" | Out-Null

    # Determinar qué pista de audio (índice 1-based) lleva el flag default.
    # Por defecto es la primera (idioma prioritario + Dolby por orden). Pero si el usuario
    # eligió "DTS" y el idioma prioritario tiene una pista DTS, el default va a esa pista DTS.
    $idxDefault = 1
    if ($audiosOrdenados.Count -gt 0) {
        $pesoMinAud = ($audiosOrdenados | Measure-Object PesoLang -Minimum).Minimum
        if ($defaultPreferidoAudio -eq "DTS") {
            for ($k = 0; $k -lt $audiosOrdenados.Count; $k++) {
                if ($audiosOrdenados[$k].PesoLang -eq $pesoMinAud -and $audiosOrdenados[$k].Format -match "(?i)DTS") {
                    $idxDefault = $k + 1
                    break
                }
            }
        } else {
            # DOLBY: primera pista del idioma prioritario que sea Dolby (normalmente la #1)
            for ($k = 0; $k -lt $audiosOrdenados.Count; $k++) {
                $a = $audiosOrdenados[$k]
                $esDolby = ($a.Format -match "(?i)E-?AC-?3|AC-?3|TrueHD|MLP" -or $a.Comm -match "(?i)Dolby|Plus|Atmos")
                if ($a.PesoLang -eq $pesoMinAud -and $esDolby) {
                    $idxDefault = $k + 1
                    break
                }
            }
        }
    }

    for ($i=1; $i -le $audiosOrdenados.Count; $i++) {
        $isDef = if ($i -eq $idxDefault) { 1 } else { 0 }
        $langVal = if ([string]::IsNullOrWhiteSpace($audiosOrdenados[$i-1].CodLang)) { "und" } else { $audiosOrdenados[$i-1].CodLang }
        Invoke-Mkvpropedit @(
            $archivoMkv, "--edit", "track:a$i",
            "--set", "name=$($audiosOrdenados[$i-1].NomFinal)",
            "--set", "language=$langVal",
            "--set", "flag-default=$isDef"
        ) "audio #$i" | Out-Null
    }

    $forzActivo = $false
    for ($i=1; $i -le $subsOrdenados.Count; $i++) {
        $s = $subsOrdenados[$i-1]
        # flag-forced va en TODOS los subs realmente forzados (es metadato de la pista, también
        # en otros idiomas; antes solo lo recibía el primero y el resto quedaba con forced=0).
        # flag-default solo en el primero: es el que el reproductor debe autoseleccionar.
        $isF = if ($s.IsForced) { 1 } else { 0 }
        $isDef = if ($s.IsForced -and -not $forzActivo) { $forzActivo = $true; 1 } else { 0 }
        $langVal = if ([string]::IsNullOrWhiteSpace($s.CodLang)) { "und" } else { $s.CodLang }
        Invoke-Mkvpropedit @(
            $archivoMkv, "--edit", "track:s$i",
            "--set", "name=$($s.NomFinal)",
            "--set", "language=$langVal",
            "--set", "flag-forced=$isF",
            "--set", "flag-default=$isDef"
        ) "sub #$i" | Out-Null
    }
}

# Genera capturas a partir del MKV final.
# Si el vídeo es HDR ($esHdr=$true), aplica tonemapping para que la captura se vea bien en SDR.
# Si es SDR, hace una captura directa (sin tonemap, que falsearía los colores).
function New-Capturas($archivoFinalDestino, $nombreFinal, $numCapturas, $esHdr = $true) {
    if ($numCapturas -le 0) { return }
    $tipoCaptura = if ($esHdr) { "tonemapped (HDR)" } else { "directas (SDR)" }
    Write-Host "`n[>>] Generando $numCapturas capturas $tipoCaptura..." -ForegroundColor Cyan
    Write-DiagPaso "Capturas: INICIO ($numCapturas, $tipoCaptura)"
    # "$( ... )" protege contra ffprobe devolviendo $null o varias líneas: antes, .Replace()
    # sobre $null lanzaba una excepción que marcaba el archivo como incidencia aunque el MKV
    # ya estuviera terminado y renombrado correctamente.
    $rawDur = "$(ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "$archivoFinalDestino" 2>$null)".Trim()
    $dur = 0.0
    $durValida = (-not [string]::IsNullOrWhiteSpace($rawDur)) -and
                 [double]::TryParse($rawDur.Replace(',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dur) -and
                 $dur -gt 0
    if ($durValida) {
        # Segundo de inicio de la primera captura, según duración:
        #  - <=5 min (300s):  20s fijo (cortos arrancan pronto)
        #  - 5 a 30 min:      escala lineal de 20s (a los 5min) hasta 300s (a los 30min)
        #  - >=30 min (1800s): 300s fijo
        if ($dur -le 300) {
            $inic = 20
        } elseif ($dur -ge 1800) {
            $inic = 300
        } else {
            $inic = 20 + (300 - 20) * ($dur - 300) / (1800 - 300)
        }
        # Margen final para no caer en créditos: 60s en vídeos largos; en cortos un 8% proporcional
        # (60s fijos se comerían casi todo el clip).
        $margenFin = if ($dur -lt 1800) { [Math]::Min(60, $dur * 0.08) } else { 60 }
        $fin = $dur - $margenFin
        # Salvaguarda para clips muy cortos donde el inicio no deja ventana válida.
        if ($fin -le $inic) {
            $inic = $dur * 0.10
            $fin  = $dur * 0.90
        }
        $ventana = $fin - $inic

        # Reparto uniforme: con n>1 dividimos entre (n-1) para que la primera captura caiga en $inic
        # y la última en $fin, esparciendo el resto por igual. Con n=1, una sola en el centro de la ventana.
        for ($i = 0; $i -lt $numCapturas; $i++) {
            if ($numCapturas -eq 1) {
                $t = [int]($inic + $ventana / 2)
            } else {
                $t = [int]($inic + ($ventana * $i / ($numCapturas - 1)))
            }
            $ts = $t.ToString("D4")
            # Las capturas se quedan en la carpeta de trabajo ($rutaCarpeta), aunque el MKV se mueva a
            # otra carpeta de salida. La subida sabe buscarlas también aquí (CarpetaOrigen en el config).
            $rutaCap = Join-Path $rutaCarpeta "$([System.IO.Path]::GetFileNameWithoutExtension($nombreFinal))_cap_$ts.jpg"
            if ($esHdr) {
                # Tonemapping HDR->SDR para que la captura no se vea lavada/oscura
                ffmpeg -noaccurate_seek -ss $t -i "$archivoFinalDestino" -frames:v 1 -vf "zscale=t=linear:npl=250,tonemap=tonemap=reinhard:desat=2,zscale=p=709:t=709:m=709,format=yuv420p" -q:v 2 "$rutaCap" -y -loglevel fatal
            } else {
                # SDR: captura directa, sin tonemap
                ffmpeg -noaccurate_seek -ss $t -i "$archivoFinalDestino" -frames:v 1 -q:v 2 "$rutaCap" -y -loglevel fatal
            }
            Write-Host "   -> Captura en segundo $t" -ForegroundColor Gray
        }
    } else {
        Write-Host "   [!] No se pudo leer la duración del archivo; se omiten las capturas." -ForegroundColor DarkYellow
        Add-Incidencia (Split-Path $archivoFinalDestino -Leaf) "No se pudieron generar capturas (duración no legible con ffprobe)."
    }
}

# Lee mediainfo del archivo MKV final ya muxeado y devuelve los componentes técnicos
# que se van a usar en el nombre del archivo: resolución, HDR detallado, códec de video
# y descripción del audio principal (códec + canales + Atmos opcional).
# Esto garantiza que el nombre refleja exactamente lo que el archivo final contiene,
# en lugar de inferirlo del archivo de origen (que en híbridos DV+HDR10+ puede no
# tener toda la metadata combinada).
function Get-DatosTecnicosMkv($rutaMkv) {
    $resultado = [PSCustomObject]@{
        Resolucion  = "1080p"
        CodecVideo  = "264"        # "264", "265" o "AV1"
        HdrVideo    = ""           # ej: "DV HDR10+", "HDR10", "DV", ...
        AudioPista  = ""           # ej: "DD+ 5.1", "TrueHD 7.1 Atmos", "DTS-HD MA 5.1"
    }

    $json = Get-MediainfoJson $rutaMkv
    if (-not $json) { return $resultado }

    # --- VIDEO ---
    $pistaVideo = $json.media.track | Where-Object { $_.'@type' -eq 'Video' } | Select-Object -First 1
    if ($pistaVideo) {
        # Resolución
        $alto = [int]$pistaVideo.Height
        $resultado.Resolucion = if ($alto -gt 1200) { "2160p" } else { "1080p" }

        # Códec de video
        $fmt = "$($pistaVideo.Format)"
        $resultado.CodecVideo = if     ($fmt -match "(?i)AV1")              { "AV1" }
                                elseif ($fmt -match "(?i)HEVC|H\.?265")     { "265" }
                                elseif ($fmt -match "(?i)AVC|H\.?264")      { "264" }
                                else { "264" }

        # HDR: leemos el campo HDR_Format que en el archivo final ya combina la info de DV + HDR10/HDR10+
        $hdrStr = "$($pistaVideo.HDR_Format) $($pistaVideo.HDR_Format_Compatibility) $($pistaVideo.HDR_Format_String)"
        $arrH = @()
        if ($hdrStr -match "(?i)Dolby Vision") { $arrH += "DV" }
        if     ($hdrStr -match "(?i)HDR10\+|SMPTE ST 2094 App 4") { $arrH += "HDR10+" }
        elseif ($hdrStr -match "(?i)HDR10")                       { $arrH += "HDR10" }

        # Detección secundaria por espacio de color cuando HDR_Format no está disponible.
        # Muchos encodes AV1 (fastflix, etc.) conservan BT.2020 + PQ + 10 bits pero no incluyen
        # los metadatos SMPTE ST 2086 que mediainfo necesita para rellenar HDR_Format.
        # Si vemos BT.2020 + PQ/SMPTE ST 2084 + >=10 bits, es HDR10 estándar.
        if ($arrH.Count -eq 0) {
            $primarias = "$($pistaVideo.colour_primaries)"
            $transferencia = "$($pistaVideo.transfer_characteristics)"
            $profBits = 0
            if ($pistaVideo.BitDepth) { $profBits = [int]"$($pistaVideo.BitDepth)" }
            if ($primarias -match "(?i)BT\.?2020" -and $transferencia -match "(?i)PQ|SMPTE ST 2084|ST 2084" -and $profBits -ge 10) {
                $arrH += "HDR10"
            }
        }

        $resultado.HdrVideo = $arrH -join " "
    }

    # --- AUDIO PRINCIPAL: la primera pista de audio del archivo final (ya viene ordenada por Format-Audios) ---
    $pistaAudio = $json.media.track | Where-Object { $_.'@type' -eq 'Audio' } | Select-Object -First 1
    if ($pistaAudio) {
        # Canales
        $cStr = switch ("$($pistaAudio.Channels)") { "8" {"7.1"} "6" {"5.1"} "2" {"2.0"} "1" {"1.0"} default { "$($pistaAudio.Channels)" } }
        if (-not $cStr) { $cStr = "2.0" }

        # Códec comercial (mismo razonamiento que Format-Audios pero releyendo del final)
        $fmtA   = "$($pistaAudio.Format)"
        $profA  = "$($pistaAudio.Format_Profile)"
        $commA  = "$($pistaAudio.Format_Commercial_IfAny)"
        $featA  = "$($pistaAudio.Format_AdditionalFeatures)"

        $codStr = $fmtA
        if     ($fmtA -match "(?i)E-?AC-?3" -or $commA -match "(?i)Plus" -or $profA -match "(?i)E-?AC-?3") { $codStr = "DD+" }
        elseif ($fmtA -match "(?i)AC-?3")         { $codStr = "DD" }
        elseif ($fmtA -match "(?i)TrueHD|MLP")    { $codStr = "TrueHD" }
        elseif ($fmtA -match "(?i)DTS") {
            # Solo dos casos: DTS-HD MA (cuando profile o commercial-string contienen MA / Master Audio)
            # o DTS a secas (todo lo demás: Core, X, HR, ES, etc.). Decisión del usuario.
            # Importante: miramos también Comm porque mediainfo a veces deja Profile vacío
            # y mete la info en Comm (ej. "DTS-HD MA + DTS:X").
            if ($profA -match "(?i)MA|Master Audio" -or $commA -match "(?i)MA|Master Audio") { $codStr = "DTS-HD MA" }
            else { $codStr = "DTS" }
        }
        elseif ($fmtA -match "(?i)AAC")           { $codStr = "AAC" }
        elseif ($fmtA -match "(?i)FLAC")          { $codStr = "FLAC" }
        elseif ($fmtA -match "(?i)PCM|LPCM")      { $codStr = "LPCM" }

        $atm = if ($commA -match "(?i)Atmos" -or $fmtA -match "(?i)Atmos") { " Atmos" } else { "" }
        $resultado.AudioPista = "$codStr $cStr$atm"
    }

    return $resultado
}

# =========================================================================
# CREACIÓN DE TORRENTS: las funciones (New-TorrentNucleo/Archivo/Pack) viven en el
# módulo compartido HDZ-Torrent.ps1, que también usa la GUI, para no duplicar el
# formato bencode. Se carga aquí.
# =========================================================================
. (Join-Path $PSScriptRoot "HDZ-Torrent.ps1")

# =========================================================================
# FASE 1: ESCANEO INICIAL Y CUESTIONARIO GLOBAL
# =========================================================================
try { Clear-Host } catch {}   # en hosts sin consola interactiva Clear-Host lanza excepción
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "       LÍNEA DE MONTAJE HDZ INICIADA     " -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Log: $rutaLog" -ForegroundColor DarkGray

Write-Host "`nComprobando entorno..." -ForegroundColor DarkGray
# D7: incluimos dovi_tool y mediainfo en el check (mediainfo ya estaba; dovi_tool faltaba)
if (-not (Test-Dependencias @("ffmpeg", "ffprobe", "mediainfo", "mkvmerge", "mkvpropedit", "mkvextract", "dovi_tool"))) {
    Write-Host "Instala las herramientas que faltan en el PATH." -ForegroundColor Yellow
    try { Stop-Transcript | Out-Null } catch {}
    exit
}

# Volcado inicial al log: entorno completo del sistema y contenido de la carpeta.
# No se muestra en pantalla (Write-Verbose en transcript), pero permite diagnosticar
# diferencias entre máquinas o detectar archivos externos inesperados.
Write-DiagEntorno
Write-DiagCarpeta "CARPETA AL INICIO"

$todosLosArchivos = @(Get-ChildItem -Include "*.mp4", "*.mkv", "*.ts" -Recurse -Depth 0 | Where-Object { $_.Name -notmatch "_FINAL" -and $_.Name -notmatch "-HDZ" })
if ($todosLosArchivos.Count -eq 0) {
    # ¿Hay archivos que SÍ existen pero fueron excluidos por tener el sufijo -HDZ?
    # En ese caso avisamos y ofrecemos procesarlos igualmente (reprocesar).
    $archivosHDZ = @(Get-ChildItem -Include "*.mp4", "*.mkv", "*.ts" -Recurse -Depth 0 | Where-Object { $_.Name -match "-HDZ" -and $_.Name -notmatch "_FINAL" })
    if ($archivosHDZ.Count -gt 0) {
        Write-Host "`n[!] No hay archivos nuevos, pero se han encontrado $($archivosHDZ.Count) archivo(s) ya procesado(s) (sufijo -HDZ):" -ForegroundColor Yellow
        foreach ($f in $archivosHDZ) { Write-Host "    - $($f.Name)" -ForegroundColor DarkGray }
        $cfgReproc = Get-CfgGui "ReprocesarHDZ"
        $respHDZ = if ($null -ne $cfgReproc) {
            Write-Host "`n[GUI] Reprocesar archivos -HDZ: $(if ($cfgReproc) { 'Sí' } else { 'No' }) (configuración)." -ForegroundColor DarkGray
            if ($cfgReproc) { "S" } else { "N" }
        } else {
            Read-Host "`n¿Quieres procesarlos igualmente? (S/N)"
        }
        if ($respHDZ -match "^[sS]") {
            $todosLosArchivos = $archivosHDZ
        } else {
            Write-Host "Operación cancelada." -ForegroundColor Red
            try { Stop-Transcript | Out-Null } catch {}
            exit
        }
    } else {
        Write-Host "No hay vídeos para procesar en la carpeta." -ForegroundColor Red
        try { Stop-Transcript | Out-Null } catch {}
        exit
    }
}

# Selección de archivos de la GUI: procesar SOLO los marcados (por nombre de archivo).
$cfgSeleccionGui = Get-CfgGui "ArchivosSeleccionados"
if ($cfgSeleccionGui) {
    $nombresSel = @($cfgSeleccionGui | ForEach-Object { "$_" })
    $antesSel = $todosLosArchivos.Count
    $todosLosArchivos = @($todosLosArchivos | Where-Object { $nombresSel -contains $_.Name })
    Write-Host "[GUI] Selección de archivos: $($todosLosArchivos.Count) de $antesSel (configuración)." -ForegroundColor DarkGray
    if ($todosLosArchivos.Count -eq 0) {
        Write-Host "La selección de la GUI no coincide con ningún archivo de la carpeta." -ForegroundColor Red
        try { Stop-Transcript | Out-Null } catch {}
        exit
    }
}

# Agrupamos archivos ignorando etiquetas para encontrar híbridos
# Agrupa archivos del mismo episodio quitando del nombre los marcadores que pueden diferir entre
# las dos versiones del par DV+HDR10 (capa HDR, códec de audio, etiqueta Atmos, formato de plataforma).
# Es case-insensitive y cubre las variaciones más comunes en releases de Apple TV, iTunes, Disney+, etc.
$normalizadorAgrupacion = '(?i)[\._\s\-\(\)]+(Dolby[\._\s\-]Vision|DolbyVision|HDR10\+|HDR10|HDR|DV|2160p|1080p|UHD|WEB[\._\s\-]?DL|Apple[\._\s\-]?TV|iTunes|Disney\+?|Atmos|TrueHD|DTS[\._\s\-]?HD[\._\s\-]?MA|DTS[\._\s\-]?HD|DTS[\._\s\-]?X|DTS|E[\._\s\-]?AC[\._\s\-]?3|EAC3|AC3|DDP\+?|DD\+|DDP|DD|AAC|FLAC|MA|HRA)(?=[\._\s\-\(\)]|$)'
$grupos = @($todosLosArchivos | Group-Object { ($_.BaseName -replace $normalizadorAgrupacion, "") -replace "\s+", " " })

Write-Host "Detectados $($grupos.Count) archivo(s) para procesar." -ForegroundColor Green
$global:totalGruposGui = $grupos.Count
Set-ProgresoGui 2 "Preparando lote" ""

# --- MODO DE LOTE (homogéneo vs heterogéneo) ---
# Si todos los archivos pertenecen al mismo proyecto (típico: episodios de una temporada),
# las preguntas de título/año/origen/etc. se hacen UNA VEZ. Si cada archivo es distinto
# (películas o episodios sueltos sin relación), se preguntan por cada uno.
$modoLote = "HOMOGENEO"
if ($grupos.Count -gt 1) {
    $cfgModoLote = Get-CfgGui "ModoLote"
    if ("$cfgModoLote" -in @("HOMOGENEO", "HETEROGENEO")) {
        $modoLote = "$cfgModoLote"
        Write-Host "[GUI] Modo de procesamiento: $modoLote (configuración)." -ForegroundColor DarkGray
    } else {
        $modoLote = Mostrar-Menu "Modo de procesamiento" @(
            @{Nombre="Mismo proyecto para todos (típico: temporada de serie)"; Valor="HOMOGENEO"},
            @{Nombre="Cada archivo es distinto (películas/episodios sueltos)"; Valor="HETEROGENEO"}
        )
    }
} else {
    # Solo hay un archivo (o un grupo): mostramos su nombre para que el usuario sepa qué se va a procesar.
    Write-Host "   -> Archivo: $($grupos[0].Group[0].Name)" -ForegroundColor Gray
}

# Datos de proyecto: en modo homogéneo se piden ahora; en heterogéneo, dentro del bucle.
$datosProyectoGlobal = $null
if ($modoLote -eq "HOMOGENEO") {
    $datosProyectoGlobal = Get-DatosProyecto $null
}

# Preguntas que SIEMPRE son globales (no cambian entre archivos)
Write-Host "`n--- OPCIONES GENERALES (aplican a todos) ---" -ForegroundColor Yellow
# La pregunta de PGS se ha movido al bucle: solo se hace si se detectan PGS reales en cada archivo.
$cfgCaps = Get-CfgGui "NumCapturas"
if ($null -ne $cfgCaps) {
    $numCapturas = [int]$cfgCaps
    Write-Host "[GUI] Capturas por archivo: $numCapturas (configuración)." -ForegroundColor DarkGray
} else {
    $numCapturas = Read-Host "Número de capturas finales por archivo (0 para saltar)"
    $numCapturas = if ([int]::TryParse($numCapturas, [ref]$null)) { [int]$numCapturas } else { 0 }
}

$cfgBorrar = Get-CfgGui "BorrarOriginales"
if ($null -ne $cfgBorrar) {
    $borrarOriginales = [bool]$cfgBorrar
    Write-Host "[GUI] Originales al terminar: $(if ($borrarOriginales) { 'borrar' } else { 'conservar (.procesado)' }) (configuración)." -ForegroundColor DarkGray
} else {
    $accionOriginales = Mostrar-Menu "¿Qué hacer con los originales al terminar?" @(@{Nombre="Borrarlos (Recomendado para ahorrar espacio)"; Valor="BORRAR"}, @{Nombre="Conservarlos (Añadir .procesado al nombre)"; Valor="CONSERVAR"})
    $borrarOriginales = ($accionOriginales -eq "BORRAR")
}

# Cambio 5: ¿añadir el sufijo -HDZ al nombre final?
$cfgSufijo = Get-CfgGui "SufijoHDZ"
if ($null -ne $cfgSufijo) {
    $anadirSufijoHDZ = [bool]$cfgSufijo
    Write-Host "[GUI] Sufijo '-HDZ' en el nombre final: $(if ($anadirSufijoHDZ) { 'Sí' } else { 'No' }) (configuración)." -ForegroundColor DarkGray
} else {
    $respSufijo = Read-Host "`n¿Añadir el sufijo '-HDZ' al nombre final? (S/N)"
    $anadirSufijoHDZ = ($respSufijo -match "^[sS]")
}

# Carpetas de salida elegidas en la GUI (vacías = junto a los vídeos de origen).
$carpetaSalidaGui  = "$(Get-CfgGui 'CarpetaSalida')".Trim()
$carpetaTorrentGui = "$(Get-CfgGui 'CarpetaTorrent')".Trim()
if ($carpetaSalidaGui)  { try { [void][System.IO.Directory]::CreateDirectory($carpetaSalidaGui) }  catch {}; Write-Host "[GUI] Carpeta de salida del archivo: $carpetaSalidaGui" -ForegroundColor DarkGray }
if ($carpetaTorrentGui) { try { [void][System.IO.Directory]::CreateDirectory($carpetaTorrentGui) } catch {}; Write-Host "[GUI] Carpeta de salida del torrent: $carpetaTorrentGui" -ForegroundColor DarkGray }

# ¿Crear archivos .torrent? Modos: NO / INDIVIDUAL (uno por archivo) / PACK (uno del
# lote completo, carpeta + multi-file) / AMBOS.
$modoTorrent = "NO"
$torrentAnnounce = ""
$cfgModoTorrent = Get-CfgGui "ModoTorrent"
$cfgTorrentCompat = Get-CfgGui "CrearTorrent"   # compatibilidad con configs antiguas (bool)
if ("$cfgModoTorrent" -in @("NO", "INDIVIDUAL", "PACK", "AMBOS")) {
    $modoTorrent = "$cfgModoTorrent"
    Write-Host "[GUI] Torrents: $modoTorrent (configuración)." -ForegroundColor DarkGray
} elseif ($null -ne $cfgTorrentCompat) {
    $modoTorrent = if ([bool]$cfgTorrentCompat) { "INDIVIDUAL" } else { "NO" }
    Write-Host "[GUI] Torrents: $modoTorrent (configuración)." -ForegroundColor DarkGray
} else {
    $modoTorrent = Mostrar-Menu "¿Crear archivos .torrent de los resultados?" @(
        @{Nombre="No"; Valor="NO"},
        @{Nombre="Sí: uno por archivo"; Valor="INDIVIDUAL"},
        @{Nombre="Sí: un PACK del lote completo (carpeta + torrent multi-archivo)"; Valor="PACK"},
        @{Nombre="Sí: ambos (por archivo y pack)"; Valor="AMBOS"}
    )
}
if ($modoTorrent -ne "NO") {
    $cfgAnn = Get-CfgGui "TorrentAnnounce"
    if ($null -ne $cfgAnn) {
        $torrentAnnounce = "$cfgAnn".Trim()
        Write-Host "[GUI] URL de anuncio tomada de la configuración." -ForegroundColor DarkGray
    } else {
        $torrentAnnounce = Read-Host "   URL de anuncio del tracker (ENTER para omitir)"
        $torrentAnnounce = "$torrentAnnounce".Trim()
    }
}

# Pre-escaneo: ¿hay algún archivo con audio DTS? Si no, ni siquiera preguntamos por la conversión.
Write-Host "`nComprobando si algún archivo tiene audio DTS..." -ForegroundColor DarkGray
$hayDTS = $false
$archivosConDTS = 0
$totalArchivos = ($grupos | ForEach-Object { $_.Group.Count } | Measure-Object -Sum).Sum
$nProc = 0
foreach ($g in $grupos) {
    foreach ($f in $g.Group) {
        $nProc++
        Write-Progress -Activity "Comprobando audio DTS en el lote" -Status "$nProc/$totalArchivos : $($f.Name)" -PercentComplete (($nProc/$totalArchivos)*100)
        Set-ProgresoGui ([int](2 + 2 * $nProc / [Math]::Max(1,$totalArchivos))) "Comprobando audio DTS ($nProc/$totalArchivos)" $f.Name
        $codecs = ffprobe -v error -select_streams a -show_entries stream=codec_name -of csv=p=0 $f.FullName 2>$null
        if ($codecs -match "(?i)dts") { 
            $hayDTS = $true
            $archivosConDTS++
            break
        }
    }
}
Write-Progress -Activity "Comprobando audio DTS en el lote" -Completed

if ($hayDTS) {
    Write-Host "   -> Detectado audio DTS en $archivosConDTS archivo(s) del lote." -ForegroundColor DarkGray
    $cfgDTS = Get-CfgGui "ModoConversionDTS"
    if ("$cfgDTS" -in @("SIEMPRE", "NUNCA", "PREGUNTAR")) {
        $modoConversionDTS = "$cfgDTS"
        Write-Host "   [GUI] Conversión DTS -> E-AC3: $modoConversionDTS (configuración)." -ForegroundColor DarkGray
    } else {
        $modoConversionDTS = Mostrar-Menu "¿Convertir pistas DTS a E-AC3?" @(
            @{Nombre="Sí, siempre que se detecte DTS sin AC3/E-AC3 ya presente"; Valor="SIEMPRE"},
            @{Nombre="No, conservar el audio DTS original"; Valor="NUNCA"},
            @{Nombre="Preguntar para cada archivo"; Valor="PREGUNTAR"}
        )
    }
} else {
    Write-Host "   -> Ningún archivo del lote tiene audio DTS. Saltando pregunta de conversión." -ForegroundColor DarkGray
    $modoConversionDTS = "NUNCA"
}

# Pre-escaneo de pistas con idioma 'und' (audio y subtítulo): si las hay, preguntar UNA vez
# por tipo a qué idioma corresponden. Va el PRIMERO porque resuelve idiomas que el resto de
# pre-escaneos (filtro de idiomas, subs únicos) van a necesitar ya normalizados.
if ($modoLote -eq "HOMOGENEO") {
    Set-ProgresoGui 4 "Detectando idiomas indeterminados (und)…" ""
    Resolve-IdiomasUndLote $grupos
}

# Pre-escaneo de idiomas de subtítulos: si hay >3 idiomas distintos, preguntar qué mantener.
# Va PRIMERO (antes que el forzado/completo) para no preguntar por subs de idiomas que se
# van a descartar. En HOMOGÉNEO se decide una vez para todo el lote.
$idiomasSubsMantenerGlobal = $null
if ($modoLote -eq "HOMOGENEO") {
    Set-ProgresoGui 6 "Analizando idiomas de subtítulos…" ""
    $idiomasSubsMantenerGlobal = Resolve-IdiomasSubsLote $grupos
}

# Pre-escaneo de subs únicos sin señal de forzado: solo aplica al modo HOMOGÉNEO.
# Se ejecuta DESPUÉS del filtro de idiomas y lo respeta: no pregunta por subs de idiomas
# que ya se van a descartar. En HETEROGÉNEO se preguntará en el bucle, archivo por archivo.
$decisionesSubUnico = @{}
if ($modoLote -eq "HOMOGENEO") {
    $decisionesSubUnico = Resolve-SubsUnicosLote $grupos $idiomasSubsMantenerGlobal
}

# Cambio 2a: ¿algún archivo del lote tiene, en su idioma de mayor prioridad, audio Dolby Y DTS
# a la vez? Si es así, preguntamos cuál de los dos lleva el flag default (la pista por defecto).
# El orden de pistas NO cambia (Dolby siempre va primero); solo cambia a qué pista del idioma
# prioritario se le pone default=1. Si no existe la casuística, no preguntamos.
$defaultPreferidoAudio = "DOLBY"   # valor por defecto si no se pregunta
$existeCasoDolbyDTS = $false
# Decisión pre-elegida en la GUI: si está definida, nos saltamos también el escaneo de la
# casuística Dolby+DTS (solo servía para decidir si hacía falta preguntar).
$cfgDefAudio = Get-CfgGui "DefaultPreferidoAudio"
if ("$cfgDefAudio" -in @("DOLBY", "DTS")) {
    $defaultPreferidoAudio = "$cfgDefAudio"
    Write-Host "[GUI] Pista de audio predeterminada: $defaultPreferidoAudio (configuración)." -ForegroundColor DarkGray
} else {
# La conversión DTS->EAC3 puede CREAR un audio Dolby donde antes solo había DTS. Por eso
# evaluamos la casuística sobre el ESTADO FINAL: para cada idioma miramos qué audios quedarán.
# Regla de conversión (igual que Invoke-ConversionDTSMultiidioma): un idioma con DTS y SIN
# AC3/EAC3 previo recibirá un EAC3 convertido si la conversión está activa. En modo PREGUNTAR
# asumimos que se convierte (decisión del usuario).
$conversionActiva = ($modoConversionDTS -eq "SIEMPRE" -or $modoConversionDTS -eq "PREGUNTAR")
foreach ($g in $grupos) {
    $hibridoInfo = Resolve-Hibrido $g
    $rutas = $hibridoInfo.OrigenesPistas
    # Recolectar audios de todos los archivos del grupo, agrupados por idioma
    $porIdioma = @{}  # lang -> @{ PesoLang; TieneDolby; TieneDTS; TieneAC3oEAC3 }
    foreach ($ruta in $rutas) {
        $mi = Get-MediainfoJson $ruta
        if (-not $mi) { continue }
        foreach ($p in $mi.media.track) {
            if ($p.'@type' -ne "Audio") { continue }
            $lang = "$($p.Language)"; if ([string]::IsNullOrWhiteSpace($lang)) { $lang = "und" }
            $lang = Get-LanguageCode $lang "$($p.Title)"
            $nom = Get-LanguageName $lang "$($p.Title)"
            $esDTS = ("$($p.Format)" -match "(?i)DTS")
            $esEAC3oAC3 = ("$($p.Format)" -match "(?i)E-?AC-?3|AC-?3")
            $esDolby = ($esEAC3oAC3 -or "$($p.Format)" -match "(?i)TrueHD|MLP" -or "$($p.Format_Commercial_IfAny)" -match "(?i)Dolby|Plus|Atmos")
            if (-not $porIdioma.ContainsKey($lang)) {
                $porIdioma[$lang] = @{ PesoLang = (Get-PesoIdioma $nom); TieneDolby = $false; TieneDTS = $false; TieneAC3oEAC3 = $false }
            }
            if ($esDolby)     { $porIdioma[$lang].TieneDolby = $true }
            if ($esDTS)       { $porIdioma[$lang].TieneDTS = $true }
            if ($esEAC3oAC3)  { $porIdioma[$lang].TieneAC3oEAC3 = $true }
        }
    }
    if ($porIdioma.Keys.Count -eq 0) { continue }
    # Idioma prioritario del grupo (menor PesoLang)
    $pesoMin = ($porIdioma.Values | ForEach-Object { $_.PesoLang } | Measure-Object -Minimum).Minimum
    foreach ($lang in $porIdioma.Keys) {
        $info = $porIdioma[$lang]
        if ($info.PesoLang -ne $pesoMin) { continue }
        # Estado FINAL de este idioma:
        $tendraDTS = $info.TieneDTS
        # Tendrá Dolby si ya lo tiene, O si se va a crear por conversión (DTS sin AC3/EAC3 previo)
        $tendraDolby = $info.TieneDolby
        if ($conversionActiva -and $info.TieneDTS -and -not $info.TieneAC3oEAC3) { $tendraDolby = $true }
        if ($tendraDolby -and $tendraDTS) { $existeCasoDolbyDTS = $true; break }
    }
    if ($existeCasoDolbyDTS) { break }
}

if ($existeCasoDolbyDTS) {
    $defaultPreferidoAudio = Mostrar-Menu "Se ha detectado audio Dolby y DTS en el mismo idioma. ¿Cuál quieres como pista de audio predeterminada?" @(
        @{Nombre="Dolby (DD+/TrueHD/DD)"; Valor="DOLBY"},
        @{Nombre="DTS (DTS-HD MA/DTS)"; Valor="DTS"}
    )
}
}   # fin del else: sin configuración GUI para la pista predeterminada

# =========================================================================
# PRE-ESCANEO DE PGS (modo HOMOGÉNEO): preguntas al inicio + extracción y pausa únicas.
# Objetivo: responder todo al principio y dejar el script trabajando sin más interrupciones.
# En HETEROGÉNEO esto se hace por archivo dentro del bucle (junto a sus preguntas).
# =========================================================================
$decisionExtraerPGS = $false       # ¿extraer PGS para OCR?
$decisionConservarPGS = "CONSERVAR_PGS"  # tras convertir: conservar PGS o solo SRT
$pgsYaExtraidos = $false           # marca para que el bucle no vuelva a extraer/pausar
if ($modoLote -eq "HOMOGENEO") {
    # ¿Hay PGS en algún archivo del lote?
    $hayPgsEnLote = $false
    foreach ($g in $grupos) {
        $hi = Resolve-Hibrido $g
        foreach ($mkv in $hi.OrigenesPistas) {
            if (Test-TienePGS $mkv) { $hayPgsEnLote = $true; break }
        }
        if ($hayPgsEnLote) { break }
    }

    if ($hayPgsEnLote) {
        $cfgExtraer = Get-CfgGui "ExtraerPGS"
        $rExtraer = if ($null -ne $cfgExtraer) {
            Write-Host "`n[GUI] Extraer PGS para OCR en Subtitle Edit: $(if ($cfgExtraer) { 'Sí' } else { 'No' }) (configuración)." -ForegroundColor DarkGray
            if ($cfgExtraer) { "S" } else { "N" }
        } else {
            Read-Host "`n[?] Hay subtítulos PGS en el lote. ¿Extraerlos para OCR en Subtitle Edit? (S/N)"
        }
        if ($rExtraer -match "^[sS]") {
            $decisionExtraerPGS = $true
            $cfgConsPGS = Get-CfgGui "DecisionConservarPGS"
            if ("$cfgConsPGS" -in @("CONSERVAR_PGS", "ELIMINAR_PGS")) {
                $decisionConservarPGS = "$cfgConsPGS"
                Write-Host "[GUI] Tras convertir los PGS: $(if ($decisionConservarPGS -eq 'CONSERVAR_PGS') { 'mantener PGS + SRT' } else { 'solo SRT (borrar PGS)' }) (configuración)." -ForegroundColor DarkGray
            } else {
                $decisionConservarPGS = Mostrar-Menu "Tras convertir los PGS a SRT, ¿qué hago con ellos?" @(
                    @{Nombre="Mantener PGS + SRT (conserva los PGS del vídeo, añade los SRT de la carpeta y cualquier PGS extra que dejes ahí)"; Valor="CONSERVAR_PGS"},
                    @{Nombre="Mantener solo SRT y borrar PGS (elimina los PGS del vídeo y deja solo los SRT de la carpeta)"; Valor="ELIMINAR_PGS"}
                )
            }
        }
    }
}

Write-Host "`n[+] Configuración guardada. Iniciando cadena de montaje..." -ForegroundColor Green
Write-DiagConfiguracion
# Rutas de los MKV finales de ESTA ejecución (las usa el torrent de PACK al terminar)
$global:archivosFinalesLote = @()
$respuestasIdiomasSubs = @{}   # Memoria global de idiomas para subs (D4: claves con path)
$respuestasIdiomasAudios = @{} # Memoria global de idiomas para audios
# Pre-carga de las decisiones POR PISTA de la GUI: con la clave ya en memoria, ni
# Resolve-IdiomasAudios/Subs ni Build-PistasBrutas vuelven a preguntar por esas pistas.
foreach ($kGui in $global:undPistasGui.Keys) {
    $partesGui = $kGui -split "\|"
    if ($partesGui.Count -lt 3) { continue }
    $llaveGui = "Interno_$($partesGui[0])_$($partesGui[1])"
    if ($partesGui[2] -eq "Audio") { $respuestasIdiomasAudios[$llaveGui] = $global:undPistasGui[$kGui] }
    else                           { $respuestasIdiomasSubs[$llaveGui]   = $global:undPistasGui[$kGui] }
    Write-Host "[GUI] Pista $($partesGui[2]) #$($partesGui[1]) de $(Split-Path $partesGui[0] -Leaf) -> idioma '$($global:undPistasGui[$kGui])' (configuración)." -ForegroundColor DarkGray
}

# Extracción + pausa ÚNICA del lote (HOMOGÉNEO), si el usuario pidió extraer PGS.
# Extraemos los PGS de todos los archivos (con prefijo para no mezclarlos), hacemos UNA sola
# pausa para que el usuario los convierta todos juntos en Subtitle Edit, y luego el bucle
# procesa todo sin más interrupciones.
if ($modoLote -eq "HOMOGENEO" -and $decisionExtraerPGS) {
    Write-Host "`n[>>] Extrayendo subtítulos PGS de todo el lote..." -ForegroundColor Cyan
    $totalExtraidos = 0
    foreach ($g in $grupos) {
        $hi = Resolve-Hibrido $g
        $prefijo = [System.IO.Path]::GetFileNameWithoutExtension($hi.ArchivoPrincipal.Name)
        foreach ($mkv in $hi.OrigenesPistas) {
            if (Test-TienePGS $mkv) {
                $totalExtraidos += Invoke-ExtraccionPGS $mkv $idiomasSubsMantenerGlobal $prefijo
                break  # solo del primer archivo con PGS del par
            }
        }
    }
    if ($totalExtraidos -gt 0) {
        Write-Host "`n=======================================================" -ForegroundColor Magenta
        Write-Host " PAUSA: Pásalos por Subtitle Edit y guarda los .srt en la carpeta en la que tienes el video." -ForegroundColor Yellow
        Write-Host " (Se procesará todo el lote sin más interrupciones al continuar.)" -ForegroundColor Yellow
        Write-Host "=======================================================" -ForegroundColor Magenta
        Read-Host "Presiona ENTER cuando hayas terminado para continuar el montaje..."
    }
    $pgsYaExtraidos = $true
}

# =========================================================================
# FASE BATCH: BUCLE PRINCIPAL DE PROCESAMIENTO
# =========================================================================
foreach ($grupo in $grupos) {

    # D2: try/catch por episodio para que un fallo no aborte todo el script
    try {

        $hibridoInfo = Resolve-Hibrido $grupo
        $esHibrido        = $hibridoInfo.EsHibrido
        $archivoDV        = $hibridoInfo.ArchivoDV
        $archivoHDR10     = $hibridoInfo.ArchivoHDR10
        $archivoPrincipal = $hibridoInfo.ArchivoPrincipal
        $origenPistas     = $hibridoInfo.OrigenPistas       # singular (1er archivo) — compat
        $origenesPistas   = $hibridoInfo.OrigenesPistas     # lista de archivos para mergear pistas

        Write-Host "`n=========================================================" -ForegroundColor Magenta
        Write-Host " PROCESANDO: $($archivoPrincipal.Name)" -ForegroundColor Magenta
        $global:archivoEnCurso = $archivoPrincipal.Name
        $global:idxGrupoGui++
        # Base de progreso de este archivo: el avance corre dentro de su "porción" (1/total).
        $global:pctBaseGui = if ($global:totalGruposGui -gt 0) { [int](100 * ($global:idxGrupoGui - 1) / $global:totalGruposGui) } else { 0 }
        $global:pctPasoGui = if ($global:totalGruposGui -gt 0) { 100.0 / $global:totalGruposGui } else { 100.0 }
        Set-ProgresoGui ([int]($global:pctBaseGui + $global:pctPasoGui * 0.05)) "Procesando $($global:idxGrupoGui)/$($global:totalGruposGui)" $archivoPrincipal.Name
        Write-DiagPaso "INICIO procesado de $($archivoPrincipal.Name)"
        Write-Host "=========================================================" -ForegroundColor Magenta

        if ($esHibrido) { Write-Host "[!] Pareja Híbrida Detectada (DV + HDR10) — pistas combinadas de ambos archivos" -ForegroundColor Green }

        # Diagnóstico: volcamos al log toda la información sobre los archivos de entrada
        # (uno solo en no-híbridos, los dos en híbridos). Es la radiografía del input.
        for ($_i = 0; $_i -lt $origenesPistas.Count; $_i++) {
            $etiqueta = if ($origenesPistas.Count -gt 1) { "ENTRADA $($_i + 1)/$($origenesPistas.Count)" } else { "ENTRADA" }
            Write-DiagArchivoEntrada $origenesPistas[$_i] $etiqueta
        }
        Write-DiagLinea "ArchivoPrincipal: $($archivoPrincipal.FullName)"
        Write-DiagLinea "EsHibrido: $esHibrido"
        Write-DiagLinea "OrigenesPistas: $($origenesPistas -join '  ;  ')"

        # Datos de proyecto: globales en homogéneo, preguntados aquí en heterogéneo
        if ($modoLote -eq "HOMOGENEO") {
            $datosProyecto = $datosProyectoGlobal
        } else {
            $datosProyecto = Get-DatosProyecto $archivoPrincipal.Name
        }
        $tituloPelicula    = $datosProyecto.Titulo
        $anoPelicula       = $datosProyecto.Ano
        $esSerie           = $datosProyecto.EsSerie
        $tipoOrigen        = $datosProyecto.TipoOrigen
        $webTipo           = $datosProyecto.WebTipo
        $plataformaFormato = $datosProyecto.PlataformaFormato
        $etiquetasExtra    = $datosProyecto.EtiquetasExtra

        # Auto-Detección de Episodio
        $serieTag = ""
        if ($esSerie) {
            $serieTag = Detectar-Episodio $archivoPrincipal.Name
            if ($serieTag) {
                Write-Host " -> Capítulo detectado automáticamente: $serieTag" -ForegroundColor Green
            } else {
                $serieTag = Read-Host " -> No pude detectar el capítulo para este archivo. Escríbelo (ej: S01E01)"
            }
        }

        # Pre-carga de mediainfo y mkvmerge -J del archivo (o ambos, si es híbrido). La primera
        # lectura puede ser lenta en archivos grandes (AV1, UHD remuxes), así que la disparamos
        # AQUÍ una sola vez con barra de progreso, en lugar de en silencio dentro de cada
        # Resolve-* posterior. Las siguientes lecturas usan la caché y son instantáneas.
        Set-ProgresoGui ([int]($global:pctBaseGui + $global:pctPasoGui * 0.15)) "Analizando pistas $($global:idxGrupoGui)/$($global:totalGruposGui)" $archivoPrincipal.Name
        Write-DiagPaso "Pre-carga mediainfo+mkvmerge: INICIO"
        $totalArch = $origenesPistas.Count
        $nArch = 0
        foreach ($ruta in $origenesPistas) {
            $nArch++
            Write-Progress -Activity "Analizando archivo" -Status "$nArch/$totalArch : $(Split-Path $ruta -Leaf)" -PercentComplete (($nArch/$totalArch)*100)
            Write-DiagPaso "  mediainfo INICIO: $(Split-Path $ruta -Leaf)"
            Get-MediainfoJson $ruta | Out-Null
            Write-DiagPaso "  mediainfo FIN, mkvmerge -J INICIO: $(Split-Path $ruta -Leaf)"
            Get-MkvmergeJson $ruta | Out-Null
            Write-DiagPaso "  mkvmerge -J FIN: $(Split-Path $ruta -Leaf)"
        }
        Write-Progress -Activity "Analizando archivo" -Completed
        Write-DiagPaso "Pre-carga mediainfo+mkvmerge: FIN"

        # Resolución de conflictos de idioma. Audios primero (los necesita la conversión DTS),
        # subtítulos después. En híbrido recorre AMBOS archivos del par.
        # En HOMOGÉNEO, si el pre-escaneo del lote ya decidió el idioma de los 'und', se aplica
        # sin volver a preguntar (pasamos el idioma global).
        Resolve-IdiomasAudios $origenesPistas $archivoPrincipal $respuestasIdiomasAudios $global:idiomaUndAudioLote
        Resolve-IdiomasSubs   $origenesPistas $archivoPrincipal $respuestasIdiomasSubs   $global:idiomaUndSubLote

        # Idiomas de subtítulos a mantener: VA PRIMERO para no preguntar por subs de idiomas
        # que se van a descartar. En HOMOGÉNEO usamos la decisión global del pre-escaneo;
        # en HETEROGÉNEO preguntamos para este archivo concreto (si tiene >3 idiomas).
        if ($modoLote -eq "HOMOGENEO") {
            $idiomasSubsMantener = $idiomasSubsMantenerGlobal
        } else {
            $idiomasSubsMantener = Resolve-IdiomasSubsArchivo $origenesPistas
        }

        # Resolución de subs únicos sin señal: DESPUÉS del filtro de idiomas (lo respeta).
        # En HETEROGÉNEO se pregunta archivo a archivo aquí. En HOMOGÉNEO solo si alguna
        # combinación quedó marcada como "__PREGUNTAR__" en el pre-escaneo.
        $hayPendientesEnEsteArchivo = $false
        if ($modoLote -eq "HETEROGENEO") {
            $hayPendientesEnEsteArchivo = $true
        } else {
            foreach ($k in $decisionesSubUnico.Keys) {
                if ($k.StartsWith("$($archivoPrincipal.FullName)|") -and $decisionesSubUnico[$k] -eq "__PREGUNTAR__") {
                    $hayPendientesEnEsteArchivo = $true
                    break
                }
            }
        }
        if ($hayPendientesEnEsteArchivo) {
            Resolve-SubsUnicosArchivo $origenesPistas $archivoPrincipal $decisionesSubUnico $idiomasSubsMantener
        }

        # Preguntas de PGS (HETEROGÉNEO): se hacen aquí, junto al resto de preguntas del archivo,
        # ANTES de la pausa. La extracción + pausa ocurren después, en la FASE 2.
        $decisionExtraerPGSHetero = $false
        $decisionConservarPGSHetero = "CONSERVAR_PGS"
        if ($modoLote -eq "HETEROGENEO") {
            $tienePgsEste = $false
            foreach ($mkv in $origenesPistas) { if (Test-TienePGS $mkv) { $tienePgsEste = $true; break } }
            if ($tienePgsEste) {
                $cfgExtraerH = Get-CfgGui "ExtraerPGS"
                $rEx = if ($null -ne $cfgExtraerH) {
                    Write-Host "`n[GUI] Extraer PGS para OCR: $(if ($cfgExtraerH) { 'Sí' } else { 'No' }) (configuración)." -ForegroundColor DarkGray
                    if ($cfgExtraerH) { "S" } else { "N" }
                } else {
                    Read-Host "`n[?] Este archivo tiene subtítulos PGS. ¿Extraerlos para OCR en Subtitle Edit? (S/N)"
                }
                if ($rEx -match "^[sS]") {
                    $decisionExtraerPGSHetero = $true
                    $cfgConsPGSH = Get-CfgGui "DecisionConservarPGS"
                    if ("$cfgConsPGSH" -in @("CONSERVAR_PGS", "ELIMINAR_PGS")) {
                        $decisionConservarPGSHetero = "$cfgConsPGSH"
                        Write-Host "[GUI] Tras convertir los PGS: $(if ($decisionConservarPGSHetero -eq 'CONSERVAR_PGS') { 'mantener PGS + SRT' } else { 'solo SRT (borrar PGS)' }) (configuración)." -ForegroundColor DarkGray
                    } else {
                        $decisionConservarPGSHetero = Mostrar-Menu "Tras convertir los PGS a SRT, ¿qué hago con ellos?" @(
                            @{Nombre="Mantener PGS + SRT (conserva los PGS del vídeo, añade los SRT de la carpeta y cualquier PGS extra que dejes ahí)"; Valor="CONSERVAR_PGS"},
                            @{Nombre="Mantener solo SRT y borrar PGS (elimina los PGS del vídeo y deja solo los SRT de la carpeta)"; Valor="ELIMINAR_PGS"}
                        )
                    }
                }
            }
        }

        # =========================================================================
        # FASE 2: PRE-PROCESAMIENTO (FUSIÓN, EXTRACCIÓN DTS Y PGS)
        # =========================================================================
        $archivosBorrables = @()
        $videoA_Muxear = $archivoPrincipal.FullName

        if ($esHibrido) {
            $resFusion = Invoke-FusionDolbyVision $archivoDV $archivoHDR10
            if (-not $resFusion) { 
                Write-Host "   Saltando este archivo para no detener la cadena..." -ForegroundColor Yellow
                Add-Incidencia $global:archivoEnCurso "Fusión Dolby Vision falló. Archivo SALTADO (no procesado)."
                continue 
            }
            $videoA_Muxear      = $resFusion.VideoFusionado
            $archivosBorrables += $resFusion.Borrables
        }

        $archivosBorrables += Invoke-ConversionDTSMultiidioma $origenesPistas $modoConversionDTS $respuestasIdiomasAudios

        # --- PGS ---
        # HOMOGÉNEO: ya se preguntó y se extrajo todo al inicio (pausa única). Solo recogemos
        # las decisiones globales; no se vuelve a preguntar ni a pausar.
        # HETEROGÉNEO: se pregunta, extrae y pausa para este archivo concreto (junto a sus datos).
        if ($modoLote -eq "HOMOGENEO") {
            $decisionPGS = $decisionConservarPGS
            # Si no se generó ningún SRT para este archivo (el usuario no lo convirtió), no podemos
            # eliminar los PGS o el archivo se quedaría sin esos subs: forzamos conservar.
            if ($decisionPGS -eq "ELIMINAR_PGS") {
                $prefijoEste = [System.IO.Path]::GetFileNameWithoutExtension($archivoPrincipal.Name)
                $prefijoEsteEsc = [regex]::Escape($prefijoEste)
                $srtDeEste = @(Get-ChildItem -LiteralPath $rutaCarpeta -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "(?i)^${prefijoEsteEsc}_sub_\d+_[a-z]{2,3}\.srt$" })
                if ($srtDeEste.Count -eq 0) {
                    Write-Host "`n[i] No se detectaron SRT convertidos; se conservarán los PGS originales de este archivo." -ForegroundColor DarkGray
                    $decisionPGS = "CONSERVAR_PGS"
                }
            }
        } else {
            # HETEROGÉNEO: pregunta + extracción + pausa para este archivo.
            $decisionPGS = "CONSERVAR_PGS"
            $tienePgsAlguno = $false
            foreach ($mkv in $origenesPistas) {
                if (Test-TienePGS $mkv) { $tienePgsAlguno = $true; break }
            }
            if ($tienePgsAlguno -and $decisionExtraerPGSHetero) {
                Write-Host "`n[>>] Extrayendo subtítulos PGS..." -ForegroundColor Cyan
                $totalEx = 0
                # Prefijo basado en el ArchivoPrincipal: en híbridos el MKV con PGS puede ser el
                # OTRO archivo del par, y el resto del flujo (detección de SRT generados, filtro
                # $prefijoSrtPropio) busca los SRT por el nombre del principal. Sin esto, los
                # SRT convertidos por el usuario no se recogían.
                $prefijoHet = [System.IO.Path]::GetFileNameWithoutExtension($archivoPrincipal.Name)
                foreach ($mkv in $origenesPistas) {
                    if (Test-TienePGS $mkv) { $totalEx += Invoke-ExtraccionPGS $mkv $idiomasSubsMantener $prefijoHet; break }
                }
                if ($totalEx -gt 0) {
                    Write-Host "`n=======================================================" -ForegroundColor Magenta
                    Write-Host " PAUSA: Pásalos por Subtitle Edit y guarda los .srt en la carpeta en la que tienes el video." -ForegroundColor Yellow
                    Write-Host "=======================================================" -ForegroundColor Magenta
                    Read-Host "Presiona ENTER cuando hayas terminado para continuar el montaje..."
                    $decisionPGS = $decisionConservarPGSHetero
                    # Si no hay SRT generados para este archivo, forzar conservar
                    $prefHet = [regex]::Escape([System.IO.Path]::GetFileNameWithoutExtension($archivoPrincipal.Name))
                    $srtGen = @(Get-ChildItem -LiteralPath $rutaCarpeta -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "(?i)^${prefHet}_sub_\d+_[a-z]{2,3}\.srt$" })
                    if ($decisionPGS -eq "ELIMINAR_PGS" -and $srtGen.Count -eq 0) {
                        Write-Host "   -> No se detectaron SRT generados; se conservarán los PGS originales." -ForegroundColor DarkGray
                        $decisionPGS = "CONSERVAR_PGS"
                    }
                }
            } elseif (-not $tienePgsAlguno) {
                Write-Host "`n[i] Sin subtítulos PGS en este archivo, paso directo." -ForegroundColor DarkGray
            }
        }

        # =========================================================================
        # FASE 3: LÓGICA REMUX (ENSAMBLAJE DE PISTAS)
        # =========================================================================
        Set-ProgresoGui ([int]($global:pctBaseGui + $global:pctPasoGui * 0.45)) "Ensamblando MKV $($global:idxGrupoGui)/$($global:totalGruposGui)" $archivoPrincipal.Name
        Write-Host "`n[>>] Construyendo Track Order y Ensamblando..." -ForegroundColor Cyan
        Write-DiagPaso "Build-PistasBrutas + Format-* + Set-Banderas: INICIO"

        # Build-PistasBrutas resuelve idiomas tanto de audios como de subs internos/externos.
        # Combinamos las dos memorias para que tenga acceso a todas las respuestas.
        $memoriaCombinada = @{}
        foreach ($k in $respuestasIdiomasAudios.Keys) { $memoriaCombinada[$k] = $respuestasIdiomasAudios[$k] }
        foreach ($k in $respuestasIdiomasSubs.Keys)   { $memoriaCombinada[$k] = $respuestasIdiomasSubs[$k] }
        $prefijoSrtPropio = [System.IO.Path]::GetFileNameWithoutExtension($archivoPrincipal.Name)
        # Preguntar forzado/completo de los subs EXTERNOS únicos (SRT sueltos o convertidos de PGS),
        # ahora que ya están en la carpeta (tras la pausa de Subtitle Edit, si la hubo).
        Resolve-SubsExternosForzados $origenesPistas $prefijoSrtPropio $decisionesSubUnico $idiomasSubsMantener
        $analisis = Build-PistasBrutas $origenesPistas $esHibrido $memoriaCombinada $decisionPGS $idiomasSubsMantener $prefijoSrtPropio
        $idVideo    = $analisis.IdVideo
        $altoVideo  = $analisis.AltoVideo
        $codecVideo = $analisis.CodecVideo
        $hdrVideo   = $analisis.HdrVideo

        $audios = @($analisis.Pistas | Where-Object { $_.Tipo -eq "Audio" })
        # Format-Audios calcula Pts (calidad), PesoFamilia y NomFinal, y ordena. DEBE ir antes del
        # filtro de híbrido, porque ese filtro se basa en Pts (que aquí aún no existe en las pistas).
        $audiosOrdenados = Format-Audios $audios

        # En HÍBRIDOS (par DV + HDR10/HDR10+), cada archivo de origen aporta sus propias pistas de
        # audio, duplicadas por idioma con calidades distintas (p.ej. el DV trae AC-3 y el HDR10+
        # trae E-AC-3 Atmos o AAC). Nos quedamos SOLO con la mejor pista por idioma (mayor Pts).
        # En no-híbridos no se toca: se conservan todas las pistas del idioma.
        if ($esHibrido -and $audiosOrdenados.Count -gt 1) {
            # Índice único por pista para identificarla sin depender de -eq sobre objetos
            # (que en PowerShell no compara por referencia de forma fiable).
            for ($i = 0; $i -lt $audiosOrdenados.Count; $i++) {
                $audiosOrdenados[$i] | Add-Member -NotePropertyName "_idxFiltro" -NotePropertyValue $i -Force
            }
            # Solo compiten las pistas INTERNAS (las duplicadas entre los dos MKV del par).
            # Las EXTERNAS (EAC3 convertidos por el script, audios del usuario) se conservan
            # siempre: si compitieran, el EAC3 convertido podía "ganar" al DTS original
            # castellano (demotado en Pts por la regla de orden) y hacer que se descartara.
            # Se compara por PtsCalidad (calidad técnica pura, sin esa demotion).
            $mejorIdxPorIdioma = @{}   # lang -> @{ Idx; Pts }
            foreach ($a in $audiosOrdenados) {
                if ($a.Origen -ne "Interno") { continue }
                $lng = "$($a.CodLang)"
                $pts = [double]$a.PtsCalidad
                if (-not $mejorIdxPorIdioma.ContainsKey($lng) -or $pts -gt $mejorIdxPorIdioma[$lng].Pts) {
                    $mejorIdxPorIdioma[$lng] = @{ Idx = $a._idxFiltro; Pts = $pts }
                }
            }
            $idxGanadores = @($mejorIdxPorIdioma.Values | ForEach-Object { $_.Idx })
            $audiosFiltrados = @($audiosOrdenados | Where-Object { $_.Origen -ne "Interno" -or $idxGanadores -contains $_._idxFiltro })
            $descartados = $audiosOrdenados.Count - $audiosFiltrados.Count
            if ($descartados -gt 0) {
                Write-Host "   [híbrido] Audios redundantes descartados: $descartados (se conserva el mejor por idioma)." -ForegroundColor DarkGray
                foreach ($a in $audiosFiltrados) {
                    Write-DiagLinea "  [hibrido mejor audio] $($a.NomLang): $($a.NomFinal) ($($a.Pts) pts) <- $(Split-Path $a.ArchivoFuenteMkv -Leaf)"
                }
            }
            $audiosOrdenados = $audiosFiltrados
        }


        $subs = @($analisis.Pistas | Where-Object { $_.Tipo -eq "Sub" })

        # Si el usuario eligió eliminar los PGS originales tras convertirlos a SRT,
        # filtramos aquí las pistas internas en formato PGS. Los SRT generados (que ya están
        # en la carpeta como subs externos) entran al archivo final en su lugar.
        if ($decisionPGS -eq "ELIMINAR_PGS") {
            $antesCount = $subs.Count
            $subs = @($subs | Where-Object {
                -not ($_.Origen -eq "Interno" -and "$($_.Format)" -match "(?i)pgs")
            })
            $eliminados = $antesCount - $subs.Count
            if ($eliminados -gt 0) {
                Write-Host "   [decisión PGS] Eliminados $eliminados sub(s) PGS originales (se conservan solo los SRT convertidos)." -ForegroundColor DarkGray
            }
        }

        $subsOrdenados = Format-Subs $subs $origenPistas $decisionesSubUnico

        # Volcado al log de las pistas que vamos a muxear (post-resolución/dedup/forzados/orden).
        # Es lo que le pasaremos a mkvmerge.
        Write-DiagPistasFinal $audiosOrdenados $subsOrdenados $idVideo $altoVideo $codecVideo $hdrVideo

        $archivoSalidaTemporal = Invoke-Ensamblaje $archivoPrincipal $origenesPistas $esHibrido $videoA_Muxear $idVideo $audiosOrdenados $subsOrdenados
        if (-not $archivoSalidaTemporal) {
            Write-Host "   Limpiando y saltando al siguiente archivo..." -ForegroundColor Yellow
            Add-Incidencia $global:archivoEnCurso "mkvmerge falló al ensamblar. Archivo SALTADO (no procesado)."
            foreach ($del in $archivosBorrables) { Remove-Item -LiteralPath $del -ErrorAction SilentlyContinue }
            continue
        }

        # =========================================================================
        # FASE 4: ETIQUETADO Y RENOMBRADO HDZ
        # =========================================================================
        Write-Host "`n[>>] Aplicando banderas y calculando nombre final..." -ForegroundColor Cyan
        Write-DiagPaso "mkvpropedit + calculo nombre: INICIO"
        Set-Banderas $archivoSalidaTemporal $audiosOrdenados $subsOrdenados $defaultPreferidoAudio

        # Leemos los datos técnicos del archivo final ya muxeado (resolución, audio, HDR, códec).
        # Es la fuente fiable: refleja exactamente lo que mediainfo verá en el .mkv resultante,
        # incluyendo la combinación DV+HDR10+ que solo aparece tras la fusión con dovi_tool.
        # Nota: hay que invalidar la cache de mediainfo para este path porque el archivo
        # acaba de ser modificado por mkvmerge/mkvpropedit.
        $cacheMediainfo.Remove($archivoSalidaTemporal) | Out-Null
        $datosTec = Get-DatosTecnicosMkv $archivoSalidaTemporal

        $resStr      = $datosTec.Resolucion
        $codecVideo  = $datosTec.CodecVideo
        $hdrVideo    = $datosTec.HdrVideo
        $audioFichero = $datosTec.AudioPista

        # Cambio 2b: el nombre del archivo se basa en el MEJOR audio del idioma prioritario,
        # prefiriendo DTS sobre Dolby. Recorremos $audiosOrdenados (ya ordenado por idioma),
        # tomamos el idioma de mayor prioridad presente, y dentro de ese idioma elegimos:
        #   - si hay DTS, el mejor DTS (DTS-HD MA > DTS)
        #   - si no hay DTS, el mejor Dolby (el primero de ese idioma)
        # El resultado sustituye al audio "por defecto" (primera pista) que daba Get-DatosTecnicosMkv.
        if ($audiosOrdenados -and $audiosOrdenados.Count -gt 0) {
            $pesoMin = ($audiosOrdenados | Measure-Object PesoLang -Minimum).Minimum
            $audiosIdiomaTop = @($audiosOrdenados | Where-Object { $_.PesoLang -eq $pesoMin })
            $audioParaNombre = $null
            # ¿hay DTS en el idioma prioritario?
            $dtsDelTop = @($audiosIdiomaTop | Where-Object { $_.Format -match "(?i)DTS" })
            if ($dtsDelTop.Count -gt 0) {
                # Mejor DTS: DTS-HD MA antes que DTS pelado
                $audioParaNombre = $dtsDelTop | Sort-Object @{Expression={ if ($_.Profile -match "(?i)MA|Master Audio" -or $_.Comm -match "(?i)MA|Master Audio") {1} else {0} }; Descending=$true} | Select-Object -First 1
            } else {
                # No hay DTS: el primer audio del idioma top (que será el mejor Dolby por orden)
                $audioParaNombre = $audiosIdiomaTop[0]
            }
            if ($audioParaNombre) {
                # Reconstruir la cadena "CODEC CANALES [Atmos]" desde la pista elegida (sin idioma)
                $aFmt = $audioParaNombre.Format; $aProf = $audioParaNombre.Profile; $aComm = $audioParaNombre.Comm
                $cN = switch ("$($audioParaNombre.Chan)") { "8" {"7.1"} "6" {"5.1"} "2" {"2.0"} "1" {"1.0"} default { "$($audioParaNombre.Chan)" } }
                if (-not $cN) { $cN = "2.0" }
                $cod = $aFmt
                if ($aFmt -match "(?i)E-?AC-?3" -or $aComm -match "(?i)Plus" -or $aProf -match "(?i)E-?AC-?3") { $cod = "DD+" }
                elseif ($aFmt -match "(?i)AC-?3")    { $cod = "DD" }
                elseif ($aFmt -match "(?i)TrueHD|MLP") { $cod = "TrueHD" }
                elseif ($aFmt -match "(?i)DTS") {
                    if ($aProf -match "(?i)MA|Master Audio" -or $aComm -match "(?i)MA|Master Audio") { $cod = "DTS-HD MA" } else { $cod = "DTS" }
                }
                elseif ($aFmt -match "(?i)AAC") { $cod = "AAC" }
                elseif ($aFmt -match "(?i)FLAC") { $cod = "FLAC" }
                elseif ($aFmt -match "(?i)PCM|LPCM") { $cod = "LPCM" }
                $atmN = if ($aComm -match "(?i)Atmos" -or $aFmt -match "(?i)Atmos") { " Atmos" } else { "" }
                $audioFichero = "$cod $cN$atmN"
            }
        }

        # Volcado al log de los datos técnicos releídos del archivo final.
        Write-DiagSeccion "DATOS TÉCNICOS RELEÍDOS DEL ARCHIVO FINAL"
        Write-DiagLinea "Resolucion: $resStr"
        Write-DiagLinea "CodecVideo: $codecVideo"
        Write-DiagLinea "HdrVideo: '$hdrVideo'  (vacío si no detectado)"
        Write-DiagLinea "AudioPista (Get-DatosTecnicosMkv): '$($datosTec.AudioPista)'"
        Write-DiagLinea "AudioFichero (para nombre, regla DTS): '$audioFichero'"

        $base = "$tituloPelicula ($anoPelicula)"
        if ($serieTag) { $base += " $serieTag" }

        # =========================================================================
        # CONSTRUCCIÓN DEL NOMBRE FINAL SEGÚN NORMAS HDZ
        # =========================================================================
        # Separador: ESPACIO en todos los casos (WEB-DL y Físico).
        # Audio: el formato leído del archivo final, sin idioma (ej: "DD+ 5.1", "TrueHD 7.1 Atmos").
        # =========================================================================

        # Sufijo -HDZ: opcional según elección del usuario (cambio 5).
        $sufijoHDZ = if ($anadirSufijoHDZ) { "-HDZ" } else { "" }

        if ($tipoOrigen -eq "WEB") {
            # Patrón WEB: Título (Año) [Etiq] [SXXEXX] 2160p PLAT WEB-DL|WEBRip AUDIO [HDR] H.265[-HDZ].mkv
            $piezas = @()
            $piezas += $base
            if ($etiquetasExtra) { $piezas += $etiquetasExtra }
            $piezas += $resStr
            $piezas += $plataformaFormato
            $piezas += $webTipo   # "WEB-DL" o "WEBRip"
            $piezas += $audioFichero
            if ($hdrVideo) { $piezas += $hdrVideo }
            $piezas += if ($codecVideo -eq "AV1") { "AV1$sufijoHDZ" } else { "H.$codecVideo$sufijoHDZ" }
        } else {
            # Patrón Físico: Título (Año) [Etiq] FORMATO RESOLUCION [CAPA] [CODEC] AUDIO[-HDZ].mkv
            # CODEC: "x265"/"x264" en Rip y MHD; "HEVC" en Remastered; vacío en Full/Remux.
            $codecTag = ""
            if ($plataformaFormato -match "Rip")           { $codecTag = if ($codecVideo -eq "AV1") { "AV1" } else { "x$codecVideo" } }
            elseif ($plataformaFormato -eq "Remastered")   { $codecTag = if ($codecVideo -eq "AV1") { "AV1" } else { "HEVC" } }
            elseif ($plataformaFormato -eq "MHD")          { $codecTag = if ($codecVideo -eq "AV1") { "AV1" } else { "x$codecVideo" } }
            # En Full/Remux el códec NO se incluye

            $piezas = @()
            $piezas += $base
            if ($etiquetasExtra) { $piezas += $etiquetasExtra }
            $piezas += $plataformaFormato
            $piezas += $resStr
            if ($hdrVideo) { $piezas += $hdrVideo }
            if ($codecTag) { $piezas += $codecTag }
            $piezas += "$audioFichero$sufijoHDZ"
        }

        $nombreFinal = (($piezas | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }) -join " ") + ".mkv"
        $nombreFinal = $nombreFinal -replace "\s+", " "

        # =========================================================================
        # FASE 5: CAPTURAS Y LIMPIEZA
        # =========================================================================
        # Limpieza de temporales (antes de renombrar el resultado).
        Write-Host "`n[>>] Limpiando archivos temporales..." -ForegroundColor Cyan
        Write-DiagPaso "Limpieza y renombrado final: INICIO"
        foreach ($del in $archivosBorrables) { Remove-Item -LiteralPath $del -ErrorAction SilentlyContinue }
        # Limpiar los .sup extraídos y los .srt convertidos de ESTE archivo. En AMBAS opciones
        # (mantener PGS+SRT / solo SRT) ya se han metido al MKV lo que correspondía, así que estos
        # archivos de paso de la carpeta ya no hacen falta. Como todos llevan el nombre del vídeo
        # delante ("<nombre>_sub_<N>_<lang>.ext"), borramos SOLO los de este archivo (los que
        # empiezan por su nombre), sin tocar los de archivos del lote aún sin procesar.
        # OJO: solo borramos los .sup EXTRAÍDOS (patrón sub_<N>_<lang>), no los .sup sueltos que
        # el usuario haya dejado a mano con otro nombre.
        $prefijoLimpieza = [System.IO.Path]::GetFileNameWithoutExtension($archivoPrincipal.Name)
        $prefijoLimpiezaEsc = [regex]::Escape($prefijoLimpieza)
        Get-ChildItem -LiteralPath $rutaCarpeta -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "(?i)^${prefijoLimpiezaEsc}_sub_\d+_[a-z]{2,3}\.(sup|srt)$" } |
            Remove-Item -ErrorAction SilentlyContinue

        # IMPORTANTE: resolvemos los archivos originales (borrar o marcar .procesado) ANTES de
        # renombrar el resultado. Esto es clave al REPROCESAR un archivo que ya lleva -HDZ: el
        # nombre final calculado coincide con el del original, así que si renombrásemos primero
        # tendríamos colisión y el resultado saldría con un "(2)". Liberando el nombre del original
        # antes, el resultado puede tomar el nombre limpio.
        # El archivo muxeado está en $archivoSalidaTemporal ("Ensamblaje_..."), distinto del original,
        # así que borrar/renombrar el original no afecta al resultado.
        # Salvaguarda: solo tocamos los originales si el muxeado existe de verdad (no perder datos).
        if (Test-Path -LiteralPath $archivoSalidaTemporal) {
            if ($borrarOriginales) {
                Write-Host "   -> Borrando archivos fuente originales..." -ForegroundColor DarkGray
                if ($esHibrido) {
                    Remove-Item -LiteralPath $archivoDV.FullName    -ErrorAction SilentlyContinue
                    Remove-Item -LiteralPath $archivoHDR10.FullName -ErrorAction SilentlyContinue
                }
                else { Remove-Item -LiteralPath $archivoPrincipal.FullName -ErrorAction SilentlyContinue }
            } else {
                Write-Host "   -> Marcando originales como procesados (.procesado)..." -ForegroundColor DarkGray
                if ($esHibrido) {
                    Rename-Item -LiteralPath $archivoDV.FullName    -NewName "$($archivoDV.Name).procesado"    -ErrorAction SilentlyContinue
                    Rename-Item -LiteralPath $archivoHDR10.FullName -NewName "$($archivoHDR10.Name).procesado" -ErrorAction SilentlyContinue
                } else {
                    Rename-Item -LiteralPath $archivoPrincipal.FullName -NewName "$($archivoPrincipal.Name).procesado" -ErrorAction SilentlyContinue
                }
            }
        }

        # Ahora el nombre está libre: renombramos el resultado (Bug 10: evitar colisiones residuales).
        $archivoFinalDestino = Rename-EvitarColision $archivoSalidaTemporal $nombreFinal
        # Carpeta de salida elegida en la GUI: movemos allí el MKV final (las capturas, que se
        # generan a continuación junto al archivo, caerán también en esa carpeta).
        if ($carpetaSalidaGui -and (Split-Path $archivoFinalDestino -Parent) -ne $carpetaSalidaGui) {
            try {
                $leafFin = Split-Path $archivoFinalDestino -Leaf
                $destFin = Join-Path $carpetaSalidaGui $leafFin
                $kk = 1
                while (Test-Path -LiteralPath $destFin) {
                    $destFin = Join-Path $carpetaSalidaGui ("{0} ({1}){2}" -f [IO.Path]::GetFileNameWithoutExtension($leafFin), $kk, [IO.Path]::GetExtension($leafFin)); $kk++
                }
                Move-Item -LiteralPath $archivoFinalDestino -Destination $destFin
                $archivoFinalDestino = $destFin
                Write-Host "   -> Movido a la carpeta de salida: $carpetaSalidaGui" -ForegroundColor DarkGray
            } catch { Add-Incidencia $global:archivoEnCurso "No se pudo mover a la carpeta de salida: $($_.Exception.Message)" }
        }
        $global:archivosFinalesLote += $archivoFinalDestino

        # Volcado al log del archivo final, post-renombrado.
        Write-DiagArchivoSalida $archivoFinalDestino $nombreFinal

        # Capturas: tonemapped solo si el vídeo es HDR. $hdrVideo viene de Get-DatosTecnicosMkv
        # (será "DV HDR10", "HDR10", "HDR10+", etc. en HDR; cadena vacía en SDR).
        $videoEsHdr = -not [string]::IsNullOrWhiteSpace($hdrVideo)
        Set-ProgresoGui ([int]($global:pctBaseGui + $global:pctPasoGui * 0.85)) "Generando capturas $($global:idxGrupoGui)/$($global:totalGruposGui)" (Split-Path $archivoFinalDestino -Leaf)
        New-Capturas $archivoFinalDestino (Split-Path $archivoFinalDestino -Leaf) $numCapturas $videoEsHdr

        # Torrent del resultado final (si se pidió). Fallo no fatal: se anota y se sigue.
        if ($modoTorrent -in @("INDIVIDUAL", "AMBOS")) {
            Set-ProgresoGui ([int]($global:pctBaseGui + $global:pctPasoGui * 0.93)) "Creando torrent $($global:idxGrupoGui)/$($global:totalGruposGui)" (Split-Path $archivoFinalDestino -Leaf)
            Write-Host "`n[>>] Creando archivo .torrent..." -ForegroundColor Cyan
            try {
                $rutaTorrentFinal = New-TorrentArchivo $archivoFinalDestino $torrentAnnounce "" $carpetaTorrentGui
                Write-Host "   -> Torrent creado: $(Split-Path $rutaTorrentFinal -Leaf)" -ForegroundColor Green
                Add-ResultadoGui "archivo" $rutaTorrentFinal $archivoFinalDestino   # la GUI crea una pestaña de subida
            } catch {
                Write-Host "   [!] No se pudo crear el .torrent: $($_.Exception.Message)" -ForegroundColor Red
                Add-Incidencia $global:archivoEnCurso "Fallo creando el .torrent: $($_.Exception.Message)"
            }
        }

        $etiquetaTipo = if ($esSerie) { "Episodio" } else { "Archivo" }
        Set-ProgresoGui ([int]($global:pctBaseGui + $global:pctPasoGui)) "Completado $($global:idxGrupoGui)/$($global:totalGruposGui)" (Split-Path $archivoFinalDestino -Leaf)
        Write-Host " ¡OK! $etiquetaTipo completado: $(Split-Path $archivoFinalDestino -Leaf) " -ForegroundColor Green
        Write-DiagPaso "FIN archivo OK: $(Split-Path $archivoFinalDestino -Leaf)"

    } catch {
        # D2: capturamos cualquier excepción no controlada para no abortar el batch
        Write-Host "`n[!] EXCEPCIÓN procesando el grupo: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
        Add-Incidencia $global:archivoEnCurso "EXCEPCIÓN no controlada: $($_.Exception.Message)"
        # Limpieza defensiva de temporales
        Remove-Item "metadata_p8.rpu","video_base.hevc","video_definitivo.hevc" -ErrorAction SilentlyContinue
        Get-ChildItem -LiteralPath $rutaCarpeta -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "(?i)_audio_[a-z]{2,3}(-[a-z0-9]{2,3})?\.eac3$" -or $_.Name -match "(?i)^audio_.*_temp\.eac3$" } |
            Remove-Item -ErrorAction SilentlyContinue
        Write-Host "   -> Continuando con el siguiente elemento..." -ForegroundColor Yellow
        continue
    }
}

# =========================================================================
# TORRENT DE PACK (lote completo): se crea una carpeta con el nombre del pack,
# se mueven dentro los MKV finales y se genera el .torrent multi-archivo junto
# a ella, listo para subir y sembrar.
# =========================================================================
if ($modoTorrent -in @("PACK", "AMBOS") -and @($global:archivosFinalesLote).Count -gt 0) {
    try {
        # Nombre automático: el del primer archivo final, sustituyendo SxxEyy(-Ezz) por Sxx
        # (convención de pack de temporada). Para películas se usa tal cual.
        $primeroPack = [System.IO.Path]::GetFileNameWithoutExtension((Split-Path $global:archivosFinalesLote[0] -Leaf))
        $nombrePackAuto = (($primeroPack -replace '(?i)\b(S\d{1,3})E\d{1,4}(?:-E?\d{1,4})*\b', '$1') -replace "\s+", " ").Trim()
        $cfgNombrePack = Get-CfgGui "TorrentPackNombre"
        $nombrePack = if ($cfgNombrePack) {
            Write-Host "`n[GUI] Nombre del pack: $("$cfgNombrePack".Trim()) (configuración)." -ForegroundColor DarkGray
            "$cfgNombrePack".Trim()
        } elseif ($global:cfgGui) {
            # Lanzado desde la GUI con el campo de nombre vacío: la GUI promete "vacío =
            # automático", así que usamos el nombre automático sin preguntar por consola.
            Write-Host "`n[GUI] Nombre del pack (automático): $nombrePackAuto" -ForegroundColor DarkGray
            $nombrePackAuto
        } else {
            $rPack = Read-Host "`nNombre del pack (ENTER = $nombrePackAuto)"
            if ([string]::IsNullOrWhiteSpace($rPack)) { $nombrePackAuto } else { $rPack.Trim() }
        }
        foreach ($chInv in [System.IO.Path]::GetInvalidFileNameChars()) { $nombrePack = $nombrePack.Replace("$chInv", "") }
        if ([string]::IsNullOrWhiteSpace($nombrePack)) { $nombrePack = "Pack HDZ" }

        Write-Host "`n[>>] Creando PACK '$nombrePack' ($(@($global:archivosFinalesLote).Count) archivo(s))..." -ForegroundColor Cyan
        # La carpeta del pack se crea en la carpeta de salida si se eligió una.
        $baseDirPack = if ($carpetaSalidaGui) { $carpetaSalidaGui } else { $rutaCarpeta }
        $rutaPack = Join-Path $baseDirPack $nombrePack
        [void][System.IO.Directory]::CreateDirectory($rutaPack)
        foreach ($fFinal in $global:archivosFinalesLote) {
            if (Test-Path -LiteralPath $fFinal) {
                $destinoPack = Join-Path $rutaPack (Split-Path $fFinal -Leaf)
                if (-not (Test-Path -LiteralPath $destinoPack)) {
                    Move-Item -LiteralPath $fFinal -Destination $destinoPack
                }
            }
        }
        $rutaTorrentPack = New-TorrentPack $rutaPack $torrentAnnounce "" $carpetaTorrentGui
        Write-Host "   -> Pack listo: carpeta '$nombrePack' + $(Split-Path $rutaTorrentPack -Leaf)" -ForegroundColor Green
        $primerVideoPack = @(Get-ChildItem -LiteralPath $rutaPack -File -Recurse -ErrorAction SilentlyContinue |
                             Where-Object { $_.Extension -match "(?i)^\.(mkv|mp4)$" } | Sort-Object FullName | Select-Object -First 1)
        Add-ResultadoGui "pack" $rutaTorrentPack $(if ($primerVideoPack.Count) { $primerVideoPack[0].FullName } else { "" })
    } catch {
        Write-Host "   [!] No se pudo crear el torrent del pack: $($_.Exception.Message)" -ForegroundColor Red
        Add-Incidencia $null "Fallo creando el torrent del PACK: $($_.Exception.Message)"
    }
}

# Resumen final: si hubo incidencias, las listamos para que no pasen desapercibidas.
if ($global:incidencias.Count -gt 0) {
    Write-Host "`n=========================================" -ForegroundColor Yellow
    Write-Host " PROCESO TERMINADO CON $($global:incidencias.Count) INCIDENCIA(S) " -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Yellow
    foreach ($inc in $global:incidencias) {
        Write-Host "  [!] $inc" -ForegroundColor Red
    }
    Write-Host "`nRevisa el log para más detalle: $rutaLog" -ForegroundColor DarkYellow
    Set-ProgresoGuiFin "Terminado con $($global:incidencias.Count) incidencia(s)"
} else {
    Write-Host "`n=========================================" -ForegroundColor Green
    Write-Host " ¡TODOS LOS ARCHIVOS PROCESADOS CON ÉXITO! " -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Set-ProgresoGuiFin "¡Todos los archivos procesados!"
}

try { Stop-Transcript | Out-Null } catch {}

# Limpieza global de seguridad: borrar cualquier .sup residual de Subtitle Edit que pudiera
# haber quedado (archivos de paso, nunca se muxean).
try {
    Get-ChildItem -LiteralPath $rutaCarpeta -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "(?i)^.+_sub_\d+_[a-z]{2,3}\.sup$" } |
        Remove-Item -ErrorAction SilentlyContinue
} catch {}

# Anexar el diagnóstico al log principal (un solo archivo final) y borrar el temporal.
# Se hace tras Stop-Transcript para que el log ya no esté bloqueado por el transcript.
try {
    if (Test-Path -LiteralPath $global:rutaDiag) {
        Add-Content -LiteralPath $rutaLog -Value "`r`n`r`n=================== DIAGNÓSTICO DETALLADO ===================`r`n" -Encoding UTF8
        Get-Content -LiteralPath $global:rutaDiag -Encoding UTF8 | Add-Content -LiteralPath $rutaLog -Encoding UTF8
        Remove-Item -LiteralPath $global:rutaDiag -ErrorAction SilentlyContinue
    }
} catch {}