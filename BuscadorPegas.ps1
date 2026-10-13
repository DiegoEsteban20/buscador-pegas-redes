<#
.SYNOPSIS
    Buscador de Pegas v2 (Conectividad y Redes): busca ofertas de trabajo en GetOnBoard,
    Computrabajo, Chiletrabajos y Talent.com, filtra por tu área, marca las NUEVAS y arma
    un Excel en Escritorio\Pegas.

.EJEMPLOS
    .\BuscadorPegas.ps1                 # Busca ahora (abre el Excel si hay ofertas nuevas)
    .\BuscadorPegas.ps1 -SiempreAbrir   # Abre el Excel aunque no haya nuevas
    .\BuscadorPegas.ps1 -Instalar       # Se ejecuta solo al iniciar sesión en Windows
    .\BuscadorPegas.ps1 -Desinstalar    # Quita la tarea programada
#>
param(
    [switch]$Instalar,
    [switch]$Desinstalar,
    [switch]$SiempreAbrir,
    [switch]$NoAbrir,
    [switch]$Forzar,
    [int]$HorasEntreEjecuciones = 6,
    [string]$Solo = '',                       # prueba un solo portal (ej. -Solo Computrabajo)
    [hashtable]$UrlBase = @{}                 # solo para pruebas
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Carpeta      = $PSScriptRoot
$ArchConfig   = Join-Path $Carpeta 'Config.psd1'
$ArchVistos   = Join-Path $Carpeta 'Vistos.csv'
$ArchLog      = Join-Path $Carpeta 'BuscadorPegas.log'
$ArchUltima   = Join-Path $Carpeta 'ultima_ejecucion.txt'
$Escritorio   = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($Escritorio)) { $Escritorio = $Carpeta }
$CarpetaExcel = Join-Path $Escritorio 'Pegas'
$NombreTarea  = 'Buscador de Pegas'
$Inv          = [Globalization.CultureInfo]::InvariantCulture

function Log([string]$m) {
    $l = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
    Add-Content -Path $ArchLog -Value $l -Encoding UTF8
    Write-Host $l
}

# ================================================================ Tarea programada
if ($Desinstalar) {
    Unregister-ScheduledTask -TaskName $NombreTarea -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Tarea '$NombreTarea' eliminada." -ForegroundColor Green
    return
}
if ($Instalar) {
    $usuario = "$env:USERDOMAIN\$env:USERNAME"
    $accion = New-ScheduledTaskAction -Execute 'powershell.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
    $disparador = New-ScheduledTaskTrigger -AtLogOn -User $usuario
    $disparador.Delay = 'PT4M'   # 4 min, para no chocar con el Buscador de Ofertas
    $principal = New-ScheduledTaskPrincipal -UserId $usuario -LogonType Interactive -RunLevel Limited
    $config = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
    Register-ScheduledTask -TaskName $NombreTarea -Action $accion -Trigger $disparador `
        -Principal $principal -Settings $config -Description 'Busca ofertas de trabajo de redes y conectividad al iniciar sesión' -Force | Out-Null
    Write-Host "Listo: '$NombreTarea' se ejecutará al iniciar sesión (4 min después), máximo cada $HorasEntreEjecuciones h." -ForegroundColor Green
    return
}

if (-not $Forzar -and -not $Solo -and (Test-Path $ArchUltima)) {
    $ultima = [datetime]::ParseExact((Get-Content $ArchUltima -Raw).Trim(), 'yyyy-MM-dd HH:mm:ss', $null)
    if ((Get-Date) - $ultima -lt (New-TimeSpan -Hours $HorasEntreEjecuciones)) {
        Log "Última búsqueda hace menos de $HorasEntreEjecuciones h ($ultima). Usa -Forzar para repetir."
        return
    }
}

# ================================================================ Configuración y utilidades
$cfg = Import-PowerShellDataFile $ArchConfig
$hoy = Get-Date

$Headers = @{
    'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36'
    'Accept-Language' = 'es-CL,es;q=0.9'
    'Accept'          = 'text/html,application/json;q=0.9,*/*;q=0.8'
}
function Bajar([string]$url) {
    $r = Invoke-WebRequest -Uri $url -Headers $Headers -UseBasicParsing -TimeoutSec 40
    [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
}
function Leer-Json([string]$txt) {
    if ($PSVersionTable.PSVersion.Major -ge 6) { return ($txt | ConvertFrom-Json -AsHashtable) }
    Add-Type -AssemblyName System.Web.Extensions
    $js = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $js.MaxJsonLength = [int]::MaxValue
    $js.DeserializeObject($txt)
}
function Limpiar([string]$s) {
    if (-not $s) { return '' }
    $s = [regex]::Replace($s, '<[^>]+>', ' ')
    $s = [Net.WebUtility]::HtmlDecode($s)
    ($s -replace '\s+', ' ').Trim()
}
function Sin-Tildes([string]$s) { if (-not $s) { return '' }; ($s.ToLower().Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '') }
function En-Santiago($o) {
    if (-not $cfg.SoloSantiago) { return $true }
    $u = Sin-Tildes $o.Ubicacion
    $c = Sin-Tildes $o.Cargo
    if ($c -match "(?i)$($cfg.OtrasCiudades)") { return $false }
    if ($u -match "(?i)$($cfg.Santiago)") { return $true }
    if ($cfg.IncluirRemoto -and ($u -match 'remot|teletrabajo|home office' -or $c -match 'remot|teletrabajo|home office')) { return $true }
    if ($c -match "(?i)$($cfg.Santiago)") { return $true }
    # Sin ciudad (solo "Chile"): se acepta si no menciona otra ciudad
    if ($u -eq '' -or $u -match '^(chile)?( · (presencial|hibrido))?$') { return ($u -notmatch "(?i)$($cfg.OtrasCiudades)") }
    return $false
}
function Base([string]$portal, [string]$porDefecto) { if ($UrlBase.ContainsKey($portal)) { $UrlBase[$portal] } else { $porDefecto } }

$meses = @{ enero = 1; febrero = 2; marzo = 3; abril = 4; mayo = 5; junio = 6; julio = 7; agosto = 8; septiembre = 9; setiembre = 9; octubre = 10; noviembre = 11; diciembre = 12 }
function Fecha-Texto([string]$t) {
    # "Hace 4 horas", "Hace 3 días", "Ayer", "12 de Octubre de 2026", "hace 2 días"
    $t = (Limpiar $t).ToLower()
    if ($t -match 'hace\s+(\d+)\s+(minuto|hora)') { return $hoy.Date }
    if ($t -match 'hoy|reci[eé]n') { return $hoy.Date }
    if ($t -match 'ayer') { return $hoy.Date.AddDays(-1) }
    if ($t -match 'hace\s+(\d+)\s+d[ií]a') { return $hoy.Date.AddDays(-[int]$matches[1]) }
    if ($t -match 'hace\s+(\d+)\s+semana') { return $hoy.Date.AddDays(-7 * [int]$matches[1]) }
    if ($t -match 'hace\s+(\d+)\s+mes') { return $hoy.Date.AddMonths(-[int]$matches[1]) }
    if ($t -match '(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})' -and $meses.ContainsKey($matches[2])) {
        return Get-Date -Year ([int]$matches[3]) -Month $meses[$matches[2]] -Day ([int]$matches[1]) -Hour 0 -Minute 0 -Second 0
    }
    return $null
}

function Nueva-Oferta($portal, $busqueda, $cargo, $empresa, $ubicacion, $publicado, $link, $sueldo, $nivel) {
    [pscustomobject]@{
        Portal = $portal; Busqueda = $busqueda; Cargo = (Limpiar $cargo); Empresa = (Limpiar $empresa)
        Ubicacion = (Limpiar $ubicacion); Publicado = $publicado; Link = $link; Sueldo = $sueldo
        Nivel = $nivel; Relevancia = 0; Estado = ''
    }
}

# ---------------------------------------------------------------- Lectores por portal
$nivelesGob = @{ 1 = 'Sin experiencia'; 2 = 'Junior'; 3 = 'Semi Senior'; 4 = 'Senior'; 5 = 'Expert' }
$modalidades = @{ fully_remote = 'Remoto'; remote_local = 'Remoto (Chile)'; hybrid = 'Híbrido'; no_remote = 'Presencial'; temporarily_remote = 'Remoto temporal' }

function Leer-GetOnBoard([string]$q) {
    $b = Base 'GetOnBoard' 'https://www.getonbrd.com'
    $json = Leer-Json (Bajar "$b/api/v0/search/jobs?query=$([uri]::EscapeDataString($q))&per_page=50&expand=%5B%22company%22%5D")
    foreach ($j in $json.data) {
        $a = $j.attributes
        $pub = ([DateTimeOffset]::FromUnixTimeSeconds([int64]$a.published_at)).LocalDateTime
        if ($pub -lt $hoy.AddDays(-[int]$cfg.DiasMaximos)) { continue }
        $nivelId = 0
        if ($a.seniority -and $a.seniority.data) { $nivelId = [int]$a.seniority.data.id }
        if ($nivelId -gt [int]$cfg.NivelMaximo) { continue }
        $paises = @($a.countries)
        if (-not ($paises | Where-Object { $cfg.Paises -contains $_ })) { continue }
        $empresa = ''
        if ($a.company -and $a.company.data -and $a.company.data.attributes) { $empresa = $a.company.data.attributes.name }
        $sueldo = ''
        if ($a.min_salary -or $a.max_salary) { $sueldo = "US$ $($a.min_salary) - $($a.max_salary)" }
        $mod = $modalidades[[string]$a.remote_modality]
        $ubic = (($paises -join ', '), $mod | Where-Object { $_ }) -join ' · '
        Nueva-Oferta 'GetOnBoard' $q $a.title $empresa $ubic $pub $j.links.public_url $sueldo $nivelesGob[$nivelId]
    }
}

function Leer-Computrabajo([string]$q) {
    $b = Base 'Computrabajo' 'https://cl.computrabajo.com'
    $slug = ($q.ToLower() -replace '[^a-z0-9ñáéíóú ]', '' -replace '\s+', '-')
    $html = Bajar "$b/trabajo-de-$slug`?pubdate=$($cfg.ComputrabajoDias)"
    foreach ($a in ($html -split '<article class="box_offer' | Select-Object -Skip 1)) {
        $t = [regex]::Match($a, '<a class="js-o-link[^"]*" href="([^"#]+)[^"]*">\s*(.*?)\s*</a>', 'Singleline')
        if (-not $t.Success) { continue }
        $co = [regex]::Match($a, 'offer-grid-article-company-url>\s*(.*?)\s*</a>', 'Singleline')
        $lo = [regex]::Match($a, '<p class="fs16 fc_base mt5">\s*<span class="mr10">\s*(.*?)\s*</span>', 'Singleline')
        $da = [regex]::Match($a, '<p class="fs13 fc_aux mt15">\s*(.*?)\s*</p>', 'Singleline')
        $link = $t.Groups[1].Value; if ($link.StartsWith('/')) { $link = 'https://cl.computrabajo.com' + $link }
        $empresa = if ($co.Success) { $co.Groups[1].Value } else { '(confidencial)' }
        Nueva-Oferta 'Computrabajo' $q $t.Groups[2].Value $empresa $lo.Groups[1].Value (Fecha-Texto $da.Groups[1].Value) $link '' ''
    }
}

function Leer-Chiletrabajos([string]$q) {
    $b = Base 'Chiletrabajos' 'https://www.chiletrabajos.cl'
    $html = Bajar "$b/encuentra-un-empleo?2=$([uri]::EscapeDataString($q))&action=search"
    foreach ($a in ($html -split '<div class="job-item' | Select-Object -Skip 1)) {
        $t = [regex]::Match($a, '<h2 class="title[^"]*">\s*<a href="([^"]+)"[^>]*>(.*?)</a>', 'Singleline')
        if (-not $t.Success) { continue }
        $metas = @([regex]::Matches($a, '<h3 class="meta">(.*?)</h3>', 'Singleline') | ForEach-Object { Limpiar $_.Groups[1].Value })
        $empresa = ''; $ubic = ''; $fecha = $null
        if ($metas.Count -ge 1) {
            $partes = $metas[0] -split ',\s*', 2
            $empresa = $partes[0]; if ($partes.Count -gt 1) { $ubic = $partes[1] }
        }
        if ($metas.Count -ge 2) { $fecha = Fecha-Texto $metas[1] }
        Nueva-Oferta 'Chiletrabajos' $q $t.Groups[2].Value $empresa $ubic $fecha $t.Groups[1].Value '' ''
    }
}

function Leer-Talent([string]$q) {
    $b = Base 'Talent' 'https://cl.talent.com'
    $html = Bajar "$b/jobs?k=$([uri]::EscapeDataString($q))&l=Chile"
    foreach ($c in ($html -split 'data-new-id="' | Select-Object -Skip 1)) {
        $id = [regex]::Match($c, '^(\d+)"').Groups[1].Value
        $t = [regex]::Match($c, 'JobCard_title[^"]*">(.*?)</h2>', 'Singleline')
        if (-not ($id -and $t.Success)) { continue }
        $co = [regex]::Match($c, 'JobCard_company[^"]*">(.*?)</span>', 'Singleline')
        $lo = [regex]::Match($c, 'JobCard_location[^"]*">(.*?)</span>', 'Singleline')
        $ti = [regex]::Match($c, 'JobCard_timeText[^"]*">(.*?)</', 'Singleline')
        Nueva-Oferta 'Talent.com' $q $t.Groups[1].Value $co.Groups[1].Value $lo.Groups[1].Value (Fecha-Texto $ti.Groups[1].Value) "https://cl.talent.com/view?id=$id" '' ''
    }
}

# ================================================================ Búsqueda
$lectores = [ordered]@{
    GetOnBoard    = ${function:Leer-GetOnBoard}
    Computrabajo  = ${function:Leer-Computrabajo}
    Chiletrabajos = ${function:Leer-Chiletrabajos}
    Talent        = ${function:Leer-Talent}
}
$rxIncluir = "(?i)$($cfg.Incluir)"
$rxExcluir = "(?i)$($cfg.Excluir)"
$rxSenior  = "(?i)$($cfg.PalabrasSenior)"
$rxMuy     = "(?i)$($cfg.MuyRelevante)"
$rxRel     = "(?i)$($cfg.Relevante)"
$rxEntrada = "(?i)$($cfg.NivelEntrada)"

$ofertas = @{}
$porTitulo = @{}
$errores = New-Object System.Collections.Generic.List[object]
$resumen = [ordered]@{}

foreach ($portal in $lectores.Keys) {
    if (-not $cfg.Portales[$portal]) { continue }
    if ($Solo -and $portal -notlike "*$Solo*") { continue }
    $n = 0; $fallos = 0
    foreach ($q in $cfg.Busquedas) {
        try {
            foreach ($o in @(& $lectores[$portal] $q)) {
                if (-not $o -or -not $o.Cargo) { continue }
                if ($o.Cargo -notmatch $rxIncluir) { continue }
                if ($o.Cargo -match $rxExcluir) { continue }
                if ($cfg.ExcluirSenior -and $o.Cargo -match $rxSenior) { continue }
                if (-not (En-Santiago $o)) { continue }
                $clave = ("{0}|{1}" -f $o.Cargo, $o.Empresa).ToLower() -replace '[^a-z0-9ñáéíóú|]', ''
                if ($ofertas.ContainsKey($o.Link) -or $porTitulo.ContainsKey($clave)) { continue }   # misma oferta en otro portal o búsqueda
                $r = 0
                $r += 3 * [regex]::Matches($o.Cargo, $rxMuy).Count
                $r += [regex]::Matches($o.Cargo, $rxRel).Count
                if ($o.Cargo -match $rxEntrada) { $r += 2; if (-not $o.Nivel) { $o.Nivel = 'Entrada / Junior' } }
                $o.Relevancia = [math]::Min($r, 10)
                $ofertas[$o.Link] = $o; $porTitulo[$clave] = 1; $n++
            }
        } catch {
            $fallos++
            if ($fallos -le 2) { $errores.Add([pscustomobject]@{ Portal = $portal; Busqueda = $q; Error = $_.Exception.Message }) }
            Log "ERROR $portal '$q': $($_.Exception.Message)"
        }
        Start-Sleep -Milliseconds 900   # no saturar los portales
    }
    $resumen[$portal] = $n
    Log "$portal`: $n ofertas del área ($fallos búsquedas con error)"
}

# ================================================================ Nuevas vs vistas
$vistos = @{}
if (Test-Path $ArchVistos) { Import-Csv $ArchVistos -Encoding UTF8 | ForEach-Object { $vistos[$_.Link] = $_.Visto } }
foreach ($o in $ofertas.Values) {
    if (-not $vistos.ContainsKey($o.Link)) { $o.Estado = 'NUEVA'; $vistos[$o.Link] = $hoy.ToString('yyyy-MM-dd') }
}
# Olvida lo visto hace más de 90 días para que el archivo no crezca sin límite
$vistos.GetEnumerator() | Where-Object { [datetime]::ParseExact($_.Value, 'yyyy-MM-dd', $null) -gt $hoy.AddDays(-90) } |
    ForEach-Object { [pscustomobject]@{ Link = $_.Key; Visto = $_.Value } } | Export-Csv $ArchVistos -NoTypeInformation -Encoding UTF8

$orden = @(
    @{ Expression = { $_.Estado -eq 'NUEVA' }; Descending = $true },
    @{ Expression = 'Relevancia'; Descending = $true },
    @{ Expression = { if ($_.Publicado) { $_.Publicado } else { [datetime]'2000-01-01' } }; Descending = $true }
)
$todas  = @($ofertas.Values | Sort-Object $orden)
$nuevas = @($todas | Where-Object Estado -eq 'NUEVA')
$top    = @($todas | Where-Object { $_.Relevancia -ge 3 } | Select-Object -First 40)
Log "Total: $($todas.Count) ofertas del área, $($nuevas.Count) nuevas"

# ================================================================ Links para portales que no se pueden leer
$terminos = @('técnico en redes', 'soporte TI', 'conectividad', 'analista NOC', 'telecomunicaciones')
$links = New-Object System.Collections.Generic.List[object]
foreach ($p in $terminos) {
    $e = [uri]::EscapeDataString($p)
    $links.Add([pscustomobject]@{ Portal = 'LinkedIn'; Busqueda = $p; Link = "https://www.linkedin.com/jobs/search/?keywords=$e&location=Santiago%2C%20Regi%C3%B3n%20Metropolitana%2C%20Chile&f_TPR=r604800&f_E=1%2C2" })
    $slugL = ($p.ToLower().Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '') -replace '\s+', '-'
    $links.Add([pscustomobject]@{ Portal = 'Laborum'; Busqueda = $p; Link = "https://www.laborum.cl/empleos-busqueda-$slugL.html" })
    $links.Add([pscustomobject]@{ Portal = 'Trabajando.com'; Busqueda = $p; Link = "https://www.trabajando.cl/trabajo-empleo/$e" })
    $links.Add([pscustomobject]@{ Portal = 'Indeed'; Busqueda = $p; Link = "https://cl.indeed.com/jobs?q=$e&l=Santiago%2C+Regi%C3%B3n+Metropolitana&fromage=7" })
}
$links.Add([pscustomobject]@{ Portal = 'DuocLaboral (con tu cuenta Duoc)'; Busqueda = '-'; Link = 'https://duoclaboral.cl/' })
$links.Add([pscustomobject]@{ Portal = 'Empleos Públicos'; Busqueda = 'Área: Informática / Telecomunicaciones'; Link = 'https://www.empleospublicos.cl/' })
$links.Add([pscustomobject]@{ Portal = 'Bolsa Nacional de Empleo'; Busqueda = '-'; Link = 'https://www.bne.cl/' })

# ================================================================ Excel
New-Item -ItemType Directory -Path $CarpetaExcel -Force | Out-Null
$archExcel = Join-Path $CarpetaExcel ("Pegas_{0}.xlsx" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))

function Estrellas([int]$r) { if ($r -ge 6) { '★★★' } elseif ($r -ge 3) { '★★' } elseif ($r -ge 1) { '★' } else { '' } }
function Filas($lista) {
    foreach ($o in $lista) {
        [pscustomobject][ordered]@{
            'Estado' = $o.Estado; 'Relevancia' = (Estrellas $o.Relevancia)
            'Publicado' = $(if ($o.Publicado) { $o.Publicado.ToString('yyyy-MM-dd') } else { '' })
            'Cargo' = $o.Cargo; 'Empresa' = $o.Empresa; 'Ubicación' = $o.Ubicacion; 'Nivel' = $o.Nivel
            'Sueldo' = $o.Sueldo; 'Portal' = $o.Portal; 'Encontrada con' = $o.Busqueda; 'Link' = $o.Link
        }
    }
}

try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Log 'Instalando módulo ImportExcel (solo la primera vez)...'
        if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
        }
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ImportExcel

    function Link-Formula($ws, [int]$col, [string]$texto) {
        if (-not $ws.Dimension) { return }
        for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
            $c = $ws.Cells[$r, $col]
            $url = [string]$c.Value
            if ($c.Hyperlink) { $url = $c.Hyperlink.OriginalString; $c.Hyperlink = $null }
            if ($url -match '^https?://') {
                $c.Formula = 'HYPERLINK("' + ($url -replace '"', '%22') + '","' + $texto + '")'
                $c.Style.Font.UnderLine = $true
                $c.Style.Font.Color.SetColor([Drawing.Color]::FromArgb(5, 99, 193))
            }
        }
    }
    function Formato($pkg, [string]$hoja) {
        $ws = $pkg.Workbook.Worksheets[$hoja]
        if (-not $ws.Dimension) { return }
        $ult = $ws.Dimension.End.Row
        Link-Formula $ws 11 'Postular'
        $anchos = @(9, 10, 11, 55, 28, 30, 16, 16, 14, 22, 9)
        for ($i = 0; $i -lt $anchos.Count; $i++) { $ws.Column($i + 1).Width = $anchos[$i] }
        $ws.Column(2).Style.Font.Color.SetColor([Drawing.Color]::FromArgb(191, 143, 0))
        Add-ConditionalFormatting -Worksheet $ws -Range "A2:A$ult" -RuleType Equal -ConditionValue '"NUEVA"' -BackgroundColor ([Drawing.Color]::FromArgb(198, 239, 206)) -ForegroundColor ([Drawing.Color]::FromArgb(0, 97, 0))
        Add-ConditionalFormatting -Worksheet $ws -Range "G2:G$ult" -RuleType NotContainsBlanks -BackgroundColor ([Drawing.Color]::FromArgb(255, 235, 156))
    }

    $comun = @{ Path = $archExcel; AutoFilter = $true; FreezeTopRow = $true; BoldTopRow = $true; PassThru = $true }
    if ($top.Count) {
        $pkg = Filas $top | Export-Excel @comun -WorksheetName 'Más afines' -TableStyle Medium2
        Formato $pkg 'Más afines'; Close-ExcelPackage $pkg
    }
    if ($nuevas.Count) {
        $pkg = Filas $nuevas | Export-Excel @comun -WorksheetName 'Nuevas' -TableStyle Medium7
        Formato $pkg 'Nuevas'; Close-ExcelPackage $pkg
    }
    if ($todas.Count) {
        $pkg = Filas $todas | Export-Excel @comun -WorksheetName 'Todas' -TableStyle Light9
        Formato $pkg 'Todas'; Close-ExcelPackage $pkg
    }
    $pkg = $links | Export-Excel @comun -WorksheetName 'Otros portales' -TableStyle Light1
    $ws = $pkg.Workbook.Worksheets['Otros portales']
    Link-Formula $ws 3 'Abrir búsqueda'
    $ws.Column(1).Width = 32; $ws.Column(2).Width = 40; $ws.Column(3).Width = 16
    Close-ExcelPackage $pkg
    if ($errores.Count) { $errores | Export-Excel -Path $archExcel -WorksheetName 'Errores' -AutoSize -BoldTopRow }
    Log "Excel creado: $archExcel"
} catch {
    $archExcel = [IO.Path]::ChangeExtension($archExcel, '.csv')
    Filas $todas | Export-Csv $archExcel -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Log "No se pudo crear el Excel ($($_.Exception.Message)). Se guardó como CSV: $archExcel"
}

Get-ChildItem $CarpetaExcel -Filter 'Pegas_*' | Sort-Object LastWriteTime -Descending | Select-Object -Skip 20 | Remove-Item -ErrorAction SilentlyContinue
if (-not $Solo) { Set-Content -Path $ArchUltima -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }

# ================================================================ Aviso
if ($nuevas.Count -and $env:OS -eq 'Windows_NT') {
    try {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        $n = New-Object System.Windows.Forms.NotifyIcon
        $n.Icon = ([type]'System.Drawing.SystemIcons')::Information
        $n.Visible = $true
        $n.ShowBalloonTip(10000, 'Buscador de Pegas', "$($nuevas.Count) ofertas nuevas de redes y conectividad. Revisa Escritorio\Pegas.", 'Info')
        Start-Sleep -Seconds 6
        $n.Dispose()
    } catch { }
}
if (-not $NoAbrir -and (Test-Path $archExcel) -and ($nuevas.Count -or $SiempreAbrir)) { Invoke-Item $archExcel }
