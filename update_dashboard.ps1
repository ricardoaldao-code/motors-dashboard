# ============================================================
# update_dashboard.ps1
# Regenera el dashboard Motors MLM y hace push a GitHub Pages
# Ejecutar diariamente via Task Scheduler
# ============================================================

Set-Location "$env:USERPROFILE\motors-dashboard"

$TODAY     = (Get-Date).ToString("yyyy-MM-dd")
$YESTERDAY = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")

# Fecha inicio semana 1 (referencia fija para calcular SEM)
$SEM_BASE  = "2026-04-08"

# Rango de datos: desde 09/04 hasta ayer
$DATE_FROM = "2026-04-09"
$DATE_TO   = $YESTERDAY

Write-Host "=== Motors Dashboard Update ===" -ForegroundColor Cyan
Write-Host "Fecha: $TODAY | Datos hasta: $YESTERDAY"

# ── 1. Query datos diarios ─────────────────────────────────
Write-Host "`n[1/4] Consultando datos diarios en BigQuery..."

$q_diario = @"
WITH active_dealers AS (
  SELECT CUS_CUST_ID, EJECUTIVO_ACT AS ASESOR, SEGMENTACION_NSM
  FROM ``meli-bi-data.WHOWNER.DM_VIS_SELLER_MO_L1``
  WHERE SIT_SITE_ID = 'MLM' AND VERTICAL = 'MO'
    AND FECHA = (SELECT MAX(FECHA) FROM ``meli-bi-data.WHOWNER.DM_VIS_SELLER_MO_L1``
                 WHERE SIT_SITE_ID = 'MLM' AND VERTICAL = 'MO' AND FECHA <= '$DATE_TO')
),
dealer_names AS (
  SELECT CUS_CUST_ID, MAX(NOMBRE_DE_LA_CUENTA) AS NOMBRE
  FROM ``meli-bi-data.WHOWNER.DM_VIS_SELLER_MO_L1``
  WHERE SIT_SITE_ID = 'MLM' AND VERTICAL = 'MO' AND NOMBRE_DE_LA_CUENTA IS NOT NULL
  GROUP BY CUS_CUST_ID
),
fechas AS (
  SELECT fecha,
    CAST(FLOOR(DATE_DIFF(fecha, DATE '$SEM_BASE', DAY)/7)+1 AS INT64) AS SEM
  FROM UNNEST(GENERATE_DATE_ARRAY('$DATE_FROM', '$DATE_TO', INTERVAL 1 DAY)) AS fecha
),
daily AS (
  SELECT v.CUS_CUST_ID, v.FECHA,
    SUM(COALESCE(v._CALL_UNICOS,0))            AS CALL_UNICOS,
    SUM(COALESCE(v.WHATSAPP_UNICOS,0))         AS WHATSAPP_UNICOS,
    SUM(COALESCE(v.PREGUNTAS_UNICOS,0))        AS PREGUNTAS_UNICOS,
    SUM(COALESCE(v._CALL_UNICOS,0)+COALESCE(v.WHATSAPP_UNICOS,0)+
        COALESCE(v.PREGUNTAS_UNICOS,0))        AS TOTAL_CONTACTOS
  FROM ``meli-bi-data.WHOWNER.BT_VIS_VIBRANCY_TOTAL_SITES`` v
  INNER JOIN active_dealers ad ON v.CUS_CUST_ID = ad.CUS_CUST_ID
  WHERE v.FECHA BETWEEN '$DATE_FROM' AND '$DATE_TO'
    AND v.VERTICAL = 'Motors' AND v.FRAUDE = 'N/A' AND v.SIT_SITE_ID = 'MLM'
  GROUP BY v.CUS_CUST_ID, v.FECHA
)
SELECT ad.CUS_CUST_ID,
  COALESCE(ad.SEGMENTACION_NSM,'(sin seg)')  AS SEGMENTACION_NSM,
  COALESCE(n.NOMBRE,'(sin nombre)')          AS NOMBRE_CLIENTE,
  COALESCE(ad.ASESOR,'(sin asesor)')         AS ASESOR,
  f.fecha AS FECHA, f.SEM,
  COALESCE(d.CALL_UNICOS,0)       AS CALL_UNICOS,
  COALESCE(d.WHATSAPP_UNICOS,0)   AS WHATSAPP_UNICOS,
  COALESCE(d.PREGUNTAS_UNICOS,0)  AS PREGUNTAS_UNICOS,
  COALESCE(d.TOTAL_CONTACTOS,0)   AS TOTAL_CONTACTOS
FROM active_dealers ad
CROSS JOIN fechas f
LEFT JOIN dealer_names n  ON ad.CUS_CUST_ID = n.CUS_CUST_ID
LEFT JOIN daily d ON ad.CUS_CUST_ID = d.CUS_CUST_ID AND d.FECHA = f.fecha
ORDER BY ad.CUS_CUST_ID, f.fecha
"@

$rows_diario = bq query --use_legacy_sql=false --project_id=meli-bi-data `
    --maximum_bytes_billed=1000000000000 --format=csv --max_rows=200000 $q_diario
$rows_diario | Out-File "$env:TEMP\motors_datos_diarios.csv" -Encoding utf8
Write-Host "  -> $($rows_diario.Count - 1) filas"

# ── 2. Query dashboard diario (agregado) ──────────────────
Write-Host "[2/4] Consultando agregado diario..."

$q_dash = @"
WITH active_dealers AS (
  SELECT DISTINCT CUS_CUST_ID
  FROM ``meli-bi-data.WHOWNER.DM_VIS_SELLER_MO_L1``
  WHERE SIT_SITE_ID = 'MLM' AND VERTICAL = 'MO'
    AND FECHA = (SELECT MAX(FECHA) FROM ``meli-bi-data.WHOWNER.DM_VIS_SELLER_MO_L1``
                 WHERE SIT_SITE_ID = 'MLM' AND VERTICAL = 'MO' AND FECHA <= '$DATE_TO')
)
SELECT v.FECHA,
  CAST(FLOOR(DATE_DIFF(v.FECHA, DATE '$SEM_BASE', DAY)/7)+1 AS INT64) AS SEM,
  FORMAT_DATE('%A', v.FECHA)  AS DIA_SEMANA,
  IF(v.FECHA >= '2026-05-07', 'Post 07-may (WA landing)', 'Pre 07-may') AS PERIODO,
  SUM(COALESCE(v._CALL_UNICOS,0))       AS CALL_UNICOS,
  SUM(COALESCE(v.WHATSAPP_UNICOS,0))    AS WHATSAPP_UNICOS,
  SUM(COALESCE(v.PREGUNTAS_UNICOS,0))   AS PREGUNTAS_UNICOS,
  SUM(COALESCE(v._CALL_UNICOS,0)+COALESCE(v.WHATSAPP_UNICOS,0)+
      COALESCE(v.PREGUNTAS_UNICOS,0))   AS TOTAL_CONTACTOS
FROM ``meli-bi-data.WHOWNER.BT_VIS_VIBRANCY_TOTAL_SITES`` v
INNER JOIN active_dealers ad ON v.CUS_CUST_ID = ad.CUS_CUST_ID
WHERE v.FECHA BETWEEN '$DATE_FROM' AND '$DATE_TO'
  AND v.VERTICAL = 'Motors' AND v.FRAUDE = 'N/A' AND v.SIT_SITE_ID = 'MLM'
GROUP BY v.FECHA, SEM, DIA_SEMANA, PERIODO
ORDER BY v.FECHA
"@

$rows_dash = bq query --use_legacy_sql=false --project_id=meli-bi-data `
    --maximum_bytes_billed=1000000000000 --format=csv --max_rows=500 $q_dash
$rows_dash | Out-File "$env:TEMP\motors_dash_diario.csv" -Encoding utf8
Write-Host "  -> $($rows_dash.Count - 1) filas"

# ── 3. Computar agregados en PowerShell ───────────────────
Write-Host "[3/4] Procesando datos..."

Import-Module ImportExcel -ErrorAction SilentlyContinue

$diario  = Import-Csv "$env:TEMP\motors_datos_diarios.csv"
$dashRaw = Import-Csv "$env:TEMP\motors_dash_diario.csv"

function Sum-Col($rows, $col) { ($rows | ForEach-Object { [double]$_.$col } | Measure-Object -Sum).Sum }

# dash_semanal
$SEM_DAYS = @{1=6;2=7;3=7;4=7;5=7;6=4;7=7;8=7;9=7;10=7}
$semanal = $dashRaw | Group-Object SEM | ForEach-Object {
    $r = $_.Group
    $sem = [int]$_.Name
    $days = if ($SEM_DAYS.ContainsKey($sem)) { $SEM_DAYS[$sem] } else { 7 }
    [PSCustomObject]@{
        SEM             = $sem
        FECHA_INICIO    = ($r | Sort-Object FECHA | Select-Object -First 1).FECHA
        FECHA_FIN       = ($r | Sort-Object FECHA | Select-Object -Last  1).FECHA
        DIAS            = $r.Count
        CALL_UNICOS     = Sum-Col $r 'CALL_UNICOS'
        WHATSAPP_UNICOS = Sum-Col $r 'WHATSAPP_UNICOS'
        PREGUNTAS_UNICOS= Sum-Col $r 'PREGUNTAS_UNICOS'
        TOTAL_CONTACTOS = Sum-Col $r 'TOTAL_CONTACTOS'
        PCT_WHATSAPP    = [math]::Round((Sum-Col $r 'WHATSAPP_UNICOS')/(Sum-Col $r 'TOTAL_CONTACTOS')*100,1)
    }
} | Sort-Object SEM

# dash_dia_semana
$DOW_ORDER = @{Monday=1;Tuesday=2;Wednesday=3;Thursday=4;Friday=5;Saturday=6;Sunday=7}
$dow = $dashRaw | Group-Object DIA_SEMANA | ForEach-Object {
    $ds = $_.Name
    $pre  = @($_.Group | Where-Object { $_.PERIODO -like 'Pre*' })
    $post = @($_.Group | Where-Object { $_.PERIODO -like 'Post*' })
    $waPre  = if ($pre.Count)  { [math]::Round((Sum-Col $pre  'WHATSAPP_UNICOS')/$pre.Count,1)  } else { 0 }
    $waPost = if ($post.Count) { [math]::Round((Sum-Col $post 'WHATSAPP_UNICOS')/$post.Count,1) } else { 0 }
    $totPre  = if ($pre.Count)  { [math]::Round((Sum-Col $pre  'TOTAL_CONTACTOS')/$pre.Count,1)  } else { 0 }
    $totPost = if ($post.Count) { [math]::Round((Sum-Col $post 'TOTAL_CONTACTOS')/$post.Count,1) } else { 0 }
    $tPre = Sum-Col $pre 'TOTAL_CONTACTOS'; $tPost = Sum-Col $post 'TOTAL_CONTACTOS'
    $wPre = Sum-Col $pre 'WHATSAPP_UNICOS'; $wPost = Sum-Col $post 'WHATSAPP_UNICOS'
    [PSCustomObject]@{
        DIA_SEMANA     = $ds
        ORDEN          = if ($DOW_ORDER.ContainsKey($ds)) { $DOW_ORDER[$ds] } else { 9 }
        DIAS_PRE       = $pre.Count
        DIAS_POST      = $post.Count
        WA_PROM_PRE    = $waPre
        WA_PROM_POST   = $waPost
        WA_DELTA_PCT   = if ($waPre -gt 0) { [math]::Round(($waPost-$waPre)/$waPre*100,1) } else { 0 }
        TOTAL_PROM_PRE = $totPre
        TOTAL_PROM_POST= $totPost
        PCT_WA_PRE     = if ($tPre  -gt 0) { [math]::Round($wPre/$tPre*100,1)   } else { 0 }
        PCT_WA_POST    = if ($tPost -gt 0) { [math]::Round($wPost/$tPost*100,1) } else { 0 }
    }
} | Sort-Object ORDEN

# detalle_cust_id por semana
$detalle = $diario | Group-Object CUS_CUST_ID, SEM | ForEach-Object {
    $r = $_.Group; $r0 = $r[0]
    [PSCustomObject]@{
        CUS_CUST_ID      = $r0.CUS_CUST_ID
        SEM              = [int]$r0.SEM
        SEGMENTACION_NSM = $r0.SEGMENTACION_NSM
        NOMBRE_CLIENTE   = $r0.NOMBRE_CLIENTE
        ASESOR           = $r0.ASESOR
        TOTAL_CONTACTOS  = Sum-Col $r 'TOTAL_CONTACTOS'
    }
} | Sort-Object CUS_CUST_ID, SEM

# ── 4. Generar HTML ───────────────────────────────────────
Write-Host "[4/4] Generando HTML..."

$dd  = $dashRaw  | ConvertTo-Json -Compress -Depth 3
$ds  = $semanal  | ConvertTo-Json -Compress -Depth 3
$dow_json = $dow | ConvertTo-Json -Compress -Depth 3
$det = $detalle  | ConvertTo-Json -Compress -Depth 3

# Calcular rango de semanas para header
$maxSem = ($semanal | Measure-Object SEM -Maximum).Maximum
$lastDate = ($dashRaw | Sort-Object FECHA | Select-Object -Last 1).FECHA

# Leer template HTML y reemplazar datos
$templatePath = "$env:USERPROFILE\motors-dashboard\template.html"
$outputPath   = "$env:USERPROFILE\motors-dashboard\index.html"

$html = [System.IO.File]::ReadAllText($templatePath, [System.Text.Encoding]::UTF8)
$html = $html.Replace('__DATA_DIARIO__',  $dd)
$html = $html.Replace('__DATA_SEMANAL__', $ds)
$html = $html.Replace('__DATA_DOW__',     $dow_json)
$html = $html.Replace('__DATA_DETALLE__', $det)
$html = $html.Replace('__LAST_UPDATE__',  $TODAY)
$html = $html.Replace('__DATE_TO__',      $lastDate)
$html = $html.Replace('__MAX_SEM__',      "Sem $maxSem")

[System.IO.File]::WriteAllText($outputPath, $html, [System.Text.Encoding]::UTF8)
$size = [math]::Round((Get-Item $outputPath).Length/1MB, 2)
Write-Host "  -> index.html generado ($size MB)"

# ── 5. Git commit y push ──────────────────────────────────
Write-Host "`nHaciendo push a GitHub..."
$gitPath = "C:\Program Files\Git\bin\git.exe"
& $gitPath add index.html
& $gitPath commit -m "Daily update $TODAY"
& $gitPath push origin main
Write-Host "`n✅ Dashboard actualizado: https://$(& $gitPath remote get-url origin | Split-Path -Leaf)" -ForegroundColor Green
