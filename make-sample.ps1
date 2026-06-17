# Convierte data/slim.csv -> sample-data.js  (window.SAMPLE_ROWS)
# Para mantener liviano el demo, toma las ULTIMAS N fechas.
param([int]$LastDays = 5)

$csv = "C:\Users\kcast\planning-app\data\slim.csv"
$out = "C:\Users\kcast\planning-app\sample-data.js"

$rows = Import-Csv -Path $csv -Encoding UTF8
Write-Output ("Filas totales: {0}" -f $rows.Count)

$fechas = $rows | Where-Object { $_.Fecha } | Select-Object -ExpandProperty Fecha | Sort-Object -Unique
$keep = $fechas | Select-Object -Last $LastDays
Write-Output ("Fechas en demo: " + ($keep -join ', '))
$keepSet = @{}; $keep | ForEach-Object { $keepSet[$_] = $true }

$sel = $rows | Where-Object { $keepSet.ContainsKey($_.Fecha) }
Write-Output ("Filas en demo: {0}" -f $sel.Count)

$recs = foreach ($r in $sel) {
  [ordered]@{
    id          = $r.idFicha
    fecha       = $r.Fecha
    hora        = $r.Hora
    campana     = $r.Campana
    tipoOfrec   = $r.TipoOfrecimiento
    tipoVenta   = $r.TipoVenta
    lineaUpsell = $r.LineaUpselling
    lineaMigrar = $r.LineaMigrar
    plan        = $r.PlanTarifario
    cargoFijo   = $r.CargoFijo
    ganancia    = $r.Ganancia
    cantFamilia = $r.CantFamiliaAdic
    tipoProducto= $r.TipoProducto
    cantAcces   = $r.CantAccesorios
    montoFinan  = $r.MontoFinanEquipos
    supervisor  = $r.Supervisor
    docVendedor = $r.DocVendedor
    vendedor    = $r.Vendedor
    estado      = $r.Estado
    subEstado   = $r.SubEstado
  }
}
$json = $recs | ConvertTo-Json -Depth 3 -Compress
$sw = New-Object System.IO.StreamWriter($out, $false, (New-Object System.Text.UTF8Encoding($false)))
$sw.Write("window.SAMPLE_ROWS = ")
$sw.Write($json)
$sw.Write(";")
$sw.Close()
Write-Output ("Escrito: {0}  ({1:N1} MB)" -f $out, ((Get-Item $out).Length/1MB))
