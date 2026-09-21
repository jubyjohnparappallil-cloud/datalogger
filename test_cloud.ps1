# End-to-end check of cloud mode: batched upload, job, download, cleanup.
$base = "http://127.0.0.1:5001"
$dir = "C:\Users\Juby John\Downloads\Aramax WH\Aramax WH"
$files = Get-ChildItem -LiteralPath $dir -Filter *.pdf | Select-Object -First 6

$health = (Invoke-WebRequest -Uri "$base/healthz" -UseBasicParsing).Content
"HEALTH: $health"

$batch = (curl.exe -s -X POST "$base/api/batch" | ConvertFrom-Json)
"BATCH: $($batch.batch_id)"

# Upload in two chunks, the way the browser does it.
$chunks = @(, $files[0..2]) + @(, $files[3..5])
foreach ($chunk in $chunks) {
    $args = @("-s", "-X", "POST", "$base/api/batch/$($batch.batch_id)")
    foreach ($f in $chunk) { $args += @("-F", "files=@$($f.FullName)") }
    $r = (& curl.exe @args | ConvertFrom-Json)
    "UPLOAD: added $($r.added), total on server $($r.total)"
}

$job = (curl.exe -s -X POST "$base/api/process" -F "batch_id=$($batch.batch_id)" | ConvertFrom-Json)
"JOB: $($job.job_id)"

for ($i = 0; $i -lt 90; $i++) {
    Start-Sleep -Seconds 2
    $s = (curl.exe -s "$base/api/status/$($job.job_id)" | ConvertFrom-Json)
    if ($s.status -eq "done" -or $s.status -eq "error") { break }
}
"STATUS: $($s.status) - $($s.message)"
if ($s.status -ne "done") { "ERROR: $($s.error)"; exit 1 }
"STATS: loggers=$($s.stats.loggers) rows=$($s.stats.rows)"
"RANGE: $($s.stats.start) -> $($s.stats.end)"
"COLUMNS: $($s.preview_loggers -join ', ')"
$row = $s.preview[0]
"ROW 1: $($row.DATE) $($row.Time) | temp DL-1=$($row.temp.'DL-1') | rh DL-1=$($row.rh.'DL-1')"

$out = "C:\data separa\output\_cloudtest.xlsx"
curl.exe -s -o $out "$base/download/$($s.download)"
"DOWNLOAD: $([math]::Round((Get-Item $out).Length/1KB,0)) KB"

$leftover = @(Get-ChildItem "C:\data separa\uploads" -Directory -ErrorAction SilentlyContinue)
"UPLOAD CLEANUP: $($leftover.Count) batch folder(s) left on disk"

# The folder-path route must be refused in cloud mode.
$blocked = (curl.exe -s -X POST "$base/api/process" -F "folder=C:\Users" | ConvertFrom-Json)
"FOLDER PATH IN CLOUD: $($blocked.error)"
