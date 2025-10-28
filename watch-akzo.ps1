# =======================
# watch-akzo.ps1
# Scans Raw\AkzoNobel for new PDFs, sends to OpenAI, writes summary to Summaries\2025Q3
# =======================

# ---- CONFIG ----
$RawDir = "C:\Users\MORAISN\OneDrive - Hexion Inc\01 - Projects\01 - Earnings Release Summary\Raw\AkzoNobel"
$SummaryDir = "C:\Users\MORAISN\OneDrive - Hexion Inc\01 - Projects\01 - Earnings Release Summary\Summaries\2025Q3"
$StateFile = "C:\Users\MORAISN\OneDrive - Hexion Inc\01 - Projects\01 - Earnings Release Summary\_automation\processed.csv"
$Model = "gpt-3.5-turbo"   # choose from your listed models

# ---- PROMPT (your standard format) ----
$Prompt = @"
You are an equity analyst. Summarize AkzoNobel Q3 2025 based on the attached report (state “n/a” if data is missing).

Return EXACTLY:
1) Overview (≤40 words)
2) Tailwinds / Headwinds / Positioning (3 bullets)
3) Regional Market Outlook (EU, AMR, APAC/MEAI)
4) Market Outlook by Segment and Year (table: segments = Decorative Paints, Performance Coatings; columns = 2025, 2026)

Also include: revenue, margins, EPS YoY (if available), key segment highlights, cash flow, leverage, guidance, and risks (pricing, China, FX).
"@

function Ensure-Paths {
  if (-not (Test-Path $RawDir)) { New-Item -ItemType Directory -Path $RawDir | Out-Null }
  if (-not (Test-Path $SummaryDir)) { New-Item -ItemType Directory -Path $SummaryDir | Out-Null }
  if (-not (Test-Path $StateFile)) { "filename,processed_at" | Out-File -FilePath $StateFile -Encoding utf8 }
}

function Is-Processed($filePath) {
  $name = [IO.Path]::GetFileName($filePath)
  Select-String -Path $StateFile -Pattern ("^" + [regex]::Escape($name) + ",") -SimpleMatch | ForEach-Object { return $true }
  return $false
}

function Mark-Processed($filePath) {
  $name = [IO.Path]::GetFileName($filePath)
  $ts = (Get-Date).ToString("s")
  Add-Content -Path $StateFile -Value "$name,$ts"
}

function Upload-To-OpenAI([string]$filePath) {
  $apiKey = $env:OPENAI_API_KEY
  if ([string]::IsNullOrWhiteSpace($apiKey)) { throw "OPENAI_API_KEY environment variable is not set." }

  # Ensure TLS 1.2+
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

  Add-Type -AssemblyName System.Net.Http
  $client = [System.Net.Http.HttpClient]::new()
  $client.DefaultRequestHeaders.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $apiKey)

  $content = [System.Net.Http.MultipartFormDataContent]::new()
  # Correct purpose for files used by Responses/Assistants API:
  $content.Add([System.Net.Http.StringContent]::new("assistants"), "purpose")

  # >>> The critical line: wrap bytes as a single argument <<<
  [byte[]]$bytes = [System.IO.File]::ReadAllBytes($filePath)
  $fileContent = [System.Net.Http.ByteArrayContent]::new($bytes)   # no splat

  $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/pdf")
  $fileName = [IO.Path]::GetFileName($filePath)
  $content.Add($fileContent, "file", $fileName)

  $resp = $client.PostAsync("https://api.openai.com/v1/files", $content).Result
  $body = $resp.Content.ReadAsStringAsync().Result

  $client.Dispose()
  $content.Dispose()
  $fileContent.Dispose()

  if (-not $resp.IsSuccessStatusCode) {
    throw "Upload failed ($($resp.StatusCode)): $body"
  }

  return (ConvertFrom-Json $body).id
}

function Create-Response([string]$fileId, [string]$prompt) {
  # TLS
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $apiKey = $env:OPENAI_API_KEY
  if ([string]::IsNullOrWhiteSpace($apiKey)) { throw "OPENAI_API_KEY is not set." }

  # Build the JSON body using the Responses API "input" param with mixed content parts
  $bodyObj = @{
    model       = $Model          # e.g., "gpt-4.1"
    temperature = 0.2
    input       = @(
      @{
        role    = "user"
        content = @(
          @{ type = "input_text"; text = $prompt },
          @{ type = "input_file"; file_id = $fileId }
        )
      }
    )
  }
  $json = $bodyObj | ConvertTo-Json -Depth 10

  # Use HttpClient & UTF-8 JSON
  Add-Type -AssemblyName System.Net.Http
  $client = [System.Net.Http.HttpClient]::new()
  $client.DefaultRequestHeaders.Authorization =
    [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $apiKey)
  $content = [System.Net.Http.StringContent]::new($json, [System.Text.Encoding]::UTF8, "application/json")

  $maxAttempts = 3
  for ($i=1; $i -le $maxAttempts; $i++) {
    try {
      $resp = $client.PostAsync("https://api.openai.com/v1/responses", $content).Result
      $text = $resp.Content.ReadAsStringAsync().Result
      if (-not $resp.IsSuccessStatusCode) {
        throw "HTTP $($resp.StatusCode): $text"
      }
      $obj = $text | ConvertFrom-Json
      if ($obj.output_text) { return $obj.output_text.Trim() }
      try { return $obj.output[0].content[0].text.Trim() } catch { return $text }
    } catch {
      if ($i -eq $maxAttempts) { throw "Responses API failed after $maxAttempts attempts: $($_.Exception.Message)" }
      Start-Sleep -Seconds (5 * $i)
      # recreate HttpContent each retry
      $content.Dispose()
      $content = [System.Net.Http.StringContent]::new($json, [System.Text.Encoding]::UTF8, "application/json")
    }
  }

  $content.Dispose()
  $client.Dispose()
}
function Save-Summary([string]$sourceFile, [string]$text) {
  $baseName = [IO.Path]::GetFileNameWithoutExtension($sourceFile)
  $outName  = "$baseName-summary.md"
  $outPath  = Join-Path $SummaryDir $outName
  $text | Out-File -FilePath $outPath -Encoding utf8
  return $outPath
}

# ---- Main one-shot run ----
Ensure-Paths
$files = Get-ChildItem -Path $RawDir -File -Filter *.pdf | Sort-Object LastWriteTime -Descending
foreach ($f in $files) {
  if (Is-Processed $f.FullName) { continue }
  try {
    Write-Host "Processing $($f.Name)..."
    $fileId  = Upload-To-OpenAI $f.FullName
    Write-Host "Uploaded. file_id=$fileId"
    $summary = Create-Response $fileId $Prompt
    $outPath = Save-Summary $f.FullName $summary
    Mark-Processed $f.FullName
    Write-Host "Summary saved -> $outPath"
  } catch {
    Write-Host "ERROR processing $($f.Name): $($_.Exception.Message)"
  }
}