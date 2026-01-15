$log = "C:\ztemp\log\send-report.log"
Start-Transcript -Path $log -Append

# ===============================
# LOAD CONFIG (JSON)
# ===============================
#$configPath = "C:\ztemp\sendmail\lpr-report.json"

# Path of the running EXE (when compiled) or script (when running as .ps1)
$baseDir =
    if ($MyInvocation.MyCommand.Path) {
        Split-Path -Parent $MyInvocation.MyCommand.Path   # .ps1
    }
    else {
        [AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')  # .exe (ps2exe)
    }

$configPath = Join-Path $baseDir "lpr-report.json"

if (-not (Test-Path -Path $configPath -PathType Leaf)) {
    throw "Config file not found next to executable/script: $configPath (baseDir=$baseDir)"
}

$config = Get-Content -Path $configPath -Raw -Encoding UTF8 | ConvertFrom-Json




if (-not (Test-Path -Path $configPath -PathType Leaf)) {
    throw "Config file not found: $configPath"
}

$config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

# Basic validation
if ([string]::IsNullOrWhiteSpace($config.endpoint))  { throw "Config missing 'endpoint'" }
if ([string]::IsNullOrWhiteSpace($config.jwtSecret)) { throw "Config missing 'jwtSecret'" }
if ([string]::IsNullOrWhiteSpace($config.reportLocation))  { throw "Config missing 'reportLocation'" }
if (-not $config.recipients -or $config.recipients.Count -lt 1) { throw "Config missing 'recipients' array" }


$uri       = [string]$config.endpoint
$jwtSecret = [string]$config.jwtSecret
$recipients = @($config.recipients | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$reportFolder = [string]$config.reportLocation
if ($recipients.Count -lt 1) { throw "No valid recipient emails found in config 'recipients'." }

# ===============================
# File Creation
# ===============================
Connect-ManagementServer

# Define yesterday time range
$start = (Get-Date).Date.AddDays(-1)   # Yesterday 00:00
$end   = (Get-Date).Date              # Today 00:00

$startUtc = $start.ToUniversalTime()
$endUtc   = $end.ToUniversalTime()

# Format date for filename (YYYY-MM-DD)
$dateString = $start.ToString("yyyy-MM-dd")

# Output file path
#$outputFile = $partialPath+"LPRReport_$dateString.csv"
$outputFile = Join-Path $reportFolder "LPRReport_$dateString.csv"

# Prefer explicit Start/End boundaries (yesterday 00:00 -> today 00:00)
$splat = @{
    StartTime = $start
    EndTime   = $end
}


#$events = Get-VmsLprEvent @splat | Select-Object Timestamp, CustomTag, ObjectValue, Message, SourceName, RuleType
$events = Get-VmsLprEvent @splat | Select-Object `
  @{ Name = "TimestampLocal"; Expression = { $_.Timestamp.ToLocalTime() } },
  @{ Name = "TimestampUtc";   Expression = { $_.Timestamp } },
  CustomTag,
  ObjectValue,
  Message,
  SourceName,
  RuleType


# Export to CSV
$events | Export-Csv $outputFile -NoTypeInformation -Encoding UTF8

Disconnect-ManagementServer

# ===============================
# Pick newest report file (optional, but keeping your logic)
# ===============================
$folder  = "C:\ztemp\rep"
$pattern = "LPRReport_*.csv"

$latest = Get-ChildItem -Path $folder -Filter $pattern -File |
          Sort-Object LastWriteTime -Descending |
          Select-Object -First 1

if (-not $latest) { throw "No file found matching $pattern in $folder" }

$filePath = $latest.FullName

# ===============================
# Email content
# ===============================
$subject = "LPR Report DistriFresh $dateString"
$body    = "Hi, report from $dateString attached."

# ===============================
# JWT GENERATION (HS256)
# ===============================
function ConvertTo-Base64Url([byte[]] $bytes) {
    [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
}

function New-JwtHS256 {
    param(
        [Parameter(Mandatory)] [string] $Secret,
        [Parameter(Mandatory)] [hashtable] $Payload
    )

    $header = @{ alg="HS256"; typ="JWT" }

    $headerJson  = ($header  | ConvertTo-Json -Compress)
    $payloadJson = ($Payload | ConvertTo-Json -Compress)

    $headerB64  = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($headerJson))
    $payloadB64 = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($payloadJson))

    $unsigned = "$headerB64.$payloadB64"

    $keyBytes = [Text.Encoding]::UTF8.GetBytes($Secret)
    $hmac = [System.Security.Cryptography.HMACSHA256]::new($keyBytes)

    $sigBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($unsigned))
    $sigB64   = ConvertTo-Base64Url $sigBytes

    "$unsigned.$sigB64"
}

$now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$jwt = New-JwtHS256 -Secret $jwtSecret -Payload @{
    iat = $now
    exp = $now + 300   # valid for 5 minutes
}

# ===============================
# MULTIPART FORM UPLOAD (PS 5.1)
# ===============================
Add-Type -AssemblyName System.Net.Http

$client = [System.Net.Http.HttpClient]::new()
$client.DefaultRequestHeaders.Authorization =
    [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $jwt)

$form = [System.Net.Http.MultipartFormDataContent]::new()

# text fields
# If your endpoint supports multiple recipients as a single comma-separated string:
$to = ($recipients -join ",")

$form.Add([System.Net.Http.StringContent]::new($to),      "to")
$form.Add([System.Net.Http.StringContent]::new($subject), "subject")
$form.Add([System.Net.Http.StringContent]::new($body),    "body")

# file field
$fileBytes   = [System.IO.File]::ReadAllBytes($filePath)
$fileContent = [System.Net.Http.ByteArrayContent]::new($fileBytes)
$fileContent.Headers.ContentType =
    [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")

$form.Add($fileContent, "file", [System.IO.Path]::GetFileName($filePath))

try {
    $response = $client.PostAsync($uri, $form).Result
    $respText = $response.Content.ReadAsStringAsync().Result

    "HTTP $([int]$response.StatusCode) $($response.ReasonPhrase)"
    $respText

    if (-not $response.IsSuccessStatusCode) {
        throw "Request failed with HTTP $([int]$response.StatusCode): $respText"
    }
}
finally {
    $form.Dispose()
    $client.Dispose()
}

Stop-Transcript
