# ===============================
# CONFIG
# ===============================
$uri       = "https://webhook.end.point"
$jwtSecret = "JWT secret"

$filePath  = "path to file"
$to        = "te e-mail address"
$subject   = "Subject"
$body      = "Body"

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
