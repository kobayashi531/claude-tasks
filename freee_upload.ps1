. "$PSScriptRoot\freee_config.ps1"

$TokenFile = "$PSScriptRoot\freee_token.json"

# ── アップロードするファイルを指定 ──────────────────────────────
$FilePath = Read-Host "アップロードするファイルのパスを入力してください"

# ── Content-Type を拡張子から自動判定 ──────────────────────────
$ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
$mimeType = switch ($ext) {
    ".png"  { "image/png" }
    ".jpg"  { "image/jpeg" }
    ".jpeg" { "image/jpeg" }
    ".pdf"  { "application/pdf" }
    default { "application/octet-stream" }
}

$TokenUrl = "https://accounts.secure.freee.co.jp/public_api/token"

function Refresh-Token($refreshToken) {
    $resp = Invoke-RestMethod -Method Post -Uri $TokenUrl `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            grant_type    = "refresh_token"
            client_id     = $ClientId
            client_secret = $ClientSecret
            refresh_token = $refreshToken
        }
    return $resp
}

function Save-Token($tokenResp) {
    $data = @{
        access_token  = $tokenResp.access_token
        refresh_token = $tokenResp.refresh_token
        expires_at    = (Get-Date).AddSeconds($tokenResp.expires_in - 60).ToString("o")
    }
    $data | ConvertTo-Json | Set-Content -Path $TokenFile -Encoding UTF8
}

$accessToken = $null

if (Test-Path $TokenFile) {
    $saved = Get-Content $TokenFile -Raw | ConvertFrom-Json
    $expiresAt = [datetime]::Parse($saved.expires_at)

    if ((Get-Date) -lt $expiresAt) {
        $accessToken = $saved.access_token
        Write-Host "Using saved token (valid until $expiresAt)"
    } else {
        Write-Host "Token expired. Refreshing..."
        try {
            $tokenResp = Refresh-Token $saved.refresh_token
            Save-Token $tokenResp
            $accessToken = $tokenResp.access_token
            Write-Host "Token refreshed OK"
        } catch {
            Write-Host "Refresh failed. Re-authenticating..."
        }
    }
}

if (-not $accessToken) {
    $authUrl = "https://accounts.secure.freee.co.jp/public_api/authorize" +
               "?client_id=$ClientId&redirect_uri=$([Uri]::EscapeDataString($RedirectUri))" +
               "&response_type=code&prompt=select_company"

    Write-Host "Opening browser for authentication..."
    Start-Process $authUrl

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("http://localhost:8080/")
    $listener.Start()
    Write-Host "Waiting for auth in browser..."

    $context  = $listener.GetContext()
    $rawUrl   = $context.Request.Url.Query
    $query    = $rawUrl.TrimStart("?")
    $authCode = ($query.Split("&") | Where-Object { $_ -like "code=*" }) -replace "code=", ""

    $html = "<html><body><h2>Auth OK! You can close this tab.</h2></body></html>"
    $buf  = [System.Text.Encoding]::UTF8.GetBytes($html)
    $context.Response.ContentLength64 = $buf.Length
    $context.Response.OutputStream.Write($buf, 0, $buf.Length)
    $context.Response.Close()
    $listener.Stop()

    if (-not $authCode) { Write-Error "Auth code not received"; Read-Host "Press Enter to close"; exit 1 }

    $tokenResp = Invoke-RestMethod -Method Post -Uri $TokenUrl `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            grant_type    = "authorization_code"
            client_id     = $ClientId
            client_secret = $ClientSecret
            code          = $authCode
            redirect_uri  = $RedirectUri
        }

    Save-Token $tokenResp
    $accessToken = $tokenResp.access_token
    Write-Host "Access token OK (saved for next time)"
}

if (-not (Test-Path $FilePath)) {
    Write-Error "File not found: $FilePath"
    Read-Host "Press Enter to close"
    exit 1
}

$boundary = [System.Guid]::NewGuid().ToString()
$fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
$fileName  = [System.IO.Path]::GetFileName($FilePath)

$bodyLines = [System.Collections.Generic.List[byte]]::new()
function Add-Text([string]$text) {
    $bodyLines.AddRange([System.Text.Encoding]::UTF8.GetBytes($text))
}

Add-Text "--$boundary`r`n"
Add-Text "Content-Disposition: form-data; name=`"company_id`"`r`n`r`n"
Add-Text "$CompanyId`r`n"
Add-Text "--$boundary`r`n"
Add-Text "Content-Disposition: form-data; name=`"receipt`"; filename=`"$fileName`"`r`n"
Add-Text "Content-Type: $mimeType`r`n`r`n"
$bodyLines.AddRange($fileBytes)
Add-Text "`r`n--$boundary--`r`n"

Write-Host "Uploading: $fileName"
$uploadResp = Invoke-RestMethod -Method Post `
    -Uri "https://api.freee.co.jp/api/1/receipts" `
    -Headers @{ Authorization = "Bearer $accessToken" } `
    -ContentType "multipart/form-data; boundary=$boundary" `
    -Body $bodyLines.ToArray()

$receiptId = $uploadResp.receipt.id
Write-Host "Upload SUCCESS! File box ID: $receiptId"
Read-Host "Press Enter to close"