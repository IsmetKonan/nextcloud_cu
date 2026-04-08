#
# Made by Ismet Konan for ece24
# Last Edited 08.04.2026
#

$VERSION     = "1.0.1"
$DEKO        = "-----------------------------------------------------------------"
$EMPTY_LINE  = "                                                                 "

$Host.UI.RawUI.ForegroundColor = 'Blue'
Write-Host $EMPTY_LINE
Write-Host $DEKO
Write-Host "    ____                    __     __ __"
Write-Host "   /  _/________ ___  ___  / /_   / //_/___  ____  ____ _____"
Write-Host "   / // ___/ __ `__ \/ _ \/ __/  / ,< / __ \/ __ \/ __ `/ __ \"
Write-Host " _/ /(__  ) / / / / /  __/ /_   / /| / /_/ / / / / /_/ / / / /"
Write-Host "/___/____/_/ /_/ /_/\___/\__/  /_/ |_\____/_/ /_/\__,_/_/ /_/"
Write-Host $DEKO
Write-Host "CC Ismet Konan"
Write-Host "$VERSION starting up ..."
Write-Host $DEKO

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Import-Module ImportExcel

$localDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path -Path $localDir -ChildPath "user_creation_log.txt"

Add-Content -Path $logFile -Value "----- User Creation Log: $(Get-Date) -----`n"

function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO"   
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    Add-Content -Path $logFile -Value $logEntry
}

$adminUser = Read-Host "Enter username"
$appPasswordSecure = Read-Host "Enter password" -AsSecureString
$appPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($appPasswordSecure)
)

Write-Host "Starting user creation" -ForegroundColor Cyan

$pair = "$($adminUser):$($appPassword)"
$encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))

$headers = @{
    Authorization    = "Basic $encodedCreds"
    "OCS-APIRequest" = "true"
}

$users = Import-Excel ".\user.xlsx"

$baseUrl = $users[0].URL

if (-not $baseUrl) {
    Write-Host "URL in F2 is empty!" -ForegroundColor Red
    exit
}

$baseUrl = $baseUrl.Trim()

if ($baseUrl -match "^http:([^/])") {
    $baseUrl = $baseUrl -replace "^http:", "http://"
}

$baseUrl = $baseUrl.TrimEnd('/')

Write-Host "Using base URL: $baseUrl" -ForegroundColor Cyan

$testUri = "$baseUrl/ocs/v1.php/cloud/users?format=json"

try {
    Invoke-RestMethod -Uri $testUri -Method Get -Headers $headers
    Write-Host "Login successful!" -ForegroundColor Green
}
catch {
    Write-Host "LOGIN FAILED!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    exit
}

$localDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$counter = 1

foreach ($user in $users) {

    # Skip if userid or password is missing
    if (-not $user.userid -or -not $user.password) {
        Write-Host "Skipping user: missing userid or password." -ForegroundColor Yellow
        Write-Log "Skipped user: missing userid or password for entry in Excel" "WARNING"
        continue
    }

    # Check if user already exists
    $checkUri = "$baseUrl/ocs/v1.php/cloud/users/$($user.userid)?format=json"
    $userExists = $false
    try {
        $resp = Invoke-RestMethod -Uri $checkUri -Method Get -Headers $headers
        if ($resp.ocs.meta.statuscode -eq 100) { 
            $userExists = $true
        }
    } catch {
        # If GET fails (404), user does not exist
        $userExists = $false
    }

    if ($userExists) {
        Write-Host "User $($user.userid) already exists, skipping." -ForegroundColor Yellow
        Write-Log "User already exists: $($user.userid)" "WARNING"
        continue
    }

    # Create user
    $uri = "$baseUrl/ocs/v1.php/cloud/users?format=json"
    $body = @{
        userid      = $user.userid
        password    = $user.password
        displayName = $user.displayName
        email       = $user.email
    }

    try {
        # API call and capture full response
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -ContentType "application/x-www-form-urlencoded"
        
        Write-Host "API Response: $($response | ConvertTo-Json -Depth 5)" -ForegroundColor Cyan

        # Check if creation succeeded
        if ($response.ocs.meta.statuscode -eq 100) {
            Write-Host "User created: $($user.userid)" -ForegroundColor Green
            Write-Log "User created: $($user.userid)" "SUCCESS"

            # Create Outlook MSG
            $outlook = New-Object -ComObject Outlook.Application
            $mail = $outlook.CreateItem(0)
            $mail.Subject = "Ihr Account wurde erstellt, $($user.displayName)!"
            $mail.To = $user.email
            $mail.Body = "Hallo $($user.displayName)!`n`n" +
                         "Ihr Account wurde erstellt, `n`n" + 
                         "die URL zum Login lautet: $baseUrl`n`n" +
                         "User ID: $($user.userid)`nPassword: $($user.password)`nBitte ändern Sie Ihr Passwort nach dem ersten Login!`n`n" +
                         "Viele Grüße,`nIhr ece24 Team"

            $fileName = "mail_$($user.userid).msg"
            $fullPath = Join-Path -Path $localDir -ChildPath $fileName
            $mail.SaveAs($fullPath, 3)
            Write-Host "MSG erstellt für: $fullPath" -ForegroundColor Cyan

            $counter++
        } else {
            Write-Host "Failed to create user: $($user.userid)" -ForegroundColor Red
            Write-Host "API returned status: $($response.ocs.meta.status) ($($response.ocs.meta.statuscode))" -ForegroundColor Yellow
            Write-Log "Failed to create user: $($user.userid) - API status: $($response.ocs.meta.status) ($($response.ocs.meta.statuscode))" "ERROR"
        }

    } catch {
        Write-Host "Failed to create user: $($user.userid)" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log "Failed to create user: $($user.userid) - $($_.Exception.Message)" "ERROR"
    }
}

Write-Host $DEKO
Write-Host 'All Done!' -ForegroundColor Green
Write-Host $DEKO
Write-Host $DEKO
pause > $null
Write-Host $DEKO
