param(
    [switch]$Sideload
)

$projectRoot = Split-Path -Parent $PSScriptRoot

Push-Location $projectRoot
try {
    npm install
    npm run dev-cert

    if ($Sideload) {
        Start-Process powershell -ArgumentList '-NoExit', '-Command', "Set-Location '$projectRoot'; npm start"
        Start-Sleep -Seconds 5
        npm run sideload
    } else {
        npm start
    }
}
finally {
    Pop-Location
}
