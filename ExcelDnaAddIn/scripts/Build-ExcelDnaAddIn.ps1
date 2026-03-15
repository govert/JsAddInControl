param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

$projectRoot = Split-Path -Parent $PSScriptRoot

Push-Location $projectRoot
try {
    dotnet build .\ExcelDnaAddIn.csproj -c $Configuration
}
finally {
    Pop-Location
}
