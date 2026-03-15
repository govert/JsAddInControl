param(
    [string]$WorkbookPath,
    [switch]$RebuildExcelDna
)

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$officeJsRoot = Join-Path $root "OfficeJsAddIn"
$excelDnaRoot = Join-Path $root "ExcelDnaAddIn"
$manifestPath = Join-Path $officeJsRoot "manifest.xml"
$xllPath = Join-Path $excelDnaRoot "bin\\Debug\\net48\\ExcelDnaAddIn-AddIn64.xll"
$pfxPath = Join-Path $officeJsRoot "localhost-devcert.pfx"
$cerPath = Join-Path $officeJsRoot "localhost-devcert.cer"
$manifestXml = [xml](Get-Content $manifestPath)
$manifestMetadata = @{
    id = $manifestXml.OfficeApp.Id
    officeAppType = $manifestXml.OfficeApp.GetAttribute("type", "http://www.w3.org/2001/XMLSchema-instance")
    version = $manifestXml.OfficeApp.Version
    manifestType = "xml"
} | ConvertTo-Json -Compress

Push-Location $officeJsRoot
try {
    if (-not (Test-Path ".\\node_modules")) {
        npm install | Out-Host
    }

    dotnet dev-certs https -ep $pfxPath -p testingcertpass | Out-Host
    $localhostCert = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq "CN=localhost" } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1
    if ($localhostCert) {
        Export-Certificate -Cert $localhostCert -FilePath $cerPath -Force | Out-Null
        certutil -user -addstore Root $cerPath | Out-Host
    }
    npx office-addin-dev-settings register $manifestPath | Out-Host

    Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue |
        ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }

    Start-Process -FilePath "node" -ArgumentList ".\\server.js" -WorkingDirectory $officeJsRoot -WindowStyle Hidden | Out-Null
    Start-Sleep -Seconds 3

    $sideloadWorkbook = node -e "const { OfficeApp } = require('office-addin-manifest'); const { generateSideloadFile } = require('office-addin-dev-settings'); (async () => { const manifest = JSON.parse(process.argv[1]); const file = await generateSideloadFile(OfficeApp.Excel, manifest); console.log(file); })().catch((error) => { console.error(error); process.exit(1); });" $manifestMetadata
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to generate the Office JS sideload workbook."
    }

    if (-not $WorkbookPath) {
        $WorkbookPath = (($sideloadWorkbook | Out-String).Trim().Split([Environment]::NewLine, [System.StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Last 1).Trim()
    }
}
finally {
    Pop-Location
}

Get-Process EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "Excel add-in *" } |
    ForEach-Object { $_.CloseMainWindow() | Out-Null }

Start-Sleep -Seconds 2

Push-Location $excelDnaRoot
try {
    if ($RebuildExcelDna -or -not (Test-Path $xllPath)) {
        dotnet build .\ExcelDnaAddIn.csproj | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to build the Excel-DNA add-in."
        }
    }
}
finally {
    Pop-Location
}

try {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $excel = [Microsoft.VisualBasic.Interaction]::GetObject("", "Excel.Application")
}
catch {
    $excel = New-Object -ComObject Excel.Application
}

$excel.Visible = $true
$excel.DisplayAlerts = $false

$registered = $excel.RegisterXLL($xllPath)
if (-not $registered) {
    throw "Excel did not register the Excel-DNA add-in."
}

$alreadyOpen = $false
foreach ($workbook in @($excel.Workbooks)) {
    if ($workbook.FullName -eq $WorkbookPath) {
        $alreadyOpen = $true
        break
    }
}

if (-not $alreadyOpen) {
    $excel.Workbooks.Open($WorkbookPath) | Out-Null
}

[pscustomobject]@{
    ExcelVersion = $excel.Version
    RegisteredXll = $registered
    OfficeJsWorkbook = $WorkbookPath
    ExcelProcessId = (Get-Process -Name EXCEL | Sort-Object StartTime | Select-Object -Last 1 -ExpandProperty Id)
}
