param(
    [string]$Path1 = ".\formulaire_création_VM_6.xlsx",
    [string]$Path2 = ".\formulaire_création_VM_vm_app_dev_1 (5).xlsx"
)

# Ensure ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

function Compare-ExcelSheet {
    param(
        [string]$SheetName,
        [int]$HeaderRow = 1,
        [int]$StartColumn = 1,
        [int]$EndColumn = 10
    )
    
    Write-Host "`n=== Comparing Sheet: $SheetName ===" -ForegroundColor Cyan
    
    $Data1 = Import-Excel -Path $Path1 -WorksheetName $SheetName -HeaderRow $HeaderRow -StartColumn $StartColumn -EndColumn $EndColumn -ErrorAction SilentlyContinue
    $Data2 = Import-Excel -Path $Path2 -WorksheetName $SheetName -HeaderRow $HeaderRow -StartColumn $StartColumn -EndColumn $EndColumn -ErrorAction SilentlyContinue

    if (-not $Data1 -and -not $Data2) {
        Write-Host "Sheet '$SheetName' is empty or could not be found." -ForegroundColor Yellow
        return
    }

    $Diff = Compare-Object -ReferenceObject $Data1 -DifferenceObject $Data2 -PassThru
    if ($Diff) {
        $Diff | Select-Object SideIndicator, * | Format-Table -AutoSize
    } else {
        Write-Host "No differences found in $SheetName." -ForegroundColor Green
    }
}

# 1. Tab 1: Principale (Key/Value pairs starting from Row 6)
Compare-ExcelSheet -SheetName "Principale" -HeaderRow 6 -StartColumn 1 -EndColumn 2

# 2. Tab 2: Publication VM (Software Stack starting at Row 16)
Compare-ExcelSheet -SheetName "Publication VM" -HeaderRow 16 -StartColumn 4 -EndColumn 6

# 3. Tab 3: Informations liées au service (Flow Matrix starting at Row 12)
Compare-ExcelSheet -SheetName "Informations liées au service" -HeaderRow 12 -StartColumn 3 -EndColumn 8

# 4. Tab 4: Suivi des Non conformités (Compliance Table starting at Row 12)
Compare-ExcelSheet -SheetName "Suivi des Non conformités" -HeaderRow 12 -StartColumn 1 -EndColumn 4