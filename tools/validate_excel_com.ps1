[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]] $Path,

    [switch] $RefreshPivots
)

$ErrorActionPreference = "Stop"
$failures = [System.Collections.Generic.List[string]]::new()
$excel = $null

function Release-ComObject {
    param([object] $Value)
    if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

# Excel answers RPC_E_CALL_REJECTED (0x80010001) while it is busy; a rejected call
# is not a verdict on the file, so retry it.
function Invoke-Com {
    param([scriptblock] $Call)
    for ($attempt = 1; ; $attempt++) {
        try {
            return & $Call
        }
        catch {
            $rejected = $_.Exception -is [System.Runtime.InteropServices.COMException] -and
                $_.Exception.HResult -eq -2147418111
            if (-not $rejected -or $attempt -ge 20) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbooks = $excel.Workbooks

    foreach ($candidate in $Path) {
        $workbook = $null
        $worksheets = $null
        try {
            $fullPath = [System.IO.Path]::GetFullPath($candidate)
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
                throw "File does not exist: $fullPath"
            }

            # Open read-only with Excel's normal loader.  A corrupt-load mode
            # is deliberately not used: silent repair would make validation
            # report a damaged file as valid.
            $workbook = Invoke-Com { $workbooks.Open($fullPath, 0, $true) }
            # (Open returns one COM object, so Invoke-Com does not enumerate it.)
            # A busy Excel can hand back a workbook whose sheets are not loaded yet
            # (Count 0 without an error); give it a moment before calling it empty.
            $sheetCount = 0
            for ($attempt = 1; $attempt -le 20; $attempt++) {
                $worksheets = $workbook.Worksheets
                $sheetCount = [int](Invoke-Com { $worksheets.Count })
                if ($sheetCount -ge 1) { break }
                Release-ComObject $worksheets
                Start-Sleep -Milliseconds 500
            }
            if ($sheetCount -lt 1) {
                throw "Workbook has no worksheets: $fullPath"
            }

            for ($sheetIndex = 1; $sheetIndex -le $sheetCount; $sheetIndex++) {
                $worksheet = $null
                $usedRange = $null
                try {
                    $worksheet = Invoke-Com { $worksheets.Item($sheetIndex) }
                    $usedRange = Invoke-Com { $worksheet.UsedRange }
                    # Force Excel to materialise the complete used region.
                    $null = Invoke-Com { $usedRange.Value2 }
                    if ($RefreshPivots) {
                        $pivotTables = $worksheet.PivotTables()
                        $pivotCount = [int](Invoke-Com { $pivotTables.Count })
                        for ($pivotIndex = 1; $pivotIndex -le $pivotCount; $pivotIndex++) {
                            $pivotTable = $null
                            try {
                                $pivotTable = Invoke-Com { $pivotTables.Item($pivotIndex) }
                                $null = Invoke-Com { $pivotTable.RefreshTable() }
                            }
                            finally {
                                Release-ComObject $pivotTable
                            }
                        }
                        Release-ComObject $pivotTables
                    }
                }
                finally {
                    Release-ComObject $usedRange
                    Release-ComObject $worksheet
                }
            }

            Write-Output ("OK`t{0}`tworksheets={1}" -f $fullPath, $sheetCount)
        }
        catch {
            $failures.Add(("FAIL`t{0}`t{1}" -f $candidate, $_.Exception.Message))
        }
        finally {
            if ($null -ne $workbook) {
                try { $workbook.Close($false) } catch { }
                Release-ComObject $workbook
            }
            Release-ComObject $worksheets
            $worksheets = $null
        }
    }
}
catch {
    $failures.Add(("Excel COM unavailable: {0}" -f $_.Exception.Message))
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        Release-ComObject $workbooks
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

foreach ($failure in $failures) {
    Write-Error $failure
}
if ($failures.Count -gt 0) {
    exit 1
}
exit 0
