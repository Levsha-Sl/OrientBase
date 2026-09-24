param()

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('OrientBase-import-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testRoot
$ansi = [System.Text.Encoding]::GetEncoding(1251)
$testExcel = $null
$testBook = $null

function Add-TestModule([string] $name, [string] $code) {
    $component = $testBook.VBProject.VBComponents.Add(1)
    $component.Name = $name
    $component.CodeModule.AddFromString($code)
}

try {
    # A separate Excel instance and a new unsaved workbook; the project workbook is never opened.
    $testExcel = New-Object -ComObject Excel.Application
    $testExcel.Visible = $false
    $testExcel.DisplayAlerts = $false
    $testExcel.EnableEvents = $false
    $testBook = $testExcel.Workbooks.Add()

    foreach ($name in @('Base', 'ClubsListData', 'ModuleCSV')) {
        $null = $testBook.VBProject.VBComponents.Import((Join-Path $projectRoot "vba-files\Module\$name.bas"))
    }
    $null = $testBook.VBProject.VBComponents.Import((Join-Path $projectRoot 'vba-files\Class\ClubModule.cls'))

    # Expose private parsing boundaries only inside the disposable test workbook.
    $testBook.VBProject.VBComponents.Item('Base').CodeModule.AddFromString(@'
Public Function TestParse(lines As Collection) As Object
    Set TestParse = ParseAthletes(lines)
End Function
'@)
    $testBook.VBProject.VBComponents.Item('ModuleCSV').CodeModule.AddFromString(@'
Public Function TestRead(path As String, charset As String, header As String) As Collection
    Set TestRead = ReadCSVFile(path, charset, header)
End Function
'@)

    # Replace presentation and dialogs with no-op boundaries; storage and reconciliation are real.
    Add-TestModule 'BaseSheet' @'
Public Sub AcceptBaseData(rows As Variant)
End Sub
'@
    Add-TestModule 'ClubsSheet' @'
Public Sub Init()
End Sub
Public Sub AcceptClubs(clubs As Object)
End Sub
Public Sub DeleteClub(id As Long)
End Sub
'@
    Add-TestModule 'ClubSelector' @'
Public Function Show(clubs As Object) As Collection
    Set Show = Nothing
End Function
'@
    Add-TestModule 'DataSheets' @'
Public Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Set GetSheet = ThisWorkbook.Worksheets.Add
        GetSheet.Name = sheetName
    End If
End Function
'@
    $testCode = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ImportTests.bas'), $ansi)
    Add-TestModule 'ImportTests' $testCode

    $header = 'FNM;BIRTH;STAT;DATESTAT;DATEINS;PERIOD;TIMESTAMP'
    $fixture = $header + "`r`n" + "Иванов Иван;01.02.2000;КМС;03.04.2025;05.06.2026;365;01.01.2026`r`n"
    [System.IO.File]::WriteAllText((Join-Path $testRoot 'ansi.csv'), $fixture, $ansi)
    [System.IO.File]::WriteAllText((Join-Path $testRoot 'utf8.csv'), $fixture, [System.Text.UTF8Encoding]::new($true))
    [System.IO.File]::WriteAllText((Join-Path $testRoot 'orgeo.csv'), "header`r`nИванов Иван;КМС`r`n", [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllText((Join-Path $testRoot 'wrong-header.csv'), "FNM;BIRTH`r`nTest;01.02.2000`r`n", $ansi)
    [System.IO.File]::WriteAllText((Join-Path $testRoot 'empty.csv'), '', $ansi)
    [System.IO.File]::WriteAllText((Join-Path $testRoot 'header-only.csv'), $header + "`r`n", $ansi)

    $testExcel.Run("'$($testBook.Name)'!ImportTests.RunAll", $testRoot)
}
finally {
    if ($null -ne $testBook) {
        $testBook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($testBook)
    }
    if ($null -ne $testExcel) {
        $testExcel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($testExcel)
    }
    $resolvedTestRoot = [System.IO.Path]::GetFullPath($testRoot)
    $resolvedTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolvedTestRoot.StartsWith($resolvedTemp, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path $resolvedTestRoot -Leaf) -like 'OrientBase-import-*') {
        Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
    }
}
