Attribute VB_Name = "Base"
Public wsBase As Worksheet
Public BaseKeys As Object

Public Function Init() As Worksheet
    If wsBase Is Nothing Then
        ' (FNM BIRTH)key STAT DATESTAT DATEINS PERIOD
        Set wsBase = DataSheets.GetSheet("Base", 1)
        Set BaseKeys = CreateObject("Scripting.Dictionary")
    End If

    Set Init = wsBase
End Function

' TODO Import
Public Sub Import()

End Sub

Public Sub Export()
    Dim filePath As Variant
    Dim fNum As Integer

    Dim baseArr As Variant
    Dim lastRow As Long, lastCol As Long: lastCol = 6 ' Фиксированное количество столбцов (FNM BIRTH STAT DATESTAT DATEINS PERIOD)

    ' Диалог
    filePath = Application.GetSaveAsFilename( _
    InitialFileName:="base.csv", _
    FileFilter:="CSV Files (*.csv), *.csv" _
    )

    If filePath = False Then Exit Sub

        lastRow = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row
        baseArr = wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(lastRow, lastCol)).Value

        fNum = FreeFile
        Open filePath For Output As #fNum

        ' --- Заголовки ---
        Print #fNum, "FNM;BIRTH;STAT;DATESTAT;DATEINS;PERIOD"
        ' --- Данные ---
        Dim i As Long, j As Long
        Dim line As String
        For i = 1 To lastRow
            line = ""
            For j = 1 To lastCol
                line = line & baseArr(i, j)
                If j < lastCol Then line = line & ";"
                Next j
                Print #fNum, line
            Next i

            Close #fNum

            MsgBox "База экспортирована в " & filePath, vbInformation
End Sub