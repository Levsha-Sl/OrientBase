Attribute VB_Name = "Data"
Public wsData As Worksheet
Public DataKeys As Object

Public Function Init() As Worksheet
    If wsData Is Nothing Then
        ' (FNM BIRTH)key STAT DATESTAT DATEINS PERIOD
        Set wsData = DataSheets.GetSheet("Data", 1)
        Set DataKeys = CreateObject("Scripting.Dictionary")
    End If

    Set Init = wsData
End Function

' TODO Import
Public Sub Import()

End Sub

Public Sub Export()
    Dim filePath As Variant
    Dim fNum As Integer

    Dim dataArr As Variant
    Dim lastRow As Long, lastCol As Long: lastCol = 6 ' Фиксированное количество столбцов (FNM BIRTH STAT DATESTAT DATEINS PERIOD)

    ' Диалог
    filePath = Application.GetSaveAsFilename( _
    InitialFileName:="data.csv", _
    FileFilter:="CSV Files (*.csv), *.csv" _
    )

    If filePath = False Then Exit Sub

        lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
        dataArr = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol)).Value

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
                line = line & dataArr(i, j)
                If j < lastCol Then line = line & ";"
                Next j
                Print #fNum, line
            Next i

            Close #fNum

            MsgBox "База экспортирована в " & filePath, vbInformation
End Sub