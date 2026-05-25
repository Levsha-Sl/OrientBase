Attribute VB_Name = "Ranks"
Public wsRanksData As Worksheet
' (STAT)key ID
Private RanksKeys As Object
' STAT PERIODAGE
Private RanksData As Variant
Private NewRanksData As Collection

Public Function Init() As Worksheet
    If wsRanksData Is Nothing Then
        ' (STAT)key PERIODAGE
        Set wsRanksData = DataSheets.GetSheet("Ranks", 2)
        If wsRanksData.Cells(1, 1).Value = "" Then 
            Call LoadPrimaryRanks
        End If

        RanksData = getRanksData()
        Set RanksKeys = CreateObject("Scripting.Dictionary")
        RanksKeys.CompareMode = 1
    End If

    Set Init = wsRanksData
End Function

Public Sub LoadToMart()
    Dim martArr() As Variant
    Dim rowCount As Long
    rowCount = UBound(RanksData, 1)
    ReDim martArr(1 To rowCount, 1 To 3)
    Dim i As Long

    For i = 1 To rowCount
        RanksKeys.Add RanksData(i, 1), i
        martArr(i, 1) = i                 ' ID
        martArr(i, 2) = RanksData(i, 1)   ' Rank
        martArr(i, 3) = RanksData(i, 2)   ' PeriodAge
    Next i

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    RankSheet.wsRanks.Range("A2").Resize(rowCount, 3).ClearContents
    RankSheet.wsRanks.Range("A2").Resize(rowCount, 3).Value = martArr
    RankSheet.wsRanks.Columns("A:C").AutoFit

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub SaveChanges()
    On Error Goto ErrorHandler 
        Dim totalRows As Long: totalRows = UBound(RanksData, 1)
        Dim totalCols As Long: totalCols = 2
        Dim resultData As Variant: ReDim resultData(1 To totalRows, 1 To totalCols)

        Dim writeRow As Long: writeRow = 0
        Dim i As Long, j As Long
        For i = 1 To totalRows
            If Not IsEmpty(RanksData(i, 2)) And RanksData(i, 2) <> "" Then
                writeRow = writeRow + 1
                For j = 1 To totalCols
                    resultData(writeRow, j) = RanksData(i, j)
                Next j
            End If
        Next i

        If writeRow > 0 Then
            With wsRanksData
                .Range("A1").Resize(totalRows, totalCols).ClearContents
                .Range("A1").Resize(writeRow, totalCols).Value = resultData
            End With
        End If
        Call Refresh

     Exit Sub
 ErrorHandler:
        MsgBox "Ошибка при сохранении изменений: " & Err.Description, vbCritical
        Resume CleanExit
End Sub

Public Sub UpdateRank(targetRank As String, id As Long, )
    If RanksKeys.Exists(targetRank) Then 
        RanksData(RanksKeys.Item(targetRank), 1) = name
        RanksData(RanksKeys.Item(targetRank), 2) = periodValue
    Else
        create
    End If
    Dim row As Long: row = RanksKeys.Item(targetRank)
End Sub

Public Sub DeleteRank(targetRank As String)
    If RanksKeys.Exists(targetRank) Then 
        RanksData(RanksKeys.Item(targetRank), 1) = ""
        RanksData(RanksKeys.Item(targetRank), 2) = ""
    Else
    End If
End Sub

Private Sub CreateRank(targetRank As String, periodValue As Long)
    If RanksKeys.Exists(targetRank) Then 
        RanksKeys.Add targetRank, (UBound(RanksData, 1) + 1)
        NewRanksData.Add Array(targetRank, periodValue)
    Else
        Err.Raise vbObjectError + 1001, "CreateRank", "Наименования разрядов не должны совпадать"
    End If
End Sub

Public Function GetRankValue(targetRank As String) As Long
    If RanksKeys.Exists(targetRank) Then
        GetRankValue = RanksData(RanksKeys.Item(targetRank), 2)
    Else
        GetRankValue = -1
    End If
End Function

Private Sub LoadPrimaryRanks()
    Dim primaryRanks As Variant
    Dim rowsCount As Long

    primaryRanks = [{"бр",0;"IIIю",2;"IIю",2;"Iю",2;"III",2;"II",2;"I",2;"КМС",2;"МС",3;"МСМК",0}]
    rowsCount = UBound(primaryRanks, 1)

    wsRanksData.Range("A1:B" &  rowsCount).Value =primaryRanks
End Sub

Private  Function getRanksData() As Variant
    Dim lastRow As Long, lastCol As Long
    lastRow = wsRanksData.Cells(wsRanksData.Rows.Count, 1).End(xlUp).Row
    getRanksData = wsRanksData.Range(wsRanksData.Cells(1, 1), wsRanksData.Cells(lastRow, 2)).Value
End Function

Private Sub Refresh()
    RanksData = getRanksData()
    Set RanksKeys = CreateObject("Scripting.Dictionary")
    RanksKeys.CompareMode = 1
    Call LoadToMart
End Sub