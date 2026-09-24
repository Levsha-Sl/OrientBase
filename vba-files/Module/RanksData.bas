Attribute VB_Name = "RanksData"
Public Const SHEET_NAME As String = "Ranks"

Public wsRanksData As Worksheet
' (STAT)key PERIODAGE
Private RanksDict As Object
' (ID)key STAT
Private RanksKeys As Object
Private MaxId As Long
Public Function getMaxId() As Long
    getMaxId = MaxId
End Function

Public Function Init() As Worksheet
    If wsRanksData Is Nothing Then
        ' (STAT)key PERIODAGE
        Set wsRanksData = DataSheets.GetSheet(SHEET_NAME)
        If wsRanksData.Cells(1, 1).value = "" Then
            Call LoadPrimaryRanks
        End If
    End If

    Set Init = wsRanksData
End Function

Public Sub LoadToShowcase()
    Dim martArr() As Variant
    Set RanksKeys = CreateObject("Scripting.Dictionary")
    RanksKeys.CompareMode = 1
    Set RanksDict = CreateObject("Scripting.Dictionary")
    RanksDict.CompareMode = 1

    Dim ranks As Variant: ranks = getRanks()
    MaxId = UBound(ranks, 1)
    ReDim martArr(1 To MaxId, 1 To 3)

    Dim i As Long
    For i = 1 To MaxId
        RanksKeys.Add i, ranks(i, 1)
        RanksDict.Add ranks(i, 1), ranks(i, 2)
        martArr(i, 1) = i            ' ID
        martArr(i, 2) = ranks(i, 1)  ' Rank
        martArr(i, 3) = ranks(i, 2)  ' PeriodAge
    Next i

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    RanksSheet.wsRanks.Range("A2").Resize(MaxId, 3).value = martArr
    RanksSheet.wsRanks.Columns("A:C").AutoFit

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub SaveChanges()
    Dim totalRows As Long: totalRows = RanksDict.Count
    Dim TotalCols As Long: TotalCols = 2
    Dim resultRanks As Variant: ReDim resultRanks(1 To totalRows, 1 To TotalCols)

    Dim i As Long
    For i = 1 To totalRows
        resultRanks(i, 1) = RanksDict.Keys()(i - 1) ' Rank
        resultRanks(i, 2) = RanksDict.Items()(i - 1) ' PeriodAge
    Next i

    If totalRows > 0 Then
        With wsRanksData
            Dim lastRow As Long: lastRow = UBound(getRanks, 1)
            .Range("A1").Resize(lastRow, TotalCols).ClearContents
            .Range("A1").Resize(totalRows, TotalCols).value = resultRanks
        End With
    End If
    Call LoadToShowcase
End Sub

Public Sub UpdateRank(updatedRank As String, periodValue As Long, id As Long)
    If Not RanksDict.Exists(updatedRank) Then
        RanksDict.Item(RanksKeys(id)) = periodValue
        RanksDict.key(RanksKeys(id)) = updatedRank
        RanksKeys(id) = updatedRank
    ElseIf (RanksDict.Item(RanksKeys(id)) <> periodValue) Then
        RanksDict.Item(RanksKeys(id)) = periodValue
    Else
        Err.Raise vbObjectError + 1101, "UpdateRank", "Наименования разрядов не должны совпадать"
    End If
End Sub

Public Function DeleteRank(id As Long) As Boolean
    If RanksKeys.Exists(id) Then
        If RanksDict.Exists(RanksKeys.Item(id)) Then
            RanksDict.Remove (RanksKeys.Item(id))
            RanksKeys.Remove (id)
        Else
            Err.Raise vbObjectError + 1101, "DeleteRank", "Разряд с таким наименованием не найден"
        End If
    Else
        Err.Raise vbObjectError + 1101, "DeleteRank", "Разряд с таким id не найден"
    End If
    DeleteRank = True
End Function

Public Function CreateRank(newRank As String, periodValue As Long) As Long
    If Not RanksDict.Exists(newRank) Then
        MaxId = MaxId + 1
        Call RanksDict.Add(newRank, periodValue)
        Call RanksKeys.Add(MaxId, newRank)
    Else
        Err.Raise vbObjectError + 1101, "CreateRank", "Наименования разрядов не должны совпадать"
    End If
    CreateRank = MaxId
End Function

Public Function GetRankValue(targetRank As String) As Long
    If RanksDict.Exists(targetRank) Then
        GetRankValue = RanksDict.Item(targetRank)
    Else
        Err.Raise vbObjectError + 1102, "GetRankValue", "Разряд с таким наименованием не найден"
    End If
End Function

Private Sub LoadPrimaryRanks()
    Dim primaryRanks As Variant
    Dim rowsCount As Long

    primaryRanks = [{"бр",0;"IIIю",2;"IIю",2;"Iю",2;"III",2;"II",2;"I",2;"КМС",2;"МС",3;"МСМК",0}]
    rowsCount = UBound(primaryRanks, 1)

    wsRanksData.Range("A1:B" & rowsCount).value = primaryRanks
End Sub

Private Function getRanks() As Variant
    Dim lastRow As Long, lastCol As Long
    lastRow = wsRanksData.Cells(wsRanksData.Rows.Count, 1).End(xlUp).row
    lastCol = 2
    getRanks = wsRanksData.Range(wsRanksData.Cells(1, 1), wsRanksData.Cells(lastRow, lastCol)).value
End Function
