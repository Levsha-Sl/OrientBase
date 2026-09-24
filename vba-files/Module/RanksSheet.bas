Attribute VB_Name = "RanksSheet"
Public wsRanks As Worksheet

Public Const SHEET_NAME As String = "Разряды"

Public Sub Init()
    Set wsRanks = GetRanksSheet
End Sub

Public Function ChangesShowcase(Target As Range) As Boolean
    On Error Goto ErrorHandler
        Dim Result As Boolean: Result = True
        Dim rowIdx As Long: rowIdx = Target.row
        Dim colIdx As Long: colIdx = Target.Column

        If (colIdx = 2 Or colIdx = 3) And rowIdx >= 2 Then
            Application.EnableEvents = False

            Dim rankId As Long, rankName As String, periodValue As Long
            Dim hasId As Boolean, hasData As Boolean

            rankId = wsRanks.Cells(rowIdx, 1).value
            rankName = wsRanks.Cells(rowIdx, 2).value
            periodValue = IIf(IsEmpty(wsRanks.Cells(rowIdx, 3).value) Or wsRanks.Cells(rowIdx, 3).value = "", -1, CLng(wsRanks.Cells(rowIdx, 3).value))

            hasId = (rankId > 0)
            hasData = (rankName <> "") And (periodValue >= 0)

            If Not hasId Then
                ' ID пуст: INSERT если есть все данные
                If hasData Then
                    wsRanks.Cells(rowIdx, 1).value = RanksData.CreateRank(rankName, periodValue)
                End If
            Else
                ' ID не пуст: UPDATE если данные есть, DELETE если нет
                If hasData Then
                    Call RanksData.UpdateRank(rankName, periodValue, rankId)
                Else
                    wsRanks.Cells(rowIdx, 1).value = IIf(RanksData.DeleteRank(rankId), "", rankId)
                End If
            End If
        End If
 CleanExit:
        ChangesShowcase = Result
        Application.EnableEvents = True
     Exit Function
 ErrorHandler:
        Select Case Err.Number
         Case vbObjectError + 1101
            MsgBox Err.Description, vbExclamation, "Ошибка изменения данных"
         Case Else
            MsgBox "Произошла ошибка при сохранении изменений: " & Err.Description, vbCritical, "Ошибка"
        End Select
        Result = False
        Resume CleanExit
End Function

Private Function GetRanksSheet() As Worksheet
    On Error Goto ErrorHandler
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName(SHEET_NAME)

        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))

            With ws
                .name = SHEET_NAME
                .Range("A1:C1").value = [{"ID","Разряд","Период (.г)"}]

                With .Range("A1:C1")
                    .Font.Bold = True
                    .AutoFilter
                End With

                With .Columns("A:C")
                    .AutoFit
                    .HorizontalAlignment = xlLeft
                End With

                Dim BtnSave As Object: Set BtnSave = ws.Buttons.Add( _
                Left:=ws.Cells(1, 5).Left, _
                Top:=ws.Cells(1, 5).Top, _
                Width:=ws.Cells(1, 5).Width + ws.Cells(1, 6).Width, _
                Height:=ws.Cells(1, 1).Height)
                With BtnSave
                    .OnAction = "RanksSheet.BtnSave"
                    .Caption = "Сохранить"
                End With
            End With
        End If

        Set GetRanksSheet = ws

 CleanExit:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
     Exit Function
 ErrorHandler:
        MsgBox "Ошибка при создании листа разрядов: " & Err.Description, vbCritical
        Resume CleanExit
End Function

Private Sub BtnSave()
    wsRanks.Range("A2").Resize(RanksData.getMaxId(), 3).ClearContents
    Call RanksData.SaveChanges
End Sub
