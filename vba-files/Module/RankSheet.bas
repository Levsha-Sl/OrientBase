Attribute VB_Name = "RanksSheet"
Public wsRanks As Worksheet

Public Sub InitRanksSheet()
    Set wsRanks = RanksSheet.GetRanksSheet
End Sub

Public Function ChangesMart(Target As Range) As Boolean
    On Error Goto ErrorHandler
        Dim result As Boolean: result = True
        Dim rowIdx As Long: rowIdx = Target.row
        Dim colIdx As Long: colIdx = Target.Column

        If (colIdx = 2 Or colIdx = 3) And rowIdx >= 2 Then
            Application.EnableEvents = False

            Dim rankId As Long, rankName As String, periodValue As Long
            Dim hasId As Boolean, hasData As Boolean

            rankId = wsRanks.Cells(rowIdx, 1).Value
            rankName = wsRanks.Cells(rowIdx, 2).Value
            periodValue = IIf(IsEmpty(wsRanks.Cells(rowIdx, 3).Value) Or wsRanks.Cells(rowIdx, 3).Value = "", -1, CLng(wsRanks.Cells(rowIdx, 3).Value))

            hasId = (rankId > 0)
            hasData = (rankName <> "") And (periodValue >= 0)

            If Not hasId Then
                ' ID пуст: INSERT если есть все данные
                If hasData Then
                    wsRanks.Cells(rowIdx, 1).Value = Ranks.CreateRank(rankName, periodValue) 
                End If
            Else
                ' ID не пуст: UPDATE если данные есть, DELETE если нет
                If hasData Then
                    Call Ranks.UpdateRank(rankName, periodValue, rankId)
                Else
                    wsRanks.Cells(rowIdx, 1).Value = IIf(Ranks.DeleteRank(rankId), "", rankId) 
                End If
            End If
        End If
 CleanExit:
        ChangesMart = result
        Application.EnableEvents = True
     Exit Function
 ErrorHandler:
        Select Case Err.Number
         Case vbObjectError + 1101
            MsgBox Err.Description, vbExclamation, "Ошибка изменения данных"
         Case Else
            MsgBox "Произошла ошибка при сохранении изменений: " & Err.Description, vbCritical, "Ошибка"
        End Select
        result = False
        Resume CleanExit
End Function

Private Function GetRanksSheet() As Worksheet
    On Error Goto ErrorHandler
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName("Разряды")

        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets. Add(Before:=ThisWorkbook.Worksheets(4))

            With ws
                .Name = "Разряды"
                .Range("A1:C1").Value = [{"ID","Разряд","Период (.г)"}]

                With .Range("A1:C1")
                    .Font.Bold = True
                    .AutoFilter
                End With

                With .Columns("A:C")
                    .AutoFit
                    .HorizontalAlignment = xlLeft
                End With

                Dim btnSave As Object: Set btnSave = ws.Buttons.Add( _
                Left:=ws.Cells(1, 5).Left, _
                Top:=ws.Cells(1, 5).Top, _
                Width:=ws.Cells(1, 5).Width + ws.Cells(1, 6).Width, _
                Height:=ws.Cells(1, 1).Height)
                With btnSave
                    .OnAction = "BtnSave"
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
        MsgBox "Ошибка при получении листа разрядов: " & Err.Description, vbCritical
        Resume CleanExit
End Function

Private Sub BtnSave()
    wsRanks.Range("A2").Resize(Ranks.getMaxId(), 3).ClearContents
    Call Ranks.SaveChanges
End Sub