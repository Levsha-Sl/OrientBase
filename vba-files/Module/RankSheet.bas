Attribute VB_Name = "RankSheet"
Public wsRanks As Worksheet

Public Sub InitRanksSheet()
    If wsRanks Is Nothing Then
        Set wsRanks = RankSheet.GetRankSheet
    End If
End Sub

Public Sub ChangesMart()
    
End Sub

Private Function GetRankSheet() As Worksheet
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
                    .OnAction = "Ranks.SaveChanges"
                    .Caption = "Сохранить"
                End With
            End With
        End If

        Set GetRankSheet = ws

 CleanExit:
        Application.EnableEvents = True
        Application.ScreenUpdating = True
     Exit Function
 ErrorHandler:
        MsgBox "Ошибка при получении листа разрядов: " & Err.Description, vbCritical
        Resume CleanExit
End Function