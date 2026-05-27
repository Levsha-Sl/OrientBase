VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    DataSheets.InitApp  
End Sub

Private Sub Workbook_SheetSelectionChange(Byval Sh As Object, Byval Target As Range)
    If Sh.name = "Клубы" Then
        ' TODO переместить в клубы
        If Target.Column <> 5 Or Target.Count > 1 Then Exit Sub
            If Target.Value <> "УДАЛИТЬ КЛУБ" Or Target.row = 1 Then Exit Sub

                Dim mainRef As String, needRef As String

                mainRef = ModuleSheet.GetSheetFromLink(Sh.Cells(Target.row, 1))
                needRef = ModuleSheet.GetSheetFromLink(Sh.Cells(Target.row, 2))

                If MsgBox("Удалить листы этой загрузки?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

                    Application.DisplayAlerts = False

                    ModuleSheet.DeleteSheetIfExists mainRef
                    ModuleSheet.DeleteSheetIfExists needRef

                    Target.EntireRow.Delete

                    Application.DisplayAlerts = True
                End If
End Sub

Private Sub Workbook_SheetFollowHyperlink(Byval Sh As Object, Byval Target As Hyperlink)
    If Sh.name = "Клубы" Then
        ' TODO переместить в клубы
        Dim sheetName As String
        sheetName = Split(Target.SubAddress, "!")(0)
        sheetName = Replace(sheetName, "'", "")

        If Not ModuleSheet.SheetExists(sheetName) Then
            MsgBox "Лист [" & sheetName & "] не найден. Ссылка удалена.", vbExclamation
            Target.Range.ClearContents
        End If
End Sub

Private Sub Workbook_SheetChange(Byval Sh As Object, Byval Target As Range)
    If Sh.name = "База" Then
        ' TODO переместить в базу
        Dim rowIdx As Long: rowIdx = Target.row
        Dim colIdx As Long: colIdx = Target.Column
        ' Проверяем, что изменение в пределах A-H и ниже заголовков
        If colIdx <= 8 And rowIdx >= 6 Then

            Dim hasFirstThree As Boolean
            Dim hasAnyOther As Boolean

            ' 1. Проверяем первые три (A, B, C) — должны быть все заполнены
            ' CountA вернет 3, если ячейки A, B и C не пустые
            hasFirstThree = (Application.WorksheetFunction.CountA(Sh.Range(Sh.Cells(rowIdx, 1), Sh.Cells(rowIdx, 3))) = 3)

            ' 2. Проверяем остальные (D, E, F, G, H) — должна быть хотя бы одна
            hasAnyOther = (Application.WorksheetFunction.CountA(Sh.Range(Sh.Cells(rowIdx, 4), Sh.Cells(rowIdx, 8))) > 0)

            Application.EnableEvents = False

            ' Если оба условия выполнены — ставим дату
            If hasFirstThree And hasAnyOther Then
                Sh.Cells(rowIdx, 9).Value = Now
            Else
                ' Если условия перестали соблюдаться (например, стерли ФИО) — очищаем дату
                Sh.Cells(rowIdx, 9).ClearContents
            End If

            Application.EnableEvents = True
        End If
    End If

    If Sh.name = "Разряды" Then
        Dim hasChangesMart As Boolean: hasChangesMart = RanksSheet.ChangesMart(Target) 
        If Not hasChangesMart Then
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
        End If
    End If
End Sub
