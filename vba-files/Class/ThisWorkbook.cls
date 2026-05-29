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
    Dim hasChangesMart As Boolean: hasChangesMart = True

    Select Case Sh.name
     Case "База"
        hasChangesMart = BaseSheet.ChangesMart(Target) 
     Case "Разряды"
        hasChangesMart = RanksSheet.ChangesMart(Target)
    End Select

    If Not hasChangesMart Then
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
    End If
End Sub
