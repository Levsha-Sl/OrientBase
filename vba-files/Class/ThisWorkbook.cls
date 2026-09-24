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

Private Sub Workbook_SheetSelectionChange(Byval sh As Object, Byval Target As Range)
    Select Case sh.name
     Case ClubsSheet.SHEET_NAME
        Call ClubsSheet.HandleSelectionChange(Target)
    End Select
End Sub

Private Sub Workbook_SheetFollowHyperlink(Byval sh As Object, Byval Target As Hyperlink)
    Select Case sh.name
     Case ClubsSheet.SHEET_NAME
        ClubsSheet.HandleFollowHyperlink Target
    End Select
End Sub

Private Sub Workbook_SheetChange(Byval sh As Object, Byval Target As Range)
    Dim hasChangesShowcase As Boolean: hasChangesShowcase = True

    Select Case sh.name
     Case BaseSheet.SHEET_NAME
        hasChangesShowcase = BaseSheet.ChangesShowcase(Target)
     Case RanksSheet.SHEET_NAME
        hasChangesShowcase = RanksSheet.ChangesShowcase(Target)
    End Select

    If sh.name Like "L*" Then
        hasChangesShowcase = ClubSheet.ChangesShowcase(sh,Target)
    End If

    If Not hasChangesShowcase Then
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets

        If sh.Visible = xlSheetVisible _
            And sh.name <> BaseSheet.SHEET_NAME _
            And sh.name <> RanksSheet.SHEET_NAME Then
            sh.Delete
        End If
    Next sh

    Application.DisplayAlerts = True
End Sub

