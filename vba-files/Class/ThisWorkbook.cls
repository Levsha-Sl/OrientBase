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
    Select Case Sh.name
     Case ClubsSheet.SHEET_NAME
        Call ClubsSheet.HandleSelectionChange(Target)
    End Select          
End Sub

Private Sub Workbook_SheetFollowHyperlink(Byval Sh As Object, Byval Target As Hyperlink)
    Select Case Sh.name
     Case ClubsSheet.SHEET_NAME
        ClubsSheet.HandleFollowHyperlink Target
    End Select 
End Sub

Private Sub Workbook_SheetChange(Byval Sh As Object, Byval Target As Range)
    Dim hasChangesMart As Boolean: hasChangesMart = True

    Select Case Sh.name
     Case BaseSheet.SHEET_NAME
        hasChangesMart = BaseSheet.ChangesMart(Target)
     Case RanksSheet.SHEET_NAME
        hasChangesMart = RanksSheet.ChangesMart(Target)
    End Select

    If Not hasChangesMart Then
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
        And sh.Name <> BaseSheet.SHEET_NAME _
        And sh.Name <> RanksSheet.SHEET_NAME Then
            sh.Delete
        End If
    Next sh

    Application.DisplayAlerts = True
End Sub
