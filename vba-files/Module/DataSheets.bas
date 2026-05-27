Attribute VB_Name = "DataSheets"
Public Sub AllInit()
    Base.Init 'wsBase
    Ranks.Init 'wsRanksData
    Clubs.Init 'wsClubsData
End Sub

Public Sub LoadAllDataToMarts()
    ' TODO Call Base.LoadToMarts
    Call Ranks.LoadToMart
    ' TODO Clubs.LoadToMart
End Sub

Public Function GetSheet(name As String, queueNumber) As Worksheet
    Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName(name)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets. _
        Add(Before:=ThisWorkbook.Worksheets(queueNumber))
        ws.Name = name
        ws.Visible = xlSheetVeryHidden
    End If

    Set GetSheet = ws
End Function

' TODO Import And Export