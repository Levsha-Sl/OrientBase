Attribute VB_Name = "DataSheets"
' TODO Debag
Public Sub InitApp()
    InitData
    InitMarts

    LoadAllDataToMarts
End Sub

Private Sub InitData()
    ranks.Init  'wsRanksData
    Base.Init   'wsBaseData
    clubs.Init  'wsClubsData
End Sub

Private Sub InitMarts()
    RanksSheet.Init  'wsRanks
    BaseSheet.Init 'wsBase
    ' ClubsSheet.Init  'wsClubs
End Sub

Private Sub LoadAllDataToMarts()
    ranks.LoadToMart
    Base.LoadToMart
    ' TODO Clubs.LoadToMart
End Sub

Public Function GetSheet(name As String) As Worksheet
    Dim ws As Worksheet: Set ws = ModuleSheet.GetSheetByName(name)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets. _
        Add(Before:=ThisWorkbook.Worksheets(1))
        ws.name = name
        ws.Visible = xlSheetVeryHidden
    End If

    Set GetSheet = ws
End Function

' TODO Import And Export
