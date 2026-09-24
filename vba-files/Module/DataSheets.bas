Attribute VB_Name = "DataSheets"
' TODO Debag
Public Sub InitApp()
    InitData
    InitShowcases

    LoadAllDataToShowcases
End Sub

Private Sub InitData()
    RanksData.Init  'wsRanksData
    Base.Init   'wsBaseData
    ' Clubs.Init  'wsClubsData
End Sub

Private Sub InitShowcases()
    RanksSheet.Init  'wsRanks
    BaseSheet.Init 'wsBase
    ' ClubsSheet.Init  'wsClubs
End Sub

Private Sub LoadAllDataToShowcases()
    RanksData.LoadToShowcase
    Base.LoadToShowcase
    ' Clubs.LoadToShowcase
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
