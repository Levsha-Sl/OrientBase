Attribute VB_Name = "Clubs"
Public wsClubsData As Worksheet
Public ClubsKeys As Object

Public Function Init() As Worksheet
    If wsClubsData Is Nothing Then
        ' LINKSHEET LINKINFOSHEET (NAME TIMELOAD)key
        Set wsClubsData = DataSheets.GetSheet("Clubs")
        Set ClubsKeys = CreateObject("Scripting.Dictionary")
    End If

    Set Init = wsClubsData
End Function