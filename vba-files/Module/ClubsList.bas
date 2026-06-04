Attribute VB_Name = "ClubsList"

Private ClubsList As Object

Public Sub Init(сlubsList As Object)
    Set ClubsList = сlubsList
    ClubsSheet.Init
    ClubsSheet.AcceptClubs ClubsList
End Sub

' TODO удаление списка кулбов и всех связанных с ним листов