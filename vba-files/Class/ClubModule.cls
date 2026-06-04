VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClubModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private pClubName As String
Private pTimeStamp As Date
Private pParticipants As Object

Public Sub Init(ByVal clubName As String, ByVal dataTime As Date, ByVal athleteData As Object)
    pClubName = clubName
    pTimeStamp = dataTime
    Set pParticipants = athleteData
End Sub

Public Property Get clubName() As String
clubName = pClubName
End Property

Public Property Get timeStamp() As Date
timeStamp = pTimeStamp
End Property

Public Property Get ParticipantCount() As Long
If pParticipants Is Nothing Then
    ParticipantCount = 0
Else
    ParticipantCount = pParticipants.Count
End If
End Property

Public Property Get Participants() As Object
Set Participants = pParticipants
End Property
