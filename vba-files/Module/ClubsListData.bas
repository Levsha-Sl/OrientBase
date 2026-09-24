Attribute VB_Name = "ClubsListData"
Option Private Module

Private ClubsList As Object
Private CombinedClubsList As Object

'/*
' 0 "Инф.базы", 1 "Обнов", 2 "Новый"
'*/
Public Property Get arrStatus() As Variant ' 0,1,2
arrStatus = Array("Инф.базы", "Обнов", "Новый")
End Property

Public Sub Init()
    Dim clubsData As Object
    Dim clubsView As Object
    Dim selectedClubName As Collection

    Dim list As Collection
    Set list = ModuleCSV.OpenCSVFile()
    If list is Nothing Then
     Exit Sub
    End If

    Set clubsData = ParseAthletes(list)
    Set clubsView = CreateObject("Scripting.Dictionary"): clubsView.CompareMode = 1

    For Each key In clubsData.Keys
        clubsView.Add key, clubsData.Item(key).Count
    Next key

    Set selectedClubName = ClubSelector.Show(clubsView)
    If selectedClubName Is Nothing Then
        MsgBox "Отмена", vbInformation
     Exit Sub
    End If

    ClubsListData.InitLists selectedClubName, clubsData
End Sub

'/*
' Dictionary Kay:"clubName" Value: "athleteDict"
' Dictionary Key:"FIO#BD" Value: "stat"  
'*/
Public Sub InitLists(selectedClubName As Collection, clubsData As Object)
    Set ClubsList = CreateObject("Scripting.Dictionary"): ClubsList.CompareMode = 1

    Dim club As ClubModule
    Dim time As Date: time = Now
    Dim i As Long: i = 0
    For Each clubName In selectedClubName
        Set club = New ClubModule
        Call club.Init(clubName, time, clubsData.Item(clubName))
        i = i + 1
        ClubsList.Add i, club
    Next clubName

    ClubsSheet.Init
    Set CombinedClubsList = CreateObject("Scripting.Dictionary"): CombinedClubsList.CompareMode = 1
    ClubsSheet.AcceptClubs ClubsList
End Sub

'/*
' Dictionary Kay:"clubName" Value: "athleteDict"
' Dictionary Key:"FIO#BD" Value: "stat"
'*/
Private Function ParseAthletes(lines As Collection) As Object
    Dim clubs As Object

    Dim athleteDict As Object
    Dim lineData As String
    Dim cols() As String
    Dim clubName As String
    Dim athleteKey As String
    Dim athleteRank As String
    Dim i As Long

    Set clubs = CreateObject("Scripting.Dictionary")
    clubs.CompareMode = 1

    For i = 1 To lines.Count
        lineData = lines(i)
        cols = Split(lineData, ";")

        ' ФИО#ДР Разряд
        If UBound(cols) >= 11 Then
            clubName = Trim(cols(1))
            'TODO big big data
            athleteRank = Trim(cols(11))
            athleteKey = Trim(cols(3)) & " " & Trim(cols(4)) & " " & Trim(cols(5)) & "#" & Format(Trim(cols(7)), "dd.mm.yyyy")
            
            If Not clubs.Exists(clubName) Then
                Set athleteDict = CreateObject("Scripting.Dictionary")
                clubs.Add clubName, athleteDict
            Else
                Set athleteDict = clubs(clubName)
            End If

            athleteDict.Add athleteKey, athleteRank
        End If
    Next i

    Set ParseAthletes = clubs
End Function

'/*
' Dictionary Kay:"clubName" Value: "athleteDict"
' Dictionary Key:"FIO#BD" Value: "stat"
'*/
Public Sub SetClubInCombinedList(idClub As Long)
    Dim club As ClubModule
    Set club = ClubsList.Item(idClub)

    Dim combParticipant As Object
    Set combParticipant = CreateObject("Scripting.Dictionary"): combParticipant.CompareMode = 1

    Dim localStatuses As Variant
    localStatuses = arrStatus

    Dim stat As String
    Dim importedData As Variant
    Dim status As String
    Dim arrData As Variant
    Dim defaultRow As Variant

    For Each pKey In club.Participants.Keys
        importedData = club.Participants.Item(pKey)
        If IsArray(importedData) Then
            ' Импорт базы уже содержит даты разряда и страховки, а также период.
            defaultRow = importedData
        Else
            stat = CStr(importedData)
            defaultRow = Array("",stat,"","","") 'ID STAT DATESTAT DATEINS PERIOD
        End If
        Dim keyParts As Variant: keyParts = Split(pKey, "#")
        arrData = Base.GetParticipant(CStr(keyParts(0)), CDate(keyParts(1)))

        If IsArray(arrData) Then
            status = localStatuses(0)
            If IsArray(importedData) Then
                defaultRow(0) = arrData(0)
            ElseIf stat = arrData(1) Then
                defaultRow = arrData
            Else
                defaultRow(0) = arrData(0)
                defaultRow(3) = arrData(3)
                defaultRow(4) = arrData(4)
            End If
        Else
            status = localStatuses(2)
            arrData = defaultRow
        End If

        combParticipant.Add pKey & "#" & status, arrData

        If status = localStatuses(0) Then
            combParticipant.Add pKey & "#" & localStatuses(1), defaultRow
        End If 
    Next pKey 

    'ФИО#ДР#состояние Id STAT DATESTAT DATEINS PERIOD
    Dim newClub As New ClubModule
    newClub.Init club.ClubName, club.TimeStamp, combParticipant

    CombinedClubsList.Add idClub, newClub
End Sub

Public Function GetClubForShowcase(idClub As Long) As Variant
    Dim club As ClubModule
    Set club = CombinedClubsList.Item(idClub)

    Dim participants As Variant
    participants = club.Participants.Items
    Dim pKeys As Variant: pKeys =  club.Participants.Keys
    Dim countP As Long: countP = club.ParticipantCount
    ReDim martArr(1 To countP, 1 To 10)

    Dim i As Long
    Dim key As String
    Dim keyParts As Variant
    For i = 1 To countP
        key = pKeys(i - 1)
        keyParts = Split(key, "#")
        martArr(i, 1) = keyParts(0)         ' FNM
        martArr(i, 2) = keyParts(2)         ' STATUS
        martArr(i, 3) = participants(i - 1)(0)    ' ID
        martArr(i, 4) = CDate(keyParts(1))         ' BIRTH
        martArr(i, 5) = participants(i - 1)(1)  ' STAT
        martArr(i, 6) = participants(i - 1)(2)  ' DATESTAT
        martArr(i, 7) = Empty               ' Formula For stat
        martArr(i, 8) = participants(i - 1)(3)  ' DATEINS
        martArr(i, 9) = participants(i - 1)(4)    ' PERIOD
        martArr(i, 10) = Empty              ' Formula For ins
    Next i

    GetClubForShowcase = martArr
End Function

'/*
' key: "fnm & "#" & birth "#" & status"
' value: id, stat, dateStat, dateIns, period
'*/
Public Sub UpdateParticipant(idClub As Long, key As String, stat As String, dateStat As Variant, dateIns As Variant, period As Variant)
    If Not CombinedClubsList.Exists(idClub) Then 
        Err.Raise vbObjectError + 1401, "UpdateParticipant", "Клуб не найден"
    End If

    Dim dictParticipants As Object
    Set dictParticipants = CombinedClubsList.Item(idClub).Participants()

    If Not dictParticipants.Exists(key) Then 
        Err.Raise vbObjectError + 1401, "UpdateParticipant", "Спортсмен не найден"
    End If

    dictParticipants.Item(key) = Array(dictParticipants.Item(key)(0), stat, dateStat, dateIns, period)
End Sub

Public Sub ExportToBase(idClub As Long)
    If Not CombinedClubsList.Exists(idClub) Then 
        Err.Raise vbObjectError + 1401, "UpdateParticipant", "Клуб не найден"
    End If

    Dim dictParticipants As Object
    Set dictParticipants = CombinedClubsList.Item(idClub).Participants()

    Dim ExportParticipants As Object
    Set ExportParticipants = CreateObject("Scripting.Dictionary")
    ExportParticipants.CompareMode = 1

    '(FNM BIRTH STATUS)key ID STAT DATESTAT DATEINS PERIOD
    ' To
    '(FNM BIRTH)key ID STAT DATESTAT DATEINS PERIOD
    Dim keyParts As Variant
    For Each participantKey In dictParticipants
        KeyParts = Split(participantKey,"#")
        If KeyParts(2) <> arrStatus(0) Then
            ExportParticipants.Add KeyParts(0) & "#" & KeyParts(1), dictParticipants.Item(ParticipantKey)
        End If
    Next participantKey 

    Base.AcceptOrgeoClub ExportParticipants
    ClubsSheet.DeleteClub idClub
End Sub

Private Sub getCombinedParticipantsClubs()

End Sub

Private Sub getCombinedParticipantsList()

End Sub
' TODO удаление списка кулбов и всех связанных с ним листов
