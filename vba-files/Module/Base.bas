Attribute VB_Name = "Base"
Public wsBaseData As Worksheet
' (FNM#BIRTH)key STAT DATESTAT DATEINS PERIOD
Private BaseDict As Object
' (ID)key FNM#BIRTH
Private BaseKeys As Object
' (FNM#BIRTH)key ID
Private UnBaseKeys As Object


Private MaxId As Long
Private TotalCols As Long

Public Function getMaxId() As Long
    getMaxId = MaxId
End Function

Public Function Init() As Worksheet
    If wsBaseData Is Nothing Then
        ' (FNM BIRTH)key STAT DATESTAT DATEINS PERIOD
        Set wsBaseData = DataSheets.GetSheet("Base")
        If wsBaseData.Cells(1, 1).value = "" Then
            Call LoadPrimaryBase
        End If
    End If
    TotalCols = 7 ' Фиксированное количество столбцов (FNM BIRTH STAT DATESTAT DATEINS PERIOD TIMESTAMP)
    Set Init = wsBaseData
End Function

Public Sub LoadToShowcase()
    Dim martArr() As Variant
    Set BaseKeys = CreateObject("Scripting.Dictionary")
    BaseKeys.CompareMode = 1
    Set BaseDict = CreateObject("Scripting.Dictionary")
    BaseDict.CompareMode = 1


    Dim baseData As Variant: baseData = GetBaseData()
    MaxId = UBound(baseData, 1)
    ReDim martArr(1 To MaxId, 1 To 10)

    Dim i As Long
    For i = 1 To MaxId
        Dim key As String: key = baseData(i, 1) & "#" & baseData(i, 2)
        BaseKeys.Add i, key
        BaseDict.Add key, Array(baseData(i, 3), baseData(i, 4), baseData(i, 5), baseData(i, 6), baseData(i, 7))

        martArr(i, 1) = i                   ' ID
        martArr(i, 2) = baseData(i, 1)      ' FNM
        martArr(i, 3) = baseData(i, 2)      ' BIRTH
        martArr(i, 4) = baseData(i, 3)      ' STAT
        martArr(i, 5) = baseData(i, 4)      ' DATESTAT
        martArr(i, 6) = Empty               ' Formula For stat
        martArr(i, 7) = baseData(i, 5)      ' DATEINS
        martArr(i, 8) = baseData(i, 6)      ' PERIOD
        martArr(i, 9) = Empty               ' Formula For ins
        martArr(i, 10) = baseData(i, 7)     ' TIMESTAMP
    Next i

    Set UnBaseKeys = Nothing
    Call BaseSheet.AcceptBaseData(martArr)
End Sub

Public Sub SaveChanges()
    Dim totalRows As Long: totalRows = BaseDict.Count
    Dim resultBase As Variant: ReDim resultBase(1 To totalRows, 1 To TotalCols)

    Dim i As Long, keyParts() As String
    Dim key As String

    For i = 1 To totalRows
        key = BaseDict.Keys()(i - 1)
        keyParts = Split(key, "#")
        resultBase(i, 1) = keyParts(0)                      ' FNM
        resultBase(i, 2) = CDate(keyParts(1))               ' BIRTH
        resultBase(i, 3) = BaseDict.Items()(i - 1)(0)       ' STAT
        resultBase(i, 4) = BaseDict.Items()(i - 1)(1)       ' DATESTAT
        resultBase(i, 5) = BaseDict.Items()(i - 1)(2)       ' DATEINS
        resultBase(i, 6) = BaseDict.Items()(i - 1)(3)       ' PERIOD
        resultBase(i, 7) = BaseDict.Items()(i - 1)(4)       ' TIMESTAMP
    Next i

    If totalRows > 0 Then
        With wsBaseData
            Dim lastRow As Long: lastRow = UBound(GetBaseData, 1)
            .Range("A1").Resize(lastRow, TotalCols).ClearContents
            .Range("A1").Resize(totalRows, TotalCols).value = resultBase
        End With
    End If
    Call LoadToShowcase
End Sub

Public Sub UpdateBase(fnm As String, birth As Date, stat As String, dateStat As Variant, dateIns As Variant, period As Variant, timeStamp As Variant, id As Long)
    Dim oldKey As String, newKey As String

    If BaseKeys.Exists(id) Then
        oldKey = BaseKeys(id)
        newKey = fnm & "#" & birth

        If oldKey <> newKey Then
            ' Ключ изменился - удаляем старый, добавляем новый
            BaseKeys(id) = newKey
            Set UnBaseKeys = Nothing
            BaseDict.key(oldKey) = newKey
        End If

        BaseDict.Item(newKey) = Array(stat, dateStat, dateIns, period, timeStamp)
    Else
        Err.Raise vbObjectError + 1201, "UpdateBase", "Спортсмен с таким ID не найден"
    End If
End Sub

Public Function DeleteBase(id As Long) As Boolean
    If BaseKeys.Exists(id) Then
        Dim key As String: key = BaseKeys(id)
        If BaseDict.Exists(key) Then
            BaseDict.Remove key
            BaseKeys.Remove id
            Set UnBaseKeys = Nothing
        Else
            Err.Raise vbObjectError + 1201, "DeleteBase", "Спортсмен с таким наименованием не найден"
        End If
    Else
        Err.Raise vbObjectError + 1201, "DeleteBase", "Спортсмен с таким ID не найден"
    End If
    DeleteBase = True
End Function

Public Function CreateBase(fnm As String, birth As Date, stat As String, dateStat As Variant, dateIns As Variant, period As Variant, timeStamp As Variant) As Long
    Dim key As String: key = fnm & "#" & birth

    If Not BaseDict.Exists(key) Then
        MaxId = MaxId + 1
        Call BaseDict.Add(key, Array(stat, dateStat, dateIns, period, timeStamp))
        Call BaseKeys.Add(MaxId, key)
        Set UnBaseKeys = Nothing
    Else
        Err.Raise vbObjectError + 1201, "CreateBase", "Спортсмен с такими ФИО и датой рождения уже существует"
    End If
    CreateBase = MaxId
End Function

'/*
' return ID STAT DATESTAT DATEINS PERIOD
'*/
Public Function GetParticipant(fnm As String, birth As Date) As Variant
    Dim resArr As Variant
    ReDim resArr(0 To 4)
    Dim key As String: key = fnm & "#" & Format(birth, "dd.mm.yyyy")
    If BaseDict.Exists(key) Then
        Dim i As Long: i = 0
        Dim dataP As Variant: dataP = BaseDict.Item(key) '0 STAT 1 DATESTAT 2 DATEINS 3 PERIOD 4 TIMESTAMP
        If UnBaseKeys Is Nothing Then
            SetUnBaseKeys
        End If

        resArr(i) = UnBaseKeys.Item(key) 'ID
        For i = 1 To 4
            resArr(i) = dataP(i - 1)
        Next i

        GetParticipant = resArr
    Else
        GetParticipant = Empty
    End If
End Function

'/*
' participant: (FNM BIRTH)key ID STAT DATESTAT DATEINS PERIOD
'*/
Public Sub AcceptOrgeoClub(participant As Object)
    Dim arrExp As Variant
    Dim id As Long
    Dim arrData(0 To 4) As Variant
    Dim timeStamp As Variant
    timeStamp = Format$(Now, "dd.mm.yyyy")
    For Each participantKay In participant
        arrExp = participant.Item(participantKay)
        id = Val(arrExp(0))
        arrData(0) = arrExp(1) ' STAT
        arrData(1) = arrExp(2) ' DATESTAT
        arrData(2) = arrExp(3) ' DATEINS
        arrData(3) = arrExp(4) ' PERIOD
        arrData(4) = timeStamp
        If BaseKeys.Exists(id) Then 
            BaseDict.Item(BaseKeys.Item(id)) = arrData
        Else
            If Not BaseDict.Exists(key) Then 
                MaxId = MaxId + 1
                BaseDict.Add participantKay, arrData
                BaseKeys.Add MaxId, participantKay
            Else

            End If
        End If
    Next participantKay 

    SaveChanges
End Sub

' TODO ниже код не трогать
Public Sub Import()
    Dim list As Collection
    Set list = ModuleCSV.OpenCSVFile()
    If list Is Nothing Then Exit Sub

    Dim lineVariant As Variant
    Dim lineStr As String
    Dim items() As String
    Dim rowNum As Long
    Dim colNum As Long
    Dim currentTimestamp As Date
    
    currentTimestamp = Now() ' Фиксируем время один раз
    rowNum = 1 ' Начинаем запись с первой строки (A1)

    ' Отключаем обновление экрана для ускорения работы
    Application.ScreenUpdating = False

    For Each lineVariant In list
        lineStr = CStr(lineVariant)
        
        ' Разбиваем строку на массив элементов по точке с запятой
        items = Split(lineStr, ";")
        
        ' Записываем элементы строки в ячейки
        For colNum = 0 To UBound(items)
            wsBaseData.Cells(rowNum, colNum + 1).Value = items(colNum)
        Next colNum
        
        ' В последний дополнительный столбец записываем TimeStamp
        wsBaseData.Cells(rowNum, UBound(items) + 2).Value = currentTimestamp
        
        rowNum = rowNum + 1 ' Переходим на следующую строку
    Next lineVariant

    Application.ScreenUpdating = True
    MsgBox "Данные успешно импортированы!", vbInformation
End Sub

Public Sub Export()
    Dim filePath As Variant
    Dim fNum As Integer

    Dim baseArr As Variant
    Dim lastRow As Long

    ' Диалог
    Dim dataTime As Date: dataTime = Now
    filePath = Application.GetSaveAsFilename( _
    InitialFileName:="base" & Format(dataTime, "_yyyymmdd_hhmmss") & ".csv", _
    FileFilter:="CSV Files (*.csv), *.csv" _
    )

    If filePath = False Then Exit Sub

        lastRow = wsBaseData.Cells(wsBaseData.Rows.Count, 1).End(xlUp).row
        baseArr = wsBaseData.Range(wsBaseData.Cells(1, 1), wsBaseData.Cells(lastRow, TotalCols)).value

        fNum = FreeFile
        Open filePath For Output As #fNum

        ' --- Заголовки ---
        Print #fNum, "FNM;BIRTH;STAT;DATESTAT;DATEINS;PERIOD;TIMESTAMP"
        ' --- Данные ---
        Dim i As Long, j As Long
        Dim line As String
        For i = 1 To lastRow
            line = ""
            For j = 1 To TotalCols
                line = line & baseArr(i, j)
                If j < TotalCols Then line = line & ";"
                Next j
                Print #fNum, line
            Next i

            Close #fNum

            MsgBox "База экспортирована в " & filePath, vbInformation
End Sub

Private Sub LoadPrimaryBase()
    wsBaseData.Range("A1:G1").value = Array("Гордон Фриман", "19.10.1982", "МСМК", "03.05.2020", "01.01.2026", "365", Format$(Now, "dd.mm.yyyy"))
End Sub

Private Function GetBaseData() As Variant
    Dim lastRow As Long, lastCol As Long
    lastRow = wsBaseData.Cells(wsBaseData.Rows.Count, 1).End(xlUp).row
    GetBaseData = wsBaseData.Range(wsBaseData.Cells(1, 1), wsBaseData.Cells(lastRow, TotalCols)).value
End Function

Private Sub SetUnBaseKeys()
    Set UnBaseKeys = CreateObject("Scripting.Dictionary"): UnBaseKeys.CompareMode = 1
    Dim key As Variant
    For Each key In BaseKeys.Keys
        UnBaseKeys.Add BaseKeys.Item(key), key
    Next key 
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
    Dim athleteKey As String

    Dim rank As String
    Dim bd As String

    Dim dataP(0 to 5) As Variant
    Dim i As Long
    Dim clubName As String: clubName = "Import"

    Set clubs = CreateObject("Scripting.Dictionary")
    clubs.CompareMode = 1
    Set dataP = CreateObject("Scripting.Dictionary")
    dataP.CompareMode = 1

    For i = 1 To lines.Count
        lineData = lines(i)
        cols = Split(lineData, ";")

        ' ФИО#ДР Разряд
        If UBound(cols) >= 5 Then
            athleteKey = Trim(cols(0))
            
            bd = Trim(cols(1))
            rank = Trim(cols(2))
            rank = Trim(cols(2))
            rank = Trim(cols(2))

            If Not clubs.Exists(clubName) Then
                Set athleteDict = CreateObject("Scripting.Dictionary")
                clubs.Add clubName, athleteDict
            Else
                Set athleteDict = clubs(clubName)
            End If

            athleteDict.Add athleteKey, dataP
        End If
    Next i

    Set ParseAthletes = clubs
End Function