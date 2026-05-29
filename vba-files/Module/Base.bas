Attribute VB_Name = "Base"
Public wsBaseData As Worksheet
' (FNM#BIRTH)key STAT DATESTAT DATEINS PERIOD
Private BaseDict As Object
' (ID)key FNM#BIRTH
Private BaseKeys As Object
Private MaxId As Long
Private TotalCols As Long

Public Function getMaxId() As Long
    getMaxId = MaxId
End Function

Public Function Init() As Worksheet
    If wsBaseData Is Nothing Then
        ' (FNM BIRTH)key STAT DATESTAT DATEINS PERIOD
        Set wsBaseData = DataSheets.GetSheet("Base")
        If wsBaseData.Cells(1, 1).Value = "" Then 
            Call LoadPrimaryBase
        End If
    End If
    TotalCols = 7 ' Фиксированное количество столбцов (FNM BIRTH STAT DATESTAT DATEINS PERIOD TIMESTAMP)
    Set Init = wsBaseData
End Function

Public Sub LoadToMart()
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
            .Range("A1").Resize(totalRows, TotalCols).Value = resultBase
        End With
    End If
    Call LoadToMart
End Sub

Public Sub UpdateBase(fnm As String, birth As Date, stat As String, dateStat As Variant, dateIns As Variant, period As Variant, timeStamp As Variant, id As Long)
    Dim oldKey As String, newKey As String

    If BaseKeys.Exists(id) Then
        oldKey = BaseKeys(id)
        newKey = fnm & "#" & birth

        If oldKey <> newKey Then
            ' Ключ изменился - удаляем старый, добавляем новый
            BaseKeys(id) = newKey
            BaseDict.Key(oldKey) = newKey
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
    Else
        Err.Raise vbObjectError + 1201, "CreateBase", "Спортсмен с такими ФИО и датой рождения уже существует"
    End If
    CreateBase = MaxId
End Function

Public Function GetBaseValue(fnm As String, birth As Date) As Variant
    Dim key As String: key = fnm & "#" & birth
    If BaseDict.Exists(key) Then
        GetBaseValue = BaseDict.Item(key)
    Else
        Err.Raise vbObjectError + 1202, "GetBaseValue", "Спортсмен с таким наименованием не найден"
    End If
End Function

' TODO ниже код не трогать
Public Sub Import()

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

        lastRow = wsBaseData.Cells(wsBaseData.Rows.Count, 1).End(xlUp).Row
        baseArr = wsBaseData.Range(wsBaseData.Cells(1, 1), wsBaseData.Cells(lastRow, TotalCols)).Value

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
    wsBaseData.Range("A1:G1").Value = Array("Гордон Фриман", "19.10.1982", "МСМК", "03.05.2020", "01.01.2026", "365", Format$(Now, "dd.mm.yyyy"))
End Sub

Private Function GetBaseData() As Variant
    Dim lastRow As Long, lastCol As Long
    lastRow = wsBaseData.Cells(wsBaseData.Rows.Count, 1).End(xlUp).Row
    GetBaseData = wsBaseData.Range(wsBaseData.Cells(1, 1), wsBaseData.Cells(lastRow, TotalCols)).Value
End Function