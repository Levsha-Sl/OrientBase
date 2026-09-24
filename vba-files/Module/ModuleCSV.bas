Attribute VB_Name = "ModuleCSV"
Option Private Module

Public Function OpenCSVFile(Optional charset As String = "utf-8", Optional expectedHeader As String = "") As Collection
    Dim filePath As Variant
    Dim lines As Collection

    filePath = SelectCSVFile()

    If VarType(filePath) = vbBoolean Then
        Set OpenCSVFile = Nothing
     Exit Function
    End If

    Set lines = ReadCSVFile(CStr(filePath), charset, expectedHeader)
    Set OpenCSVFile = lines
End Function

Private Function SelectCSVFile() As Variant
    SelectCSVFile = Application.GetOpenFilename( _
    FileFilter:="CSV Files (*.csv), *.csv", _
    Title:="Выберите CSV файл для импорта (ADODB)")
End Function

Private Function ReadCSVFile(filePath As String, charset As String, expectedHeader As String) As Collection
    On Error GoTo ErrorHandler
    Dim objStream As Object
    Dim lineData As String
    Dim lines As Collection

    Set lines = New Collection

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary: проверяем BOM перед выбором кодировки.
    objStream.Open
    objStream.LoadFromFile filePath
    If objStream.Size >= 3 Then
        Dim prefix As Variant
        prefix = objStream.Read(3)
        If prefix(0) = &HEF And prefix(1) = &HBB And prefix(2) = &HBF Then charset = "utf-8"
    End If
    objStream.Position = 0
    objStream.Type = 2 ' adTypeText
    objStream.Charset = charset

    If Not objStream.EOS Then
        lineData = objStream.ReadText(-2)
    End If
    If expectedHeader <> "" Then
        If StrComp(Trim$(lineData), expectedHeader, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 1202, "ModuleCSV.ReadCSVFile", "Неверный заголовок CSV. Выберите файл экспорта базы."
        End If
    End If

    Do Until objStream.EOS
        lineData = objStream.ReadText(-2)
        If Trim(lineData) <> "" Then
            lines.Add lineData
        End If
    Loop

    objStream.Close
    Set ReadCSVFile = lines
    Exit Function

ErrorHandler:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not objStream Is Nothing Then objStream.Close
    On Error GoTo 0
    Err.Raise errorNumber, "ModuleCSV.ReadCSVFile", errorDescription
End Function
