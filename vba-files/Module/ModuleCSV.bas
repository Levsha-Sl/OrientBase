Attribute VB_Name = "ModuleCSV"
Option Private Module

Public Function OpenCSVFile() As Collection
    Dim filePath As Variant
    Dim lines As Collection

    filePath = SelectCSVFile()

    If VarType(filePath) = vbBoolean Then
        Set OpenCSVFile = Nothing
     Exit Function
    End If

    Set lines = ReadCSVFile(CStr(filePath))
    Set OpenCSVFile = lines
End Function

Private Function SelectCSVFile() As Variant
    SelectCSVFile = Application.GetOpenFilename( _
    FileFilter:="CSV Files (*.csv), *.csv", _
    Title:="Выберите CSV файл для импорта (ADODB)")
End Function

Private Function ReadCSVFile(filePath As String) As Collection
    Dim objStream As Object
    Dim lineData As String
    Dim lines As Collection

    Set lines = New Collection

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile filePath

    If Not objStream.EOS Then
        objStream.ReadText (-2)
    End If

    Do Until objStream.EOS
        lineData = objStream.ReadText(-2)
        If Trim(lineData) <> "" Then
            lines.Add lineData
        End If
    Loop

    objStream.Close
    Set ReadCSVFile = lines
End Function
