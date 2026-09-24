Option Explicit

Private checks As Long
Private Const HEADER As String = "FNM;BIRTH;STAT;DATESTAT;DATEINS;PERIOD;TIMESTAMP"

Private Sub Check(condition As Boolean, message As String)
    If Not condition Then Err.Raise vbObjectError + 1900, "ImportTests", message
    checks = checks + 1
End Sub

Private Function Lines(ParamArray values() As Variant) As Collection
    Dim value As Variant
    Set Lines = New Collection
    For Each value In values
        Lines.Add CStr(value)
    Next value
End Function

Private Sub InvalidRows(rows As Collection, expectedMessage As String)
    Dim result As Object
    Dim errorNumber As Long, errorMessage As String
    On Error Resume Next
    Set result = Base.TestParse(rows)
    errorNumber = Err.Number
    errorMessage = Err.Description
    On Error GoTo 0
    Check errorNumber = vbObjectError + 1202, "Malformed CSV was accepted: " & expectedMessage
    Check InStr(errorMessage, expectedMessage) > 0, "Missing diagnostic: " & errorMessage
End Sub

Private Sub InvalidFile(path As String)
    Dim result As Collection
    Dim errorNumber As Long
    On Error Resume Next
    Set result = ModuleCSV.TestRead(path, "windows-1251", HEADER)
    errorNumber = Err.Number
    On Error GoTo 0
    Check errorNumber = vbObjectError + 1202, "Invalid header was accepted: " & path
End Sub

Private Sub Review(clubs As Object)
    Dim selected As New Collection
    selected.Add "Import"
    ClubsListData.InitLists selected, clubs
    ClubsListData.SetClubInCombinedList 1
End Sub

Public Function RunAll(folder As String) As String
    checks = 0
    Dim rows As Collection, clubs As Object, athletes As Object
    Dim data As Variant, view As Variant, current As Variant
    Dim key As Variant, file As Variant

    For Each file In Array("ansi.csv", "utf8.csv")
        Set rows = ModuleCSV.TestRead(folder & "\" & file, "windows-1251", HEADER)
        Check rows.Count = 1, "CSV row count: " & file
        Set clubs = Base.TestParse(rows)
        Set athletes = clubs.Item("Import")
        Check athletes.Exists("Иванов Иван#01.02.2000"), "Cyrillic name/date corrupted: " & file
        data = athletes.Item("Иванов Иван#01.02.2000")
        Check data(1) = "КМС", "Cyrillic rank corrupted"
        Check data(2) = DateSerial(2025, 4, 3), "Rank date lost"
        Check data(3) = DateSerial(2026, 6, 5), "Insurance date lost"
        Check data(4) = 365, "Insurance period lost"
    Next file
    Set rows = ModuleCSV.TestRead(folder & "\orgeo.csv", "utf-8", "")
    Check rows(1) = "Иванов Иван;КМС", "Orgeo UTF-8 reading changed"
    Set rows = ModuleCSV.TestRead(folder & "\header-only.csv", "windows-1251", HEADER)
    Check rows.Count = 0, "Header-only file should have no participants"
    InvalidFile folder & "\wrong-header.csv"
    InvalidFile folder & "\empty.csv"

    InvalidRows Lines("Test;01.02.2000;КМС;;;365"), "7 столбцов"
    InvalidRows Lines("Test;01.02.2000;КМС;;;365;;extra"), "7 столбцов"
    InvalidRows Lines(";01.02.2000;КМС;;;365;"), "ФИО"
    InvalidRows Lines("Test#Other;01.02.2000;КМС;;;365;"), "ФИО"
    InvalidRows Lines("Test;not a date;КМС;;;365;"), "дата рождения"
    InvalidRows Lines("Test;01.02.2000;КМС;bad;;365;"), "столбце 4"
    InvalidRows Lines("Test;01.02.2000;КМС;;bad;365;"), "столбце 5"
    InvalidRows Lines("Test;01.02.2000;КМС;;;bad;"), "период страховки"
    InvalidRows Lines("Test;01.02.2000;КМС;;;-1;"), "целым неотрицательным"
    InvalidRows Lines("Test;01.02.2000;КМС;;;1,5;"), "целым неотрицательным"
    InvalidRows Lines("Test;01.02.2000;КМС;;;2147483648;"), "целым неотрицательным"
    InvalidRows Lines("Test;01.02.2000;КМС;;;365;", " test ;01.02.2000;КМС;;;;"), "повтор участника"

    Base.Init
    Base.wsBaseData.Range("A1:G1").Value = Array("Иванов Иван", DateSerial(2000, 2, 1), "I", DateSerial(2024, 1, 1), DateSerial(2026, 1, 1), 180, DateSerial(2026, 1, 1))
    Base.wsBaseData.Range("A2:G2").Value = Array("Unrelated", DateSerial(1990, 1, 1), "I", "", "", "", DateSerial(2026, 1, 1))
    Base.LoadToShowcase
    ' Unsaved edits in the cache must participate in the comparison.
    Base.UpdateBase "Иванов Иван", DateSerial(2000, 2, 1), "МС", DateSerial(2024, 2, 1), DateSerial(2026, 1, 1), 180, DateSerial(2026, 1, 1), 1
    Set rows = ModuleCSV.TestRead(folder & "\ansi.csv", "windows-1251", HEADER)
    rows.Add "New Athlete;03.04.2005;;;;;"
    Set clubs = Base.TestParse(rows)
    Review clubs
    view = ClubsListData.GetClubForShowcase(1)
    Check UBound(view, 1) = 3, "Expected base/update pair and one new athlete"
    Check view(1, 2) = ClubsListData.arrStatus(0), "Missing base reference row"
    Check view(1, 5) = "МС", "Comparison ignored the current cache"
    Check view(1, 6) = DateSerial(2024, 2, 1), "Base reference row changed"
    Check view(2, 2) = ClubsListData.arrStatus(1), "Missing update row"
    Check view(2, 3) = 1, "Existing participant ID lost"
    Check view(2, 5) = "КМС", "Imported rank lost"
    Check view(2, 6) = DateSerial(2025, 4, 3), "Imported rank date lost in review"
    Check view(2, 8) = DateSerial(2026, 6, 5), "Imported insurance date lost in review"
    Check view(2, 9) = 365, "Imported period lost in review"
    Check view(3, 2) = ClubsListData.arrStatus(2), "Missing new participant status"
    Check view(3, 6) = "" And view(3, 8) = "" And view(3, 9) = "", "Optional blanks became dates or numbers"
    Check Base.getMaxId = 2, "Review changed base size"
    current = Base.GetParticipant("Иванов Иван", DateSerial(2000, 2, 1))
    Check current(1) = "МС", "Review overwrote the cache"
    Check Base.wsBaseData.Cells(1, 3).Value = "I", "Review overwrote persisted data"

    ' The existing review-edit-accept path must work for CSV without export changes.
    ClubsListData.UpdateParticipant 1, "Иванов Иван#01.02.2000#" & ClubsListData.arrStatus(1), "КМС", DateSerial(2025, 4, 3), DateSerial(2026, 6, 5), 366
    ClubsListData.ExportToBase 1
    current = Base.GetParticipant("Иванов Иван", DateSerial(2000, 2, 1))
    Check current(1) = "КМС" And current(4) = 366, "Reviewed edits were not accepted"
    Check Base.getMaxId = 3, "New participant not inserted exactly once"
    Check IsArray(Base.GetParticipant("New Athlete", DateSerial(2005, 4, 3))), "New participant missing"
    Check IsArray(Base.GetParticipant("Unrelated", DateSerial(1990, 1, 1))), "Unrelated participant removed"
    Check CDate(Base.wsBaseData.Cells(1, 7).Value) = Date, "Acceptance timestamp is not current"
    Review clubs
    view = ClubsListData.GetClubForShowcase(1)
    Check UBound(view, 1) = 4, "Reimport should match both participants"
    ClubsListData.ExportToBase 1
    Check Base.getMaxId = 3, "Reimport duplicated participants"

    ' Even when rank is unchanged, full CSV data must supply the new dates/blanks.
    Set clubs = Base.TestParse(Lines("Иванов Иван;01.02.2000;КМС;01.07.2026;;;"))
    Review clubs
    view = ClubsListData.GetClubForShowcase(1)
    Check view(2, 6) = DateSerial(2026, 7, 1), "Same-rank CSV date was discarded"
    Check view(2, 8) = "" And view(2, 9) = "", "Explicit CSV blanks were discarded"

    ' Legacy Orgeo rank-only payload retains its existing merge rules.
    Set athletes = CreateObject("Scripting.Dictionary")
    athletes.Add "Иванов Иван#01.02.2000", "КМС"
    athletes.Add "Orgeo New#01.01.2007", "I"
    Set clubs = CreateObject("Scripting.Dictionary")
    clubs.Add "Import", athletes
    Review clubs
    view = ClubsListData.GetClubForShowcase(1)
    Check view(2, 6) = DateSerial(2025, 4, 3), "Orgeo same-rank date preservation changed"
    Check view(2, 8) = DateSerial(2026, 6, 5) And view(2, 9) = 365, "Orgeo insurance preservation changed"
    Check view(3, 5) = "I" And view(3, 2) = ClubsListData.arrStatus(2), "Orgeo new participant changed"
    athletes.Item("Иванов Иван#01.02.2000") = "МС"
    Review clubs
    view = ClubsListData.GetClubForShowcase(1)
    Check view(2, 5) = "МС" And view(2, 6) = "", "Orgeo changed-rank rule changed"
    Check view(2, 8) = DateSerial(2026, 6, 5) And view(2, 9) = 365, "Orgeo changed-rank insurance rule changed"

    RunAll = "PASS: " & checks & " assertions (CSV, reconciliation, acceptance, reimport, Orgeo regression)"
End Function
