Attribute VB_Name = "ClubSelector"
Option Explicit

Public Function Show(ByVal clubs As Object) As Collection
    Dim frm As frmClubSelector
    Set frm = New frmClubSelector

    frm.Init clubs
    frm.Show vbModal

    If frm.Cancelled Then
        Set Show = Nothing
    Else
        Set Show = frm.Result
    End If

    Unload frm
    Set frm = Nothing
End Function
