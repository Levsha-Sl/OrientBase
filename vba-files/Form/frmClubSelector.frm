VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClubSelector 
   Caption         =   "Загурзка инофрмации с Orgeo"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "frmClubSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClubSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private pResult As Collection
Private pCancelled As Boolean

Public Property Get Result() As Collection
    Set Result = pResult
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = pCancelled
End Property

Public Sub Init(ByVal clubs As Object)

    Dim key As Variant
    Dim chk As MSForms.CheckBox
    Dim topPos As Double

    topPos = 6

    fraList.ScrollTop = 0

    ' очищаем старые чекбоксы
    Do While fraList.Controls.Count > 0
        fraList.Controls.Remove fraList.Controls(0).name
    Loop

    For Each key In clubs.Keys

        Set chk = fraList.Controls.Add("Forms.CheckBox.1")

        With chk
            .Caption = CStr(key) & " (" & clubs(key) & ")"
            .name = "chk_" & Replace(CStr(key), " ", "_")
            .Left = 8
            .Top = topPos
            .Width = fraList.Width - 25
            .Height = 18
            .value = True
        End With

        topPos = topPos + 20

    Next key

    fraList.ScrollHeight = topPos + 10

End Sub

Private Sub cmdSelectAll_Click()

    Dim ctrl As Control

    For Each ctrl In fraList.Controls
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = True
        End If
    Next ctrl

End Sub

Private Sub cmdUnselectAll_Click()

    Dim ctrl As Control

    For Each ctrl In fraList.Controls
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.value = False
        End If
    Next ctrl

End Sub

Private Sub cmdLoad_Click()

    Dim ctrl As Control
    Dim txt As String

    Set pResult = New Collection
    pCancelled = False

    For Each ctrl In fraList.Controls

        If TypeName(ctrl) = "CheckBox" Then

            If ctrl.value = True Then

                txt = ctrl.Caption

                txt = Left$(txt, InStrRev(txt, "(") - 2)

                pResult.Add txt

            End If

        End If

    Next ctrl

    Me.Hide

End Sub

Private Sub cmdCancel_Click()

    pCancelled = True
    Me.Hide

End Sub
