VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClearCont 
   Caption         =   "Что чистим?"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2865
   OleObjectBlob   =   "ClearCont.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ClearCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ChListPerenos_Click()
If ChListPerenos = False And ChListSO = False And ChListVR = False Then ChListSO = True
End Sub

Private Sub ChListSO_Click()
If ChListPerenos = False And ChListSO = False And ChListVR = False Then ChListPerenos = True
End Sub

Private Sub ChListVR_Click()
If ChListPerenos = False And ChListSO = False And ChListVR = False Then ChListPerenos = True
End Sub
Private Sub ClearAllList_Butt_Click()
Dim a As String, b As String, c As String
Dim e As Byte
Application.ScreenUpdating = False
    Чистка_Печати
    Worksheets("Спецификация").Activate
If ChListPerenos = True Then
    a = d & "Перенос" & d & " "
    e = e + 1
End If
If ChListSO = True Then
    b = d & "СО" & d & " "
    e = e + 1
End If
If ChListVR = True Then
    c = d & "ВР" & d
    e = e + 1
End If
ClearCont.Hide
Application.ScreenUpdating = True
    If e > 1 Then
        MsgBox ("Листы " & a & b & c & " Очищены")
        Else
        MsgBox ("Лист " & a & b & c & " Очищен")
    End If
End Sub

Function d()
    d = Chr(34)
End Function
