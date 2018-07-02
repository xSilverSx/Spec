VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Podgotovka 
   Caption         =   "Подготовить спецификацию"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   OleObjectBlob   =   "Podgotovka.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Podgotovka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Private Sub Begin_Click()
' Dim t As Single
'    t = Timer
    Перенос
    Podgotovka.Hide
'    MsgBox "Обработка данных продолжалась  " & Timer - t & " сек.", vbInformation
End Sub
Private Sub Cancel_Click()
    Podgotovka.Hide
End Sub

Private Sub Perenos_Click()
    If Perenos = True And SO = True And VR = True Then VR.Value = False
End Sub

Private Sub SO_Click()
    If SO = False And VR = False Then VR.Value = True
    If Perenos = True And SO = True And VR = True Then VR.Value = False
End Sub
Private Sub VR_Click()
    If SO = False And VR = False Then SO.Value = True
    If Perenos = True And SO = True And VR = True Then SO.Value = False
End Sub
Private Sub Closed_Click()
    Unload Podgotovka
End Sub

