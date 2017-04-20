VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} barForm 
   ClientHeight    =   165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1890
   OleObjectBlob   =   "barForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "barForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ans = MsgBox("Нажмите ""Да"", чтобы остановить процесс." & Chr(13) & _
            "Нажмите ""Нет"", ""Отмена"" или закройте это окно, чтобы закрыть только окно загрузки.", _
            vbYesNoCancel + vbQuestion, "Остановить процесс?")
        If ans = 6 Then
            End
        Else
            Unload barForm
        End If
    End If
End Sub
