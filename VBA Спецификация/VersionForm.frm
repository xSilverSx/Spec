VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VersionForm 
   Caption         =   "Список версий"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   OleObjectBlob   =   "VersionForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VersionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
'Public iQuestion As Integer

Private Sub CommandButton1_Click()
    Оставить_одну_версию
End Sub

Private Sub CommandButton2_Click()
    Сохранить_Версию_Спецификации
End Sub

Private Sub CommandButton3_Click()
    Dim a As Integer
    If ListBox1.ListIndex = -1 Then
        MsgBox "Не выбрана строка", vbCritical
    Else
    a = MsgBox("Действительно удалить эту версию спецификации?", vbYesNo)
    If a = vbYes Then
    i = ListBox1.Value
        If IsWorkSheetExist("Спецификация_" & i) Then
            Sheets("Спецификация_" & i).Visible = True
            Sheets("СО_" & i).Visible = True
            Sheets("ВР_" & i).Visible = True
                Application.DisplayAlerts = False
                    Sheets("Спецификация_" & i).Delete
                    Sheets("СО_" & i).Delete
                    Sheets("ВР_" & i).Delete
                Application.DisplayAlerts = True
    i = ListBox1.ListIndex + 2
    ActiveWorkbook.Sheets("Версии").Unprotect 'Снятие защиты листа
        Sheets("Версии").Rows(i & ":" & i).Delete
    End If
    End If
    End If
    СнятьЗащиту
End Sub

Private Sub CommandButton4_Click()
Dim iVersion As Integer, lLastRow As Integer, iQuestion As Integer
    If ListBox1.ListIndex = -1 Then
        MsgBox "Не выбрана версия на которую слудует заменить", vbCritical
        Exit Sub
    End If
    iVersion = ListBox1.Value
        If IsWorkSheetExist("Спецификация_" & iVersion) = False Then
            MsgBox "Выбрана текущая версия, или этой версии не существует", vbCritical
            Exit Sub
        End If
    iQuestion = MsgBox("Текущая версия будет перезаписана.", vbYesNoCancel)
        If iQuestion = vbCancel Or iQuestion = vbNo Then Exit Sub
'        ElseIf iQuestion = vbYes Then
'            Сохранить_Версию_Спецификации
'        End If
Application.DisplayAlerts = False
    Sheets("Спецификация_" & iVersion).Visible = True
    Sheets("СО_" & iVersion).Visible = True
    Sheets("ВР_" & iVersion).Visible = True
                Sheets("Спецификация").Delete
                Sheets("СО").Delete
                Sheets("ВР").Delete
        Sheets("ВР_" & iVersion).Copy After:=ActiveWorkbook.Sheets(1)
        Sheets("СО_" & iVersion).Copy After:=ActiveWorkbook.Sheets(2)
        Sheets("Спецификация_" & iVersion).Copy After:=ActiveWorkbook.Sheets(3)
            Sheets("Спецификация_" & iVersion & " (2)").Name = "Спецификация"
            Sheets("СО_" & iVersion & " (2)").Name = "СО"
            Sheets("ВР_" & iVersion & " (2)").Name = "ВР"
    Sheets("Спецификация_" & iVersion).Visible = False
    Sheets("СО_" & iVersion).Visible = False
    Sheets("ВР_" & iVersion).Visible = False
    Обновить_дату_последней_версии
    
    lLastRow = ActiveWorkbook.Sheets("Версии").UsedRange.Row + Sheets("Версии").UsedRange.Rows.Count - 1
    iVersion = Sheets("Версии").Cells(lLastRow, 1).Value
    Sheets("Спецификация").Cells(1, 25).Value = iVersion
    Sheets("СО").Cells(1, 1).Value = iVersion 'обозначить версию на листе
    Sheets("СО").Cells(1, 1).Font.ThemeColor = xlThemeColorDark1 'скрыть видимость версии
    Sheets("ВР").Cells(1, 1).Value = iVersion
    Sheets("ВР").Cells(1, 1).Font.ThemeColor = xlThemeColorDark1
Application.DisplayAlerts = True
End Sub

Private Sub ListBox1_Change()
    TextBox1.Value = ListBox1.Column(3)
End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim iVersion As Integer
    iVersion = ListBox1.Value
    If IsWorkSheetExist("Спецификация_" & iVersion) Then
        If Sheets("Спецификация_" & iVersion).Visible = xlSheetVisible Then
            Sheets("Спецификация_" & iVersion).Visible = xlSheetVeryHidden
            Sheets("СО_" & iVersion).Visible = xlSheetVeryHidden
            Sheets("ВР_" & iVersion).Visible = xlSheetVeryHidden
        Else
            Sheets("Спецификация_" & iVersion).Visible = xlSheetVisible
            Sheets("СО_" & iVersion).Visible = xlSheetVisible
            Sheets("ВР_" & iVersion).Visible = xlSheetVisible
            Sheets("Спецификация_" & iVersion).Activate
        End If
    Else
    MsgBox "Выбрана текущая версия, или этой версии не существует", vbCritical
    End If
End Sub

Private Sub UserForm_Initialize()
Dim sItemList As String
Dim lLastRow As Integer
If IsWorkSheetExist("Версии") = False Then
    MsgBox "В этом файле только одна версия", vbCritical
    End
Else
    lLastRow = ActiveWorkbook.Sheets("Версии").UsedRange.Row + Sheets("Версии").UsedRange.Rows.Count - 1
    sItemList = "=Версии!A2:D" & lLastRow
    ListBox1.ColumnWidths = "40;60;30;0"
    ListBox1.RowSource = sItemList
End If
End Sub


