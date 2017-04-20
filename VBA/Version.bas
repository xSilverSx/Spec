Attribute VB_Name = "Version"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit
Public iVersion As Integer
'Public iQuestion As Integer

Sub Сохранить_Версию_Спецификации()
    Dim a As Integer
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
            If IsWorkSheetExist("Версии") = False Then
                Запись_Новой_Версии
            End If
        a = MsgBox("Сохранить в архиве старую версию?", vbYesNo)
            
            If a = vbYes Then
                ActiveWorkbook.Sheets("Спецификация").Copy After:=ActiveWorkbook.Sheets(1)
                ActiveWorkbook.Sheets("СО").Copy After:=ActiveWorkbook.Sheets(2)
                ActiveWorkbook.Sheets("ВР").Copy After:=ActiveWorkbook.Sheets(3)
                iVersion = Sheets("Спецификация").Cells(1, 25).Value
                Sheets("Спецификация (2)").Name = "Спецификация_" & iVersion
                Sheets("СО (2)").Name = "СО_" & iVersion
                Sheets("ВР (2)").Name = "ВР_" & iVersion
                Sheets("Спецификация_" & iVersion).Visible = xlSheetVeryHidden
                Sheets("СО_" & iVersion).Visible = xlSheetVeryHidden
                Sheets("ВР_" & iVersion).Visible = xlSheetVeryHidden
                Запись_Новой_Версии
            End If
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
End Sub

Sub Обновить_дату_последней_версии()
    Dim lLastRow As Integer
        If IsWorkSheetExist("Версии") = False Then
            Запись_Новой_Версии
        End If
        lLastRow = ActiveWorkbook.Sheets("Версии").UsedRange.Row + Sheets("Версии").UsedRange.Rows.Count - 1
        Sheets("Версии").Cells(lLastRow, 2).Value = Date
        
        ActiveWorkbook.BuiltinDocumentProperties(32).Value = "Версий: " & lLastRow - 1 & _
        "; " & Date & "-Дата последней версии"
         ActiveWorkbook.Sheets("Версии").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Sub Оставить_одну_версию()
    Dim lLastRow As Integer, iDelList As Integer, RwsDel As String
    Dim a As Integer
        
        If IsWorkSheetExist("Версии") = False Then
                MsgBox "В этом файле только одна версия", vbCritical
                End
            Else
                ActiveWorkbook.Sheets("Версии").Unprotect 'Снятие защиты листа
                Application.ScreenUpdating = False
                Application.DisplayAlerts = False
                a = MsgBox("Оставить только последнюю версию?", vbYesNo)
                If a = vbYes Then
                    lLastRow = ActiveWorkbook.Sheets("Версии").UsedRange.Row + Sheets("Версии").UsedRange.Rows.Count - 1
                    iDelList = Sheets("Версии").Cells(lLastRow, 1).Value
                    For i = 1 To lLastRow
                        If IsWorkSheetExist("Спецификация_" & i) Then
                            Sheets("Спецификация_" & i).Visible = True
                            Sheets("СО_" & i).Visible = True
                            Sheets("ВР_" & i).Visible = True
                            Sheets("Спецификация_" & i).Delete
                            Sheets("СО_" & i).Delete
                            Sheets("ВР_" & i).Delete
                        End If
                    Next i
                RwsDel = "2:" & lLastRow
                Sheets("Версии").Rows(RwsDel).Delete
                Запись_Новой_Версии
        End If
        End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub Показать_Форму_Версии()
        VersionForm.Show
End Sub

Sub Запись_Новой_Версии()
    Dim lLastRow As Integer, iVersion As Integer
        If IsWorkSheetExist("Версии") = False Then
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets("Версии").Copy Before:=ActiveWorkbook.Sheets(1)
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
        End If
        lLastRow = ActiveWorkbook.Sheets("Версии").UsedRange.Row + Sheets("Версии").UsedRange.Rows.Count - 1 'определить количество заполненых ячеек
        If Sheets("Версии").Cells(lLastRow, 1).Value = "" Then
            Sheets("Версии").Cells(lLastRow, 1).Value = 1
            Sheets("Версии").Cells(lLastRow, 2).Value = Date
            iVersion = 1
        Else
            iVersion = Sheets("Версии").Cells(lLastRow, 1).Value + 1
            Sheets("Версии").Cells(lLastRow + 1, 1).Value = iVersion
            Sheets("Версии").Cells(lLastRow + 1, 2).Value = Date
        End If
            Sheets("Спецификация").Cells(1, 25).Value = iVersion
            Sheets("СО").Cells(1, 1).Value = iVersion 'обозначить версию на листе
            Sheets("СО").Cells(1, 1).Font.ThemeColor = xlThemeColorDark1 'скрыть видимость версии
            Sheets("ВР").Cells(1, 1).Value = iVersion
            Sheets("ВР").Cells(1, 1).Font.ThemeColor = xlThemeColorDark1
            
            ActiveWorkbook.BuiltinDocumentProperties(32).Value = "Версий: " & lLastRow & _
            "; " & Date & "-Дата последней версии"
            ActiveWorkbook.Sheets("Версии").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'Включение защиты листа
            
End Sub

Sub СнятьЗащиту() 'Снять защиту с ячеек которые можно изменять
    Dim lLastRow As Integer
        ActiveWorkbook.Sheets("Версии").Unprotect 'Снятие защиты листа
        ActiveWorkbook.Sheets("Версии").Activate
        lLastRow = ActiveWorkbook.Sheets("Версии").UsedRange.Row + Sheets("Версии").UsedRange.Rows.Count - 1 'определить количество заполненых ячеек
        Range(Cells(2, 1), Cells(lLastRow, 4)).Locked = False
        ActiveWorkbook.Sheets("Версии").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True 'Включение защиты листа
End Sub




