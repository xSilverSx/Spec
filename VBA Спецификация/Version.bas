Attribute VB_Name = "Version"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit
Public iVersion As Integer
Public ShActive As Variant

Sub CreateNewSpec() 'Создать новый лист спецификации, сохранить его в папку со спецификацией
Dim b As Byte
    If IsBookOpen("Спецификация 1.xlsx") = False Then
        Delete_File (ThisWorkbook.Path & "\" & "Спецификация 1.xlsx")
        b = OpenFolderBook("Template-Spec", "xlsx")   'Открыть файл шаблона
        WbOpenFile.SaveAs ThisWorkbook.Path & "\" & "Спецификация 1"    'Сохранить новый файл под другим именем
    Else
        MsgBox "Книга Спецификация 1 уже открыта, закройте или переименуйте ее чтобы создать новый шаблон"
    End If
End Sub

Sub AddWorkbookFooter() 'Добавить основную надпись в книгу
Dim b As Boolean
Dim ShActive As Sheets
    b = ListSpec 'Проверяем что книга является спецификацией
    If b Then
        AddFooterStamp ("СО")
        AddFooterStamp ("ВР")
    End If
End Sub

Sub AddFooterStamp(StrList As String) 'Добавить картинки штампов в колонтитул слева на листы
Dim StrFolder As String 'Путь расположения надстройки
StrFolder = ThisWorkbook.Path
    If FileLocation(StrFolder & "\Page2.png") And FileLocation(StrFolder & "\Page1.png") Then
        ActiveWorkbook.Sheets(StrList).PageSetup.LeftFooterPicture.filename = StrFolder & "\Page2.png"
        ActiveWorkbook.Sheets(StrList).Columns("A:D").Clear
        With ActiveWorkbook.Sheets(StrList)
            .Rows("28:28").RowHeight = 24   'Корректировка высоты строки, для выравнивания штампа
            With .PageSetup.LeftFooterPicture
                .Height = 247.5
                .Width = 52.5
            End With
            .PageSetup.FirstPage.LeftFooter.Picture.filename = StrFolder & "\Page1.png"
            With .PageSetup.FirstPage.LeftFooter.Picture
                .Height = 135.55
                .Width = 52.5
            End With
            With .PageSetup
                .LeftFooter = "&G"
                .FirstPage.LeftFooter.Text = "&G"
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = True
                .ScaleWithDocHeaderFooter = True
                .AlignMarginsHeaderFooter = True
                .LeftMargin = Application.InchesToPoints(0.196850393700787)
                .RightMargin = Application.InchesToPoints(0.196850393700787)
                .TopMargin = Application.InchesToPoints(0.196850393700787)
                .BottomMargin = Application.InchesToPoints(0.196850393700787)
                .HeaderMargin = Application.InchesToPoints(0.196850393700787)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
            End With
        End With
    Else
    MsgBox "Основная надпись не обновлена" & Chr(13) & "Файлы Page1.png и(или) Page2.png не найдены." & _
    Chr(13) & "Файлы можно найти по ссылке https://github.com/xSilverSx/Spec", vbCritical
    End If
End Sub

Sub Сохранить_Версию_Спецификации()
    Dim A As Integer
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
            If IsWorkSheetExist("Версии") = False Then
                Запись_Новой_Версии
            End If
        A = MsgBox("Сохранить в архиве старую версию?", vbYesNo)
            If A = vbYes Then
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
    Dim A As Integer
        
        If IsWorkSheetExist("Версии") = False Then
                MsgBox "В этом файле только одна версия", vbCritical
                End
            Else
                ActiveWorkbook.Sheets("Версии").Unprotect 'Снятие защиты листа
                Application.ScreenUpdating = False
                Application.DisplayAlerts = False
                A = MsgBox("Оставить только последнюю версию?", vbYesNo)
                If A = vbYes Then
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
        If IsWorkSheetExist("Версии") = False Then VersionInsert
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

Sub VersionInsert() 'Вставить лист Версии если он отсутсвует
Dim b As Byte
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        Set WbActive = ActiveWorkbook
        b = OpenFolderBook("Template-Spec", "xlsx")
        If b = FileOpenTrue Or b = FileOpenBefore Then WbOpenFile.Sheets("Версии").Copy Before:=WbActive.Sheets(1)
        If b = FileOpenTrue Then WbOpenFile.Close
        WbActive.Sheets("Версии").Unprotect
        WbActive.Sheets("Версии").Rows("2:2").Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
