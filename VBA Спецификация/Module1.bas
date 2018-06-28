Attribute VB_Name = "Module1"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit
Public TheClosedBook As Boolean

Sub Main()
    ActiveWorkbook.Worksheets("Спецификация").Activate
        If IsWorkSheetExist("База_СО") = False Then 'В случае отсутствия базы данных вывести сообщение
            MsgBox "База данных не установлена!", vbCritical
        Else
            SO_Zapolnen.Show
        End If
End Sub
Sub addInTheBase()
    ActiveWorkbook.Worksheets("Спецификация").Activate
    addBase.Show
End Sub
Sub Выгрузить_Форму()
    TheClosedBook = True
    Unload VBAProjectSO.SO_Zapolnen
    Unload VBAProjectSO.addBase
End Sub
Sub Редактор_Книги() 'Делает книгу доступной для редактрования
    If ThisWorkbook.IsAddin = False Then
    ThisWorkbook.IsAddin = True
    Exit Sub
    End If
    If ThisWorkbook.IsAddin = True Then ThisWorkbook.IsAddin = False
End Sub
Sub Сохранить_Книгу() 'Сохраняет книгу
Dim a As Byte
a = MsgBox("Действительно пересохранить файл Надстройки?", vbYesNo)
If a = vbYes Then ThisWorkbook.Save
End Sub
Sub Сортировка_Базы()
Dim m As Boolean    'Логическое для проверки видимости листа
Application.ScreenUpdating = False
If ThisWorkbook.IsAddin = True Then
m = True
Редактор_Книги
End If

Call Заменить("Раздел", "ЯЯЯЯЯЯЯРаздел", True, Range("A:A")) 'Замена для того чтобы "Раздел" был в конце

    ThisWorkbook.Worksheets("База_СО").ListObjects("Таблица1").Sort.SortFields. _
        Clear
    ThisWorkbook.Worksheets("База_СО").ListObjects("Таблица1").Sort.SortFields. _
        Add Key:=Range("Таблица1[Категория]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("База_СО").ListObjects("Таблица1").Sort.SortFields. _
        Add Key:=Range("Таблица1[Подкатегория]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("База_СО").ListObjects("Таблица1").Sort.SortFields. _
        Add Key:=Range("Таблица1[Краткое Наименование]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("База_СО").ListObjects("Таблица1").Sort.SortFields. _
        Add Key:=Range("Таблица1[Сортировка]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("База_СО").ListObjects("Таблица1").Sort.SortFields. _
        Add Key:=Range("Таблица1[[Тип ]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ThisWorkbook.Worksheets("База_СО").ListObjects("Таблица1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Call Заменить("ЯЯЯЯЯЯЯРаздел", "Раздел", True, Range("A:A"))

If m Then Редактор_Книги
Application.ScreenUpdating = True
MsgBox "Сортировка Базы данных проведена"
End Sub
Sub Change_ReferenceStyle() 'Замена стилей R1C1
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub
