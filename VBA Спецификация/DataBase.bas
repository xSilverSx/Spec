Attribute VB_Name = "DataBase"
Option Explicit
Public TheClosedBook As Boolean

Sub Main() 'Функция запуска формы для вставки позиций из базы
    ActiveWorkbook.Worksheets("Спецификация").Activate
        If IsWorkSheetExistXLAM("База_СО") = False Then
            MsgBox "База данных не установлена!", vbCritical
        Else
            SO_Zapolnen.Show
        End If
End Sub

Sub Выгрузить_Форму() 'Функция выгрузки формы из памяти
    TheClosedBook = True
    Unload VBAProjectSO.SO_Zapolnen
    Unload VBAProjectSO.addBase
End Sub

Sub Подключить_Базу_Данных() 'Подключить базу данных в Надстройку Если база отсутсвует, переподключить если база была открыта до этого
Dim b As Byte
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    ThisWorkbook.IsAddin = False
	Выгрузить_Форму 'Закрытие формы перед подключением базы
    Set WbActive = ThisWorkbook
    b = OpenFolderBook("SpecDataBase", "xlsx") 'Открыть файл базы данных
    If b = FileOpenTrue Or b = FileOpenBefore Then _
    WbOpenFile.Sheets("База_СО").Copy Before:=WbActive.Sheets(1)
    If b = FileOpenTrue Then WbOpenFile.Close
    If IsWorkSheetExistXLAM("База_СО (2)") Then     'При удаление листа "До копирования" возникает ошибка, поэтому лист удаляем после копирования с переименованием
        WbActive.Sheets("База_СО").Name = "Удалить"
        WbActive.Sheets("Удалить").Delete
        WbActive.Sheets("База_СО (2)").Name = "База_СО"
        MsgBox "База данных переподключена"
    End If
    bOpen = False
    WbActive.IsAddin = True
    WbActive.Saved = True 'Исключение сохранения запроса на сохранение надстройки
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub УдалитьБазуДанных()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    If ThisWorkbook.IsAddin = True Then ThisWorkbook.IsAddin = False
        If IsWorkSheetExistXLAM("База_СО") = True Then
            ThisWorkbook.Sheets("База_СО").Delete
        End If
    ThisWorkbook.IsAddin = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub Сортировка_Базы()
If IsBookOpen("SpecDataBase.xlsx") = True Then
    Workbooks("SpecDataBase.xlsx").Sheets("База_СО").Activate
    Call Заменить("Раздел", "ЯЯЯЯЯЯЯРаздел", True, Range("A:A")) 'Замена для того чтобы "Раздел" при сортировке оказался в конце
        Worksheets("База_СО").ListObjects("Таблица").Sort.SortFields. _
            Clear
        Worksheets("База_СО").ListObjects("Таблица").Sort.SortFields. _
            Add Key:=Range("Таблица[Категория]"), SortOn:=xlSortOnValues, Order:= _
            xlAscending, DataOption:=xlSortNormal
        Worksheets("База_СО").ListObjects("Таблица").Sort.SortFields. _
            Add Key:=Range("Таблица[Подкатегория]"), SortOn:=xlSortOnValues, Order:= _
            xlAscending, DataOption:=xlSortNormal
        Worksheets("База_СО").ListObjects("Таблица").Sort.SortFields. _
            Add Key:=Range("Таблица[Краткое Наименование]"), SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        Worksheets("База_СО").ListObjects("Таблица").Sort.SortFields. _
            Add Key:=Range("Таблица[Сортировка]"), SortOn:=xlSortOnValues, Order:= _
            xlAscending, DataOption:=xlSortNormal
        Worksheets("База_СО").ListObjects("Таблица").Sort.SortFields. _
            Add Key:=Range("Таблица[[Тип ]]"), SortOn:=xlSortOnValues, Order:= _
            xlAscending, DataOption:=xlSortNormal
        With Worksheets("База_СО").ListObjects("Таблица").Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Call Заменить("ЯЯЯЯЯЯЯРаздел", "Раздел", True, Range("A:A"))
    MsgBox "Сортировка Базы данных проведена"
Else
    MsgBox "База данных закрыта. Нелья провести сортировку", vbCritical
End If
End Sub

