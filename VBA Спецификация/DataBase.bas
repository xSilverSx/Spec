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

Sub Подключить_Базу_Данных()            'Подключить базу данных в Надстройку
Dim Name As String                      'Путь расположения надстройки
Dim strPath As String                   'Путь к базе данных
Dim bOpen As Boolean                    'Запоминает была ли открыта книга базы данных
Dim Wb As Workbook
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    Name = ThisWorkbook.Path                'Путь расположения надстройки
    strPath = Name & "\База данных.xlsx"    'Определяем полный путь файла
    If FileLocation(strPath) = True Then    'Проверяем есть ли файл базы данных
        If IsBookOpen(strPath) = False Then 'Проверяем открыт ли этот файл
            bOpen = False
            Set Wb = Workbooks.Open(strPath)
            КопированиеЛиста
            Wb.Close
        End If
    Else
        MsgBox ("Файл базы не найден")
    End If
ThisWorkbook.Saved = True 'Исключение сохранения запроса на сохранение надстройки
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub КопированиеЛиста()
    If ThisWorkbook.IsAddin = True Then Редактор_Книги
    Workbooks("База данных.xlsx").Sheets("База_СО").Copy Before:=Workbooks("Спецификация Надстройка.xlam").Sheets("Шаблоны")
    ThisWorkbook.Saved = True 'Исключение сохранения запроса сохранения надстройки
    Редактор_Книги
End Sub

Sub УдалитьБазуДанных()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    If ThisWorkbook.IsAddin = True Then Редактор_Книги
        If IsWorkSheetExistXLAM("База_СО") = True Then
            ThisWorkbook.Sheets("База_СО").Delete
'            MsgBox "Лист База_СО удален"
'        Else
'            MsgBox "Лист База_СО не найден"
        End If
    Редактор_Книги
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub ПереподключитьБазуДанных()
    УдалитьБазуДанных
    Подключить_Базу_Данных
    MsgBox "База данных успешно подключена!"
End Sub

Sub ОткрытьБазуДанных()
Dim Name As String
Dim strPath As String
Name = ThisWorkbook.Path                    'Путь расположения надстройки
strPath = Name & "\База данных.xlsx"        'Определяем полный путь файла
If FileLocation(strPath) = True Then        'Проверяем есть ли файл базы данных
        If IsBookOpen("База данных.xlsx") = False Then 'Проверяем открыт ли этот файл
            Workbooks.Open (strPath)
        Else
        MsgBox "База данных уже открыта"
            Workbooks("База данных.xlsx").Sheets("База_СО").Activate
        End If
Else
    MsgBox ("Файл базы не найден")
End If
End Sub

Sub Сортировка_Базы()
If IsBookOpen("База данных.xlsx") = True Then
    Workbooks("База данных.xlsx").Sheets("База_СО").Activate
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

