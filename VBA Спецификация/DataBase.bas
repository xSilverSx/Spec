Attribute VB_Name = "DataBase"
Option Explicit

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
ThisWorkbook.Saved = True 'Исключение сохранения запроса сохранения надстройки
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub КопированиеЛиста()
    Редактор_Книги
    Workbooks("База данных.xlsx").Sheets("База_СО").Copy Before:=Workbooks("Спецификация Надстройка.xlam").Sheets("Шаблоны")
    ThisWorkbook.Saved = True 'Исключение сохранения запроса сохранения надстройки
    Редактор_Книги
End Sub

