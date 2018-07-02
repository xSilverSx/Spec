VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PaketnayaObrabotka 
   Caption         =   "Пакетная обработка"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3075
   OleObjectBlob   =   "PaketnayaObrabotka.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PaketnayaObrabotka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Public Vopros As Boolean
Public This_Wbk As Workbook
Public First_Wbk As Workbook 'для записи начальной книги
Public strAdress As String 'адресс ячеек для копирования


Private Sub CommandButton1_Click()
Dim i As Byte, bQuest As Byte
Dim СписокФайлов As FileDialogSelectedItems
Dim a
Dim Pos As Integer
Dim strPos As String, BoolStrAdress As Boolean
Dim OpenBook As Boolean
Dim File As Variant
Dim CountFile As Long, Progress As Integer
Dim bar As Progressbar
Application.DisplayAlerts = False
Set First_Wbk = ActiveWorkbook
Me.Hide

    If OButton4.Value = True Or OButton5.Value = True Then strAdress = Какие_Ячейки_Выбрать("Задайте диапазон для копирования", "Диапазон")

Set СписокФайлов = GetFilenamesCollection("Выберите файлы для отправки на печать", ActiveWorkbook.Path)   ' выводим окно выбора
    ' ===================== другие варианты вызова функции =====================
    ' стартовая папка не указана, заголовок окна по умолчанию
'           Set СписокФайлов = GetFilenamesCollection
    ' обзор файлов начинается с папки "Рабочий стол"
'           СтартоваяПапка = CreateObject("WScript.Shell").SpecialFolders("Desktop")
'           Set СписокФайлов = GetFilenamesCollection("Выберите файлы на рабочем столе", СтартоваяПапка)
    ' ==========================================================================
                                          
                                          
If СписокФайлов Is Nothing Then Exit Sub  ' выход, если пользователь отказался от выбора файлов
    
CountFile = СписокФайлов.Count
'If CountFile > 9 Then
    Set bar = New Progressbar
    bar.createtimeFinish    ' вывод строки для оставшегося времени
    bar.createLoadingBar    ' вывод полосы загрузки
    bar.createString    ' вывод строки пройденных этапов из общего количества с указанием процента
    bar.createtimeDuration  ' текущая время обработки процесса
    bar.createTextBox   ' вывод пустого текстового поля
    bar.setParameters CountFile, 0, 1   ' Задание параметров для последующей обработки:
                                        ' 1 - указание числа этапов процесса;
                                        ' 2 - интервал обновления формы, в данном случае ноль, но можно вовсе опустить
                                        ' 3 - интервал обновления в секундах, применяется, только если предыдущий _
                                              аргумент равен нулю или опущен
    bar.Start "Идет пакетная обработка"
'End If
    
    Application.ScreenUpdating = False
        For Each File In СписокФайлов
'  Debug.Print File
Pos = InStrRev(File, "\") 'Определяем имя файла (без пути)
'    Debug.Print Pos
strPos = Mid(File, Pos + 1)
'    Debug.Print strPos
    If IsBookOpen(strPos) Then
'        MsgBox "Книга открыта", vbInformation, "Сообщение"
        Workbooks(strPos).Activate
        OpenBook = True
    Else
'        MsgBox "Книга закрыта", vbInformation, "Сообщение"
        OpenBook = False
        Workbooks.Open filename:=File
    End If
        Set This_Wbk = ActiveWorkbook
                
                        If OButton1.Value = True Then 'Обновить версии спецификации
                            Копировать_Листы
                        ElseIf OButton2.Value = True Then
                            Создать_кнопки
                        ElseIf OButton3.Value = True Then
                            Обновить_дату_последней_версии
                        ElseIf OButton4.Value = True Then
                                Активация_И_Копия_Фрагмента
                                This_Wbk.Activate
                                Замена_Основной_Надписи
                        ElseIf OButton5.Value = True Then
                            ActiveWorkbook.Sheets("СО").Activate
                            Range(strAdress).Copy
                            ActiveWorkbook.Sheets("ВР").Activate
                            Range(strAdress).Activate
                            ActiveSheet.Paste
                        End If
                            ActiveWorkbook.Save
                If OpenBook = False Then 'Закрываем если книга не была открыта
                    This_Wbk.Close False
                End If
    '                Set This_Wbk = ActiveWorkbook
    '                Set ActSheet = ActiveSheet
'            If CountFile > 9 Then
                Progress = Progress + 1
                bar.Update Progress, "Идет обработка - " & strPos 'Progress + 1
'            End If
'            Debug.Print File
        Next
    bar.exitBar ' Закрываем прогресс бар
    Set bar = Nothing ' удаляем экземпляр класса прогресс бара
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Me.Show
End Sub

Sub Активация_И_Копия_Фрагмента()
    First_Wbk.Activate
    ActiveWorkbook.Sheets("СО").Activate
    Range(strAdress).Copy
End Sub

Sub Замена_Основной_Надписи()
    ActiveWorkbook.Sheets("ВР").Activate
    Range(strAdress).Activate
        Draws_In_Selection_Select
    ActiveSheet.Paste
    ActiveWorkbook.Sheets("СО").Activate
    Range(strAdress).Activate
        Draws_In_Selection_Select
    ActiveSheet.Paste
End Sub

