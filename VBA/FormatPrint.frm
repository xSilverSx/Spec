VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormatPrint 
   Caption         =   "Отправить на печать или создать PDF"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   OleObjectBlob   =   "FormatPrint.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormatPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Public Vopros As Boolean 'Действительно ли отпрвить на печать текущий лист
Public This_Wbk As Workbook


Private Sub PechatNeskolco_Click()
    Dim i As Byte
    Dim СписокФайлов As FileDialogSelectedItems
    Dim a
    Dim Pos As Integer
    Dim strPos As String
    Dim OpenBook As Boolean
    Dim File As Variant
    Dim CountFile As Long, Progress As Integer
    Dim bar As Progressbar

Application.DisplayAlerts = False
        If CheckBoxSO.Value = False And CheckBoxVR.Value = False Then
            MsgBox ("Не выбрано ни одного листа для печати")
            Exit Sub
        End If
            If PrintA3 Or PrintA4 = True Then
                i = MsgBox("Действительно отправить спецификации на печать?", vbYesNo)
                    If i = vbNo Then: Exit Sub
            End If
    Vopros = True
        
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
    bar.Start "Идет печать"
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
                
                        If CheckBoxSO.Value = True Then 'Отправляем на печать или в пдф нужные файлы
                                Worksheets("СО").Activate
                                PrintFormat_Click
                        End If
                        If CheckBoxVR.Value = True Then
                                Worksheets("ВР").Activate
                                PrintFormat_Click
                        End If
                If OpenBook = False Then 'Закрываем если книга не была открыта
                    This_Wbk.Close False
                End If
    '                Set This_Wbk = ActiveWorkbook
    '                Set ActSheet = ActiveSheet
'            If CountFile > 9 Then
                Progress = Progress + 1
                bar.Update Progress, "Печать книги - " & strPos 'Progress + 1
'            End If
'            Debug.Print File
        Next
    bar.exitBar ' Закрываем прогресс бар
    Set bar = Nothing ' удаляем экземпляр класса прогресс бара
    Vopros = False
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Private Sub PrintFormat_Click()
Dim a
Dim i As Byte
Select Case a
    Case PrintPDF.Value = False
'            If Vopros = False Then
'                i = MsgBox("Вы уверены?", vbYesNo)
'                    If i = vbNo Then: Exit Sub
'            End If
        Создать_PDF
    Case PrintA4.Value = False
            If Vopros = False Then
                i = MsgBox("Отправить на печать текущий лист? Формат листа А4", vbYesNo)
                    If i = vbNo Then: Exit Sub
            End If
        Печать_на_А4
        Отправить_на_печать
    Case PrintA3.Value = False
            If Vopros = False Then
                i = MsgBox("Отправить на печать текущий лист? Формат листа А3", vbYesNo)
                    If i = vbNo Then: Exit Sub
            End If
        Печать_на_А3
        Отправить_на_печать
    End Select
FormatPrint.Hide
End Sub






