Attribute VB_Name = "ImportExport"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 19/04/2017г.
Option Explicit
Public Name As String

Private Sub ExportModule() 'Экспортирование всех модулей из книги, на рабочий стол в папку Модули
    Dim objVBProjFrom As Object, objVBProjTo As Object, objVBComp As Object
    Dim sModuleName As String, sFullName As String
    Dim xCount As Byte, i As Integer
    Dim NameItem As String, NameType As Byte
    
    Set objVBProjFrom = ThisWorkbook.VBProject
    xCount = ThisWorkbook.VBProject.VBComponents.Count
    For i = 1 To xCount
        NameItem = ThisWorkbook.VBProject.VBComponents.Item(i).Name
'        Debug.Print NameItem
        NameType = ThisWorkbook.VBProject.VBComponents.Item(i).Type
'        Debug.Print NameType
        Call ExportModFunc(NameItem, NameType)
    Next i
End Sub

Private Sub ImportModule() 'Импортирование Модулей
Dim СписокФайлов As FileDialogSelectedItems
Dim File As Variant, sFile As String, sFullName As String
Dim i As Integer
Dim sRas As String 'расширение файлов
Dim sModuleName As String
Set СписокФайлов = GetFilenamesCollection("Выберите файлы для отправки на печать", CreateObject("WScript.Shell").SpecialFolders("Desktop"))   ' выводим окно выбора
    ' ===================== другие варианты вызова функции =====================
    ' стартовая папка не указана, заголовок окна по умолчанию
'           Set СписокФайлов = GetFilenamesCollection
    ' обзор файлов начинается с папки "Рабочий стол"
'           СтартоваяПапка = CreateObject("WScript.Shell").SpecialFolders("Desktop")
'           Set СписокФайлов = GetFilenamesCollection("Выберите файлы на рабочем столе", СтартоваяПапка)
    ' ==========================================================================
If СписокФайлов Is Nothing Then Exit Sub  ' выход, если пользователь отказался от выбора файлов
i = 1
For Each File In СписокФайлов
sRas = Right(File, 4) 'Определяем расширение файла
Select Case sRas
    Case ".bas", ".cls", ".frm"
        
        sFile = File
        Call ChangeFileCharset(sFile, "Windows-1251", "UTF-8") 'Переводим файл в "Windows-1251" для корректного импорта
        sModuleName = FindTextInString(sFile, 2, "\", ".")
        RemoveModule (sModuleName)
        ThisWorkbook.VBProject.VBComponents.Import filename:=File
        Call ChangeFileCharset(sFile, "UTF-8", "Windows-1251") 'Возвращаем файлу кодировку UTF-8
'    Case ".cls"
'        sModuleName = a.FindTextInString(File, 2, "/", ".")
'        RemoveModule (sModuleName)
'        objVBProjTo.VBComponents.Import Filename:=File
'    Case ".frm"
'        sModuleName = a.FindTextInString(File, 2, "/", ".")
'        RemoveModule (sModuleName)
'        objVBProjTo.VBComponents.Import Filename:=File
    Case Else
End Select
'Debug.Print sRas
'Debug.Print СписокФайлов.Item(i)
i = i + 1
Next
End Sub

Private Function ExportModFunc(sModuleName As String, TypeModule As Byte) 'Функция экспортирования файлов
    Dim objVBProjFrom As Object, objVBProjTo As Object, objVBComp As Object
    Dim sFullName As String
    'расширение стандартного модуля
    Dim sExt As String 'Определяем расширение по типу файла
        Select Case TypeModule
            Case 1
                sExt = ".bas" 'Раширение стандартного модуля
            Case 2
                sExt = ".cls" 'Расширение класса
            Case 3
                sExt = ".frm" 'Расширение формы
            Case Else
                Exit Function 'При других типах выходим из функции
        End Select
    Name = SpecialFolderPath 'Путь рабочего стола
    If Dir(Name & "\Модули", vbDirectory) = "" _
    Then MkDir (Name & "\Модули")  'Создание папки для сохранения

    'имя модуля для копирования
'    sModuleName = "Module1"
    On Error Resume Next
    'проект книги, из которой копируем модуль
    Set objVBProjFrom = ThisWorkbook.VBProject
    'необходимый компонент
    Set objVBComp = objVBProjFrom.VBComponents(sModuleName)
    'если указанного модуля не существует
    If objVBComp Is Nothing Then
        MsgBox "Модуль с именем '" & sModuleName & "' отсутствует в книге.", vbCritical, "Error"
        Exit Function
    End If
    'проект книги для добавления модуля
'    Set objVBProjTo = ActiveWorkbook.VBProject
    'полный путь для экспорта/импорта модуля. К папке должен быть доступ на запись/чтение
    sFullName = Name & "\Модули\" & sModuleName & sExt
    objVBComp.Export filename:=sFullName
    
    'Перекодируем файл в UTF-8
    Call ChangeFileCharset(sFullName, "UTF-8", "Windows-1251")
'    objVBProjTo.VBComponents.Import Filename:=sFullName
    'удаляем временный файл для импорта
'    Kill sFullName
End Function

Function RemoveModule(moduleName As String) 'Удаление модуля
    Dim ModuleDel As Object
    Set ModuleDel = ThisWorkbook.VBProject.VBComponents(moduleName)
    ThisWorkbook.VBProject.VBComponents.Remove ModuleDel
End Function

Private Function FindTextInString(FindindText As String, Left_Mid_Right_1_2_3 As String, _
            First_Symb As String, Optional Second_Symb As String) As String
'Поиск слова в строке
Dim intSymbOne As Integer, intLengText As Integer
Dim intSymbTwo As Integer
'Left_Mid_Right_1_2_3 = LCase(Left_Mid_Right_1_2_3) 'Перевод знаков в нижний регистр
Select Case Left_Mid_Right_1_2_3 ' Выбираем для какого случая делаем
    Case 1
        intSymbOne = InStr(1, FindindText, First_Symb) 'Поиск порядка нужного символа в нашей строке
        FindTextInString = Left(FindindText, intSymbOne - 1) 'Вывод строки
    Case 2
        intLengText = Len(FindindText) 'Определяем длину строки
        intSymbOne = InStrRev(FindindText, First_Symb, intLengText)
        intSymbTwo = InStrRev(FindindText, Second_Symb, intLengText)
        FindTextInString = Mid(FindindText, intSymbOne + 1, intSymbTwo - intSymbOne - 1)
    Case 3
        intLengText = Len(FindindText) 'Определяем длину строки
        intSymbOne = InStrRev(FindindText, First_Symb, intLengText) 'Поиск порядка нужного символа в нашей строке
        FindTextInString = Right(FindindText, intLengText - intSymbOne) 'Вывод строки
    Case Else
        MsgBox "Не верное значение", vbCritical
End Select
End Function

Function GetFilenamesCollection(Optional ByVal title As String = "Выберите файлы для обработки", _
                             Optional ByVal InitialPath As String = "c:\") As FileDialogSelectedItems
    ' функция выводит диалоговое окно выбора нескольких файлов с заголовком Title,
    ' начиная обзор диска с папки InitialPath
    ' возвращает массив путей к выбранным файлам, или пустую строку в случае отказа от выбора
    With Application.FileDialog(3) ' msoFileDialogFilePicker
        .ButtonName = "Выбрать": .title = title: .InitialFileName = InitialPath
        If .Show <> -1 Then Exit Function
        Set GetFilenamesCollection = .SelectedItems
    End With
End Function

Function SpecialFolderPath() As String 'определяет путь рабочего стола
    Dim objWSHShell As Object
    Dim strSpecialFolderPath
    Dim strSpecialFolder

    Set objWSHShell = CreateObject("WScript.Shell")
    SpecialFolderPath = objWSHShell.SpecialFolders("Desktop")
    Set objWSHShell = Nothing
    Exit Function
ErrorHandler:
     MsgBox "Error finding " & strSpecialFolder, vbCritical + vbOKOnly, "Error"
End Function


'====================================================================================================
'Функции перекодирования текста http://excelvba.ru/code/Encode

Sub ПримерИспользования_ChangeTextCharset()
Dim ИсходнаяСтрока As String
    ИсходнаяСтрока = "бНОПНЯ"
    ' вызываем функцию ChangeTextCharset с указанием кодировок
    ' (меняем кодировку с KOI8-R на Windows-1251)
    ПерекодированнаяСтрока = ChangeTextCharset(ИсходнаяСтрока, "Windows-1251", "KOI8-R")
 
    MsgBox "Результат перекодировки: """ & ПерекодированнаяСтрока & """", _
           vbInformation, "Исходная строка: """ & ИсходнаяСтрока & """"
 
End Sub

Function ChangeFileCharset(ByVal filename$, ByVal DestCharset$, _
                           Optional ByVal SourceCharset$) As Boolean
    ' функция перекодировки (смены кодировки) текстового файла
    ' В качестве параметров функция получает путь filename$ к текстовому файлу,
    ' и название кодировки DestCharset$ (в которую будет переведён файл)
    ' Функция возвращает TRUE, если перекодировка прошла успешно
    On Error Resume Next: Err.Clear
    Dim FileContent$
    
    
    With CreateObject("ADODB.Stream")
        .Type = 2
        If Len(SourceCharset$) Then .Charset = SourceCharset$    ' указываем исходную кодировку
        .Open
        .LoadFromFile filename$    ' загружаем данные из файла
        FileContent$ = .ReadText   ' считываем текст файла в переменную FileContent$
        .Close
        .Charset = DestCharset$    ' назначаем новую кодировку
        .Open
        .WriteText FileContent$
        .SaveToFile filename$, 2   ' сохраняем файл уже в новой кодировке
        .Close
    End With
    ChangeFileCharset = Err = 0
End Function

Function ChangeTextCharset(ByVal txt$, ByVal DestCharset$, _
                           Optional ByVal SourceCharset$) As String
    ' функция перекодировки (смены кодировки) текстовоq строки
    ' В качестве параметров функция получает текстовую строку txt$,
    ' и название кодировки DestCharset$ (в которую будет переведён текст)
    ' Функция возвращает текст в новой кодировке
    On Error Resume Next: Err.Clear
    With CreateObject("ADODB.Stream")
        .Type = 2: .Mode = 3
        If Len(SourceCharset$) Then .Charset = SourceCharset$    ' указываем исходную кодировку
        .Open
        .WriteText txt$
        .Position = 0
        .Charset = DestCharset$    ' назначаем новую кодировку
        ChangeTextCharset = .ReadText
        .Close
    End With
End Function

Function ChangeFileCharset_UTF8noBOM(ByVal filename$, Optional ByVal SourceCharset$) As Boolean
    ' функция перекодировки (смены кодировки) текстового файла
    ' В качестве параметров функция получает путь filename$ к текстовому файлу,
    ' Функция возвращает TRUE, если перекодировка прошла успешно
    On Error Resume Next: Err.Clear
    DestCharset$ = "utf-8"
    With CreateObject("ADODB.Stream")
        .Type = 2
        If Len(SourceCharset$) Then .Charset = SourceCharset$        ' указываем исходную кодировку
        .Open
        .LoadFromFile filename$        ' загружаем данные из файла
        FileContent$ = .ReadText        ' считываем текст файла в переменную FileContent$
        .Close
        .Charset = DestCharset$        ' назначаем новую кодировку "utf-8"
        .Open
        .WriteText FileContent$
 
        'Write your data into the stream.

        Dim binaryStream As Object
        Set binaryStream = CreateObject("ADODB.Stream")
        binaryStream.Type = 1
        binaryStream.Mode = 3
        binaryStream.Open
        'Skip BOM bytes
        .Position = 3
        .CopyTo binaryStream
        .flush
        .Close
        binaryStream.SaveToFile filename$, 2
        binaryStream.Close
    End With
    ChangeFileCharset_UTF8noBOM = Err = 0
End Function

Function EncodeUTF8noBOM(ByVal txt As String) As String
    For i = 1 To Len(txt)
        l = Mid(txt, i, 1)
        Select Case AscW(l)
            Case Is > 4095: t = Chr(AscW(l) \ 64 \ 64 + 224) & Chr(AscW(l) \ 64) & Chr(8 * 16 + AscW(l) Mod 64)
            Case Is > 127: t = Chr(AscW(l) \ 64 + 192) & Chr(8 * 16 + AscW(l) Mod 64)
            Case Else: t = l
        End Select
        EncodeUTF8noBOM = EncodeUTF8noBOM & t
    Next
End Function

Function SaveTextToFile(ByVal txt$, ByVal filename$, Optional ByVal encoding$ = "windows-1251") As Boolean
    ' функция сохраняет текст txt в кодировке Charset$ в файл filename$
    On Error Resume Next: Err.Clear
    Select Case encoding$
 
        Case "windows-1251", "", "ansi"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(filename, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing
 
        Case "utf-16", "utf-16LE"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(filename, True, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing
 
        Case "utf-8noBOM"
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = "utf-8": .Open
                .WriteText txt$
 
                Set binaryStream = CreateObject("ADODB.Stream")
                binaryStream.Type = 1: binaryStream.Mode = 3: binaryStream.Open
                .Position = 3: .CopyTo binaryStream        'Skip BOM bytes
                .flush: .Close
                binaryStream.SaveToFile filename$, 2
                binaryStream.Close
            End With
 
        Case Else
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = encoding$: .Open
                .WriteText txt$
                .SaveToFile filename$, 2        ' сохраняем файл в заданной кодировке
                .Close
            End With
    End Select
    SaveTextToFile = Err = 0: DoEvents
End Function

