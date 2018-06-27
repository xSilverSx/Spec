VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public bClose As Byte
Private WithEvents xlsApp As Application
Attribute xlsApp.VB_VarHelpID = -1
Public iQuestion As Integer
Public wbOpen As Workbook
Public bOpen As Boolean

Private Sub Workbook_Open()
    Set xlsApp = ThisWorkbook.Application
    AutoOpen
    'Обнуляем значения
    iQuestion = 0
    bClose = 0
End Sub

Sub Oshibka_na_liste_Perenos()
    Set xlsApp = ThisWorkbook.Application
End Sub
 
Private Sub xlsApp_SheetChange(ByVal Sh As Object, ByVal Target As Range) 'Разбивка по строкам на листе Пернос
On Error Resume Next   'Иногда возникает ошибка на строке If Selection.Count > 1 Then Exit Sub при обработке закоментировать
If (TypeOf Sh Is Worksheet) And (Sh.Name = "Перенос") Then
If Selection.Count > 1 Then Exit Sub
Dim CW As Integer, CS As Integer, Pr As Integer
'Application.EnableEvents = 0 '*1
If IsError(Target) Then
'    Dim sNameList As String, sNameCell As String
    MsgBox "Ошибка в одной из ячеек" & Chr(13) & _
        "Проверьте ячейку - " & Target.Address & ", на листе - " & Target.Worksheet.Name, vbCritical
    End
End If
   
    With Target
        CS = Len(.Value) 'длина значения текста
        CW = Int(.EntireColumn.ColumnWidth) 'ширина столбца
        CW = InStrRev(.Value, " ", CW) 'Позиция пробела
        Pr = InStr(.Value, Chr(10)) 'Позиция символа переноса
        If Pr <> 0 Then If (Pr < CW) Or (CW = 0) Then CW = Pr
        'сравнение значений положения символа переноса
        If CW < 1 Then Exit Sub
        If CS > CW Then
            .Offset(1, 0).Value = Mid(.Value, CW + 1, CS - CW)
            .Value = Left(.Value, CW - 1)
        End If
    '    .EntireRow.Insert
    End With
'Application.EnableEvents = 3 '*1
'-----------------------------------------------------
''*1 - стереть строки если кромсать надо несколько раз
End If
End Sub

Private Sub AutoOpen()
    Dim bQuest As Byte
    Dim BaseName As String
Application.ScreenUpdating = False
            If bClose = 1 Then
                bClose = 0
                Application.ScreenUpdating = True
                End
            End If
'Подключить_Базу_Данных
'Иногда при открытии файла эксель не отображаются таблицы, такой код немного исправляет ситуацию
            If Workbooks.Count = 0 Then
                bOpen = True
                Set wbOpen = Application.Workbooks.Add
            End If
            
            AddIns.Add filename:=ThisWorkbook.FullName 'Включение надстройки при открытии файла
            BaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(ThisWorkbook.FullName)
            If AddIns(BaseName).Installed = False Then
                bQuest = MsgBox("Установить Надстройку?", vbYesNo)
                    bClose = 1
                    If bQuest = vbYes Then AddIns(BaseName).Installed = True
            End If
Application.ScreenUpdating = True
End Sub

'Function Открытие_файла() As Boolean
'Dim st As String 'Переменная пути рабочего стола
'Dim Q As String 'Переменная вопроса
'Const FILEO = "\Open list.txt" 'Постоянная для пути файла с настройками
'Dim sFile As String 'Расположение файла с настройками
'        st = ThisWorkbook.Path
'        sFile = st & FILEO
'        If FileLocation(sFile) = True Then 'Проверка существования файла
'            Call ChangeFileCharset(sFile, "Windows-1251", "UTF-8")
'            Open st & FILEO For Input As #1
'                Input #1, Q
'            If Q = "Да" Then Открытие_файла = True
'                Close #1
'            Call ChangeFileCharset(sFile, "UTF-8", "Windows-1251")
'        Else: Открытие_файла = False
'        End If
'End Function

Sub Подключить_Базу_Данных()
Dim Name As String 'Путь расположения надстройки
Dim strPath As String 'Путь к базе данных
Dim bOpen As Boolean 'Запоминает была ли открыта книга базы данных
Dim Wb As Workbook
Name = ThisWorkbook.Path 'Путь расположения надстройки
strPath = Name & "\База данных.xlsx" 'Определяем полный путь файла
If FileLocation(strPath) = True Then 'Проверяем есть ли файл базы данных
    If IsBookOpen(strPath) = False Then 'Проверяем открыт ли этот файл
        bOpen = False
        Set Wb = Workbooks.Open(strPath)
        Wb.Close
    End If
Else
    MsgBox ("Файл базы не найден")
End If
End Sub

Private Sub xlsApp_WorkbookDeactivate(ByVal Wb As Workbook)
    If bOpen = True Then
        bOpen = False
        wbOpen.Close 'Закрыть книгу после исправления
    End If
End Sub
