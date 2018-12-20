Attribute VB_Name = "PrintAndOther"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit
Public BaseName As String, Name As String
Public inPos As Integer
Public PathBook As String, strFileN As String

Sub Создать_PDF()
    Печать_на_А3
    Name = SpecialFolderPath 'Путь рабочего стола
    If Dir(Name & "\PDF Спецификации", vbDirectory) = "" _
    Then MkDir (Name & "\PDF Спецификации")  'Создание папки для сохранения
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        Name & "\PDF Спецификации\" & Название_документа & "-" & Название_листа & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
End Sub

Sub Создать_PDF_СО_ВР()
Application.ScreenUpdating = False
    Worksheets("СО").Activate
    Создать_PDF
    Worksheets("ВР").Activate
    Создать_PDF
    Worksheets("Спецификация").Activate
Application.ScreenUpdating = True
End Sub

Function Название_документа() As String
    Название_документа = ActiveWorkbook.Name
End Function

Function Название_листа() As String
    Название_листа = ActiveSheet.Name
End Function

Sub Печать_на_А4()
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        .Zoom = 70
        .LeftMargin = Application.InchesToPoints(0.196850393700787)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.196850393700787)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0.196850393700787)
        .FooterMargin = Application.InchesToPoints(0.24)
    End With
End Sub

Sub Печать_на_А3()
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA3
        .Zoom = 100
        .LeftMargin = Application.InchesToPoints(0.196850393700787)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.196850393700787)
        .BottomMargin = Application.InchesToPoints(0.196850393700787)
        .HeaderMargin = Application.InchesToPoints(0.196850393700787)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
    End With
End Sub

Sub Отправить_на_печать()
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub

Sub Print_Show()
    FormatPrint.Show
End Sub

Sub Changing_Show()
    Changing.Show
End Sub



Sub Копирование_КнопокСО()
    Columns("AS:AZ").Delete
        ActiveSheet.Buttons.Add(1170, 0, 115, 30).Select
            Selection.OnAction = "Добавить_Лист"
            Selection.Characters.Text = "Добавить Лист"
            Selection.Name = "Button 1"
        ActiveSheet.Buttons.Add(1170, 55, 115, 30).Select
            Selection.OnAction = "Print_Show"
            Selection.Characters.Text = "Отправить на печать"
            Selection.Name = "Button 2"
    Range("e2").Activate
End Sub

Sub Создать_кнопки()
Dim A As Byte
Application.ScreenUpdating = False
If ListSpec = False Then
    Exit Sub
Else
    AllObjButtonDelete 'Удаление кнопок на листах
        Sheets("Спецификация").Activate 'Создание кнопок Спецификация
            ActiveSheet.Buttons.Add(50, 0, 50, 20).Select
                Selection.OnAction = "Podgotovka_Show"
                Selection.Characters.Text = "Перенос"
                Selection.Name = "Button 1"

            ActiveSheet.Buttons.Add(100, 0, 100, 20).Select
                Selection.OnAction = "main"
                Selection.Characters.Text = "Добавить из базы"
                Selection.Name = "Button 2"

            ActiveSheet.Buttons.Add(335, 0, 60, 20).Select
                Selection.OnAction = "Очистить_всё"
                Selection.Characters.Text = "Очистка"
                Selection.Name = "Button 3"

            ActiveSheet.Buttons.Add(938, 0, 105, 20).Select
                Selection.OnAction = "Создать_PDF_СО_ВР"
                Selection.Characters.Text = "Создать PDF СО и ВР"
                Selection.Name = "Button 4"
        
        
        Sheets("Перенос").Activate 'Создание кнопок Перенос
            ActiveSheet.Buttons.Add(55, 0, 125, 20).Select
                Selection.OnAction = "Перенос_по_строкам"
                Selection.Characters.Text = "Перенос по строкам"
                Selection.Name = "Button 1"
                
            ActiveSheet.Buttons.Add(350, 0, 125, 20).Select
                Selection.OnAction = "На_печать_выборочно"
                Selection.Characters.Text = "Перенос по листам"
                Selection.Name = "Button 2"
                
        Worksheets("СО").Activate
            Копирование_КнопокСО
        Worksheets("ВР").Activate
            Копирование_КнопокСО
Sheets("Спецификация").Activate
If ActiveWorkbook.Sheets("ВР").Range("A15").Value <> "" And ActiveWorkbook.Sheets("ВР").Range("A15").Value <> "" Then
    A = MsgBox("Обновить до ГОСТ Р 21.1101-2013?", vbYesNo)
    If A = vbYes Then AddWorkbookFooter
End If
Application.ScreenUpdating = True
End If
End Sub
 
 Sub Удалить_лист_Шаблон()
    If IsWorkSheetExist("Шаблоны") = True Then
            Sheets("Шаблоны").Delete
    End If
 End Sub
 
Sub Сохранить_XLSX()
Dim A As Boolean, b As Boolean
Dim Full As String, NewFullName As String
Dim i As Integer

PathBook = ActiveWorkbook.Path
'MsgBox (a)   'полный путь к папке файла
BaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(Название_документа)

strFileN = PathBook & "\" & BaseName & ".xlsx"
Full = PathBook & "\" & BaseName & ".xlsm"
Do While FileLocation(strFileN)
    i = i + 1
    strFileN = PathBook & "\" & BaseName & "-" & i & ".xlsx"
Loop

    ActiveWorkbook.SaveAs filename:=strFileN, FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False
A = FileLocation(strFileN)
b = FileLocation(Full)
If A And b Then
        i = 0
        Name = SpecialFolderPath 'Путь рабочего стола
        
        If Dir(Name & "\PDF Спецификации", vbDirectory) = "" _
        Then MkDir (Name & "\PDF Спецификации")  'Создание папки для сохранения
        
        If Dir(Name & "\PDF Спецификации\XLSM", vbDirectory) = "" _
        Then MkDir (Name & "\PDF Спецификации\XLSM")  'Создание папки для сохранения
              
        NewFullName = Name & "\PDF Спецификации\XLSM\" & BaseName & ".xlsm"
'        Debug.Print NewFullName
            Do While FileLocation(NewFullName)
                i = i + 1
                NewFullName = Name & "\PDF Спецификации\XLSM\" & BaseName & "-" & i & ".xlsm"
            Loop
    Name Full As NewFullName
    NewFullName = Replace(NewFullName, ".xlsm", ".txt")
'    Debug.Print NewFullName
    Запомнить NewFullName, Full
End If
End Sub

'Sub Свойства_Файла()
'ActiveWorkbook.BuiltinDocumentProperties(32).Value = "Версия 17" 'Простановка версии файла в графе "Состояние содержимого"
'End Sub
 
Function NameNoDate() As String
BaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(Название_документа)
inPos = InStrRev(BaseName, " 20")
    If inPos = 0 Then
        NameNoDate = BaseName
            Else
        NameNoDate = Left(BaseName, inPos - 1)
    End If
End Function
Sub Сохранить_Сегодня()
Dim strDate As String, strNameNoDate As String
Dim i As Byte
PathBook = ActiveWorkbook.Path
strNameNoDate = NameNoDate
strDate = Format(Now(), "yyyy.mm.dd")
strFileN = PathBook & "\" & strNameNoDate & " " & strDate & ".xlsx"
Do While FileLocation(strFileN)
    i = i + 1
    strFileN = PathBook & "\" & strNameNoDate & " " & strDate & "-" & i & ".xlsx"
Loop
ActiveWorkbook.SaveAs filename:=strFileN, FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False
MsgBox ("Файл сохранен:" & vbCr & strFileN)
End Sub

Private Sub Запомнить(ИмяФайла As String, СтарыйПуть As String) 'Создает текстовый документ, с тем же именем что и перемещаемый файл и сохраняет там бывший путь
Dim FS As Object
Set FS = CreateObject("Scripting.FileSystemObject")
Set A = FS.CreateTextFile(ИмяФайла, True)
A.WriteLine ("Старый путь нахождения файла" & vbNewLine & СтарыйПуть)
A.Close
End Sub

