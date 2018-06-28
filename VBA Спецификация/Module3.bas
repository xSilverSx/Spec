Attribute VB_Name = "Module3"
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
    End With
End Sub

Sub Печать_на_А3()
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA3
        .Zoom = 100
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

Sub Копировать_Листы() 'Обновление спецификации, и сохранение ее как xlsx
Dim rngAll As Range
Dim lnRow As Integer
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Спецификация").Copy After:=ActiveWorkbook.Sheets(1) 'Копировать листы из шаблона
    ThisWorkbook.Sheets("Перенос").Copy After:=ActiveWorkbook.Sheets(2)

        Worksheets("СО").Activate
            Копирование_КнопокСО
        Worksheets("ВР").Activate
            Копирование_КнопокСО
    
    Sheets("Спецификация").Activate  'Выделение необходимой области на листе спецификация
    Set rngAll = Range("A1").CurrentRegion
    lnRow = rngAll.Rows.Count 'Подсчет числа колонок
    Range(Cells(2, 1), Cells(lnRow, 25)).Copy 'Копирование области
    
    Sheets("Спецификация (2)").Activate 'выбор листа на который произвести копирование
    Range("a2").Activate
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("a2").Select
    
    Sheets("Перенос").Activate  'Выделение необходимой области на листе спецификация
    Range(Cells(2, 1), Cells(1000, 15)).Copy
    Sheets("Перенос (2)").Activate 'выбор листа на который произвести копирование
    Range("a2").Activate
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("a2").Select
    
    Удалить_лист_Шаблон
'    Application.DisplayAlerts = False
        Sheets("Спецификация").Delete
        Sheets("Перенос").Delete
'    Application.DisplayAlerts = True
    Sheets("Спецификация (2)").Name = "Спецификация"
    Создать_кнопки
    Sheets("Спецификация").Activate
'    If Range("j2").Value = "" Then
'    Range("j2:K2").Value = "9"
'    End If
    Sheets("Перенос (2)").Name = "Перенос"
'   Свойства_Файла
    
'    ThisWorkbook.Sheets("СО").Copy After:=ActiveWorkbook.Sheets(3)
'    ThisWorkbook.Sheets("ВР").Copy After:=ActiveWorkbook.Sheets(4)
Сохранить_XLSX
MsgBox ("Файл Обновлен")
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub Копирование_КнопокСО()
    
    Columns("AS:AZ").Delete
        ActiveSheet.Buttons.Add(1170, 0, 115, 30).Select
            Selection.OnAction = "VBAProjectSO.Module2.Добавить_Лист"
            Selection.Characters.Text = "Добавить Лист"
        ActiveSheet.Buttons.Add(1170, 55, 115, 30).Select
            Selection.OnAction = "VBAProjectSO.Module3.Print_Show"
            Selection.Characters.Text = "Отправить на печать"
    Range("e2").Activate
  
End Sub

Sub Создать_кнопки()
Dim a As Byte
Application.ScreenUpdating = False
If List = False Then
    Exit Sub
Else
        
        
        Sheets("Спецификация").Activate 'Создание кнопок Спецификация
        Удалить_Объекты
            ActiveSheet.Buttons.Add(50, 0, 50, 20).Select
                Selection.OnAction = "VBAProjectSO.Module2.Podgotovka_Show"
                Selection.Characters.Text = "Печать"

            ActiveSheet.Buttons.Add(100, 0, 100, 20).Select
                Selection.OnAction = "VBAProjectSO.Module1.main"
                Selection.Characters.Text = "Добавить из базы"

            ActiveSheet.Buttons.Add(335, 0, 60, 20).Select
                Selection.OnAction = "VBAProjectSO.Module2.Очистить_всё"
                Selection.Characters.Text = "Очистка"

            ActiveSheet.Buttons.Add(938, 0, 105, 20).Select
                Selection.OnAction = "VBAProjectSO.Module3.Создать_PDF_СО_ВР"
                Selection.Characters.Text = "Создать PDF СО и ВР"
        
        
        Sheets("Перенос").Activate 'Создание кнопок Перенос
        Удалить_Объекты
            ActiveSheet.Buttons.Add(55, 0, 125, 20).Select
                Selection.OnAction = "VBAProjectSO.Module2.Перенос_по_строкам"
                Selection.Characters.Text = "Перенос по строкам"
                
            ActiveSheet.Buttons.Add(350, 0, 125, 20).Select
                Selection.OnAction = "VBAProjectSO.Module2.На_печать_выборочно"
                Selection.Characters.Text = "Перенос по листам"
                
        Worksheets("СО").Activate
            Копирование_КнопокСО
        Worksheets("ВР").Activate
            Копирование_КнопокСО
                
            
           
Sheets("Спецификация").Activate

If ActiveWorkbook.Sheets("ВР").Range("A15").Value <> "" And ActiveWorkbook.Sheets("ВР").Range("A15").Value <> "" Then
    a = MsgBox("Обновить до ГОСТ Р 21.1101-2013?", vbYesNo)
    If a = vbYes Then Обновить_до_ГОСТ
End If
Application.ScreenUpdating = True
End If
End Sub

Function List() As Boolean 'Проверка существования листов
Dim a As Boolean, b As Boolean, c As Boolean, d As Boolean
    a = IsWorkSheetExist("Спецификация")
    b = IsWorkSheetExist("Перенос")
    c = IsWorkSheetExist("СО")
    d = IsWorkSheetExist("ВР")
        List = a And b And c And d
'            MsgBox (List)
End Function

 Function Удалить_Объекты()
    ActiveSheet.DrawingObjects.Select
    Selection.Delete
 End Function
 
 Sub Удалить_лист_Шаблон()
    
    If IsWorkSheetExist("Шаблоны") = True Then
        
            Sheets("Шаблоны").Delete
        
    End If
 End Sub
 
Sub Сохранить_XLSX()
Dim a As Boolean, b As Boolean
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
a = FileLocation(strFileN)
b = FileLocation(Full)
If a And b Then
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
Set a = FS.CreateTextFile(ИмяФайла, True)
a.WriteLine ("Старый путь нахождения файла" & vbNewLine & СтарыйПуть)
a.Close
End Sub


Sub Posicii()
On Error GoTo Error:
Dim lLastRow As Integer, i As Integer, iNum As Integer
Dim rRange As Range, rPosit As Range
Dim st As String
Dim bNum As Boolean
    If Cells(3, 1).Value <> 1 Then
    MsgBox "Нет возможности пересчитать позиции", vbCritical
    End
    End If
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Set rRange = Range(Cells(4, 1), Cells(lLastRow, 1))
    For Each rPosit In rRange
        i = i + 1
        iNum = iNum + 1
        st = rRange(i, 1).Value
        If IsNumeric(st) Then
Error:      rRange(i, 1).Value = "=R[-" & iNum & "]C+1"
            iNum = 0
        End If
    Next rPosit

End Sub

Function bCount() As Byte
Dim lLastRow As Integer, i As Integer, bMod As Byte
Dim rRange As Range
Set rRange = ActiveWorkbook.Sheets("Перенос").Range("J2")
lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
For i = 4 To lLastRow
    Set rRange = rRange.Offset(1, 0)
'    rRange.Activate
    If rRange.Value <> "" Or i = 29 Then
        Set rRange = rRange.Offset(1, 0)
        Exit For
    End If
Next i
If lLastRow = i Then
    bCount = 1
Else
    bCount = 2
End If
bMod = 1
For i = i To lLastRow
    If rRange.Value <> "" Or bMod = 31 Then
        bCount = bCount + 1
        bMod = 1
    End If
    bMod = bMod + 1
    Set rRange = rRange.Offset(1, 0)
'    rRange.Activate
Next i
MsgBox bCount
End Function

Function Обновить_до_ГОСТ()
    If ActiveWorkbook.Sheets("ВР").Range("A15").Value <> "" Then 'Перенос строки согласовано
        ActiveWorkbook.Sheets("ВР").Range("A15:A23").Cut Destination:=ActiveWorkbook.Sheets("ВР").Range("B15:B23")
    End If
    
    If ActiveWorkbook.Sheets("СО").Range("A15").Value <> "" Then
        ActiveWorkbook.Sheets("СО").Range("A15:A23").Cut Destination:=ActiveWorkbook.Sheets("СО").Range("B15:B23")
    End If
End Function


