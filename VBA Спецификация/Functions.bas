Attribute VB_Name = "Functions"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit
Function IsBookOpen(wbName As String) As Boolean 'Проверка на открытие файла
    Dim wbBook As Workbook: On Error Resume Next
    Set wbBook = Workbooks(wbName)
    IsBookOpen = Not wbBook Is Nothing
End Function

Function Какие_Ячейки_Выбрать(Описание As String, Заголовок As String) As String
Dim rRange As Range
Set rRange = Application.InputBox(Prompt:=Описание, title:=Заголовок, Type:=8)
Какие_Ячейки_Выбрать = rRange.Address
End Function


Sub Draws_In_Selection_Select() ' выделить все рисунки в выбранном диапазоне и удалить
'обсуждение http://www.planetaexcel.ru/forum/index.php?FID=8&PAGE_NAME=read&TID=37169
If TypeName(Selection) <> "Range" Then Exit Sub
Dim oDraw
On Error Resume Next
With CreateObject("Scripting.Dictionary")
For Each oDraw In ActiveSheet.DrawingObjects '.ShapeRange
If Not Intersect(Selection, Range(oDraw.TopLeftCell, oDraw.BottomRightCell)) Is Nothing Then .Add oDraw.Name, oDraw
Next
If .Count > 0 Then
    ActiveSheet.Shapes.Range(.keys).Select
    Selection.Delete
End If
End With
End Sub

Function FileLocation(strFileName As String) As Boolean 'Проверка существования файла (полное имя)
'   Dim strFileName As String
   ' Имя искомого файла
'   strFileName = strFileN
   ' Проверка наличия файла (функция Dir возвращает пустую _
    строку, если по указанному пути файл обнаружить не удалось)

   If Dir(strFileName) <> "" Then
      FileLocation = True 'MsgBox "Файл " & strFileName & " найден"
   Else
      FileLocation = False '"Файл " & strFileName & " не найден"
   End If
End Function

Sub Диапазон_в_ячейку()
Dim msgQ As Byte
Dim rRange As Range, r1Range As Range
Dim sRange As String
Dim sCell As String 'Значение для вставки в ячейку
    ThisWorkbook.Sheets("Лист1").Activate
'    Do
        sRange = Какие_Ячейки_Выбрать("Выберите ячейки для копирования", "Диапазон ячеек")
        
        Set rRange = ActiveSheet.Range(sRange)
            For Each r1Range In rRange
                sCell = sCell + r1Range.Value + ", "
            Next
            msgQ = Len(sCell)
            
            sCell = Left(sCell, msgQ - 2)
'            MsgBox sCell
'        msgQ = MsgBox(sCell, vbYesNoCancel, "Выбранный диапазон")
'        If msgQ = vbCancel Then Exit Sub
'    Loop Until msgQ = vbYes
        ThisWorkbook.Sheets("Лист2").Activate
    Do
        sRange = Какие_Ячейки_Выбрать("Выберите ячейку для вставки", "Выбор ячейки")
            msgQ = InStr(1, sRange, ":")
            If msgQ > 0 Then Call MsgBox("Нужна только одна ячейка!!!", vbCritical)
    Loop Until msgQ = 0
    Set rRange = ActiveSheet.Range(sRange)
    rRange.Value = sCell
    
End Sub
'Функции в листе
Function IsWorkSheetExist(sSName As String) As Boolean 'Проверка существования листа активной книги
Dim c As Object
On Error GoTo errНandle:
'   Set c = Sheets(sName)
   ' Альтернативный вариант :
    ActiveWorkbook.Sheets(sSName).Unprotect
    Worksheets(sSName).Cells(1, 1) = Worksheets(sSName).Cells(1, 1)
    IsWorkSheetExist = True
Exit Function
errНandle:
   IsWorkSheetExist = False
End Function

Function IsWorkSheetExistXLAM(sSName As String) As Boolean 'Проверка существования листа в надстройке
Dim c As Object
On Error GoTo errНandle:
    ThisWorkbook.Sheets(sSName).Unprotect
    ThisWorkbook.Worksheets(sSName).Cells(1, 1) = ThisWorkbook.Worksheets(sSName).Cells(1, 1)
    IsWorkSheetExistXLAM = True
Exit Function
errНandle:
   IsWorkSheetExistXLAM = False
End Function

Sub Редактор_Книги() 'Делает книгу надстройки доступной для редактрования
    If ThisWorkbook.IsAddin = False Then
    ThisWorkbook.IsAddin = True
    Exit Sub
    End If
    If ThisWorkbook.IsAddin = True Then ThisWorkbook.IsAddin = False
End Sub

Sub Change_ReferenceStyle() 'Замена стилей R1C1
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub

Sub Сохранить_Книгу() 'Сохраняет книгу надстройки
Dim a As Byte
    a = MsgBox("Действительно пересохранить файл Надстройки?", vbYesNo)
    If a = vbYes Then ThisWorkbook.Save
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

Sub Удалить_пробелы()
  Dim v As Range
  For Each v In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants)
    v.Value = Trim(v)
    While InStr(1, v, "  ", vbTextCompare) > 0
      v.Value = Replace(v, "  ", " ")
    Wend
  Next
End Sub

Sub Заменить(sWhat As String, sReplacement As String, Целиком As Boolean, rRange As Range)
'Что меняем, На что меняем, Ячейка целиком или часть текста, Диапазон
    If Целиком Then
        rRange.Replace What:=sWhat, Replacement:=sReplacement, LookAt:=xlWhole
    Else
        rRange.Replace What:=sWhat, Replacement:=sReplacement, LookAt:=xlPart
    End If
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



