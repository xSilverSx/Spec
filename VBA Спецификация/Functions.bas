Attribute VB_Name = "Functions"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.

Function IsBookOpen(wbName As String) As Boolean
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




