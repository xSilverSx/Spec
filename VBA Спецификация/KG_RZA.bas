Attribute VB_Name = "KG_RZA"
'Модуль для извлечения данных кабельного журнала РЗА, для основной программы не требуется
'Извлекает данные пригодные для программы ZCAD
Option Explicit

Sub Перенос_ZCAD()
Application.ScreenUpdating = False
Dim rngStart As Range, rngAll As Range, angX As Range, rngSumm As Range, KG As Range, Cops As Range
Dim i As Integer, b As Integer
Dim lnRow, lnCol, lnR As Long ' lnR Разница между LnRow и старым lnrow
Dim k As Byte
Dim a As Boolean 'Логическое значение для первой строчки
Set Sheet1 = Worksheets("Кабельный журнал")
Set Sheet2 = Worksheets("ZCAD")
Sheet2.Activate
Rows("2:10000").Select
Selection.Delete
Range("a2").Select

Set KG = Sheet1.Range("N4:AT7")
Set Cops = Sheet2.Range("A2:F3")
         
For i = 1 To 5000

    If KG.Cells(1, 1) <> "" Then
    Cops.Cells(1, 1) = KG.Cells(1, 1) 'Номер кабеля
    Cops.Cells(1, 4) = "@NET0"
'    MsgBox KG.Cells(3, 8).Address
    Cops.Cells(1, 5) = KG.Cells(1, 5) & KG.Cells(3, 5) & KG.Cells(3, 7) & KG.Cells(3, 8) 'Марка кабеля
    Cops.Cells(1, 6) = KG.Cells(2, 10)
    Cops.Cells(1, 7) = KG.Cells(2, 22)
    Set Cops = Cops.Offset(1, 0)
    End If
    Set KG = KG.Offset(4, 0)
Next i
Удалить_пробелы
Range("a2").Activate
End Sub


