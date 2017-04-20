Attribute VB_Name = "Func"
Option Explicit
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
