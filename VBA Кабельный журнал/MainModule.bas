Attribute VB_Name = "MainModule"
Sub Создать_Лист_КЖ()
'
' Создать_Лист_КЖ Макрос
' Создание нового листа КЖ
'

'
    Application.ScreenUpdating = False
    
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = False
    Application.Goto Reference:="КЖ"
    Selection.Copy
    Application.Goto Reference:="Разметка"
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    ActiveCell.Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A20"), Type:= _
                                    xlFillCopy
    ActiveCell.Range("A1:A20").Select
    ActiveWindow.SmallScroll Down:=-6
    ActiveCell.Offset(22, 16).Range("A1:A2").Select
    Selection.FormulaR1C1 = "=R[-24]C+1"
    ActiveCell.Offset(-1, -4).Range("A1:D3").Select
    Selection.FormulaR1C1 = "=Содержание!R29C36"
    ActiveCell.Offset(-21, -11).Range("A1:P20").Select
    Selection.ClearContents
    ActiveCell.Offset(0, 1).Range("A1:O20").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ActiveCell.Offset(0, -1).Range("A1:P24").Select
    ActiveCell.Offset(0, 13).Range("A1").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    Application.ScreenUpdating = True
End Sub

Sub Удалить_Лист_КЖ()
'
' Удалить_Лист_КЖ Макрос
' Удаление последнего листа КЖ
'

'
    Application.ScreenUpdating = False
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = False
    Application.Goto Reference:="Разметка"
    Selection.End(xlDown).Select
    ActiveCell.Offset(-23, -18).Range("A1:S24").Select
    ActiveCell.Activate
    Selection.Clear
    
    Application.Goto Reference:="Разметка"
    Selection.End(xlDown).Select
    ActiveCell.Offset(-3, -17).Range("A1:P4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    Application.ScreenUpdating = True
   
End Sub



Sub ЗаменитьИксы() 'Корректирует все сивмолы Х, . Латинские/русские
Dim A
Dim VarStr As Variant, strR As String
A = Array("x", "X", "Х", ".")
For Each VarStr In A
strR = VarStr
    If strR <> "." Then
        Call Заменить(strR, "х", False, _
        ThisWorkbook.Worksheets("КЖ").Range(Cells(4, 5), Cells(10000, 5)))
    Else
        Call Заменить(strR, ",", False, _
        ThisWorkbook.Worksheets("КЖ").Range(Cells(4, 5), Cells(10000, 5)))
    End If
Next
End Sub
 
Sub Печать_в_PDF()
    Dim Name As String
    Скрыть (True)
    Name = SpecialFolderPath 'Путь рабочего стола
    If Dir(Name & "\PDF Спецификации", vbDirectory) = "" _
        Then MkDir (Name & "\PDF Спецификации")  'Создание папки для сохранения
    
    Sheets(Array("Содержание", "Сводная ведомость A3", "КЖ")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        Name & "\PDF Спецификации\" & ActiveWorkbook.Name & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
    Sheets("КЖ").Select
End Sub


Sub СкрытьОтобразить()
    If Sheets("Сводная ведомость A3").Rows("1:2").EntireRow.Hidden Then
        Скрыть (False)
    Else
        Скрыть (True)
    End If
End Sub





