Attribute VB_Name = "MainModule"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit
Public Sheet1 As Worksheet, Sheet2 As Worksheet, Sheet3 As Worksheet 'Объявлены перенные Листов
Public Spec As Range, Pechat As Range, Posit As Range
Public AArray
Public i As Integer
Public bPerenos As Boolean 'Обозначает что выполняется функция переноса

Function Array_A() 'Объявление массивов
    AArray = Array(0, 1, 2, 3, 4, 11, 20, 24, 28, 33) 'Массив колонок для содержимого
End Function

Sub Podgotovka_Show() 'Показать форму для переноса
    Podgotovka.Show
End Sub

Sub Перенос()
Application.ScreenUpdating = False
Dim rngStart As Range, rngAll As Range, angX As Range, rngSumm As Range, Spec As Range, Cops As Range
Dim i As Integer, b As Integer
Dim lnRow, lnCol, lnR As Long ' lnR Разница между LnRow и старым lnrow
Dim k As Byte
Dim A As Boolean 'Логическое значение для первой строчки
Dim rPropusk As String
Application.Run "VBAProjectSO.ЭтаКнига.Oshibka_na_liste_Perenos" 'Запуск функции на листе перенос
Set Sheet1 = Worksheets("Спецификация")
Set Sheet2 = Worksheets("Перенос")
    If Podgotovka.Perenos = True Then 'Скрыть/отобразить лист Перенос, по необходимости
            Sheet2.Visible = xlSheetVisible
        Else: Sheet2.Visible = xlSheetHidden
    End If
For k = 1 To 2 'Перенос всех листов СО и(или) ВР
    If Podgotovka.SO = False And k = 2 Then Exit For
    If Podgotovka.VR = True And k = 1 Then Sheet2.Range("O1") = "ВР" 'Записывает что будет создаваться ведомость объемов работ
    If Podgotovka.VR = False And k = 1 Then k = 2
    If Podgotovka.SO = True And k = 2 Then Sheet2.Range("O1") = "СО" 'Записывает что будет создаваться спецификация
    A = False
    lnR = 0
PerenosClear 'Очистка листа перенос
Sheet2.Activate
Set Spec = Sheet1.Range("A2:I2")
Set Cops = Sheet2.Range("A3:I3")
bPerenos = True
Do While Spec.Cells(2) <> ""
                If A Then      'Прописывается разделитель для СО или ВР
                    If Podgotovka.VR = True And k = 1 Then Cops.Cells(1, 10) = Spec.Cells(1, 11)
                    If Podgotovka.SO = True And k = 2 Then Cops.Cells(1, 10) = Spec.Cells(1, 10)
                End If
                A = True
    If Sheet2.Range("O1") = "ВР" Then
        For i = 1 To 9
            Cops.Cells(1, i) = Spec.Cells(1, i)
            If Cops.Cells(1, i) = "" Then Cops.Cells(1, i) = 0
        Next i
        Set rngAll = Cops.CurrentRegion
        lnRow = rngAll.Rows.Count
        Set Spec = Spec.Offset(1, 0)
    Else
            If Spec.Cells(1) = "вр" Then
                Set Spec = Spec.Offset(1, 0)
            Else
            For i = 1 To 9
                Cops.Cells(1, i) = Spec.Cells(1, i)
                If Cops.Cells(1, i) = "" Then Cops.Cells(1, i) = 0
            Next i
            Set rngAll = Cops.CurrentRegion
            lnRow = rngAll.Rows.Count
            Set Spec = Spec.Offset(1, 0)
    End If
    End If
            Select Case Spec.Cells(1).Value
                Case "", "вр"
                    Set Cops = Cops.Offset(lnRow - lnR, 0)
                    lnR = lnRow
                Case Else
                        rPropusk = Right(Spec.Cells(1).Value, 1) 'Записать последний символ в строке позиции
                        If Podgotovka.NePropusk = True Or rPropusk = "_" Then 'Если стоит галочка не пропускать, или последний символ "_" то строка не пропускается
                                Set Cops = Cops.Offset(lnRow - lnR, 0)
                                lnR = lnRow
                                Else
                                Set Cops = Cops.Offset(lnRow - lnR + 1, 0)
                                lnR = 0
                        End If
            End Select
Loop
bPerenos = False
Удалить_пробелы
Range("a2").Activate
            If Podgotovka.Perenos = False Then На_печать_все_листы 'Подготовить печать полностью, или только на лист перенос
Next k
    Application.ScreenUpdating = True
End Sub

Sub Перенос_Строк_На_СО_ВР()
Dim b As Byte
            If Spec.Cells(1, 1).Value = "ч" Then 'убрать или добавить подчеркивание
                Pechat.Cells(1, AArray(2)).Font.Underline = xlUnderlineStyleSingle
                Else: Pechat.Cells(1, AArray(2)).Font.Underline = xlUnderlineStyleNone
            End If
                  If i = 1 Or i = 2 Then 'Принудительно провести перенос первой строчки
                        For b = 1 To 9
                            Pechat.Cells(1, AArray(b)) = Spec.Cells(1, b)
                        Next b
                                Set Spec = Spec.Offset(1, 0)
                                Set Pechat = Pechat.Offset(1, 0)
                  Else
                        If Spec.Cells(1, 10).Value = "" Then 'Проверка метки переноса
                            For b = 1 To 9
                                    Pechat.Cells(1, AArray(b)) = Spec.Cells(1, b)
                            Next b
                                Set Spec = Spec.Offset(1, 0)
                                Set Pechat = Pechat.Offset(1, 0)
                         ElseIf i = 1 Then                      'не переносить первую строку
                            For b = 1 To 9
                                    Pechat.Cells(1, AArray(b)) = Spec.Cells(1, b)
                            Next b
                                Set Spec = Spec.Offset(1, 0)
                                Set Pechat = Pechat.Offset(1, 0)
                         Else
                                For b = 1 To 9
                                        Pechat.Cells(1, AArray(b)) = " " 'очистить содержимое ячеек которые не переносились
                                Next b
                            Set Pechat = Pechat.Offset(1, 0)
                        End If
            End If
End Sub
Function List_Count() As Byte 'Количество листов
Set Posit = Range("AP73")
List_Count = 1
    Do While Posit.Value <> ""
            List_Count = List_Count + 1
            Set Posit = Posit.Offset(34, 0)
    Loop
'    MsgBox (List_Count)
End Function

Sub На_печать_все_листы()
Dim b, intS, c As Integer, Buf As Integer
'Неверное_Количество_Листов
Array_A
Set Sheet2 = Worksheets("Перенос")

        If Sheet2.Range("O1") = "СО" Then 'Убрать получение листов из Спецификации
            Set Sheet3 = Worksheets("СО")
'            intS = Worksheets("Спецификация").Range("J2").Value
        Else
            Set Sheet3 = Worksheets("ВР")
'            intS = Worksheets("Спецификация").Range("K2").Value
        End If
                intS = bCount 'листы с листа "Перенос"
                Buf = intS
Set Spec = Sheet2.Range("A3")
Set Pechat = Sheet3.Range("E2")
Sheet3.Activate
c = List_Count ' Количество листов на листе СО или ВР
        Do While intS > c
            Добавить_Лист
            intS = intS - 1
        Loop
        If c > intS Then
        i = MsgBox("Указано меньшее количество листов, чем было до этого. Удалить последние листы?", vbYesNo + vbDefaultButton1)
            If i = vbYes Then
                Range(Cells(40 + (35 * (intS - 1)), 1), Cells(40 + (35 * (c - 1)), 50)).Delete 'удаление лишних листов в случае если их больше чем указано
                Range("AO35").Value = intS
            End If
        End If
For i = 1 To 26 'Заполняет первый лист
    Перенос_Строк_На_СО_ВР
Next i
Set Pechat = Pechat.Offset(12, 0) 'окончание первого листа
intS = Buf
If intS > 1 Then
    For intS = 1 To intS - 1
            For i = 1 To 31 'Заполенние последующих листов
                Перенос_Строк_На_СО_ВР
            Next i
        Set Pechat = Pechat.Offset(4, 0)
    Next intS
Else
End If
Удалить_Знаки
Sheet3.Activate
End Sub

Sub На_печать_выборочно()
Dim b, intS As Integer
Dim strSpec, strPec, Vozvrat As Range
Array_A
Set Sheet2 = Worksheets("Перенос")
        If Sheet2.Range("O1") = "СО" Then
            Set Sheet3 = Worksheets("СО")
        Else
            Set Sheet3 = Worksheets("ВР")
        End If
Sheet2.Activate
Set strSpec = Application.InputBox(Prompt:="Ячейка спецификации для переноса", title:="Задайте начальную ячейку переноса (ячейка с позицией)", Type:=8)
Sheet3.Activate
Set strPec = Application.InputBox(Prompt:="Ячейка спецификации СО", title:="Задайте начальную ячейку переноса (ячейка с позицией)", Type:=8)

Set Spec = Sheet2.Range(strSpec.Address)
Set Pechat = Sheet3.Range(strPec.Address)
Set Vozvrat = strPec

intS = InputBox("Задайте количество листов для переноса спецификации Значения от 1 - 50", "Листов для переноса", "1")

Do While intS <= 0 Or intS > 50
    MsgBox "Вы задали не коректное число листов, задайте другое число", vbCritical
    intS = InputBox("Вы задали не коректное число листов, попробуйте еще раз значения от 1 - 50", "Ошибка ввода", "2")
Loop

    For intS = 1 To intS
        For i = 1 To 31
            Перенос_Строк_На_СО_ВР
        Next i
    Set Pechat = Pechat.Offset(4, 0)
    Next intS
    Sheet3.Activate
    Удалить_Знаки
    Sheet3.Range(Vozvrat.Address).Select
End Sub
Sub Перенос_по_строкам()
    Dim b, intS As Integer
    Dim strSpec, strPec, Vozvrat As Range
Array_A
    Set Sheet2 = Worksheets("Перенос")
            If Sheet2.Range("O1") = "СО" Then
                Set Sheet3 = Worksheets("СО")
            Else
                Set Sheet3 = Worksheets("ВР")
            End If
    Sheet2.Activate
    Set strSpec = Application.InputBox(Prompt:="Ячейка спецификации для переноса", title:="Задайте начальную ячейку переноса (ячейка с позицией)", Type:=8)
    Sheet3.Activate
    Set strPec = Application.InputBox(Prompt:="Ячейка спецификации СО", title:="Задайте начальную ячейку переноса (ячейка с позицией)", Type:=8)
    
    Set Spec = Sheet2.Range(strSpec.Address)
    Set Pechat = Sheet3.Range(strPec.Address)
    Set Vozvrat = strPec
    
    intS = InputBox("Задайте количество строк для переноса спецификации Значения от 1 - 31" & vbCr & "Прим. на первом листе 26 строк" & vbCr & "На последующих 31", _
    "Строк для переноса", "2")
    
    
    Do While intS <= 0 Or intS > 31
        MsgBox "Вы задали не коректное число строк, задайте другое число", vbCritical
        intS = InputBox("Задайте количество строк для переноса спецификации Значения от 1 - 31. Прим. на первом листе 26 строк, на последующих 31", _
        "Строк для переноса", "2")
    Loop
    
        For intS = 1 To intS
            
                    If Spec.Cells(1, 1).Value = "ч" Then 'убрать или добавить подчеркивание
                        Pechat.Cells(1, AArray(2)).Font.Underline = xlUnderlineStyleSingle
                    Else: Pechat.Cells(1, AArray(2)).Font.Underline = xlUnderlineStyleNone
                    End If
            
            
            For b = 1 To 9
                    Pechat.Cells(1, AArray(b)) = Spec.Cells(1, b)
            Next b
                    Set Spec = Spec.Offset(1, 0)
                    Set Pechat = Pechat.Offset(1, 0)
        Next intS
        Sheet3.Activate
        Удалить_Знаки
        Sheet3.Range(Vozvrat.Address).Select
End Sub

Sub Добавить_Лист()
    Dim b As Integer
    Dim rngX As Range
      ThisWorkbook.Sheets("Шаблоны").Rows("1:35").Copy
        Set rngX = Range("ap73")
            Do While rngX.Value <> ""
                Set rngX = rngX.Offset(34, 0)
            Loop
        Set rngX = rngX.Offset(-33, -41)
        rngX.Select
        Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone
        Set rngX = rngX.Offset(33, 41)
            If rngX.Address <> "$AP$73" Then
                rngX.FormulaR1C1 = "=R[-35]C+1"
                Else: rngX.FormulaR1C1 = "2"
            End If
            Range("AO35").Formula = rngX.Value
            rngX.Offset(-33, -34).Select
End Sub

Sub Удалить_пробелы()
  Dim v As Range
  For Each v In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants)
    v.Value = Trim(v)
    While InStr(1, v, "  ", vbTextCompare) > 0
      v.Value = Replace(v, "  ", " ")
    Wend
  Next
End Sub

Sub Удалить_Знаки()
    Call Заменить("_", " ", False, Cells)
    Call Заменить("вр", " ", True, Range("E:E"))
    Call Заменить("ч", " ", True, Range("E:E"))
    Call Заменить(0, "", True, Cells)
    Call Заменить(" ", "", True, Cells)
End Sub

Sub Чистка_Печати()
    If ClearCont.ChListPerenos = True Then 'Чистка листа перенос
        PerenosClear
    End If
    If ClearCont.ChListSO = True Then 'Чистка листа СО
        SoVrClear ("СО")
    End If
    If ClearCont.ChListVR = True Then 'Чистка листа ВР
        SoVrClear ("ВР")
    End If
End Sub

Sub PerenosClear()
    ActiveWorkbook.Sheets("Перенос").Rows("2:10000").Delete
End Sub

Sub SoVrClear(strSheet As String)
    With ActiveWorkbook.Sheets(strSheet)
        .Rows("40:10000").Delete
        .Range("E2:AK27").Value = ""
        .Range("AO35").Value = "1"
    End With
End Sub

Sub Очистить_всё()
    ClearCont.Show
End Sub

Function bCount() As Byte 'Подсчитывает количество листов которое будет в спецификации
Dim lLastRow As Integer, i As Integer, bMod As Byte
Dim rRange As Range
Set rRange = ActiveWorkbook.Sheets("Перенос").Range("J2")
lLastRow = ActiveWorkbook.Sheets("Перенос").UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
For i = 4 To lLastRow
    Set rRange = rRange.Offset(1, 0)
    If rRange.Value <> "" Or i = 29 Then
        Set rRange = rRange.Offset(1, 0)
        Exit For
    End If
Next i
If lLastRow < i Then
    bCount = 1
Else
    bCount = 2
End If
bMod = 1
For i = i To lLastRow
    If rRange.Value <> "" Or bMod = 31 Then
        If bMod <> 2 Then
                bCount = bCount + 1
                bMod = 0
        Else
        End If
    End If
    bMod = bMod + 1
    Set rRange = rRange.Offset(1, 0)
Next i
End Function

