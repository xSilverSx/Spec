VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SO_Zapolnen 
   Caption         =   "Выберите элемент для спецификации"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   8640.001
   OleObjectBlob   =   "SO_Zapolnen.frx":0000
   ShowModal       =   0   'False
   Tag             =   "0"
End
Attribute VB_Name = "SO_Zapolnen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public sL As Single, sSh As Single, Massa As Single 'Длина и ширина и масса полосы
Public arr, i
'arr(i, 1) - Категория
'arr(i, 2) - Подкатегория
'arr(i, 3) - Краткое наименование
'arr(i, 4) - Полное наименование
'arr(i, 5) - Тип
'arr(i, 6) - Нормативный документ
'arr(i, 7) - Код оборудования
'arr(i, 8) - Завод
'arr(i, 9) - Единица измерения
'arr(i, 10) - Масса


Private Sub UserForm_Initialize() 'Выбор ячейки категория
Dim a As Integer
With ThisWorkbook.Sheets("База_СО") 'Заполнение ячейки категория
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "A").End(xlUp)).Value
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            .Item(arr(i, 1)) = 1
        Next
        Me.ComboBox1.List = .keys
        ComboBox1.Value = ComboBox1.List(0)
    End With
For a = 500 To 3000 Step 500 'Заполнение ячейки Длинна
    ComboBox12.AddItem a
    ComboBox12.Value = "1000"
Next a
For a = 25 To 100 Step 5    'Заполнение ячейки ширина
    ComboBox13.AddItem a
    ComboBox13.Value = "50"
Next a
ThisWorkbook.Saved = True
End Sub

Private Sub ComboBox1_Change() 'Заполнение ячейки подкатегория
With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "B").End(xlUp)).Value
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 1) = ComboBox1.Value Then 'Учет значений Категории
                .Item(arr(i, 2)) = 1
            End If
        Next
        Me.ComboBox2.List = .keys
        ComboBox2.Value = ComboBox2.List(0)
    End With
End Sub

Private Sub ComboBox2_Change() 'Заполнение ячейки Краткое наименование
With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "C").End(xlUp)).Value
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 1) = ComboBox1.Value Then
                If arr(i, 2) = ComboBox2.Value Then
                    .Item(arr(i, 3)) = 1
                End If
            End If
        Next

        Me.ComboBox3.List = .keys
        ComboBox3.Value = ComboBox3.List(0)
    End With
End Sub

Private Sub ComboBox3_Change()

With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "D").End(xlUp)).Value 'Тип оборудования
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    .Item(arr(i, 4)) = 1
                End If
            End If
        Next
        
        Me.ComboBox4.List = .keys
        ComboBox4.Value = ComboBox4.List(0)
End With
ComboBox3.ControlTipText = ComboBox3.Value
End Sub

Private Sub ComboBox4_Change()
On Error Resume Next 'При ошибке продалжает работу программы, иначе вызывает ошибку индекса в "Коде оборудования" combobox6
ComboBox10.Clear 'Очистка комбобокса, перед добавлением новых элементов
With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "E").End(xlUp)).Value 'Запись данных в полное наименование
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    If arr(i, 4) = ComboBox4.Value Then
                        ComboBox10.AddItem (arr(i, 5))
'                        .Item(arr(i, 5)) = 1
                    End If
                End If
            End If
        Next
End With

With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "F").End(xlUp)).Value 'Запись данных в Нормативный документ
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    .Item(arr(i, 6)) = 1
                End If
            End If
        Next
        Me.ComboBox5.List = .keys
        ComboBox5.Value = ComboBox5.List(0)
End With

With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "G").End(xlUp)).Value 'Запись данных в код оборудования
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    .Item(arr(i, 7)) = 1
                End If
            End If
        Next
        Me.ComboBox6.List = .keys
        ComboBox6.Value = ComboBox6.List(ComboBox4.ListIndex)
End With

With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "H").End(xlUp)).Value 'Запись данных в Завод
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    .Item(arr(i, 8)) = 1
                End If
            End If
        Next
        Me.ComboBox7.List = .keys
        ComboBox7.Value = ComboBox7.List(0)
End With

With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "I").End(xlUp)).Value 'Запись данных в единицу измерения
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    If arr(i, 4) = ComboBox4.Value Then
                        .Item(arr(i, 9)) = 1
                    End If
                End If
            End If
        Next
        Me.ComboBox8.List = .keys
        ComboBox8.Value = ComboBox8.List(0) 'ComboBox8.List(ComboBox4.ListIndex)
End With

With ThisWorkbook.Sheets("База_СО")
arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "J").End(xlUp)).Value 'Запись данных в массу
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    If arr(i, 4) = ComboBox4.Value Then
                        .Item(arr(i, 10)) = 1
                    End If
                End If
            End If
        Next
        Me.ComboBox9.List = .keys
        ComboBox9.Value = ComboBox9.List(0)
End With

With ThisWorkbook.Sheets("База_СО")
arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "M").End(xlUp)).Value 'Вставка примечания
End With
    With CreateObject("Scripting.Dictionary")
        For i = LBound(arr) To UBound(arr)
            If arr(i, 2) = ComboBox2.Value Then
                If arr(i, 3) = ComboBox3.Value Then
                    If arr(i, 4) = ComboBox4.Value Then
                        Me.ComboBox4.ControlTipText = arr(i, 11)
                    End If
                End If
            End If
        Next
End With


ComboBox5_Change
ComboBox10.Value = ComboBox10.List(0)
Скрыть_Лишнее
K_T_Change
End Sub

Function Скрыть_Лишнее()
'Скрыть лишние и пересчитать при необходимости добавить нукжные разделы сюда
Select Case ComboBox2.Value
    Case "Лотки", "Рукава", "Арматура", "Двутавры", "Прокат", "Шины", "Трубы хризотилцементные"
    'Работает для данных подкатегорий
        Select Case ComboBox3.Value
            Case "Шина медная", "Лист стальной", "Полоса стальная"
    'Показывает длину и ширину
                РазделПоказать (True)
                ДлинШирПоказать True, True
            Case Else
    'Для остальных показывает только длинну
                РазделПоказать (True)
                ДлинШирПоказать True, False
        End Select
    Case "Болты"
    'Для болтов вывести "пакет"
                РазделПоказать (True)
                ДлинШирПоказать False, False
                Frame18.Left = 372
    Case "Раздел"
    'Установить раздел
                РазделПоказать (False)
    Case Else
    'Скрыть раздел скрыть длинну и ширину
                РазделПоказать (True)
                ДлинШирПоказать False, False
End Select
End Function
Function РазделПоказать(bVisible As Boolean) 'Скрыть показать кнопки раздела
    Frame2.Visible = bVisible
    Frame4.Visible = bVisible
    Frame18.Left = 432
    If bVisible Then
            Me.Height = 200
            Me.Width = 437
            BtnVst.Left = 372
        Else
            BtnVst.Left = 145
            Me.Height = 165
            Me.Width = 209
            ComboBox17.Value = Null
            ComboBox11.Value = Null
    End If
End Function
Function ДлинШирПоказать(bDlina As Boolean, bShirina As Boolean)
    Frame12.Visible = bDlina
    Frame13.Visible = bShirina
    If bDlina And bShirina Then
        ComboBox11.Value = Massa_Polosy(Val(Replace(ComboBox9.Value, ",", ".")), _
        Val(Replace(ComboBox12.Value, ",", ".")), Val(Replace(ComboBox13.Value, ",", ".")))
    ElseIf bDlina Or bShirina Then
        ComboBox11.Value = Massa_Polosy(Val(Replace(ComboBox9.Value, ",", ".")), _
        Val(Replace(ComboBox12.Value, ",", ".")), 1000)
    Else
        ComboBox12.Value = "1000"
        ComboBox11.Value = ComboBox9.Value
    End If
End Function
Function Massa_Polosy(Massa, sL, sSh) As Single 'Подсчет полосы по длине и ширине
    Massa_Polosy = Massa * sL * sSh / 1000000
End Function

Private Sub BtnVst_Click()
Dim i As Integer, l As Integer, k As Integer
Dim s, adDr, tmpArr, AddrTmp As String
Dim val_arr(0 To 7)
'Dim ExitWhile As Boolean
Application.ScreenUpdating = False
val_arr(0) = ComboBox15.Value   'Позиция
val_arr(1) = ComboBox10.Value   'Наименование
val_arr(2) = ComboBox17.Value   'Обозначение
val_arr(3) = ComboBox6.Value    'Код
val_arr(4) = ComboBox7.Value    'Завод
val_arr(5) = ComboBox8.Value    'Единица измерения
val_arr(6) = ComboBox16.Value   'Количество
val_arr(7) = ComboBox11.Value   'Масса единицы
    If ActiveCell.Row = 1 Then Exit Sub 'Условие не вставки в первую ячейку
    adDr = ActiveCell.Address 'Получение адреса ячейки
    tmpArr = Split(adDr, "$") 'Разбиение на массив строки и столбца
    AddrTmp = tmpArr(1) & tmpArr(2)
    tmpArr(1) = "A"
    s = Replace("A" & (tmpArr(2) + 1), " ", "") 'Адрес ячейки на которую необходимо будет перейти
    For i = 0 To 7
    s = Replace(Chr(Asc(tmpArr(1)) + i) & tmpArr(2), " ", "")
    Range(s).Select
        Select Case i
            Case 1, 6, 7
                Range(s).FormulaLocal = Replace(val_arr(i), ".", ",")
            Case 0
                Range(s).FormulaLocal = "A" 'Заглушка (иначе у таблицы появляется итог)
                k = Val(tmpArr(2))
                    Do While IsNumeric(Cells(k, 1)) = False Or Cells(k, 1) = "" Xor Cells(k, 1).Address = "$A$1"
                    'Проверяет на число/не число, что ячейка не пустая, что ячейки не пустые
                        l = l + 1
                        k = k - 1
                    Loop
                    If Cells(k, 1).Address = "$A$1" And val_arr(i) <> "ч" And val_arr(i) <> "вр" Then
                        Range(s).FormulaLocal = 1
                    Else
                        Range(s).FormulaLocal = Replace(val_arr(i), "-1", "-" & l)
                    End If
                        If Range(s).Value = "ч" Then
                        Range(Chr(Asc(tmpArr(1)) + 1) & tmpArr(2)).Font.Underline = xlUnderlineStyleSingle
                        Range(Chr(Asc(tmpArr(1)) + 1) & tmpArr(2)).Font.Bold = True
                        Else
                        Range(Chr(Asc(tmpArr(1)) + 1) & tmpArr(2)).Font.Underline = xlUnderlineStyleNone
                        Range(Chr(Asc(tmpArr(1)) + 1) & tmpArr(2)).Font.Bold = False
                        End If
            Case Else
                Range(s).FormulaLocal = val_arr(i)
        End Select
    Next i
    s = Replace(Chr(Asc(tmpArr(1))) & (tmpArr(2) + 1), " ", "")
    Range(s).Activate
Application.ScreenUpdating = True
End Sub

Private Sub ComboBox10_Change()

Select Case ComboBox1.Value
    Case "Раздел"
        ComboBox15.Value = "ч"
    Case "Сметы"
        ComboBox15.Value = "вр"
    Case Else
        ComboBox15.Value = "=R[-1]C+1"
End Select
ComboBox10.ControlTipText = ComboBox10.Value
End Sub

Private Sub ComboBox11_Change()
    ComboBox5_Change
End Sub

Private Sub ComboBox12_Change()
ComboBox9_Change
Select Case ComboBox12.Value
    Case "1000"
        ComboBox10.Value = ComboBox10.List(0)
        ComboBox8.Value = "м"
    Case Else
        ComboBox10.Value = ComboBox10.List(0) & " L=" & ComboBox12.Value 'ComboBox10.List(ComboBox3.ListIndex) & " L=" & ComboBox12.Value
        ComboBox8.Value = "шт."
End Select

End Sub
Private Sub ComboBox13_Change()
    ComboBox9_Change
End Sub

Private Sub ComboBox5_Change()
Dim i As Integer
Dim s As String
i = InStr(ComboBox4, "Обозначение_")
If i > 0 Then
    s = Mid(ComboBox4, 13)
    Select Case ComboBox3.Value 'Заполнение типа, марки обозначения документа
        Case "Шина медная", "Лист стальной", "Полоса стальная" 'При форме "Обозначение" х "Ширина" "Нормативный документ"
            ComboBox17.Value = s & "x" & ComboBox13.Value & " " & ComboBox5.Value
           
'        Case "Арматура A-I", "Рукав гибкий металлический", "Двутавр", "Двутавр нормальный", _
'        "Двутавр широкополочный", "Прокат круглый", "Труба  сварная квадратная", "Труба  сварная прямоугольная", _
'        "Труба стальная водогазопроводная", "Уголок неравнополочный", "Уголок равнополочный", _
'        "Швеллер гнутый равнополочный", "Швеллер с наклонными полками", "Швеллер с параллельными полк.", _
'        "Шина алюминиевая"
'            ComboBox17.Value = s & " " & ComboBox5.Value
'            'При форме "Обозначение" "Нормативный документ"
        Case Else
            ComboBox17.Value = s & " " & ComboBox5.Value
    End Select
Else
    ComboBox17.Value = ComboBox4.Value & " " & ComboBox5.Value
End If
End Sub

Private Sub ComboBox9_Change()
    Скрыть_Лишнее
End Sub

Private Sub K_T_Change() 'В случае добавления колонок подправить функцию
If K_T = True Then
    ThisWorkbook.Worksheets("База_СО").Range("V20") = ComboBox3.Value
    ComboBox9.Value = ThisWorkbook.Worksheets("База_СО").Range("W20") + ComboBox9.List(0)
    ComboBox10.Value = "Болт с гайкой и двумя шайбами"
    ComboBox8.Value = "к-т"
Else
    ComboBox9.Value = ComboBox9.List(0)
    ComboBox10.Value = ComboBox10.List(0)
    ComboBox8.Value = ComboBox8.List(0)
End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If TheClosedBook Then
    Cancel = False
    TheClosedBook = False
Else
    Cancel = True
    SO_Zapolnen.Hide
End If
End Sub


