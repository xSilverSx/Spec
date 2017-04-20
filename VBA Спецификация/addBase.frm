VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addBase 
   Caption         =   "Добавить позицию в базу"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   7455
   OleObjectBlob   =   "addBase.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "addBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub Button1_Click()
'On Error GoTo err_ch
Dim i As Integer, l As Integer, k As Integer
Dim s, adDr, tmpArr, AddrTmp As String
Dim val_arr(1 To 7)
'Application.ScreenUpdating = False
If ActiveCell.Row = 1 Then Exit Sub 'Условие не вставки в первую ячейку
adDr = ActiveCell.Address 'Получение адреса ячейки
tmpArr = Split(adDr, "$") 'Разбиение на массив строки и столбца
AddrTmp = tmpArr(1) & tmpArr(2)
tmpArr(1) = "A"
s = Replace("A" & (tmpArr(2) + 1), " ", "") 'Адрес ячейки на которую необходимо будет перейти
For i = 1 To 7
s = Replace(Chr(Asc(tmpArr(1)) + i) & tmpArr(2), " ", "")
Range(s).Select
val_arr(i) = Range(s).Value
'MsgBox val_arr(i)

Next i
    ComboBox10.Value = val_arr(1)    'Наименование
    ComboBox3.Value = val_arr(1)
    ComboBox17.Value = val_arr(2)     'Обозначение
    ComboBox6.Value = val_arr(3)      'Код
    ComboBox7.Value = val_arr(4)      'Завод
    ComboBox8.Value = val_arr(5)      'Единица измерения
    ComboBox11.Value = val_arr(7)     'Масса единицы
l = InStr(1, val_arr(2), " ")
ComboBox4.Value = Left(val_arr(2), l)
ComboBox5.Value = Mid(val_arr(2), l + 1)
Range(adDr).Activate
'TextBox5.Value = Prefix_N_Opor(TextBox5.Value)
'Application.ScreenUpdating = True
'Exit Sub
'err_ch:
'MsgBox "Ошибка N " & Err.Number & vbCrLf & Err.Description
'Application.ScreenUpdating = True
End Sub

Private Sub Button2_Click()
Dim i As Integer, l As Integer, k As Integer
Dim s, adDr, AddrTmp As String
Dim val_arr(0 To 9)
Dim arr
Dim tmpArr(1 To 2)
With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(1, 1), .Cells(Rows.Count, "A").End(xlUp)).Value
End With
'MsgBox UBound(arr)
    val_arr(0) = ComboBox1.Value   'Категория
    val_arr(1) = ComboBox2.Value   'Подкатегория
    val_arr(2) = ComboBox3.Value   'Краткое наименование
    val_arr(3) = ComboBox10.Value    'Полное наименование
    val_arr(4) = ComboBox4.Value    'Тип (обозначение)
    val_arr(5) = ComboBox5.Value    'Нормативный документ
    val_arr(6) = ComboBox6.Value   'Код оборудования
    val_arr(7) = ComboBox7.Value   'Завод
    val_arr(8) = ComboBox8.Value   'Единица измерения
    val_arr(9) = ComboBox11.Value   'Масса

   
tmpArr(1) = "A"
tmpArr(2) = UBound(arr) + 1
s = tmpArr(1) & tmpArr(2) + 1 'Адрес ячейки на которую необходимо будет перейти
For i = 0 To 7
If i = 7 Then
s = Chr(Asc(tmpArr(1)) + i + 5) & tmpArr(2)
ThisWorkbook.Sheets("База_СО").Range(s).FormulaLocal = "Нов."
End If
s = Chr(Asc(tmpArr(1)) + i) & tmpArr(2)
ThisWorkbook.Sheets("База_СО").Range(s).FormulaLocal = val_arr(i)
Next i

End Sub


Private Sub UserForm_Initialize() 'Выбор ячейки категория
Dim a As Integer, i As Integer
Dim arr
With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 1), .Cells(Rows.Count, "A").End(xlUp)).Value
End With
    With CreateObject("Scripting.Dictionary")
  
        For i = LBound(arr) To UBound(arr)
            .Item(arr(i, 1)) = 1
        Next

        Me.ComboBox1.List = .keys
        ComboBox1.Value = ComboBox1.List(0)
End With
With ThisWorkbook.Sheets("База_СО")
    arr = .Range(.Cells(2, 2), .Cells(Rows.Count, "B").End(xlUp)).Value
End With
    With CreateObject("Scripting.Dictionary")
  
        For i = LBound(arr) To UBound(arr)
            .Item(arr(i, 1)) = 1
        Next

        Me.ComboBox2.List = .keys
        ComboBox2.Value = ComboBox2.List(0)
End With
Button1_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If TheClosedBook Then
    Cancel = False
    TheClosedBook = False
Else
    Cancel = True
    Me.Hide
End If
End Sub
