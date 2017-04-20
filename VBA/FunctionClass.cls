VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Function FindTextInString(FindindText As String, Left_Mid_Right_1_2_3 As String, _
            First_Symb As String, Optional Second_Symb As String) As String
'Поиск слова в строке
Dim intSymbOne As Integer, intLengText As Integer
Dim intSymbTwo As Integer
'Left_Mid_Right_1_2_3 = LCase(Left_Mid_Right_1_2_3) 'Перевод знаков в нижний регистр
Select Case Left_Mid_Right_1_2_3 ' Выбираем для какого случая делаем
    Case 1
        intSymbOne = InStr(1, FindindText, First_Symb) 'Поиск порядка нужного символа в нашей строке
        FindTextInString = Left(FindindText, intSymbOne - 1) 'Вывод строки
    Case 2
        intLengText = Len(FindindText) 'Определяем длину строки
        intSymbOne = InStrRev(FindindText, First_Symb, intLengText)
        intSymbTwo = InStrRev(FindindText, Second_Symb, intLengText)
        FindTextInString = Mid(FindindText, intSymbOne + 1, intSymbTwo - intSymbOne - 1)
    Case 3
        intLengText = Len(FindindText) 'Определяем длину строки
        intSymbOne = InStrRev(FindindText, First_Symb, intLengText) 'Поиск порядка нужного символа в нашей строке
        FindTextInString = Right(FindindText, intLengText - intSymbOne) 'Вывод строки
    Case Else
        MsgBox "Не верное значение", vbCritical
End Select
End Function
