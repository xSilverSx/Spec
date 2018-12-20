Attribute VB_Name = "Error"
'Модуль обработки ошибок фактически не используется
Option Explicit
 
'Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub Error() 'Функция обработки ошибок
Dim e As String, e1 As String, e2 As String
Dim f As String

    e = Err.Description & " " & Err.Number
    e1 = "Время Ошибки: " & Now() & " " & e '& Chr(vbNewLine) &
    e2 = "Документ от: " & ThisWorkbook.BuiltinDocumentProperties(12).Value & _
    " " & ThisWorkbook.BuiltinDocumentProperties(32).Value '& Computer
    Debug.Print e1
    MsgBox "Ошибка: " & e
On Error Resume Next
f = FreeFile
Open "X:\ДЕПАРТАМЕНТ ПРОЕКТИРОВАНИЯ\!ОТДЕЛ ПРОЕКТИРОВАНИЯ ЭТО\Для работы ЭТО\Программы\DOSBoxPortable\log.txt" For Append As #f
Print #f, e1
Print #f, e2
Close #f
Err.Clear
End Sub


'Function Computer()
'Dim scomp As String
'Dim h As String
'scomp = Space(255)
'h = GetComputerName(scomp, 255)
'Computer = Trim(scomp)
'End Function



Sub nulll()
On Error GoTo Error:
Dim A As String
Dim x As Integer
A = "string"
x = A

Exit Sub
Error:
Call Error
End Sub




 
