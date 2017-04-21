Attribute VB_Name = "Functions"

Function Заменить(sWhat As String, sReplacement As String, Целиком As Boolean, rRange As Range)
'Что меняем, На что меняем, Ячейка целиком или часть текста, Диапазон
    If Целиком Then
        rRange.Replace What:=sWhat, Replacement:=sReplacement, LookAt:=xlWhole
    Else
        rRange.Replace What:=sWhat, Replacement:=sReplacement, LookAt:=xlPart
    End If
End Function

Function Вы_уверены()
Dim intKey As Integer

intKey = MsgBox("Вы уверены что хотите удалить последний лист (это необратимо)", _
                vbQuestion + vbYesNo + vbDefaultButton2)
              
If intKey = vbYes Then
Удалить_Лист_КЖ
End If

End Function

Sub Скрыть(YesNo As Boolean)
If YesNo Then
    Sheets("Сводная ведомость A3").Rows("1:2").EntireRow.Hidden = True
    Sheets("КЖ").Columns("A:A").EntireColumn.Hidden = True
Else
    Sheets("Сводная ведомость A3").Rows("1:2").EntireRow.Hidden = False
    Sheets("КЖ").Columns("A:A").EntireColumn.Hidden = False
End If
End Sub

'Function SpecialFolderPath() As String 'определяет путь рабочего стола
'    Dim objWSHShell As Object
'    Dim strSpecialFolderPath
'    Dim strSpecialFolder
'
'    Set objWSHShell = CreateObject("WScript.Shell")
'    SpecialFolderPath = objWSHShell.SpecialFolders("Desktop")
'    Set objWSHShell = Nothing
'    Exit Function
'ErrorHandler:
'     MsgBox "Error finding " & strSpecialFolder, vbCritical + vbOKOnly, "Error"
'End Function

