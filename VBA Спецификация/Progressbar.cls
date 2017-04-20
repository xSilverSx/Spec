VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private timeStart As Date
Private allProcess As Long
Private UpdInterval As Long
Private UpdTimeInterval As Long ' в секундах
Private prevUpdate As Double
Private prev As Long
Private formProgress As New barForm

Private Sub Class_Initialize()
    Load formProgress
    prev = 0
    prevUpdate = 0
    UpdInterval = Empty
    UpdTimeInterval = Empty
End Sub
Sub Start(Optional title As String = "Процесс выполнения")
    timeStart = Now
    formProgress.Caption = title
    formProgress.Show False
End Sub
Sub exitBar()
    Unload formProgress
End Sub
Sub createString()
' создание элемента формы для сообщения текстовой информации пользователю
' перебираем элементы формы для изменения параметров
t = 5
h = 49
For Each cntr In formProgress.Controls

    ' если элемент уже добавлялся, то прекращаем
    If cntr.Name = "String" Then
        MsgBox "Элемент уже был создан!", vbCritical, "Ошибка"
        Exit Sub
        
    ' проверяем на наличие других элементов
    ElseIf TypeName(cntr) = "TextBox" Then
        cntr.Top = cntr.Top + 20
        h = h + 55
    ElseIf cntr.Name = "bar" Then
        h = h + 20
        t = t + 20
    ElseIf cntr.Name = "Duration" Or cntr.Name = "Finish" Then
        cntr.Top = cntr.Top + 20
        h = h + 20
    End If
Next cntr

formProgress.Width = 340
formProgress.Height = h

' создаем lable с пустым текстом
    Set vCommand = formProgress.Controls.Add("Forms.Label.1")
    With vCommand
        .Left = 10
        .Top = t
        .Width = 325
        .Height = 15
        .Name = "String"
        .Caption = ""
    End With

End Sub
Sub createTextBox()
' создание элемента формы для множественных сообщений пользователю
' перебираем элементы формы для изменения параметров
t = 5
h = 84

For Each cntr In formProgress.Controls

    ' если элемент уже добавлялся, то прекращаем
    If TypeName(cntr) = "TextBox" Then
        MsgBox "Элемент уже был создан!", vbCritical, "Ошибка"
        Exit Sub
        
    ' проверяем на наличие других элементов
    ElseIf cntr.Name = "String" Then
        t = t + 20
        h = h + 20
    ElseIf cntr.Name = "bar" Or cntr.Name = "Duration" Or cntr.Name = "Finish" Then
        h = h + 20
        t = t + 20
    End If
Next cntr

formProgress.Width = 340
formProgress.Height = h

' создаем TextBox с пустым текстом
    Set vCommand = formProgress.Controls.Add("Forms.TextBox.1")
    With vCommand
        .Left = 5
        .Top = t
        .Width = 325
        .Height = 50
        .Name = "Text"
        .MultiLine = True
        .WordWrap = True
    End With

End Sub
Sub createtimeDuration()
' создание элемента формы для сообщения продолжительности процесса пользователю
' перебираем элементы формы для изменения параметров
t = 5
h = 49
For Each cntr In formProgress.Controls

    ' если элемент уже добавлялся, то прекращаем
    If cntr.Name = "Duration" Then
        MsgBox "Элемент уже был создан!", vbCritical, "Ошибка"
        Exit Sub
        
    ' проверяем на наличие других элементов
    ElseIf TypeName(cntr) = "TextBox" Then
        cntr.Top = cntr.Top + 20
        h = h + 55
    ElseIf cntr.Name = "bar" Or cntr.Name = "String" Then
        h = h + 20
        t = t + 20
    ElseIf cntr.Name = "Finish" Then
        h = h + 20
        cntr.Top = cntr.Top + 20
    End If
Next cntr

formProgress.Width = 340
formProgress.Height = h

' создаем lable с пустым текстом
    Set vCommand = formProgress.Controls.Add("Forms.Label.1")
    With vCommand
        .Left = 10
        .Top = t
        .Width = 325
        .Height = 15
        .Name = "Duration"
        .Caption = ""
    End With

End Sub
Sub createtimeFinish()
' создание элемента формы для сообщения оставшегося времени процесса пользователю
' перебираем элементы формы для изменения параметров
t = 5
h = 49
For Each cntr In formProgress.Controls

    ' если элемент уже добавлялся, то прекращаем
    If cntr.Name = "Finish" Then
        MsgBox "Элемент уже был создан!", vbCritical, "Ошибка"
        Exit Sub
        
    ' проверяем на наличие других элементов
    ElseIf TypeName(cntr) = "TextBox" Then
        cntr.Top = cntr.Top + 20
        h = h + 55
    ElseIf cntr.Name = "bar" Or cntr.Name = "String" Or cntr.Name = "Duration" Then
        h = h + 20
        t = t + 20
    End If
Next cntr

formProgress.Width = 340
formProgress.Height = h

' создаем lable с пустым текстом
    Set vCommand = formProgress.Controls.Add("Forms.Label.1")
    With vCommand
        .Left = 10
        .Top = t
        .Width = 325
        .Height = 15
        .Name = "Finish"
        .Caption = ""
    End With

End Sub
Sub createLoadingBar()
' создание элемента формы для полосы загрузки
' перебираем элементы формы для изменения параметров
h = 49

For Each cntr In formProgress.Controls

    ' если элемент уже добавлялся, то прекращаем
    If cntr.Name = "container" Or cntr.Name = "bar" Then
        MsgBox "Элемент loadingBar уже был создан!", vbCritical, "Ошибка"
        Exit Sub
    
    ' проверяем на наличие других элементов
    ElseIf TypeName(cntr) = "TextBox" Then
        cntr.Top = cntr.Top + 20
        h = h + 55
    ElseIf cntr.Name = "Finish" Or cntr.Name = "Duration" Or cntr.Name = "String" Then
        cntr.Top = cntr.Top + 20
        h = h + 20
    End If
Next cntr

formProgress.Width = 340
formProgress.Height = h

' создаем lable с пустым текстом
    Set vCommand = formProgress.Controls.Add("Forms.Label.1")
    With vCommand
        .Left = 5
        .Top = 5
        .Width = 325
        .Height = 15
        .Name = "container"
        .Caption = ""
        .TabIndex = 0
        .BackColor = &H80000004
    End With
Set vCommand = Nothing
    Set vCommand = formProgress.Controls.Add("Forms.Label.1")
    With vCommand
        .Left = 5
        .Top = 5
        .Width = 0
        .Height = 15
        .Name = "bar"
        .Caption = ""
        .TabIndex = 1
        .BackColor = &H80000002
    End With

End Sub
Sub setParameters(expProcess_INT As Long, Optional UpdateInterval_INT As Long, Optional UpdTimeInterval_INT_SEC As Long)
    ' устанавливаем ожидаемое число процессов для прогресс бара
    allProcess = expProcess_INT
    If Not IsMissing(UpdateInterval_INT) And UpdateInterval_INT > 0 Then
        UpdInterval = UpdateInterval_INT
    ElseIf Not IsMissing(UpdTimeInterval_INT_SEC) And UpdTimeInterval_INT_SEC > 0 Then
                                                                    UpdTimeInterval = UpdTimeInterval_INT_SEC
    Else: UpdTimeInterval = 1
    End If
End Sub
Function Update(Optional ByVal curProcess As Long, Optional ByVal stringTextBox As String)
If formProgress.Controls.Count = 0 Then Exit Function
' Проверяем, прошел ли интервал времени. Если нет, то не проводим обновление
timeDur = Now - timeStart
If UpdInterval > 0 And (curProcess - prevUpdate) >= UpdInterval Then
    prevUpdate = curProcess
ElseIf UpdTimeInterval > 0 And (Round(timeDur / (1 / 24 / 60 / 60), 0) - prevUpdate) >= UpdTimeInterval Then
    prevUpdate = Round(timeDur / (1 / 24 / 60 / 60), 0)
Else
    Exit Function
End If

' обновление формы
For Each cntl In formProgress.Controls
    Select Case cntl.Name
    Case "bar"
        If Not IsMissing(curProcess) Then
        ' опрежеляем процент выполнения от общего числа процессов
            rat = curProcess / allProcess
        ' устанавливаем ширину элемента для полосы загрузки в определенном проценте от общей ширины для полосы загрузки
            cntl.Width = Round(325 * rat, 2)
        End If
    Case "String"
    ' задаем строку сколько этапов процесса обработано
        If Not IsMissing(curProcess) Then cntl.Caption = "Обработано: " & curProcess & " из " & allProcess & " - " _
                                                            & Round(curProcess / allProcess * 100, 2) & "%"
    Case "Duration"
    ' определяет продолжительность процесса и переводим строку в виде времени
        h = WorksheetFunction.RoundDown((Now - timeStart) / (1 / 24), 0)
        If Len(Trim(CStr(h))) = 1 Then h = "0" & Trim(CStr(h))
        
        m = WorksheetFunction.RoundDown(((Now - timeStart) / (1 / 24 / 60)) - h * 60, 0)
        If Len(Trim(CStr(m))) = 1 Then m = "0" & Trim(CStr(m))
        
        s = Round((Now - timeStart) / (1 / 24 / 60 / 60) - h * 60 * 60 - m * 60, 0)
        If Len(Trim(CStr(s))) = 1 Then s = "0" & Trim(CStr(s))
        
        timeStr = h & ":" & m & ":" & s
        
        cntl.Caption = "Продолжительность обработки: " & timeStr
    Case "Finish"
    ' определяем оставшееся время исходя из того, сколько прошло до текущего этапа процесса, и сколько этапов осталось
        If Not IsMissing(curProcess) Then
        
                timeOneProcess = timeDur / curProcess ' времени на один этап процесса
                
                remTime = timeOneProcess * (allProcess - curProcess) ' оставшееся время
                
                h = WorksheetFunction.RoundDown(remTime / (1 / 24), 0)
                If Len(Trim(CStr(h))) = 1 Then h = "0" & Trim(CStr(h))
            
                m = WorksheetFunction.RoundDown((remTime / (1 / 24 / 60)) - h * 60, 0)
                If Len(Trim(CStr(m))) = 1 Then m = "0" & Trim(CStr(m))
            
                s = Round(remTime / (1 / 24 / 60 / 60) - h * 60 * 60 - m * 60, 0)
                If Len(Trim(CStr(s))) = 1 Then s = "0" & Trim(CStr(s))
                
                remTime = h & ":" & m & ":" & s
                
                cntl.Caption = "Оставшееся время обработки: " & remTime
        End If
    Case "Text"
    ' добавляем строку в textbox с переносом, если там уже что-то есть
        If Not IsMissing(stringTextBox) Then
            If cntl.TextLength = 0 Then
                cntl.Text = stringTextBox
            Else
                cntl.Text = stringTextBox & Chr(13) & cntl.Text
            End If
        End If
    End Select
Next cntl

formProgress.Repaint
DoEvents

End Function
Function getForm()
Set getForm = formProgress
End Function

