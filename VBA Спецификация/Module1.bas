Attribute VB_Name = "Module1"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit
Public TheClosedBook As Boolean

Sub Main()
    ActiveWorkbook.Worksheets("Спецификация").Activate
        If IsWorkSheetExistXLAM("База_СО") = False Then
            MsgBox "База данных не установлена!", vbCritical
        Else
            SO_Zapolnen.Show
        End If
End Sub

Sub addInTheBase()
    ActiveWorkbook.Worksheets("Спецификация").Activate
    addBase.Show
End Sub

Sub Выгрузить_Форму()
    TheClosedBook = True
    Unload VBAProjectSO.SO_Zapolnen
    Unload VBAProjectSO.addBase
End Sub

Sub Редактор_Книги() 'Делает книгу доступной для редактрования
    If ThisWorkbook.IsAddin = False Then
    ThisWorkbook.IsAddin = True
    Exit Sub
    End If
    If ThisWorkbook.IsAddin = True Then ThisWorkbook.IsAddin = False
End Sub

Sub Сохранить_Книгу() 'Сохраняет книгу
Dim a As Byte
a = MsgBox("Действительно пересохранить файл Надстройки?", vbYesNo)
If a = vbYes Then ThisWorkbook.Save
End Sub

Sub Change_ReferenceStyle() 'Замена стилей R1C1
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub
