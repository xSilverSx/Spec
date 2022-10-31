Sub НомерДобавить()
Dim StrFolder As String 'Путь расположения надстройки
StrFolder = ThisWorkbook.Path
    If FileLocation(StrFolder & "\NumberSpace.png") Then
        ActiveSheet.PageSetup.RightHeaderPicture.filename = _
            StrFolder & "\NumberSpace.png"
        ActiveSheet.PageSetup.FirstPage.RightHeader.Picture.filename = _
            StrFolder & "\NumberSpace.png"
        Application.PrintCommunication = False
        
        With ActiveSheet.PageSetup
            .RightHeader = "&P&G"
            .FirstPage.RightHeader.Text = "&P&G"
        End With
        Application.PrintCommunication = True
        
        Call НомерРамка
    Else
        MsgBox "Номер не добавлен" & Chr(13) & "Не найден файл \NumberSpace.png" & _
        Chr(13) & "Файлы можно найти по ссылке https://github.com/xSilverSx/Spec", vbCritical
    End If
End Sub
Sub НомерРамка()

    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 1110, _
        0.5, 1110, 18).Select
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Weight = 1.25
    End With
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 1110, _
        18, 1146, 18).Select
    Selection.ShapeRange.ShapeStyle = msoLineStylePreset1
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Weight = 1.25
    End With

End Sub


