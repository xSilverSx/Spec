Attribute VB_Name = "RibbonCallbacks"
'Романов Владимир Анатольевич e-hoooo@yandex.ru 20/04/2016г.
Option Explicit 'Потребовать явного объявления всех переменных в файле

'ObnovSpec (компонент: button, атрибут: onAction), 2007
Sub ObnovSpec(control As IRibbonControl)
    Копировать_Листы
End Sub

'ObnovKnopok (компонент: button, атрибут: onAction), 2007
Sub ObnovKnopok(control As IRibbonControl)
    Создать_кнопки
End Sub

'NaPerenos (компонент: button, атрибут: onAction), 2007
Sub NaPerenos(control As IRibbonControl)
    Podgotovka_Show
End Sub

'SaveToday (компонент: button, атрибут: onAction), 2007
Sub SaveToday(control As IRibbonControl)
    Сохранить_Сегодня
End Sub

'RCChange (компонент: button, атрибут: onAction), 2007
Sub RCChange(control As IRibbonControl)
    Change_ReferenceStyle
End Sub

'Pos (компонент: button, атрибут: onAction), 2007
Sub Pos(control As IRibbonControl)
    Posicii
End Sub

'ClearVersion (компонент: button, атрибут: onAction), 2007
Sub ClearVersion(control As IRibbonControl)
    Оставить_одну_версию
End Sub

'CreateVersion (компонент: button, атрибут: onAction), 2007
Sub CreateVersion(control As IRibbonControl)
    Сохранить_Версию_Спецификации
End Sub

'ShowFormVersion (компонент: button, атрибут: onAction), 2007
Sub ShowFormVersion(control As IRibbonControl)
    Показать_Форму_Версии
End Sub

'DateVersion (компонент: button, атрибут: onAction), 2007
Sub DateVersion(control As IRibbonControl)
    Обновить_дату_последней_версии
End Sub

'Packet (компонент: button, атрибут: onAction), 2007
Sub Packet(control As IRibbonControl)
    PaketnayaObrabotka.Show
End Sub

'Locked (компонент: button, атрибут: onAction), 2007
Sub Locked(control As IRibbonControl)
    СнятьЗащиту
End Sub

'RedactBook (компонент: button, атрибут: onAction), 2007
Sub ReloadDB(control As IRibbonControl)
    ПереподключитьБазуДанных
End Sub

'UnLoadForm (компонент: button, атрибут: onAction), 2007
Sub UnLoadForm(control As IRibbonControl)
    Выгрузить_Форму
End Sub

'ДобавитьИзБазы (компонент: button, атрибут: onAction), 2007
Sub Добавить_Из_Базы(control As IRibbonControl)
    Main
End Sub

'SortBase (компонент: button, атрибут: onAction), 2007
Sub SortBase(control As IRibbonControl)
    Сортировка_Базы
End Sub

'OpenDataBase (компонент: button, атрибут: onAction), 2007
Sub OpenDataBase(control As IRibbonControl)
    ОткрытьБазуДанных
End Sub

''addInBase (компонент: button, атрибут: onAction), 2010/2013
'Sub addInBase(control As IRibbonControl)
'    addInTheBase
'End Sub

