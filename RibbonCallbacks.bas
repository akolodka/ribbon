Attribute VB_Name = "RibbonCallbacks"
Option Explicit 'Потребовать явного объявления всех переменных в файле

'FirstButton (элемент: button, атрибут: onAction), 2010+
Private Sub OnFirstRibbonButtonPressed(ByVal control As IRibbonControl)
    MsgBox "Сработала процедура, заданная в onAction элемента " & control.ID
End Sub

'SecondButton (элемент: button, атрибут: onAction), 2010+
Private Sub OnSecondRibbonButtonPressed(ByVal control As IRibbonControl)
    MsgBox "Сработала процедура, заданная в onAction элемента " & control.ID
End Sub

'ThirdButton (элемент: button, атрибут: onAction), 2010+
Private Sub OnThirdRibbonButtonPressed(ByVal control As IRibbonControl)
    MsgBox "Сработала процедура, заданная в onAction элемента " & control.ID
End Sub

