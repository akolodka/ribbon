Attribute VB_Name = "RibbonCallbacks"
Option Explicit '����������� ������ ���������� ���� ���������� � �����

'FirstButton (�������: button, �������: onAction), 2010+
Private Sub OnFirstRibbonButtonPressed(ByVal control As IRibbonControl)
    MsgBox "��������� ���������, �������� � onAction �������� " & control.ID
End Sub

'SecondButton (�������: button, �������: onAction), 2010+
Private Sub OnSecondRibbonButtonPressed(ByVal control As IRibbonControl)
    MsgBox "��������� ���������, �������� � onAction �������� " & control.ID
End Sub

'ThirdButton (�������: button, �������: onAction), 2010+
Private Sub OnThirdRibbonButtonPressed(ByVal control As IRibbonControl)
    MsgBox "��������� ���������, �������� � onAction �������� " & control.ID
End Sub

