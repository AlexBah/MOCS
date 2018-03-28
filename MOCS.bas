' Create Button in menu following my macro in this workbook
Sub AddMyButton()
	Dim myControl As CommandBarButton
    Set myControl = CommandBars(1).Controls.Add(msoControlButton)
    With myControl
        .Caption = "CaptionMyMacro" 'Caption of button
		.OnAction = "GoClose"
        .FaceId = 16
        .Style = msoButtonCaption
    End With
	ThisWorkbook.Close
End Sub

Sub GoClose()
	MyMacro
	ThisWorkbook.Close 'Close this workbook after finishing macro
End Sub

Private Sub MyMacro()
	'Here my code
End Sub


