Attribute VB_Name = "ButtonFlat"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000

Public Function btnFlat(Button As CommandButton)
    SetWindowLong Button.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
