Attribute VB_Name = "KeyStates"
Option Explicit

Declare PtrSafe Function GetKeyState Lib "User32" (ByVal vKey As Integer) As Integer

Const SHIFT_KEY = 16
Const CTRL_KEY = 17
Const ALT_KEY = 18

Function IsShiftPressed() As Boolean
    IsShiftPressed = GetKeyState(SHIFT_KEY) And &H8000
End Function

Function IsCtrlPressed() As Boolean
    IsCtrlPressed = GetKeyState(CTRL_KEY) And &H8000
End Function

Function IsAltPressed() As Boolean
    IsAltPressed = GetKeyState(ALT_KEY) And &H8000
End Function

