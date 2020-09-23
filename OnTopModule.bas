Attribute VB_Name = "OnTopModule"
Option Explicit

' Declare the API function to put a window on top
    Declare Sub SetWindowPos Lib "User32" _
    (ByVal hWnd As Integer, _
    ByVal hWndInsertAfter As Integer, _
    ByVal X As Integer, _
    ByVal Y As Integer, _
    ByVal cx As Integer, _
    ByVal cy As Integer, _
    ByVal wFlags As Integer)

' Declare constants
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const HWND_TOPMOST = -1
Global ckIt As Boolean
Public Sub modPutWindowOnTop()
Dim Para
' Display the form
    frmPrintLandscape.Show
    'Make the ON TOP call
Para = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    'Specify the Form/Window name (TDWin)
    SetWindowPos frmPrintLandscape.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Para

End Sub

