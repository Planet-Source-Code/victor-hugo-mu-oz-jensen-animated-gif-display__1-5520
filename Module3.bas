Attribute VB_Name = "Module3"
Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Global Const HWND_TOPMOST = -1
    Global Const HWND_NOTOPMOST = -2
    Global AOTValue As Boolean
