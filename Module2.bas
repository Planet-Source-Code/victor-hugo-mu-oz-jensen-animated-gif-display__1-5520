Attribute VB_Name = "Module2"
Option Explicit
#If Win16 Then
    Type RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type
#Else
    Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
#End If
#If Win16 Then
    Declare Sub GetWindowRect Lib "User" (ByVal hwnd As Integer, lpRect As RECT)
    Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
    Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer
    Declare Sub SetBkColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
    Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
    Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
    Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

Sub ExplodeForm(f As Form, Movement As Integer)
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, Cx%, Cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
    For i = 1 To Movement
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - Cx) / 2
        Y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, X, Y, X + Cx, Y + Cy
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
    
End Sub


Public Sub ImplodeForm(f As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'****************************************************************
'*Author: Carl Slutter
'*
'*Description:
'*The larger the "Movement" value, the slower the "Implosion"
'*
'*Creation Date: Thursday  23 January 1997  2:42 pm
'*Revision Date: Thursday  23 January 1997  2:42 pm
'*
'*Version Number: 1.00
'****************************************************************
    
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, Cx%, Cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
        For i = Movement To 1 Step -1
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - Cx) / 2
        Y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, X, Y, X + Cx, Y + Cy
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub

