VERSION 5.00
Begin VB.Form form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Animated Gif Demo Project by Hugo Muñoz"
   ClientHeight    =   4680
   ClientLeft      =   1065
   ClientTop       =   2055
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6390
   Begin VB.Timer AnimationTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "hugo@ecuabox.com"
      BeginProperty Font 
         Name            =   "Short Hand"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Image AnimatedGIF 
      Appearance      =   0  'Flat
      Height          =   900
      Index           =   0
      Left            =   480
      Top             =   840
      Width           =   7020
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RepeatTimes&
Dim RepeatCount&
Dim FrameCount&
Dim TotalFrames&

Private Sub Check1_Click()
If Check1 = Unchecked Then AOTValue = False
If Check1 = Checked Then AOTValue = True
If AOTValue = True Then
SetWindowPos form1.hwnd, HWND_TOPMOST, form1.Left / 15, form1.Top / 15, form1.Width / 15, form1.Height / 15, 3 Or 8 'AOTValue = True
Else
SetWindowPos form1.hwnd, HWND_NOTOPMOST, form1.Left / 15, form1.Top / 15, form1.Width / 15, form1.Height / 15, 3 Or 8
End If
End Sub

Private Sub Check2_Click()
Dim Vitesse
If Check2 = Unchecked Then Timer1.Enabled = False
If Check2 = Checked Then
Vitesse = InputBox("Which interval ( Milliseconds )", "Refresh", 1000, Screen.Width / 2, Screen.Height / 2)
If Vitesse = "" Then Vitesse = 1000
Timer1.Interval = Vitesse
Timer1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText List1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
List1.Clear
Call GetPasswords
End Sub

Private Sub Command4_Click()
CreateFile "Pass.txt"
End Sub

Private Sub Form_Load()

'Check1.Value = 1
Dim path As String
form1.MouseIcon = LoadPicture(App.path & "\hand.cur")
form1.MousePointer = vbDefault
path = App.path & "\anixw.gif"
Call LoadAniGif(path, AnimatedGIF)
End Sub

Private Sub Form_Unload(Cancel As Integer)

End
End Sub


Private Sub Timer1_Timer()
Command3_Click
End Sub

Sub LoadAniGif(xFile As String, xImgArray)

    If Not IIf(Dir$(xFile) = "", False, True) Or xFile = "" Then
        MsgBox "File not found.", vbExclamation, "gif not found"
        Exit Sub
    End If
        
    Dim F1, F2
    Dim AnimatedGIFs() As String
    Dim imgHeader As String
    Static buf$, picbuf$
    Dim fileHeader As String
    Dim imgCount
    Dim i&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd
    GifEnd = Chr(0) & "!ù"
    
    AnimationTimer.Enabled = False
    For i = 1 To xImgArray.Count - 1
        Unload xImgArray(i)
    Next i
    
    F1 = FreeFile
On Error GoTo badFile:
    Open xFile For Binary Access Read As F1
        buf = String(LOF(F1), Chr(0))
        Get #F1, , buf
    Close F1
    
    i = 1
    imgCount = 0
    
    j = (InStr(1, buf, GifEnd) + Len(GifEnd)) - 2
    fileHeader = Left(buf, j)
    i = j + 2
    
    If Len(fileHeader) >= 127 Then
        RepeatTimes& = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * CLng(256))
    Else
        RepeatTimes = 0
    End If


    Do
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + Len(GifEnd)
        If j > Len(GifEnd) Then
            F2 = FreeFile
            Open "tmp.gif" For Binary As F2
                picbuf = String(Len(fileHeader) + j - i, Chr(0))
                picbuf = fileHeader & Mid(buf, i - 1, j - i)
                Put #F2, 1, picbuf
                imgHeader = Left(Mid(buf, i - 1, j - i), 16)
            Close F2
            
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * CLng(256))) * CLng(10)
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * CLng(256))
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * CLng(256))
                Load xImgArray(imgCount - 1)
                xImgArray(imgCount - 1).ZOrder 0
                xImgArray(imgCount - 1).Left = xImgArray(0).Left + (xOff * CLng(15))
                xImgArray(imgCount - 1).Top = xImgArray(0).Top + (yOff * CLng(15))
            End If
            xImgArray(imgCount - 1).Tag = TimeWait
            xImgArray(imgCount - 1).Picture = LoadPicture("tmp.gif")
            Kill ("tmp.gif")
            
            i = j '+ 1
        End If
        DoEvents
    Loop Until j = Len(GifEnd)
    
    If i < Len(buf) Then
        F2 = FreeFile
        Open "tmp.gif" For Binary As F2
            picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
            picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
            Put #F2, 1, picbuf
            imgHeader = Left(Mid(buf, i - 1, Len(buf) - i), 16)
        Close F2

        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * CLng(256))) * CLng(10)
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * CLng(256))
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * CLng(256))
            Load xImgArray(imgCount - 1)
            xImgArray(imgCount - 1).ZOrder 0
            xImgArray(imgCount - 1).Left = xImgArray(0).Left + (xOff * CLng(15))
            xImgArray(imgCount - 1).Top = xImgArray(0).Top + (yOff * CLng(15))
        End If
        xImgArray(imgCount - 1).Tag = TimeWait
        xImgArray(imgCount - 1).Picture = LoadPicture("tmp.gif")
        Kill ("tmp.gif")
    End If
    
    FrameCount = 0
    TotalFrames = xImgArray.Count - 1
    
On Error GoTo badTime
    AnimationTimer.Interval = CInt(xImgArray(0).Tag)
badTime:
    AnimationTimer.Enabled = True
Exit Sub
badFile:
    MsgBox "File not found.", vbExclamation, "File Error"

End Sub

Private Sub AnimatedGIF_Click(Index As Integer)
MsgBox "Send me your comments at hugo@ecuabox.com", vbInformation
End Sub

Private Sub AnimatedGIF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
form1.MousePointer = 99
End Sub

Private Sub AnimationTimer_Timer()
    If FrameCount < TotalFrames Then
        FrameCount = FrameCount + 1
        AnimatedGIF(FrameCount).Visible = True
        AnimationTimer.Interval = CLng(AnimatedGIF(FrameCount).Tag)
    Else
        FrameCount = 0
        For i = 1 To AnimatedGIF.Count - 1
            AnimatedGIF(i).Visible = False
        Next i
        AnimationTimer.Interval = CLng(AnimatedGIF(FrameCount).Tag)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
form1.MousePointer = vbDefault
End Sub
