VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form VbScroller 
   Caption         =   "VbScroller"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   3735
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   5535
   End
   Begin VB.PictureBox PicBgColor 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox PicTextColor 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin MSComDlg.CommonDialog ColorDlg 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicColor 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   5280
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   5280
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scroll it!"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   120
      MouseIcon       =   "Form1.frx":08CA
      ScaleHeight     =   2115
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2055
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3625
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         MousePointer    =   4
         Appearance      =   0
         TextRTF         =   $"Form1.frx":1194
         MouseIcon       =   "Form1.frx":1218
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Background color"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Text color"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Hyperlink color"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "VbScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const TVM_SETBKCOLOR = 4381&
Private Const EM_CHARFROMPOS& = &HD7
Dim RtbPosY As Integer
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
As Long) As Long

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As _
Any) As Long
Private hyperlink As String
Private paused As Boolean
Private UnderLineColor As OLE_COLOR
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub



Private Sub Form_Load()
paused = False
Text1.Text = Text1.Text & "This is just a test... www.jhajha.com" & vbCrLf
Text1.Text = Text1.Text & "I repeat; this is a test" & vbCrLf
Text1.Text = Text1.Text & "Visit www.winzip.com" & vbCrLf
Text1.Text = Text1.Text & "Yep. It works alright" & vbCrLf
Text1.Text = Text1.Text & "Hmmm.. yadayada..." & vbCrLf
Text1.Text = Text1.Text & "Check out www.pscode.com now!" & vbCrLf
Text1.Text = Text1.Text & "Blabla whatever... hmmmm" & vbCrLf
Text1.Text = Text1.Text & "mailto:danne.r@aland.net" & vbCrLf
Text1.Text = Text1.Text & "Advertiseing is kewl :)" & vbCrLf
Text1.Text = Text1.Text & "Well, bye now.." & vbCrLf
rtb.Text = Text1.Text
rtb.SelStart = 0
rtb.SelLength = Len(rtb.Text)
rtb.SelColor = vbWhite
rtb.SelLength = 0
' = ""
rtb.Top = Picture1.ScaleHeight + 300
RtbPosY = rtb.Top
UnderLineColor = vbRed
UnderlineHyperlink
End Sub



Private Sub PicBgColor_Click()
On Error GoTo trap
With ColorDlg
    .CancelError = True
    .ShowColor
    Picture1.BackColor = .Color
    PicBgColor.BackColor = .Color
    rtb.BackColor = .Color
End With
trap:
End Sub

Private Sub PicColor_Click()
On Error GoTo trap
With ColorDlg
    .CancelError = True
    .ShowColor
    PicColor.BackColor = .Color
    UnderLineColor = .Color
    UnderlineHyperlink
End With
trap:
End Sub

Private Sub PicTextColor_Click()
On Error GoTo trap
With ColorDlg
    .CancelError = True
    .ShowColor
    PicTextColor.BackColor = .Color
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    rtb.SelColor = .Color
    rtb.SelLength = 0
    UnderlineHyperlink
End With
trap:
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
hyperlink = CheckHyperlink(x, y)
End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    If Len(hyperlink) > 0 Then
        ShellExecute Me.hwnd, "Open", hyperlink, vbNullString, vbNullString, vbShow
        paused = False
    End If
End If
End Sub

Private Sub Text1_Change()
rtb.Text = Text1.Text
End Sub

Private Sub Timer1_Timer()
If paused = False Then
RtbPosY = RtbPosY - 15: If RtbPosY < (-rtb.Height) Then RtbPosY = Picture1.Height
rtb.Top = RtbPosY
End If
End Sub

Private Sub UnderlineHyperlink()
On Error Resume Next
Dim pos As Long
Dim posEnd As Long
Dim char As String
Dim link As String
pos = InStr(1, LCase$(rtb.Text), "mailto:")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 7 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = UnderLineColor
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "ftp://")
Loop
pos = InStr(1, LCase$(rtb.Text), "ftp://")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 6 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = UnderLineColor
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "ftp://")
Loop
pos = InStr(1, LCase$(rtb.Text), "http://")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 7 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = UnderLineColor
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "http://")
Loop
pos = InStr(1, LCase$(rtb.Text), "www.")

Do While pos > 0

For posEnd = pos To Len(rtb.Text)
char = Mid$(rtb.Text, posEnd, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next posEnd
link = Mid$(rtb.Text, pos, posEnd - pos)
char = Right$(link, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?" Or Len(char) <> 1
link = Left$(link, Len(link) - 1)
char = Right$(link, 1)
Loop

If Len(link) > 4 Then
rtb.SelStart = pos - 1
rtb.SelLength = Len(link)
rtb.SelUnderline = True
rtb.SelColor = UnderLineColor
End If
pos = InStr(posEnd + 1, LCase$(rtb.Text), "www.")
Loop
rtb.SelStart = Len(rtb.Text)
End Sub

Public Function CheckHyperlink(x As Single, y As Single) As String
Dim point As POINTAPI
Dim charpos As Long
Dim pos_start As Long
Dim pos_end As Long
Dim char As String
Dim word As String
point.x = x \ Screen.TwipsPerPixelX
point.y = y \ Screen.TwipsPerPixelY
charpos = SendMessage(rtb.hwnd, EM_CHARFROMPOS, 0&, point)

If charpos <= 0 Or charpos = Len(rtb.Text) Then
rtb.MousePointer = 1
paused = False
CheckHyperlink = vbNullString
Exit Function
End If

For pos_start = charpos To 1 Step -1

If Mid$(rtb.Text, pos_start + 1, 1) = Chr$(13) Then
rtb.MousePointer = 1
paused = False
CheckHyperlink = vbNullString
Exit Function
End If
char = Mid$(rtb.Text, pos_start, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next pos_start
pos_start = pos_start + 1

For pos_end = charpos To Len(rtb.Text)
char = Mid$(rtb.Text, pos_end, 1)
If char = Chr$(32) Or char = Chr$(10) Or char = Chr$(13) Then Exit For
Next pos_end
pos_end = pos_end - 1
If pos_start <= pos_end Then word = LCase$(Mid$(rtb.Text, pos_start, pos_end - pos_start + 1))
If Left$(word, 7) = "http://" Or Left$(word, 4) = "www." Or Left$(word, 6) = "ftp://" Or Left$(word, 7) = "mailto:" Then

char = Right$(word, 1)

Do While char = "." Or char = "," Or char = "!" Or char = "?"
If Len(char) = 0 Then Exit Do
word = Left$(word, Len(word) - 1)
char = Right$(word, 1)
Loop

If Len(word) < 4 Then
rtb.MousePointer = 1
paused = False
CheckHyperlink = vbNullString
Else
rtb.MousePointer = 99
paused = True
CheckHyperlink = word
End If
Else
rtb.MousePointer = 1
paused = False
End If
End Function


