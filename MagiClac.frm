VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "The Magic Calculator 1.0"
   ClientHeight    =   5250
   ClientLeft      =   3375
   ClientTop       =   2955
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MagiClac.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Tag             =   "80"
   Begin VB.PictureBox PicVoteLink 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      MouseIcon       =   "MagiClac.frx":B1E2
      MousePointer    =   99  'Custom
      ScaleHeight     =   255
      ScaleWidth      =   4215
      TabIndex        =   30
      Top             =   500
      Width           =   4215
   End
   Begin VB.Timer TmrLinkMouseOver 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   4680
   End
   Begin VB.CommandButton CmdBS 
      Caption         =   "< BS"
      Height          =   615
      Left            =   4095
      TabIndex        =   22
      Top             =   3255
      Width           =   975
   End
   Begin VB.CommandButton CmdPower 
      Caption         =   "^"
      Height          =   615
      Left            =   3120
      TabIndex        =   21
      Top             =   3255
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "+"
      Height          =   615
      Left            =   4095
      TabIndex        =   20
      Top             =   4515
      Width           =   975
   End
   Begin VB.CommandButton CmdSub 
      Caption         =   "---"
      Height          =   615
      Left            =   3120
      TabIndex        =   19
      Top             =   4515
      Width           =   975
   End
   Begin VB.CommandButton CmdMult 
      Caption         =   "*"
      Height          =   615
      Left            =   4095
      TabIndex        =   18
      Top             =   3885
      Width           =   975
   End
   Begin VB.CommandButton CmdDivide 
      Caption         =   "/"
      Height          =   615
      Left            =   3120
      TabIndex        =   17
      Top             =   3885
      Width           =   975
   End
   Begin VB.CommandButton CmdCP 
      Caption         =   ")"
      Height          =   615
      Left            =   4095
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdOP 
      Caption         =   "("
      Height          =   615
      Left            =   3120
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdDot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2085
      TabIndex        =   14
      Top             =   4515
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   2085
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1095
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   2085
      TabIndex        =   9
      Top             =   3255
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1095
      TabIndex        =   8
      Top             =   3255
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   3255
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2085
      TabIndex        =   12
      Top             =   3885
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1095
      TabIndex        =   11
      Top             =   3885
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3885
      Width           =   975
   End
   Begin VB.CommandButton CmdNum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   4515
      Width           =   1945
   End
   Begin VB.Timer TmrExpandMouseOver 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7200
      Tag             =   "Reset"
      Top             =   4680
   End
   Begin VB.PictureBox PicExpand 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   85
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   7770
      TabIndex        =   28
      Top             =   2430
      Width           =   7765
   End
   Begin RichTextLib.RichTextBox TxtExp 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   529
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"MagiClac.frx":B334
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
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   2700
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton CmdSolve 
      Caption         =   "&Solve"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"MagiClac.frx":B3AF
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1695
      Left            =   5280
      TabIndex        =   29
      Top             =   3000
      Width           =   2280
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      FillColor       =   &H00FFE0E0&
      Height          =   2730
      Left            =   15
      Top             =   2520
      Width           =   7815
   End
   Begin VB.Label LblExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   7460
      TabIndex        =   27
      Top             =   60
      Width           =   290
   End
   Begin VB.Label LblMin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   6970
      TabIndex        =   26
      Top             =   60
      Width           =   285
   End
   Begin VB.Label LblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   6520
      TabIndex        =   25
      Top             =   60
      Width           =   285
   End
   Begin VB.Label LblResult 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your expresion here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFE0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      FillColor       =   &H00E0E0E0&
      Height          =   2550
      Left            =   15
      Top             =   0
      Width           =   7815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      FillColor       =   &H00FFE0E0&
      FillStyle       =   0  'Solid
      Height          =   2510
      Left            =   5160
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed By : Yehia Muhsen
'Date          : 8-13-2003
'Copyrights    : Copyright © 2003 Yehia Muhsen, All rights reserved.
'Description   : This program is one of my best programs. Basically it's a calculator,
'                but it's not like any calculator. I called it the MagiCalc 1.0
'                because it's very powerful, and can accurately evaluate any mathematical
'                expression no matter how complicated it is as long as it meets conditions.

'                The operations allowed are *, / , + , - , and ^. You can use parenthesis too.
'                For example:+(2((5+4)/2*6+1.56845)4 + 3^(3-(--5)/2))/(25^.5)+2-3*4/-2^2.2*5

'                In addition to evaluating expressions, this program is supported by a very
'                powerful errors handling. It deals with almost all kind of errors,
'                and it tells you exaclty what and where the error is so that you can
'                fix it. One of the good things about this program is that it uses the
'                memory so effectively. The following expression has almost all kind of
'                errors that this program handles, use that expression to see how Syntax
'                errors get handled. ()+*5.334.2+)453+&76/6(7+

'                A good feature of this calcualtor is that it recolor the expression after the
'                evaluation (Operations in purple, numbers in light blue,invalid inputs in red ,
'                positive and negative signs in dark green, and parenthesis in randome colors).
'                That make the expresssion very clear .

'                Moreover, it's so easy to use. To copy the result to the clipboard, just click
'                on the result. You also can review the old expressoin you entered by using
'                Up and Down keys.

'                If you would like to use the mouse for input, then you can expand the form to
'                view an input keyboard by clicking the dark bar at the bottom. To hide the
'                input keyboard, click on the dark bar again.

'                Also the code has a lot of comments, and I hope it will be easy for you to
'                understand, however if you have any quesions, then please feel free to email
'                me at yehia_sm@hotmail.com
'
'Special Thanks: For Issam Hijazi for inspiring me with the idea.

Option Explicit

'Declare the API-Functions
'Use this function to minimize the window
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_MINIMIZE = 6

'For mouseover event
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'For voting
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)


'For moving the object when by the mouse
Dim StartMoving As Boolean
Dim InitialX As Long, InitialY As Long
'For saving expressions
Dim Expressions() As String, nExp As Integer, nTemp As Integer

Private Sub CmdAdd_Click()
AddExp "+"
End Sub

Private Sub CmdCP_Click()
AddExp ")"
End Sub

Private Sub CmdDivide_Click()
AddExp "/"
End Sub

Private Sub CmdDot_Click()
AddExp "."
End Sub

Private Sub CmdBS_Click()
Dim sStart As Integer
'Get the old input position
sStart = TxtExp.SelStart


If TxtExp.SelLength > 0 Then
    'Delete the selected part
    TxtExp.Text = Mid(TxtExp.Text, 1, sStart) + Mid(TxtExp.Text, sStart + 1 + TxtExp.SelLength)
    sStart = sStart + 1
Else
    'If the input position is at the begining the do nothing
    If sStart = 0 Then GoTo Ex
    'Just delete one character to the left of the input position
    TxtExp.Text = Mid(TxtExp.Text, 1, sStart - 1) + Mid(TxtExp.Text, sStart + 1)
End If

'Set a new input position = Old positon +1
TxtExp.SelStart = sStart - 1

Ex:
TxtExp.SetFocus
End Sub

Private Sub CmdExit_Click()
Dim Re As Byte
Re = MsgBox("Are you sure you want to exit this program?", vbYesNo + vbQuestion, "Confirmation")
If Re = vbYes Then Unload Form1
End Sub

Private Sub CmdMult_Click()
AddExp "*"
End Sub

Private Sub CmdNum_Click(Index As Integer)
AddExp Right(Str(Index), 1)
End Sub

Private Sub CmdOP_Click()
AddExp "("
End Sub

Private Sub CmdPower_Click()
AddExp "^"
End Sub

Private Sub CmdSolve_Click()
Dim Exp As String
Dim Result As Double

'Desable buttons
CmdSolve.Enabled = False
CmdExit.Enabled = False
LblExit.Enabled = False

'Get Expressoin
Exp = Trim(TxtExp.Text)
'Get rid of spaces
Exp = Replace(Exp, " ", "")
'Rewrite Expression
TxtExp = Exp

'Evaluate expression and show result
LblResult = "0"

'Reset all colors
TxtExp.SelStart = 0
TxtExp.SelLength = Len(Exp)
TxtExp.SelColor = &HFF8080
TxtExp.SelBold = False
'
If EvaluateExp(Exp, Result, 0, TxtExp, LblResult) Then
    'Show Result
    LblResult = Result
    'Save the expression so that you can get it back any time
    'Check if the expression is not like the old one
    If Exp <> Expressions(nExp) Then
        nExp = nExp + 1
        ReDim Preserve Expressions(nExp)
        Expressions(nExp) = TxtExp.Text
        nTemp = nExp
    End If
End If

'Renable buttons
CmdSolve.Enabled = True
CmdExit.Enabled = True
LblExit.Enabled = True
TxtExp.SetFocus

End Sub

Private Sub CmdClear_Click()
TxtExp.Text = ""
TxtExp.SetFocus
End Sub

Private Sub Command4_Click()

End Sub

Private Sub CmdSub_Click()
AddExp "-"
End Sub

Private Sub Form_Load()
'Initial values
nExp = 0
ReDim Preserve Expressions(nExp)
Expressions(0) = ""

'Set the color
TxtExp.SelColor = &HFF8080

'Initial height
Me.Height = 2550

'************************ ( This part can be deleted )
'Initial expression
TxtExp = "+(2((5+4)/2*6+1.56845)4 + 3^(3-(--5)/2))/(25^.5)+2-3*4/-2^2.2*5"
TxtExp.SelStart = 0
TxtExp.SelLength = Len(TxtExp)
TxtExp.SelColor = &HFF8080
TxtExp.SelLength = 0

'For the vote
PicVoteLink.ForeColor = &HFF0000
PicVoteLink.Print "Pleae click here to vote for me at PSC. I'll appreciate that ..."
'************************* ( This part can be deleted )

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
InitialX = X
InitialY = Y
StartMoving = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If StartMoving Then
    Me.Left = Me.Left + (X - InitialX)
    Me.Top = Me.Top + (Y - InitialY)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Stop moving the form
StartMoving = False
End Sub

Private Sub Label2_Click()
LblHelp_Click
End Sub

Private Sub LblExit_Click()
CmdExit_Click
End Sub

Private Sub LblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblExit.BorderStyle = 1
End Sub

Private Sub LblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblExit.BorderStyle = 0
End Sub

Private Sub LblHelp_Click()
'
Dim MSG As String

'The message
MSG = "The MagiCalc 1.0" + vbCrLf
MSG = MSG + "Copyright © 2003 Yehia Muhsen, All rights reserved.." + vbCrLf + vbCrLf
MSG = MSG + "* Conditions for the input:" + vbCrLf
MSG = MSG + "  - All numbers are allowed (e.g 73245.29143)." + vbCrLf
MSG = MSG + "  - Only the following mathematical operations are allowed: ^ * / + -." + vbCrLf
MSG = MSG + "  - Parenthesis are allowed, but not brackets." + vbCrLf
MSG = MSG + "  - Missing operation around parenthesis is considered multiplication." + vbCrLf
MSG = MSG + "  - Negative and positive signs are allowed." + vbCrLf
MSG = MSG + "  - Spaces are allow, but other characters are not." + vbCrLf & vbCrLf

MSG = MSG + "* What gets evaluated first:" & vbCrLf
MSG = MSG + "  - Parenthesis , ( ^ ) , negation, ( * , / ) , then ( + , - ) ." + vbCrLf
'MSG = MSG + "  - Power operation ' ^ ' ." + vbCrLf
'MSG = MSG + "  - Positive and negative signs ' - ', ' + '." + vbCrLf
'MSG = MSG + "  - Muliplication and division ' * ', ' / ' ." + vbCrLf
'MSG = MSG + "  - Addition and subtraction ' + ', ' - ' ." + vbCrLf + vbCrLf
MSG = MSG + vbCrLf

MSG = MSG + "* To copy the result to the clipboard, just click on it." + vbCrLf
MSG = MSG + "* To view previouse inputs, use the up and down arrows." + vbCrLf
MSG = MSG + "* To expand the form for input keyboard, click the dark bar at the bottom." + vbCrLf + vbCrLf

MSG = MSG + "Any questions or comments, email me at yehia_sm@hotmail.com."

'Show the message
MsgBox MSG, vbInformation, "About The MagiCalc 1.0"

End Sub

Private Sub LblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblHelp.BorderStyle = 1
End Sub

Private Sub LblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblHelp.BorderStyle = 0
End Sub

Private Sub LblMin_Click()
'Minimize window to the taskbar
ShowWindow Me.hwnd, SW_MINIMIZE
End Sub

Private Sub LblMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblMin.BorderStyle = 1
End Sub

Private Sub LblMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblMin.BorderStyle = 0
End Sub

Private Sub LblResult_Click()
'If the result is a number then copy it to the clipboard
If IsNumeric(LblResult.Caption) Then
    Clipboard.Clear
    Clipboard.SetText LblResult.Caption
End If
End Sub

Private Sub PicExpand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.Height = 2550 Then
    Me.Height = 5250
Else
    Me.Height = 2550
End If
TxtExp.SetFocus
End Sub

Private Sub PicExpand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Change color and run the timer
PicExpand.BackColor = &H808080
TmrExpandMouseOver.Enabled = True
End Sub

Private Sub PicVoteLink_Click()
Dim MyArticleAddr As String
Dim MSG As String

'Vote
MyArticleAddr = "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=48154&optCodeRatingValue=5"
ShellExecute &O0, "Open", MyArticleAddr, vbNullString, vbNullString, 1

'Message
MSG = "Thank you very much for you vote. I really appreciate that . " & vbCrLf & vbCrLf
MSG = MSG & "To delete the link from the form, delete the picture box ( PicVoteLinke)," & vbCrLf
MSG = MSG & "then delete the voting part, in the Form_Load precudure ."
MSG = MSG & vbCrLf & vbCrLf & "Thanks again :) ..."

MsgBox MSG, vbInformation, "Thank you"
End Sub

Private Sub PicVoteLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicVoteLink.ForeColor = &H80FF&
PicVoteLink.CurrentX = 0
PicVoteLink.CurrentY = 0
PicVoteLink.Print "Pleae click here to vote for me at PSC. I'll appreciate that ..."
TmrLinkMouseOver.Enabled = True
End Sub

Private Sub TmrExpandMouseOver_Timer()
Dim PT As POINTAPI
Dim lHwnd As Long

'Get the mouse coordinates
GetCursorPos PT
'Get the handle of the window where the mouse is
lHwnd = WindowFromPoint(PT.X, PT.Y)
'If the mouse leaves the object then reset colors
If PicExpand.hwnd <> lHwnd Then
    PicExpand.BackColor = &HC0C0C0
    TmrExpandMouseOver.Enabled = False
End If
End Sub

Private Sub TmrLinkMouseOver_Timer()
Dim PT As POINTAPI
Dim lHwnd As Long

'Get the mouse coordinates
GetCursorPos PT
'Get the handle of the window where the mouse is
lHwnd = WindowFromPoint(PT.X, PT.Y)
'If the mouse leaves the object then reset colors
If PicVoteLink.hwnd <> lHwnd Then
    PicVoteLink.ForeColor = &HFF0000
    PicVoteLink.Cls
    PicVoteLink.Print "Pleae click here to vote for me at PSC. I'll appreciate that ..."
    TmrLinkMouseOver.Enabled = False
End If
End Sub

Private Sub TxtExp_KeyDown(KeyCode As Integer, Shift As Integer)

'When hit the upper arrow
If KeyCode = 38 And nTemp > 1 Then
    nTemp = nTemp - 1
    TxtExp.Text = Expressions(nTemp)
'When hit the down arrow
ElseIf KeyCode = 40 And nTemp < nExp Then
    nTemp = nTemp + 1
    TxtExp.Text = Expressions(nTemp)
End If

'If there's no selected character then change the color to default (&HFF8080)
If TxtExp.SelLength = 0 Then
    TxtExp.SelColor = &HFF8080
    TxtExp.SelBold = False
End If

End Sub

Private Sub AddExp(ByVal Exp As String)
Dim sStart As Integer

'Get the old input position
sStart = TxtExp.SelStart
'Insert the new expression after the input postion
'If there is a selected text, it'll be removed
TxtExp.Text = Mid(TxtExp.Text, 1, sStart) + Exp + Mid(TxtExp.Text, sStart + 1 + TxtExp.SelLength)

'Reset the color to default (&HFF8080)
TxtExp.SelStart = 0
TxtExp.SelLength = Len(TxtExp)
TxtExp.SelColor = &HFF8080
TxtExp.SelBold = False
TxtExp.SelLength = 0

'Set a new input position = Old positon +1
TxtExp.SelStart = sStart + 1

TxtExp.SetFocus
End Sub
