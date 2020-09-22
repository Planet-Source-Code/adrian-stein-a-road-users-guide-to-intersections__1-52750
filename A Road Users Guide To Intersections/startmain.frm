VERSION 5.00
Begin VB.Form Form36 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WELCOME"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command21 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   6360
      Top             =   2280
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H0000C000&
      Caption         =   "Quit"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H0000C000&
      Caption         =   "Previous Results"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H0000C000&
      Caption         =   "Read me"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0000C000&
      Caption         =   "Website"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0000C000&
      Caption         =   "View intro"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000C000&
      Caption         =   "Glossary"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H000080FF&
      Caption         =   "Cross roads"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H000080FF&
      Caption         =   "T - Intersections"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000080FF&
      Caption         =   "3 Way"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000080FF&
      Caption         =   "Round-a-bout"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   2295
      Left            =   5160
      TabIndex        =   21
      Top             =   2880
      Width           =   3495
      Begin VB.Label Label1 
         Caption         =   $"startmain.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   3840
      Top             =   4920
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   3840
      Top             =   5400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   2880
      Top             =   5400
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   3360
      Top             =   4920
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   3360
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   2880
      Top             =   4920
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H000080FF&
      Caption         =   "Malfunctions"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Caption         =   "Traffic lights"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000080FF&
      Caption         =   "Indrodunction"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Caption         =   "T- intersections"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "3 Way "
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Round-a-bout"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Malfunctions"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Traffic lights"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Introduction"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Main3 
      Caption         =   "Options"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton main2 
      Caption         =   "Tests"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Main1 
      Caption         =   "Lessons"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000000FF&
      Caption         =   "Cross roads"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   2655
      TabIndex        =   20
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   2880
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Stuart Morrison"
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Alex White"
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Adrian Stein"
      Height          =   375
      Left            =   3840
      TabIndex        =   29
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Made By:"
      Height          =   375
      Left            =   3000
      TabIndex        =   28
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   1065
      Left            =   -6960
      Picture         =   "startmain.frx":00C0
      Top             =   1440
      Visible         =   0   'False
      Width           =   7155
   End
   Begin VB.Image Image3 
      Height          =   990
      Left            =   6000
      Picture         =   "startmain.frx":18E2A
      Top             =   5880
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   9480
      Picture         =   "startmain.frx":1F16C
      Top             =   240
      Visible         =   0   'False
      Width           =   8820
   End
   Begin VB.Label Label3 
      Height          =   8055
      Left            =   2760
      TabIndex        =   25
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time1 As Integer
Dim Time2 As Integer
Dim time3 As Integer
Dim time4 As Integer
Dim time5 As Integer
Dim time6 As Integer
Dim maxidx As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long








Private Sub Command1_Click()
Form36.Hide
Form27.Show


End Sub

Private Sub Command10_Click()
MsgBox "you must start on the first test."
End Sub

Private Sub Command11_Click()
MsgBox "you must start on the first test."
End Sub

Private Sub Command12_Click()
MsgBox "you must start on the first test."
End Sub

Private Sub Command13_Click()
MsgBox "you must start on the first test."
End Sub

Private Sub Command14_Click()
MsgBox "you must start on the first test."
End Sub

Private Sub Command15_Click()
Form25.Show
End Sub

Private Sub Command16_Click()
Form28.Show
Form36.Hide
Form28.Picture1.Left = 5640
sndPlaySound App.Path + "\boy.wav", 1
Form28.Timer1.Enabled = True
Form28.Timer3.Enabled = True
Form28.Label10.Visible = False
Form28.Command1.Visible = False
Form28.Label11.Visible = False
Form28.Label9.Visible = False
Form28.Label10.Visible = False
Form28.Label8.Visible = True
Form28.Label1.Caption = "A"
Form28.Label2.Caption = "d"
Form28.Label3.Caption = "r"
Form28.Label4.Caption = "i"
Form28.Label5.Caption = "a"
Form28.Label6.Caption = "n"
End Sub

Private Sub Command17_Click()
  Dim ret&
  ret& = ShellExecute(Me.hwnd, "Open", "http://rta.nsw.gov.au/licensing/tests/driverknowledgetest/demonstrationdriverknowledgetest/index.html", "", App.Path, 1)
End Sub

Private Sub Command18_Click()
Form26.Show

End Sub

Private Sub Command19_Click()
Form36.Hide
Form3.Show
End Sub

Private Sub Command2_Click()
Form36.Hide
Form30.Show

End Sub

Private Sub Command20_Click()
Form36.Enabled = False
Form24.Show

End Sub

Private Sub Command21_Click()
Form36.Enabled = False
Form24.Show
End Sub

Private Sub Command21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command21.BackColor = &HFFFF00

End Sub

Private Sub Command3_Click()
Form36.Hide
Form31.Show
End Sub

Private Sub Command4_Click()
Form36.Hide
Form32.Show
End Sub

Private Sub Command5_Click()
Form36.Hide
Form33.Show
End Sub

Private Sub Command6_Click()
Form36.Hide
Form34.Show
End Sub

Private Sub Command7_Click()
Form36.Hide
Form35.Show
End Sub

Private Sub clear_Click()
'Static maxidx
'If maxidx = 0 Then
'    maxidx = 0
'End If
'If maxidx < 7 Then

 '   Command1(maxidx).BackColor = vbRed
'    maxidx = maxidx + 1
'Else: maxidx = 0
'End If
End Sub

Private Sub Command8_Click()
Form36.Hide
Form4.Show

End Sub

Private Sub Command9_Click()
Form36.Hide
Form6.Show

End Sub

Private Sub Form_Load()
time1 = 0
Time2 = 0
End Sub

Private Sub Label2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
MsgBox "yeep"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.BackColor = vbRed
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Timer2.Enabled = True
time1 = 0

Timer3.Enabled = False
Timer4.Enabled = True
time3 = 0

Timer5.Enabled = False
Timer6.Enabled = True
time5 = 0

Command1.BackColor = vbRed
Command2.BackColor = vbRed
Command3.BackColor = vbRed
Command4.BackColor = vbRed
Command5.BackColor = vbRed
Command6.BackColor = vbRed
Command7.BackColor = vbRed
Command8.BackColor = &H80FF&
Command9.BackColor = &H80FF&
Command10.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Command12.BackColor = &H80FF&
Command13.BackColor = &H80FF&
Command14.BackColor = &H80FF&
Command15.BackColor = &HC000&
Command16.BackColor = &HC000&
Command17.BackColor = &HC000&
Command18.BackColor = &HC000&
Command19.BackColor = &HC000&
Command20.BackColor = &HC000&
Command21.BackColor = &H8000000F
Label1.Caption = "This program is designed to teach you all about intersections, and then test you on your knowledge with a series of multiple choice questions. Just mouse over a button to see your options."
End Sub

Private Sub Main1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
Timer2.Enabled = False
Time2 = 0

Timer3.Enabled = False
Timer4.Enabled = True
time3 = 0

Timer5.Enabled = False
Timer6.Enabled = True
time5 = 0

End Sub

Private Sub main2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer3.Enabled = True
Timer4.Enabled = False
time4 = 0

Timer1.Enabled = False
Timer2.Enabled = True
time1 = 0

Timer5.Enabled = False
Timer6.Enabled = True
time5 = 0
End Sub

Private Sub Main3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer5.Enabled = True
Timer6.Enabled = False
time6 = 0

Timer1.Enabled = False
Timer2.Enabled = True
time1 = 0

Timer3.Enabled = False
Timer4.Enabled = True
time3 = 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Timer2.Enabled = True
time1 = 0
Timer3.Enabled = False
Timer4.Enabled = True
time3 = 0

Timer5.Enabled = False
Timer6.Enabled = True
time5 = 0

Command1.BackColor = vbRed
Command2.BackColor = vbRed
Command3.BackColor = vbRed
Command4.BackColor = vbRed
Command5.BackColor = vbRed
Command6.BackColor = vbRed
Command7.BackColor = vbRed
Command8.BackColor = &H80FF&
Command9.BackColor = &H80FF&
Command10.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Command12.BackColor = &H80FF&
Command13.BackColor = &H80FF&
Command14.BackColor = &H80FF&
Command15.BackColor = &HC000&
Command16.BackColor = &HC000&
Command17.BackColor = &HC000&
Command18.BackColor = &HC000&
Command19.BackColor = &HC000&
Command20.BackColor = &HC000&
Label1.Caption = "This program is designed to teach you all about intersections, and then test you on your knowledge with a series of multiple choice questions. Just mouse over a button to see your options."
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
time1 = time1 + 1
Select Case time1
    Case Is = 1
        Command1.Visible = True
    Case Is = 2
        Command2.Visible = True
    Case Is = 3
        Command3.Visible = True
    Case Is = 4
        Command4.Visible = True
    Case Is = 5
        Command5.Visible = True
    Case Is = 6
        Command6.Visible = True
    Case Is = 7
        Command7.Visible = True


End Select
End Sub

Private Sub Timer2_Timer()
Time2 = Time2 + 1
Select Case Time2
    Case Is = 1
        Command7.Visible = False
    Case Is = 2
        Command6.Visible = False
    Case Is = 3
        Command5.Visible = False
    Case Is = 4
        Command4.Visible = False
    Case Is = 5
        Command3.Visible = False
    Case Is = 6
        Command2.Visible = False
    Case Is = 7
        Command1.Visible = False

End Select
End Sub

Private Sub Timer3_Timer()
time3 = time3 + 1
Select Case time3
    Case Is = 1
        Command8.Visible = True
    Case Is = 2
        Command9.Visible = True
    Case Is = 3
        Command10.Visible = True
    Case Is = 4
        Command11.Visible = True
    Case Is = 5
        Command12.Visible = True
    Case Is = 6
        Command13.Visible = True
    Case Is = 7
        Command14.Visible = True

End Select
End Sub

Private Sub Timer4_Timer()
time4 = time4 + 1
Select Case time4
    Case Is = 1
        Command14.Visible = False
    Case Is = 2
        Command13.Visible = False
    Case Is = 3
        Command12.Visible = False
    Case Is = 4
        Command11.Visible = False
    Case Is = 5
        Command10.Visible = False
    Case Is = 6
        Command9.Visible = False
    Case Is = 7
        Command8.Visible = False


End Select
End Sub

Private Sub Timer5_Timer()
time5 = time5 + 1
Select Case time5
    Case Is = 1
        Command15.Visible = True
    Case Is = 2
        Command16.Visible = True
    Case Is = 3
        Command17.Visible = True
    Case Is = 4
        Command18.Visible = True
    Case Is = 5
        Command19.Visible = True
    Case Is = 6
        Command20.Visible = True

End Select
End Sub

Private Sub Timer6_Timer()
time6 = time6 + 1
Select Case time6
    Case Is = 1
        Command20.Visible = False
    Case Is = 2
        Command19.Visible = False
    Case Is = 3
        Command18.Visible = False
    Case Is = 4
        Command17.Visible = False
    Case Is = 5
        Command16.Visible = False
    Case Is = 6
        Command15.Visible = False

End Select
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HFFFF&
Command2.BackColor = vbRed
Command3.BackColor = vbRed
Command6.BackColor = vbRed
Command7.BackColor = vbRed
Command5.BackColor = vbRed
Command4.BackColor = vbRed
Label1.Caption = "Press this button for a General Introduction to the lessons."
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HFFFF&
Command1.BackColor = vbRed
Command3.BackColor = vbRed
Label1.Caption = "This is a lesson on Traffic light intersections"
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = &HFFFF&
Command2.BackColor = vbRed
Command4.BackColor = vbRed
Label1.Caption = "This is a lesson on Malfunctions in intersections"
End Sub
Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.BackColor = &HFFFF&
Command5.BackColor = vbRed
Command3.BackColor = vbRed
Label1.Caption = "This is a lesson on Round-a-bout intersections"
End Sub
Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BackColor = &HFFFF&
Command4.BackColor = vbRed
Command6.BackColor = vbRed
Label1.Caption = "This is a lesson on 3-way intersections"
End Sub
Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.BackColor = &HFFFF&
Command7.BackColor = vbRed
Command5.BackColor = vbRed
Label1.Caption = "This is a lesson on T-intersections"
End Sub
Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.BackColor = &HFFFF&
Command6.BackColor = vbRed
Command5.BackColor = vbRed
Command1.BackColor = vbRed
Command2.BackColor = vbRed
Command3.BackColor = vbRed
Command6.BackColor = vbRed
Label1.Caption = "This is a lesson on Cross road intersections"
Command5.BackColor = vbRed
Command4.BackColor = vbRed
End Sub
Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command8.BackColor = &HFFFF&
Command9.BackColor = &H80FF&
Command10.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Label1.Caption = "This is a General Introduction to the tests and how they are done."
End Sub
Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command9.BackColor = &HFFFF&
Command8.BackColor = &H80FF&
Command10.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Label1.Caption = "This is a Test on Traffic light intersections"
End Sub
Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command10.BackColor = &HFFFF&
Command9.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Command12.BackColor = &H80FF&
Label1.Caption = "This is a Test on Malfunctions in intersections"
End Sub
Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command11.BackColor = &HFFFF&
Command10.BackColor = &H80FF&
Command9.BackColor = &H80FF&
Command12.BackColor = &H80FF&
Label1.Caption = "This is a Test on Round-a-bout intersections"
End Sub
Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command12.BackColor = &HFFFF&
Command10.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Command13.BackColor = &H80FF&
Label1.Caption = "This is a Test on 3-way intersections"
End Sub
Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command13.BackColor = &HFFFF&
Command12.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Command14.BackColor = &H80FF&
Label1.Caption = "This is a Test on T-intersections"
End Sub
Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command14.BackColor = &HFFFF&
Command13.BackColor = &H80FF&
Command12.BackColor = &H80FF&
Command11.BackColor = &H80FF&
Label1.Caption = "This is a Test on Cross road intersections"
End Sub
Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command15.BackColor = &HFFFF&
Command14.BackColor = &HC000&
Command16.BackColor = &HC000&
Command17.BackColor = &HC000&
Label1.Caption = "This button leads you to the Glossary page, to find out the meaning of more complex words"
End Sub
Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command16.BackColor = &HFFFF&
Command14.BackColor = &HC000&
Command15.BackColor = &HC000&
Command17.BackColor = &HC000&
Label1.Caption = "This button takes you to the Intro."
End Sub
Private Sub Command17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command17.BackColor = &HFFFF&
Command16.BackColor = &HC000&
Command18.BackColor = &HC000&
Command19.BackColor = &HC000&
Label1.Caption = "This will take you to te RTA website."


End Sub
Private Sub Command18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command18.BackColor = &HFFFF&
Command17.BackColor = &HC000&
Command19.BackColor = &HC000&
Command16.BackColor = &HC000&
Label1.Caption = "This will take you to the readme file"
End Sub
Private Sub Command19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command19.BackColor = &HFFFF&
Command18.BackColor = &HC000&
Command17.BackColor = &HC000&
Command20.BackColor = &HC000&
Label1.Caption = "This is a list of previous results from the test"
End Sub
Private Sub Command20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command20.BackColor = &HFFFF&
Command19.BackColor = &HC000&
Command18.BackColor = &HC000&
Command17.BackColor = &HC000&
Label1.Caption = "This will quit the program."
End Sub

Private Sub Timer7_Timer()
Image1.Visible = True
Image2.Visible = True
If Image1.Left > 600 Then
    Image1.Left = Image1.Left - 42
End If
If Image2.Left < 1320 Then
    Image2.Left = Image2.Left + 39
    Label1.Visible = True
End If

End Sub
