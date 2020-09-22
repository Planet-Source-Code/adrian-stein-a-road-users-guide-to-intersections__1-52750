VERSION 5.00
Begin VB.Form Form33 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Three Way Intersections Lesson"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Print"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Repeat"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4080
      Left            =   4800
      Picture         =   "3 way int less.frx":0000
      ScaleHeight     =   4080
      ScaleWidth      =   4095
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   1920
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   1440
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Lesson"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Lesson"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Glossary"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   3240
      Picture         =   "3 way int less.frx":36782
      ScaleHeight     =   3600
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   6480
      Picture         =   "3 way int less.frx":6CF04
      ScaleHeight     =   3600
      ScaleWidth      =   3015
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4080
      Left            =   4800
      Picture         =   "3 way int less.frx":A3686
      ScaleHeight     =   4080
      ScaleWidth      =   4095
      TabIndex        =   10
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "3 way Intersection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   9495
   End
End
Attribute VB_Name = "Form33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Command7.Visible = False
Command4.Visible = True
Picture4.Visible = False
Label1.Caption = "3 way intersections are a combination of round-a-bouts and t-intersections. A 3way intersection has a small round lump in the middle of the road. This lump is to be driven around, but you may drive over it if it is required. It has all the same rules as round-a-bouts (CLOCKWISE, Give way to your right).Press continue"
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Form33.Hide
Form32.Show
End Sub

Private Sub Command2_Click()
Command7.Visible = False
Command4.Visible = True
Picture4.Visible = False
Label1.Caption = "3 way intersections are a combination of round-a-bouts and t-intersections. A 3way intersection has a small round lump in the middle of the road. This lump is to be driven around, but you may drive over it if it is required. It has all the same rules as round-a-bouts (CLOCKWISE, Give way to your right).Press continue"
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Form33.Hide
Form34.Show

End Sub

Private Sub Command3_Click()
Form25.Show
End Sub

Private Sub Command4_Click()
Picture3.Visible = True
Picture2.Visible = False
Picture1.Visible = True
Label1.Caption = "These two pictures are correct, in the left picture, car two is allowed to go around the lump. And in the right picture, car two is allowed to slightly go over the lump, because of a sharp turn. Press Continue "
Command4.Visible = False
Command6.Visible = True
End Sub

Private Sub Command5_Click()
Command7.Visible = False
Command4.Visible = True
Picture4.Visible = False
Label1.Caption = "3 way intersections are a combination of round-a-bouts and t-intersections. A 3way intersection has a small round lump in the middle of the road. This lump is to be driven around, but you may drive over it if it is required. It has all the same rules as round-a-bouts (CLOCKWISE, Give way to your right).Press continue"
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Form33.Hide
Form36.Show
End Sub

Private Sub Command6_Click()
Label1.Caption = "This picture is incorrect, car two is  not allowed to cut the corner of the intersection, nor is  allowed to go directly over the lump. This can cause confusion and collision with other drivers. Press next Lesson."
Command6.Visible = False
Command4.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = True
Command7.Visible = True

End Sub

Private Sub Command7_Click()
Command7.Visible = False
Command4.Visible = True
Picture4.Visible = False
Label1.Caption = "3 way intersections are a combination of round-a-bouts and t-intersections. A 3 way intersection has a small round lump in the middle of the road. This lump is to be driven around, but you may drive over it, if it is required. It has all the same rules as round-a-bouts (CLOCKWISE, Give way to your right).Press continue"
Picture2.Visible = True
End Sub

Private Sub Command8_Click()
Form33.PrintForm

End Sub

Private Sub Form_Load()
Label1.Caption = "3 way intersections are a combination of round-a-bouts and t-intersections. A 3 way intersection has a small round lump in the middle of the road. This lump is to be driven around, but you may drive over it, if it is required. It has all the same rules as round-a-bouts (CLOCKWISE, Give way to your right).Press continue"
End Sub





