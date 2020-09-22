VERSION 5.00
Begin VB.Form Form34 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Intersection Lesson"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Print"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Repeat"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   3840
      Picture         =   "tint less.frx":0000
      ScaleHeight     =   4050
      ScaleWidth      =   4560
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   3840
      Picture         =   "tint less.frx":3C222
      ScaleHeight     =   3975
      ScaleWidth      =   4425
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   4425
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Continue"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4170
      Left            =   3840
      Picture         =   "tint less.frx":79814
      ScaleHeight     =   4170
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   240
      Width           =   4515
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Glossary"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Lesson"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Lesson"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   1440
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   1920
      Top             =   1800
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   9135
   End
   Begin VB.Label Label2 
      Caption         =   "T-Intersection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Command7.Visible = False
Command6.Visible = False
Command4.Visible = True
Label1.Caption = "A T-intersection is a crossroad, but with on of the roads terminating, and one continuing. The general rule for T-INtersections is, 'The cars of the terminating road must give way at all times'.Press continue"
Picture2.Visible = False
Picture1.Visible = True
Picture3.Visible = False
Form34.Hide
Form33.Show
End Sub

Private Sub Command2_Click()
Command7.Visible = False
Command6.Visible = False
Command4.Visible = True
Label1.Caption = "A T-intersection is a crossroad, but with on of the roads terminating, and one continuing. The general rule for T-INtersections is, 'The cars of the terminating road must give way at all times'.Press continue"
Picture2.Visible = False
Picture1.Visible = True
Picture3.Visible = False
Form34.Hide
Form35.Show

End Sub

Private Sub Command3_Click()
Form36.Hide
Form25.Show
End Sub

Private Sub Command4_Click()
Picture1.Visible = False
Picture3.Visible = True
Label1.Caption = "THe road labeled B is referrred to as the CONTINUING ROAD. The road labled A is referred to as the terminating road. As you can see Car two must give way to Car one. Press Continue "
Command4.Visible = False
Command6.Visible = True
End Sub

Private Sub Command5_Click()
Command7.Visible = False
Command6.Visible = False
Command4.Visible = True
Label1.Caption = "A T-intersection is a crossroad, but with on of the roads terminating, and one continuing. The general rule for T-Intersections is, 'The cars of the terminating road must give way at all times'.Press continue"
Picture2.Visible = False
Picture1.Visible = True
Picture3.Visible = False
Form34.Hide
Form36.Show
End Sub

Private Sub Command6_Click()
Label1.Caption = "Some times T-Intersections have traffic lights. These are the only time where the continuing road gives way to the Terminating road, the they have the green light. Press Next Lesson."
Command6.Visible = False
Command4.Visible = False
Command7.Visible = True
Picture3.Visible = False
Picture2.Visible = True
End Sub

Private Sub Command7_Click()
Command7.Visible = False
Command4.Visible = True
Label1.Caption = "A T-intersection is a crossroad, but with on of the roads terminating, and one continuing. The general rule for T-Intersections is, 'The cars of the terminating road must give way at all times'.Press continue"
Picture2.Visible = False
Picture1.Visible = True
Picture3.Visible = False
End Sub

Private Sub Command8_Click()
Form34.PrintForm

End Sub

Private Sub Form_Load()
Label1.Caption = "A T-intersection is a crossroad, but with on of the roads terminating, and one continuing. The general rule for T-Intersections is, 'The cars of the terminating road must give way at all times'.Press continue"
End Sub



