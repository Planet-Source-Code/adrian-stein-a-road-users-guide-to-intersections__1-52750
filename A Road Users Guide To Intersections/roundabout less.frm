VERSION 5.00
Begin VB.Form Form32 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roundabout Lesson"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
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
      Left            =   5280
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Repeat"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   2040
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   1560
      Top             =   1800
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   3480
      Picture         =   "roundabout less.frx":0000
      ScaleHeight     =   4110
      ScaleWidth      =   4860
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   3480
      Picture         =   "roundabout less.frx":4109A
      ScaleHeight     =   4110
      ScaleWidth      =   4860
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   3480
      Picture         =   "roundabout less.frx":82134
      ScaleHeight     =   4110
      ScaleWidth      =   4860
      TabIndex        =   9
      Top             =   0
      Width           =   4860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Lesson"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Lesson"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Glossary"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   3480
      Picture         =   "roundabout less.frx":C31CE
      ScaleHeight     =   4110
      ScaleWidth      =   4860
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Roundabouts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   9495
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Label1.Caption = "Round-a-bouts are used as a alternative to traffic lights. They stop the frustration of wiating for the lights to change when there isn't anyone else there. With round-a-bouts there is one general rule, GIVE WAY TO YOUR RIGHT.  Press to continue."
Picture1.Visible = False
Picture4.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Timer2.Enabled = False
Timer1.Enabled = False
Form32.Hide
Form31.Show
End Sub

Private Sub Command2_Click()
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Label1.Caption = "Round-a-bouts are used as a alternative to traffic lights. They stop the frustration of wiating for the lights to change when there isn't anyone else there. With round-a-bouts there is one general rule, GIVE WAY TO YOUR RIGHT.  Press to continue."
Picture1.Visible = False
Picture4.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Timer2.Enabled = False
Timer1.Enabled = False
Form32.Hide
Form33.Show

End Sub

Private Sub Command3_Click()
Form25.Show
End Sub

Private Sub Command4_Click()
Picture2.Visible = False
Picture1.Visible = True
Label1.Caption = "Car one here has right of way, as it is on the RIGHT of car 2. You can only enter the round-a-bout when there is no cars on you right coming around, that you will considerably slow down or interfere with if you enter. Press continue."
Command4.Visible = False
Command6.Visible = True
End Sub

Private Sub Command5_Click()
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Label1.Caption = "Round-a-bouts are used as a alternative to traffic lights. They stop the frustration of wiating for the lights to change when there isn't anyone else there. With round-a-bouts there is one general rule, GIVE WAY TO YOUR RIGHT.  Press to continue."
Picture1.Visible = False
Picture4.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Timer2.Enabled = False
Timer1.Enabled = False
Form32.Hide
Form36.Show
End Sub

Private Sub Command6_Click()

Command7.Visible = True
Command6.Visible = False
Picture1.Visible = True
Picture3.Visible = True
Timer2.Enabled = True
Timer1.Enabled = True
Label1.Caption = " You may only travel CLOCKWISE around a round-a-bout. "
End Sub

Private Sub Command7_Click()
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Label1.Caption = "Round-a-bouts are used as a alternative to traffic lights. They stop the frustration of wiating for the lights to change when there isn't anyone else there. With round-a-bouts there is one general rule, GIVE WAY TO YOUR RIGHT.  Press continue."
Picture1.Visible = False
Picture4.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Timer2.Enabled = False
Timer1.Enabled = False

End Sub


Private Sub Command8_Click()
Form32.PrintForm

End Sub

Private Sub Form_Load()
Label1.Caption = "Round-a-bouts are used as a alternative to traffic lights. They stop the frustration of wiating for the lights to change when there isn't anyone else there. With round-a-bouts there is one general rule, GIVE WAY TO YOUR RIGHT.  Press continue."
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()
Picture4.Visible = False


End Sub

Private Sub Timer2_Timer()
Picture4.Visible = True

End Sub
