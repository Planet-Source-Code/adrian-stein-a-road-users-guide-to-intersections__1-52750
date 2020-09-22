VERSION 5.00
Begin VB.Form Form35 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crossroads"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Print"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Take Test"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Repeat"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2790
      Left            =   4080
      Picture         =   "crossroad less.frx":0000
      ScaleHeight     =   2790
      ScaleWidth      =   4185
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   4185
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   1800
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   1320
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Lesson"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Lesson"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Glossary"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   4320
      Picture         =   "crossroad less.frx":26292
      ScaleHeight     =   3675
      ScaleWidth      =   3690
      TabIndex        =   1
      Top             =   720
      Width           =   3690
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Crossroads"
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
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   9135
   End
End
Attribute VB_Name = "Form35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Command4.Visible = True
Command7.Visible = False
Label1.Caption = "Crossroads are two roads that meet at a point. To control the traffic, they can either put in traffic lights, or they can put in special signs that control the flow. Press continue"
Picture1.Visible = True
Picture2.Visible = False
Form35.Hide
Form34.Show
End Sub

Private Sub Command2_Click()
Command4.Visible = True
Command7.Visible = False
Label1.Caption = "Crossroads are two roads that meet at a point. To control the traffic, they can either put in traffic lights, or they can put in special signs that control the flow. Press continue"
Picture1.Visible = True
Picture2.Visible = False
Form2.Hide
Form3.Show

End Sub

Private Sub Command3_Click()
Form36.Hide
Form25.Show
End Sub

Private Sub Command4_Click()

Picture2.Visible = True
Label1.Caption = "These are the two signs. The sign on the left, is a give way sign. If you come to a give way sign, you must give way to ALL CARS, not just on your right. The sign on te Right, is a STOP sign. When confronted by one, you must come to a COMPLETE stop, and after, give way to all cars. Press Continue "
Command4.Visible = False
Command6.Visible = True
Picture1.Visible = False
End Sub

Private Sub Command5_Click()
Command4.Visible = True
Command7.Visible = False
Label1.Caption = "Crossroads are two roads that meet at a point. To control the traffic, they can either put in traffic lights, or they can put in special signs that control the flow. Press continue"
Picture1.Visible = True
Picture2.Visible = False
Form35.Hide
Form36.Show
End Sub

Private Sub Command6_Click()
Label1.Caption = "When approached by a sign, you have to wait until there is enough room for you to enter the intersection without causing other cars to slow down. Now you are ready for the test!"
Command6.Visible = False
Command4.Visible = False
Command7.Visible = True
End Sub

Private Sub Command7_Click()
Command4.Visible = True
Command7.Visible = False
Label1.Caption = "Crossroads are two roads that meet at a point. To control the traffic, they can either put in traffic lights, or they can put in special signs that control the flow. Press continue"
Picture1.Visible = True
Picture2.Visible = False
End Sub

Private Sub Command8_Click()
Command4.Visible = True
Command7.Visible = False
Label1.Caption = "Crossroads are two roads that meet at a point. To control the traffic, they can either put in traffic lights, or they can put in special signs that control the flow. Press continue"
Picture1.Visible = True
Picture2.Visible = False
Form35.Hide
Form4.Show
End Sub

Private Sub Command9_Click()
Form35.PrintForm

End Sub

Private Sub Form_Load()
Label1.Caption = "Crossroads are two roads that meet at a point. To control the traffic, they can either put in traffic lights, or they can put in special signs that control the flow. Press continue"
End Sub

