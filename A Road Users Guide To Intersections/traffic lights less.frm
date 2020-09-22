VERSION 5.00
Begin VB.Form Form30 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traffic Lights Lesson"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Print"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Repeat"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   7200
      Picture         =   "traffic lights less.frx":0000
      ScaleHeight     =   2040
      ScaleWidth      =   1020
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Glossary"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Lesson"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Lesson"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3795
      Left            =   5400
      Picture         =   "traffic lights less.frx":6CA2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4290
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   9495
   End
   Begin VB.Label Label2 
      Caption         =   "Traffic lights"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = "This is a Traffic lights intersection. These intersections contain four (4) roads crossing at a point. Obviously, this whould cause crashes, so that is why they have traffic lights. Press the button to continue"
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
End Sub

Private Sub Command2_Click()
Label1.Caption = "This is a Traffic lights intersection. These intersections contain four (4) roads crossing at a point. Obviously, this whould cause crashes, so that is why they have traffic lights. Press the button to continue"
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Form30.Hide
Form31.Show

End Sub

Private Sub Command3_Click()
Form25.Show
End Sub

Private Sub Command4_Click()
Image1.Visible = False
Picture1.Visible = True
Label1.Caption = "This is a traffic light, it has three lights which tell the driver what to do. You must obey the CLOSEST LIGHT ON YOUR LEFT, or more generally, the ones that are facing you. The red light (top) indicates to come to a complete stop on the line, or close behind a car. The Amber light (middle) indicates to stop, but you can travel throught it if comming to a complete stop could cause a crash. The amber light flashes for a few seconds after the green and before the red. The green light (bottom) indicates to proceed through, watching for on-coming cars or the lights to change. Press continue."
Command4.Visible = False
Command6.Visible = True
End Sub

Private Sub Command5_Click()
Label1.Caption = "This is a Traffic lights intersection. These intersections contain four (4) roads crossing at a point. Obviously, this would cause crashes, so that is why they have traffic lights. Press the button to continue"
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Form30.Hide
Form36.Show
End Sub

Private Sub Command6_Click()
Command7.Visible = True
Command6.Visible = False
Image1.Visible = True
Picture1.Visible = False
Label1.Caption = " You MUST obey the traffic lights, unless there is a policeman directing traffic, or the lights are not working, which we will cover next, press Next Lesson"
End Sub

Private Sub Command7_Click()
Label1.Caption = "This is a Traffic lights intersection. These intersections contain four (4) roads crossing at a point. Obviously, this would cause crashes, so that is why they have traffic lights. Press the button to continue"
Command7.Visible = False
Command4.Visible = True

End Sub

Private Sub Command8_Click()
Form30.PrintForm
End Sub

Private Sub Form_Load()
Label1.Caption = "This is a Traffic lights intersection. These intersections contain four (4) roads crossing at a point. Obviously, this would cause crashes, so that is why they have traffic lights. Press the button to continue"
End Sub

Private Sub Picture2_Click()

End Sub

