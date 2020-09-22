VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Intersection Test 3"
   ClientHeight    =   6240
   ClientLeft      =   2910
   ClientTop       =   2235
   ClientWidth     =   8940
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6240
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B"
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Quit"
      Height          =   735
      Left            =   7200
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Skip"
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A"
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "/20"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Question 15"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   3240
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   3240
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   3240
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "B."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "A."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Both"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Bike"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Car"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "What has right of way at this intersection, the car, the bike or both?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "T-Intersection Test 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   720
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   2760
      Left            =   5160
      Picture         =   "T-intersection33.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3555
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command5.Visible = True
Command1.BackColor = vbYellow
Command2.BackColor = &H8000000F
Command3.BackColor = &H8000000F
End Sub

Private Sub Command2_Click()
Command5.Visible = True
Command2.BackColor = vbYellow
Command1.BackColor = &H8000000F
Command3.BackColor = &H8000000F
End Sub

Private Sub Command3_Click()
Command5.Visible = True
Command3.BackColor = vbYellow
Command2.BackColor = &H8000000F
Command1.BackColor = &H8000000F
End Sub

Private Sub Command4_Click()
Text2.Text = "2"
Form14.Hide
Form18.Show
End Sub

Private Sub Command5_Click()
If Command1.BackColor = vbYellow Then
Image5.Picture = LoadPicture(App.Path + "/Pics and sounds\arrow.gif")
Image6.Picture = LoadPicture(App.Path + "/Pics and sounds\cross.gif")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

If Command3.BackColor = vbYellow Then
Image5.Picture = LoadPicture(App.Path + "/Pics and sounds\tick.gif")
Text1.Text = "1"
Form1.Label11.Caption = Val(Form1.Label11.Caption) + "1"
Form10.Label11.Caption = Val(Form10.Label11.Caption) + "1"
Form11.Label11.Caption = Val(Form11.Label11.Caption) + "1"
Form12.Label11.Caption = Val(Form12.Label11.Caption) + "1"
Form13.Label11.Caption = Val(Form13.Label11.Caption) + "1"
Form14.Label11.Caption = Val(Form14.Label11.Caption) + "1"
Form15.Label11.Caption = Val(Form15.Label11.Caption) + "1"
Form16.Label11.Caption = Val(Form16.Label11.Caption) + "1"
Form17.Label11.Caption = Val(Form17.Label11.Caption) + "1"
Form18.Label11.Caption = Val(Form18.Label11.Caption) + "1"
Form19.Label11.Caption = Val(Form19.Label11.Caption) + "1"
Form2.Label11.Caption = Val(Form2.Label11.Caption) + "1"
Form20.Label11.Caption = Val(Form20.Label11.Caption) + "1"
Form21.Label11.Caption = Val(Form21.Label11.Caption) + "1"
Form22.Label11.Caption = Val(Form22.Label11.Caption) + "1"
Form5.Label11.Caption = Val(Form5.Label11.Caption) + "1"
Form6.Label11.Caption = Val(Form6.Label11.Caption) + "1"
Form7.Label11.Caption = Val(Form7.Label11.Caption) + "1"
Form8.Label11.Caption = Val(Form8.Label11.Caption) + "1"
Form9.Label11.Caption = Val(Form9.Label11.Caption) + "1"
Form3.Label12.Caption = Val(Form3.Label12.Caption) + "1"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

If Command2.BackColor = vbYellow Then
Image5.Picture = LoadPicture(App.Path + "/Pics and sounds\arrow.gif")
Image4.Picture = LoadPicture(App.Path + "/Pics and sounds\cross.gif")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

End Sub

Private Sub Command6_Click()
Form23.Show
End Sub


Private Sub Command7_Click()
If Text2.Text = "" Then
Form14.Hide
Form18.Show
End If

If Text2.Text = "2" Then
If Form18.Text2 = "2" Then
Form18.Command4.Enabled = False
Form14.Hide
Form18.Show
ElseIf Form19.Text2 = "2" Then
Form19.Command4.Enabled = False
Form14.Hide
Form19.Show
ElseIf Form20.Text2 = "2" Then
Form20.Command4.Enabled = False
Form14.Hide
Form20.Show
ElseIf Form21.Text2 = "2" Then
Form21.Command4.Enabled = False
Form14.Hide
Form21.Show
ElseIf Form22.Text2 = "2" Then
Form22.Command4.Enabled = False
Form14.Hide
Form22.Show
Else
Form14.Hide
Form3.Show
End If
End If
End Sub

