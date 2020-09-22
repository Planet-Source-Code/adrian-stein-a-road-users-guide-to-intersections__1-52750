VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traffic Lights Not Working Test 1"
   ClientHeight    =   6240
   ClientLeft      =   2955
   ClientTop       =   2235
   ClientWidth     =   8940
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   7
      Top             =   4080
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
      Caption         =   "A"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Quit"
      Height          =   735
      Left            =   7200
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "B"
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "/20"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Question 4"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   3360
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   3360
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   3360
      Top             =   3000
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
      Left            =   360
      TabIndex        =   15
      Top             =   3840
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
      Left            =   360
      TabIndex        =   14
      Top             =   3240
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
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Any car can go at any time."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   12
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Car 2."
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
      Left            =   840
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Car 1."
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
      Left            =   960
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Which car has right of way at this intersection?"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Image Image2 
      Height          =   2880
      Left            =   5640
      Picture         =   "Traffic Lights Not Working1.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3195
   End
   Begin VB.Label Label1 
      Caption         =   "Traffic Lights Not Working Test 1"
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
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "Form9"
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
Form9.Hide
Form10.Show
End Sub

Private Sub Command5_Click()
If Command1.BackColor = vbYellow Then
Image6.Picture = LoadPicture(App.Path + "/Pics and sounds\tick.gif")
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
Form3.Label9.Caption = Val(Form3.Label9.Caption) + "1"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

If Command3.BackColor = vbYellow Then
Image6.Picture = LoadPicture(App.Path + "/Pics and sounds\arrow.gif")
Image5.Picture = LoadPicture(App.Path + "/Pics and sounds\cross.gif")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

If Command2.BackColor = vbYellow Then
Image6.Picture = LoadPicture(App.Path + "/Pics and sounds\arrow.gif")
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
Form9.Hide
Form10.Show
End If

If Text2.Text = "2" Then
If Form10.Text2 = "2" Then
Form10.Command4.Enabled = False
Form9.Hide
Form10.Show
ElseIf Form11.Text2 = "2" Then
Form11.Command4.Enabled = False
Form9.Hide
Form11.Show
ElseIf Form2.Text2 = "2" Then
Form2.Command4.Enabled = False
Form9.Hide
Form2.Show
ElseIf Form1.Text2 = "2" Then
Form1.Command4.Enabled = False
Form9.Hide
Form1.Show
ElseIf Form5.Text2 = "2" Then
Form5.Command4.Enabled = False
Form9.Hide
Form5.Show
ElseIf Form15.Text2 = "2" Then
Form15.Command4.Enabled = False
Form9.Hide
Form15.Show
ElseIf Form16.Text2 = "2" Then
Form16.Command4.Enabled = False
Form9.Hide
Form16.Show
ElseIf Form17.Text2 = "2" Then
Form17.Command4.Enabled = False
Form9.Hide
Form17.Show
ElseIf Form12.Text2 = "2" Then
Form12.Command4.Enabled = False
Form9.Hide
Form12.Show
ElseIf Form13.Text2 = "2" Then
Form13.Command4.Enabled = False
Form9.Hide
Form13.Show
ElseIf Form14.Text2 = "2" Then
Form14.Command4.Enabled = False
Form9.Hide
Form14.Show
ElseIf Form18.Text2 = "2" Then
Form18.Command4.Enabled = False
Form9.Hide
Form18.Show
ElseIf Form19.Text2 = "2" Then
Form19.Command4.Enabled = False
Form9.Hide
Form19.Show
ElseIf Form20.Text2 = "2" Then
Form20.Command4.Enabled = False
Form9.Hide
Form20.Show
ElseIf Form21.Text2 = "2" Then
Form21.Command4.Enabled = False
Form9.Hide
Form21.Show
ElseIf Form22.Text2 = "2" Then
Form22.Command4.Enabled = False
Form9.Hide
Form22.Show
Else
Form9.Hide
Form3.Show
End If
End If
End Sub


