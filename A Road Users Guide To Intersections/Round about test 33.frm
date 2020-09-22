VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roundabout Test 3"
   ClientHeight    =   6240
   ClientLeft      =   3015
   ClientTop       =   2235
   ClientWidth     =   8940
   LinkTopic       =   "Form5"
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
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C"
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "A"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
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
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
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
   Begin VB.CommandButton Command1 
      Caption         =   "B"
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "/20"
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Question 9"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   5400
      Picture         =   "Round about test 33.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Roundabout Test 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Can car 2 move onto the roundabout before car 1 passes him?"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "No"
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
      Left            =   840
      TabIndex        =   13
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Yes"
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
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Only if he speeds out quickly."
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
      TabIndex        =   11
      Top             =   2280
      Width           =   2295
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
      TabIndex        =   10
      Top             =   2280
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
      TabIndex        =   9
      Top             =   3120
      Width           =   375
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
      TabIndex        =   8
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   3480
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   3480
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   3480
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
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
Form5.Hide
Form15.Show
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
Form3.Label3.Caption = Val(Form3.Label3.Caption) + "1"
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
Form5.Hide
Form15.Show
End If

If Text2.Text = "2" Then
If Form15.Text2 = "2" Then
Form15.Command4.Enabled = False
Form5.Hide
Form15.Show
ElseIf Form16.Text2 = "2" Then
Form16.Command4.Enabled = False
Form5.Hide
Form16.Show
ElseIf Form17.Text2 = "2" Then
Form17.Command4.Enabled = False
Form5.Hide
Form17.Show
ElseIf Form12.Text2 = "2" Then
Form12.Command4.Enabled = False
Form5.Hide
Form12.Show
ElseIf Form13.Text2 = "2" Then
Form13.Command4.Enabled = False
Form5.Hide
Form13.Show
ElseIf Form14.Text2 = "2" Then
Form14.Command4.Enabled = False
Form5.Hide
Form14.Show
ElseIf Form18.Text2 = "2" Then
Form18.Command4.Enabled = False
Form5.Hide
Form18.Show
ElseIf Form19.Text2 = "2" Then
Form19.Command4.Enabled = False
Form5.Hide
Form19.Show
ElseIf Form20.Text2 = "2" Then
Form20.Command4.Enabled = False
Form5.Hide
Form20.Show
ElseIf Form21.Text2 = "2" Then
Form21.Command4.Enabled = False
Form5.Hide
Form21.Show
ElseIf Form22.Text2 = "2" Then
Form22.Command4.Enabled = False
Form5.Hide
Form22.Show
Else
Form5.Hide
Form3.Show
End If
End If
End Sub

