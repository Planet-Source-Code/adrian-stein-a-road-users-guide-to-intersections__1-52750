VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cross Road Test 2"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "B"
      Height          =   735
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Skip"
      Height          =   735
      Left            =   5520
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Quit"
      Height          =   735
      Left            =   7200
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   2760
      Left            =   5280
      Picture         =   "test2.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "Cross Roads Test 2"
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
      Left            =   720
      TabIndex        =   15
      Top             =   480
      Width           =   3735
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
      TabIndex        =   14
      Top             =   1560
      Width           =   4935
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
      Left            =   840
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
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
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
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
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
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
      Top             =   3000
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
      Left            =   3360
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   3360
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   3360
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command5.Visible = True
Command1.BackColor = vbRed
Command2.BackColor = &H8000000F
Command3.BackColor = &H8000000F
End Sub

Private Sub Command2_Click()
Command5.Visible = True
Command2.BackColor = vbRed
Command1.BackColor = &H8000000F
Command3.BackColor = &H8000000F
End Sub

Private Sub Command3_Click()
Command5.Visible = True
Command3.BackColor = vbRed
Command2.BackColor = &H8000000F
Command1.BackColor = &H8000000F
End Sub

Private Sub Command5_Click()
If Command1.BackColor = vbRed Then
Image5.Picture = LoadPicture("J:\SDD\group pro\arrow.gif")
Image6.Picture = LoadPicture("J:\SDD\group pro\cross.gif")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

If Command3.BackColor = vbRed Then
Image5.Picture = LoadPicture("J:\SDD\group pro\tick.gif")
Text1.Text = "1"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

If Command2.BackColor = vbRed Then
Image5.Picture = LoadPicture("J:\SDD\group pro\arrow.gif")
Image4.Picture = LoadPicture("J:\SDD\group pro\cross.gif")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command7.Visible = True
Command5.Visible = False
End If

End Sub

Private Sub Command6_Click()
End
End Sub


Private Sub Command7_Click()
If Text1.Text <> "" Then
Form3.Label3.Caption = Val(Form3.Label3.Caption) + "1"
End If
Form1.Hide
Form3.Show

End Sub

