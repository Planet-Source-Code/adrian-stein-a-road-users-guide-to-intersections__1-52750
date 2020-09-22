VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Results"
   ClientHeight    =   5415
   ClientLeft      =   5265
   ClientTop       =   2685
   ClientWidth     =   5070
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Print"
      Height          =   375
      Left            =   3480
      TabIndex        =   28
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Lessons"
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Redo Test"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label25 
      Caption         =   "\20"
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "\2"
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label19 
      Caption         =   "\3"
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "\3"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "\3"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "\3"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "\3"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "\3"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label24 
      Caption         =   "0"
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label22 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label21 
      Caption         =   "Vocabulary"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   20
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label18 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label17 
      Caption         =   "Three Way Intersections"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "T-Intersections"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Cross Roads"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "Traffic Lights Not Working"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Traffic Lights"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Roundabouts"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form36.Show
End Sub

Private Sub Command2_Click()

Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form10.Text1.Text = ""
Form10.Text2.Text = ""
Form11.Text1.Text = ""
Form11.Text2.Text = ""
Form12.Text1.Text = ""
Form12.Text2.Text = ""
Form13.Text1.Text = ""
Form13.Text2.Text = ""
Form14.Text1.Text = ""
Form14.Text2.Text = ""
Form15.Text1.Text = ""
Form15.Text2.Text = ""
Form16.Text1.Text = ""
Form16.Text2.Text = ""
Form17.Text1.Text = ""
Form17.Text2.Text = ""
Form18.Text1.Text = ""
Form18.Text2.Text = ""
Form19.Text1.Text = ""
Form19.Text2.Text = ""
Form2.Text1.Text = ""
Form2.Text2.Text = ""
Form20.Text1.Text = ""
Form20.Text2.Text = ""
Form21.Text1.Text = ""
Form21.Text2.Text = ""
Form22.Text1.Text = ""
Form22.Text2.Text = ""
Form5.Text1.Text = ""
Form5.Text2.Text = ""
Form6.Text1.Text = ""
Form6.Text2.Text = ""
Form7.Text1.Text = ""
Form7.Text2.Text = ""
Form8.Text1.Text = ""
Form8.Text2.Text = ""
Form9.Text1.Text = ""
Form9.Text2.Text = ""

Label3.Caption = "0"
Label6.Caption = "0"
Label9.Caption = "0"
Label12.Caption = "0"
Label15.Caption = "0"
Label18.Caption = "0"
Label22.Caption = "0"
Label24.Caption = "0"
Label20.Caption = ""

Form1.Command1.Enabled = True
Form1.Command2.Enabled = True
Form1.Command3.Enabled = True
Form1.Command4.Enabled = True
Form1.Command5.Enabled = True
Form1.Command7.Enabled = True
Form10.Command1.Enabled = True
Form10.Command2.Enabled = True
Form10.Command3.Enabled = True
Form10.Command4.Enabled = True
Form10.Command5.Enabled = True
Form10.Command7.Enabled = True
Form11.Command1.Enabled = True
Form11.Command2.Enabled = True
Form11.Command3.Enabled = True
Form11.Command4.Enabled = True
Form11.Command5.Enabled = True
Form1.Command7.Enabled = True
Form12.Command1.Enabled = True
Form12.Command2.Enabled = True
Form12.Command3.Enabled = True
Form12.Command4.Enabled = True
Form12.Command5.Enabled = True
Form12.Command7.Enabled = True
Form13.Command1.Enabled = True
Form13.Command2.Enabled = True
Form13.Command3.Enabled = True
Form13.Command4.Enabled = True
Form13.Command5.Enabled = True
Form1.Command7.Enabled = True
Form14.Command1.Enabled = True
Form14.Command2.Enabled = True
Form14.Command3.Enabled = True
Form14.Command4.Enabled = True
Form14.Command5.Enabled = True
Form14.Command7.Enabled = True
Form15.Command1.Enabled = True
Form15.Command2.Enabled = True
Form15.Command3.Enabled = True
Form15.Command4.Enabled = True
Form15.Command5.Enabled = True
Form15.Command7.Enabled = True
Form16.Command1.Enabled = True
Form16.Command2.Enabled = True
Form16.Command3.Enabled = True
Form16.Command4.Enabled = True
Form16.Command5.Enabled = True
Form16.Command7.Enabled = True
Form17.Command1.Enabled = True
Form17.Command2.Enabled = True
Form17.Command3.Enabled = True
Form17.Command4.Enabled = True
Form17.Command5.Enabled = True
Form17.Command7.Enabled = True
Form18.Command1.Enabled = True
Form18.Command2.Enabled = True
Form18.Command3.Enabled = True
Form18.Command4.Enabled = True
Form18.Command5.Enabled = True
Form18.Command7.Enabled = True
Form19.Command1.Enabled = True
Form19.Command2.Enabled = True
Form19.Command3.Enabled = True
Form19.Command4.Enabled = True
Form19.Command5.Enabled = True
Form19.Command7.Enabled = True
Form2.Command1.Enabled = True
Form2.Command2.Enabled = True
Form2.Command3.Enabled = True
Form2.Command4.Enabled = True
Form2.Command5.Enabled = True
Form2.Command7.Enabled = True
Form21.Command1.Enabled = True
Form21.Command2.Enabled = True
Form21.Command3.Enabled = True
Form21.Command4.Enabled = True
Form21.Command5.Enabled = True
Form21.Command7.Enabled = True
Form20.Command1.Enabled = True
Form20.Command2.Enabled = True
Form20.Command3.Enabled = True
Form20.Command4.Enabled = True
Form20.Command5.Enabled = True
Form20.Command7.Enabled = True
Form22.Command1.Enabled = True
Form22.Command2.Enabled = True
Form22.Command3.Enabled = True
Form22.Command4.Enabled = True
Form22.Command5.Enabled = True
Form22.Command7.Enabled = True
Form5.Command1.Enabled = True
Form5.Command2.Enabled = True
Form5.Command3.Enabled = True
Form5.Command4.Enabled = True
Form5.Command5.Enabled = True
Form1.Command7.Enabled = True
Form6.Command1.Enabled = True
Form6.Command2.Enabled = True
Form6.Command3.Enabled = True
Form6.Command4.Enabled = True
Form6.Command5.Enabled = True
Form6.Command7.Enabled = True
Form7.Command1.Enabled = True
Form7.Command2.Enabled = True
Form7.Command3.Enabled = True
Form7.Command4.Enabled = True
Form7.Command5.Enabled = True
Form7.Command7.Enabled = True
Form8.Command1.Enabled = True
Form8.Command2.Enabled = True
Form8.Command3.Enabled = True
Form8.Command4.Enabled = True
Form8.Command5.Enabled = True
Form8.Command7.Enabled = True
Form9.Command1.Enabled = True
Form9.Command2.Enabled = True
Form9.Command3.Enabled = True
Form9.Command4.Enabled = True
Form9.Command5.Enabled = True
Form9.Command7.Enabled = True

Form1.Command1.BackColor = &H8000000F
Form1.Command2.BackColor = &H8000000F
Form1.Command3.BackColor = &H8000000F
Form10.Command1.BackColor = &H8000000F
Form10.Command2.BackColor = &H8000000F
Form10.Command3.BackColor = &H8000000F
Form11.Command1.BackColor = &H8000000F
Form11.Command2.BackColor = &H8000000F
Form11.Command3.BackColor = &H8000000F
Form12.Command1.BackColor = &H8000000F
Form12.Command2.BackColor = &H8000000F
Form12.Command3.BackColor = &H8000000F
Form13.Command1.BackColor = &H8000000F
Form13.Command2.BackColor = &H8000000F
Form13.Command3.BackColor = &H8000000F
Form14.Command1.BackColor = &H8000000F
Form14.Command2.BackColor = &H8000000F
Form14.Command3.BackColor = &H8000000F
Form15.Command1.BackColor = &H8000000F
Form15.Command2.BackColor = &H8000000F
Form15.Command3.BackColor = &H8000000F
Form16.Command1.BackColor = &H8000000F
Form16.Command2.BackColor = &H8000000F
Form16.Command3.BackColor = &H8000000F
Form17.Command1.BackColor = &H8000000F
Form17.Command2.BackColor = &H8000000F
Form17.Command3.BackColor = &H8000000F
Form18.Command1.BackColor = &H8000000F
Form18.Command2.BackColor = &H8000000F
Form18.Command3.BackColor = &H8000000F
Form19.Command1.BackColor = &H8000000F
Form19.Command2.BackColor = &H8000000F
Form19.Command3.BackColor = &H8000000F
Form2.Command1.BackColor = &H8000000F
Form2.Command2.BackColor = &H8000000F
Form2.Command3.BackColor = &H8000000F
Form20.Command1.BackColor = &H8000000F
Form20.Command2.BackColor = &H8000000F
Form20.Command3.BackColor = &H8000000F
Form21.Command1.BackColor = &H8000000F
Form21.Command2.BackColor = &H8000000F
Form21.Command3.BackColor = &H8000000F
Form22.Command1.BackColor = &H8000000F
Form22.Command2.BackColor = &H8000000F
Form22.Command3.BackColor = &H8000000F
Form5.Command1.BackColor = &H8000000F
Form5.Command2.BackColor = &H8000000F
Form5.Command3.BackColor = &H8000000F
Form6.Command1.BackColor = &H8000000F
Form6.Command2.BackColor = &H8000000F
Form6.Command3.BackColor = &H8000000F
Form7.Command1.BackColor = &H8000000F
Form7.Command2.BackColor = &H8000000F
Form7.Command3.BackColor = &H8000000F
Form8.Command1.BackColor = &H8000000F
Form8.Command2.BackColor = &H8000000F
Form8.Command3.BackColor = &H8000000F
Form9.Command1.BackColor = &H8000000F
Form9.Command2.BackColor = &H8000000F
Form9.Command3.BackColor = &H8000000F

Form1.Image4.Picture = LoadPicture("")
Form1.Image5.Picture = LoadPicture("")
Form1.Image6.Picture = LoadPicture("")
Form10.Image4.Picture = LoadPicture("")
Form10.Image5.Picture = LoadPicture("")
Form10.Image6.Picture = LoadPicture("")
Form11.Image4.Picture = LoadPicture("")
Form11.Image5.Picture = LoadPicture("")
Form11.Image6.Picture = LoadPicture("")
Form12.Image4.Picture = LoadPicture("")
Form12.Image5.Picture = LoadPicture("")
Form12.Image6.Picture = LoadPicture("")
Form13.Image4.Picture = LoadPicture("")
Form13.Image5.Picture = LoadPicture("")
Form13.Image6.Picture = LoadPicture("")
Form14.Image4.Picture = LoadPicture("")
Form14.Image5.Picture = LoadPicture("")
Form14.Image6.Picture = LoadPicture("")
Form15.Image4.Picture = LoadPicture("")
Form15.Image5.Picture = LoadPicture("")
Form15.Image6.Picture = LoadPicture("")
Form16.Image4.Picture = LoadPicture("")
Form16.Image5.Picture = LoadPicture("")
Form16.Image6.Picture = LoadPicture("")
Form17.Image4.Picture = LoadPicture("")
Form17.Image5.Picture = LoadPicture("")
Form17.Image6.Picture = LoadPicture("")
Form18.Image4.Picture = LoadPicture("")
Form18.Image5.Picture = LoadPicture("")
Form18.Image6.Picture = LoadPicture("")
Form19.Image4.Picture = LoadPicture("")
Form19.Image5.Picture = LoadPicture("")
Form19.Image6.Picture = LoadPicture("")
Form21.Image4.Picture = LoadPicture("")
Form21.Image5.Picture = LoadPicture("")
Form21.Image6.Picture = LoadPicture("")
Form22.Image4.Picture = LoadPicture("")
Form22.Image5.Picture = LoadPicture("")
Form22.Image6.Picture = LoadPicture("")
Form2.Image4.Picture = LoadPicture("")
Form2.Image5.Picture = LoadPicture("")
Form2.Image6.Picture = LoadPicture("")
Form20.Image4.Picture = LoadPicture("")
Form20.Image5.Picture = LoadPicture("")
Form20.Image6.Picture = LoadPicture("")
Form5.Image4.Picture = LoadPicture("")
Form5.Image5.Picture = LoadPicture("")
Form5.Image6.Picture = LoadPicture("")
Form6.Image4.Picture = LoadPicture("")
Form6.Image5.Picture = LoadPicture("")
Form6.Image6.Picture = LoadPicture("")
Form7.Image4.Picture = LoadPicture("")
Form7.Image5.Picture = LoadPicture("")
Form7.Image6.Picture = LoadPicture("")
Form8.Image4.Picture = LoadPicture("")
Form8.Image5.Picture = LoadPicture("")
Form8.Image6.Picture = LoadPicture("")
Form9.Image4.Picture = LoadPicture("")
Form9.Image5.Picture = LoadPicture("")
Form9.Image6.Picture = LoadPicture("")

Form1.Command7.Visible = False
Form10.Command7.Visible = False
Form11.Command7.Visible = False
Form12.Command7.Visible = False
Form13.Command7.Visible = False
Form14.Command7.Visible = False
Form15.Command7.Visible = False
Form16.Command7.Visible = False
Form17.Command7.Visible = False
Form18.Command7.Visible = False
Form19.Command7.Visible = False
Form20.Command7.Visible = False
Form2.Command7.Visible = False
Form21.Command7.Visible = False
Form22.Command7.Visible = False
Form5.Command7.Visible = False
Form6.Command7.Visible = False
Form7.Command7.Visible = False
Form8.Command7.Visible = False
Form9.Command7.Visible = False

Form1.Label11.Caption = ""
Form10.Label11.Caption = ""
Form11.Label11.Caption = ""
Form12.Label11.Caption = ""
Form13.Label11.Caption = ""
Form14.Label11.Caption = ""
Form15.Label11.Caption = ""
Form16.Label11.Caption = ""
Form17.Label11.Caption = ""
Form18.Label11.Caption = ""
Form19.Label11.Caption = ""
Form2.Label11.Caption = ""
Form20.Label11.Caption = ""
Form21.Label11.Caption = ""
Form22.Label11.Caption = ""
Form5.Label11.Caption = ""
Form6.Label11.Caption = ""
Form7.Label11.Caption = ""
Form8.Label11.Caption = ""
Form9.Label11.Caption = ""

Form3.Hide
Form4.Show
End Sub

Private Sub Command3_Click()
Form3.Hide
Form30.Show

End Sub

Private Sub Command4_Click()
Form3.PrintForm
End Sub

Private Sub Label12_Change()
Label24.Caption = ""
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)
End Sub
Private Sub Label15_Change()
Label24.Caption = ""
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)
End Sub
Private Sub Label18_Change()
Label24.Caption = ""
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)
End Sub

Private Sub Label22_Change()
Label24.Caption = ""
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)
End Sub

Private Sub Label24_Change()
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)


If Val(Label24.Caption) < 17 Then
Label20.Caption = "FAILED"
Else
Label20.Caption = "PASSED"
End If
End Sub


Private Sub Label3_Change()
Label24.Caption = ""
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)
End Sub
Private Sub Label6_Change()
Label24.Caption = ""
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)
End Sub
Private Sub Label9_Change()
Label24.Caption = ""
Label24.Caption = Val(Label3.Caption) + Val(Label6.Caption) + Val(Label9.Caption) + Val(Label12.Caption) + Val(Label15.Caption) + Val(Label18.Caption) + Val(Label22.Caption)
End Sub
