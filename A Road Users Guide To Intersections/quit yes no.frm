VERSION 5.00
Begin VB.Form Form23 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "If you do you will lose what questions you have answed up to this point."
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Are you sure you want to quit the test now?"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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

Form3.Label3.Caption = "0"
Form3.Label6.Caption = "0"
Form3.Label9.Caption = "0"
Form3.Label12.Caption = "0"
Form3.Label15.Caption = "0"
Form3.Label18.Caption = "0"
Form3.Label22.Caption = "0"
Form3.Label24.Caption = "0"
Form3.Label20.Caption = ""

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

Form1.Enabled = True
Form10.Enabled = True
Form11.Enabled = True
Form12.Enabled = True
Form13.Enabled = True
Form14.Enabled = True
Form15.Enabled = True
Form16.Enabled = True
Form17.Enabled = True
Form18.Enabled = True
Form19.Enabled = True
Form20.Enabled = True
Form2.Enabled = True
Form21.Enabled = True
Form22.Enabled = True
Form5.Enabled = True
Form6.Enabled = True
Form7.Enabled = True
Form8.Enabled = True
Form9.Enabled = True
Form1.Hide

Form10.Hide
Form11.Hide
Form12.Hide
Form13.Hide
Form14.Hide
Form15.Hide
Form16.Hide
Form17.Hide
Form18.Hide
Form19.Hide
Form20.Hide
Form2.Hide
Form21.Hide
Form22.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form23.Hide
Form36.Show
End Sub

Private Sub Command2_Click()
Form1.Enabled = True
Form10.Enabled = True
Form11.Enabled = True
Form12.Enabled = True
Form13.Enabled = True
Form14.Enabled = True
Form15.Enabled = True
Form16.Enabled = True
Form17.Enabled = True
Form18.Enabled = True
Form19.Enabled = True
Form20.Enabled = True
Form2.Enabled = True
Form21.Enabled = True
Form22.Enabled = True
Form5.Enabled = True
Form6.Enabled = True
Form7.Enabled = True
Form8.Enabled = True
Form9.Enabled = True
Form23.Hide

End Sub

Private Sub Form_Load()
Form1.Enabled = False
Form10.Enabled = False
Form11.Enabled = False
Form12.Enabled = False
Form13.Enabled = False
Form14.Enabled = False
Form15.Enabled = False
Form16.Enabled = False
Form17.Enabled = False
Form18.Enabled = False
Form19.Enabled = False
Form20.Enabled = False
Form2.Enabled = False
Form21.Enabled = False
Form22.Enabled = False
Form5.Enabled = False
Form6.Enabled = False
Form7.Enabled = False
Form8.Enabled = False
Form9.Enabled = False
End Sub

