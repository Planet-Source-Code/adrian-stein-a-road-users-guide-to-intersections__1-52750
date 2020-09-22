VERSION 5.00
Begin VB.Form Form28 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WELCOME"
   ClientHeight    =   4530
   ClientLeft      =   3180
   ClientTop       =   2430
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   120
      Top             =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Skip Intro"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   120
      Top             =   2160
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   5640
      Picture         =   "animation.frx":0000
      ScaleHeight     =   1935
      ScaleWidth      =   3255
      TabIndex        =   13
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label11 
      Height          =   735
      Left            =   1320
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1440
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label9 
      Caption         =   "AND"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Made By"
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
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "A Road Users Guide To Intersections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim bob As Integer
Dim boby As Integer
Dim Add As Integer


Private Sub Command1_Click()
sndPlaySound App.Path + "v.wav", 1
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
bob = 0
boby = 0
Add = 0
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Form28.Hide
Form36.Show
Picture1.Visible = True
End Sub

Private Sub Command2_Click()
sndPlaySound App.Path + "v.wav", 1
Timer1.Enabled = False
bob = 0
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False

End Sub


Private Sub Command3_Click()
sndPlaySound App.Path + "v.wav", 1
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
bob = 0
boby = 0
Add = 0
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Form28.Hide
Form36.Show

End Sub

Private Sub Form_Load()
'sndPlaySound App.Path + "\boy.wav", 1
bob = 0
sndPlaySound App.Path + "\boy.wav", 1
Timer1.Enabled = True
Timer3.Enabled = True

End Sub


Private Sub StopHuhk_Click()
sndPlaySound App.Path + "v.wav", 1
End Sub

Private Sub Timer1_Timer()
'Dim bob As Integer
'Randomize
'bob = 0 + Rnd



bob = bob + 1


'twice
If bob = 1 Then
Timer1.Interval = 80
End If

If bob = 1 Then
Label4.Visible = True
End If
If bob = 2 Then
Label4.Visible = False
End If



If bob = 3 Then
Label3.Visible = True
End If
If bob = 4 Then
Label3.Visible = False
End If

If bob = 5 Then
Label6.Visible = True
End If
If bob = 6 Then
Label6.Visible = False
End If

If bob = 7 Then
Label2.Visible = True
End If
If bob = 8 Then
Label2.Visible = False
End If

If bob = 9 Then
Label4.Visible = True
End If
If bob = 10 Then
Label4.Visible = False
End If

If bob = 11 Then
Label5.Visible = True
End If
If bob = 12 Then
Label5.Visible = False
End If

Label9.Visible = False

'first

If bob = 13 Then
Label1.Visible = True
End If
If bob = 14 Then
Label1.Visible = False
End If

If bob = 15 Then
Label4.Visible = True
End If
If bob = 16 Then
Label4.Visible = False
End If

If bob = 17 Then
Label6.Visible = True
End If
If bob = 18 Then
Label6.Visible = False
End If

If bob = 19 Then
Label2.Visible = True
End If
If bob = 20 Then
Label2.Visible = False
End If

If bob = 21 Then
Label3.Visible = True
End If
If bob = 22 Then
Label3.Visible = False
End If

If bob = 23 Then
Label5.Visible = True
End If
If bob = 24 Then
Label5.Visible = False
End If

'once
If bob = 25 Then
Timer1.Interval = 50
End If


If bob = 26 Then
Label6.Visible = True
End If
If bob = 27 Then
Label6.Visible = False
End If

If bob = 28 Then
Label4.Visible = True
End If
If bob = 29 Then
Label4.Visible = False
End If

If bob = 30 Then
Label1.Visible = True
End If
If bob = 31 Then
Label1.Visible = False
End If

If bob = 32 Then
Label3.Visible = True
End If
If bob = 33 Then
Label3.Visible = False
End If

If bob = 34 Then
Label5.Visible = True
End If
If bob = 35 Then
Label5.Visible = False
End If

If bob = 36 Then
Label6.Visible = True
End If
If bob = 37 Then
Label6.Visible = False
End If

'twice
If bob = 38 Then
Timer1.Interval = 30
End If

If bob = 39 Then
Label4.Visible = True
End If
If bob = 40 Then
Label4.Visible = False
End If

If bob = 41 Then
Label1.Visible = True
End If
If bob = 42 Then
Label1.Visible = False
End If

If bob = 43 Then
Label6.Visible = True
End If
If bob = 44 Then
Label6.Visible = False
End If

If bob = 45 Then
Label2.Visible = True
End If
If bob = 46 Then
Label2.Visible = False
End If

If bob = 47 Then
Label4.Visible = True
End If
If bob = 48 Then
Label4.Visible = False
End If

If bob = 49 Then
Label5.Visible = True
End If
If bob = 50 Then
Label5.Visible = False
End If

'more
If bob = 51 Then
Timer1.Interval = 10
End If

If bob = 52 Then
Label4.Visible = True
End If
If bob = 53 Then
Label4.Visible = False
End If

If bob = 54 Then
Label1.Visible = True
End If
If bob = 55 Then
Label1.Visible = False
End If

If bob = 56 Then
Label6.Visible = True
End If
If bob = 57 Then
Label6.Visible = False
End If

If bob = 58 Then
Label2.Visible = True
End If
If bob = 59 Then
Label2.Visible = False
End If

If bob = 60 Then
Label4.Visible = True
End If
If bob = 61 Then
Label4.Visible = False
End If

If bob = 62 Then
Label5.Visible = True
End If
If bob = 63 Then
Label5.Visible = False
End If


'last

If bob = 64 Then
Timer1.Interval = 1
End If

If bob = 65 Then
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
End If

'end

If bob = 66 Then
Label3.Visible = False
End If


If bob = 67 Then
Label1.Visible = False
End If

If bob = 68 Then
Label6.Visible = False
End If

If bob = 69 Then
Label2.Visible = False
End If

If bob = 70 Then
Label4.Visible = False
End If

If bob = 71 Then
Label5.Visible = False
End If

If bob = 72 Then
Timer1.Interval = 100
End If


If bob = 72 Then
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
End If

If bob = 73 Then
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
End If

If bob = 74 Then
Label9.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
End If

If bob = 75 Then
Label9.Visible = True
'Label1.ForeColor = vbBlack
'Label2.ForeColor = vbBlack
'Label3.ForeColor = vbBlack
'Label4.ForeColor = vbBlack
'Label5.ForeColor = vbBlack
'Label6.ForeColor = vbBlack

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
End If

If bob = 76 Then
Label9.Visible = True
Label9.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
End If

If bob = 77 Then
Label9.Visible = True
End If

If bob = 78 Then
Label9.Visible = True
End If
If bob = 79 Then
Label9.Visible = True
End If
If bob = 80 Then
Label9.Visible = True
End If
If bob = 81 Then
Label9.Visible = True
End If

If bob = 81 Then
Label1.Caption = "S"
Label2.Caption = "t"
Label3.Caption = "u"
Label4.Caption = "a"
Label5.Caption = "r"
Label6.Caption = "t"
bob = 0
Timer2.Enabled = True
End If


End Sub



Private Sub Timer2_Timer()

boby = boby + 1

If boby = 59 Then
Label1.Caption = ""
Label2.Caption = "A"
Label3.Caption = "l"
Label4.Caption = "e"
Label5.Caption = "x"
Label6.Caption = ""
End If

If boby = 112 Then
Timer1.Enabled = False
End If

If boby = 114 Then
Label9.Visible = False
Label8.Visible = False
Label10.Visible = True
Label11.Visible = True

End If

If boby = 120 Then
Command1.Visible = True
Timer2.Enabled = False
Timer3.Enabled = False
Picture1.Visible = False
End If

End Sub

Private Sub Timer3_Timer()
Add = Add + 1

If Add = 1 Then
Picture1.Left = Picture1.Left - 100
Add = 0
End If

If Picture1.Left <= -3260 Then
Picture1.Left = 5640
End If

End Sub
