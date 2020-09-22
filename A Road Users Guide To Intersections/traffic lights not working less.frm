VERSION 5.00
Begin VB.Form Form31 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Malfunctions Lesson"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10455
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
      Left            =   3240
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   4200
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   3600
      Top             =   960
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   6360
      Picture         =   "traffic lights not working less.frx":0000
      ScaleHeight     =   2040
      ScaleWidth      =   1020
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   6360
      Picture         =   "traffic lights not working less.frx":6CA2
      ScaleHeight     =   2040
      ScaleWidth      =   1020
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   7320
      Picture         =   "traffic lights not working less.frx":D944
      ScaleHeight     =   2040
      ScaleWidth      =   1020
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Lesson"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Lesson"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
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
   Begin VB.CommandButton Command5 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Continue"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "(Traffic Lights Not Working)"
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
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   3915
      Left            =   5280
      Picture         =   "traffic lights not working less.frx":145E6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4410
   End
   Begin VB.Label Label2 
      Caption         =   "Malfunctions"
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
      TabIndex        =   7
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   9495
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bob As Integer
Dim steve As Integer
Private Sub Command1_Click()
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Image1.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Label1.Caption = "Malfunctioning traffic lights occur when there is a power failure or the lights are damaged in some way. Malfunctioning traffic lights are rare, some even have back-up power sources in heavy built up areas. But when a malfunction occurs, it is important to know the rules as it is easy to become part of a collision."

Form31.Hide
Form30.Show
End Sub

Private Sub Command2_Click()
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Image1.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Label1.Caption = "Malfunctioning traffic lights occur when there is a power failure or the lights are damaged in some way. Malfunctioning traffic lights are rare, some even have back-up power sources in heavy built up areas. But when a malfunction occurs, it is important to know the rules as it is easy to become part of a collision."

Form31.Hide
Form32.Show

End Sub

Private Sub Command3_Click()
Form25.Show
End Sub

Private Sub Command4_Click()
Image1.Visible = False
Picture1.Visible = True
Label1.Caption = "As you can see, these two sets of lights have malfunctioned. When coming to a malfunctioning set of traffic lights, first, come to a complete stop. Then look at the other cars at the traffic lights. You must give way to the cars on your right. If there are four cars waiting at an intersection, you are allowed to slowly move out, watching for the other three cars doing the same. Press continue."
Command4.Visible = False
Command6.Visible = True
Picture2.Visible = True
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Command5_Click()
Command7.Visible = False
Command4.Visible = True
Command6.Visible = False
Image1.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Label1.Caption = "Malfunctioning traffic lights occur when there is a power failure or the lights are damaged in some way. Malfunctioning traffic lights are rare, some even have back-up power sources in heavy built up areas. But when a malfunction occurs, it is important to know the rules as it is easy to become part of a collision."

Form31.Hide
Form36.Show
End Sub

Private Sub Command6_Click()
Command6.Visible = False
Command7.Visible = True
Label1.Caption = " After observing a malfunctioning traffic light, it is good moral to ring and inform the local council that the particular lights aren't working. This can help to prevent further crashes. Press next lesson"
End Sub

Private Sub Command7_Click()
Command7.Visible = False
Command4.Visible = True
Image1.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Label1.Caption = "Malfunctioning traffic lights occur when there is a power failure or the lights are damaged in some way. Malfunctioning traffic lights are rare, some even have back-up power sources in heavy built up areas. But when a malfunction occurs, it is important to know the rules as it is easy to become part of a collision."

End Sub

Private Sub Command8_Click()
Form31.PrintForm

End Sub

Private Sub Form_Load()
Label1.Caption = "Malfunctioning traffic lights occur when there is a power failure or the lights are damaged in some way. Malfunctioning traffic lights are rare, some even have back-up power sources in heavy built up areas. But when a malfunction occurs, it is important to know the rules as it is easy to become part of a collision."

End Sub

Private Sub Timer1_Timer()
Picture3.Visible = True
End Sub

Private Sub Timer2_Timer()
Picture3.Visible = False
End Sub

Private Sub Timer3_Timer()
Dim bob As Integer
Randomize
bob = Rnd * 10
Picture4.Left = Picture4.Left + bob
End Sub

Private Sub Timer4_Timer()

Randomize
bob = Rnd * 100
If steve < 5 Then
    Picture4.Left = Picture4.Left + bob
    steve = steve + 1
End If
    
If steve > 5 Then
    Picture4.Left = Picture4.Left - bob
End If
If steve = 9 Then
    steve = 0
End If
End Sub
