VERSION 5.00
Begin VB.Form Form26 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Read Me File"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print "
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close "
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   7440
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7455
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Read Me File"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "This program is easy to use and can be very useful to learn about different types off intersections."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   $"READ_M~1.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   $"READ_M~1.frx":027D
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   $"READ_M~1.frx":0401
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   5175
   End
   Begin VB.Label Label6 
      Caption         =   $"READ_M~1.frx":05CA
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   5175
   End
   Begin VB.Label Label7 
      Caption         =   "The makers: Adrian Stein, Alex White and Stuart Morrison Copyright 2003."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   5295
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form26.Hide
Form36.Show

End Sub

Private Sub Command2_Click()
Form26.PrintForm
End Sub

