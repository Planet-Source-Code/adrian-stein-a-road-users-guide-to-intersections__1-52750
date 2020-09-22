VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introduction"
   ClientHeight    =   3840
   ClientLeft      =   5370
   ClientTop       =   3855
   ClientWidth     =   4290
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Start Test"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   1080
      Picture         =   "test intro.frx":0000
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   $"test intro.frx":6342
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Inroduction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Hide
Form6.Show
End Sub

