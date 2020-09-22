VERSION 5.00
Begin VB.Form Form27 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introduction"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "View Lessons"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   3000
      Width           =   1935
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
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   1200
      Picture         =   "lesson intro.frx":0000
      Top             =   2040
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   $"lesson intro.frx":6342
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form27.Hide
Form30.Show

End Sub
