VERSION 5.00
Begin VB.Form Form25 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Glossary"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   1800
      TabIndex        =   111
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame27 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   108
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Frame Frame26 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   107
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Frame Frame25 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   106
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Frame Frame24 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   105
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Frame Frame23 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   104
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Frame Frame22 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   103
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Frame Frame21 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   98
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label80 
         Caption         =   "Traffic lights - the box on a pole on the side of the road that tells cars when to go, get ready to go or stop or stop."
         Height          =   495
         Left            =   360
         TabIndex        =   102
         Top             =   1920
         Width           =   5895
      End
      Begin VB.Label Label79 
         Caption         =   "Traffic - other cars, buses, bikes, and other vehicles on the road."
         Height          =   255
         Left            =   360
         TabIndex        =   101
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Label Label78 
         Caption         =   "Three way intersection - where there three roads meet with a small hump in the middle of the road."
         Height          =   375
         Left            =   360
         TabIndex        =   100
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label Label77 
         Caption         =   "T-intersection - an intersection were one road meets another in the shape of a T."
         Height          =   255
         Left            =   360
         TabIndex        =   99
         Top             =   600
         Width           =   5895
      End
   End
   Begin VB.Frame Frame20 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   91
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label76 
         Caption         =   "Stop sign - a sign that tells a vehicle to stop."
         Height          =   255
         Left            =   360
         TabIndex        =   97
         Top             =   2400
         Width           =   4335
      End
      Begin VB.Label Label75 
         Caption         =   "Stop - to not proceed forwards or backwards."
         Height          =   255
         Left            =   360
         TabIndex        =   96
         Top             =   2040
         Width           =   4815
      End
      Begin VB.Label Label74 
         Caption         =   "Speed limit - the maximum speed the car is allowed to go on the certain road."
         Height          =   255
         Left            =   360
         TabIndex        =   95
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Label Label73 
         Caption         =   "Speed - how fast the car is traveling."
         Height          =   255
         Left            =   360
         TabIndex        =   94
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label72 
         Caption         =   "Signal - to indicate or show which direction the car is going."
         Height          =   375
         Left            =   360
         TabIndex        =   93
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label71 
         Caption         =   "Sign - a metal or wooden object telling you something, often the speed limit or to stop."
         Height          =   375
         Left            =   360
         TabIndex        =   92
         Top             =   600
         Width           =   6015
      End
   End
   Begin VB.Frame Frame19 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   85
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label70 
         Caption         =   "Route - the path of the car that it takes to get from one place to the other."
         Height          =   255
         Left            =   480
         TabIndex        =   90
         Top             =   2280
         Width           =   5535
      End
      Begin VB.Label Label69 
         Caption         =   $"Glossary.frx":0000
         Height          =   495
         Left            =   480
         TabIndex        =   89
         Top             =   1680
         Width           =   5775
      End
      Begin VB.Label Label68 
         Caption         =   "Road - the surface which vehicles are driven on."
         Height          =   255
         Left            =   480
         TabIndex        =   88
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label67 
         Caption         =   "Right of way - which car can has the priority when two cars meet."
         Height          =   255
         Left            =   480
         TabIndex        =   87
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label66 
         Caption         =   "Right lane - the right lane of traffic or the lane on the right of the road."
         Height          =   255
         Left            =   480
         TabIndex        =   86
         Top             =   600
         Width           =   5535
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   83
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label65 
         Caption         =   "Queue - a line or pile up of car one behind another waiting for something."
         Height          =   255
         Left            =   480
         TabIndex        =   84
         Top             =   600
         Width           =   5535
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   80
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label64 
         Caption         =   "Policeman - a man that enforces the law who can also control traffic."
         Height          =   255
         Left            =   480
         TabIndex        =   82
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label Label63 
         Caption         =   "Pedestrian - people walking by the side of the road."
         Height          =   375
         Left            =   480
         TabIndex        =   81
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   78
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label62 
         Caption         =   "On coming - something coming towards you."
         Height          =   255
         Left            =   360
         TabIndex        =   79
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   77
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Frame Frame14 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   73
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label61 
         Caption         =   "Mirror - the reflecting glass used to see behind or to the side of the car."
         Height          =   255
         Left            =   360
         TabIndex        =   76
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label60 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   75
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label60 
         Caption         =   "Malfunction - when something goes wrong or does not work"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   74
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   68
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label59 
         Caption         =   "Lights - also know as traffic lights, the three green, yellow, and red circles on the traffic lights."
         Height          =   495
         Left            =   240
         TabIndex        =   72
         Top             =   1680
         Width           =   6135
      End
      Begin VB.Label Label58 
         Caption         =   "Left lane - the left lane of traffic or the lane on the left of the road"
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Label Label57 
         Caption         =   "Lane - the road in which the traffic travels often there are teo lanes of traffic on each road going each direction."
         Height          =   495
         Left            =   240
         TabIndex        =   70
         Top             =   720
         Width           =   6135
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   67
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label56 
         Caption         =   "Kerb - the edge of the road."
         Height          =   375
         Left            =   360
         TabIndex        =   69
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   65
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label55 
         Caption         =   "Join - where two things join or meet like the joining of two roads."
         Height          =   255
         Left            =   480
         TabIndex        =   66
         Top             =   720
         Width           =   5775
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   62
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label54 
         Caption         =   "Intersection - where two or more roads meet."
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label53 
         Caption         =   "Indicator - the flashing light on the side of the car used to say which direction the car is turning."
         Height          =   495
         Left            =   360
         TabIndex        =   63
         Top             =   600
         Width           =   6015
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   60
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label52 
         Caption         =   "Hazard - warning or danger."
         Height          =   375
         Left            =   360
         TabIndex        =   61
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   56
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label51 
         Caption         =   "Give way sign - a sign that tells you to give way."
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label50 
         Caption         =   "Give way - to let another car go before you."
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label49 
         Caption         =   "Garage - a space or room where cars are stored."
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   54
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label48 
         Caption         =   "Flash - to put on and off continuously (indicators flash)."
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   600
         Width           =   4095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   51
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label47 
         Caption         =   "Exit - to leave or proceed out of something."
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label46 
         Caption         =   "Enter - to go into or proceed onto something."
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   47
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label45 
         Caption         =   "Driving - the act of driving a car or controlling it through streets and roads."
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Label Label44 
         Caption         =   "Diver - the person driving the car."
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label43 
         Caption         =   "Direction - the direction or bearing the car is traveling: north, south, east or west."
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   600
         Width           =   6015
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   41
      Top             =   3720
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Label Label42 
         Caption         =   "Cyclist - a person traveling on a bike."
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label Label41 
         Caption         =   "Crossroad - an intersection where four roads meet of two rods cross each other."
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1800
         Width           =   5775
      End
      Begin VB.Label Label40 
         Caption         =   "Commercial vehicles - public transport, trucks and vans and council owned vehicles."
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1440
         Width           =   6135
      End
      Begin VB.Label Label39 
         Caption         =   "Center lane - the central lane of traffic or the lane in the middle of the left and right lanes."
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label38 
         Caption         =   "Car - an automobile that moves transporting passengers around the place."
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   5895
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   32
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   31
      Top             =   3720
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Label Label37 
         Caption         =   "Brake pedal - the pedal on the car that slows it down."
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label Label36 
         Caption         =   "Brake - to slow down decease the speed of the car. "
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label35 
         Caption         =   "Blind spot - the area of the road that is not visible using the mirrors."
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label Label34 
         Caption         =   "Bike - a two wheeled vehicle."
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   6615
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   13
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label22 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label23 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label24 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label26 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   2
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   1
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label81 
         BackStyle       =   0  'Transparent
         Caption         =   " -      -       -      -        -       -       -      -        -       -       -       -       -      -        -       -       -"
         Height          =   255
         Left            =   240
         TabIndex        =   109
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label82 
         Caption         =   "-       -       -       -       -       -       -       -       -       -       -"
         Height          =   255
         Left            =   1800
         TabIndex        =   110
         Top             =   840
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Label Label33 
         Caption         =   "Approach - to come up to something that is stationary, such as a roundabout."
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   5895
      End
      Begin VB.Label Label32 
         Caption         =   "Amber - the yellow color on the middle light of a traffic light."
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label31 
         Caption         =   "Alternate - to change from one to another."
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label30 
         Caption         =   "Accelerator pedal - the pedal that speeds the car up."
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label29 
         Caption         =   "Accelerate - to speed up or go faster than you are going."
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   240
      X2              =   6720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   240
      X2              =   6720
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Glossary"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   28
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label28 
      Caption         =   $"Glossary.frx":0087
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   27
      Top             =   2520
      Width           =   6375
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form25.Hide

End Sub

Private Sub Command2_Click()
Form25.PrintForm

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80000012
Label3.ForeColor = &H80000012
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label7.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label9.ForeColor = &H80000012
Label10.ForeColor = &H80000012
Label11.ForeColor = &H80000012
Label12.ForeColor = &H80000012
Label13.ForeColor = &H80000012
Label14.ForeColor = &H80000012
Label15.ForeColor = &H80000012
Label16.ForeColor = &H80000012
Label17.ForeColor = &H80000012
Label18.ForeColor = &H80000012
Label19.ForeColor = &H80000012
Label20.ForeColor = &H80000012
Label21.ForeColor = &H80000012
Label22.ForeColor = &H80000012
Label23.ForeColor = &H80000012
Label24.ForeColor = &H80000012
Label25.ForeColor = &H80000012
Label26.ForeColor = &H80000012
Label27.ForeColor = &H80000012
End Sub

Private Sub Label10_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame2.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame10.Visible = True
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = vbRed
End Sub

Private Sub Label11_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame2.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame11.Visible = True
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = vbRed
End Sub

Private Sub Label12_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame2.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame12.Visible = True
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = vbRed
End Sub

Private Sub Label13_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame2.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame13.Visible = True
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = vbRed
End Sub

Private Sub Label14_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame2.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame14.Visible = True
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = vbRed
End Sub

Private Sub Label15_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame2.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame15.Visible = True
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.ForeColor = vbRed
End Sub

Private Sub Label16_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame2.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame16.Visible = True
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.ForeColor = vbRed
End Sub

Private Sub Label17_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame2.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame17.Visible = True
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.ForeColor = vbRed
End Sub

Private Sub Label18_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame2.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame18.Visible = True
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.ForeColor = vbRed
End Sub

Private Sub Label19_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame2.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame19.Visible = True
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.ForeColor = vbRed
End Sub

Private Sub Label2_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame2.Visible = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbRed

End Sub

Private Sub Label20_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame2.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame20.Visible = True
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.ForeColor = vbRed
End Sub

Private Sub Label21_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame2.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame21.Visible = True
End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label21.ForeColor = vbRed
End Sub

Private Sub Label22_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame2.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame22.Visible = True
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label22.ForeColor = vbRed
End Sub

Private Sub Label23_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame2.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame23.Visible = True
End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label23.ForeColor = vbRed
End Sub

Private Sub Label24_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame2.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame24.Visible = True
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label24.ForeColor = vbRed
End Sub

Private Sub Label25_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame2.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame25.Visible = True
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label25.ForeColor = vbRed
End Sub

Private Sub Label26_Click()
Frame27.Visible = False
Frame2.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame26.Visible = True
End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label26.ForeColor = vbRed
End Sub

Private Sub Label27_Click()
Frame2.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame27.Visible = True
End Sub

Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label27.ForeColor = vbRed
End Sub

Private Sub Label3_Click()
Frame2.Visible = False
Frame2.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame2.Visible = False
Frame3.Visible = True

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed
End Sub

Private Sub Label4_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
End Sub

Private Sub Label5_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame5.Visible = True
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbRed
End Sub

Private Sub Label6_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame6.Visible = True
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbRed
End Sub

Private Sub Label7_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame8.Visible = False
Frame2.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame7.Visible = True
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbRed
End Sub

Private Sub Label8_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame9.Visible = False
Frame2.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame8.Visible = True
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = vbRed
End Sub

Private Sub Label81_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80000012
Label3.ForeColor = &H80000012
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label7.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label9.ForeColor = &H80000012
Label10.ForeColor = &H80000012
Label11.ForeColor = &H80000012
Label12.ForeColor = &H80000012
Label13.ForeColor = &H80000012
Label14.ForeColor = &H80000012
Label15.ForeColor = &H80000012
Label16.ForeColor = &H80000012
Label17.ForeColor = &H80000012
Label18.ForeColor = &H80000012
Label19.ForeColor = &H80000012
Label20.ForeColor = &H80000012
Label21.ForeColor = &H80000012
Label22.ForeColor = &H80000012
Label23.ForeColor = &H80000012
Label24.ForeColor = &H80000012
Label25.ForeColor = &H80000012
Label26.ForeColor = &H80000012
Label27.ForeColor = &H80000012
End Sub

Private Sub Label82_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H80000012
Label3.ForeColor = &H80000012
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label7.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label9.ForeColor = &H80000012
Label10.ForeColor = &H80000012
Label11.ForeColor = &H80000012
Label12.ForeColor = &H80000012
Label13.ForeColor = &H80000012
Label14.ForeColor = &H80000012
Label15.ForeColor = &H80000012
Label16.ForeColor = &H80000012
Label17.ForeColor = &H80000012
Label18.ForeColor = &H80000012
Label19.ForeColor = &H80000012
Label20.ForeColor = &H80000012
Label21.ForeColor = &H80000012
Label22.ForeColor = &H80000012
Label23.ForeColor = &H80000012
Label24.ForeColor = &H80000012
Label25.ForeColor = &H80000012
Label26.ForeColor = &H80000012
Label27.ForeColor = &H80000012
End Sub

Private Sub Label9_Click()
Frame27.Visible = False
Frame26.Visible = False
Frame25.Visible = False
Frame24.Visible = False
Frame23.Visible = False
Frame22.Visible = False
Frame21.Visible = False
Frame20.Visible = False
Frame19.Visible = False
Frame18.Visible = False
Frame17.Visible = False
Frame16.Visible = False
Frame15.Visible = False
Frame14.Visible = False
Frame13.Visible = False
Frame12.Visible = False
Frame11.Visible = False
Frame10.Visible = False
Frame2.Visible = False
Frame8.Visible = False
Frame7.Visible = False
Frame6.Visible = False
Frame5.Visible = False
Frame4.Visible = False
Frame3.Visible = False
Frame9.Visible = True
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = vbRed
End Sub
