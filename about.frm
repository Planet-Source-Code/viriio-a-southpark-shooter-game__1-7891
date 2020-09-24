VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help on this Game"
   ClientHeight    =   3510
   ClientLeft      =   2955
   ClientTop       =   1455
   ClientWidth     =   4860
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "about.frx":030A
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4860
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"about.frx":0614
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "VIRIIO's Game"
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
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   3480
      Picture         =   "about.frx":06C5
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "The more times you hit Cartman the faster he pops up and down. If there's any questions email me: grant_84@mailcity.com "
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"about.frx":0C08
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
