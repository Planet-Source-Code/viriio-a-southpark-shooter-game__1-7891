VERSION 5.00
Begin VB.Form frmCharacters 
   Caption         =   "Characters"
   ClientHeight    =   3660
   ClientLeft      =   3495
   ClientTop       =   2415
   ClientWidth     =   3990
   Icon            =   "frmCharacters.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3660
   ScaleWidth      =   3990
   Begin VB.OptionButton all 
      Caption         =   "All Characters"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.OptionButton stan 
      Caption         =   "Stan"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.OptionButton kyle 
      Caption         =   "Kyle"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.OptionButton ike 
      Caption         =   "Ike"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton kenny 
      Caption         =   "Kenny"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton cartman 
      Caption         =   "Cartman"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Choose your Character to hit."
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Characters"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   1440
      Picture         =   "frmCharacters.frx":030A
      Top             =   960
      Width           =   750
   End
   Begin VB.Image Image4 
      Height          =   750
      Left            =   2640
      Picture         =   "frmCharacters.frx":088C
      Top             =   960
      Width           =   750
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   1320
      Picture         =   "frmCharacters.frx":0D64
      Top             =   2400
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   240
      Picture         =   "frmCharacters.frx":1419
      Top             =   2400
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   240
      Picture         =   "frmCharacters.frx":1AB0
      Top             =   960
      Width           =   750
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cartman.Value = True Then
g = "0"
allcharacters = False
End If
If kenny.Value = True Then
g = "1"
allcharacters = False
End If
If ike.Value = True Then
g = "2"
allcharacters = False
End If
If kyle.Value = True Then
g = "3"
allcharacters = False
End If
If stan.Value = True Then
g = "4"
allcharacters = False
End If
If All.Value = True Then
allcharacters = True
End If
Unload Me
End Sub

