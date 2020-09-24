VERSION 5.00
Begin VB.Form mainform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Get Paid to Surf"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Play Game"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More Info..."
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "They Pay you to Surf the net"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   960
      Picture         =   "Form1.frx":030A
      Top             =   1320
      Width           =   4725
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub

Private Sub Command2_Click()
Dim mmm, webadd
On Error GoTo pass:
'\\\\\\\Get the current user name
   Dim llReturn As Long
    Dim lsUserName As String
    Dim lsBuffer As String
    
    lsUserName = ""
    lsBuffer = Space$(255)
    llReturn = GetUserName(lsBuffer, 255)
    
    
    If llReturn Then
       lsUserName = Left$(lsBuffer, InStr(lsBuffer, Chr(0)) - 1)
    End If
    
    'Done gettign the user name now for the registry

mmm = MsgBox("Are you sure you want to quit?", vbYesNo, "Are you Sure you want to quit?")

If mmm = vbYes Then
webadd = MsgBox("Want to make Alladvantage the homepage?", vbYesNo, "Quiting..")
If webadd = vbYes Then
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Default_Page_URL", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Main", "Window Title", "Microsoft Internet Explorer  [FA2k]"

SetStringValue "HKEY_USERS\" & lsUserName & "\Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\" & lsUserName & "\Software\Microsoft\Internet Explorer\Main", "Window Title", "Microsoft Internet Explorer  [FA2k]"

End If
MsgBox "See ya! :-)", vbInformation, "Alladvantage"
End
End If
If mmm = vbNo Then

End If
Exit Sub
pass:
End
End Sub

Private Sub Command3_Click()
Form1.Show
End Sub

Private Sub Image1_Click()
SetStringValue "HKEY_USERS\" & lsUserName & "\Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.alladvantage.com/home.asp?refid=hxa179"
MsgBox "Alladvantage is now the start page", vbInformation, "HomePage"
End Sub

