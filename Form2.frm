VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Money Stuff"
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   4200
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808000&
      Caption         =   "Make Alladvantage the HomePage"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Command1_Click()
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

MsgBox "See ya! :-)", vbInformation, "Alladvantage"
End
End If
If mmm = vbNo Then

End If
If Check1.Value = Checked Then
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Default_Page_URL", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Main", "Window Title", "Microsoft Internet Explorer  [FA2k]"

SetStringValue "HKEY_USERS\" & lsUserName & "\Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\" & lsUserName & "\Software\Microsoft\Internet Explorer\Main", "Window Title", "Microsoft Internet Explorer  [FA2k]"

End If
Exit Sub
pass:
End Sub

Private Sub Form_Load()

End Sub
