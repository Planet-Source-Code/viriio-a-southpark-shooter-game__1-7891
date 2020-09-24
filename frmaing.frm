VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VIRIIO's "
   ClientHeight    =   4830
   ClientLeft      =   2445
   ClientTop       =   1920
   ClientWidth     =   5640
   Icon            =   "frmaing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmaing.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   4830
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000C000&
      Caption         =   "Make Alladvantage the HomePage"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.Timer moletimer 
      Index           =   0
      Interval        =   3000
      Left            =   3720
      Top             =   2520
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "Start Game"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image imghalf 
      Height          =   750
      Index           =   4
      Left            =   8640
      Picture         =   "frmaing.frx":0614
      Top             =   360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghole 
      Height          =   750
      Index           =   4
      Left            =   8640
      Picture         =   "frmaing.frx":0BC4
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgup 
      Height          =   750
      Index           =   4
      Left            =   8640
      Picture         =   "frmaing.frx":0FC0
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgmiss 
      Height          =   750
      Index           =   4
      Left            =   8640
      Picture         =   "frmaing.frx":1657
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgmiss 
      Height          =   750
      Index           =   3
      Left            =   7800
      Picture         =   "frmaing.frx":1B0C
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgup 
      Height          =   750
      Index           =   3
      Left            =   7800
      Picture         =   "frmaing.frx":1FC1
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghole 
      Height          =   750
      Index           =   3
      Left            =   7800
      Picture         =   "frmaing.frx":2676
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghalf 
      Height          =   750
      Index           =   3
      Left            =   7800
      Picture         =   "frmaing.frx":2A72
      Top             =   360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghalf 
      Height          =   750
      Index           =   2
      Left            =   6960
      Picture         =   "frmaing.frx":3022
      Top             =   360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghole 
      Height          =   750
      Index           =   2
      Left            =   6960
      Picture         =   "frmaing.frx":34A2
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgup 
      Height          =   750
      Index           =   2
      Left            =   6960
      Picture         =   "frmaing.frx":389E
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgmiss 
      Height          =   750
      Index           =   2
      Left            =   6960
      Picture         =   "frmaing.frx":3D76
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgmiss 
      Height          =   750
      Index           =   1
      Left            =   6120
      Picture         =   "frmaing.frx":422B
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgup 
      Height          =   750
      Index           =   1
      Left            =   6120
      Picture         =   "frmaing.frx":46E0
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghole 
      Height          =   750
      Index           =   1
      Left            =   6120
      Picture         =   "frmaing.frx":4C62
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghalf 
      Height          =   750
      Index           =   1
      Left            =   6120
      Picture         =   "frmaing.frx":505E
      Top             =   360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgPlay 
      Height          =   480
      Left            =   1320
      Picture         =   "frmaing.frx":557E
      Top             =   4920
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Life:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   3120
      Picture         =   "frmaing.frx":5888
      Top             =   3120
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   120
      Picture         =   "frmaing.frx":5CCA
      Top             =   3960
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   4680
      Picture         =   "frmaing.frx":610C
      Top             =   0
      Width           =   645
   End
   Begin VB.Label gameover 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblhit 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblmiss 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   4440
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.Image imgmiss 
      Height          =   750
      Index           =   0
      Left            =   5280
      Picture         =   "frmaing.frx":654E
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgup 
      Height          =   750
      Index           =   0
      Left            =   5280
      Picture         =   "frmaing.frx":6A03
      Top             =   1680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghalf 
      Height          =   750
      Index           =   0
      Left            =   5280
      Picture         =   "frmaing.frx":6F46
      Top             =   360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imghole 
      Height          =   750
      Index           =   0
      Left            =   5280
      Picture         =   "frmaing.frx":73E3
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgmole 
      Height          =   750
      Index           =   0
      Left            =   1320
      Picture         =   "frmaing.frx":77DF
      Top             =   1560
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label hit 
      BackColor       =   &H0000C000&
      Caption         =   "Hits: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblinstr 
      BackColor       =   &H0000C000&
      Caption         =   "The aim of the game is to hit Cartman, using the mouse."
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   1800
      Picture         =   "frmaing.frx":7BA9
      Top             =   840
      Width           =   645
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnucharacters 
         Caption         =   "C&haracters"
      End
      Begin VB.Menu mnuhide 
         Caption         =   "&Hide"
         Shortcut        =   ^H
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuthisgame 
         Caption         =   "Help on this Game"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnudummy 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "&About"
         Index           =   1
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim paused
Dim pause123, pause1234

Dim hits ' counts no of hits
Dim misses ' and misses
Dim basetime ' timeing
Dim playing 'true or false?
'trayicon
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim T As NOTIFYICONDATA
'trayicon finished

Private Sub cmdgo_Click()
gameover.Visible = False
If paused = True Then
Else
Call startgame
End If
End Sub
Sub startgame()
Dim a
basetime = 4500
cmdgo.Enabled = False
lblinstr.Caption = ""
playing = True
hits = 0
misses = 8
lblhit.Caption = hits
lblmiss.Caption = misses
ReDim molestate(7) ' sets all to zero
For a = 0 To 7
imgmole(a).Picture = imghole(g).Picture
moletimer(a).Interval = 500 + basetime * 4 * Rnd
moletimer(a).Enabled = True
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
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
    'Unload this form. Important: always end with "unload me".
    T.cbSize = Len(T)
    T.hWnd = pichook.hWnd
    T.uId = 1&
    Shell_NotifyIcon NIM_DELETE, T
    End
   g = "noneselected"
   If Check1.Value = Checked Then
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Default_Page_URL", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Internet Explorer\Main", "Window Title", "Microsoft Internet Explorer  [FA2k]"

SetStringValue "HKEY_USERS\" & lsUserName & "\Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.alladvantage.com/home.asp?refid=hxa179"
SetStringValue "HKEY_USERS\" & lsUserName & "\Software\Microsoft\Internet Explorer\Main", "Window Title", "Microsoft Internet Explorer  [FA2k]"

End If
End Sub
Sub Form_Load()
Dim a
Randomize

On Error Resume Next
allcharacters = True
For a = 0 To 9

Load moletimer(a) ' create timer control
moletimer(a).Enabled = False

Load imgmole(a) ' create img control
imgmole(a).Picture = imghole(g).Picture
'it shows the "hole" image
If a < 5 Then 'top row
imgmole(a).Move 510 + 900 * a, 1200
Else 'secound row
imgmole(a).Move 510 + 900 * (a - 5), 2100
End If
imgmole(a).Visible = True
Next

On Error GoTo 0

End Sub
Sub stopping()
Dim h
For h = 0 To 9
moletimer(h).Enabled = False
imgmole(h).Visible = False
Next
paused = True
End Sub
Private Sub pause123_Click()
pause123.Visible = False
pause1234.Visible = False

Dim h
For h = 0 To 9
moletimer(h).Enabled = True
imgmole(h).Visible = True
imgmole(h).Refresh
Next
paused = False
End Sub

Sub miss()
If gameover.Visible = False Then
misses = misses - 1
basetime = basetime + 150
lblmiss.Caption = misses
End If
If misses <= 0 Then
playing = False
gameover.Visible = True
cmdgo.Enabled = True
End If
End Sub



Sub imgmole_click(Index As Integer)
On Error GoTo passme
If molestate(Index) = 0 Or molestate(Index) = 4 Then
Exit Sub
End If
If playing = False Then
Exit Sub
'must be a hit
End If
moletimer(Index).Enabled = False
molestate(Index) = 0
imgmole(Index).Picture = imghole(g).Picture
moletimer(Index).Interval = 500 + basetime * 4 * Rnd
hits = hits + 1
lblhit.Caption = hits
basetime = basetime * 0.988
moletimer(Index).Enabled = True
passme:
End Sub

Private Sub mnuabout_Click()
frmAbout.Show 1
End Sub

Private Sub mnucharacters_Click()
frmCharacters.Show
End Sub

Private Sub mnuexit_Click()
 Dim Msg, Style, Title, Response, MyString
Msg = "Are you Sure you Want to Quit?"    ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Quiting..."    ' Define title.
        ' context.
        ' Display message.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
  End
Else    ' User chose No.
    End If
        
End Sub

Private Sub mnuhide_Click()
 'Setup initial Tray Icon
    T.cbSize = Len(T)
    T.hWnd = pichook.hWnd
    T.uId = 1&
    T.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    T.ucallbackMessage = WM_MOUSEMOVE
    T.hIcon = imgPlay.Picture
    T.szTip = "Recent" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, T
    
    'Hide this form
    Me.Hide
End Sub

Private Sub mnuhigh_Click()
high.Show

End Sub

Private Sub mnunew_Click()
hits = "0"
misses = "0"
gameover.Visible = False
Call startgame
End Sub

Private Sub mnupause_Click()
pause123.Visible = True
pause1234.Visible = True
Call stopping
End Sub

Private Sub mnuthisgame_Click()
about.Show 1
End Sub


Private Sub moletimer_Timer(Index As Integer)
moletimer(Index).Enabled = False

Randomize
If paused = False Then
imgmole(Index).Refresh
Select Case molestate(Index)
Case 0
If allcharacters = True Then
g = Rnd * 4
End If
'from in hole to half way up
molestate(Index) = 1
imgmole(Index).Picture = imghalf(g).Picture
moletimer(Index).Interval = 80

Case 1
'from half up
molestate(Index) = 2
imgmole(Index).Picture = imgup(g).Picture
moletimer(Index).Interval = basetime

Case 2
'from up to half-down
molestate(Index) = 3
imgmole(Index).Picture = imghalf(g).Picture
moletimer(Index).Interval = 80

Case 3
'from half way down to a miss
molestate(Index) = 4
imgmole(Index).Picture = imgmiss(g).Picture
moletimer(Index).Interval = 1500
Call miss

Case 4
' return from a miss back to the hole
molestate(Index) = 0
imgmole(Index).Picture = imghole(g).Picture
moletimer(Index).Interval = 500 + basetime * 4 * Rnd
End Select
End If
imgmole(Index).Refresh
moletimer(Index).Enabled = True


End Sub

Private Sub mnuPopUp_Click(Index As Integer)
    Select Case Index
        Case 0:
          Form1.Show
                  Case 1:
        frmAbout.Show
        Case 3:
        Dim Msg, Style, Title, Response, MyString
Msg = "Are you Sure you Want to Quit?"    ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Quiting..."    ' Define title.
        ' context.
        ' Display message.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
  End
Else    ' User chose No.
    End If
        
    End Select
End Sub



Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Static State As Boolean
 Static Popped As Boolean
 Static Msg As Long
   Me.Hide
    Msg = X / Screen.TwipsPerPixelX
    If Popped = False Then
        'As we don't want any popups during the processing,
        'turn off the MouseMove event by setting popped=true
        Popped = True
        
        'Se the general section on how to interpret Msg
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                mnuPopUp_Click 0
                
            Case WM_LBUTTONDOWN:
                'In this case we have a twostate button, On or Off.
                State = Not State
                If State Then
                    'Set picture
                             Else
                    T.hIcon = imgPlay.Picture
                    T.szTip = "On" + Chr$(0)
                End If
                Shell_NotifyIcon NIM_MODIFY, T
                
            Case WM_LBUTTONUP:
            PopupMenu mnudummy, 2, , , mnuPopUp(0)
            Case WM_RBUTTONDBLCLK:
            
            Case WM_RBUTTONDOWN:
            
            Case WM_RBUTTONUP:
                PopupMenu mnudummy, 2, , , mnuPopUp(0)
                
        End Select
        'OK to popup again
        Popped = False
    End If
End Sub

