VERSION 5.00
Begin VB.Form trayicon 
   Caption         =   "VIRIIOs Tray Icon"
   ClientHeight    =   3045
   ClientLeft      =   2505
   ClientTop       =   2190
   ClientWidth     =   4080
   Icon            =   "TrayIcon.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3045
   ScaleWidth      =   4080
   Begin VB.Menu mnuDummy 
      Caption         =   "PopUp"
      Begin VB.Menu mnuPopUp 
         Caption         =   "Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuPopUp 
         Caption         =   "&About"
         Index           =   1
         Shortcut        =   ^A
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
Attribute VB_Name = "trayicon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()
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


Private Sub Form_Unload(Cancel As Integer)
    'Unload this form. Important: always end with "unload me".
    T.cbSize = Len(T)
    T.hWnd = pichook.hWnd
    T.uId = 1&
    Shell_NotifyIcon NIM_DELETE, T
    End
End Sub
Private Sub mnuPopUp_Click(Index As Integer)
    Select Case Index
        Case 0:
          Form1.Show
                  Case 2:
            Unload Me
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


