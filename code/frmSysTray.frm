VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "Form1"
   ClientHeight    =   465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu m 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu monitor 
         Caption         =   "Monitor"
      End
      Begin VB.Menu refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu Ex 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''icon_tray''''''''''''''
Const MAX_TOOLTIP As Integer = 64
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206

Private Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uId              As Long
    uFlags           As Long
    uCallBackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type
Private nfIconData As NOTIFYICONDATA

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
    Me.Left = 5
    Me.Top = -1000
    prcCreateSysTrayIcon
    Me.WindowState = 1
    frmLeadCount.Show
End Sub


Sub prcRemoveSysTrayIcon()

'
' Remove this application from the System Tray.
'
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub



Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then
        Me.Visible = False
    End If
End Sub

Private Sub monitor_Click()
    If frmLeadCount.Visible = True Then
        frmLeadCount.Visible = False
    Else
        frmLeadCount.Visible = True
    End If
End Sub


Sub prcCreateSysTrayIcon()
    '
    ' Add this application's icon to the system tray.
    '
    ' Parm 1 = Handle of the window to receive notification messages
    '          associated with an icon in the taskbar status area.
    ' Parm 2 = Icon to display.
    ' Parm 3 = Handle of icon to display.
    ' Parm 4 = Tooltip displayed when cursor moves over system tray icon.
    '
    With nfIconData
        .hwnd = Me.hwnd
        .uId = Me.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.handle
        '.szTip = "1"
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
    'Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    prcRemoveSysTrayIcon
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lMsg As Single
'
' Determine the event that happened to the System Tray icon.
' Left clicking the icon displays a message box.
' Right clicking the icon creates an instance of an object from an
' ActiveX Code component then invokes a method to display a message.
'
lMsg = x / Screen.TwipsPerPixelX
Select Case lMsg
    Case WM_LBUTTONUP
            monitor_Click 'MsgBox "Modern Consumer" & vbNewLine & "Toolbar", vbInformation, "MC_Toolbar"
    Case WM_RBUTTONUP
            Me.PopupMenu m
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONDBLCLK
    Case Else
End Select
    
End Sub

Private Sub refresh_click()
    frmLeadCount.prcBrowseLeadcount

End Sub


Private Sub Ex_Click()
    Unload frmLeadCount
    Unload Me
End Sub
