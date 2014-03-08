VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmLeadCount 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "LeadCount.dat"
   ClientHeight    =   2400
   ClientLeft      =   3000
   ClientTop       =   3000
   ClientWidth     =   4725
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmLeadCount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Left            =   2460
      Top             =   780
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Text            =   "¯¯END!"
      Top             =   300
      Width           =   795
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   2235
      Left            =   3300
      TabIndex        =   0
      Top             =   60
      Width           =   2220
      ExtentX         =   3916
      ExtentY         =   3942
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2460
      Top             =   1260
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H80000005&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   1485
      Index           =   2
      Left            =   240
      Top             =   300
      Width           =   1275
   End
End
Attribute VB_Name = "frmLeadCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Public Interval As Long

Dim mbDontNavigateNow As Boolean

Public sLeft As Long
Public sDown As Long
Public sWidth As Long
Public sHeight As Long

Public sSideBorderSpacing As Long
Public sBottomborderSpacing As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''transparent vars''''''''''''''''''''''''
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   x As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private ResultRegion As Long

Private Const RectXRound As Integer = 28
Private Const RectYRound As Integer = 28
''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''fade'''''''''''''''''''''''''''''

'A few API Declarations
Private Declare Function SetLayeredWindowAttributes Lib _
  "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
  ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean

Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Sub lwa_FadeIn(ByVal hwnd As Long, Optional ByVal iStep As Integer = 1)
  
  Dim bAlpha As Integer
  
  bAlpha = 0
  While bAlpha < frmToolbar.FADE_GRAD
    If bAlpha > frmToolbar.FADE_GRAD Then bAlpha = frmToolbar.FADE_GRAD
    SetLayeredWindowAttributes hwnd, 0, bAlpha, _
      LWA_ALPHA
    DoEvents
    
    bAlpha = bAlpha + iStep
  Wend
End Sub

Public Sub lwa_FadeOut(ByVal hwnd As Long, Optional ByVal iStep As Integer = 1)
 
  Dim bAlpha As Integer
  
  bAlpha = frmToolbar.FADE_GRAD
  While bAlpha > 0
    If bAlpha < 0 Then bAlpha = 0
    SetLayeredWindowAttributes hwnd, 0, bAlpha, LWA_ALPHA
    DoEvents
    
    bAlpha = bAlpha - iStep
  Wend
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Visible = False
    ElseIf Me.WindowState <> vbMaximized Then
        Dim intWidth As Long
        Dim intHeight As Long
        prcDetermineResolution intWidth, intHeight
        
        Me.Top = (intHeight * 15) - (Me.Height + 800) '(sDown + sBottomborderSpacing)
        Me.Left = (intWidth * 15) - (Me.Width + 10) '(sLeft + sSideBorderSpacing)
        brwWebBrowser.Width = Me.Width - 200
        brwWebBrowser.Height = Me.Height - 200
    Else
        brwWebBrowser.Width = Me.Width - 200
        brwWebBrowser.Height = Me.Height - 200
    End If

End Sub

Private Sub Form_Load()

On Error Resume Next
    
    sDown = 800
    sLeft = 800
    
    sBottomborderSpacing = 10
    sSideBorderSpacing = 10
    
    Me.Width = 2445
    Me.Height = 2715
    
    prcGrabMonitorReg
    
    If frmToolbar.FADE_GRAD = Empty Then
        frmToolbar.FADE_GRAD = 0
    End If
    
    cboAddress.Left = 2640
    cboAddress.Visible = False
        
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(1, 1, 1, 0), True)
    
    '''''''''''''''fade''''''''''''''''''''''''''''''
    
    prcFade
    '''''''''''''''''''''''''''''''''''''''''''''''''
  
    prcBrowseLeadcount
    'Pause 0.1
    'Text1.Text = Me.brwWebBrowser.Document.documentElement.innerHTML
    'Pause 0.1
    'Text1.Refresh
End Sub

Sub prcFade()
    If frmToolbar.FADE_GRAD = 0 Then
        frmToolbar.FADE_GRAD = 200
    End If
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, _
      GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
      
    Me.Show
    DoEvents
    
    lwa_FadeIn Me.hwnd, 1
End Sub

Sub prcGrabMonitorReg()
    frmToolbar.FADE_GRAD = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "transgrade")
    If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "width") <> Empty Then Me.Width = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "width")
    If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "height") <> Empty Then Me.Height = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "height")
    
    'MsgBox Me.Height
End Sub



Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim Corraction As Integer
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim i As Integer
    
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    
    For i = shpBorder.LBound To shpBorder.UBound
        Select Case shpBorder(i).Shape
            Case 0: 'rectangle & square
                ObjectRegion = CreateRectRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
            Case 1: 'circle
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case 4:  'round square
                ObjectRegion = CreateRoundRectRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                        RectXRound, RectYRound)
            Case 5: 'round square
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                End If
            Case 3: 'circle
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case Else:  'oval
                shpBorder(i).Shape = 2
                ObjectRegion = CreateEllipticRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
        End Select
        nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
        nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
        DeleteObject ObjectRegion
    Next i
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function

Sub prcBrowseLeadcount()

    Timer1.Enabled = False
    StartingAddress = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "url")
    Interval = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "interval")
    brwWebBrowser.ToolTipText = Interval & ":url: " & StartingAddress & ", Interval: " & Interval & " Mins"
    cboAddress.Text = StartingAddress
    cboAddress.AddItem cboAddress.Text
    'try to navigate to the starting address
    Timer1.Interval = 60000 * Interval
    Timer1.Enabled = True
    timTimer.Enabled = True
    'brwWebBrowser.Navigate StartingAddress
    
    prcNav
    
End Sub


Private Sub brwWebBrowser_DownloadComplete()
    Dim sSplit1() As String
    Dim sSplit2() As String
    
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
    
    Text1.Text = Me.brwWebBrowser.Document.documentElement.innerHTML
    sSplit1 = Split(Trim(Text1.Text), "<PRE>")
    sSplit2 = Split(sSplit1(1), "</PRE>")
    Text1.Text = sSplit2(0)
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
    
End Sub

Sub prcNav()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_Click()
    prcNav
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    lwa_FadeOut Me.hwnd, 1
  Unload Me
End Sub

Private Sub Text1_Click()
    Me.WindowState = 0
    Me.Visible = False
End Sub


Private Sub Timer1_Timer()
    prcBrowseLeadcount
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub


