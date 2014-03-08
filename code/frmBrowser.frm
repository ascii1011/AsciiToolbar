VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   ClientHeight    =   6045
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   1740
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   960
      TabIndex        =   6
      ToolTipText     =   "Double click me to browse"
      Top             =   1260
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   735
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   5640
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":02E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0E6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3735
      Left            =   60
      TabIndex        =   0
      Top             =   2220
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   6588
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
      Location        =   ""
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5640
      Top             =   2580
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   7935
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Text            =   "¯¯END!"
         Top             =   300
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6540
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   540
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub Command1_Click()
    frmAddLink.Show
End Sub

Private Sub Command2_Click()
    Dim Response
    Dim sitems
    Dim sitem As Long
    Dim sLnkPath As String
    Dim vResult
    
    If List1.ListCount > 0 Then
        Response = MsgBox("Are you sure you want to delete this?: " & List1.List(List1.ListIndex), vbExclamation + vbYesNo, "Delete Confirmation")
        If Response <> vbYes Then
            Exit Sub
        End If
    
        sLnkPath = "Software\ModCon\Toolbar\web\links\"
        sitems = Split(List1.List(List1.ListIndex), ".)")
        vResult = DeleteKey(HKEY_LOCAL_MACHINE, sLnkPath & sitems(0))
        If vResult = 0 Then
            sitem = sitems(0)
            prcReposition sitem
            prcFillLinks
        End If
    End If
    Command2.Enabled = False
End Sub

Sub prcReposition(lStart As Long)
        Dim sLnkPath As String
    Dim lIndex As Long
                
    sLnkPath = "Software\ModCon\Toolbar\web\links\"
    
    lIndex = QueryValue(HKEY_LOCAL_MACHINE, sLnkPath, "index")
    
    If (lIndex - 1) = 0 Then
        SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath, "index", 0, REG_SZ
    Else
        Dim e As Long
        Dim indx As Long
        Dim URL As String
        Dim cmd As String
        
        For e = (lStart + 1) To lIndex
        
            indx = QueryValue(HKEY_LOCAL_MACHINE, sLnkPath & e, "index")
            Text1.Text = indx
            indx = Trim(Text1.Text)
            URL = QueryValue(HKEY_LOCAL_MACHINE, sLnkPath & e, "url")
            Text1.Text = URL
            URL = Trim(Text1.Text)
            cmd = QueryValue(HKEY_LOCAL_MACHINE, sLnkPath & e, "command")
            Text1.Text = cmd
            cmd = Trim(Text1.Text)
                        
            CreateNewKey HKEY_LOCAL_MACHINE, sLnkPath & e - 1
            SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath & e - 1, "index", e - 1, REG_SZ
            SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath & e - 1, "url", URL, REG_SZ
            SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath & e - 1, "command", cmd, REG_SZ
            
            DeleteKey HKEY_LOCAL_MACHINE, sLnkPath & e
            
        Next
        
        SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath, "index", (lIndex - 1), REG_SZ
        prcFillLinks
    End If
    
        
        
End Sub

Sub prcFillLinks()
    Dim lResult As Long
    Dim indx As String, URL As String, cmd As String
    
    List1.Clear
    
    lResult = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "index")
    If lResult > 0 Then
        Dim Y As Long
        For Y = 1 To lResult
            indx = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links\" & Y, "index")
            Text1.Text = indx
            indx = Trim(Text1.Text)
            URL = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links\" & Y, "url")
            Text1.Text = URL
            URL = Trim(Text1.Text)
            cmd = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links\" & Y, "command")
            Text1.Text = cmd
            cmd = Trim(Text1.Text)
            
            List1.AddItem indx & ".)" & URL & cmd
        Next
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
    Form_Resize
    Command2.Enabled = False
    prcFillLinks

    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    StartingAddress = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "Home")
        If StartingAddress = "" Then
            StartingAddress = "www.modernconsumer.com"
        End If
    
    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        
        brwWebBrowser.Navigate StartingAddress
        
    End If

End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
    If Trim(cboAddress.Text) <> "www.modernconsumer.com" Then
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\history", "url", Trim(cboAddress.Text), REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\archive", "url", Now & ":" & Trim(cboAddress.Text), REG_SZ
    End If
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

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    cboAddress.Width = Me.ScaleWidth - 100
    List1.Width = Me.ScaleWidth - 950
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 1100
End Sub

Private Sub List1_Click()
    Dim sitems
    
    If List1.ListCount > 0 Then
        Command2.Enabled = True
    End If
End Sub

Private Sub List1_DblClick()
    Dim sitems
    
    If List1.ListCount > 0 Then
        sitems = Split(List1.List(List1.ListIndex), ".)")
        cboAddress.Text = sitems(1)
        cboAddress_Click
    End If
    'brwWebBrowser.Navigate
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select

End Sub

