VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmToolbar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8925
   Icon            =   "frmToolbar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   3660
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1980
      Visible         =   0   'False
      Width           =   315
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7500
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   4440
      TabIndex        =   1
      Top             =   1320
      Width           =   1995
      Begin VB.PictureBox picSmall 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1200
         ScaleHeight     =   240
         ScaleMode       =   0  'User
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   180
         Width           =   240
      End
      Begin VB.PictureBox picLarge 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   240
         ScaleHeight     =   555
         ScaleMode       =   0  'User
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   180
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolbar.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   5
      Top             =   630
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label lblfile 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   8835
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' This example demonstrates how to:
'   Extract both large and small icons from executables and dll's.
'   Draw them to a control with a device context handle (.hdc) such as a PictureBox.
'   Draw them to a control without an .hdc property such as an ImageList.
'   Dynamically populate both ImageList and ToolBar controls.
'

'
'file contains
'iSize:0 or 1 for small or large size
'lIndex:long number for index of last image inserted, minimum = 2
'images:image index#:path:name:command line:key
'
Dim glLargeIcons() As Long
Dim glSmallIcons() As Long
Dim sImageParams() As String
Dim lIndex         As Long
Dim lIcons         As Long
Dim sExeName       As String
Public iSize       As Integer
Dim lImageIndex    As Long
Public iImageMaxNum    As Integer
Public iTitleBar As Integer
Public iPosition As Integer
Public iOS As Integer
Public sSysPath As String

'''''''''''for fading'''''''''''
Public FADE_GRAD As Long
'''''''''''scroll on settings''''''
Public iScroll1 As Integer



Private Sub Form_Load()
    sSettingsForm = 0
    frmToolbar.BorderStyle = 0
    iScroll1 = 0
    prcInitVars
    
    prcFillToolBar
    'frmSysTray.Show
End Sub





Function prcSetReg() As Integer
    Dim vResult
    Dim sResult1
    Dim sResult2
    
    sResult1 = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "version")
    Text1.Text = sResult1
    sResult1 = Trim(Text1.Text)
    sResult2 = App.Major & "." & App.Minor & "." & App.Revision
            
    If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar") = False Then
        prcDoRegSet
    ElseIf CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main") = False Then
        'if toolbar does not exsist
        'vResult = DeleteKey(HKEY_LOCAL_MACHINE, "Software\ModCon")
        prcDoRegSet
    ElseIf sResult1 <> sResult2 Then
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "version", App.Major & "." & App.Minor & "." & App.Revision, REG_SZ     'size of buttons
        'reinit vars if version wrong
        'vResult = DeleteKey(HKEY_LOCAL_MACHINE, "Software\ModCon")
        prcDoRegSet
    End If
    
End Function

Sub prcDoRegSet()
    Dim vResult
    Dim sResult1
    Dim sResult2
    
    sResult1 = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "version")
    Text1.Text = sResult1
    sResult1 = Trim(Text1.Text)
    sResult2 = App.Major & "." & App.Minor & "." & App.Revision

        '''''''''''''''sys info'''''''''''''''
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main"
        If sResult1 <> sResult2 Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "version", App.Major & "." & App.Minor & "." & App.Revision, REG_SZ     'size of buttons
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "iSize") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "iSize", 0, REG_SZ     'size of buttons
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "lIndex") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "lIndex", DEFAULT_INDEX, REG_SZ   'how many images
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "TitleBar") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "TitleBar", 0, REG_SZ   'show title bar = 1
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "pos") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "pos", 2, REG_SZ     'pos of toolbar
        
        ''''''''''''''buttons'''''''''''''''
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0"
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1"
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2"
        
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "index") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "index", 0, REG_SZ    'index number of button
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "path") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "path", "c:\" & sSysPath & "\system32\ctfmon.exe", REG_SZ     'path to where the file is
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "name") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "name", "Settings", REG_SZ    'personalized name for this button
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "key") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\0", "key", "settings", REG_SZ     'app button key name
                
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "index") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "index", 1, REG_SZ    'index number of button
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "path") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "path", "c:\" & sSysPath & "\system32\dxdiag.exe", REG_SZ     'path to where the file is
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "name") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "name", "Exit", REG_SZ    'personalized name for this button
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "key") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\1", "key", "exit", REG_SZ
        
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "index") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "index", 2, REG_SZ    'index number of button
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "path") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "path", "C:\Program Files\Internet Explorer\IEXPLORE.EXE", REG_SZ     'path to where the file is
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "name") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "name", "Modcon Browser", REG_SZ    'personalized name for this button
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "key") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\2", "key", "browser", REG_SZ 'app button key name
        '''''''''''''''''''''''''''''''''''
        
        ''''''''''''''''''browser'''''''''''''''''''''''''''''''''''
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links"
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\history") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\history"
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\archive") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\archive"
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "Home") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "Home", "www.modernconsumer.com", REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "index") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "index", "0", REG_SZ
        
        ''''''''''''''''''''''monitor''''''''''''''''''''''''''''''''
        If CheckRegistryKey(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor") = False Then CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor"
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "url") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "url", "www.driverloans.com/leadcount.dat", REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "interval") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "interval", 1, REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "transgrade") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "transgrade", 200, REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "width") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "width", "2445", REG_SZ
        If QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "height") = Empty Then SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "height", "2715", REG_SZ
        
End Sub

Sub prcGrabReg()
    Dim lWidth As Long
    
    iSize = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "iSize")
    iImageMaxNum = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "lIndex")
    iTitleBar = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "TitleBar")
    iPosition = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "pos")
        
    If iSize = 0 Then
        Me.Height = 400
        lWidth = 350
        Toolbar1.Visible = True
        Toolbar2.Visible = False
    ElseIf iSize = 1 Then
        Me.Height = 670
        lWidth = 590
        Toolbar1.Visible = False
        Toolbar2.Visible = True
    End If
    
    If iTitleBar = 0 Then
        frmToolbar.BorderStyle = 0
    Else
        frmToolbar.BorderStyle = iTitleBar
    End If
    
    
    Me.Width = 1
    If iImageMaxNum >= 1 Then
        Dim icount As Long
        ReDim sImageParams(iImageMaxNum, 3)
        For icount = 0 To iImageMaxNum - 1
            sImageParams(icount, 0) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "index"))
            If sImageParams(icount, 0) <> "" Then
                Me.Width = Me.Width + lWidth
                sImageParams(icount, 1) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "path"))
                sImageParams(icount, 2) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "name"))
                sImageParams(icount, 3) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "key"))
                
                sImageParams(icount, 0) = Left(sImageParams(icount, 0), Len(sImageParams(icount, 0)) - 1)
                sImageParams(icount, 1) = Left(sImageParams(icount, 1), Len(sImageParams(icount, 1)) - 1)
                sImageParams(icount, 2) = Left(sImageParams(icount, 2), Len(sImageParams(icount, 2)) - 1)
                sImageParams(icount, 3) = Left(sImageParams(icount, 3), Len(sImageParams(icount, 3)) - 1)
            End If
        Next
    End If
    
    frmSettings.prcSetPos
End Sub
    

Sub prcInitVars()
    
    Me.Top = 5
    Me.Left = 1000
    sSysPath = ""
    
    prcGetSysInfo
    
    If sSysPath <> "" Then
        prcSetReg
        prcGrabReg
        '
        ' Align the toolbars to the top of the form.
        '
        With Toolbar1
            .Align = vbAlignTop
            .AllowCustomize = False
            .Wrappable = False
            .BorderStyle = ccNone
        End With
        
        picLarge.Height = LARGE_ICON * Screen.TwipsPerPixelY
        picLarge.Width = LARGE_ICON * Screen.TwipsPerPixelX
        picSmall.Height = SMALL_ICON * Screen.TwipsPerPixelY
        picSmall.Width = SMALL_ICON * Screen.TwipsPerPixelX
    Else
        MsgBox "This operating system is not supported.  Good bye."
        Unload Me
    End If
    
    
End Sub

Sub prcFillToolBar()

    prcClearButtons
    
    'iSize = UBound(sImageParams)
    For lImageIndex = 0 To iImageMaxNum - 1
        If Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & lImageIndex, "index")) <> "" Then
            prcDisplayButtons
        End If
    Next
End Sub

Sub prcClearButtons()
    Toolbar1.buttons.Clear
    Toolbar1.ImageList = Nothing
    Toolbar2.buttons.Clear
    Toolbar2.ImageList = Nothing
    
    With ImageList2
        .ListImages.Clear
        .ImageHeight = LARGE_ICON
        .ImageWidth = LARGE_ICON
    End With
    
    With ImageList1
        .ListImages.Clear
        .ImageHeight = SMALL_ICON
        .ImageWidth = SMALL_ICON
    End With
End Sub
    




Sub prcDisplayButtons()
    Dim btn    As Button
    Dim imgObj As ListImage
    Dim ind

    lIndex = 0
    lblfile = ""
    picSmall.Picture = LoadPicture("")
    picLarge.Picture = LoadPicture("")

    picSmall.Enabled = True
    picLarge.Enabled = True
    
    ReDim glLargeIcons(0)
    ReDim glSmallIcons(0)
    Call pGetIcon
    
    ind = sImageParams(lImageIndex, 0)
    
    Set imgObj = ImageList2.ListImages.Add(lImageIndex + 1, sImageParams(lImageIndex, 1), picLarge.Image)
    With Toolbar2
        .ImageList = ImageList2
        Set btn = .buttons.Add(.buttons.Count + 1, sImageParams(lImageIndex, 1), , , sImageParams(lImageIndex, 1))
        .buttons(lImageIndex + 1).ToolTipText = sImageParams(lImageIndex, 2)
        .buttons(lImageIndex + 1).Key = sImageParams(lImageIndex, 3)
    End With
    
    Set imgObj = ImageList1.ListImages.Add(lImageIndex + 1, sImageParams(lImageIndex, 1), picSmall.Image)
    With Toolbar1
        .ImageList = ImageList1
        Set btn = .buttons.Add(.buttons.Count + 1, sImageParams(lImageIndex, 1), , , sImageParams(lImageIndex, 1))
        .buttons(lImageIndex + 1).ToolTipText = sImageParams(lImageIndex, 2)
        .buttons(lImageIndex + 1).Key = sImageParams(lImageIndex, 3)
    End With

End Sub




Public Sub WriteDefault()
    Dim f As Integer
    Dim sParams As String

On Error GoTo errHandle:

    sParams = "iSize:" & iSize
    f = FreeFile()
    
    Open "c:\Program Files\Toolbar_settings.asc" For Append As #f
    Print #f, sParams
    Close #f
    
errHandle:
        'Print #f, Now & " Err = " & Err.Description & " File Error!"
        Close #f
        Screen.MousePointer = vbDefault
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Me.Top = Y
    'Me.Left = X
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    'Me.Top = Y
    'Me.Left = X
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmSettings
    Unload frmBrowser
    Unload frmLeadCount
    Unload frmSysTray
    Unload frmInfo
    Unload frmAddLink
    Unload frmResolutionCntrl
    
End Sub


Private Sub toolBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Response

On Error GoTo errHandle

    Select Case (Button.Key)
        Case (sImageParams(Button.index - 1, 3)):
            If Button.index = 1 Then
                prcSettings
            ElseIf Button.index = 2 Then
                Unload Me
            ElseIf Button.index = 3 Then
                frmBrowser.Show
            Else
                Call Shell(sImageParams(Button.index - 1, 1), vbNormalFocus)
            End If
                            
    End Select
    
    Exit Sub

errHandle:
    Select Case (Err.Number)
        Case (91):
            Resume Next
        Case (438):
            MsgBox Me.Caption & " does not support this operation", vbInformation, "SYSTEM"
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Login run time error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub



Private Sub toolBar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Response

On Error GoTo errHandle

    Select Case (Button.Key)
        Case (sImageParams(Button.index - 1, 3)):
                                                    If Button.index = 1 Then
                                                        prcSettings
                                                    ElseIf Button.index = 2 Then
                                                        Unload Me
                                                    ElseIf Button.index = 3 Then
                                                        frmBrowser.Show
                                                    Else
                                                        Call Shell(sImageParams(Button.index - 1, 1), vbNormalFocus)
                                                    End If
                            
    End Select
    
    Exit Sub

errHandle:
    Select Case (Err.Number)
        Case (91):
            Resume Next
        Case (438):
            MsgBox Me.Caption & " does not support this operation", vbInformation, "SYSTEM"
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Login run time error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub



Sub prcSettings()
    If sSettingsForm = 1 Then
        Unload frmSettings
    End If
    frmSettings.Show
End Sub






Public Sub pGetIcon()
    Dim l As Long
    'MsgBox sImageParams(lImageIndex, 1)
    Call ExtractIconEx(sImageParams(lImageIndex, 1), lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)
    
    With picLarge
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        Call DrawIconEx(.hdc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
        .refresh
    End With
    
    With picSmall
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        Call DrawIconEx(.hdc, 0, 0, glSmallIcons(lIndex), SMALL_ICON, SMALL_ICON, 0, 0, DI_NORMAL)
        .refresh
    End With

End Sub

Sub prcGetSysInfo()
    Dim msg As String         ' Status information.
    Dim NewLine As String     ' New-line.
    Dim ret As Integer        ' OS Information
    Dim ver_major As Integer  ' OS Version
    Dim ver_minor As Integer  ' Minor Os Version
    Dim Build As Long         ' OS Build

    iOS = 0
    
    ' Get operating system and version.
    Dim verinfo As OSVERSIONINFO
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    ret = GetVersionEx(verinfo)
    If ret = 0 Then
        MsgBox "Error Getting Version Information"
        End
    End If
        
    
    ver_major = verinfo.dwMajorVersion
    ver_minor = verinfo.dwMinorVersion
    Build = verinfo.dwBuildNumber
    
    'msg = msg & Build
    
    'Get_User_Name
    'Label2 = msg  '''''', vbOKOnly, "System Info"
    
    If verinfo.dwPlatformId = 2 And verinfo.dwMajorVersion = 5 And verinfo.dwMinorVersion = 1 And verinfo.dwBuildNumber = 2600 Then
        'Label2 = "XP-" & msg & ": " & strUserTsr
        'Label3 = strCompName
        iOS = 1
        sSysPath = "windows"
    Else
        If verinfo.dwPlatformId = 2 And verinfo.dwMajorVersion = 5 And verinfo.dwMinorVersion = 0 And verinfo.dwBuildNumber = 2195 Then
            'Label2 = "2K-" & msg & "-" & strUserTsr
            'Label3 = strCompName
            iOS = 2
            sSysPath = "winnt"
        Else
            'Command3.Enabled = False
            'Label2 = "Unknown-" & msg & "-" & strUserTsr
            'Label3 = strCompName
            sSysPath = ""
            iOS = 10
        End If
    End If
End Sub

