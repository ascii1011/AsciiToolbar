VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6030
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   3180
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   3180
      Width           =   795
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3075
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Under Construction"
      Top             =   0
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5424
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Settings"
      TabPicture(0)   =   "frmSettings.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmResolutionCntrl"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Add Link"
      TabPicture(1)   =   "frmSettings.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Remove Link"
      TabPicture(2)   =   "frmSettings.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Monitor"
      TabPicture(3)   =   "frmSettings.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      Begin VB.Frame frmResolutionCntrl 
         Height          =   2535
         Left            =   120
         TabIndex        =   36
         Top             =   420
         Width           =   5535
         Begin VB.CommandButton Command14 
            Caption         =   "R"
            Height          =   315
            Left            =   5040
            TabIndex        =   61
            Top             =   1980
            Width           =   315
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Close Res"
            Height          =   315
            Left            =   3960
            TabIndex        =   60
            Top             =   1980
            Width           =   915
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Resolution"
            Height          =   315
            Left            =   3960
            TabIndex        =   59
            Top             =   1440
            Width           =   915
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Reset All"
            Height          =   315
            Left            =   240
            TabIndex        =   53
            Top             =   1320
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Small"
            Height          =   195
            Left            =   360
            TabIndex        =   46
            Top             =   540
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Large"
            Height          =   195
            Left            =   360
            TabIndex        =   45
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Show"
            Height          =   195
            Left            =   1440
            TabIndex        =   44
            Top             =   2280
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   2820
            Picture         =   "frmSettings.frx":04B2
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   43
            Top             =   780
            Width           =   495
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   3480
            Picture         =   "frmSettings.frx":08F4
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   42
            Top             =   780
            Width           =   495
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   2160
            Picture         =   "frmSettings.frx":0D36
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   41
            Top             =   780
            Width           =   495
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   2400
            Picture         =   "frmSettings.frx":1178
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   40
            Top             =   1560
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.PictureBox Picture5 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   3240
            Picture         =   "frmSettings.frx":15BA
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   39
            Top             =   1560
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Display Properties"
            Height          =   495
            Left            =   240
            TabIndex        =   38
            Top             =   1920
            Width           =   915
         End
         Begin VB.PictureBox Picture6 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   4620
            Picture         =   "frmSettings.frx":19FC
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   37
            ToolTipText     =   "If only one monitor is installed, this will put the toolbar out of view."
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Button Size:"
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Show Title Bar:"
            Height          =   195
            Left            =   1260
            TabIndex        =   49
            Top             =   1980
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Move Toolbar"
            Height          =   195
            Left            =   2520
            TabIndex        =   48
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label15 
            Caption         =   "Extended windows only!"
            Height          =   435
            Left            =   4320
            TabIndex        =   47
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   25
         Top             =   420
         Width           =   5535
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   720
            TabIndex        =   31
            ToolTipText     =   "Type a name for this link"
            Top             =   540
            Width           =   3555
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Browse"
            Height          =   315
            Left            =   1260
            TabIndex        =   30
            Top             =   1680
            Width           =   795
         End
         Begin VB.Frame Frame2 
            Height          =   1335
            Left            =   4380
            TabIndex        =   27
            Top             =   240
            Width           =   975
            Begin VB.PictureBox picSmall 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   300
               ScaleHeight     =   240
               ScaleMode       =   0  'User
               ScaleWidth      =   240
               TabIndex        =   29
               Top             =   1020
               Width           =   240
            End
            Begin VB.PictureBox picLarge 
               BorderStyle     =   0  'None
               Height          =   555
               Left            =   180
               ScaleHeight     =   555
               ScaleMode       =   0  'User
               ScaleWidth      =   495
               TabIndex        =   28
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Add Link"
            Height          =   315
            Left            =   2940
            TabIndex        =   26
            Top             =   1680
            Width           =   795
         End
         Begin MSComDlg.CommonDialog cdlOpen 
            Left            =   4680
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label2 
            Caption         =   "Path:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1140
            Width           =   435
         End
         Begin VB.Label lblfile 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   720
            TabIndex        =   34
            Top             =   1080
            Width           =   3555
         End
         Begin VB.Label Label4 
            Caption         =   "Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000C&
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2220
            Width           =   5295
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   12
         Top             =   420
         Width           =   5535
         Begin VB.ListBox List1 
            Height          =   1815
            Left            =   180
            TabIndex        =   19
            Top             =   240
            Width           =   1275
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2460
            TabIndex        =   18
            ToolTipText     =   "Type a name for this link"
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   2460
            TabIndex        =   17
            ToolTipText     =   "Type a name for this link"
            Top             =   660
            Width           =   2775
         End
         Begin VB.TextBox Text4 
            Height          =   315
            Left            =   2460
            TabIndex        =   16
            ToolTipText     =   "Type a name for this link"
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Left            =   2460
            TabIndex        =   15
            ToolTipText     =   "Type a name for this link"
            Top             =   1500
            Width           =   2775
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Update"
            Height          =   315
            Left            =   2460
            TabIndex        =   14
            Top             =   1920
            Width           =   795
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Remove"
            Height          =   315
            Left            =   4080
            TabIndex        =   13
            Top             =   1920
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "Index:"
            Height          =   195
            Left            =   1860
            TabIndex        =   24
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Path:"
            Height          =   195
            Left            =   1860
            TabIndex        =   23
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Name:"
            Height          =   195
            Left            =   1860
            TabIndex        =   22
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Key:"
            Height          =   195
            Left            =   1860
            TabIndex        =   21
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label9 
            BackColor       =   &H8000000C&
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   2280
            Width           =   5295
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   3
         Top             =   420
         Width           =   5535
         Begin VB.TextBox Text10 
            Height          =   255
            Left            =   4320
            TabIndex        =   58
            Top             =   1080
            Width           =   435
         End
         Begin VB.TextBox Text9 
            Height          =   315
            Left            =   2580
            TabIndex        =   56
            ToolTipText     =   "Press ""Enter"" to preview results"
            Top             =   1440
            Width           =   795
         End
         Begin VB.TextBox Text8 
            Height          =   315
            Left            =   780
            TabIndex        =   54
            ToolTipText     =   "Press ""Enter"" to preview results"
            Top             =   1440
            Width           =   795
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   1860
            Max             =   255
            Min             =   1
            TabIndex        =   51
            Top             =   1080
            Value           =   1
            Width           =   2415
         End
         Begin VB.TextBox Text7 
            Height          =   315
            Left            =   660
            TabIndex        =   8
            Top             =   660
            Width           =   795
         End
         Begin VB.TextBox Text6 
            Height          =   315
            Left            =   660
            TabIndex        =   7
            Top             =   240
            Width           =   3555
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Update"
            Height          =   435
            Left            =   4140
            TabIndex        =   6
            Top             =   1980
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Show Monitor"
            Height          =   435
            Left            =   240
            TabIndex        =   5
            Top             =   1980
            Width           =   975
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Refresh Monitor"
            Height          =   435
            Left            =   2220
            TabIndex        =   4
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Height:"
            Height          =   195
            Left            =   2040
            TabIndex        =   57
            Top             =   1500
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Width:"
            Height          =   195
            Left            =   240
            TabIndex        =   55
            Top             =   1500
            Width           =   495
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   5340
            Y1              =   1860
            Y2              =   1860
         End
         Begin VB.Label Label16 
            Caption         =   "Transparent Gradient"
            Height          =   195
            Left            =   180
            TabIndex        =   52
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "URL:"
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Time:"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Minutes"
            Height          =   195
            Left            =   1560
            TabIndex        =   9
            Top             =   720
            Width           =   675
         End
      End
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000A&
      BorderColor     =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   3540
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   5955
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim glLargeIcons() As Long
Dim glSmallIcons() As Long
Dim sImageParams() As String
Dim lIndex         As Long
Dim lIcons         As Long
Dim sExeName       As String
Dim sTmpImageParams() As String

'monitor
Dim sMonitorAddress As String
Dim sMonitorInterval As Long
Dim sMonitorWidth As Long
Dim sMonitorHeight As Long
Dim sMonitorTransgrade As Long

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





Private Sub Check1_Click()
    prcShowTitle
End Sub

Sub prcShowTitle()
    If Check1.Value = 0 Then
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "TitleBar", 0, REG_SZ     'size of buttons
    Else
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "TitleBar", 1, REG_SZ     'size of buttons
    End If
    frmToolbar.prcGrabReg
End Sub

Private Sub Command1_Click()
    prcBrowse
End Sub

Private Sub prcBrowse()
    lIndex = 0
    lblfile = ""
    picSmall.Picture = LoadPicture("")
    picLarge.Picture = LoadPicture("")
    
    cdlOpen.FLAGS = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    cdlOpen.FileName = ""
    cdlOpen.Filter = "Executable Files (*.exe) | *.exe|Application Extension (*.dll) | *.dll"
    
On Error GoTo CancelButton
    cdlOpen.Action = 1
    sExeName = cdlOpen.FileName
    lblfile = sExeName

    picSmall.Enabled = True
    picLarge.Enabled = True
    ReDim glLargeIcons(0)
    ReDim glSmallIcons(0)
    Call pGetIcon
CancelButton:
    'We end up here when hitting Cancel on the Open File dialog.
End Sub

Private Sub Command10_Click()
    prcDisplayProperties
End Sub

Private Sub Command12_Click()
    frmResCntrl.Show
End Sub

Private Sub Command13_Click()
    Unload frmResCntrl
End Sub

Private Sub Command14_Click()
    SetResolution 1280, 800, 32
    Pause 0.1
    ResetResolution
End Sub

Private Sub Command2_Click()
    frmInfo.Show
End Sub

Private Sub Command3_Click()
    prcAddLink
End Sub

Sub prcAddLink()
    Dim vResult
    Dim sLnkPath As String
    Dim sTmpKeys
    Dim sTmpKey As String
    
        sTmpKeys = Split(Trim(lblfile.Caption), "\")
        If UBound(sTmpKeys) > 0 Then
            sTmpKey = Trim(sTmpKeys(UBound(sTmpKeys)))
            sTmpKey = Left(sTmpKey, Len(sTmpKey) - 4)
        End If
        
        If Trim(Text1.Text) = "" Then
            Text1.Text = sTmpKey
        End If
        
        vResult = 0
        sLnkPath = "Software\ModCon\Toolbar\Buttons\"
    
        vResult = vResult + CreateNewKey(HKEY_LOCAL_MACHINE, sLnkPath & frmToolbar.iImageMaxNum)
        vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, sLnkPath & frmToolbar.iImageMaxNum, "index", frmToolbar.iImageMaxNum, REG_SZ)    'index number of button
        vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, sLnkPath & frmToolbar.iImageMaxNum, "path", Trim(lblfile.Caption), REG_SZ)       'path to where the file is
        vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, sLnkPath & frmToolbar.iImageMaxNum, "name", Trim(Text1.Text), REG_SZ)     'personalized name for this button
        
        vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, sLnkPath & frmToolbar.iImageMaxNum, "key", sTmpKey, REG_SZ)      'app button key name
        
        If vResult = 0 Then
            vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "lIndex", frmToolbar.iImageMaxNum + 1, REG_SZ)
        End If
        If vResult <> 0 Then
            Label3.Caption = "an Error has occured while adding this link"
            DeleteKey HKEY_LOCAL_MACHINE, sLnkPath & frmToolbar.iImageMaxNum
            prcClearForm
            Exit Sub
        End If
        prcClearForm
        Label3.Caption = "link has been added."
        frmToolbar.prcGrabReg
        frmToolbar.prcFillToolBar
End Sub

Sub prcClearForm()
    Text1.Text = ""
    lblfile.Caption = ""
    Command3.Enabled = False
    picLarge.Cls
    picSmall.Cls
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub



Sub prcBrowseRegLinks()
    Dim max
    
    List1.Clear
    
    sMonitorAddress = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "url")
    sMonitorInterval = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "interval")
    sMonitorWidth = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "width")
    sMonitorHeight = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "height")
    sMonitorTransgrade = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "transgrade")
    
    Text6.Text = sMonitorAddress
    Text7.Text = sMonitorInterval
    Text8.Text = sMonitorWidth
    Text9.Text = sMonitorHeight
    Text10.Text = sMonitorTransgrade
        
    frmToolbar.iScroll1 = 1
    HScroll1.Value = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "transgrade")
    
    frmToolbar.iImageMaxNum = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "lindex"))
            
    max = frmToolbar.iImageMaxNum
    If max >= 1 Then
        Dim icount As Long
        ReDim sTmpImageParams(max, 3)
        For icount = 3 To max - 1
        
            sTmpImageParams(icount, 0) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "index"))
            If sTmpImageParams(icount, 0) <> "" Then
                sTmpImageParams(icount, 1) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "path"))
                sTmpImageParams(icount, 2) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "name"))
                sTmpImageParams(icount, 3) = Trim(QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\Buttons\" & icount, "key"))
                
                sTmpImageParams(icount, 0) = Left(sTmpImageParams(icount, 0), Len(sTmpImageParams(icount, 0)) - 1)
                sTmpImageParams(icount, 1) = Left(sTmpImageParams(icount, 1), Len(sTmpImageParams(icount, 1)) - 1)
                sTmpImageParams(icount, 2) = Left(sTmpImageParams(icount, 2), Len(sTmpImageParams(icount, 2)) - 1)
                sTmpImageParams(icount, 3) = Left(sTmpImageParams(icount, 3), Len(sTmpImageParams(icount, 3)) - 1)
                
                List1.AddItem sTmpImageParams(icount, 2)
            End If
            
        Next
    End If
    
End Sub

Private Sub Command5_Click()
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "url", Trim(Text6.Text), REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "interval", Trim(Text7.Text), REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "transgrade", Trim(Text10.Text), REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "width", Trim(Text8.Text), REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\monitor", "height", Trim(Text9.Text), REG_SZ
End Sub

Private Sub Command6_Click()
    prcUpdateLink
End Sub

Sub prcUpdateLink()
    Dim vResult
    Dim sLnkPath As String
    Dim lLnkIndex As Long
        
        vResult = 0
        sLnkPath = "Software\ModCon\Toolbar\Buttons\"
        lLnkIndex = Trim(Text2.Text)
        
        vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, sLnkPath & lLnkIndex, "path", Trim(Text3.Text), REG_SZ)       'path to where the file is
        vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, sLnkPath & lLnkIndex, "name", Trim(Text4.Text), REG_SZ)     'personalized name for this button
        vResult = vResult + SetKeyValue(HKEY_LOCAL_MACHINE, sLnkPath & lLnkIndex, "key", Trim(Text5.Text), REG_SZ)      'app button key name
                
        If vResult <> 0 Then
            Label9.Caption = "an Error has occured while updating this link"
        Else
            Label9.Caption = "link has been udpated."
        End If
        
        prcBrowseRegLinks
        frmToolbar.prcGrabReg
        frmToolbar.prcFillToolBar
End Sub

Private Sub Command7_Click()
    prcRemoveLink
End Sub

Sub prcRemoveLink()
    Dim vResult
    Dim sLnkPath As String
    Dim lLnkIndex As Long
        
        vResult = 0
        sLnkPath = "Software\ModCon\Toolbar\Buttons\"
        lLnkIndex = Trim(Text2.Text)
        
        vResult = DeleteKey(HKEY_LOCAL_MACHINE, sLnkPath & lLnkIndex)
        
        If vResult <> 0 Then
            Label9.Caption = "an Error has occured while removing this link"
        Else
            frmToolbar.iImageMaxNum = frmToolbar.iImageMaxNum - 1
            SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "lIndex", frmToolbar.iImageMaxNum, REG_SZ
            Label9.Caption = "link has been deleted."
        End If
        prcClearRemoveform
        prcCleanUpLinks (lLnkIndex)
        frmToolbar.prcGrabReg
        frmToolbar.prcFillToolBar
        prcBrowseRegLinks
End Sub

Sub prcClearRemoveform()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Command6.Enabled = False
    Command7.Enabled = False
End Sub

Sub prcCleanUpLinks(linkid As Long)
    Dim skIndex
    Dim skPath
    Dim skName
    Dim skKey
    Dim sLnkPath As String
    
    sLnkPath = "Software\ModCon\Toolbar\Buttons\"
    
    If linkid <> frmToolbar.iImageMaxNum Then
        Dim G As Long
        For G = linkid To frmToolbar.iImageMaxNum - 1
            skIndex = Trim(QueryValue(HKEY_LOCAL_MACHINE, sLnkPath & G + 1, "index"))
            If skIndex <> "" Then
                skPath = Trim(QueryValue(HKEY_LOCAL_MACHINE, sLnkPath & G + 1, "path"))
                skName = Trim(QueryValue(HKEY_LOCAL_MACHINE, sLnkPath & G + 1, "name"))
                skKey = Trim(QueryValue(HKEY_LOCAL_MACHINE, sLnkPath & G + 1, "key"))
                
                skIndex = Left(skIndex, Len(skIndex) - 1)
                skPath = Left(skPath, Len(skPath) - 1)
                skName = Left(skName, Len(skName) - 1)
                skKey = Left(skKey, Len(skKey) - 1)
                
                CreateNewKey HKEY_LOCAL_MACHINE, sLnkPath & G
                SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath & G, "index", skIndex - 1, REG_SZ
                SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath & G, "path", skPath, REG_SZ
                SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath & G, "name", skName, REG_SZ
                SetKeyValue HKEY_LOCAL_MACHINE, sLnkPath & G, "key", skKey, REG_SZ
                
                DeleteKey HKEY_LOCAL_MACHINE, sLnkPath & G + 1
            End If
        Next
    End If
End Sub

Private Sub Command8_Click()
    frmSysTray.Show
End Sub

Private Sub monitor_Click()
    
    Me.WindowState = 0
    Me.Visible = True
End Sub

Private Sub Command9_Click()
    frmLeadCount.prcBrowseLeadcount
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


Private Sub Form_Load()
    Dim btn    As Button
    Dim imgObj As ListImage
    Dim ind As Integer
    Dim i
    
    sSettingsForm = 1
    
    If frmToolbar.iPosition < 4 Then
        Me.Top = 1000
        Me.Left = 1000
    Else
        Me.Top = 1000
        Me.Left = 22000
    End If
    
    Check1.Value = frmToolbar.iTitleBar
    If frmToolbar.iSize = 0 Then
        Option1.Value = True
        Option2.Value = False
    Else
        Option1.Value = False
        Option2.Value = True
    End If
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    Command6.Enabled = False
    Command7.Enabled = False
    
    prcClearForm
    
    picLarge.Height = LARGE_ICON * Screen.TwipsPerPixelY
    picLarge.Width = LARGE_ICON * Screen.TwipsPerPixelX
    picSmall.Height = SMALL_ICON * Screen.TwipsPerPixelY
    picSmall.Width = SMALL_ICON * Screen.TwipsPerPixelX
    
End Sub




Public Sub pGetIcon()
    Dim l As Long
    
    Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)
    
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



Sub prcDisplayProperties()
    Dim dblReturn As Double
    'dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3", 5)
    
    dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3", 5)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub





Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    sSettingsForm = 0
End Sub

Private Sub HScroll1_Change()
    
    If frmToolbar.iScroll1 > 1 Then
        frmToolbar.FADE_GRAD = HScroll1.Value
        Text10.Text = HScroll1.Value
        frmLeadCount.prcFade
    Else
        frmToolbar.iScroll1 = 2
    End If
End Sub

Private Sub lblfile_Change()

    If lblfile.Caption = "" Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If
End Sub

Private Sub List1_Click()
    Dim max
    
    max = frmToolbar.iImageMaxNum
    If max >= 1 Then
        Dim icount As Long
        
        For icount = 0 To max - 1
        
            If sTmpImageParams(icount, 2) = List1.List(List1.ListIndex) Then
                Text2.Text = sTmpImageParams(icount, 0)
                Text3.Text = sTmpImageParams(icount, 1)
                Text4.Text = sTmpImageParams(icount, 2)
                Text5.Text = sTmpImageParams(icount, 3)
                icount = max
            End If
            
        Next
    End If
End Sub

Private Sub Option1_Click()
    prcBtnSize
End Sub

Sub prcBtnSize()
    If Option1.Value = True Then
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "iSize", 0, REG_SZ     'size of buttons
    Else
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "iSize", 1, REG_SZ     'size of buttons
    End If
    frmToolbar.prcGrabReg
End Sub

Private Sub Option2_Click()
    prcBtnSize
End Sub

Private Sub Picture1_Click()
    prcPosLeft
End Sub

Sub prcPosLeft()
    frmToolbar.iPosition = 1
    prcSetPosReg
    frmToolbar.Left = 10
    frmToolbar.Top = 5
End Sub

Private Sub Picture2_Click()
    prcPosRight
End Sub

Sub prcPosXtrmeRight()
    Dim intWidth As Long
    Dim intHeight As Long
    
    frmToolbar.iPosition = 4
    prcSetPosReg
    prcDetermineResolution intWidth, intHeight
    frmToolbar.Top = 5
    frmToolbar.Left = ((intWidth * 15) + 10)
End Sub

Sub prcPosRight()
    Dim intWidth As Long
    Dim intHeight As Long
    
    frmToolbar.iPosition = 3
    prcSetPosReg
    prcDetermineResolution intWidth, intHeight
    frmToolbar.Top = 5
    frmToolbar.Left = ((intWidth * 15) - (frmToolbar.Width + 5))
End Sub

Private Sub Picture3_Click()
    prcPosCenter
End Sub

Sub prcPosCenter()
    Dim intWidth As Long
    Dim intHeight As Long
    
    frmToolbar.iPosition = 2
    prcSetPosReg
    prcDetermineResolution intWidth, intHeight
    frmToolbar.Top = 5
    frmToolbar.Left = (((intWidth * 15) / 2) - ((frmToolbar.Width + 5) / 2))
End Sub

Sub prcSetPos()
    If frmToolbar.iPosition = 1 Then
        prcPosLeft
    ElseIf frmToolbar.iPosition = 2 Then
        prcPosCenter
    ElseIf frmToolbar.iPosition = 3 Then
        prcPosRight
    ElseIf frmToolbar.iPosition = 4 Then
        prcPosXtrmeRight
    End If
End Sub

Sub prcSetPosReg()
    SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\main", "pos", frmToolbar.iPosition, REG_SZ     'size of buttons
End Sub

Private Sub Picture6_Click()
    prcPosXtrmeRight
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    prcBrowseRegLinks
End Sub


Private Sub Text2_Change()
    If Trim(Text2.Text) = "" Then
        Command6.Enabled = False
        Command7.Enabled = False
    Else
        Command6.Enabled = True
        Command7.Enabled = True
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmLeadCount.Width = Trim(Text8.Text)
        frmLeadCount.Height = Trim(Text9.Text)
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmLeadCount.Width = Trim(Text8.Text)
        frmLeadCount.Height = Trim(Text9.Text)
    End If
End Sub
