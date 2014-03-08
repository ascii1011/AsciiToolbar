VERSION 5.00
Begin VB.Form frmAddLink 
   Caption         =   "Add a Link"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frmAddLink.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   300
      Width           =   6615
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   315
      Left            =   6180
      TabIndex        =   1
      Top             =   1380
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   360
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Command:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   795
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   60
      Top             =   60
      Width           =   7695
   End
End
Attribute VB_Name = "frmAddLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim lIndex As Long
    
    lIndex = QueryValue(HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "index")
    
    lIndex = lIndex + 1
    SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links", "index", lIndex, REG_SZ
    CreateNewKey HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links\" & lIndex
    SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links\" & lIndex, "index", lIndex, REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links\" & lIndex, "url", Trim(Text1.Text), REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, "Software\ModCon\Toolbar\web\links\" & lIndex, "command", Trim(Text2.Text), REG_SZ
    frmBrowser.prcFillLinks
        
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SetNewShape
End Sub
Private Sub SetNewShape()
    With Shape1
        .BackColor = &H8000000F
        .BackStyle = 0
        .BorderColor = &H8000000C
        .BorderStyle = 1
        .BorderWidth = 1
        .FillStyle = 1
        .Shape = 4
    End With

End Sub
