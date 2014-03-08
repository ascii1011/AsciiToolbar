VERSION 5.00
Begin VB.Form frmResCntrl 
   Caption         =   "Resolution"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command4 
         Caption         =   "reset"
         Height          =   255
         Left            =   4140
         TabIndex        =   11
         Top             =   3180
         Width           =   795
      End
      Begin VB.CommandButton Command3 
         Caption         =   "set"
         Height          =   255
         Left            =   4140
         TabIndex        =   10
         Top             =   1620
         Width           =   795
      End
      Begin VB.CommandButton Command2 
         Caption         =   "grab list"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   3180
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Height          =   255
         Left            =   4020
         TabIndex        =   8
         Top             =   480
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "grab"
         Height          =   255
         Left            =   4020
         TabIndex        =   7
         Top             =   840
         Width           =   795
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   180
         TabIndex        =   6
         Top             =   780
         Width           =   3195
      End
      Begin VB.TextBox Text2 
         Height          =   255
         Left            =   4140
         TabIndex        =   5
         Top             =   1980
         Width           =   795
      End
      Begin VB.TextBox Text3 
         Height          =   255
         Left            =   4140
         TabIndex        =   4
         Top             =   2280
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   255
         Left            =   4140
         TabIndex        =   3
         Top             =   2580
         Width           =   795
      End
      Begin VB.TextBox Text5 
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox Text6 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "supported res's:"
         Height          =   195
         Left            =   4020
         TabIndex        =   19
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         Height          =   195
         Left            =   3960
         TabIndex        =   18
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Y"
         Height          =   195
         Left            =   3960
         TabIndex        =   17
         Top             =   2340
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "Quality"
         Height          =   195
         Left            =   3600
         TabIndex        =   16
         Top             =   2580
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "bit"
         Height          =   195
         Left            =   4980
         TabIndex        =   15
         Top             =   2640
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "X"
         Height          =   195
         Left            =   420
         TabIndex        =   14
         Top             =   420
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Y"
         Height          =   195
         Left            =   1740
         TabIndex        =   13
         Top             =   420
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "Current Resolution:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmResCntrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    If frmToolbar.iPosition < 4 Then
        Me.Top = 1000
        Me.Left = 2000
    Else
        Me.Top = 1000
        Me.Left = 23000
    End If
    
    prcDetermineResolution Xpix, Ypix
    Text5.Text = Xpix
    Text6.Text = Ypix
    
    Text2.Text = 1280
    Text3.Text = 800
    Text4.Text = 32
End Sub

Private Sub Command1_Click()
    Text1.Text = GetSupportedResolutionCount
End Sub

Private Sub Command2_Click()
    Dim lIndex As Integer
    Dim sStr As String
    
    List1.Clear
    sStr = ""
On Error GoTo Err1:
    If Trim(Text1.Text) <> "" Or Trim(Text1.Text) <> 0 Then
        
        For lIndex = 0 To Trim(Text1.Text)
            sStr = lIndex & ") X," & GetSupportedResolutionIndex(lIndex, "X")
            sStr = sStr & ",Y," & GetSupportedResolutionIndex(lIndex, "Y")
            sStr = sStr & ",B," & GetSupportedResolutionIndex(lIndex, "B")
            sStr = sStr & ",R," & GetSupportedResolutionIndex(lIndex, "R")
            List1.AddItem sStr
        Next
    End If
Err1:
    Resume Next
End Sub

Private Sub Command3_Click()
    Dim lX As Long, lY As Long, lB As Long
    
    If Trim(Text2.Text) = "" Or Trim(Text3.Text) = "" Or Trim(Text4.Text) = "" Then
        Exit Sub
    End If
    
    lX = Trim(Text2.Text)
    lY = Trim(Text3.Text)
    lB = Trim(Text4.Text)
    
    SetResolution lX, lY, lB
    'Pause 5
    'ResetResolution
End Sub

Private Sub Command4_Click()
    ResetResolution
End Sub

Private Sub List1_Click()
    Dim stemparray
    
    stemparray = Split(List1.List(List1.ListIndex), ",")
    If UBound(stemparray) > 0 And UBound(stemparray) <= 7 Then
        Text2.Text = stemparray(1)
        Text3.Text = stemparray(3)
        Text4.Text = stemparray(5)
    End If
End Sub

