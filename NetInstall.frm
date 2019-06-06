VERSION 5.00
Begin VB.Form frmNetInstall 
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12810
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   11400
      TabIndex        =   7
      Top             =   8400
      Width           =   1215
   End
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "&Open folder"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   8175
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   1244
   End
   Begin VB.Label Label8 
      Caption         =   " Although this process could have been provided, you may encounter file access isues which will be easier for you to resolve."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   10455
   End
   Begin VB.Label Label7 
      Caption         =   "If you haven't already you'll need to setup a folder where all the users of MMOS can access it."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   10455
   End
   Begin VB.Label Label1 
      Caption         =   "Here you'll be provided with a zip file which you'll need install onto a network drive."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   10455
   End
   Begin VB.Label Label6 
      Caption         =   "The location must be accessible by this program and users who will use MMOS."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   8415
   End
   Begin VB.Label Label5 
      Caption         =   "'Server.zip' should be extracted to a server or network folder."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   "You'll find a file called 'Server.zip' in the current folder, click on 'Open Folder' to open that folder."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   9135
   End
   Begin VB.Label Label2 
      Caption         =   "Welcome to the MMOS network configuration tool."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5055
   End
End
Attribute VB_Name = "frmNetInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdBack_Click()

    Unload Me
    frmAbout.Show
    
End Sub

Private Sub cmdOpenFolder_Click()

    Dim cQUOTE As String: cQUOTE = Chr$(34)
    Dim lstrShellPath As String
    lstrShellPath = Replace(AppPath, cQUOTE, cQUOTE & cQUOTE)
    Shell "explorer.exe /e, " & lstrShellPath, vbNormalFocus
    
End Sub

Private Sub cmdProceed_Click()

    Load frmConfigure
    Unload Me
    
    frmConfigure.Show

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    ShowBanner Me, gconstrConfigNetInstall
    
End Sub
Private Sub Form_Paint()

    Const lintStatusHeight As Integer = 270
    Const lintMenuBarHeight As Integer = 300
    Dim lintBottomPanelHeight As Integer: lintBottomPanelHeight = (ctlBottomLine1.Height / 2)
    
    With cmdBack
        .Left = Me.Width - .Width - 200
        .Top = (mdiMain.Height - mdiMain.sbStatusBar.Height - (ctlBottomLine1.Height) - lintMenuBarHeight) - (.Height / 2)
    End With
    
End Sub
