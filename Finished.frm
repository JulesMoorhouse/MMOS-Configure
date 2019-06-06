VERSION 5.00
Begin VB.Form frmFinished 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   3
      Top             =   4485
      Width           =   4680
      _extentx        =   8255
      _extenty        =   1244
   End
   Begin VB.Label Label1 
      Caption         =   "Network Install - Finished"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmFinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

    Load frmTables
    Unload Me
    
    frmTables.Show

End Sub

Private Sub cmdOK_Click()

    End

End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    
End Sub


Private Sub Form_Paint()

    Const lintStatusHeight As Integer = 270
    Const lintMenuBarHeight As Integer = 300
    Dim lintBottomPanelHeight As Integer: lintBottomPanelHeight = (ctlBottomLine1.Height / 2)
    
    With cmdOK
        .Left = Me.width - .width - 200
        .top = (mdiMain.Height - mdiMain.sbStatusBar.Height - (ctlBottomLine1.Height) - lintMenuBarHeight) - (.Height / 2)
    End With
    
    With cmdBack
        .top = cmdOK.top
        .Left = cmdOK.Left - 120 - .width
    End With
    
End Sub
