VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTables 
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13875
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   WindowState     =   2  'Maximized
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   1852
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   12480
      TabIndex        =   10
      Top             =   9000
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtStaticLdr 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   360
      Left            =   7560
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdReattach 
      Caption         =   "&Reattach"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   8790
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   1244
   End
   Begin VB.Label Label7 
      Caption         =   $"Tables.frx":0000
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Width           =   10815
   End
   Begin VB.Label Label1 
      Caption         =   "If you haven't already done so, please create a 'static.ldr' file within the 'Files && Paths' screen."
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Label Label6 
      Caption         =   $"Tables.frx":0090
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   13095
   End
   Begin VB.Label Label5 
      Caption         =   "You’ll also need to copy ‘static.ldr’ to the folder where MMOS installed."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   9255
   End
   Begin VB.Label Label4 
      Caption         =   "You will need to install this on every machine you where you want to use MMOS."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   9135
   End
   Begin VB.Label Label3 
      Caption         =   "OK now you can download the client setup from ..."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   8415
   End
   Begin VB.Label Label2 
      Caption         =   "OK, now you have all the files setup, we need to re-attach some tables."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   7575
   End
   Begin VB.Label lblStaticLdr 
      Caption         =   "Previous static.ldr file"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbooFormLoaded As Boolean

Private Type StarticText
    Text As String
    TextParam As String
    TextDesc As String
End Type

Dim mstrFilePaths(cTOTAL_STATIC_ATTRIBUTES) As StarticText

Private Sub AssignFieldsFromArray()

   With gstrStatic
        mstrFilePaths(FilePathArrayIndexes.ServerPath).Text = .strServerPath
        mstrFilePaths(FilePathArrayIndexes.ServerTestPath).Text = .strServerTestNewPath
        
        mstrFilePaths(FilePathArrayIndexes.SupportPath).Text = .strSupportPath
        mstrFilePaths(FilePathArrayIndexes.SupportTestPath).Text = .strSupportTestPath
        
        mstrFilePaths(FilePathArrayIndexes.LocalDatabase).Text = .strLocalDBFile
        mstrFilePaths(FilePathArrayIndexes.LocalTestDatabase).Text = .strLocalTestingDBFile
        
        mstrFilePaths(FilePathArrayIndexes.CentralDatabase).Text = .strCentralDBFile
        mstrFilePaths(FilePathArrayIndexes.CentralTestDatabase).Text = .strCentralTestingDBFile
        
        mstrFilePaths(FilePathArrayIndexes.ReportsDatabase).Text = .strReportsDBFile
        mstrFilePaths(FilePathArrayIndexes.ReportsTestDatabase).Text = .strReportsTestingDBFile
        
        mstrFilePaths(FilePathArrayIndexes.Program1).Text = .strPrograms(0).strProgram
        mstrFilePaths(FilePathArrayIndexes.Program1).TextParam = .strPrograms(0).strParam
        mstrFilePaths(FilePathArrayIndexes.Program1).TextDesc = .strPrograms(0).strDesc
        
        mstrFilePaths(FilePathArrayIndexes.Program2).Text = .strPrograms(1).strProgram
        mstrFilePaths(FilePathArrayIndexes.Program2).TextParam = .strPrograms(1).strParam
        mstrFilePaths(FilePathArrayIndexes.Program2).TextDesc = .strPrograms(1).strDesc
        
        mstrFilePaths(FilePathArrayIndexes.Program3).Text = .strPrograms(2).strProgram
        mstrFilePaths(FilePathArrayIndexes.Program3).TextParam = .strPrograms(2).strParam
        mstrFilePaths(FilePathArrayIndexes.Program3).TextDesc = .strPrograms(2).strDesc
        
        mstrFilePaths(FilePathArrayIndexes.LoggingFile).Text = .strVerLogBStatus
        mstrFilePaths(FilePathArrayIndexes.ParcelExportFile).Text = .strPFElecFile
    End With
    
End Sub
Private Sub LoadCheckFile(pstrFile As String)

    Dim fs As Object
    Dim booHasError As Boolean
    
    Busy True
    Decrypt txtStaticLdr.Text, "STATIC"
    AssignFieldsFromArray
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim Index As Integer: Index = 0

    For Index = 0 To cTOTAL_STATIC_ATTRIBUTES
        Dim strFile As String: strFile = mstrFilePaths(Index).Text
        Dim objPathType As PathType: objPathType = mobjStaticAttributes(Index).ItemPathType
        Dim strParentPath As String: strParentPath = ""
        
        Select Case mobjStaticAttributes(Index).ParentFolderIndex
        Case FilePathArrayIndexes.ServerPath, FilePathArrayIndexes.ServerTestPath
            strParentPath = mstrFilePaths(mobjStaticAttributes(Index).ParentFolderIndex).Text
        Case FilePathArrayIndexes.LocalProgramFilesFolder
            strParentPath = AppPath
        Case FilePathArrayIndexes.NotApplicable
            strParentPath = ""
        Case Else
            MsgBox "ERROR: Unknown parent path!", , gconstrTitlPrefix & "Tables"
        End Select
        
        booHasError = False
        
        Dim strErrorText As String
        FindSetAnError Index, fs, strFile, objPathType, strParentPath, booHasError, strErrorText
        
        If booHasError = True Then
            Exit For
        End If
    Next Index
    
    If booHasError = True Then
        MsgBox "Sorry this file contains files or paths which are invalid!", , gconstrTitlPrefix & "Tables"
        cmdReattach.Enabled = False
    Else
        cmdReattach.Enabled = True
    End If
    
    DoEvents
    
    Busy False
    
End Sub
Public Sub cmdBack_Click()

'    Load frmConfigure
'    Unload Me
'
'    frmConfigure.Show

    Unload Me
    frmAbout.Show
    
End Sub
Private Sub cmdLoad_Click()

    txtStaticLdr.Text = SelectFile(CommonDialog1, txtStaticLdr.Text, "ldr", "Static.ldr File|*.ldr", cdlOFNFileMustExist, False)
    
    If Trim$(txtStaticLdr.Text & "") <> "" And Dir(txtStaticLdr.Text) <> "" Then
        LoadCheckFile txtStaticLdr.Text
    End If
        
End Sub
Private Sub cmdReattach_Click()
Dim lstrServerDB As String
Dim lstrCentralDBName As String
Dim lstrExtReportingDBName As String
Dim lstrLocalDBName As String
Dim lintRetVal
        
    lintRetVal = MsgBox("This feature will update the linked tables in the local database." & vbCrLf & vbCrLf & _
        "Are you sure you wish to proceed?", vbYesNo, gconstrTitlPrefix & "Auto Linked Table")

    If lintRetVal = vbYes Then
        If Trim$(gstrStatic.strServerPath) <> "" Then
            If Dir(gstrStatic.strServerPath) <> "" Then
            
                InitDb txtStaticLdr.Text
                
                Dim lintArrInc As Integer: lintArrInc = 0
                
                For lintArrInc = 0 To 1
                    If lintArrInc = 0 Then
                        lstrCentralDBName = gstrStatic.strCentralDBFile
                        lstrExtReportingDBName = gstrStatic.strReportsDBFile
                        lstrLocalDBName = gstrStatic.strLocalDBFile
                    Else
                        lstrCentralDBName = gstrStatic.strCentralTestingDBFile
                        lstrExtReportingDBName = gstrStatic.strReportsTestingDBFile
                        lstrLocalDBName = gstrStatic.strLocalTestingDBFile
                    End If
                    
                    If Not gdatLocalDatabase Is Nothing Then
                        gdatLocalDatabase.Close
                        Set gdatLocalDatabase = Nothing
                    End If
                    
                    Set gdatLocalDatabase = OpenDatabase(lstrLocalDBName, , False)
                    
                    TableDetach "ListDetailsMaster", gdatLocalDatabase
                    TableDetach "ListsMaster", gdatLocalDatabase
                    TableDetach "OrderLinesMaster", gdatLocalDatabase
                    TableDetach "ProductsMaster", gdatLocalDatabase
                    TableDetach "System", gdatLocalDatabase
                    
                    TableAttach "ListDetailsMaster", gdatLocalDatabase, lstrCentralDBName
                    TableAttach "ListsMaster", gdatLocalDatabase, lstrCentralDBName
                    TableAttach "OrderLinesMaster", gdatLocalDatabase, lstrCentralDBName
                    TableAttach "ProductsMaster", gdatLocalDatabase, lstrCentralDBName
                    TableAttach "System", gdatLocalDatabase, lstrCentralDBName
                
                    Dim gdatExtReporting As Database
                    Set gdatExtReporting = OpenDatabase(lstrExtReportingDBName, , False)
                    TableDetach "AdviceNotes", gdatExtReporting
                    TableDetach "CashBook", gdatExtReporting
                    TableDetach "CustAccounts", gdatExtReporting
                    TableDetach "ListDetails", gdatExtReporting
                    TableDetach "Lists", gdatExtReporting
                    TableDetach "OrderLinesMaster", gdatExtReporting
                    TableDetach "Products", gdatExtReporting
                    TableDetach "Remarks", gdatExtReporting
                    
                    TableAttach "AdviceNotes", gdatExtReporting, lstrCentralDBName
                    TableAttach "CashBook", gdatExtReporting, lstrCentralDBName
                    TableAttach "CustAccounts", gdatExtReporting, lstrCentralDBName
                    TableAttach "ListDetails", gdatExtReporting, lstrLocalDBName
                    TableAttach "Lists", gdatExtReporting, lstrLocalDBName
                    TableAttach "OrderLinesMaster", gdatExtReporting, lstrCentralDBName
                    TableAttach "Products", gdatExtReporting, lstrLocalDBName
                    TableAttach "Remarks", gdatExtReporting, lstrCentralDBName
                    gdatExtReporting.Close
                    Set gdatExtReporting = Nothing
                    
                    gdatLocalDatabase.Close
                    Set gdatLocalDatabase = Nothing
                    
                    If lintArrInc = 0 Then
                        FileCopy gstrStatic.strLocalDBFile, gstrStatic.strServerPath & gstrStatic.strShortLocalDBFile
                        FileCopy gstrStatic.strReportsDBFile, gstrStatic.strServerPath & gstrStatic.strShortReportsDBFile
                    Else
                        FileCopy gstrStatic.strLocalTestingDBFile, gstrStatic.strServerPath & gstrStatic.strShortLocalTestingDBFile
                        FileCopy gstrStatic.strReportsTestingDBFile, gstrStatic.strServerPath & gstrStatic.strShortReportsTestingDBFile
                    End If
                Next lintArrInc
                
                Set gdatLocalDatabase = OpenDatabase(gstrStatic.strLocalDBFile, , False)
                
                MsgBox "Process complete!", , gconstrTitlPrefix & "Tables"
            Else
                MsgBox "Specified DB not found!", , gconstrTitlPrefix & "Auto Linked Table"
            End If
        Else
            MsgBox "You have not entered a path, process halted!", , gconstrTitlPrefix & "Auto Linked Table"
        End If
    End If
    
End Sub
Private Sub Form_Activate()

    If mbooFormLoaded = False Then
        Dim lstrLastGood As String
        lstrLastGood = GetPrivateINI(AppPath & "configure.ini", "General", "LastGood")
        If lstrLastGood <> "" And Dir(lstrLastGood) <> "" Then
            txtStaticLdr.Text = lstrLastGood
            LoadCheckFile txtStaticLdr.Text
        End If
        mbooFormLoaded = True
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    mbooFormLoaded = False
    
End Sub
Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    ShowBanner Me, gconstrConfigTables
    
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

Private Sub txtStaticLdr_Change()

    cmdReattach.Enabled = False

End Sub

Private Sub txtStaticLdr_GotFocus()

    On Error Resume Next
    
    txtStaticLdr.SelStart = 0
    txtStaticLdr.SelLength = Len(txtStaticLdr.Text)
        
    On Error GoTo 0
    
End Sub
Private Sub txtStaticLdr_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 65 And Shift = 2 Then
        txtStaticLdr.SelStart = 0
        txtStaticLdr.SelLength = Len(txtStaticLdr.Text)
    End If
    
End Sub
Private Sub txtStaticLdr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    txtStaticLdr.SelStart = 0
    txtStaticLdr.SelLength = Len(txtStaticLdr.Text)
        
    On Error GoTo 0
    
End Sub
