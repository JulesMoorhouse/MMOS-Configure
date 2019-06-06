VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfigure 
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14730
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11025
   ScaleWidth      =   14730
   WindowState     =   2  'Maximized
   Begin MMOS.ctlBanner ctlBanner1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   14730
      _extentx        =   25982
      _extenty        =   1852
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Set &Defaults"
      Height          =   360
      Left            =   120
      TabIndex        =   29
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtHelpTooltip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   22
      Tag             =   "9999"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   0
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtStaticLdr 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   4215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   360
      Left            =   13320
      TabIndex        =   19
      Top             =   10560
      Width           =   1215
   End
   Begin MMOS.ctlBottomLine ctlBottomLine1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   20
      Top             =   10320
      Width           =   14730
      _extentx        =   25982
      _extenty        =   1244
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   21
      Top             =   0
      Width           =   0
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   1
      Left            =   5520
      TabIndex        =   5
      Top             =   1800
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   2
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   3
      Left            =   5520
      TabIndex        =   7
      Top             =   2760
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   4
      Left            =   5520
      TabIndex        =   8
      Top             =   3240
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   5
      Left            =   5520
      TabIndex        =   9
      Top             =   3720
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   6
      Left            =   5520
      TabIndex        =   10
      Top             =   4200
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   7
      Left            =   5520
      TabIndex        =   11
      Top             =   4680
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   8
      Left            =   5520
      TabIndex        =   12
      Top             =   5160
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   9
      Left            =   5520
      TabIndex        =   13
      Top             =   5640
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   13
      Left            =   5520
      TabIndex        =   17
      Top             =   9000
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   480
      Index           =   14
      Left            =   5520
      TabIndex        =   18
      Top             =   9480
      Width           =   8895
      _extentx        =   15690
      _extenty        =   847
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   855
      Index           =   10
      Left            =   5520
      TabIndex        =   14
      Top             =   6120
      Width           =   8895
      _extentx        =   15690
      _extenty        =   1508
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   855
      Index           =   11
      Left            =   5520
      TabIndex        =   15
      Top             =   7080
      Width           =   8895
      _extentx        =   15690
      _extenty        =   1508
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin MMOS.ctlPathPicker ctlPathPicker 
      Height          =   855
      Index           =   12
      Left            =   5520
      TabIndex        =   16
      Top             =   8055
      Width           =   8895
      _extentx        =   15690
      _extenty        =   1508
      hasparamdesc    =   0   'False
      textparam       =   ""
      textdesc        =   ""
   End
   Begin VB.Label Label7 
      Caption         =   "You will need a Static.ldr file to use the next screen."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Label Label6 
      Caption         =   "If you need to make changes to your file in the future, use the 'Load' button."
      Height          =   855
      Left            =   120
      TabIndex        =   27
      Top             =   6480
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "Once you have finished and have no warnings, use the 'Save' button before to save your changes."
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Please provide the files / paths shown on the right"
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "You will also need a copy of this file in each project folder if you making changes to the source code."
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   $"Configure.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lblStaticLdr 
      Caption         =   "Static.ldr file"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   2415
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbooFormLoaded As Boolean

'Private Sub GetLocalFields()
'
'    Dim serverPath As String: serverPath = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "ServerPath")
'    ctlPathPicker(FilePathArrayIndexes.serverPath).Text = serverPath
'
'    Dim SrvTestPth As String: SrvTestPth = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "SrvTestPth")
'    ctlPathPicker(FilePathArrayIndexes.ServerTestPath).Text = SrvTestPth
'
'    Dim SuppPath As String: SuppPath = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "SuppPath")
'    ctlPathPicker(FilePathArrayIndexes.SupportPath).Text = SuppPath
'
'    Dim SupTestPth As String: SupTestPth = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "SupTestPth")
'    ctlPathPicker(FilePathArrayIndexes.SupportTestPath).Text = SupTestPth
'
'    Dim strLocal As String: strLocal = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Local")
'    ctlPathPicker(FilePathArrayIndexes.LocalDatabase).Text = strLocal
'
'    Dim LocalTest As String: LocalTest = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "LocalTest")
'    ctlPathPicker(FilePathArrayIndexes.LocalTestDatabase).Text = LocalTest
'
'    Dim Central As String: Central = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Central")
'    ctlPathPicker(FilePathArrayIndexes.CentralDatabase).Text = Central
'
'    Dim CentralTest As String: CentralTest = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "CentraTest")
'    ctlPathPicker(FilePathArrayIndexes.CentralTestDatabase).Text = CentralTest
'
'    Dim Reps As String: Reps = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "Reps")
'    ctlPathPicker(FilePathArrayIndexes.ReportsDatabase).Text = Reps
'
'    Dim RepsTest As String: RepsTest = GetPrivateINI(AppPath & gconstrStaticIni, "DB", "RepsTest")
'    ctlPathPicker(FilePathArrayIndexes.ReportsTestDatabase).Text = RepsTest
'
'    Dim Logging As String: Logging = GetPrivateINI(AppPath & gconstrStaticIni, "Verbose Logging", "BSTAT")
'    ctlPathPicker(FilePathArrayIndexes.LoggingFile).Text = Logging
'
'    Dim parcel As String: parcel = GetPrivateINI(AppPath & gconstrStaticIni, "SysFileInfo", "PFEFile")
'    ctlPathPicker(FilePathArrayIndexes.ParcelExportFile).Text = parcel
'
'    DoEvents
'
'End Sub
'Private Sub SetLocalFields()
'
'    SetPrivateINI AppPath & gconstrStaticIni, "SysFileInfo", "ServerPath", ctlPathPicker(FilePathArrayIndexes.serverPath).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "SysFileInfo", "SrvTestPth", ctlPathPicker(FilePathArrayIndexes.ServerTestPath).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "SysFileInfo", "SuppPath", ctlPathPicker(FilePathArrayIndexes.SupportPath).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "SysFileInfo", "SupTestPth", ctlPathPicker(FilePathArrayIndexes.SupportTestPath).Text
'
'    SetPrivateINI AppPath & gconstrStaticIni, "DB", "Local", ctlPathPicker(FilePathArrayIndexes.LocalDatabase).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "DB", "LocalTest", ctlPathPicker(FilePathArrayIndexes.LocalTestDatabase).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "DB", "Central", ctlPathPicker(FilePathArrayIndexes.CentralDatabase).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "DB", "CentraTest", ctlPathPicker(FilePathArrayIndexes.CentralTestDatabase).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "DB", "Reps", ctlPathPicker(FilePathArrayIndexes.ReportsDatabase).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "DB", "RepsTest", ctlPathPicker(FilePathArrayIndexes.ReportsTestDatabase).Text
'
'    SetPrivateINI AppPath & gconstrStaticIni, "Verbose Logging", "BSTAT", ctlPathPicker(FilePathArrayIndexes.LoggingFile).Text
'    SetPrivateINI AppPath & gconstrStaticIni, "SysFileInfo", "PFEFile", ctlPathPicker(FilePathArrayIndexes.ParcelExportFile).Text
'
'End Sub
Private Sub AssignArray()

    With gstrStatic
        .strServerPath = ctlPathPicker(FilePathArrayIndexes.ServerPath).Text
        .strServerTestNewPath = ctlPathPicker(FilePathArrayIndexes.ServerTestPath).Text
        
        .strSupportPath = ctlPathPicker(FilePathArrayIndexes.SupportPath).Text
        .strSupportTestPath = ctlPathPicker(FilePathArrayIndexes.SupportTestPath).Text
        
        .strLocalDBFile = ctlPathPicker(FilePathArrayIndexes.LocalDatabase).Text
        .strLocalTestingDBFile = ctlPathPicker(FilePathArrayIndexes.LocalTestDatabase).Text
        
        .strCentralDBFile = ctlPathPicker(FilePathArrayIndexes.CentralDatabase).Text
        .strCentralTestingDBFile = ctlPathPicker(FilePathArrayIndexes.CentralTestDatabase).Text
        
        .strReportsDBFile = ctlPathPicker(FilePathArrayIndexes.ReportsDatabase).Text
        .strReportsTestingDBFile = ctlPathPicker(FilePathArrayIndexes.ReportsTestDatabase).Text
        
        .strPrograms(0).strProgram = ctlPathPicker(FilePathArrayIndexes.Program1).Text
        .strPrograms(0).strParam = ctlPathPicker(FilePathArrayIndexes.Program1).TextParam
        .strPrograms(0).strDesc = ctlPathPicker(FilePathArrayIndexes.Program1).TextDesc
        
        .strPrograms(1).strProgram = ctlPathPicker(FilePathArrayIndexes.Program2).Text
        .strPrograms(1).strParam = ctlPathPicker(FilePathArrayIndexes.Program2).TextParam
        .strPrograms(1).strDesc = ctlPathPicker(FilePathArrayIndexes.Program2).TextDesc
          
        .strPrograms(2).strProgram = ctlPathPicker(FilePathArrayIndexes.Program3).Text
        .strPrograms(2).strParam = ctlPathPicker(FilePathArrayIndexes.Program3).TextParam
        .strPrograms(2).strDesc = ctlPathPicker(FilePathArrayIndexes.Program3).TextDesc
        
        .strVerLogBStatus = ctlPathPicker(FilePathArrayIndexes.LoggingFile).Text
        .strPFElecFile = ctlPathPicker(FilePathArrayIndexes.ParcelExportFile).Text
    End With
    
End Sub
Private Sub AssignFieldsFromArray()

   With gstrStatic
        ctlPathPicker(FilePathArrayIndexes.ServerPath).Text = .strServerPath
        ctlPathPicker(FilePathArrayIndexes.ServerTestPath).Text = .strServerTestNewPath
        
        ctlPathPicker(FilePathArrayIndexes.SupportPath).Text = .strSupportPath
        ctlPathPicker(FilePathArrayIndexes.SupportTestPath).Text = .strSupportTestPath
        
        ctlPathPicker(FilePathArrayIndexes.LocalDatabase).Text = .strLocalDBFile
        ctlPathPicker(FilePathArrayIndexes.LocalTestDatabase).Text = .strLocalTestingDBFile
        
        ctlPathPicker(FilePathArrayIndexes.CentralDatabase).Text = .strCentralDBFile
        ctlPathPicker(FilePathArrayIndexes.CentralTestDatabase).Text = .strCentralTestingDBFile
        
        ctlPathPicker(FilePathArrayIndexes.ReportsDatabase).Text = .strReportsDBFile
        ctlPathPicker(FilePathArrayIndexes.ReportsTestDatabase).Text = .strReportsTestingDBFile
        
        ctlPathPicker(FilePathArrayIndexes.LoggingFile).Text = .strVerLogBStatus
        ctlPathPicker(FilePathArrayIndexes.ParcelExportFile).Text = .strPFElecFile
        
        ctlPathPicker(FilePathArrayIndexes.Program1).Text = .strPrograms(0).strProgram
        ctlPathPicker(FilePathArrayIndexes.Program1).TextParam = .strPrograms(0).strParam
        ctlPathPicker(FilePathArrayIndexes.Program1).TextDesc = .strPrograms(0).strDesc
        
        ctlPathPicker(FilePathArrayIndexes.Program2).Text = .strPrograms(1).strProgram
        ctlPathPicker(FilePathArrayIndexes.Program2).TextParam = .strPrograms(1).strParam
        ctlPathPicker(FilePathArrayIndexes.Program2).TextDesc = .strPrograms(1).strDesc
        
        ctlPathPicker(FilePathArrayIndexes.Program3).Text = .strPrograms(2).strProgram
        ctlPathPicker(FilePathArrayIndexes.Program3).TextParam = .strPrograms(2).strParam
        ctlPathPicker(FilePathArrayIndexes.Program3).TextDesc = .strPrograms(2).strDesc
    End With
    
End Sub
Private Sub SetAnyErrors()

    Dim fs As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim Index As Integer: Index = 0
    Dim lstrServerPath As String: lstrServerPath = ctlPathPicker(FilePathArrayIndexes.ServerPath).Text
    Dim lstrServerTestPath As String: lstrServerTestPath = ctlPathPicker(FilePathArrayIndexes.ServerTestPath).Text
    Dim lstrLocalPath As String: lstrLocalPath = AppPath
    
    For Index = 0 To cTOTAL_STATIC_ATTRIBUTES
        Dim strFile As String: strFile = ctlPathPicker(Index).Text
        Dim objPathType As PathType: objPathType = ctlPathPicker(Index).ItemPathType
        Dim booHasError As Boolean: booHasError = False
        Dim strErrorText As String
        Dim strParentPath As String: strParentPath = ""
        
        Select Case mobjStaticAttributes(Index).ParentFolderIndex
        Case FilePathArrayIndexes.ServerPath
            strParentPath = lstrServerPath
        Case FilePathArrayIndexes.ServerTestPath
            strParentPath = lstrServerTestPath
        Case FilePathArrayIndexes.LocalProgramFilesFolder
            strParentPath = lstrLocalPath
        Case FilePathArrayIndexes.NotApplicable
            strParentPath = ""
        Case Else
            MsgBox "ERROR: Unknown parent path!", , gconstrTitlPrefix & "Configure"
        End Select
        
        FindSetAnError Index, fs, strFile, objPathType, strParentPath, booHasError, strErrorText
        
        'ctlPathPicker(Index).hasError = True 'in the next code can be ovveridden
        ctlPathPicker(Index).SetError booHasError, strErrorText
    Next Index
    DoEvents
    
End Sub
Function AnyErrors() As Boolean

    SetAnyErrors

    AnyErrors = False
    
    Dim Index As Integer
    
    For Index = 0 To cTOTAL_STATIC_ATTRIBUTES
        If ctlPathPicker(Index).hasError = True Then
            AnyErrors = True
            Exit For
        End If
    Next Index

End Function
Function SetPathPickers()

Dim Index As Integer

    For Index = 0 To cTOTAL_STATIC_ATTRIBUTES
        SetPathPicker (Index)
    Next Index
    
End Function
Sub SetPathPicker(Index As Integer)

    ctlPathPicker(Index).ItemPathType = mobjStaticAttributes(Index).ItemPathType
    ctlPathPicker(Index).FileExtType = mobjStaticAttributes(Index).ext
    ctlPathPicker(Index).Label = mobjStaticAttributes(Index).Label
    ctlPathPicker(Index).HelpTooltip = mobjStaticAttributes(Index).Tooltip
    ctlPathPicker(Index).HasParamDesc = mobjStaticAttributes(Index).HasParamDesc
    ctlPathPicker(Index).HelpTooltipParam = mobjStaticAttributes(Index).TooltipParam
    ctlPathPicker(Index).HelpTooltipDesc = mobjStaticAttributes(Index).TooltipDesc
    ctlPathPicker(Index).ParentObj = Me
    
End Sub
Sub SetDefaults()

    ctlPathPicker(FilePathArrayIndexes.LocalDatabase).Text = "Local.mdb"
    ctlPathPicker(FilePathArrayIndexes.LocalTestDatabase).Text = "LocalTest.mdb"
    
    ctlPathPicker(FilePathArrayIndexes.CentralDatabase).Text = "Central.mdb"
    ctlPathPicker(FilePathArrayIndexes.CentralTestDatabase).Text = "CentralTest.mdb"
    
    ctlPathPicker(FilePathArrayIndexes.ReportsDatabase).Text = "Reps.mdb"
    ctlPathPicker(FilePathArrayIndexes.ReportsTestDatabase).Text = "RepsTest.mdb"
    
    ctlPathPicker(FilePathArrayIndexes.Program1).Text = "Mmos.exe"
    ctlPathPicker(FilePathArrayIndexes.Program1).TextParam = "X"
    ctlPathPicker(FilePathArrayIndexes.Program1).TextDesc = "Client program"
    
    ctlPathPicker(FilePathArrayIndexes.Program2).Text = "MAdmin.exe"
    ctlPathPicker(FilePathArrayIndexes.Program2).TextParam = "ADMIN"
    ctlPathPicker(FilePathArrayIndexes.Program2).TextDesc = "Admin Program"
    
    ctlPathPicker(FilePathArrayIndexes.Program3).Text = "MReps.exe"
    ctlPathPicker(FilePathArrayIndexes.Program3).TextParam = "REPORT"
    ctlPathPicker(FilePathArrayIndexes.Program3).TextDesc = "Reporting Program"
    
End Sub
Public Sub cmdBack_Click()

'    Load frmNetInstall
'    Unload Me
'
'    frmNetInstall.Show

    Unload Me
    frmAbout.Show

End Sub

Private Sub cmdDefaults_Click()

    SetDefaults
    
End Sub

Private Sub cmdProceed_Click()
    
    Load frmTables
    Unload Me

    frmTables.Show

End Sub
Private Sub cmdLoad_Click()

    txtStaticLdr.Text = SelectFile(CommonDialog1, txtStaticLdr.Text, "ldr", "Static.ldr File|*.ldr", cdlOFNFileMustExist, False)
    
    If Trim$(txtStaticLdr.Text & "") <> "" And Dir(txtStaticLdr.Text) <> "" Then
        Busy True
        Decrypt txtStaticLdr.Text, "STATIC"
        AssignFieldsFromArray
        SetAnyErrors
        Busy False
    Else
        MsgBox "Please select a valid file to save your static.ldr", , gconstrTitlPrefix & "Configure"
    End If
    
End Sub
Private Sub cmdSave_Click()

    Dim lbooCommonDialogCancelled As Boolean
    Dim lstrDefaultFile As String: lstrDefaultFile = "Static.ldr"
    If Trim$(txtStaticLdr.Text) <> "" Then
        lstrDefaultFile = txtStaticLdr.Text
    End If
    
    txtStaticLdr.Text = SelectFile(CommonDialog1, lstrDefaultFile, "ldr", "Static.ldr File|*.ldr", _
        cdlOFNCreatePrompt, lbooCommonDialogCancelled, False)
    
    If lbooCommonDialogCancelled = False Then
        AssignArray
        If AnyErrors = False Then
            Busy True
            Encrypt txtStaticLdr.Text, "STATIC"
            SetPrivateINI AppPath & "configure.ini", "General", "LastGood", txtStaticLdr.Text
            Busy False
            
            Dim lstrMessage As String
            Dim lintRetVal As Integer
            lstrMessage = "Static data has been saved!" & vbCrLf & vbCrLf & _
                "Would you also like to copy this file to the server folder and current Program Files Folder? (Recommended)"
    
            lintRetVal = MsgBox(lstrMessage, vbYesNo, gconstrTitlPrefix & "Configure")

            If lintRetVal = vbYes Then
                Dim lstrDestFile As String: lstrDestFile = gstrStatic.strServerPath & "Static.ldr"
                If UCase$(txtStaticLdr.Text) <> UCase$(lstrDestFile) Then
                    FileCopy txtStaticLdr.Text, lstrDestFile
                End If
                lstrDestFile = AppPath() & "Static.ldr"
                If UCase$(txtStaticLdr.Text) <> UCase$(lstrDestFile) Then
                    FileCopy txtStaticLdr.Text, AppPath() & "Static.ldr"
                End If
                MsgBox "File saved / copied successfully!", , gconstrTitlPrefix & "Configure"
            End If
        Else
            MsgBox "Please fix the errors which are highlighted!", , gconstrTitlPrefix & "Configure"
        End If
    Else
        MsgBox "Please select a valid file to save your static.ldr", , gconstrTitlPrefix & "Configure"
    End If
    
End Sub
Private Sub ctlPathPicker_OnChange(Index As Integer, ByVal Text As String, ByVal pobjPathType As PathType)

    If mbooFormLoaded = True Then
'        Dim fs As Object
'        Set fs = CreateObject("Scripting.FileSystemObject")
'        Dim lstrServerPath As String: lstrServerPath = ctlPathPicker(FilePathArrayIndexes.ServerPath).Text
'        Dim lstrServerTestPath As String: lstrServerTestPath = ctlPathPicker(FilePathArrayIndexes.ServerTestPath).Text
'        Dim lstrLocalPath As String: lstrLocalPath = AppPath
'
'        Dim strParentPath As String: strParentPath = ""
'
'        Select Case mobjStaticAttributes(Index).ParentFolderIndex
'        Case FilePathArrayIndexes.ServerPath
'            strParentPath = lstrServerPath
'        Case FilePathArrayIndexes.ServerTestPath
'            strParentPath = lstrServerTestPath
'        Case FilePathArrayIndexes.LocalProgramFilesFolder
'            strParentPath = lstrLocalPath
'        Case FilePathArrayIndexes.NotApplicable
'            strParentPath = ""
'        Case Else
'            MsgBox "ERROR: Unknown parent path!"
'        End Select
'
'        Dim booHasError As Boolean
'        Dim strErrorText As String
'
'        FindSetAnError Index, fs, Text, pobjPathType, strParentPath, booHasError, strErrorText
'
'        ctlPathPicker(Index).hasError = True 'in the next code can be ovveridden
'        ctlPathPicker(Index).SetError booHasError, strErrorText
'
        SetAnyErrors

        DoEvents
    End If
    
End Sub
   
Private Sub ctlPathPicker_OnFileSelected(Index As Integer, ByVal Text As String, ByVal pobjPathType As PathType)

    Dim lstrServerPath As String: lstrServerPath = ctlPathPicker(FilePathArrayIndexes.ServerPath).Text
    Dim lstrServerTestPath As String: lstrServerTestPath = ctlPathPicker(FilePathArrayIndexes.ServerTestPath).Text
    Dim lstrLocalPath As String: lstrLocalPath = AppPath
        
    Dim strParentPath As String: strParentPath = ""
    
    Select Case mobjStaticAttributes(Index).ParentFolderIndex
    Case FilePathArrayIndexes.ServerPath
        strParentPath = lstrServerPath
    Case FilePathArrayIndexes.ServerTestPath
        strParentPath = lstrServerTestPath
    Case FilePathArrayIndexes.LocalProgramFilesFolder
        strParentPath = lstrLocalPath
    Case FilePathArrayIndexes.NotApplicable
        strParentPath = ""
    Case Else
        MsgBox "ERROR: Unknown parent path!", , gconstrTitlPrefix & "Configure"
    End Select
        
    Select Case pobjPathType
    Case JustFile
        ctlPathPicker(Index).Text = Replace(Text, strParentPath, "", 1, -1, vbTextCompare)
    End Select
    
End Sub

Private Sub ctlPathPicker_OnHelpHover(Index As Integer, ByVal HelpText As String, Left As Integer, Top As Integer)

    With txtHelpTooltip
        If CInt(.Tag) <> Index Then
            .Text = HelpText
            .Top = ctlPathPicker(Index).Top + Top + 400
            .Width = TextWidth(HelpText) + 120
            .Height = TextHeight(HelpText) + 120
            .Left = (ctlPathPicker(Index).Left + Left) - (.Width / 2)
            If (.Top + .Height) > ctlBottomLine1.Top Then
                .Top = (.Top - .Height) - 400
            End If
            .Tag = Index
            .Visible = True
        End If
    End With
    
End Sub

Private Sub ctlPathPicker_onHelpHoverOff(Index As Integer)

    txtHelpTooltip.Visible = False
    txtHelpTooltip.Tag = 9999
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtHelpTooltip.Visible = False
    txtHelpTooltip.Tag = 9999
    
End Sub
Private Sub Form_Activate()

    If mbooFormLoaded = False Then
        Dim lstrLastGood As String
        lstrLastGood = GetPrivateINI(AppPath & "configure.ini", "General", "LastGood")
        If lstrLastGood <> "" And Dir(lstrLastGood) <> "" Then
            txtStaticLdr.Text = lstrLastGood
            Busy True
            DoEvents
            'GetLocalFields
            Decrypt txtStaticLdr.Text, "STATIC"
            AssignFieldsFromArray
            SetAnyErrors
            Busy False
        Else
            SetDefaults
        End If
        mbooFormLoaded = True
    End If
    
End Sub

Private Sub Form_Load()

    If gbooJustPreLoading Then
        Exit Sub
    End If
    
    NameForm Me
    SetPathPickers
    ShowBanner Me, gconstrConfigFilesPaths
    
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
Private Sub Form_Unload(Cancel As Integer)

    mbooFormLoaded = False
    
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
