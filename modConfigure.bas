Attribute VB_Name = "modConfigure"
Option Explicit

Public Enum FilePathArrayIndexes
    ServerPath = 0
    ServerTestPath = 1
    SupportPath = 2
    SupportTestPath = 3
    LocalDatabase = 4
    LocalTestDatabase = 5
    CentralDatabase = 6
    CentralTestDatabase = 7
    ReportsDatabase = 8
    ReportsTestDatabase = 9
    Program1 = 10
    Program2 = 11
    Program3 = 12
    LoggingFile = 13
    ParcelExportFile = 14
    LocalProgramFilesFolder = 15
    NotApplicable = 99
End Enum

Public Enum PathType
    NotApplicable = 0 'should never be set
    JustFile = 10
    Path = 20
    FileWithPath = 30
End Enum

Public Enum FileExtType
    NotApplicable = 0 'should never be set
    mdb = 1
    txt = 2
    exe = 3
End Enum

Public Type StaticAttributes
    ItemPathType As PathType
    Tooltip As String
    Label As String
    ext As FileExtType
    ParentFolderIndex As Integer
    HasParamDesc As Boolean
    TooltipParam As String
    TooltipDesc As String
End Type

Public Const cTOTAL_STATIC_ATTRIBUTES As Integer = 14

Public mobjStaticAttributes(cTOTAL_STATIC_ATTRIBUTES) As StaticAttributes

Sub Main()

    gdatSystemStartTime = Now()
    
    gstrSystemRoute = srStandardRoute
    
    'lstrThisHelpFile = MCLDebugChoices
        
    SetSystemNames
    SetStaticAttributes
        
    frmSplash.Show
    frmSplash.Refresh
    
    Select Case CheckForOtherMMosprog(frmSplash.hwnd)
    Case True
        MsgBox "You may only run one " & gconstrProductFullName & " program at once!", , gconstrTitlPrefix & "Startup"
        Unhook
        End
    Case False
        'MsgBox "no other prog found!"
    End Select
    
    Busy True
    
    gbooJustPreLoading = True
    
    ShowStatus 0
    DoEvents
    
    Load frmConfigure
    Unload frmConfigure
    Load frmTables
    Unload frmTables
    Load frmNetInstall
    Unload frmNetInstall
    
    'CheckDBNotMoved
    gbooJustPreLoading = False
    
    Load frmAbout
    
    Unload frmSplash
    
    frmAbout.Show
    
    Busy False
    
End Sub

Sub SetStaticAttributes()

    With mobjStaticAttributes(FilePathArrayIndexes.ServerPath)
        .ItemPathType = Path
        .Label = "Server path"
        .Tooltip = _
            "The path for the server where users will all have " & vbCrLf & _
            "access."
            
        .ParentFolderIndex = FilePathArrayIndexes.NotApplicable
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.ServerTestPath)
        .ItemPathType = Path
        .Label = "Server test path"
        .Tooltip = _
            "A testing area path for the server where users will " & vbCrLf & _
            "all have access."
            
        .ParentFolderIndex = FilePathArrayIndexes.NotApplicable
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.SupportPath)
        .ItemPathType = Path
        .Label = "Support path"
        .Tooltip = _
            "The folder where the client setup will be placed."
            
        .ParentFolderIndex = FilePathArrayIndexes.NotApplicable
    End With

    With mobjStaticAttributes(FilePathArrayIndexes.SupportTestPath)
        .ItemPathType = Path
        .Label = "Support test path"
        .Tooltip = _
            "The folder where the testing area client setup will " & vbCrLf & _
            "be placed."
            
        .ParentFolderIndex = FilePathArrayIndexes.NotApplicable
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.LocalDatabase)
        .ItemPathType = JustFile
        .Label = "Local database"
        .Tooltip = "The name of the file for the local / caching " & vbCrLf & _
            "database. " & vbCrLf & _
            "You should select this file from the program files " & vbCrLf & _
            "folder (or where ever the current program is " & vbCrLf & _
            "installed). " & vbCrLf & _
            "e.g. Local.mdb"
            
        .ext = mdb
        .ParentFolderIndex = FilePathArrayIndexes.LocalProgramFilesFolder
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.LocalTestDatabase)
        .ItemPathType = JustFile
        .Label = "Local test database"
        .Tooltip = _
            "The name of the file for the local / caching " & vbCrLf & _
            "database for te testing area. " & vbCrLf & _
            "You should select this file from the program files " & vbCrLf & _
            "folder (or where ever the current program is " & vbCrLf & _
            "installed)." & vbCrLf & _
            "e.g. LocalTest.mdb"
            
        .ext = mdb
        .ParentFolderIndex = FilePathArrayIndexes.LocalProgramFilesFolder
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.CentralDatabase)
        .ItemPathType = JustFile
        .Label = "Central database"
        .Tooltip = _
            "The name of your central database file, you should " & vbCrLf & _
            "select this from the server path you chose above." & vbCrLf & _
            "e.g. Central.mdb"
            
        .ext = mdb
        .ParentFolderIndex = FilePathArrayIndexes.ServerPath
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.CentralTestDatabase)
        .ItemPathType = JustFile
        .Label = "Central test database"
        .Tooltip = _
            "The name of your testing area central database file, " & vbCrLf & _
            "you should select this from the server path you " & vbCrLf & _
            "choose above." & vbCrLf & _
            "e.g. CentralTest.mdb"
            
        .ext = mdb
        .ParentFolderIndex = FilePathArrayIndexes.ServerTestPath
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.ReportsDatabase)
        .ItemPathType = JustFile
        .Label = "Reports database"
        .Tooltip = _
            "The name of the file for the local reporting database. " & vbCrLf & _
            "You should select this file from the program files " & vbCrLf & _
            "folder (or where ever the current program is installed)." & vbCrLf & _
            "e.g. Reps.mdb"
            
        .ext = mdb
        .ParentFolderIndex = FilePathArrayIndexes.LocalProgramFilesFolder
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.ReportsTestDatabase)
        .ItemPathType = JustFile
        .Label = "Reports test database"
        .Tooltip = _
            "The name of the file for the local reporting database " & vbCrLf & _
            "for te testing area. " & vbCrLf & _
            "You should select this file from the program files " & vbCrLf & _
            "folder (or where ever the current program is installed)." & vbCrLf & _
            "e.g. RepsTest.mdb"
            
        .ext = mdb
        .ParentFolderIndex = FilePathArrayIndexes.LocalProgramFilesFolder
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.Program1)
        .ItemPathType = JustFile
        .Label = "Client program"
        .Tooltip = _
            "The name of the client program (this is used by the " & vbCrLf & _
            "Loader), you should select this from the server path " & vbCrLf & _
            "you choose above." & vbCrLf & _
            "e.g. MMOS.exe"
            
        .TooltipParam = _
            "The command line argument for the loader program used to " & vbCrLf & _
            "open the client program. You should leave this as X "
        
        .TooltipDesc = "A short description of the client program"
        
        .ext = exe
        .ParentFolderIndex = FilePathArrayIndexes.ServerPath
        .HasParamDesc = True
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.Program2)
        .ItemPathType = JustFile
        .Label = "Admin program"
        .Tooltip = _
            "The name of the admin program (this is used by the " & vbCrLf & _
            "Loader), you should select this from the server path " & vbCrLf & _
            "you choose above." & vbCrLf & _
            "e.g. MAdmin.exe"
            
        .TooltipParam = _
            "The command line argument for the loader program used to " & vbCrLf & _
            "open the admin program. You should leave this as ADMIN "
        
        .TooltipDesc = "A short description of the admin program"
        
        .ext = exe
        .ParentFolderIndex = FilePathArrayIndexes.ServerPath
        .HasParamDesc = True
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.Program3)
        .ItemPathType = JustFile
        .Label = "Reporting program"
        .Tooltip = _
            "The name of the reporting program (this is used by the " & vbCrLf & _
            "Loader), you should select this from the server path " & vbCrLf & _
            "you choose above." & vbCrLf & _
            "e.g. MReps.exe"
            
        .TooltipParam = _
            "The command line argument for the loader program used to " & vbCrLf & _
            "open the reporting program. You should leave this as " & vbCrLf & _
            "REPORTING "
        
        .TooltipDesc = "A short description of the reporting program"
        
        .ext = exe
        .ParentFolderIndex = FilePathArrayIndexes.ServerPath
        .HasParamDesc = True
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.LoggingFile)
        .ItemPathType = FileWithPath
        .Label = "Logging file"
        .Tooltip = _
            "This is a full path with file name for a logging file. " & vbCrLf & _
            "You can create this from anywhere on your local hard " & vbCrLf & _
            "disk." & vbCrLf & _
            "e.g. C:\logging.txt"
        
        .ext = txt
        .ParentFolderIndex = FilePathArrayIndexes.NotApplicable
    End With
    
    With mobjStaticAttributes(FilePathArrayIndexes.ParcelExportFile)
        .ItemPathType = FileWithPath
        .Label = "Parcel export file"
        .Tooltip = _
            "This is a full path with file name for the parcel " & vbCrLf & _
            "export file. You can create this from anywhere on your " & vbCrLf & _
            "local hard disk." & vbCrLf & _
            "e.g. C:\package.txt"
            
        .ext = txt
        .ParentFolderIndex = FilePathArrayIndexes.NotApplicable
    End With
    
End Sub
Public Function SelectFile(cd As CommonDialog, _
    pstrDefaultFile As String, _
    pstrExt As String, _
    pstrFiler As String, _
    flags As FileOpenConstants, _
    ByRef pbooCancel As Boolean, _
    Optional booOpen As Boolean = True) As String
    
    On Error Resume Next
    cd.DefaultExt = pstrExt
    cd.FileName = pstrDefaultFile
    cd.Filter = pstrFiler
    cd.flags = cdlOFNFileMustExist
    cd.CancelError = True
    
    If booOpen = True Then
        cd.ShowOpen
    Else
        cd.ShowSave
    End If
                              
    If Err.Number <> &H7FF3 Then ' cancel
        If Err.Number = 32755 Or cd.FileName = "" Then
            MsgBox "No file was selected", , gconstrTitlPrefix & "Update File Selection!"
            Exit Function
        End If
        
        SelectFile = cd.FileName
        pbooCancel = False
    Else
        'SelectFile = pstrDefaultFile
        pbooCancel = True
    End If
    
    DoEvents
    
End Function
Public Sub FindSetAnError( _
    Index As Integer, _
    fs As Object, _
    Text As String, _
    pobjPathType As PathType, _
    pstrAlternatePath As String, _
    ByRef hasError As Boolean, _
    ByRef errorText As String)

    Dim itemLength As Integer: itemLength = Len(Trim(Text & ""))
            
    If pobjPathType = FileWithPath Then
        If itemLength = 0 Or fs.FileExists(Text) = False Then
            hasError = True
            errorText = "This file can not be found"
        End If
    ElseIf pobjPathType = JustFile Then
        If itemLength = 0 Then
            hasError = True
            errorText = "You have not provided a file!"
        End If
        If hasError = False Then
            Dim fullPath As String: fullPath = pstrAlternatePath & "\" & Text
            If fs.FileExists(fullPath) = False Then
                hasError = True
                errorText = "This file as part of the server path has not been found"
            End If
        End If
    ElseIf pobjPathType = Path Then
        If itemLength = 0 Or fs.folderexists(Text) = False Then
            hasError = True
            errorText = "This path can not be found"
        ElseIf Right(Text, 1) <> "\" Then
            hasError = True
            errorText = "This path does not end with a back slash"
        End If
    End If
    
End Sub
