VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ctlPathPicker 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   ForeColor       =   &H8000000C&
   ScaleHeight     =   990
   ScaleWidth      =   8970
   Begin VB.TextBox txtTextDesc 
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   540
      Width           =   1815
   End
   Begin VB.TextBox txtTextParam 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   540
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   10
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnChoose 
      Caption         =   "Choose"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   9000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblHelpDesc 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7080
      TabIndex        =   12
      Top             =   540
      Width           =   255
   End
   Begin VB.Label lblHelpParam 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      Top             =   540
      Width           =   255
   End
   Begin VB.Label lblLabelDesc 
      Caption         =   "Description"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   610
      Width           =   855
   End
   Begin VB.Label lblErrorDesc 
      BackColor       =   &H000000FF&
      Height          =   40
      Left            =   5160
      TabIndex        =   9
      Top             =   910
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabelParam 
      Caption         =   "Param"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   610
      Width           =   735
   End
   Begin VB.Label lblErrorParam 
      BackColor       =   &H000000FF&
      Height          =   40
      Left            =   2760
      TabIndex        =   6
      Top             =   910
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblHelp 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7080
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   60
      Width           =   255
   End
   Begin VB.Label lblError 
      BackColor       =   &H000000FF&
      Height          =   40
      Left            =   2760
      TabIndex        =   3
      Top             =   430
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   130
      Width           =   1935
   End
End
Attribute VB_Name = "ctlPathPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const cPATH_TYPE As String = "PathType"
Const cFILE_TYPE As String = "FileExtType"
Const cTEXT As String = "Text"
Const cLABEL As String = "Label"
Const cHELP As String = "Help"
Const cHAS_PARM_DESC As String = "HasParamDesc"
Const cHASERROR As String = "HasError"

Const cTEXT_PARAM As String = "TextParam"
Const cTEXT_DESC As String = "TextDesc"
Const cHELP_PARAM As String = "HelpParam"
Const cHELP_DESC As String = "HelpDesc"
Const cHASERROR_PARAM As String = "HasErrorParam"
Const cHASERROR_DESC As String = "HasErrorDesc"
Const cPARENT As String = "Parent"

Private mobjPathType As PathType
Private mobjFileExtType As FileExtType
Private mstrText As String
Private mstrLabel As String
Private mbooHasError As Boolean
Private mstrHelpTooltip As String
Private mstrHelpTooltipKept As String
Private mbooHasParamDesc As Boolean
Private mobjParent As Form

Private mstrTextParam As String
Private mstrTextDesc As String
Private mstrHelpTooltipParam As String
Private mstrHelpTooltipParamKept As String
Private mstrHelpTooltipDesc As String
Private mstrHelpTooltipDescKept As String
Private mbooHasErrorParam As Boolean
Private mbooHasErrorDesc As Boolean

Public Event OnChange(ByVal Text As String, ByVal pobjPathType As PathType)
Public Event OnFileSelected(ByVal Text As String, ByVal pobjPathType As PathType)
Public Event OnHelpHover(ByVal HelpText As String, Left As Integer, Top As Integer)
Public Event onHelpHoverOff()

Public Property Get ParentObj() As Form

    Set ParentObj = mobjParent
    
End Property
Public Property Let ParentObj(pobjParent As Form)

    Set mobjParent = pobjParent
    PropertyChanged cPARENT
    
End Property
Public Property Get HasParamDesc() As Boolean

    HasParamDesc = mbooHasParamDesc
    
End Property
Public Property Let HasParamDesc(pbooHasParamDesc As Boolean)
    
    mbooHasParamDesc = pbooHasParamDesc
    PropertyChanged cHAS_PARM_DESC
    
    SetAddtionalFields
    
End Property
Public Property Get ItemPathType() As PathType

    ItemPathType = mobjPathType

End Property
Public Property Let ItemPathType(pobjPathType As PathType)
    
    mobjPathType = pobjPathType
    PropertyChanged cPATH_TYPE

End Property
Public Property Get FileExtType() As FileExtType
    
    FileExtType = mobjFileExtType

End Property
Public Property Let FileExtType(pobjFileExtType As FileExtType)
    
    mobjFileExtType = pobjFileExtType
    PropertyChanged cFILE_TYPE

End Property
Public Property Get Text() As String
    
    Text = mstrText

End Property
Public Property Let Text(pstrText As String)
    
    mstrText = pstrText
    PropertyChanged cTEXT
    txtText.Text = pstrText

End Property
Public Property Get TextParam() As String
    
    TextParam = mstrTextParam

End Property
Public Property Let TextParam(pstrTextParam As String)
    
    mstrTextParam = pstrTextParam
    PropertyChanged cTEXT_PARAM
    txtTextParam.Text = pstrTextParam

End Property
Public Property Get TextDesc() As String
    
    TextDesc = mstrTextDesc

End Property
Public Property Let TextDesc(pstrTextDesc As String)
    
    mstrTextDesc = pstrTextDesc
    PropertyChanged cTEXT_DESC
    txtTextDesc.Text = pstrTextDesc

End Property
Public Property Get Label() As String
    
    Label = mstrLabel

End Property
Public Property Let Label(pstrLabel As String)
    
    mstrLabel = pstrLabel
    PropertyChanged cLABEL
    lblLabel.Caption = pstrLabel
    
End Property
Public Property Get hasError() As Boolean
    
    hasError = mbooHasError

End Property
Public Property Let hasError(pbooValue As Boolean)
    
    mbooHasError = pbooValue
    PropertyChanged cHASERROR

End Property
Public Property Get hasErrorParam() As Boolean
    
    hasErrorParam = mbooHasErrorParam

End Property
Public Property Let hasErrorParam(pbooValueParam As Boolean)
    
    mbooHasErrorParam = pbooValueParam
    PropertyChanged cHASERROR_PARAM

End Property
Public Property Get hasErrorDesc() As Boolean
    
    hasErrorDesc = mbooHasErrorDesc

End Property
Public Property Let hasErrorDesc(pbooValueDesc As Boolean)
    
    mbooHasErrorDesc = pbooValueDesc
    PropertyChanged cHASERROR_DESC

End Property
Public Property Get HelpTooltip() As String
    
    HelpTooltip = mstrHelpTooltip

End Property
Public Property Let HelpTooltip(pstrHelpTooltip As String)
    
    mstrHelpTooltip = pstrHelpTooltip
    mstrHelpTooltipKept = mstrHelpTooltip
    PropertyChanged cHELP
    
End Property
Public Property Get HelpTooltipParam() As String
    
    HelpTooltipParam = mstrHelpTooltipParam

End Property
Public Property Let HelpTooltipParam(pstrHelpTooltipParam As String)
    
    mstrHelpTooltipParam = pstrHelpTooltipParam
    mstrHelpTooltipParamKept = mstrHelpTooltipParam
    PropertyChanged cHELP_PARAM
    
End Property
Public Property Get HelpTooltipDesc() As String
    
    HelpTooltipDesc = mstrHelpTooltipDesc

End Property
Public Property Let HelpTooltipDesc(pstrHelpTooltipDesc As String)
    
    mstrHelpTooltipDesc = pstrHelpTooltipDesc
    mstrHelpTooltipDescKept = mstrHelpTooltipDesc
    PropertyChanged cHELP_DESC
    
End Property
Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent OnHelpHover(mstrHelpTooltip, lblHelp.Left, lblHelp.Top)
    
End Sub
Private Sub lblHelpDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent OnHelpHover(mstrHelpTooltipDesc, lblHelpDesc.Left, lblHelpDesc.Top)

End Sub
Private Sub lblHelpParam_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent OnHelpHover(mstrHelpTooltipParam, lblHelpParam.Left, lblHelpParam.Top)
    
End Sub
Private Sub txtText_Change()

    mstrText = txtText.Text
    RaiseEvent OnChange(txtText.Text, mobjPathType)

End Sub

Private Sub txtText_GotFocus()

    On Error Resume Next
    
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.Text)
        
    On Error GoTo 0
    
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 65 And Shift = 2 Then
        txtText.SelStart = 0
        txtText.SelLength = Len(txtText.Text)
    End If
    
End Sub

Private Sub txtText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.Text)
        
    On Error GoTo 0
    
End Sub

Private Sub txtTextDesc_GotFocus()

    On Error Resume Next
    
    txtTextDesc.SelStart = 0
    txtTextDesc.SelLength = Len(txtTextDesc.Text)
        
    On Error GoTo 0
    
End Sub

Private Sub txtTextDesc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    txtTextDesc.SelStart = 0
    txtTextDesc.SelLength = Len(txtTextDesc.Text)
        
    On Error GoTo 0
    
End Sub

Private Sub txtTextParam_GotFocus()

    On Error Resume Next
    
    txtTextParam.SelStart = 0
    txtTextParam.SelLength = Len(txtTextParam.Text)
        
    On Error GoTo 0
    
End Sub

Private Sub txtTextParam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    txtTextParam.SelStart = 0
    txtTextParam.SelLength = Len(txtTextParam.Text)
        
    On Error GoTo 0
    
End Sub
Private Sub txtTextParam_Change()

    If mbooHasParamDesc Then
        mstrTextParam = txtTextParam.Text
        If Len(Trim$(txtTextParam.Text & "")) = 0 Then
            mbooHasErrorParam = True
            lblHelpParam.ForeColor = vbRed
        Else
            mbooHasErrorParam = False
            lblHelpParam.ForeColor = vbBlue
        End If
        
        lblErrorParam.Visible = mbooHasErrorParam
    Else
        lblHelpParam.ForeColor = vbBlue
        lblErrorParam.Visible = False
    End If
    
End Sub
Private Sub txtTextDesc_Change()

    If mbooHasParamDesc Then
        mstrTextDesc = txtTextDesc.Text
        If Len(Trim$(txtTextDesc.Text & "")) = 0 Then
            mbooHasErrorDesc = True
            lblHelpDesc.ForeColor = vbRed
        Else
            mbooHasErrorDesc = False
            lblHelpDesc.ForeColor = vbBlue
        End If
        
        lblErrorDesc.Visible = mbooHasErrorDesc
    Else
        lblErrorDesc.Visible = False
        lblHelpDesc.ForeColor = vbBlue
    End If
    
End Sub
Public Sub SetError(pbooHasError As Boolean, pstrErrorText As String)

    mbooHasError = pbooHasError
    lblError.Visible = pbooHasError
    If hasError = True Then
        lblHelp.ForeColor = vbRed
        mstrHelpTooltip = mstrHelpTooltipKept & vbCrLf & " " & vbCrLf & pstrErrorText
    Else
        lblHelp.ForeColor = vbBlue
        mstrHelpTooltip = mstrHelpTooltipKept
    End If
        
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent onHelpHoverOff
    
End Sub
Private Sub UserControl_ReadProperties(pobjPropBag As PropertyBag)
    
    mobjPathType = pobjPropBag.ReadProperty(cPATH_TYPE, JustFile)
    mstrText = pobjPropBag.ReadProperty(cTEXT, "")
    mstrLabel = pobjPropBag.ReadProperty(cLABEL, "")
    mstrHelpTooltip = pobjPropBag.ReadProperty(cHELP, "")
    mbooHasParamDesc = pobjPropBag.ReadProperty(cHAS_PARM_DESC, False)
    mstrTextParam = pobjPropBag.ReadProperty(cTEXT_PARAM, "")
    mstrTextDesc = pobjPropBag.ReadProperty(cTEXT_DESC, "")
'    Set mobjParent = pobjPropBag.ReadProperty(cPARENT, Null)
    
    txtText.Text = mstrText
    txtTextParam.Text = mstrTextParam
    txtTextDesc.Text = mstrTextDesc
    
    lblLabel = mstrLabel
    lblHelp.ToolTipText = ""
    mbooHasError = True

    SetAddtionalFields

End Sub
Private Sub UserControl_WriteProperties(pobjPropBag As PropertyBag)

    Call pobjPropBag.WriteProperty(cPATH_TYPE, mobjPathType, JustFile)
    Call pobjPropBag.WriteProperty(cTEXT, mstrText, "")
    Call pobjPropBag.WriteProperty(cLABEL, mstrLabel, "")
    Call pobjPropBag.WriteProperty(cHELP, mstrHelpTooltip, "")
    Call pobjPropBag.WriteProperty(cHAS_PARM_DESC, False)
    Call pobjPropBag.WriteProperty(cTEXT_PARAM, "")
    Call pobjPropBag.WriteProperty(cTEXT_DESC, "")
    'Call pobjPropBag.WriteProperty(cPARENT, Null)
    
End Sub
Private Sub btnChoose_Click()

    Dim lstrFile As String
    Dim lbooCommonDialogCancelled As Boolean
        
    If mobjPathType = Path Then
    
        Dim lstrPath As String: lstrPath = ""
        lstrPath = GetNetDir(mobjParent, CSIDL_NETWORK)
        If lstrPath <> "" Then
            txtText.Text = lstrPath
            RaiseEvent OnFileSelected(txtText.Text, mobjPathType)
        End If
    Else
        Dim lstrExt As String: lstrExt = ""
        Dim lstrFilter As String: lstrFilter = ""
    
        Select Case mobjFileExtType
        Case mdb
            lstrExt = "mdb"
            lstrFilter = "MS Access Database File (*.mdb)|*.mdb"
        Case txt
            lstrExt = "txt"
            lstrFilter = "Text File (*.txt)|*.txt"
        Case exe
            lstrExt = "exe"
            lstrFilter = "Application (*.exe)|*.exe"
        End Select
        
        Dim lobjFlags As FileOpenConstants
        Dim lbooShowSave As Boolean
        
        If mobjPathType = FileWithPath Then
            lobjFlags = cdlOFNCreatePrompt
            lbooShowSave = False
        Else
            lobjFlags = cdlOFNFileMustExist
            lbooShowSave = True
        End If
        
        lstrFile = SelectFile(CommonDialog1, txtText.Text, lstrExt, lstrFilter, lobjFlags, lbooCommonDialogCancelled, lbooShowSave)
        
        If lbooCommonDialogCancelled = False And lbooShowSave = False And Trim$(Dir$(lstrFile)) = "" Then
            Dim lintFileNum As Integer: lintFileNum = FreeFile
            Open lstrFile For Output As #lintFileNum
            Close #lintFileNum
        End If
        
        txtText.Text = lstrFile
    End If
    
    If lbooCommonDialogCancelled = False Then
        RaiseEvent OnFileSelected(txtText.Text, mobjPathType)
    End If
    
End Sub
Private Sub SetAddtionalFields()

    lblErrorDesc.Visible = mbooHasParamDesc
    lblErrorParam.Visible = mbooHasParamDesc
    txtTextDesc.Visible = mbooHasParamDesc
    txtTextParam.Visible = mbooHasParamDesc
    lblLabelDesc.Visible = mbooHasParamDesc
    lblLabelParam.Visible = mbooHasParamDesc
    lblHelpDesc.Visible = mbooHasParamDesc
    lblHelpParam.Visible = mbooHasParamDesc

End Sub

