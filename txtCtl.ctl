VERSION 5.00
Begin VB.UserControl txtCtl 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   840
   ScaleWidth      =   2940
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   30
      Width           =   2235
   End
End
Attribute VB_Name = "txtCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum LetterCase
    Upper
    Lower
    Simple
End Enum
Private m_LetterCase As LetterCase
Private m_ForeColor As OLE_COLOR
Private m_Text As String

Private Sub Text1_Change()
    Dim pos As Integer
    pos = Text1.SelStart
    Select Case m_LetterCase
        Case Upper
            Text1.Text = VBA.UCase(Text1.Text)
        Case Lower
            Text1.Text = VBA.LCase(Text1.Text)
    End Select
    Text1.SelStart = pos
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_LetterCase = PropBag.ReadProperty("LetterCase", Simple)
    m_ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    m_Text = PropBag.ReadProperty("Text", "")
    Text1.Forecolor = m_ForeColor
    Call Text1_Change
End Sub

Private Sub UserControl_Resize()
    Text1.Left = 0
    Text1.Top = 0
    Text1.Height = UserControl.Height
    Text1.Width = UserControl.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("LetterCase", m_LetterCase, Simple)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, vbBlack)
    Call PropBag.WriteProperty("Text", m_Text, "")
End Sub

Public Property Get LetterCase() As LetterCase
        LetterCase = m_LetterCase
End Property

Public Property Let LetterCase(v_LetterCase As LetterCase)
    m_LetterCase = v_LetterCase
    Call Text1_Change
    PropertyChanged "LetterCase"
End Property

Public Property Get Forecolor() As OLE_COLOR
    Forecolor = m_ForeColor
End Property

Public Property Let Forecolor(v_ForeColor As OLE_COLOR)
    m_ForeColor = v_ForeColor
    Text1.Forecolor = m_ForeColor
    PropertyChanged "Forecolor"
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(v_Text As String)
    m_Text = v_Text
    PropertyChanged "Text"
End Property

