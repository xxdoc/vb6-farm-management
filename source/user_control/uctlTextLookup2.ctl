VERSION 5.00
Begin VB.UserControl uctlTextLookup2 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ScaleHeight     =   945
   ScaleWidth      =   6795
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3915
   End
   Begin VB.TextBox txtToCode 
      Height          =   435
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtFromCode 
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "uctlTextLookup2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Change()
Private m_ClearText As Boolean
Public Sub SetSelectText1(Start As Long, L As Long)
   txtFromCode.SelStart = Start
   txtFromCode.SelLength = L
End Sub
Public Sub SetSelectText2(Start As Long, L As Long)
   txtToCode.SelStart = Start
   txtToCode.SelLength = L
End Sub
Public Property Get Text1() As String
   Text1 = txtFromCode.Text
End Property
Public Property Get Text2() As String
   Text2 = txtToCode.Text
End Property
Public Property Let Text1(S As String)
   txtFromCode.Text = S
End Property
Public Property Let Text2(S As String)
   txtToCode.Text = S
End Property
Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(S As Boolean)
   UserControl.Enabled = S
   Call SetEnableDisableTextBox(txtFromCode, S)
   Call SetEnableDisableTextBox(txtToCode, S)
End Property
Private Sub cboName_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cboName_LostFocus()
   Call CloseCombo
End Sub

Private Sub txtFromCode_Change()
   RaiseEvent Change
End Sub
Private Sub txtToCode_Change()
   RaiseEvent Change
End Sub
Private Sub txtFromCode_GotFocus()
   Call SetSelect(txtFromCode)
End Sub
Private Sub txtToCode_GotFocus()
   Call SetSelect(txtToCode)
End Sub
Private Sub txtFromCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   ElseIf KeyAscii = 43 Then
      Call ShowCombo(txtFromCode)
   End If
End Sub
Private Sub txtToCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   ElseIf KeyAscii = 43 Then
      Call ShowCombo(txtToCode)
   End If
End Sub
Private Sub UserControl_Initialize()
   Call InitTextBox(txtFromCode, "")
   Call InitTextBox(txtToCode, "")
   Call InitCombo(cboName)
End Sub
Public Sub SetFocus()
   If txtFromCode.Visible Then
      txtFromCode.SetFocus
   End If
End Sub
Public Sub SetTextLenType(TT As TEXT_BOX_TYPE, L As Long)
   If TT = TEXT_FLOAT_MONEY Or TT = TEXT_INTEGER_MONEY Then
      txtFromCode.Alignment = 1
      txtToCode.Alignment = 1
   End If
   
   UserControl.Tag = TT
   txtFromCode.MaxLength = L
   txtToCode.MaxLength = L
End Sub
Public Sub ShowCombo(TBox As TextBox)
   cboName.Visible = True
   cboName.Top = TBox.Top
   cboName.Left = TBox.Left + TBox.Width
   UserControl.Width = TBox.Left + TBox.Width + cboName.Width
   Call cboName.SetFocus
End Sub
Public Sub CloseCombo()
   cboName.Visible = False
   UserControl.Width = txtFromCode.Width + txtToCode.Width
End Sub
Private Sub cboName_Click()
Dim O As Object
Dim TempID As Long

   RaiseEvent Change
   
   If cboName.ListIndex <= 0 Then
      If m_ClearText Then
         'txtCode.Text = ""
      End If
      Exit Sub
   End If
   
   TempID = cboName.ItemData(Minus2Zero(cboName.ListIndex))
   'Set O = MyCollection.Item(Trim(Str(TempID)))
      
   'txtCode.Text = O.KEY_LOOKUP
End Sub
