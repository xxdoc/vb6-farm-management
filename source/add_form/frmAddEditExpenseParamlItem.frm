VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExpenseParamItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditExpenseParamlItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   7646
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlBuyDate 
         Height          =   405
         Left            =   1710
         TabIndex        =   0
         Top             =   300
         Width           =   3825
         _extentx        =   6747
         _extenty        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1710
         TabIndex        =   4
         Top             =   1650
         Width           =   1935
         _extentx        =   9763
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlParttypeLookup 
         Height          =   405
         Left            =   1710
         TabIndex        =   2
         Top             =   750
         Width           =   5385
         _extentx        =   9499
         _extenty        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtADG 
         Height          =   435
         Left            =   1710
         TabIndex        =   5
         Top             =   2100
         Width           =   1935
         _extentx        =   9763
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgWeight 
         Height          =   435
         Left            =   1710
         TabIndex        =   6
         Top             =   2550
         Width           =   1935
         _extentx        =   9763
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExpenseName 
         Height          =   435
         Left            =   1710
         TabIndex        =   3
         Top             =   1200
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin Threed.SSCheck chkDepreciationFlag 
         Height          =   435
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblExpenseName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblAvgWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblBuyDate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblADG 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   780
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2040
         TabIndex        =   7
         Top             =   3330
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3690
         TabIndex        =   8
         Top             =   3330
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5340
         TabIndex        =   9
         Top             =   3330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1710
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditExpenseParamItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private m_PartTypes As Collection
Private m_PartItems As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboDrCr_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboDrCr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkDepreciationFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkDepreciationFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblBuyDate, MapText("�ѹ������"))
   Call InitNormalLabel(lblAmount, MapText("�ӹǹ"))
   Call InitNormalLabel(lblPartType, MapText("������ �.�.�."))
   Call InitNormalLabel(lblADG, MapText("�Ҥ�/˹���"))
   Call InitNormalLabel(lblExpenseName, MapText("��¨���"))
   Call InitNormalLabel(lblAvgWeight, MapText("�Ҥ����"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtADG.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtAvgWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtAvgWeight.Enabled = False
   Call txtExpenseName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitCheckBox(chkDepreciationFlag, "����������Ҥ����Դ� CASH FLOW")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdNext, MapText("�Ѵ�"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Ji As CParamItem
         
         Set Ji = TempCollection.Item(ID)

         uctlBuyDate.ShowDate = Ji.GetFieldValue("EXPENSE_DATE")
         txtAmount.Text = Ji.GetFieldValue("EXP_AMOUNT")
         txtADG.Text = Ji.GetFieldValue("EXP_AVG_PRICE")
         txtAvgWeight.Text = Ji.GetFieldValue("EXP_TOTAL_PRICE")
         txtExpenseName.Text = Ji.GetFieldValue("EXPENSE_NAME")
         uctlParttypeLookup.MyCombo.ListIndex = IDToListIndex(uctlParttypeLookup.MyCombo, Ji.GetFieldValue("EXPENSE_TYPE"))
         chkDepreciationFlag.Value = FlagToCheck(Ji.GetFieldValue("DEPRECIATION_FLAG"))
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If

   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "�֧�ä�����ش��������"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid(True)
         Exit Sub
      End If

      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      uctlBuyDate.ShowDate = -1
      uctlParttypeLookup.MyCombo.ListIndex = -1
      txtAmount.Text = ""
      txtADG.Text = ""
      txtAvgWeight.Text = ""
   End If
   Call QueryData(True)
   Call uctlBuyDate.SetFocus
   Call ParentForm.RefreshGrid(True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ji As CParamItem
   If ShowMode = SHOW_ADD Then
      Set Ji = New CParamItem
      Ji.Flag = "A"
      Call TempCollection.Add(Ji)
   Else
      Set Ji = TempCollection.Item(ID)
      If Ji.Flag <> "A" Then
         Ji.Flag = "E"
      End If
   End If

   Call Ji.SetFieldValue("EXP_AMOUNT", Val(txtAmount.Text))
   Call Ji.SetFieldValue("EXP_AVG_PRICE", Val(txtADG.Text))
   Call Ji.SetFieldValue("EXP_TOTAL_PRICE", Val(txtAvgWeight.Text))
   Call Ji.SetFieldValue("EXPENSE_DATE", uctlBuyDate.ShowDate)
   Call Ji.SetFieldValue("EXPENSE_TYPE", uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex)))
   Call Ji.SetFieldValue("EXPENSE_TYPE_NAME", uctlParttypeLookup.MyCombo.Text)
   Call Ji.SetFieldValue("EXPENSE_NAME", txtExpenseName.Text)
   Call Ji.SetFieldValue("DEPRECIATION_FLAG", Check2Flag(chkDepreciationFlag.Value))
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadExpenseType(uctlParttypeLookup.MyCombo, m_PartTypes)
      Set uctlParttypeLookup.MyCollection = m_PartTypes
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txtADG_Change()
   m_HasModify = True
   txtAvgWeight.Text = Val(txtAmount.Text) * Val(txtADG.Text)
End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
   txtAvgWeight.Text = Val(txtAmount.Text) * Val(txtADG.Text)
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtMoo_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtVillage_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccountLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtExpenseName_Change()
   m_HasModify = True
End Sub

Private Sub uctlBuyDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlParttypeLookup_Change()
   m_HasModify = True
End Sub
