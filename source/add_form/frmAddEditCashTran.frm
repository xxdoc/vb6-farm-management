VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCashTran 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmAddEditCashTran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   11880
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPaymentType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1020
         Width           =   3495
      End
      Begin prjFarmManagement.uctlTextLookup uctlChequeType 
         Height          =   405
         Left            =   1860
         TabIndex        =   4
         Top             =   2820
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlChequeDate 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtChequeNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtChequeAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   4620
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlEffectiveDate 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   2370
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlBank 
         Height          =   405
         Left            =   1860
         TabIndex        =   5
         Top             =   3270
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankBranch 
         Height          =   405
         Left            =   1860
         TabIndex        =   6
         Top             =   3720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlBankAccountLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   7
         Top             =   4170
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFeeAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   5040
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin VB.Label lblFeeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   26
         Top             =   5100
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4080
         TabIndex        =   25
         Top             =   5100
         Width           =   1575
      End
      Begin VB.Label lblBankAccount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   24
         Top             =   4290
         Width           =   1575
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2205
         TabIndex        =   10
         Top             =   5760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashTran.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblPaymentType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   23
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4050
         TabIndex        =   22
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   21
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   20
         Top             =   3390
         Width           =   1575
      End
      Begin VB.Label lblChequeType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   19
         Top             =   2940
         Width           =   1575
      End
      Begin VB.Label lblEffectiveDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblChequeDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   17
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblChequeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblChequeNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   1530
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5505
         TabIndex        =   12
         Top             =   5760
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3855
         TabIndex        =   11
         Top             =   5760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashTran.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCashTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Cheque As CCheque

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ChequeType As Long
Public ParentForm As Object
Public TempCollection As Collection
Public Area As Long

Private Mr As CMasterRef
Private m_ChequeTypes As Collection
Private m_ApAr As Collection
Private m_Banks As Collection
Private m_BankBranchs As Collection
Private m_BankAccounts As Collection
Private m_ApArMas As CCustomer

Private Sub cmdPasswd_Click()

End Sub

Private Sub cboUserGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboPaymentType_Click()
Dim TempID As Long
   
   m_HasModify = True
   
   TempID = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   If TempID = CASH_PMT Or TempID = CASHRET_PMT Then
      txtChequeNo.Enabled = False
      uctlChequeDate.Enable = False
      uctlEffectiveDate.Enable = False
      uctlChequeType.Enabled = False
      uctlBank.Enabled = False
      uctlBankBranch.Enabled = False
      uctlBankAccountLookup.Enabled = False
      txtChequeAmount.Enabled = True
      txtFeeAmount.Enabled = False
   ElseIf TempID = BANKTRF_PMT Then
      txtChequeNo.Enabled = False
      uctlChequeDate.Enable = False
      uctlEffectiveDate.Enable = False
      uctlChequeType.Enabled = False
      uctlBank.Enabled = True
      uctlBankBranch.Enabled = True
      uctlBankAccountLookup.Enabled = True
      txtChequeAmount.Enabled = True
      txtFeeAmount.Enabled = True
   ElseIf TempID = CHECK_PMT Then
      txtChequeNo.Enabled = True
      uctlChequeDate.Enable = True
      uctlEffectiveDate.Enable = True
      uctlChequeType.Enabled = True
      uctlBank.Enabled = True
      If Area = 1 Then
         uctlBankAccountLookup.Enabled = False '���Ѻ����ͧ�к���ش�ѭ����з�� pay in ����ѧ
      ElseIf Area = 2 Then
         uctlBankAccountLookup.Enabled = True '�礨��µ�ͧ�к���ش�ѭ����ҵѴ�ҡ�ѭ���
      End If
      uctlBankBranch.Enabled = True
      txtChequeAmount.Enabled = True
      txtFeeAmount.Enabled = False
   Else
      txtChequeNo.Enabled = False
      uctlChequeDate.Enable = False
      uctlEffectiveDate.Enable = False
      uctlChequeType.Enabled = False
      uctlBank.Enabled = False
      uctlBankBranch.Enabled = False
      txtChequeAmount.Enabled = False
   End If
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
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      cboPaymentType.ListIndex = -1
      txtChequeNo.Text = ""
      uctlChequeDate.ShowDate = -1
      uctlEffectiveDate.ShowDate = -1
      uctlChequeType.MyCombo.ListIndex = -1
      uctlBank.MyCombo.ListIndex = -1
      uctlBankBranch.MyCombo.ListIndex = -1
      uctlBankAccountLookup.MyCombo.ListIndex = -1
      txtChequeAmount.Text = ""
   End If
   
   Call ParentForm.RefreshGrid
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me

End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim PaymentType As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Dim Ji As CCashTran
      Set Ji = TempCollection.Item(ID)
      
      PaymentType = Ji.GetFieldValue("PAYMENT_TYPE")
      cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, PaymentType)
      txtChequeNo.Text = Ji.Cheque.GetFieldValue("CHEQUE_NO")
      txtChequeAmount.Text = Ji.GetFieldValue("AMOUNT")
      txtFeeAmount.Text = Ji.GetFieldValue("FEE_AMOUNT")
      uctlChequeDate.ShowDate = Ji.Cheque.GetFieldValue("CHEQUE_DATE")
      uctlEffectiveDate.ShowDate = Ji.Cheque.GetFieldValue("EFFECTIVE_DATE")
      uctlChequeType.MyCombo.ListIndex = IDToListIndex(uctlChequeType.MyCombo, Ji.Cheque.GetFieldValue("CHEQUE_TYPE"))
      If PaymentType = 2 Then
         uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, Ji.GetFieldValue("BANK_ID"))
         uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, Ji.GetFieldValue("BANK_BRANCH"))
      ElseIf PaymentType = 3 Then
         uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, Ji.Cheque.GetFieldValue("BANK_ID"))
         uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, Ji.Cheque.GetFieldValue("BANK_BRANCH"))
      ElseIf PaymentType = 1 Then
         uctlBank.MyCombo.ListIndex = -1
         uctlBankBranch.MyCombo.ListIndex = -1
      End If
      uctlBankAccountLookup.MyCombo.ListIndex = IDToListIndex(uctlBankAccountLookup.MyCombo, Ji.GetFieldValue("BANK_ACCOUNT"))
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim PaymentType As Long

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("LEDGER_CASH_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("LEDGER_CASH_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
   
   If Not VerifyTextControl(lblChequeNo, txtChequeNo, Not txtChequeNo.Enabled) Then
      Exit Function
   End If
   If Not VerifyDate(lblChequeDate, uctlChequeDate, Not txtChequeNo.Enabled) Then
      Exit Function
   End If
   If Not VerifyDate(lblEffectiveDate, uctlEffectiveDate, Not txtChequeNo.Enabled) Then
      Exit Function
   End If
'   If Not VerifyCombo(lblChequeType, uctlChequeType.MyCombo, Not uctlChequeType.Enabled) Then
'      Exit Function
'   End If
   If Not VerifyCombo(lblBank, uctlBank.MyCombo, Not uctlBank.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankBranch, uctlBankBranch.MyCombo, Not uctlBankBranch.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankAccount, uctlBankAccountLookup.MyCombo, Not uctlBankAccountLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblChequeAmount, txtChequeAmount, Not txtChequeAmount.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPaymentType, cboPaymentType, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(USERNAME_UNIQUE, txtChequeNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtChequeNo.Text & " " & MapText("������к�����")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ji As CCashTran
   If ShowMode = SHOW_ADD Then
      Set Ji = New CCashTran
      Ji.Flag = "A"
      Call TempCollection.Add(Ji)
   Else
      Set Ji = TempCollection.Item(ID)
      If Ji.Flag <> "A" Then
         Ji.Flag = "E"
      End If
   End If
   
   PaymentType = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   Call Ji.SetFieldValue("PAYMENT_TYPE", PaymentType)
   Call Ji.SetFieldValue("PAYMENT_TYPE_NAME", PaymentTypeToText(cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))))
   Call Ji.SetFieldValue("AMOUNT", Val(txtChequeAmount.Text))
   Call Ji.SetFieldValue("FEE_AMOUNT", Val(txtFeeAmount.Text))
   Call Ji.SetFieldValue("NET_AMOUNT", Val(txtChequeAmount.Text) - Val(txtFeeAmount.Text))
   If PaymentType = CASH_PMT Then
      Call Ji.SetFieldValue("BANK_ID", -1)
      Call Ji.SetFieldValue("BANK_BRANCH", -1)
      Call Ji.SetFieldValue("BANK_ACCOUNT", -1)
   ElseIf PaymentType = BANKTRF_PMT Then
      Call Ji.SetFieldValue("BANK_ID", uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex)))
      Call Ji.SetFieldValue("BANK_BRANCH", uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex)))
      Call Ji.SetFieldValue("BANK_ACCOUNT", uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex)))
      Call Ji.SetFieldValue("BANK_NAME", uctlBank.MyCombo.Text)
      Call Ji.SetFieldValue("BRANCH_NAME", uctlBankBranch.MyCombo.Text)
      Call Ji.SetFieldValue("ACCOUNT_NAME", uctlBankAccountLookup.MyCombo.Text)
   ElseIf PaymentType = CHECK_PMT Then
      Call Ji.Cheque.SetFieldValue("BANK_ID", uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex)))
      Call Ji.Cheque.SetFieldValue("BANK_BRANCH", uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex)))
      Call Ji.Cheque.SetFieldValue("BANK_NAME", uctlBank.MyCombo.Text)
      Call Ji.Cheque.SetFieldValue("BRANCH_NAME", uctlBankBranch.MyCombo.Text)
      'Call Ji.SetFieldValue("BANK_ACCOUNT", uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex)))
      If Area = 2 Then '�礨���
         Call Ji.SetFieldValue("BANK_ID", uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex)))
         Call Ji.SetFieldValue("BANK_BRANCH", uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex)))
         Call Ji.SetFieldValue("BANK_ACCOUNT", uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex)))
      End If
   End If
   
   If Area = 1 Then
      Call Ji.SetFieldValue("TX_TYPE", "I")
      Call Ji.Cheque.SetFieldValue("DIRECTION", 1)
   ElseIf Area = 2 Then
      Call Ji.SetFieldValue("TX_TYPE", "E")
      Call Ji.Cheque.SetFieldValue("DIRECTION", 2)
   End If
   
   Call Ji.Cheque.SetFieldValue("CHEQUE_NO", txtChequeNo.Text)
   Call Ji.Cheque.SetFieldValue("CHEQUE_AMOUNT", Val(txtChequeAmount.Text))
   Call Ji.Cheque.SetFieldValue("CHEQUE_DATE", uctlChequeDate.ShowDate)
   Call Ji.Cheque.SetFieldValue("EFFECTIVE_DATE", uctlEffectiveDate.ShowDate)
   Call Ji.Cheque.SetFieldValue("CHEQUE_TYPE", uctlChequeType.MyCombo.ItemData(Minus2Zero(uctlChequeType.MyCombo.ListIndex)))
   Call Ji.Cheque.SetFieldValue("APAR_ID", -1)
   Call Ji.Cheque.SetFieldValue("CHEQUE_STATUS", 1)
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(uctlChequeType.MyCombo, m_ChequeTypes, CHEQUE_TYPE)
      Set uctlChequeType.MyCollection = m_ChequeTypes
      
      Call LoadBank(uctlBank.MyCombo, m_Banks)
      Set uctlBank.MyCollection = m_Banks
      
      Call LoadBankBranch(uctlBankBranch.MyCombo, m_BankBranchs)
      Set uctlBankBranch.MyCollection = m_BankBranchs
                  
      Call LoadBankAccount(uctlBankAccountLookup.MyCombo, m_BankAccounts)
      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
      
      Call InitPaymentTypeEx(cboPaymentType)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblChequeNo, MapText("�Ţ�����"))
   Call InitNormalLabel(lblChequeAmount, MapText("�ӹǹ�Թ"))
   Call InitNormalLabel(lblChequeDate, MapText("�ѹ�����"))
   Call InitNormalLabel(lblEffectiveDate, MapText("�ѹ������Թ"))
   Call InitNormalLabel(lblChequeType, MapText("��������"))
   Call InitNormalLabel(lblBank, MapText("��Ҥ��"))
   Call InitNormalLabel(lblBankBranch, MapText("�ҢҸ�Ҥ��"))
   Call InitNormalLabel(Label1, MapText("�ҷ"))
   Call InitNormalLabel(Label2, MapText("�ҷ"))
   Call InitNormalLabel(lblPaymentType, MapText("��ê����Թ"))
   Call InitNormalLabel(lblBankAccount, MapText("�Ţ���ѭ��"))
   Call InitNormalLabel(lblFeeAmount, MapText("��Ҹ�������"))
   
   Call txtChequeNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtChequeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboPaymentType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdNext, MapText("�Ѵ�"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Cheque = New CCheque
   Set m_Rs = New ADODB.Recordset
   Set m_Cheque = New CCheque
   Set Mr = New CMasterRef
   
   Set m_ChequeTypes = New Collection
   Set m_ApAr = New Collection
   Set m_Banks = New Collection
   Set m_BankBranchs = New Collection
   Set m_ApArMas = New CCustomer
   Set m_BankAccounts = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Cheque = Nothing
   Set Mr = Nothing
   
   Set m_ChequeTypes = Nothing
   Set m_ApAr = Nothing
   Set m_Banks = Nothing
   Set m_BankBranchs = Nothing
   Set m_ApArMas = Nothing
   Set m_BankAccounts = Nothing
End Sub

Private Sub txtChequeAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtUserDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtChequeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub uctlAPAR_Change()
   m_HasModify = True
End Sub

Private Sub txtFeeAmount_Change()
   m_HasModify = True
End Sub
Private Sub uctlBank_Change()
Dim TempID As Long
Dim BB As CBankBranch
   TempID = uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex))
   
   If TempID > 0 Then
      Call LoadBankBranch(uctlBankBranch.MyCombo, m_BankBranchs, TempID)
      Set uctlBankBranch.MyCollection = m_BankBranchs
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlBankAccountLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlBankBranch_Change()
Dim TempID1 As Long
Dim TempID2 As Long
   
   TempID1 = uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex))
   TempID2 = uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex))
   
   If TempID2 > 0 Then
      Call LoadBankAccount(uctlBankAccountLookup.MyCombo, m_BankAccounts, TempID1, TempID2)
      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlChequeDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlChequeType_Change()
   m_HasModify = True
End Sub

Private Sub uctlEffectiveDate_HasChange()
   m_HasModify = True
End Sub
