VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditCustomer 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboArea 
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
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3330
         Width           =   2265
      End
      Begin VB.ComboBox cboEnterpriseType 
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
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2880
         Width           =   2265
      End
      Begin VB.ComboBox cboBusinessType 
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
         TabIndex        =   7
         Top             =   2880
         Width           =   3495
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   11
         Top             =   4170
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1530
         Width           =   3525
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtShortName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtEmail 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   1980
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWebSite 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   2430
         Width           =   9225
         _ExtentX        =   16960
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBusinessDesc 
         Height          =   450
         Left            =   1860
         TabIndex        =   9
         Top             =   3330
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   794
      End
      Begin prjFarmManagement.uctlTextBox txtCredit 
         Height          =   435
         Left            =   5700
         TabIndex        =   1
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   8070
         TabIndex        =   2
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3015
         Left            =   150
         TabIndex        =   12
         Top             =   4710
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5318
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditCustomer.frx":27A2
         Column(2)       =   "frmAddEditCustomer.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditCustomer.frx":290E
         FormatStyle(2)  =   "frmAddEditCustomer.frx":2A6A
         FormatStyle(3)  =   "frmAddEditCustomer.frx":2B1A
         FormatStyle(4)  =   "frmAddEditCustomer.frx":2BCE
         FormatStyle(5)  =   "frmAddEditCustomer.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditCustomer.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtCreditAmount 
         Height          =   435
         Left            =   9870
         TabIndex        =   4
         Top             =   1980
         Width           =   1125
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTaxID 
         Height          =   435
         Left            =   6240
         TabIndex        =   36
         Top             =   1530
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCustomerBranch 
         Height          =   435
         Left            =   9600
         TabIndex        =   37
         Top             =   1530
         Width           =   1635
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin VB.Label lblCustomerBranch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8760
         TabIndex        =   38
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblTaxID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5400
         TabIndex        =   35
         Top             =   1560
         Width           =   735
      End
      Begin Threed.SSCheck chkKeyAccount 
         Height          =   435
         Left            =   9240
         TabIndex        =   34
         Top             =   3840
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkTakeAwayFlag 
         Height          =   435
         Left            =   9240
         TabIndex        =   33
         Top             =   1080
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblArea 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7230
         TabIndex        =   32
         Top             =   3390
         Width           =   1485
      End
      Begin VB.Label lblCreditAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8910
         TabIndex        =   31
         Top             =   2070
         Width           =   855
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   11100
         TabIndex        =   30
         Top             =   2070
         Width           =   495
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   17
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   14
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblDiscountPercent 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6900
         TabIndex        =   28
         Top             =   1140
         Width           =   1065
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   6570
         TabIndex        =   27
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   26
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblBusinessDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   25
         Top             =   3450
         Width           =   1695
      End
      Begin VB.Label lblWebsite 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   23
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   22
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label lblEnterpriseType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7230
         TabIndex        =   21
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label lblBusinessType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   20
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   1620
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Customer As CCustomer

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Customer.CUSTOMER_ID = ID
      If Not glbDaily.QueryCustomer(m_Customer, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_Customer.PopulateFromRS(1, m_Rs)
      
      txtEmail.Text = m_Customer.EMAIL
      txtWebSite.Text = m_Customer.WEBSITE
      cboBusinessType.ListIndex = IDToListIndex(cboBusinessType, m_Customer.CUSTOMER_TYPE)
      cboEnterpriseType.ListIndex = IDToListIndex(cboEnterpriseType, m_Customer.CUSTOMER_GRADE)
      txtShortName.Text = m_Customer.CUSTOMER_CODE
      txtBusinessDesc.Text = m_Customer.BUSINESS_DESC
      txtCredit.Text = m_Customer.Credit
      txtDiscountPercent.Text = m_Customer.NORMAL_DISCOUNT
      txtCreditAmount.Text = m_Customer.CREDIT_AMOUNT
      cboArea.ListIndex = IDToListIndex(cboArea, m_Customer.REGION_ID)
      chkTakeAwayFlag.Value = FlagToCheck(m_Customer.TAKE_AWAY_FLAG)
      chkKeyAccount.Value = FlagToCheck(m_Customer.TAKE_KEY_ACCOUNT)
      
      txtCustomerBranch.Text = m_Customer.CUSTOMER_BRANCH
      txtTaxID.Text = m_Customer.TAX_ID
      
      Dim Name As cName
      Dim CstName As CCustomerName
      If (Not m_Customer.CstNames Is Nothing) And (m_Customer.CstNames.Count > 0) Then
         Set CstName = m_Customer.CstNames(1)
         Set Name = CstName.Name
         txtName.Text = Name.LONG_NAME
      Else
         txtName.Text = ""
      End If
   Else
      ShowMode = SHOW_ADD
   End If
   
   If ShowMode = SHOW_ADD Then
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyAccessRight("MAIN_CUSTOMER_EDIT") Then
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not VerifyTextControl(lblShortName, txtShortName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBusinessType, cboBusinessType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblEnterpriseType, cboEnterpriseType, False) Then
      Exit Function
   End If

   If Not CheckUniqueNs(CUSTCODE_UNIQUE, txtShortName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtShortName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_Customer.AddEditMode = ShowMode
   m_Customer.BIRTH_DATE = -1
   m_Customer.CUSTOMER_PASSWORD = ""
   m_Customer.EMAIL = txtEmail.Text
   m_Customer.WEBSITE = txtWebSite.Text
   m_Customer.CUSTOMER_TYPE = cboBusinessType.ItemData(Minus2Zero(cboBusinessType.ListIndex))
   m_Customer.CUSTOMER_GRADE = cboEnterpriseType.ItemData(Minus2Zero(cboEnterpriseType.ListIndex))
   m_Customer.Credit = Val(txtCredit.Text)
   m_Customer.CUSTOMER_CODE = txtShortName.Text
   m_Customer.BUSINESS_DESC = txtBusinessDesc.Text
   m_Customer.NORMAL_DISCOUNT = Val(txtDiscountPercent.Text)
   m_Customer.RESPONSE_BY = -1
   m_Customer.CREDIT_AMOUNT = Val(txtCreditAmount.Text)
   m_Customer.REGION_ID = cboArea.ItemData(Minus2Zero(cboArea.ListIndex))
   
   m_Customer.TAKE_AWAY_FLAG = Check2Flag(chkTakeAwayFlag.Value)
   m_Customer.TAKE_KEY_ACCOUNT = Check2Flag(chkKeyAccount.Value)
   
   m_Customer.CUSTOMER_BRANCH = txtCustomerBranch.Text
   m_Customer.TAX_ID = txtTaxID.Text
   
   'Create Dummy account
   If m_Customer.CstAccounts.Count <= 0 Then
      Dim Acc As CAccount
      
      Set Acc = New CAccount
      
      Acc.ACCOUNT_NO = m_Customer.CUSTOMER_CODE
      Acc.Flag = "A"
      
      Call m_Customer.CstAccounts.Add(Acc)
      
      Set Acc = Nothing
   End If

   Dim CstName As CCustomerName
   If m_Customer.CstNames.Count <= 0 Then
      Set CstName = New CCustomerName
      CstName.Flag = "A"
      Call m_Customer.CstNames.Add(CstName)
   Else
      Set CstName = m_Customer.CstNames.Item(1)
      CstName.Flag = "E"
   End If
   
   Dim Name As cName
   If m_Customer.CstNames.Count <= 0 Then
      Set Name = CstName.Name
      Name.LONG_NAME = txtName.Text
      Name.SHORT_NAME = txtShortName.Text
      Name.Flag = "A"
   Else
      Set Name = CstName.Name
      Name.LONG_NAME = txtName.Text
      Name.SHORT_NAME = txtShortName.Text
      Name.Flag = "E"
   End If
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCustomer(m_Customer, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboArea_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub chkKeyAccount_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkTakeAwayFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkTakeAwayFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkKeyAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditCustomerAddress.TempCollection = m_Customer.CstAddr
      frmAddEditCustomerAddress.ShowMode = SHOW_ADD
      frmAddEditCustomerAddress.HeaderText = MapText("เพิ่มที่อยู่")
      Load frmAddEditCustomerAddress
      frmAddEditCustomerAddress.Show 1

      OKClick = frmAddEditCustomerAddress.OKClick

      Unload frmAddEditCustomerAddress
      Set frmAddEditCustomerAddress = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
       Set frmAddEditCustomerPackage.TempCollection = m_Customer.CstPackage
       Set frmAddEditCustomerPackage.ParentForm = Me
      frmAddEditCustomerPackage.ShowMode = SHOW_ADD
      frmAddEditCustomerPackage.HeaderText = MapText("เพิ่มแบบการตั้งราคาขาย")
      Load frmAddEditCustomerPackage
      frmAddEditCustomerPackage.Show 1

      OKClick = frmAddEditCustomerPackage.OKClick

      Unload frmAddEditCustomerPackage
      Set frmAddEditCustomerPackage = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstPackage)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_Customer.CstAddr.Remove (ID2)
      Else
         m_Customer.CstAddr.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_Customer.CstAddr)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_Customer.CstPackage.Remove (ID2)
      Else
         m_Customer.CstPackage.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_Customer.CstPackage)
      GridEX1.Rebind
      m_HasModify = True
   
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditCustomerAddress.ID = ID
      Set frmAddEditCustomerAddress.TempCollection = m_Customer.CstAddr
      frmAddEditCustomerAddress.HeaderText = MapText("แก้ไขที่อยู่")
      frmAddEditCustomerAddress.ShowMode = SHOW_EDIT
      Load frmAddEditCustomerAddress
      frmAddEditCustomerAddress.Show 1

      OKClick = frmAddEditCustomerAddress.OKClick

      Unload frmAddEditCustomerAddress
      Set frmAddEditCustomerAddress = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditCustomerPackage.ID = ID
      Set frmAddEditCustomerPackage.TempCollection = m_Customer.CstPackage
      Set frmAddEditCustomerPackage.ParentForm = Me
      frmAddEditCustomerPackage.HeaderText = MapText("แก้ไขรูปแบบการตั้งราคาขาย")
      frmAddEditCustomerPackage.ShowMode = SHOW_EDIT
      Load frmAddEditCustomerPackage
      frmAddEditCustomerPackage.Show 1

      OKClick = frmAddEditCustomerPackage.OKClick

      Unload frmAddEditCustomerPackage
      Set frmAddEditCustomerPackage = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstPackage)
         GridEX1.Rebind
      End If
   
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()

   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadCustomerType(cboBusinessType)
      Call LoadCustomerGrade(cboEnterpriseType)
      Call LoadRegion(cboArea)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Customer.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Customer.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Customer = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 11550
   Col.Caption = MapText("ที่อยู่")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblWebsite, MapText("เว็บไซต์"))
   Call InitNormalLabel(lblShortName, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblEnterpriseType, MapText("ระดับลูกค้า"))
   Call InitNormalLabel(lblName, MapText("ชื่อลูกค้า"))
   Call InitNormalLabel(lblEmail, MapText("อีเมลล์"))
   Call InitNormalLabel(lblBusinessType, MapText("ประเภทลูกค้า"))
   Call InitNormalLabel(lblBusinessDesc, MapText("รายละเอียดลูกค้า"))
   Call InitNormalLabel(lblCredit, MapText("เครดิต"))
   Call InitNormalLabel(Label2, MapText("วัน"))
   Call InitNormalLabel(lblDiscountPercent, MapText("% ส่วนลด"))
   Call InitNormalLabel(lblCreditAmount, MapText("วงเงิน"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(lblArea, MapText("เขตการขาย"))
   Call InitNormalLabel(lblTaxID, MapText("TAX ID"))
   Call InitNormalLabel(lblCustomerBranch, MapText("สาขา"))
   
   Call InitCombo(cboBusinessType)
   Call InitCombo(cboEnterpriseType)
   Call InitCombo(cboArea)
   
   Call txtShortName.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtEmail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWebSite.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBusinessDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtCreditAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCheckBox(chkTakeAwayFlag, "ลูกค้าหน้าฟาร์ม")
   Call InitCheckBox(chkKeyAccount, "ลูกค้า KeyAccount")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("ที่อยู่")
   TabStrip1.Tabs.Add().Caption = MapText("ราคาขาย")
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Customer = New CCustomer
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Customer.CstAddr Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CCustomerAddress
      Dim Addr As CAddress
      If m_Customer.CstAddr.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Customer.CstAddr, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Set Addr = CR.Addresses

      Values(1) = Addr.ADDRESS_ID
      Values(2) = RealIndex
      Values(3) = Addr.PackAddress
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_Customer.CstPackage Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CP As CCustomerPackage
      
      If m_Customer.CstPackage.Count <= 0 Then
         Exit Sub
      End If
      Set CP = GetItem(m_Customer.CstPackage, RowIndex, RealIndex)
      If CP Is Nothing Then
         Exit Sub
      End If

      Values(1) = CP.CUSTOMER_PACKAGE_ID
      Values(2) = RealIndex
      Values(3) = CP.PKG_NAME
      Values(4) = IDToPackageName(CP.PKG_TYPE)
    
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub



Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.itemcount = CountItem(m_Customer.CstAddr)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      GridEX1.itemcount = CountItem(m_Customer.CstPackage)
      GridEX1.Rebind
   
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtBusinessDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtCredit_Change()
   m_HasModify = True
   If Val(txtCredit.Text) = 0 Then
      txtDiscountPercent.Enabled = True
   Else
      txtDiscountPercent.Enabled = False
   End If
End Sub

Private Sub txtCreditAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtCustomerBranch_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtShortName_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtTaxID_Change()
   m_HasModify = True
End Sub

Private Sub txtWebSite_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()

End Sub
Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 6000
   Col.Caption = MapText("แบบการตั้งราคา")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 6000
   Col.Caption = MapText("ประเภทการตั้งราคา")
End Sub

Public Sub ShowCustomerPackageGrid()
   GridEX1.itemcount = CountItem(m_Customer.CstPackage)
   GridEX1.Rebind
   
   m_HasModify = True
End Sub

