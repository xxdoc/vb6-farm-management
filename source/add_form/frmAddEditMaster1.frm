VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMaster1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMaster1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame Frame1 
      Height          =   2085
      Left            =   -30
      TabIndex        =   5
      Top             =   420
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   3678
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBBranch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cboBank 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cboGroup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5340
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   390
         Visible         =   0   'False
         Width           =   2655
      End
      Begin prjFarmManagement.uctlTextBox txtCode 
         Height          =   435
         Left            =   2250
         TabIndex        =   0
         Top             =   450
         Width           =   1845
         _extentx        =   4683
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   2250
         TabIndex        =   2
         Top             =   900
         Width           =   5745
         _extentx        =   4683
         _extenty        =   767
      End
      Begin Threed.SSCheck chkFlag2 
         Height          =   435
         Left            =   8040
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkFlag1 
         Height          =   435
         Left            =   8040
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   1965
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   435
         Left            =   6210
         TabIndex        =   13
         Top             =   420
         Visible         =   0   'False
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkFlag 
         Height          =   435
         Left            =   4110
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   60
         TabIndex        =   11
         Top             =   930
         Width           =   2055
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   150
         TabIndex        =   6
         Top             =   480
         Width           =   1965
      End
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   825
      Left            =   0
      TabIndex        =   7
      Top             =   2490
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1455
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   3630
         TabIndex        =   3
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5265
         TabIndex        =   4
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   615
         Index           =   0
         Left            =   11130
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   13230
         TabIndex        =   8
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMaster1"
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
Public MasterKey As String

Private m_PartType As CPartType
Private m_Location As CLocation
Private m_ProductType As CProductType
Private m_ProductStatus As CProductStatus
Private m_House As CHouse
Private m_Country As CCountry
Private m_CustomerType As CCustomerType
Private m_CustomerGrade As CCustomerGrade
Private m_SupplierType As CSupplierType
Private m_SupplierGrade As CSupplierGrade
Private m_SupplierStatus As CSupplierStatus
Private m_Position As CEmpPosition
Private m_Unit As CUnit
Private m_PartGroup As CPartGroup
Private m_ExposeType As CExposeType
Private m_DocumentType As CDocumentType
Private m_Bank As CBank
Private m_BankBranch As CBankBranch
Private m_Region As CRegion
Private m_RevenueType As CRevenueType
Private m_CnDnReasons As CCnDnReason
Private m_BankAccount As CBankAccount
Private m_StatusType As CStatusType
Private m_PackageType As CPackageType

Private m_MasterRef As CMasterRef

Public MasterMode As Long

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboBank_Click()
Dim ID1 As Long
   ID1 = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
   If ID1 > 0 Then
      Call LoadBankBranch(cboBBranch, , ID1)
   End If
   m_HasModify = True
End Sub
Private Sub cboBBranch_Click()
   m_HasModify = True
End Sub
Private Sub cboGroup_Click()
   m_HasModify = True
End Sub
Private Sub chkFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkFlag1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkFlag2_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCode, "")
   Call InitNormalLabel(lblName, "")
   
   If MasterKey = ROOT_TREE & " 1-1" Then
      Call InitCombo(cboGroup)
      Call LoadPartGroup(cboGroup)
      cboGroup.Visible = True
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทวัตถุดิบ"))
      Call InitNormalLabel(lblName, MapText("ประเภทวัตถุดิบ"))
   ElseIf MasterKey = ROOT_TREE & " 1-2" Then
      chkFlag.Visible = True
      Call InitCheckBox(chkFlag, "คลังใหญ่")
      Call InitNormalLabel(lblCode, MapText("รหัสสถานที่จัดเก็บ"))
      Call InitNormalLabel(lblName, MapText("สถานที่จัดเก็บ"))
   ElseIf MasterKey = ROOT_TREE & " 1-3" Then
      Call InitNormalLabel(lblCode, MapText("รหัสหน่วยวัด"))
      Call InitNormalLabel(lblName, MapText("หน่วยวัด"))
   ElseIf MasterKey = ROOT_TREE & " 1-4" Then
      Call InitNormalLabel(lblCode, MapText("รหัสกลุ่มวัตถุดิบ"))
      Call InitNormalLabel(lblName, MapText("กลุ่มวัตถุดิบ"))
   ElseIf MasterKey = ROOT_TREE & " 1-5" Then
      Call InitNormalLabel(lblCode, MapText("รหัสกลุ่มอาหาร"))
      Call InitNormalLabel(lblName, MapText("กลุ่มอาหาร"))
    ElseIf MasterKey = ROOT_TREE & " 1-6" Then
      Call InitNormalLabel(lblCode, MapText("รหัส"))
      Call InitNormalLabel(lblName, MapText("ประเภทการโอน"))
   ElseIf MasterKey = ROOT_TREE & " 2-1" Then
      chkFlag.Visible = True
      Call InitCheckBox(chkFlag, "เป็นพ่อแม่")
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทสุกร"))
      Call InitNormalLabel(lblName, MapText("ประเภทสุกร"))
      chkFlag2.Visible = True
      Call InitCheckBox(chkFlag2, "ล็อคน้ำหนัก")
   ElseIf MasterKey = ROOT_TREE & " 2-2" Then
      chkFlag.Visible = True
      Call InitCombo(cboGroup)
      Call LoadStatusType(cboGroup)
      cboGroup.Visible = True
      Call InitCheckBox(chkFlag, "คิดต้นทุน")
      chkFlag1.Visible = True
      Call InitCheckBox(chkFlag1, "ตายขายไม่ได้")
      Call InitNormalLabel(lblCode, MapText("รหัสสถานะสุกร"))
      Call InitNormalLabel(lblName, MapText("สถานะสุกร"))
      chkFlag2.Visible = True
      Call InitCheckBox(chkFlag2, "ล็อคน้ำหนัก")
   ElseIf MasterKey = ROOT_TREE & " 2-3" Then
      chkFlag.Visible = True
      SSCheck1.Visible = True
      Call InitCheckBox(chkFlag, "เรือนขาย")
      Call InitCheckBox(SSCheck1, "คิดต้นทุนลูกเกิด")
      Call InitNormalLabel(lblCode, MapText("รหัสโรงเรือน"))
      Call InitNormalLabel(lblName, MapText("โรงเรือน"))
      chkFlag2.Visible = True
      Call InitCheckBox(chkFlag2, "ล็อคน้ำหนัก")
   ElseIf MasterKey = ROOT_TREE & " 2-7" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทสถานะ"))
      Call InitNormalLabel(lblName, MapText("ประเภทสถานะ"))
   ElseIf MasterKey = ROOT_TREE & " 3-1" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเทศ"))
      Call InitNormalLabel(lblName, MapText("ประเทศ"))
   ElseIf MasterKey = ROOT_TREE & " 3-2" Then
      Call InitNormalLabel(lblCode, MapText("รหัสระดับลูกค้า"))
      Call InitNormalLabel(lblName, MapText("ระดับลูกค้า"))
   ElseIf MasterKey = ROOT_TREE & " 3-3" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทลูกค้า"))
      Call InitNormalLabel(lblName, MapText("ประเภทลูกค้า"))
   ElseIf MasterKey = ROOT_TREE & " 3-4" Then
      Call InitNormalLabel(lblCode, MapText("รหัสระดับซัพ ฯ"))
      Call InitNormalLabel(lblName, MapText("ระดับซัพ ฯ"))
   ElseIf MasterKey = ROOT_TREE & " 3-5" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทซัพ ฯ"))
      Call InitNormalLabel(lblName, MapText("ประเภทซัพ ฯ"))
   ElseIf MasterKey = ROOT_TREE & " 3-6" Then
      Call InitNormalLabel(lblCode, MapText("รหัสสถานะซัพ ฯ"))
      Call InitNormalLabel(lblName, MapText("สถานะซัพ ฯ"))
   ElseIf MasterKey = ROOT_TREE & " 3-7" Then
      Call InitNormalLabel(lblCode, MapText("รหัสตำแหน่ง"))
      Call InitNormalLabel(lblName, MapText("ตำแหน่ง"))
   ElseIf MasterKey = ROOT_TREE & " 4-2" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทเอกสาร"))
      Call InitNormalLabel(lblName, MapText("ประเภทเอกสาร"))
   ElseIf MasterKey = ROOT_TREE & " 4-3" Then
      Call InitNormalLabel(lblCode, MapText("รหัสธนาคาร"))
      Call InitNormalLabel(lblName, MapText("ธนาคาร"))
   ElseIf MasterKey = ROOT_TREE & " 4-4" Then
      cboGroup.Visible = True
      Call InitCombo(cboGroup)
      Call LoadBank(cboGroup)
      
      Call InitNormalLabel(lblCode, MapText("รหัสสาขาธนาคาร"))
      Call InitNormalLabel(lblName, MapText("สาขาธนาคาร"))
   ElseIf MasterKey = ROOT_TREE & " 4-5" Then
      Call InitNormalLabel(lblCode, MapText("รหัสเขตการขาย"))
      Call InitNormalLabel(lblName, MapText("เขตการขาย"))
   ElseIf MasterKey = ROOT_TREE & " 4-6" Then
      Call InitNormalLabel(lblCode, MapText("รหัสรายรับ"))
      Call InitNormalLabel(lblName, MapText("รายรับ"))
   ElseIf MasterKey = ROOT_TREE & " 4-7" Then
      Call InitNormalLabel(lblCode, MapText("รหัสสาเหตุ"))
      Call InitNormalLabel(lblName, MapText("สาเหตุเพิ่ม/ลดหนี้"))
   ElseIf MasterKey = ROOT_TREE & " 4-8" Then
      Call InitNormalLabel(lblCode, MapText("รหัสบัญชี"))
      Call InitNormalLabel(lblName, MapText("ชื่อบัญชี"))
      Call InitNormalLabel(lblBank, MapText("ธนาคาร"))
      lblBank.Visible = True
      Call InitCombo(cboBank)
      Call InitCombo(cboBBranch)
      Call LoadBank(cboBank)
      Call LoadBankBranch(cboBBranch)
      cboBank.Visible = True
      cboBBranch.Visible = True
   ElseIf MasterKey = ROOT_TREE & " 4-9" Then
      Call InitNormalLabel(lblCode, MapText("รหัสประเภทเช็ค"))
      Call InitNormalLabel(lblName, MapText("ประเภทเช็ค"))
      
   ElseIf MasterKey = ROOT_TREE & " 5-1" Then
      Call InitNormalLabel(lblCode, MapText("รหัสการตั้งราคา"))
      Call InitNormalLabel(lblName, MapText("ประเภทการตั้งราคา"))
   
   
   End If

   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)

   Call InitMainButton(cmdSave, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
      
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If MasterKey = ROOT_TREE & " 1-1" Then
         m_PartType.PART_TYPE_ID = ID
         Call m_PartType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_PartType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_PartType.PART_TYPE_NO
            txtName.Text = m_PartType.PART_TYPE_NAME
            cboGroup.ListIndex = IDToListIndex(cboGroup, m_PartType.PART_GROUP_ID)
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-2" Then
         m_Location.LOCATION_ID = ID
         m_Location.LOCATION_TYPE = 2
         Call m_Location.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_Location.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Location.LOCATION_NO
            txtName.Text = m_Location.LOCATION_NAME
            chkFlag.Value = FlagToCheck(m_Location.MASTER_FLAG)
            chkFlag2.Value = FlagToCheck(m_Location.LOCK_WEIGHT_FLAG)
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-3" Then
         m_Unit.UNIT_ID = ID
         Call m_Unit.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_Unit.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Unit.UNIT_NO
            txtName.Text = m_Unit.UNIT_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-4" Then
         m_PartGroup.PART_GROUP_ID = ID
         Call m_PartGroup.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_PartGroup.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_PartGroup.PART_GROUP_NO
            txtName.Text = m_PartGroup.PART_GROUP_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-5" Then
         m_MasterRef.KEY_ID = ID
         Call m_MasterRef.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_MasterRef.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_MasterRef.KEY_CODE
            txtName.Text = m_MasterRef.KEY_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 1-6" Then
         m_ExposeType.EXPOSE_TYPE_ID = ID
         Call m_ExposeType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_ExposeType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_ExposeType.EXPOSE_TYPE_NO
            txtName.Text = m_ExposeType.EXPOSE_TYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-1" Then
         m_ProductType.PRODUCT_TYPE_ID = ID
         Call m_ProductType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_ProductType.PopulateFromRS(1, m_Rs)
            chkFlag.Value = FlagToCheck(m_ProductType.CAPITAL_FLAG)
            txtCode.Text = m_ProductType.PRODUCT_TYPE_NO
            txtName.Text = m_ProductType.PRODUCT_TYPE_NAME
            chkFlag2.Value = FlagToCheck(m_ProductType.LOCK_WEIGHT_FLAG)
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-2" Then
         m_ProductStatus.PRODUCT_STATUS_ID = ID
         Call m_ProductStatus.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_ProductStatus.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_ProductStatus.PRODUCT_STATUS_NO
            txtName.Text = m_ProductStatus.PRODUCT_STATUS_NAME
            chkFlag.Value = FlagToCheck(m_ProductStatus.CAPITAL_MOVE_FLAG)
            chkFlag1.Value = FlagToCheck(m_ProductStatus.NON_SALE_FLAG)
            cboGroup.ListIndex = IDToListIndex(cboGroup, m_ProductStatus.STATUS_TYPE)
            chkFlag2.Value = FlagToCheck(m_ProductStatus.LOCK_WEIGHT_FLAG)
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-3" Then
         m_Location.LOCATION_ID = ID
         m_Location.LOCATION_TYPE = 1
         Call m_Location.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_Location.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Location.LOCATION_NO
            txtName.Text = m_Location.LOCATION_NAME
            chkFlag.Value = FlagToCheck(m_Location.SALE_FLAG)
            SSCheck1.Value = FlagToCheck(m_Location.CAPITAL_BIRTH_FLAG)
            chkFlag2.Value = FlagToCheck(m_Location.LOCK_WEIGHT_FLAG)
         End If
      ElseIf MasterKey = ROOT_TREE & " 2-7" Then
         m_StatusType.STATUS_TYPE_ID = ID
         Call m_StatusType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_StatusType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_StatusType.STATUS_TYPE_NO
            txtName.Text = m_StatusType.STATUS_TYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-1" Then
         m_Country.COUNTRY_ID = ID
         Call m_Country.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_Country.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Country.COUNTRY_NO
            txtName.Text = m_Country.COUNTRY_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-2" Then
         m_CustomerGrade.CSTGRADE_ID = ID
         Call m_CustomerGrade.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_CustomerGrade.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_CustomerGrade.CSTGRADE_NO
            txtName.Text = m_CustomerGrade.CSTGRADE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-3" Then
         m_CustomerType.CSTTYPE_ID = ID
         Call m_CustomerType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_CustomerType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_CustomerType.CSTTYPE_NO
            txtName.Text = m_CustomerType.CSTTYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-4" Then
         m_SupplierGrade.SUPPLIER_GRADE_ID = ID
         Call m_SupplierGrade.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_SupplierGrade.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_SupplierGrade.SUPPLIER_GRADE_NO
            txtName.Text = m_SupplierGrade.SUPPLIER_GRADE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-5" Then
         m_SupplierType.SUPPLIER_TYPE_ID = ID
         Call m_SupplierType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_SupplierType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_SupplierType.SUPPLIER_TYPE_NO
            txtName.Text = m_SupplierType.SUPPLIER_TYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-6" Then
         m_SupplierStatus.SUPPLIER_STATUS_ID = ID
         Call m_SupplierStatus.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_SupplierStatus.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_SupplierStatus.SUPPLIER_STATUS_NO
            txtName.Text = m_SupplierStatus.SUPPLIER_STATUS_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 3-7" Then
         m_Position.POSITION_ID = ID
         Call m_Position.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_Position.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Position.POSITION_NAME
            txtName.Text = m_Position.POSITION_DESC
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-2" Then
         m_DocumentType.DOCUMENT_TYPE_ID = ID
         Call m_DocumentType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_DocumentType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_DocumentType.DOCUMENT_TYPE_NO
            txtName.Text = m_DocumentType.DOCUMENT_TYPE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-3" Then
         m_Bank.BANK_ID = ID
         Call m_Bank.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_Bank.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Bank.BANK_NO
            txtName.Text = m_Bank.BANK_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-4" Then
         m_BankBranch.BBRANCH_ID = ID
         Call m_BankBranch.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_BankBranch.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_BankBranch.BBRANCH_NO
            txtName.Text = m_BankBranch.BBRANCH_NAME
            cboGroup.ListIndex = IDToListIndex(cboGroup, m_BankBranch.BANK_ID)
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-5" Then
         m_Region.REGION_ID = ID
         Call m_Region.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_Region.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_Region.REGION_NO
            txtName.Text = m_Region.REGION_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-6" Then
         m_RevenueType.REVENUE_TYPE_ID = ID
         Call m_RevenueType.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_RevenueType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_RevenueType.REVENUE_NO
            txtName.Text = m_RevenueType.REVENUE_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-7" Then
         m_CnDnReasons.REASON_ID = ID
         Call m_CnDnReasons.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_CnDnReasons.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_CnDnReasons.REASON_NO
            txtName.Text = m_CnDnReasons.REASON_NAME
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-8" Then
         m_BankAccount.BANK_ACCOUNT_ID = ID
         Call m_BankAccount.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_BankAccount.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_BankAccount.ACCOUNT_NO
            txtName.Text = m_BankAccount.ACCOUNT_NAME
            cboBank.ListIndex = IDToListIndex(cboBank, m_BankAccount.BANK_ID)
            cboBBranch.ListIndex = IDToListIndex(cboBBranch, m_BankAccount.BBRANCH_ID)
         End If
      ElseIf MasterKey = ROOT_TREE & " 4-9" Then
         m_MasterRef.KEY_ID = ID
         Call m_MasterRef.QueryData(m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_MasterRef.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_MasterRef.KEY_CODE
            txtName.Text = m_MasterRef.KEY_NAME
         End If
      
      ElseIf MasterKey = ROOT_TREE & " 5-1" Then
         m_PackageType.PACKAGE_TYPE_ID = ID
         Call m_PackageType.QueryData(1, m_Rs, ItemCount)
         If ItemCount > 0 Then
            Call m_PackageType.PopulateFromRS(1, m_Rs)
            txtCode.Text = m_PackageType.PACKAGE_TYPE_CODE
            txtName.Text = m_PackageType.PACKAGE_TYPE_NAME
         End If
      
      End If
   
      Call EnableForm(Me, True)
   End If
   
   IsOK = True
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSave_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
On Error GoTo ErrorHandler
Dim IsOK As Boolean

   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, Not txtName.Visible) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
      
   If MasterKey = ROOT_TREE & " 1-1" Then
      If Not VerifyCombo(lblCode, cboGroup, False) Then
         Exit Function
      End If
      
      If Not CheckUniqueNs(PARTTYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(PARTTYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_PartType.AddEditMode = ShowMode
      m_PartType.PART_TYPE_NAME = txtName.Text
      m_PartType.RAW_FLAG = "Y"
      m_PartType.PART_TYPE_NO = txtCode.Text
      m_PartType.PART_GROUP_ID = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      Call glbMaster.AddEditPartType(m_PartType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-2" Then
      If Not CheckUniqueNs(LOCATION_NO, txtCode.Text & "2", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(LOCATION_NAME, txtName.Text & "2", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_Location.AddEditMode = ShowMode
      m_Location.LOCATION_NAME = txtName.Text
      m_Location.LOCATION_NO = txtCode.Text
      m_Location.LOCATION_TYPE = 2 'คลัง
      m_Location.SALE_FLAG = "N"
      m_Location.MASTER_FLAG = Check2Flag(chkFlag.Value)
      Call glbMaster.AddEditLocation(m_Location, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-3" Then
      If Not CheckUniqueNs(UNIT_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(UNIT_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_Unit.AddEditMode = ShowMode
      m_Unit.UNIT_NAME = txtName.Text
      m_Unit.UNIT_NO = txtCode.Text
      Call glbMaster.AddEditUnit(m_Unit, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-4" Then
      If Not CheckUniqueNs(PARTGROUP_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(PARTGROUP_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_PartGroup.AddEditMode = ShowMode
      m_PartGroup.PART_GROUP_NAME = txtName.Text
      m_PartGroup.PART_GROUP_NO = txtCode.Text
      Call glbMaster.AddEditPartGroup(m_PartGroup, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-5" Then
      m_MasterRef.AddEditMode = ShowMode
      m_MasterRef.MASTER_AREA = FEED_GROUP
      m_MasterRef.KEY_NAME = txtName.Text
      m_MasterRef.KEY_CODE = txtCode.Text
      Call glbMaster.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 1-6" Then
      If Not CheckUniqueNs(EXPOSE_TYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(EXPOSE_TYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
      
      m_ExposeType.AddEditMode = ShowMode
      m_ExposeType.EXPOSE_TYPE_NAME = txtName.Text
      m_ExposeType.EXPOSE_TYPE_NO = txtCode.Text
      Call glbMaster.AddEditExposeType(m_ExposeType, IsOK, glbErrorLog)
   
   ElseIf MasterKey = ROOT_TREE & " 2-1" Then
      If Not CheckUniqueNs(PRODUCTTYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(PRODUCTTYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_ProductType.AddEditMode = ShowMode
      m_ProductType.PRODUCT_TYPE_NAME = txtName.Text
      m_ProductType.PRODUCT_TYPE_NO = txtCode.Text
      m_ProductType.CAPITAL_FLAG = Check2Flag(chkFlag.Value)
      m_ProductType.LOCK_WEIGHT_FLAG = Check2Flag(chkFlag2.Value)
      
      Call glbMaster.AddEditProductType(m_ProductType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 2-2" Then
      If Not CheckUniqueNs(PRODUCTSTATUS_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(PRODUCTSTATUS_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_ProductStatus.AddEditMode = ShowMode
      m_ProductStatus.PRODUCT_STATUS_NAME = txtName.Text
      m_ProductStatus.PRODUCT_STATUS_NO = txtCode.Text
      m_ProductStatus.CAPITAL_MOVE_FLAG = Check2Flag(chkFlag.Value)
      m_ProductStatus.NON_SALE_FLAG = Check2Flag(chkFlag1.Value)
      m_ProductStatus.STATUS_TYPE = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      m_ProductStatus.LOCK_WEIGHT_FLAG = Check2Flag(chkFlag2.Value)
      Call glbMaster.AddEditProductStatus(m_ProductStatus, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 2-3" Then
      If Not CheckUniqueNs(LOCATION_NO, txtCode.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(LOCATION_NAME, txtName.Text & "1", ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_Location.AddEditMode = ShowMode
      m_Location.LOCATION_NAME = txtName.Text
      m_Location.LOCATION_NO = txtCode.Text
      m_Location.LOCATION_TYPE = 1 'โรงเรือน
      m_Location.SALE_FLAG = Check2Flag(chkFlag.Value)
      m_Location.CAPITAL_BIRTH_FLAG = Check2Flag(SSCheck1.Value)
      m_Location.LOCK_WEIGHT_FLAG = Check2Flag(chkFlag2.Value)
      Call glbMaster.AddEditLocation(m_Location, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 2-7" Then
'      If Not CheckUniqueNs(LOCATION_NO, txtCode.Text & "1", ID) Then
'         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
'         glbErrorLog.ShowUserError
'
'         Call EnableForm(Me, True)
'         txtCode.SetFocus
'         Exit Function
'      End If
'
'      If Not CheckUniqueNs(LOCATION_NAME, txtName.Text & "1", ID) Then
'         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
'         glbErrorLog.ShowUserError
'
'         Call EnableForm(Me, True)
'         txtName.SetFocus
'         Exit Function
'      End If
   
      m_StatusType.AddEditMode = ShowMode
      m_StatusType.STATUS_TYPE_NAME = txtName.Text
      m_StatusType.STATUS_TYPE_NO = txtCode.Text
      Call glbMaster.AddEditStatusType(m_StatusType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-1" Then
      If Not CheckUniqueNs(COUNTRY_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(COUNTRY_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_Country.AddEditMode = ShowMode
      m_Country.COUNTRY_NAME = txtName.Text
      m_Country.COUNTRY_NO = txtCode.Text
      Call glbMaster.AddEditCountry(m_Country, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-2" Then
      If Not CheckUniqueNs(CSTGRADE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(CSTGRADE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_CustomerGrade.AddEditMode = ShowMode
      m_CustomerGrade.CSTGRADE_NAME = txtName.Text
      m_CustomerGrade.CSTGRADE_NO = txtCode.Text
      Call glbMaster.AddEditCustomerGrade(m_CustomerGrade, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-3" Then
      If Not CheckUniqueNs(CSTTYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(CSTTYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_CustomerType.AddEditMode = ShowMode
      m_CustomerType.CSTTYPE_NAME = txtName.Text
      m_CustomerType.CSTTYPE_NO = txtCode.Text
      Call glbMaster.AddEditCustomerType(m_CustomerType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-4" Then
      If Not CheckUniqueNs(SUPPLIERGRADE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(SUPPLIERGRADE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_SupplierGrade.AddEditMode = ShowMode
      m_SupplierGrade.SUPPLIER_GRADE_NAME = txtName.Text
      m_SupplierGrade.SUPPLIER_GRADE_NO = txtCode.Text
      Call glbMaster.AddEditSupplierGrade(m_SupplierGrade, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-5" Then
      If Not CheckUniqueNs(SUPPLIERTYPE_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(SUPPLIERYPE_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_SupplierType.AddEditMode = ShowMode
      m_SupplierType.SUPPLIER_TYPE_NAME = txtName.Text
      m_SupplierType.SUPPLIER_TYPE_NO = txtCode.Text
      Call glbMaster.AddEditSupplierType(m_SupplierType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-6" Then
      If Not CheckUniqueNs(SUPPLIERSTATUS_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
   
      If Not CheckUniqueNs(SUPPLIERSTATUS_NAME, txtName.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtName.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtName.SetFocus
         Exit Function
      End If
   
      m_SupplierStatus.AddEditMode = ShowMode
      m_SupplierStatus.SUPPLIER_STATUS_NAME = txtName.Text
      m_SupplierStatus.SUPPLIER_STATUS_NO = txtCode.Text
      Call glbMaster.AddEditSupplierStatus(m_SupplierStatus, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 3-7" Then
      If Not CheckUniqueNs(POSITION_NO, txtCode.Text, ID) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtCode.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
                  
         Call EnableForm(Me, True)
         txtCode.SetFocus
         Exit Function
      End If
      
      m_Position.AddEditMode = ShowMode
      m_Position.POSITION_DESC = txtName.Text
      m_Position.POSITION_NAME = txtCode.Text
      Call glbMaster.AddEditPosition(m_Position, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-2" Then
      m_DocumentType.AddEditMode = ShowMode
      m_DocumentType.DOCUMENT_TYPE_NAME = txtName.Text
      m_DocumentType.DOCUMENT_TYPE_NO = txtCode.Text
      Call glbMaster.AddEditDocumentType(m_DocumentType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-3" Then
      m_Bank.AddEditMode = ShowMode
      m_Bank.BANK_NAME = txtName.Text
      m_Bank.BANK_NO = txtCode.Text
      Call glbMaster.AddEditBank(m_Bank, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-4" Then
      m_BankBranch.AddEditMode = ShowMode
      m_BankBranch.BBRANCH_NAME = txtName.Text
      m_BankBranch.BBRANCH_NO = txtCode.Text
      m_BankBranch.BANK_ID = cboGroup.ItemData(Minus2Zero(cboGroup.ListIndex))
      Call glbMaster.AddEditBankBranch(m_BankBranch, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-5" Then
      m_Region.AddEditMode = ShowMode
      m_Region.REGION_NAME = txtName.Text
      m_Region.REGION_NO = txtCode.Text
      Call glbMaster.AddEditRegion(m_Region, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-6" Then
      m_RevenueType.AddEditMode = ShowMode
      m_RevenueType.REVENUE_NAME = txtName.Text
      m_RevenueType.REVENUE_NO = txtCode.Text
      Call glbMaster.AddEditRevenueType(m_RevenueType, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-7" Then
      m_CnDnReasons.AddEditMode = ShowMode
      m_CnDnReasons.REASON_NAME = txtName.Text
      m_CnDnReasons.REASON_NO = txtCode.Text
      Call glbMaster.AddEditCnDnReason(m_CnDnReasons, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-8" Then
      If Not VerifyCombo(lblBank, cboBank, False) Then
         Exit Function
      End If
      If Not VerifyCombo(lblBank, cboBBranch, False) Then
         Exit Function
      End If
      m_BankAccount.AddEditMode = ShowMode
      m_BankAccount.ACCOUNT_NAME = txtName.Text
      m_BankAccount.ACCOUNT_NO = txtCode.Text
      m_BankAccount.BANK_ID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
      m_BankAccount.BBRANCH_ID = cboBBranch.ItemData(Minus2Zero(cboBBranch.ListIndex))
      
      Call glbMaster.AddEditBankAccount(m_BankAccount, IsOK, glbErrorLog)
   ElseIf MasterKey = ROOT_TREE & " 4-9" Then
      m_MasterRef.AddEditMode = ShowMode
      m_MasterRef.MASTER_AREA = CHEQUE_TYPE
      m_MasterRef.KEY_NAME = txtName.Text
      m_MasterRef.KEY_CODE = txtCode.Text
      Call glbMaster.AddEditMasterRef(m_MasterRef, IsOK, True, glbErrorLog)
    ElseIf MasterKey = ROOT_TREE & " 5-1" Then
      m_PackageType.AddEditMode = ShowMode
      m_PackageType.PACKAGE_TYPE_NAME = txtName.Text
      m_PackageType.PACKAGE_TYPE_CODE = txtCode.Text
      Call glbMaster.AddEditPackageType(m_PackageType, IsOK, glbErrorLog)
   End If
   
   IsOK = True
   Call EnableForm(Me, True)
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   Call EnableForm(Me, True)
   SaveData = False
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdSave_Click
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
   
   Set m_PartType = New CPartType
   Set m_Location = New CLocation
   Set m_ProductType = New CProductType
   Set m_ProductStatus = New CProductStatus
   Set m_House = New CHouse
   Set m_Country = New CCountry
   Set m_CustomerGrade = New CCustomerGrade
   Set m_CustomerType = New CCustomerType
   Set m_SupplierGrade = New CSupplierGrade
   Set m_SupplierType = New CSupplierType
   Set m_SupplierStatus = New CSupplierStatus
   Set m_Position = New CEmpPosition
   Set m_Unit = New CUnit
   Set m_PartGroup = New CPartGroup
   Set m_ExposeType = New CExposeType
   Set m_DocumentType = New CDocumentType
   Set m_Bank = New CBank
   Set m_BankBranch = New CBankBranch
   Set m_Region = New CRegion
   Set m_RevenueType = New CRevenueType
   Set m_CnDnReasons = New CCnDnReason
   Set m_BankAccount = New CBankAccount
   Set m_StatusType = New CStatusType
   Set m_PackageType = New CPackageType
   
   Set m_MasterRef = New CMasterRef
   
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   Set m_PartType = Nothing
   Set m_Location = Nothing
   Set m_ProductType = Nothing
   Set m_ProductStatus = Nothing
   Set m_House = Nothing
   Set m_Country = Nothing
   Set m_CustomerGrade = Nothing
   Set m_CustomerType = Nothing
   Set m_SupplierGrade = Nothing
   Set m_SupplierType = Nothing
   Set m_SupplierStatus = Nothing
   Set m_Position = Nothing
   Set m_Unit = Nothing
   Set m_PartGroup = Nothing
   Set m_ExposeType = Nothing
   Set m_DocumentType = Nothing
   Set m_Bank = Nothing
   Set m_BankBranch = Nothing
   Set m_Region = Nothing
   Set m_RevenueType = Nothing
   Set m_CnDnReasons = Nothing
   Set m_BankAccount = Nothing
   Set m_StatusType = Nothing
   Set m_PackageType = Nothing
   
   Set m_MasterRef = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub
Private Sub SSCheck1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub
