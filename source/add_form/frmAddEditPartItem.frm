VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPartItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "frmAddEditPartItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10440
   StartUpPosition =   1  'CenterOwner
   Begin prjFarmManagement.uctlTextBox txtDrug 
      Height          =   435
      Left            =   1860
      TabIndex        =   21
      Top             =   2340
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   767
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8730
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   15399
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlSupplierLookup 
         Height          =   375
         Left            =   1860
         TabIndex        =   22
         Top             =   2800
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboFeedGroup 
         Height          =   315
         Left            =   7140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3240
         Width           =   2955
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3720
         Width           =   2955
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3240
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1410
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   960
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3015
         Left            =   360
         TabIndex        =   14
         Top             =   4440
         Width           =   9885
         _ExtentX        =   17436
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
         Column(1)       =   "frmAddEditPartItem.frx":27A2
         Column(2)       =   "frmAddEditPartItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPartItem.frx":290E
         FormatStyle(2)  =   "frmAddEditPartItem.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPartItem.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPartItem.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPartItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPartItem.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtRegistNum 
         Height          =   435
         Left            =   1860
         TabIndex        =   17
         Top             =   1880
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSize 
         Height          =   435
         Left            =   7140
         TabIndex        =   18
         Top             =   1880
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin VB.Label lblSupplierNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2800
         Width           =   1575
      End
      Begin VB.Label lblDrug 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   20
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label lblSize 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5490
         TabIndex        =   19
         Top             =   1880
         Width           =   1575
      End
      Begin VB.Label lblRegistNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1880
         Width           =   1575
      End
      Begin VB.Label lblFeedGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5490
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Top             =   3720
         Width           =   1575
      End
      Begin Threed.SSCheck chkIntake 
         Height          =   345
         Left            =   6420
         TabIndex        =   1
         Top             =   990
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   131073
         Caption         =   "SSCheck1"
         TripleState     =   -1  'True
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1050
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5235
         TabIndex        =   7
         Top             =   7680
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3600
         TabIndex        =   6
         Top             =   7680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":2F36
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPartItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartItem As CPartItem
Private m_Suppliers As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_InventoryBalances As Collection

Private Sub cmdPasswd_Click()

End Sub

Private Sub cboFeedGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
End Sub

Private Sub chkIntake_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
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
   Col.Width = 2685
   Col.Caption = MapText("สถานที่จัดเก็บ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2415
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนคงคลัง")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2550
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาเฉลี่ย")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2205
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคาหลังสุด")
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_PartItem.PART_ITEM_ID = ID
      m_PartItem.QueryFlag = 0
      If Not glbDaily.QueryPartItem(m_PartItem, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_PartItem.PopulateFromRS(1, m_Rs)
      
      txtName.Text = m_PartItem.PART_DESC
      txtPartNo.Text = m_PartItem.PART_NO
      cboPartType.ListIndex = IDToListIndex(cboPartType, m_PartItem.PART_TYPE)
      cboUnit.ListIndex = IDToListIndex(cboUnit, m_PartItem.UNIT_COUNT)
      cboFeedGroup.ListIndex = IDToListIndex(cboFeedGroup, m_PartItem.FEED_GROUP)
      chkIntake.Value = FlagToCheck(m_PartItem.INTAKE_FLAG)
      
      txtRegistNum.Text = m_PartItem.REGIST_NUM
      txtSize.Text = m_PartItem.DRUG_SIZE
      txtDrug.Text = m_PartItem.DRUG_NAME
      uctlSupplierLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierLookup.MyCombo, m_PartItem.SUPPLIER_ID)
      
      GridEX1.ItemCount = CountItem(m_PartItem.PartLocations)
      GridEX1.Rebind
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_InventoryBalances = Nothing
   Set m_Suppliers = Nothing
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim Ba As CBalanceAccum

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_PartItem.PartLocations Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CPartLocation
   If m_PartItem.PartLocations.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_PartItem.PartLocations, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Set Ba = GetBalanceAccum(m_InventoryBalances, CR.LOCATION_ID & "-" & CR.PART_ITEM_ID)
   
   Values(1) = CR.PART_LOCATION_ID
   Values(2) = RealIndex
   Values(3) = CR.LOCATION_NAME
   Values(4) = FormatNumber(Ba.BALANCE_AMOUNT)
   Values(5) = FormatNumber(Ba.AVG_PRICE)
   Values(6) = FormatNumber(Ba.AVG_PRICE)
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("INVENTORY_PART_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("INVENTORY_PART_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblPartNo, txtPartNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
  If Not VerifyCombo(lblSupplierNo, uctlSupplierLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, cboPartType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(PARTNO_UNIQUE, txtPartNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPartNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_PartItem.PART_ITEM_ID = ID
   m_PartItem.AddEditMode = ShowMode
   m_PartItem.PIG_FLAG = "N"
   m_PartItem.INTAKE_FLAG = Check2Flag(chkIntake.Value)
   m_PartItem.PART_NO = txtPartNo.Text
   m_PartItem.PART_DESC = txtName.Text
   m_PartItem.PART_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
   m_PartItem.UNIT_COUNT = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
   m_PartItem.FEED_GROUP = cboFeedGroup.ItemData(Minus2Zero(cboFeedGroup.ListIndex))
   
   m_PartItem.REGIST_NUM = txtRegistNum.Text
   m_PartItem.DRUG_SIZE = txtSize.Text
   m_PartItem.DRUG_NAME = txtDrug.Text
   m_PartItem.SUPPLIER_ID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditPartItem(m_PartItem, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
Dim NewDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(cboPartType)
      Call LoadUnit(cboUnit)
      Call LoadMaster(cboFeedGroup, , FEED_GROUP)
      Call LoadSupplier(uctlSupplierLookup.MyCombo, m_Suppliers)
      Set uctlSupplierLookup.MyCollection = m_Suppliers
      
      If ShowMode = SHOW_EDIT Then
'         NewDate = DateAdd("D", 1, Now)
'         Call LoadInventoryBalanceEx(Nothing, m_InventoryBalances, NewDate, -1, "", , ID)
      
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
   Call InitGrid1
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblName, MapText("ชื่อวัตถุดิบ"))
   Call InitNormalLabel(lblPartNo, MapText("หมายเลขวัตถุดิบ"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   Call InitNormalLabel(lblFeedGroup, MapText("กลุ่มอาหาร"))
   Call InitNormalLabel(lblSupplierNo, MapText("รหัสซัพ ฯ"))
   
   Call InitCheckBox(chkIntake, "นำไปคิด Intake")
   Call InitNormalLabel(lblRegistNum, MapText("เลขทะเบียน"))
   Call InitNormalLabel(lblDrug, MapText("ตัวยา"))
   Call InitNormalLabel(lblSize, MapText("ขนาดบรรจุ"))
   Call InitNormalLabel(lblSupplierNo, MapText("ซัพพลายเออร์"))

   chkIntake.Value = FlagToCheck("Y")
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboPartType)
   Call InitCombo(cboUnit)
   Call InitCombo(cboFeedGroup)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_PartItem = New CPartItem
   Set m_Rs = New ADODB.Recordset

   Set m_InventoryBalances = New Collection
   Set m_Suppliers = New Collection
   
   Call EnableForm(Me, False)
   m_HasActivate = False
      
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub txtDrug_Change()
   m_HasModify = True
End Sub

Private Sub txtPartNo_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub txtRegistNum_Change()
   m_HasModify = True
End Sub

Private Sub txtSize_Change()
   m_HasModify = True
End Sub
Private Sub uctlSupplierLookup_Change()
   m_HasModify = True
End Sub
