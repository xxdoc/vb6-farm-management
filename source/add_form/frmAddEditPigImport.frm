VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPigImport 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPigImport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5475
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   9657
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1785
         TabIndex        =   6
         Top             =   2550
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   7
         Top             =   3000
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   2100
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlExpenseTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3900
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigStatusLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3450
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeight 
         Height          =   435
         Left            =   5430
         TabIndex        =   4
         Top             =   1650
         Width           =   1755
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPigNo 
         Height          =   435
         Left            =   1770
         TabIndex        =   2
         Top             =   1200
         Width           =   2805
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin VB.Label lblPigNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   510
         TabIndex        =   26
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4170
         TabIndex        =   25
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   7305
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblPigStatus 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   23
         Top             =   3510
         Width           =   1485
      End
      Begin VB.Label lblExpenseType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   22
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   21
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   3885
         TabIndex        =   20
         Top             =   2100
         Width           =   1005
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3870
         TabIndex        =   19
         Top             =   2550
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   10
         Top             =   4650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPigImport.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   11
         Top             =   4650
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   18
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   16
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   15
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   3060
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPigImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String

Private m_PigTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_ExpenseTypes As Collection
Private m_PigStatuss As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
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
      
   Call InitNormalLabel(lblPartType, MapText("ประเภทสุกร"))
   Call InitNormalLabel(lblPart, MapText("สัปดาห์เกิด"))
   Call InitNormalLabel(lblQuantity, MapText("จำนวน"))
   Call InitNormalLabel(lblPrice, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblLocation, MapText("โรงเรือนนำเข้า"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblPigStatus, MapText("สถานะสุกร"))
   Call InitNormalLabel(lblExpenseType, MapText("ประเภทรายจ่าย"))
   Call InitNormalLabel(lblWeight, MapText("น้ำหนัก"))
   Call InitNormalLabel(lblPigNo, MapText("เบอร์หู"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
txtPigNo.Enabled = False

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CImportItem
         
         Set EnpAddr = TempCollection.Item(ID)
         
         uctlParttypeLookup.MyCombo.ListIndex = IDToListIndex(uctlParttypeLookup.MyCombo, PigCodeToID(EnpAddr.PIG_TYPE))
         If EnpAddr.PARENT_ID > 0 Then
            uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PARENT_ID)
            txtPigNo.Text = EnpAddr.PART_NO
         Else
            uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         End If
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)
         uctlPigStatusLookup.MyCombo.ListIndex = IDToListIndex(uctlPigStatusLookup.MyCombo, EnpAddr.PIG_STATUS)
         uctlExpenseTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlExpenseTypeLookup.MyCombo, EnpAddr.EXPENSE_TYPE)
      
         txtQuantity.Text = EnpAddr.IMPORT_AMOUNT
         txtPrice.Text = EnpAddr.ACTUAL_UNIT_PRICE
         txtTotalPrice.Text = EnpAddr.TOTAL_ACTUAL_PRICE
         txtWeight.Text = EnpAddr.TOTAL_WEIGHT
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim Pi As CPartItem

   If Not VerifyCombo(lblPartType, uctlParttypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblWeight, txtWeight, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPrice, txtPrice, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigStatus, uctlPigStatusLookup.MyCombo, Not uctlPigStatusLookup.Enabled) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CImportItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CImportItem
      EnpAddress.Flag = "A"
      Call TempCollection.Add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If

'   Set Pi = New CPartItem
'   Pi.PART_DESC = uctlPartLookup.MyCombo.Text
'   Pi.PART_NO = txtPigNo.Text
'   Pi.INTAKE_FLAG = "N"
'   Pi.PIG_FLAG = "Y"
'   Pi.SPECIFIC_FLAG = "N" '"Y"
'   Pi.PIG_TYPE = PigTypeToCode(uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex)))
'   Pi.PARENT_ID = -1 'uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'   Pi.UNIT_COUNT = -1
'   Pi.PART_TYPE = -1

'txtPigNo.Text = ""
'EnpAddress.OLD_PARENT_ID = 0
'
'   '==== Begin create new PART_ITEM ===
'   If (ShowMode = SHOW_ADD) And (Len(Trim(txtPigNo.Text)) > 0) Then
'      Pi.AddEditMode = SHOW_ADD
'      Call glbDaily.AddEditPartItem(Pi, IsOK, True, glbErrorLog)
'      'Now Pi.PART_ITEM_ID set to new PART_ITEM_ID
'   ElseIf (ShowMode = SHOW_EDIT) Then
'      If (EnpAddress.OLD_PARENT_ID <= 0) And (Len(Trim(txtPigNo.Text)) > 0) Then
'         ' Week ----> Specific
'         Pi.AddEditMode = SHOW_ADD
'         Call glbDaily.AddEditPartItem(Pi, IsOK, True, glbErrorLog)
'      ElseIf (EnpAddress.OLD_PARENT_ID > 0) And (Len(Trim(txtPigNo.Text)) <= 0) Then
'         'Specific ----> Week
'         Pi.PART_ITEM_ID = EnpAddress.PARENT_ID
'         Pi.PARENT_ID = -1
'         Pi.PART_NO = uctlPartLookup.MyTextBox.Text
'      Else
'         'Specific ----> Specific  (Pig NO. change)
'         Pi.AddEditMode = SHOW_EDIT
'         Pi.PART_ITEM_ID = EnpAddress.PART_ITEM_ID
'         Call glbDaily.AddEditPartItem(Pi, IsOK, True, glbErrorLog)
'      End If
'   Else
'      Pi.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
'      Pi.PART_NO = uctlPartLookup.MyTextBox.Text
'   End If
'   '==== End create new PART_ITEM ===

   EnpAddress.PARENT_ID = -1 'Pi.PARENT_ID
   EnpAddress.OLD_PARENT_ID = -1 'Pi.PARENT_ID
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)) 'Pi.PART_ITEM_ID
   
   EnpAddress.PIG_TYPE = PigTypeToCode(uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex)))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.PIG_STATUS = uctlPigStatusLookup.MyCombo.ItemData(Minus2Zero(uctlPigStatusLookup.MyCombo.ListIndex))
   EnpAddress.EXPENSE_TYPE = uctlExpenseTypeLookup.MyCombo.ItemData(Minus2Zero(uctlExpenseTypeLookup.MyCombo.ListIndex))
   EnpAddress.IMPORT_AMOUNT = txtQuantity.Text
   EnpAddress.ACTUAL_UNIT_PRICE = txtPrice.Text
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.CALCULATE_FLAG = "Y"
   EnpAddress.TOTAL_ACTUAL_PRICE = Val(txtTotalPrice.Text)
   EnpAddress.TOTAL_WEIGHT = Val(txtWeight.Text)

   Set Pi = Nothing
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadProductType(uctlParttypeLookup.MyCombo, m_PigTypes)
      Set uctlParttypeLookup.MyCollection = m_PigTypes
      
      Call LoadExpenseType(uctlExpenseTypeLookup.MyCombo, m_ExpenseTypes, "Y")
      Set uctlExpenseTypeLookup.MyCollection = m_ExpenseTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 1, "")
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Call LoadProductStatus(uctlPigStatusLookup.MyCombo, m_PigStatuss)
      Set uctlPigStatusLookup.MyCollection = m_PigStatuss
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         uctlPigStatusLookup.Enabled = False
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PigTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_ExpenseTypes = New Collection
   Set m_PigStatuss = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PigTypes = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_ExpenseTypes = Nothing
   Set m_PigStatuss = Nothing
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

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPigNo_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalPrice_Change()
   m_HasModify = True
   If Val(txtQuantity.Text) <> 0 Then
      txtPrice.Text = Format(Val(txtTotalPrice.Text) / Val(txtQuantity.Text), "0.00")
   End If
End Sub

Private Sub txtWeight_Change()
   m_HasModify = True
End Sub

Private Sub uctlExpenseTypeLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
Dim LocationID As Long
Dim Lc As CLocation

   LocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   If LocationID > 0 Then
      Set Lc = m_Locations(Trim(Str(LocationID)))
      
      If Lc.SALE_FLAG = "Y" Then
         uctlPigStatusLookup.Enabled = True
      Else
         uctlPigStatusLookup.MyTextBox.Text = ""
         uctlPigStatusLookup.MyCombo.ListIndex = -1
         uctlPigStatusLookup.Enabled = False
      End If
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlParttypeLookup_Change()
Dim PigTypeCode As String

   m_HasModify = True

   PigTypeCode = PigTypeToCode(uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex)))
   If PigTypeCode <> "" Then
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, -1, "Y", PigTypeCode, "N")
      Set uctlPartLookup.MyCollection = m_Parts
   End If
End Sub

Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
End Sub
