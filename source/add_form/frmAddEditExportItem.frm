VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditExportItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   4380
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
   Icon            =   "frmAddEditExportItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6420
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   11324
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlMFGDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   5760
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlDate uctlEXPDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   6195
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1785
         TabIndex        =   4
         Top             =   2100
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
         TabIndex        =   2
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlHouseLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   2550
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigWeekLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   7
         Top             =   3450
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   3000
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlExposeTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   5280
         Visible         =   0   'False
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblExpdate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   4800
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblMFGdate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   4350
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblExposetype 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   26
         Top             =   3960
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   23
         Top             =   3060
         Width           =   1485
      End
      Begin VB.Label lblPigWeek 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   22
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label lblHouse 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   21
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   2130
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3210
         TabIndex        =   11
         Top             =   4680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExportItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4920
         TabIndex        =   12
         Top             =   4680
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
         TabIndex        =   19
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   18
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   16
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   330
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditExportItem"
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
Public DefaultLocationID As Long
Public DefaultHouseID As Long

Private m_PartTypes As Collection
'Private m_ExposeType As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_PigTypes As Collection
Private m_Expose As Collection

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

   Call InitNormalLabel(lblPartType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblPart, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblPrice, MapText("ราคา"))
   Call InitNormalLabel(lblLocation, MapText("จากคลัง"))
   Call InitNormalLabel(lblHouse, MapText("โรงเรือน"))
   Call InitNormalLabel(lblPigWeek, MapText("สัปดาห์เกิด"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(lblPigType, MapText("ประเภทสุกร"))
'   Call InitNormalLabel(lblExposetype, MapText("ประเภทการเบิก"))
'   Call InitNormalLabel(lblMFGdate, MapText("วันที่ผลิต"))
'   Call InitNormalLabel(lblExpdate, MapText("วันที่หมดอายุ"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CExportItem
         
         Set EnpAddr = TempCollection.Item(ID)
         
         uctlParttypeLookup.MyCombo.ListIndex = IDToListIndex(uctlParttypeLookup.MyCombo, EnpAddr.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)
         uctlHouseLookup.MyCombo.ListIndex = IDToListIndex(uctlHouseLookup.MyCombo, EnpAddr.HOUSE_ID)
         
         uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, PigCodeToID(EnpAddr.PART_PIG_TYPE))
         uctlPigWeekLookup.MyCombo.ListIndex = IDToListIndex(uctlPigWeekLookup.MyCombo, EnpAddr.PIG_ID)
         
       '  uctlExposeTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlExposeTypeLookup.MyCombo, EnpAddr.EXPOSE_TYPE_ID)
       '  uctlMFGDate.ShowDate = EnpAddr.MFG_DATE
       '  uctlEXPDate.ShowDate = EnpAddr.EXP_DATE
 
         txtQuantity.Text = EnpAddr.EXPORT_AMOUNT
         txtPrice.Text = EnpAddr.EXPORT_AVG_PRICE
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      ElseIf ShowMode = SHOW_ADD Then
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, DefaultLocationID)
         uctlHouseLookup.MyCombo.ListIndex = IDToListIndex(uctlHouseLookup.MyCombo, DefaultHouseID)
      End If
      uctlLocationLookup.SetTextFocus = True
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

   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, uctlParttypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblHouse, uctlHouseLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigType, uctlPigTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigWeek, uctlPigWeekLookup.MyCombo, False) Then
      Exit Function
   End If
'   If Not VerifyCombo(lblExposetype, uctlExposeTypeLookup.MyCombo, False) Then
'      Exit Function
'   End If
'   If Not VerifyDate(lblMFGdate, uctlMFGDate, False) Then
'      Exit Function
'   End If
'   If Not VerifyDate(lblExpdate, uctlEXPDate, False) Then
'      Exit Function
'   End If
   
'   If Not VerifyTextControl(lblPrice, txtPrice, True) Then
'      Exit Function
'   End If

   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CExportItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CExportItem
      EnpAddress.Flag = "A"
      Call TempCollection.Add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If

   EnpAddress.PART_TYPE = uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex))
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.EXPORT_AMOUNT = txtQuantity.Text
   EnpAddress.EXPORT_AVG_PRICE = Val(txtPrice.Text)
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PIG_NO = uctlPigWeekLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.HOUSE_ID = uctlHouseLookup.MyCombo.ItemData(Minus2Zero(uctlHouseLookup.MyCombo.ListIndex))
   EnpAddress.HOUSE_NAME = uctlHouseLookup.MyCombo.Text
   EnpAddress.PIG_ID = uctlPigWeekLookup.MyCombo.ItemData(Minus2Zero(uctlPigWeekLookup.MyCombo.ListIndex))
   EnpAddress.PART_PIG_TYPE = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
'   EnpAddress.EXPOSE_TYPE_ID = uctlExposeTypeLookup.MyCombo.ItemData(Minus2Zero(uctlExposeTypeLookup.MyCombo.ListIndex))
'   EnpAddress.MFG_DATE = uctlMFGDate.ShowDate
'   EnpAddress.EXP_DATE = uctlEXPDate.ShowDate
   
   EnpAddress.CALCULATE_FLAG = "Y"
   
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadProductType(uctlPigTypeLookup.MyCombo, m_PigTypes)
      Set uctlPigTypeLookup.MyCollection = m_PigTypes
      
      Call LoadPartType(uctlParttypeLookup.MyCombo, m_PartTypes)
      Set uctlParttypeLookup.MyCollection = m_PartTypes
      
'      Call LoadExposeType(uctlExposeTypeLookup.MyCombo, m_ExposeType)
'      Set uctlExposeTypeLookup.MyCollection = m_ExposeType
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
      Call LoadLocation(uctlHouseLookup.MyCombo, m_Houses, 1)
      Set uctlHouseLookup.MyCollection = m_Houses

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
 '  Set m_ExposeType = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Houses = New Collection
   Set m_Pigs = New Collection
   Set m_PigTypes = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Expose = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_Houses = Nothing
   Set m_Pigs = Nothing
   Set m_PigTypes = Nothing
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

'Private Sub uctlMFGDate_HasChange()
'   m_HasModify = True
'End Sub
'
'Private Sub uctlExpDate_HasChange()
'   m_HasModify = True
'End Sub

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

Private Sub uctlHouseLookup_Change()
   m_HasModify = True
End Sub
'Private Sub uctlExposeTypeLookup_Change()
'   m_HasModify = True
'End Sub
Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
Dim PartItemID As Long
Dim LocationID As Long
Dim PL As CPartLocation
Dim iCount As Long

   m_HasModify = True
   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   LocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   
   If (PartItemID <= 0) Or (LocationID <= 0) Then
      Exit Sub
   End If
   
   Set PL = New CPartLocation
   PL.PART_LOCATION_ID = -1
   PL.PART_ITEM_ID = PartItemID
   PL.LOCATION_ID = LocationID
   Call PL.QueryData(1, m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call PL.PopulateFromRS(m_Rs)
      txtPrice.Text = Format(PL.AVG_PRICE, "0.00")
   Else
      txtPrice.Text = Format(0, "0.00")
   End If
   
   Set PL = Nothing
End Sub

Private Sub uctlParttypeLookup_Change()
Dim PartTypeID As Long

   PartTypeID = uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex))
   
   Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
   Set uctlPartLookup.MyCollection = m_Parts
   
   m_HasModify = True
End Sub

Private Sub uctlPigTypeLookup_Change()
Dim PigTypeCode As String

   m_HasModify = True
   
   PigTypeCode = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
   If PigTypeCode <> "" Then
      Call LoadPartItem(uctlPigWeekLookup.MyCombo, m_Pigs, -1, "Y", PigTypeCode)
      Set uctlPigWeekLookup.MyCollection = m_Pigs
   End If
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
