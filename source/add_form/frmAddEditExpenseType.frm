VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddEditExpenseType 
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   Icon            =   "frmAddEditExpenseType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2940
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5186
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ListView lsvMain 
         Height          =   2775
         Left            =   9630
         TabIndex        =   10
         Top             =   5820
         Visible         =   0   'False
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "โรงเรือน"
            Object.Width           =   11324
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "อัตราส่วน"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   690
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddEditExpenseType.frx":27A2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2655
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMotherNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1530
         Width           =   5175
         _ExtentX        =   11404
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSCheck ChkDepreciationGoodFlag 
         Height          =   435
         Left            =   7440
         TabIndex        =   12
         Top             =   1560
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkSellParentFlag 
         Height          =   435
         Left            =   7440
         TabIndex        =   11
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkDeplicateFlag 
         Height          =   435
         Left            =   5790
         TabIndex        =   2
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkExtraFlag 
         Height          =   435
         Left            =   4590
         TabIndex        =   1
         Top             =   1080
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2805
         TabIndex        =   4
         Top             =   2100
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditExpenseType.frx":307C
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   4455
         TabIndex        =   5
         Top             =   2100
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblMotherNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   1140
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditExpenseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_ExpenseType As CExpenseType
Private m_Houses As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String
Private m_SumUnit As Double
Private m_OldPartItemID As Long
Private m_Locations As Collection

Private Sub LoadLocationTreeView(Col As Collection)
Dim C As CExpenseRatio
Dim N As ListItem
      
      lsvMain.ListItems.Clear
      For Each C In Col
         Set N = lsvMain.ListItems.Add(, Trim(Str(C.LOCATION_ID)) & "-X", C.LOCATION_NAME & " (" & C.LOCATION_NO & ")", 1, 1)
         N.ListSubItems.Add.Text = C.RATIO & "%"
         N.Tag = C.LOCATION_ID
         If C.SELECT_FLAG = "Y" Then
            N.Checked = True
         Else
            N.Checked = False
         End If
      Next C
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_ExpenseType.EXPENSE_TYPE_ID = ID
      If Not glbMaster.QueryExpenseType(m_ExpenseType, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_ExpenseType.PopulateFromRS(1, m_Rs)
      txtDocumentNo.Text = m_ExpenseType.EXPENSE_TYPE_NO
      txtMotherNo.Text = m_ExpenseType.EXPENSE_TYPE_NAME
      chkExtraFlag.Value = FlagToCheck(m_ExpenseType.BUY_FLAG)
      chkDeplicateFlag.Value = FlagToCheck(m_ExpenseType.DEPLICATE_FLAG)
      ChkDepreciationGoodFlag.Value = FlagToCheck(m_ExpenseType.DEPRECIATION_GOOD_FLAG)
      chkSellParentFlag.Value = FlagToCheck(m_ExpenseType.SELL_PARENT_FLAG)
      
      
'      GridEX1.ItemCount = CountItem(m_ExpenseType.ExpenseRatios)
'      GridEX1.Rebind
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

Private Function CheckIntregrity() As Boolean
Dim N As ListItem
Dim Sum As Double

   Sum = 0
   For Each N In lsvMain.ListItems
      If N.Checked Then
         Sum = Sum + Val(Replace(N.ListSubItems(1).Text, "%", ""))
      End If
   Next N
   
   If CLng(Sum) <> 100# Then
      CheckIntregrity = False
   Else
      CheckIntregrity = True
   End If
End Function

Private Function Is100(Col As Collection) As Boolean
Dim D As CExptypeRatio
Dim TempSum As Double

   TempSum = 0
   For Each D In Col
      If D.Flag <> "D" Then
         TempSum = TempSum + D.RATIO
      End If
   Next D
   
   Is100 = (TempSum = 100 Or TempSum = 0)
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Pi As CPartItem
   
   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMotherNo, txtMotherNo, False) Then
      Exit Function
   End If
   
'   If Not CheckIntregrity Then
'      glbErrorLog.LocalErrorMsg = "จำนวนรวมของอัตราส่วนทั้งหมดจะต้องเท่ากับ 100"
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
'   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If m_ExpenseType.ExpenseRatios.Count > 0 Then
      If Not Is100(m_ExpenseType.ExpenseRatios) Then
         glbErrorLog.LocalErrorMsg = "จำนวนรวมของอัตราส่วนจะต้องเท่ากับ 100"
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_ExpenseType.AddEditMode = ShowMode
   m_ExpenseType.EXPENSE_TYPE_ID = ID
   m_ExpenseType.EXPENSE_TYPE_NO = txtDocumentNo.Text
   m_ExpenseType.EXPENSE_TYPE_NAME = txtMotherNo.Text
   m_ExpenseType.BUY_FLAG = Check2Flag(chkExtraFlag.Value)
   m_ExpenseType.DEPLICATE_FLAG = Check2Flag(chkDeplicateFlag.Value)
   m_ExpenseType.SELL_PARENT_FLAG = Check2Flag(chkSellParentFlag.Value)
   m_ExpenseType.DEPRECIATION_GOOD_FLAG = Check2Flag(ChkDepreciationGoodFlag.Value)
   
   Call EnableForm(Me, False)
   If Not glbMaster.AddEditExpenseType(m_ExpenseType, IsOK, glbErrorLog) Then
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

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub
Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkDeplicateFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ChkDepreciationGoodFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkExtraFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Function IsIn(T As CExptypeRatio, Col As Collection) As Boolean
Dim D As CExptypeRatio

   IsIn = False
   For Each D In Col
      If (D.Flag <> "D") And (T.LOCATION_ID = D.LOCATION_ID) Then
         IsIn = True
         Exit Function
      End If
   Next D
End Function

Private Sub chkSellParentFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
'Private Sub cmdAdd_Click()
'Dim OKClick As Boolean
'Dim D As CExportItem
'Dim lMenuChosen As Long
'Dim oMenu As cPopupMenu
'Dim Locations As Collection
'Dim Lc As CLocation
'Dim Et As CExptypeRatio
'
'   If Not cmdAdd.Enabled Then
'      Exit Sub
'   End If
'
'   Set oMenu = New cPopupMenu
'   lMenuChosen = oMenu.Popup("เพิ่มโรงเรือนทั้งหมด", "-", "เพิ่มทีละรายการ")
'   If lMenuChosen = 0 Then
'      Exit Sub
'   End If
'
'   If lMenuChosen = 2 Then
'      OKClick = False
'      Set frmExpTypeRatio.TempCollection = m_ExpenseType.ExpenseRatios
'      frmExpTypeRatio.ShowMode = SHOW_ADD
'      frmExpTypeRatio.HeaderText = MapText("เพิ่มส่วนแบ่งโรงเรือน")
'      Load frmExpTypeRatio
'      frmExpTypeRatio.Show 1
'
'      OKClick = frmExpTypeRatio.OKClick
'
'      Unload frmExpTypeRatio
'      Set frmExpTypeRatio = Nothing
'
'      If OKClick Then
'         GridEX1.ItemCount = CountItem(m_ExpenseType.ExpenseRatios)
'         GridEX1.Rebind
'      End If
'
'      If OKClick Then
'         m_HasModify = True
'      End If
'   Else
'      Set Locations = New Collection
'      Call LoadLocation(Nothing, Locations, 1, "")
'      For Each Lc In Locations
'         Set Et = New CExptypeRatio
'         Et.Flag = "A"
'         Et.LOCATION_ID = Lc.LOCATION_ID
'         Et.LOCATION_NAME = Lc.LOCATION_NAME
'         Et.LOCATION_NO = Lc.LOCATION_NO
'         Et.RATIO = 0
'         If Not IsIn(Et, m_ExpenseType.ExpenseRatios) Then
'            Call m_ExpenseType.ExpenseRatios.Add(Et)
'         End If
'         Set Et = Nothing
'      Next Lc
'      Set Locations = Nothing
'
'      GridEX1.ItemCount = CountItem(m_ExpenseType.ExpenseRatios)
'      GridEX1.Rebind
'      m_HasModify = True
'   End If
'End Sub
'
'Private Sub cmdDelete_Click()
'Dim ID1 As Long
'Dim ID2 As Long
'
'   If Not cmdDelete.Enabled Then
'      Exit Sub
'   End If
'
'   If Not VerifyGrid(GridEX1.Value(1)) Then
'      Exit Sub
'   End If
'
'   If Not ConfirmDelete(GridEX1.Value(3)) Then
'      Exit Sub
'   End If
'
'   ID2 = GridEX1.Value(2)
'   ID1 = GridEX1.Value(1)
'
'   If ID1 <= 0 Then
'      m_ExpenseType.ExpenseRatios.Remove (ID2)
'   Else
'      m_ExpenseType.ExpenseRatios.Item(ID2).Flag = "D"
'   End If
'
'   GridEX1.ItemCount = CountItem(m_ExpenseType.ExpenseRatios)
'   GridEX1.Rebind
'   m_HasModify = True
'End Sub

Private Sub cmdDeplicateFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_ExpenseType.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_ExpenseType.QueryFlag = 0
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_ExpenseType = Nothing
   Set m_Houses = Nothing
   Set m_Locations = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub
'
'Private Sub InitGrid1()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.Add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 6795
'   Col.Caption = MapText("โรงเรือน")
'
'   Set Col = GridEX1.Columns.Add '4
'   Col.Width = 1785
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("อัตราส่วน")
'End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblMotherNo, MapText("ประเภทรายจ่าย"))
   Call InitNormalLabel(lblDocumentNo, MapText("รหัสรายจ่าย"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtMotherNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call InitCheckBox(chkExtraFlag, "ซื้อสุกร")
   Call InitCheckBox(chkDeplicateFlag, "เสื่อมพันธ์")
   Call InitCheckBox(ChkDepreciationGoodFlag, "เสื่อมราคา")
   Call InitCheckBox(chkSellParentFlag, "ขายพ่อแม่")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
'   Call InitTreeView
'   Call InitGrid1
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
'   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
'   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
'   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
End Sub

Private Sub InitTreeView()
   lsvMain.Font.Name = GLB_FONT
   lsvMain.Font.Size = 14
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
   Set m_ExpenseType = New CExpenseType
   Set m_Houses = New Collection
   Set m_Locations = New Collection
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtParentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
End Sub
'
'Private Sub GridEX1_DblClick()
'   Call cmdEdit_Click
'End Sub
'
'Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
'On Error GoTo ErrorHandler
'Dim RealIndex As Long
'
'   glbErrorLog.ModuleName = Me.Name
'   glbErrorLog.RoutineName = "UnboundReadData"
'
'   If m_ExpenseType.ExpenseRatios Is Nothing Then
'      Exit Sub
'   End If
'
'   If RowIndex <= 0 Then
'      Exit Sub
'   End If
'
'   Dim CR As CExptypeRatio
'   If m_ExpenseType.ExpenseRatios.Count <= 0 Then
'      Exit Sub
'   End If
'   Set CR = GetItem(m_ExpenseType.ExpenseRatios, RowIndex, RealIndex)
'   If CR Is Nothing Then
'      Exit Sub
'   End If
'
'   Values(1) = CR.EXPTYPE_RATIO_ID
'   Values(2) = RealIndex
'   Values(3) = CR.LOCATION_NAME
'   Values(4) = FormatNumber(CR.RATIO)
'
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub

Private Sub lsvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
''debug.print ColumnHeader.Index & " " & ColumnHeader.Width
End Sub

Private Sub lsvMain_DblClick()
   If lsvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   frmExpenseRatio.ShowMode = SHOW_EDIT
   frmExpenseRatio.ID = lsvMain.SelectedItem.Index
   Set frmExpenseRatio.TempCollection = m_ExpenseType.ExpenseRatios
   Load frmExpenseRatio
   frmExpenseRatio.Show 1
   
   If frmExpenseRatio.OKClick Then
      Call LoadLocationTreeView(m_ExpenseType.ExpenseRatios)
      m_HasModify = True
   End If
   
   Unload frmExpenseRatio
   Set frmExpenseRatio = Nothing
End Sub

Private Sub lsvMain_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtMotherNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlHouseLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlResponseByLookup_Change()
   m_HasModify = True
End Sub
