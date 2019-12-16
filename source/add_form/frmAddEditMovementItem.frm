VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditMovementItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7350
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
   Icon            =   "frmAddEditMovementItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6765
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   11933
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3405
         Left            =   180
         TabIndex        =   3
         Top             =   1800
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   6006
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
         HeaderFontBold  =   -1  'True
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditMovementItem.frx":08CA
         Column(2)       =   "frmAddEditMovementItem.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditMovementItem.frx":0A36
         FormatStyle(2)  =   "frmAddEditMovementItem.frx":0B92
         FormatStyle(3)  =   "frmAddEditMovementItem.frx":0C42
         FormatStyle(4)  =   "frmAddEditMovementItem.frx":0CF6
         FormatStyle(5)  =   "frmAddEditMovementItem.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmAddEditMovementItem.frx":0E86
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   690
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1140
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblPig 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   13
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   12
         Top             =   750
         Width           =   1485
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3450
         TabIndex        =   6
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMovementItem.frx":105E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   180
         TabIndex        =   4
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMovementItem.frx":1378
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1800
         TabIndex        =   5
         Top             =   5280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   11
         Top             =   300
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   7
         Top             =   6030
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMovementItem.frx":1692
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   8
         Top             =   6030
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMovementItem"
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
Public TempCollection2 As Collection
Public COMMIT_FLAG As String
Public DOCUMENT_DATE As Date

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_PigStatuss As Collection
Private m_PigTypes As Collection

Private m_MovementItems As Collection
Public FromDate As Date
Public ToDate As Date

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim D As CExportItem
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
   
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ประเภทค่าใช้จ่าย", "-", "กลุ่มวัตถุดิบ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      frmCapitalRatio.ExpensePartMode = 1
   ElseIf lMenuChosen = 3 Then
      frmCapitalRatio.ExpensePartMode = 2
   End If
   
   OKClick = False
   
   Set frmCapitalRatio.TempCollection = m_MovementItems
   frmCapitalRatio.ShowMode = SHOW_ADD
   frmCapitalRatio.HeaderText = MapText("เพิ่มมูลค่าต้นทุน")
   Load frmCapitalRatio
   frmCapitalRatio.Show 1

   OKClick = frmCapitalRatio.OKClick

   Unload frmCapitalRatio
   Set frmCapitalRatio = Nothing

   If OKClick Then
      GridEX1.ItemCount = CountItem(m_MovementItems)
      GridEX1.Rebind
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim TempID As Long
Dim Er As CExpenseRatio
Dim iCount As Long
Dim D As CExpenseRatio
Dim PigImportAmounts As Collection
Dim FirstDate As Date
Dim LastDate As Date
Dim II As CImportItem
Dim PigCount As Long

'   Set oMenu = New cPopupMenu
'   lMenuChosen = oMenu.Popup("ทุกโรงเรือน", "ตามประเภทรายจ่าย")
'   If lMenuChosen = 0 Then
'      Exit Sub
'   End If

   Set PigImportAmounts = New Collection
Call GetFirstLastDate(DOCUMENT_DATE, FirstDate, LastDate)
   
lMenuChosen = 1

   If lMenuChosen <> 1 Then
      TempID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
      If TempID < 0 Then
         glbErrorLog.LocalErrorMsg = "กรุณาทำการเลือกประเภทรายจ่ายก่อน"
         glbErrorLog.ShowUserError
         
         uctlToLocationLookup.SetFocus
         Exit Sub
      End If
      
      Set Er = New CExpenseRatio
      Er.EXPENSE_RATIO_ID = -1
      Er.RO_ITEM_ID = TempID
      Er.SELECT_FLAG = "Y"
      Call Er.QueryData(1, m_Rs, iCount)
      Set Er = Nothing
      
      For Each Er In m_MovementItems
         Er.Flag = "D"
      Next Er
      
      While Not m_Rs.EOF
         Set Er = New CExpenseRatio
         Call Er.PopulateFromRS(1, m_Rs)
         If Er.SELECT_FLAG = "Y" Then
            Er.Flag = "A"
            Call m_MovementItems.Add(Er)
         End If
         Set Er = Nothing
         m_Rs.MoveNext
      Wend
   Else
      Dim Locations As Collection
      Dim Lc As CLocation
      
      For Each Er In m_MovementItems
         Er.Flag = "D"
      Next Er
      
      Set Locations = New Collection
      Call LoadLocation(Nothing, Locations, 1, "")
      For Each Lc In Locations
'         Call LoadPigImportAmount(Nothing, PigImportAmounts, FirstDate, LastDate, "", Lc.LOCATION_ID)
'         PigCount = 0
'         For Each Ii In PigImportAmounts
'            PigCount = PigCount + Ii.IMPORT_AMOUNT
'         Next Ii
PigCount = 1
         'เอาแต่โรงเรือนที่มีหมูมาแบ่งเปอร์เซ็นต์
         If PigCount > 0 Then
            Set Er = New CExpenseRatio
            Er.LOCATION_ID = Lc.LOCATION_ID
            Er.LOCATION_NAME = Lc.LOCATION_NAME
            Er.RATIO = 0
            Er.PIG_COUNT = 0
            Er.RATIO_AMOUNT = 0
            Er.SELECT_FLAG = "Y"
            Er.Flag = "A"
            
            Call m_MovementItems.Add(Er)
            Set Er = Nothing
         End If
      Next Lc
      
      Set Locations = Nothing
   End If
   
   GridEX1.ItemCount = CountItem(m_MovementItems)
   GridEX1.Rebind
   
   Set PigImportAmounts = Nothing
   m_HasModify = True
   Set Er = Nothing
End Sub

Private Sub cmdCalculate_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("แบ่งเท่า ๆ กัน", "แบ่งตามกำหนด", "แบ่งตามสัดส่วนสุกร")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call CalculateRatio(m_MovementItems, lMenuChosen)
   GridEX1.ItemCount = CountItem(m_MovementItems)
   GridEX1.Rebind
   
   m_HasModify = True
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
      
   If ID1 <= 0 Then
      m_MovementItems.Remove (ID2)
   Else
      m_MovementItems.Item(ID2).Flag = "D"
   End If

   GridEX1.ItemCount = CountItem(m_MovementItems)
   GridEX1.Rebind
   m_HasModify = True
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_MovementItems Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CMovementItem
   If m_MovementItems.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_MovementItems, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.MOVEMENT_ITEM_ID
   Values(2) = RealIndex
   If CR.EXPENSE_TYPE > 0 Then
      Values(3) = CR.EXPENSE_TYPE_NAME
   Else
      Values(3) = CR.PART_GROUP_NAME
   End If
   Values(4) = FormatNumber(CR.CAPITAL_AMOUNT)
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

'   If Not VerifyAccessRight("GROUP_QUERY_RIGHT") Then
'      Exit Sub
'   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False

   frmCapitalRatio.ID = ID
   Set frmCapitalRatio.TempCollection = m_MovementItems
   frmCapitalRatio.HeaderText = MapText("แก้ไขมูลค่าต้นทุน")
   frmCapitalRatio.ShowMode = SHOW_EDIT
   Load frmCapitalRatio
   frmCapitalRatio.Show 1

   OKClick = frmCapitalRatio.OKClick

   Unload frmCapitalRatio
   Set frmCapitalRatio = Nothing

   If OKClick Then
      GridEX1.ItemCount = CountItem(m_MovementItems)
      GridEX1.Rebind
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
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
   Col.Width = 6405
   Col.Caption = MapText("ประเภทรายจ่าย/กลุ่มวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2625
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ต้นทุน")
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
   
   Call InitNormalLabel(lblToLocation, MapText("โรงเรือน"))
   Call InitNormalLabel(lblPigType, MapText("ประเภทสุกร"))
   Call InitNormalLabel(lblPig, MapText("สัปดาห์เกิด"))
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข"))
   Call InitMainButton(cmdDelete, MapText("ลบ"))
   
   Call InitGrid1
End Sub

Private Sub CalculateRatio(TempCol As Collection, Ind As Long)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CCapitalMovement
         
         Set Di = TempCollection.Item(ID)
         
         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, Di.FROM_HOUSE_ID)
         uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, PigCodeToID(Di.PIG_TYPE))
         uctlPigLookup.MyCombo.ListIndex = IDToListIndex(uctlPigLookup.MyCombo, Di.PIG_ID)

         Set m_MovementItems = Di.MovementItems
         
         GridEX1.ItemCount = CountItem(m_MovementItems)
         GridEX1.Rebind
      
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub CalculateTotalCapital(D As CCapitalMovement)
Dim Mi As CMovementItem
Dim Sum As Double

   Sum = 0
   For Each Mi In D.MovementItems
      If Mi.Flag <> "D" Then
         Sum = Sum + Mi.CAPITAL_AMOUNT
      End If
   Next Mi
   D.TOTAL_CAPITAL = Sum
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

   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Function
   End If
'   If Not VerifyTextControl(lblWeight, txtWeight, False) Then
'      Exit Function
'   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Di As CCapitalMovement
   If ShowMode = SHOW_ADD Then
      Set Di = New CCapitalMovement
      
      Di.Flag = "A"
      Call TempCollection.Add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If

   Di.FROM_HOUSE_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   Di.HOUSE_NAME = uctlToLocationLookup.MyCombo.Text
   Di.PIG_TYPE = uctlPigTypeLookup.MyTextBox.Text
   Di.PIG_NO = uctlPigLookup.MyTextBox.Text
   Di.PIG_ID = uctlPigLookup.MyCombo.ItemData(Minus2Zero(uctlPigLookup.MyCombo.ListIndex))
   Di.COMMIT_FLAG = "N"
   Di.DOCUMENT_CATEGORY = 3
   Di.TX_TYPE = "I"
   
   Set Di.MovementItems = m_MovementItems
   Call CalculateTotalCapital(Di)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 1, "")
      Set uctlToLocationLookup.MyCollection = m_Houses

      Call LoadProductType(uctlPigTypeLookup.MyCombo, m_PigTypes)
      Set uctlPigTypeLookup.MyCollection = m_PigTypes
      
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
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Houses = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatuss = New Collection
   Set m_PigTypes = New Collection
   Set m_MovementItems = New Collection
   
   Call GetFirstLastDate(DOCUMENT_DATE, FromDate, ToDate)
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_Houses = Nothing
   Set m_Pigs = Nothing
   Set m_PigStatuss = Nothing
   Set m_PigTypes = Nothing
   Set m_MovementItems = Nothing
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

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigTypeLookup_Change()
Dim PigTypeCode As String

   m_HasModify = True
   
   PigTypeCode = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
   If PigTypeCode <> "" Then
      Call LoadPartItem(uctlPigLookup.MyCombo, m_Pigs, -1, "Y", PigTypeCode)
      Set uctlPigLookup.MyCollection = m_Pigs
   End If
End Sub

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
