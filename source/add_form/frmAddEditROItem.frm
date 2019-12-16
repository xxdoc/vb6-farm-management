VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditRoItem 
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
   Icon            =   "frmAddEditROItem.frx":0000
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
      TabIndex        =   14
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
      TabIndex        =   15
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   11933
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeight 
         Height          =   435
         Left            =   5190
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   1800
         TabIndex        =   5
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   5205
         TabIndex        =   6
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExpenseDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   750
         Width           =   7335
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2835
         Left            =   180
         TabIndex        =   7
         Top             =   2340
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   5001
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
         Column(1)       =   "frmAddEditROItem.frx":08CA
         Column(2)       =   "frmAddEditROItem.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditROItem.frx":0A36
         FormatStyle(2)  =   "frmAddEditROItem.frx":0B92
         FormatStyle(3)  =   "frmAddEditROItem.frx":0C42
         FormatStyle(4)  =   "frmAddEditROItem.frx":0CF6
         FormatStyle(5)  =   "frmAddEditROItem.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmAddEditROItem.frx":0E86
      End
      Begin Threed.SSCommand cmdCalculate 
         Height          =   525
         Left            =   7680
         TabIndex        =   11
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditROItem.frx":105E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   7200
         TabIndex        =   1
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditROItem.frx":1378
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3450
         TabIndex        =   10
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditROItem.frx":1692
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   180
         TabIndex        =   8
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditROItem.frx":19AC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1800
         TabIndex        =   9
         Top             =   5280
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblExpenseDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   23
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   22
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3975
         TabIndex        =   21
         Top             =   1710
         Width           =   1125
      End
      Begin VB.Label Label2 
         Height          =   345
         Left            =   7215
         TabIndex        =   20
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   7200
         TabIndex        =   19
         Top             =   1230
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3960
         TabIndex        =   18
         Top             =   1260
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   17
         Top             =   360
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   12
         Top             =   6030
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditROItem.frx":1CC6
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   13
         Top             =   6030
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   16
         Top             =   1260
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditRoItem"
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

Private m_ExpenseRatios As Collection
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

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   
   frmExpenseRatio.TotalAmount = Val(txtTotalPrice.Text)
   Set frmExpenseRatio.TempCollection = m_ExpenseRatios
   frmExpenseRatio.ShowMode = SHOW_ADD
   frmExpenseRatio.HeaderText = MapText("เพิ่มส่วนแบ่งโรงเรือน")
   Load frmExpenseRatio
   frmExpenseRatio.Show 1

   OKClick = frmExpenseRatio.OKClick

   Unload frmExpenseRatio
   Set frmExpenseRatio = Nothing

   If OKClick Then
      GridEX1.ItemCount = CountItem(m_ExpenseRatios)
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
      
      For Each Er In m_ExpenseRatios
         Er.Flag = "D"
      Next Er
      
      While Not m_Rs.EOF
         Set Er = New CExpenseRatio
         Call Er.PopulateFromRS(1, m_Rs)
         If Er.SELECT_FLAG = "Y" Then
            Er.Flag = "A"
            Call m_ExpenseRatios.Add(Er)
         End If
         Set Er = Nothing
         m_Rs.MoveNext
      Wend
   Else
      Dim Locations As Collection
      Dim Lc As CLocation
      
      For Each Er In m_ExpenseRatios
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
            
            Call m_ExpenseRatios.Add(Er)
            Set Er = Nothing
         End If
      Next Lc
      
      Set Locations = Nothing
   End If
   
   GridEX1.ItemCount = CountItem(m_ExpenseRatios)
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

   Call CalculateRatio(m_ExpenseRatios, lMenuChosen)
   GridEX1.ItemCount = CountItem(m_ExpenseRatios)
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
      m_ExpenseRatios.Remove (ID2)
   Else
      m_ExpenseRatios.Item(ID2).Flag = "D"
   End If

   GridEX1.ItemCount = CountItem(m_ExpenseRatios)
   GridEX1.Rebind
   m_HasModify = True
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_ExpenseRatios Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CExpenseRatio
   If m_ExpenseRatios.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_ExpenseRatios, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.EXPENSE_RATIO_ID
   Values(2) = RealIndex
   Values(3) = CR.LOCATION_NAME
   Values(4) = FormatNumber(CR.RATIO)
   Values(5) = FormatNumber(CR.RATIO_AMOUNT)
   Values(6) = FormatNumber(CR.PIG_COUNT)
      
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

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False

   frmExpenseRatio.TotalAmount = Val(txtTotalPrice.Text)
   frmExpenseRatio.ID = ID
   Set frmExpenseRatio.TempCollection = m_ExpenseRatios
   frmExpenseRatio.HeaderText = MapText("แก้ไขส่วนแบ่งโรงเรือน")
   frmExpenseRatio.ShowMode = SHOW_EDIT
   Load frmExpenseRatio
   frmExpenseRatio.Show 1

   OKClick = frmExpenseRatio.OKClick

   Unload frmExpenseRatio
   Set frmExpenseRatio = Nothing

   If OKClick Then
      GridEX1.ItemCount = CountItem(m_ExpenseRatios)
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
   Col.Width = 4275
   Col.Caption = MapText("โรงเรือน")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1875
   Col.Caption = MapText("เปอร์เซนต์")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2625
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน")

   Set Col = GridEX1.Columns.Add '6
   Col.TextAlignment = jgexAlignRight
   Col.Width = 0
   Col.Caption = MapText("จำนวนสุกร")
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
   
   Call InitNormalLabel(lblExpenseDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblToLocation, MapText("ประเภทรายจ่าย"))
   Call InitNormalLabel(lblWeight, MapText("น้ำหนัก"))
   Call InitNormalLabel(Label1, MapText("หน่วย"))
   Call InitNormalLabel(Label2, MapText("บาท/หน่วย"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblAvgPrice, MapText("ราคา"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtAvgPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   txtAvgPrice.Enabled = False
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtExpenseDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCalculate.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข"))
   Call InitMainButton(cmdDelete, MapText("ลบ"))
   Call InitMainButton(cmdAuto, MapText("..."))
   Call InitMainButton(cmdCalculate, MapText("คำนวณ"))
   
   Call InitGrid1
End Sub

Private Sub CalculateRatio(TempCol As Collection, Ind As Long)
Dim D As CExpenseRatio
Dim Sum As Double
Dim Amt As Double
Dim Percent As Double
Dim Count As Long
Dim PigBirths As Collection
Dim PigImports As Collection
Dim PigExports As Collection
Dim II As CImportItem
Dim Ii2 As CImportItem
Dim EI As CExportItem
Dim PigAmount As Double
Dim NewDate As Date
Dim TempSum As Double

   Count = 0
   Percent = 0
   Sum = 0
   Amt = Val(txtTotalPrice.Text)
   
   Set PigBirths = New Collection
   Set PigImports = New Collection
   Set PigExports = New Collection
   
   If Ind = 1 Then
      For Each D In TempCol
         If D.Flag <> "D" Then
            Count = Count + 1
         End If
      Next D
      
      For Each D In TempCol
         If D.Flag <> "D" Then
            D.RATIO = MyDiff(100, Count)
            D.RATIO_AMOUNT = (D.RATIO * Amt) / 100
            
            If D.Flag <> "A" Then
               D.Flag = "E"
            End If
         End If
      Next D
   ElseIf Ind = 2 Then
      For Each D In TempCol
         If D.Flag <> "D" Then
            Percent = Percent + D.RATIO
            D.RATIO_AMOUNT = (D.RATIO * Amt) / 100
         
            If D.Flag <> "A" Then
               D.Flag = "E"
            End If
         End If
      Next D
   ElseIf Ind = 3 Then
      'แบ่งตามจำนวนหมู ยกมา + เกิด
      Call LoadHousePigBirthAmount(Nothing, PigBirths, FromDate, ToDate)
      PigAmount = 0
      For Each D In TempCol
         If D.Flag <> "D" Then
            Set II = GetImportItem(PigBirths, Trim(Str(D.LOCATION_ID)))
            PigAmount = PigAmount + II.IMPORT_AMOUNT
         End If
      Next D
   
      NewDate = DateAdd("D", -1, FromDate)
      If FromDate > 0 Then
         Call LoadPigImportByHouse(Nothing, PigImports, -1, NewDate)
      End If
      For Each D In TempCol
         If D.Flag <> "D" Then
            Set II = GetImportItem(PigImports, Trim(Str(D.LOCATION_ID)))
            PigAmount = PigAmount + II.IMPORT_AMOUNT
         End If
      Next D
      
      If FromDate > 0 Then
         Call LoadPigExportByHouse(Nothing, PigExports, -1, NewDate)
      End If
      For Each D In TempCol
         If D.Flag <> "D" Then
            Set EI = GetExportItem(PigExports, Trim(Str(D.LOCATION_ID)))
            PigAmount = PigAmount - EI.EXPORT_AMOUNT
         End If
      Next D
      
TempSum = 0
      For Each D In TempCol
         If D.Flag <> "D" Then
            Set II = GetImportItem(PigBirths, Trim(Str(D.LOCATION_ID))) 'เกิด
            
            Set Ii2 = GetImportItem(PigImports, Trim(Str(D.LOCATION_ID))) 'ยกมาเข้า
            Set EI = GetExportItem(PigExports, Trim(Str(D.LOCATION_ID))) 'ยกมาออก
            
            D.RATIO = MyDiffEx(CDbl(II.IMPORT_AMOUNT + (Ii2.IMPORT_AMOUNT - EI.EXPORT_AMOUNT)), PigAmount) * 100#
            D.RATIO_AMOUNT = MyDiffEx(Amt, PigAmount) * (CDbl(II.IMPORT_AMOUNT + (Ii2.IMPORT_AMOUNT - EI.EXPORT_AMOUNT)))
TempSum = TempSum + D.RATIO_AMOUNT
            If D.Flag <> "A" Then
               D.Flag = "E"
            End If
         End If
      Next D
   End If
''debug.print "== " & TempSum
   Set PigBirths = Nothing
   Set PigImports = Nothing
   Set PigExports = Nothing
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
         Dim Di As CROItem
         
         Set Di = TempCollection.Item(ID)
         
         txtQuantity.Text = Di.ITEM_AMOUNT
         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, Di.EXPENSE_TYPE)
         txtExpenseDesc.Text = Di.EXPENSE_DESC
         txtWeight.Text = Di.TOTAL_WEIGHT
         txtTotalPrice.Text = Di.TOTAL_PRICE
         txtAvgPrice.Text = Di.AVG_PRICE
         
         Set m_ExpenseRatios = Di.ExpenseRatios
         GridEX1.ItemCount = CountItem(m_ExpenseRatios)
         GridEX1.Rebind
      
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

   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblExpenseDesc, txtExpenseDesc, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
'   If Not VerifyTextControl(lblWeight, txtWeight, False) Then
'      Exit Function
'   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Di As CROItem
   If ShowMode = SHOW_ADD Then
      Set Di = New CROItem
      
      Di.Flag = "A"
      Call TempCollection.Add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If

   Di.ITEM_AMOUNT = txtQuantity.Text
   Di.EXPENSE_TYPE = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   Di.EXPENSE_TYPE_NAME = uctlToLocationLookup.MyCombo.Text
   Di.EXPENSE_DESC = txtExpenseDesc.Text
   Di.TOTAL_WEIGHT = Val(txtWeight.Text)
   Di.TOTAL_PRICE = Val(txtTotalPrice.Text)
   Di.AVG_PRICE = Val(txtAvgPrice.Text)
   Di.AVG_WEIGHT = 0
   Set Di.ExpenseRatios = m_ExpenseRatios
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadExpenseType(uctlToLocationLookup.MyCombo, m_Houses)
      Set uctlToLocationLookup.MyCollection = m_Houses

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
   Set m_ExpenseRatios = New Collection
   
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
   Set m_ExpenseRatios = Nothing
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

Private Sub txtExpenseDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
   If Val(txtQuantity.Text) > 0 Then
      txtAvgPrice.Text = Format(Val(txtTotalPrice.Text) / Val(txtQuantity.Text), "0.00")
   Else
      txtAvgPrice.Text = "0.00"
   End If
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
   If Val(txtQuantity.Text) > 0 Then
      txtAvgPrice.Text = Format(Val(txtTotalPrice.Text) / Val(txtQuantity.Text), "0.00")
   Else
      txtAvgPrice.Text = "0.00"
   End If
End Sub

Private Sub txtWeight_Change()
   m_HasModify = True
   If Val(txtQuantity.Text) > 0 Then
      txtAvgPrice.Text = Format(Val(txtTotalPrice.Text) / Val(txtQuantity.Text), "0.00")
   Else
      txtAvgPrice.Text = "0.00"
   End If
End Sub

Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
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
