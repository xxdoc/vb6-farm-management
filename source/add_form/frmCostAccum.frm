VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCostAccum 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmCostAccum.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4665
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   8229
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBatch 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3180
         Width           =   2985
      End
      Begin VB.ComboBox cboCommitType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   2985
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   2370
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   12
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   6
         Top             =   2700
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblBatch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   19
         Top             =   3240
         Width           =   1575
      End
      Begin Threed.SSCheck chkLoss 
         Height          =   375
         Left            =   5940
         TabIndex        =   3
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCommitType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1980
         Width           =   1725
      End
      Begin Threed.SSCheck chkBalanceFlag 
         Height          =   375
         Left            =   5940
         TabIndex        =   1
         Top             =   1050
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   8
         Top             =   3870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCostAccum.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   17
         Top             =   2820
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   2850
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   10
         Top             =   3870
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   9
         Top             =   3870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCostAccum.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCostAccum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_TempSearchs1 As Collection
Private m_MovementItemSearchs1 As Collection
Private m_MovementItemSearchs2 As Collection
Private m_MovementItemSearchs3 As Collection
Private m_PigBirthInMonthLocations As Collection

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Jill ADD
Private m_CostAccumSearchs As Collection
Private m_CostAccumSearchExs As Collection
Private m_ExportIDs As Collection
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Jill ADD

Private m_ImportItems As Collection
Private m_ExportItems As Collection
Private m_ImportPigs As Collection
Private m_ExportPigs As Collection

Private m_ProductStatuss As Collection
Private m_PigTypes As Collection
Private m_PartItems As Collection
Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboPosition_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboCommitType_Click()
   m_HasModify = True
End Sub

Private Sub chkBalanceFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkLoss_Click(Value As Integer)
   m_HasModify = True
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

   If Flag Then
'      Call EnableForm(Me, False)
'
'      m_Employee.EMP_ID = ID
'      m_Employee.QueryFlag = 1
'      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Employee.PopulateFromRS(1, m_Rs)
      
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("ADMIN_GROUP_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("ADMIN_GROUP_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
   
'   If Not VerifyCombo(lblPartType, cboPartType, False) Then
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Employee.EMP_ID = ID
   m_Employee.AddEditMode = ShowMode
   m_Employee.PASS_STATUS = "Y"
   
   m_Employee.EmpName.AddEditMode = ShowMode
   m_Employee.EName.AddEditMode = ShowMode
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEmployee(m_Employee, IsOK, True, glbErrorLog) Then
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

Private Function IsIn(PigType As String) As Boolean
Dim Ps As CProductType

   IsIn = False
   For Each Ps In m_PigTypes
      If (PigType = Ps.PRODUCT_TYPE_NO) And (Ps.CAPITAL_FLAG = "Y") Then
         IsIn = True
         Exit Function
      End If
   Next Ps
End Function

Private Sub GenerateExportMoveMent(EI As CExportItem, BatchID As Long, Optional FileID As Long)
Dim CA As CCost_Accum
Dim CB As CCost_Accum
Dim CostSearch As CCost_Accum
Dim CostSearchEx As CCost_Accum

   'ไม่ต้องทำถ้าเบิกไปให้หมู G, L, B
   If IsIn(EI.TO_PIG_TYPE) Then
      Exit Sub
   End If
      
   ' ไม่ต้องมีในส่วนของ จำนวนหมู
      
   Set CA = New CCost_Accum
   Set CB = New CCost_Accum
      
   CA.AddEditMode = SHOW_ADD
   CA.DOCUMENT_DATE = EI.DOCUMENT_DATE                            '1
   CA.DOCUMENT_TYPE = EI.DOCUMENT_TYPE                              '2
   CA.LOCATION_ID = EI.HOUSE_ID                                                       '3
   CA.BATCH_ID = BatchID                                                                           '4
   CA.PART_ITEM_ID = EI.PIG_ID                                                             '5
   CA.COST_RAW = EI.EXPORT_AVG_PRICE * EI.EXPORT_AMOUNT                       '6
   CA.COST_EXP = 0                                                                                                                           '7
   CA.CUS_ID = EI.CUS_ID                                                                                                                           '8
   CA.DOCUMENT_CATEGORY = 1                                                                                                                           '8
   
   Print #FileID, "---------------------------"
   Print #FileID, "ใบเบิกวัตถุดิบ" & "              " & CA.GetKey1 & "/               " & CA.COST_RAW & "                 " & CA.COST_EXP & "                 " & CA.COST_PB
   
   Set CostSearch = GetCostAccumSearch(m_CostAccumSearchs, CA.GetKey1)
   If CostSearch Is Nothing Then
      Call m_CostAccumSearchs.Add(CA, CA.GetKey1)
   Else
      CostSearch.COST_RAW = CostSearch.COST_RAW + CA.COST_RAW
      Print #FileID, "ใบเบิกวัตถุดิบ" & "              " & CA.GetKey1 & "/               " & CostSearch.COST_RAW & "                 " & CostSearch.COST_EXP & "                 " & CostSearch.COST_PB
      Print #FileID, "---------------------------"
   End If
   
   
   CB.AddEditMode = SHOW_ADD
   CB.DOCUMENT_DATE = EI.DOCUMENT_DATE                            '1
   CB.DOCUMENT_TYPE = EI.DOCUMENT_TYPE                              '2
   CB.LOCATION_ID = EI.HOUSE_ID                                                       '3
   CB.BATCH_ID = BatchID                                                                           '4
   CB.PART_ITEM_ID = EI.PIG_ID                                                             '5
   CB.COST_RAW = EI.EXPORT_AVG_PRICE * EI.EXPORT_AMOUNT                       '6
   CB.COST_EXP = 0                                                                                                                           '7
   CB.CUS_ID = EI.CUS_ID                                                                                                                           '8
   
   Set CostSearchEx = GetCostAccumSearch(m_CostAccumSearchExs, CB.GetKey2)
   If CostSearchEx Is Nothing Then
      Call m_CostAccumSearchExs.Add(CB, CB.GetKey2)
   Else
      CostSearchEx.COST_RAW = CostSearchEx.COST_RAW + CB.COST_RAW
   End If
   
   Set CA = Nothing
   Set CB = Nothing
   
End Sub

Private Sub AddImportItem(II As CImportItem)
Dim TempII As CImportItem

   Set TempII = New CImportItem
   Call TempII.CopyObject(II)
   Call m_ImportItems.Add(II, Trim(Str(II.GUI_ID)))
   Set TempII = Nothing
End Sub

Private Sub AddExportItem(EI As CExportItem)
Dim TempEi As CExportItem

   Set TempEi = New CExportItem
   Call TempEi.CopyObject(EI)
   Call m_ExportItems.Add(EI, Trim(Str(EI.GUI_ID)))
   Set TempEi = Nothing
End Sub

Private Sub GenerateExportMoveMentEx1(EI As CExportItem)
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1

   'ไม่ต้องทำถ้าเบิกไปให้หมู G, L, B
   If IsIn(EI.TO_PIG_TYPE) Then
      Exit Sub
   End If
   
   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   
   Cm.AddEditMode = SHOW_ADD
   Cm.COMMIT_FLAG = EI.COMMIT_FLAG
   Cm.DOCUMENT_NO = EI.DOCUMENT_NO
   Cm.DOCUMENT_DATE = EI.DOCUMENT_DATE
   Cm.IVD_ID = EI.INVENTORY_DOC_ID
   Cm.DOCUMENT_CATEGORY = 1
   Cm.DOCUMENT_TYPE = EI.DOCUMENT_TYPE
   Cm.TX_TYPE = EI.TX_TYPE
   Cm.TX_AMOUNT = 0
   Cm.FROM_HOUSE_ID = EI.HOUSE_ID
   Cm.TO_HOUSE_ID = 0
   Cm.PIG_ID = EI.PIG_ID
   Cm.PIG_STATUS = 0
   Cm.TX_SEQ = EI.TRANSACTION_SEQ
   Cm.REPLACE_FLAG = EI.REPLACE_FLAG
   Call Cm.AddEditData

   Mi.AddEditMode = SHOW_ADD
   Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
   Mi.PART_ITEM_ID = EI.PART_ITEM_ID
   Mi.EXPENSE_TYPE = 0
   Mi.CAPITAL_AMOUNT = EI.EXPORT_AVG_PRICE * EI.EXPORT_AMOUNT
   Call Mi.AddEditData

   'วัตถุดิบ
   Set S = New CMovementItemSearch1
   S.PART_ITEM_ID = Mi.PART_ITEM_ID
   S.PIG_ID = Cm.PIG_ID
   S.HOUSE_ID = Cm.FROM_HOUSE_ID
   S.CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
   Set TempSearch = GetMovementSearch1(m_MovementItemSearchs1, S.GetKey1)
   If TempSearch Is Nothing Then
      Call m_MovementItemSearchs1.Add(S, S.GetKey1)
   Else
      TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
   End If
   Set S = Nothing
   
   Call AddExportItem(EI)
   
   Set Mi = Nothing
   Set Cm = Nothing
End Sub

Private Sub GeneratePigImportMoveMent(II As CImportItem, BatchID As Long, Optional FileID As Long)
Dim CA As CCost_Accum
Dim CB As CCost_Accum
Dim CostSearch As CCost_Accum
Dim CostSearchEx As CCost_Accum

   Set CA = New CCost_Accum
   Set CB = New CCost_Accum
   
   CA.AddEditMode = SHOW_ADD
   CA.DOCUMENT_DATE = II.DOCUMENT_DATE                            '1
   CA.DOCUMENT_TYPE = II.DOCUMENT_TYPE                              '2
   CA.LOCATION_ID = II.LOCATION_ID                                                       '3
   CA.BATCH_ID = BatchID                                                                           '4
   CA.PART_ITEM_ID = II.PART_ITEM_ID                                                             '5
   CA.COST_RAW = 0                                                                                       '6
   CA.COST_EXP = II.TOTAL_INCLUDE_PRICE                                       '7
   CA.CUS_ID = II.CUS_ID                                       '8
   CA.DOCUMENT_CATEGORY = 1
   
   Print #FileID, "---------------------------"
   Print #FileID, "ใบซื้อหมู" & "              " & CA.GetKey1 & "/               " & CA.COST_RAW & "                 " & CA.COST_EXP & "                 " & CA.COST_PB
   
   
   Set CostSearch = GetCostAccumSearch(m_CostAccumSearchs, CA.GetKey1)
   If CostSearch Is Nothing Then
      Call m_CostAccumSearchs.Add(CA, CA.GetKey1)
   Else
      CostSearch.COST_EXP = CostSearch.COST_EXP + CA.COST_EXP
      Print #FileID, "ใบซื้อหมู" & "              " & CA.GetKey1 & "/               " & CostSearch.COST_RAW & "                 " & CostSearch.COST_EXP & "                 " & CostSearch.COST_PB
      Print #FileID, "---------------------------"
   End If
   
   CB.AddEditMode = SHOW_ADD
   CB.DOCUMENT_DATE = II.DOCUMENT_DATE                            '1
   CB.DOCUMENT_TYPE = II.DOCUMENT_TYPE                              '2
   CB.LOCATION_ID = II.LOCATION_ID                                                       '3
   CB.BATCH_ID = BatchID                                                                           '4
   CB.PART_ITEM_ID = II.PART_ITEM_ID                                                             '5
   CB.COST_RAW = 0                                                                                       '6
   CB.COST_EXP = II.TOTAL_INCLUDE_PRICE                                       '7
   CB.CUS_ID = II.CUS_ID                                       '8
   CB.ITEM_AMOUNT = II.IMPORT_AMOUNT                                                '9
      
   Set CostSearchEx = GetCostAccumSearch(m_CostAccumSearchExs, CB.GetKey2)
   If CostSearchEx Is Nothing Then
      Call m_CostAccumSearchExs.Add(CB, CB.GetKey2)
   Else
      CostSearchEx.COST_EXP = CostSearchEx.COST_EXP + CB.COST_EXP
      CostSearchEx.ITEM_AMOUNT = CostSearchEx.ITEM_AMOUNT + CB.ITEM_AMOUNT
   End If
   
   Set CA = Nothing
   Set CB = Nothing
   
End Sub

Private Sub GeneratePigImportMoveMentEx1(II As CImportItem)
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem

   Cm.AddEditMode = SHOW_ADD
   Cm.COMMIT_FLAG = II.COMMIT_FLAG
   Cm.DOCUMENT_NO = II.DOCUMENT_NO
   Cm.DOCUMENT_DATE = II.DOCUMENT_DATE
   Cm.IVD_ID = II.INVENTORY_DOC_ID
   Cm.DOCUMENT_CATEGORY = 1
   Cm.DOCUMENT_TYPE = II.DOCUMENT_TYPE
   Cm.TX_AMOUNT = II.IMPORT_AMOUNT
   Cm.TX_TYPE = II.TX_TYPE
   Cm.FROM_HOUSE_ID = II.LOCATION_ID
   Cm.TO_HOUSE_ID = 0
   Cm.PIG_ID = II.PART_ITEM_ID
   Cm.PIG_STATUS = 0
   Cm.TX_SEQ = II.TRANSACTION_SEQ
   Cm.REPLACE_FLAG = II.REPLACE_FLAG
   Call Cm.AddEditData

   Mi.AddEditMode = SHOW_ADD
   Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
   Mi.PART_ITEM_ID = 0
   Mi.EXPENSE_TYPE = II.EXPENSE_TYPE
   Mi.CAPITAL_AMOUNT = II.TOTAL_INCLUDE_PRICE
   Call Mi.AddEditData
   
   'ค่าใช้จ่าย
   Set S = New CMovementItemSearch1
   S.EXPENSE_TYPE = Mi.EXPENSE_TYPE
   S.PIG_ID = Cm.PIG_ID
   S.HOUSE_ID = Cm.FROM_HOUSE_ID
   S.CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
   Set TempSearch = GetMovementSearch1(m_MovementItemSearchs2, S.GetKey3)
   If TempSearch Is Nothing Then
      Call m_MovementItemSearchs2.Add(S, S.GetKey3)
   Else
      TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
   End If
   Set S = Nothing
   
   Call AddImportItem(II)
   
   Set Mi = Nothing
   Set Cm = Nothing
End Sub

Private Function GetPigBirthInMonth(FirstDate As Date, LastDate As Date, Optional BatchID As Long = -1) As Long
Static TempFirstDate As Date
Static TempLastDate As Date
Static TempCol As Collection

   If (FirstDate <> TempFirstDate) Or (LastDate <> TempLastDate) Then
      Set TempCol = New Collection
      
      Call LoadPigBirthAmount(Nothing, TempCol, FirstDate, LastDate, CommitTypeToFlag(cboCommitType.ItemData(Minus2Zero(cboCommitType.ListIndex))), , , BatchID)
      TempFirstDate = FirstDate
      TempLastDate = LastDate
   End If
   
   If TempCol.Count > 0 Then
      GetPigBirthInMonth = TempCol(1).IMPORT_AMOUNT
   Else
      GetPigBirthInMonth = 0
   End If
End Function

Private Function GetPigBirthInMonthInHouse(FirstDate As Date, LastDate As Date, HouseId As Long, PigID As Long, BatchID As Long) As Long
Static TempFirstDate As Date
Static TempLastDate As Date
Static TempCol As Collection
Static TempHouseID As Long
Static TempPigID As Long

   If (FirstDate <> TempFirstDate) Or (LastDate <> TempLastDate) Or (HouseId <> TempHouseID) Or (PigID <> TempPigID) Then
      Set TempCol = Nothing
      Set TempCol = New Collection

      Call LoadPigBirthAmount(Nothing, TempCol, FirstDate, LastDate, CommitTypeToFlag(cboCommitType.ItemData(Minus2Zero(cboCommitType.ListIndex))), HouseId, PigID, BatchID)
      TempFirstDate = FirstDate
      TempLastDate = LastDate
      TempHouseID = HouseId
      TempPigID = PigID
   End If
   
   If TempCol.Count > 0 Then
      GetPigBirthInMonthInHouse = TempCol(1).IMPORT_AMOUNT
   Else
      GetPigBirthInMonthInHouse = 0
   End If
End Function
Private Sub GeneratePigBirthMoveMent(II As CImportItem, BatchID As Long, Optional FileID As Long)
Static FirstDate As Date
Static LastDate As Date
Dim TempFromDate As Date
Dim TempToDate As Date
Dim CA As CCost_Accum
Dim Export As CExportItem
Dim PigBirthInMonth As Long
Dim PigBirthInMonthInHouse As Long

Dim CB As CCost_Accum

Dim CostSearch As CCost_Accum
Dim CostSearchEx As CCost_Accum

Static TempCol As Collection

   Set CA = New CCost_Accum
   Set CB = New CCost_Accum
   
   Call GetFirstLastDate(II.DOCUMENT_DATE, TempFromDate, TempToDate)
   
   
   
   If TempFromDate <> FirstDate Or TempToDate <> LastDate Then
      FirstDate = TempFromDate
      LastDate = TempToDate
      Set TempCol = Nothing
      Set TempCol = New Collection
      Call LoadPigParentUseAmountEx(Nothing, TempCol, FirstDate, LastDate, CommitTypeToFlag(cboCommitType.ItemData(Minus2Zero(cboCommitType.ListIndex))), BatchID)
      'PigBirthInMonth = GetPigBirthInMonth(FirstDate, LastDate, BatchId)
   End If
   'Set Export = GetExportItem(TempCol, Trim(II.PART_ITEM_ID & "-" & II.LOCATION_ID))
   Set Export = GetExportItem(TempCol, Trim(II.LOCATION_ID))
   
   'PigBirthInMonthInHouse = GetPigBirthInMonthInHouse(FirstDate, LastDate, II.LOCATION_ID, II.PART_ITEM_ID, BatchID)
   PigBirthInMonthInHouse = GetPigBirthInMonthInHouse(FirstDate, LastDate, II.LOCATION_ID, -1, BatchID)
   
   CA.AddEditMode = SHOW_ADD
   CA.DOCUMENT_DATE = II.DOCUMENT_DATE                            '1
   CA.DOCUMENT_TYPE = II.DOCUMENT_TYPE                              '2
   CA.LOCATION_ID = II.LOCATION_ID                                                       '3
   CA.BATCH_ID = BatchID                                                                           '4
   CA.PART_ITEM_ID = II.PART_ITEM_ID                                                             '5
   
   CA.CUS_ID = II.CUS_ID                                                                                                                           '8
   CA.DOCUMENT_CATEGORY = 1                                                                                                                           '8
   CA.COST_PB = (Export.EXPORT_TOTAL_PRICE / PigBirthInMonthInHouse) * II.IMPORT_AMOUNT
   
   Print #FileID, "---------------------------"
   Print #FileID, "ใบหมูเกิด" & "              " & CA.GetKey1 & "/               " & CA.COST_RAW & "                 " & CA.COST_EXP & "                 " & CA.COST_PB
   
   Set CostSearch = GetCostAccumSearch(m_CostAccumSearchs, CA.GetKey1)
   If CostSearch Is Nothing Then
      Call m_CostAccumSearchs.Add(CA, CA.GetKey1)
   Else
      CostSearch.COST_PB = CostSearch.COST_PB + CA.COST_PB
      Print #FileID, "ใบหมูเกิด" & "              " & CA.GetKey1 & "/               " & CostSearch.COST_RAW & "                 " & CostSearch.COST_EXP & "                 " & CostSearch.COST_PB
      Print #FileID, "---------------------------"
   End If
   
   CB.AddEditMode = SHOW_ADD
   CB.DOCUMENT_DATE = II.DOCUMENT_DATE                            '1
   CB.DOCUMENT_TYPE = II.DOCUMENT_TYPE                              '2
   CB.LOCATION_ID = II.LOCATION_ID                                                       '3
   CB.BATCH_ID = BatchID                                                                           '4
   CB.PART_ITEM_ID = II.PART_ITEM_ID                                                             '5
   
   CB.CUS_ID = II.CUS_ID                                                                                                                           '8
   CB.DOCUMENT_CATEGORY = 1                                                                                                                           '8
   CB.COST_PB = (Export.EXPORT_TOTAL_PRICE / PigBirthInMonthInHouse) * II.IMPORT_AMOUNT
   CB.ITEM_AMOUNT = II.IMPORT_AMOUNT
   
   Set CostSearchEx = GetCostAccumSearch(m_CostAccumSearchExs, CB.GetKey2)
   If CostSearchEx Is Nothing Then
      Call m_CostAccumSearchExs.Add(CB, CB.GetKey2)
   Else
      CostSearchEx.COST_PB = CostSearchEx.COST_PB + CB.COST_PB
      CostSearchEx.ITEM_AMOUNT = CostSearchEx.ITEM_AMOUNT + CB.ITEM_AMOUNT
   End If
   
   Set CA = Nothing
   Set CB = Nothing
End Sub

Private Sub GeneratePigBirthMoveMentEx(II As CImportItem, BatchID As Long)
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim FirstDate As Date
Dim LastDate As Date
Dim FirstDateStr As String
Dim LastDateStr As String
Static TempFirstDateStr As String
Static TempLastDateStr As String
Static TempHouseID As Long
Static TempPigID As Long
Static TempCol As Collection
Dim EI As CExportItem
Dim TempII As CImportItem
Dim PigBirthInMonth As Long
Dim PigBirthInMonthInHouse As Long
Dim S As CMovementItemSearch1
Dim Ms As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1
Dim iCount As Long
Static TempCol2 As Collection
Dim Er As CExpenseRatio
Dim PigBalance As Long
Static HousePigBirths As Collection
Static ImportPigBals As Collection
Static ExportPigBals As Collection
Static ImportPigParents As Collection
Static ExportPigParents As Collection
Dim NewDate As Date
Dim PigBirthInHouse1 As Long
Dim PigBalInHouse1 As Long
Dim PigParentInHouse1 As Long

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   Set Ms = New CMovementItemSearch1

   Call GetFirstLastDate(II.DOCUMENT_DATE, FirstDate, LastDate)
   FirstDateStr = DateToStringInt(FirstDate)
   LastDateStr = DateToStringInt(LastDate)

   'ถ้ามีใบเกิดก่อนหน้านั้นที่โรงเรือนเดียวกัน เดือนเดียวกัน ก็ไม่ต้องทำอะไร
   Set TempII = New CImportItem

   TempII.IMPORT_ITEM_ID = -1
   TempII.FROM_DATE = FirstDate
   TempII.TO_DATE = LastDate
   TempII.TO_TX_SEQ = II.TRANSACTION_SEQ - 1
   TempII.FROM_TX_SEQ = -1
   TempII.LOCATION_ID = II.LOCATION_ID
   TempII.PART_ITEM_ID = II.PART_ITEM_ID
   TempII.DOCUMENT_TYPE = 5
   TempII.BATCH_ID = BatchID
   Call TempII.QueryData(1, m_Rs, iCount)
   If Not m_Rs.EOF Then
      Set TempII = Nothing
      Exit Sub
   End If

   Set TempII = Nothing

   If (TempFirstDateStr <> FirstDateStr) Or (TempLastDateStr <> LastDateStr) Then
      TempFirstDateStr = FirstDateStr
      TempLastDateStr = LastDateStr
      TempHouseID = 0

      Set TempCol = Nothing
      Set TempCol = New Collection
      Call LoadPigParentUseAmountEx3(Nothing, TempCol, FirstDate, LastDate, CommitTypeToFlag(cboCommitType.ItemData(Minus2Zero(cboCommitType.ListIndex))), BatchID)
      
      Set HousePigBirths = Nothing
      Set HousePigBirths = New Collection
      Call LoadHousePigBirthAmount(Nothing, HousePigBirths, FirstDate, LastDate, , , , BatchID)
      
      NewDate = DateAdd("D", -1, FirstDate)
      Set ImportPigBals = Nothing
      Set ImportPigBals = New Collection
      Set ExportPigBals = Nothing
      Set ExportPigBals = New Collection
      If NewDate > 0 Then
         Call LoadPigImportByHouse(Nothing, ImportPigBals, -1, NewDate, , , , , BatchID)
         Call LoadPigExportByHouse(Nothing, ExportPigBals, -1, NewDate, , , , , , BatchID)
      End If
   End If

   Ms.HOUSE_ID = II.LOCATION_ID
   Ms.PIG_ID = II.PART_ITEM_ID
   Set TempSearch = GetMovementSearch1(m_TempSearchs1, Ms.GetKey2)
   If TempSearch Is Nothing Then
      Call m_TempSearchs1.Add(Ms, Ms.GetKey2)

      PigBirthInMonth = GetPigBirthInMonth(FirstDate, LastDate)
      PigBirthInMonthInHouse = GetPigBirthInMonthInHouse(FirstDate, LastDate, II.LOCATION_ID, II.PART_ITEM_ID, BatchID)

      'ให้ใบแรกเป็นตัวแทนของสุกรคลอดทั้งเดือน
      Cm.AddEditMode = SHOW_ADD
      Cm.COMMIT_FLAG = II.COMMIT_FLAG
      Cm.DOCUMENT_NO = II.DOCUMENT_NO
      Cm.DOCUMENT_DATE = II.DOCUMENT_DATE
      Cm.IVD_ID = II.INVENTORY_DOC_ID
      Cm.DOCUMENT_CATEGORY = 1
      Cm.DOCUMENT_TYPE = II.DOCUMENT_TYPE
      Cm.TX_AMOUNT = PigBirthInMonthInHouse
      Cm.TX_TYPE = II.TX_TYPE
      Cm.FROM_HOUSE_ID = II.LOCATION_ID
      Cm.TO_HOUSE_ID = 0
      Cm.PIG_ID = II.PART_ITEM_ID
      Cm.PIG_STATUS = 0
      Cm.IMPORT_ITEM_ID = II.IMPORT_ITEM_ID
      Cm.EXPORT_ITEM_ID = 0
      Cm.TX_SEQ = II.TRANSACTION_SEQ
      Cm.REPLACE_FLAG = II.REPLACE_FLAG
      Cm.BATCH_ID = BatchID
      Call Cm.AddEditData

      For Each EI In TempCol
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.PART_ITEM_ID = EI.PART_ITEM_ID
         Mi.EXPENSE_TYPE = 0
         Mi.CAPITAL_AMOUNT = MyDiffEx(EI.EXPORT_TOTAL_PRICE, PigBirthInMonth) * PigBirthInMonthInHouse
         Call Mi.AddEditData

         'วัตถุดิบ
         Set S = New CMovementItemSearch1
         S.EXPENSE_TYPE = 0
         S.PART_ITEM_ID = Mi.PART_ITEM_ID
         S.PIG_ID = Cm.PIG_ID
         S.HOUSE_ID = Cm.FROM_HOUSE_ID
         S.CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
         Set TempSearch = GetMovementSearch1(m_MovementItemSearchs1, S.GetKey1)
         If TempSearch Is Nothing Then
            Call m_MovementItemSearchs1.Add(S, S.GetKey1)
         Else
            TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
         End If
         Set S = Nothing
      Next EI
   End If

   Set Ms = Nothing
   Set Mi = Nothing
   Set Cm = Nothing
End Sub
Private Sub GenerateExpenseMovement(Ri As CROItem, BatchID As Long, Optional FileID As Long)
Dim CA As CCost_Accum
Dim TempCA As CCost_Accum

Dim CB As CCost_Accum
Dim TempCB As CCost_Accum
Dim SearchCB As CCost_Accum

Dim AvgExpensePerUnit As Double
Dim SumTotalPig As Long
   
   Set CA = New CCost_Accum
   Set CB = New CCost_Accum
   
   SumTotalPig = 0
   For Each SearchCB In m_CostAccumSearchExs
      SumTotalPig = SumTotalPig + SearchCB.ITEM_AMOUNT                     ' จำนวนหมู ขณะที่เข้ามาใน FuncTion นี้
   Next SearchCB
   
   If SumTotalPig <> 0 Then
      AvgExpensePerUnit = Ri.TOTAL_PRICE / SumTotalPig
   End If
   
   For Each SearchCB In m_CostAccumSearchExs
      Set CA = New CCost_Accum
      Set CB = New CCost_Accum
      CB.LOCATION_ID = SearchCB.LOCATION_ID
      CB.PART_ITEM_ID = SearchCB.PART_ITEM_ID
      CB.COST_EXP = SearchCB.ITEM_AMOUNT * AvgExpensePerUnit
      Set TempCB = GetCostAccumSearch(m_CostAccumSearchExs, CB.GetKey2)
      If TempCB Is Nothing Then
         Call m_CostAccumSearchExs.Add(CB, CB.GetKey2)
      Else
         TempCB.COST_EXP = TempCB.COST_EXP + CB.COST_EXP
      End If
      
      CA.LOCATION_ID = SearchCB.LOCATION_ID
      CA.PART_ITEM_ID = SearchCB.PART_ITEM_ID
      CA.DOCUMENT_DATE = Ri.DOCUMENT_DATE
      CA.DOCUMENT_TYPE = Ri.DOCUMENT_TYPE
      CA.DOCUMENT_CATEGORY = 2
      CA.COST_EXP = SearchCB.ITEM_AMOUNT * AvgExpensePerUnit                              'เท่ากันรึเปล่า ระหว่าง coll1 กับ coll2
      CA.BATCH_ID = BatchID
      
      Print #FileID, "---------------------------"
      Print #FileID, "ใบค่าใช้จ่าย" & "              " & CA.GetKey1 & "/               " & CA.COST_RAW & "                 " & CA.COST_EXP & "                 " & CA.COST_PB
   
      Set TempCA = GetCostAccumSearch(m_CostAccumSearchs, CA.GetKey1)
      If TempCA Is Nothing Then
         Call m_CostAccumSearchs.Add(CA, CA.GetKey1)
         Print #FileID, "---------------------------"
      Else
         TempCA.COST_EXP = TempCA.COST_EXP + CA.COST_EXP
         Print #FileID, "ใบค่าใช้จ่าย" & "              " & CA.GetKey1 & "/              " & TempCA.COST_RAW & "                 " & TempCA.COST_EXP & "                 " & TempCA.COST_PB
         Print #FileID, "---------------------------"
      End If
      
   Next SearchCB
      
   Set CA = Nothing
   Set CB = Nothing

End Sub

Private Function GetPreviousAmountEx(O As Object) As Double
Dim EI As CExportItem
Dim II As CImportItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   Set TempRs = New ADODB.Recordset
   Set EI = New CExportItem
   Set II = New CImportItem

   EI.EXPORT_ITEM_ID = -1
   EI.PART_ITEM_ID1 = O.PART_ITEM_ID
   EI.LOCATION_ID1 = O.LOCATION_ID
   EI.TRANSACTION_SEQ = O.TRANSACTION_SEQ
   Call EI.QueryData(28, TempRs, iCount)
   If Not TempRs.EOF Then
      Call EI.PopulateFromRS(28, TempRs)
   Else
      EI.TRANSACTION_SEQ = 0
   End If
   
   II.IMPORT_ITEM_ID = -1
   II.PART_ITEM_ID1 = O.PART_ITEM_ID
   II.LOCATION_ID1 = O.LOCATION_ID
   II.TRANSACTION_SEQ = O.TRANSACTION_SEQ
   Call II.QueryData(17, TempRs, iCount)
   If Not TempRs.EOF Then
      Call II.PopulateFromRS(17, TempRs)
   Else
      EI.TRANSACTION_SEQ = 0
   End If
   
   If II.TRANSACTION_SEQ > EI.TRANSACTION_SEQ Then
      GetPreviousAmountEx = II.CURRENT_AMOUNT
   Else
      GetPreviousAmountEx = EI.CURRENT_AMOUNT
   End If
   
   Set II = Nothing
   Set EI = Nothing
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Function
Private Function GetPreviousAmount(O As Object) As Double
On Error Resume Next
Dim II As CImportItem
Dim Key As String
Dim Amt As Double
Dim Ba As CBalanceAccum
Dim iCount As Long

   If O.TX_TYPE = "I" Then
      Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID
   ElseIf O.TX_TYPE = "E" Then
      Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID
   Else
      Key = ""
   End If
   Set II = m_MovementItemSearchs3(Key)
   Amt = 0
   If II Is Nothing Then
      GetPreviousAmount = Amt
   Else
      GetPreviousAmount = Minus2Zero(II.CURRENT_AMOUNT)
   End If
End Function

Private Sub GeneratePigTransferMovementExp(EI As CExportItem, II As CImportItem, BatchID As Long, Optional FileID As Long)
Dim PigCount As Double
Dim CA As CCost_Accum
Dim CB As CCost_Accum
Dim S As CCost_Accum
Dim TempSearch As CCost_Accum
Dim TempSearchEx As CCost_Accum
Dim ExId As CExportId
Dim AmtRaw As Double
Dim AmtExp As Double
Dim AmtPb As Double
Dim TempAmtRaw As Double
Dim TempAmtExp As Double
Dim TempAmtPb As Double
Dim SearchCB As CCost_Accum

    'PigCount = GetPreviousAmount(EI)
   PigCount = 0
'  For Each SearchCB In m_CostAccumSearchExs
'      PigCount = PigCount + SearchCB.ITEM_AMOUNT                     ' จำนวนหมู ขณะที่เข้ามาใน FuncTion นี้
'   Next SearchCB

   Set SearchCB = Nothing
   
   Set CA = New CCost_Accum
   Set CB = New CCost_Accum
   Set ExId = New CExportId
   
   CA.AddEditMode = SHOW_ADD
   CA.DOCUMENT_DATE = EI.DOCUMENT_DATE
   CA.DOCUMENT_TYPE = EI.DOCUMENT_TYPE
   CA.LOCATION_ID = EI.LOCATION_ID
   CA.PART_ITEM_ID = EI.PART_ITEM_ID
   CA.BATCH_ID = BatchID
   CA.CUS_ID = EI.CUS_ID
   CA.DOCUMENT_CATEGORY = 1
   
   TempAmtRaw = 0
   TempAmtExp = 0
   TempAmtPb = 0
   
   Set S = GetCostAccumSearch(m_CostAccumSearchExs, CA.GetKey2)
   If Not (S Is Nothing) Then
      TempAmtRaw = S.COST_RAW
      TempAmtExp = S.COST_EXP
      TempAmtPb = S.COST_PB
      PigCount = S.ITEM_AMOUNT
   End If
   'PigCount จำนวนหมู ใน Table Balance Accum โดย เลือกจากวันที่                                       ' อาจจะต้องทำใหม่
   AmtRaw = (MyDiffEx(TempAmtRaw, PigCount)) * EI.EXPORT_AMOUNT                  'ราคาเฉลี่ยที่นำออกมาของอาหารหมู
   AmtExp = (MyDiffEx(TempAmtExp, PigCount)) * EI.EXPORT_AMOUNT                     'ราคาเฉลี่ยที่นำออกมาของค่าใช้จ่าย
   AmtPb = (MyDiffEx(TempAmtPb, PigCount)) * EI.EXPORT_AMOUNT                     'ราคาเฉลี่ยที่นำออกมาของค่าใช้จ่าย
   CA.COST_RAW = (-1) * AmtRaw
   CA.COST_EXP = (-1) * AmtExp
   CA.COST_PB = (-1) * AmtPb
   CB.COST_RAW = (-1) * AmtRaw
   CB.COST_EXP = (-1) * AmtExp
   CB.COST_PB = (-1) * AmtPb
      
   Print #FileID, "---------------------------"
   Print #FileID, "ใบโอนออก" & "              " & CA.GetKey1 & "/               " & CA.COST_RAW & "                 " & CA.COST_EXP & "                 " & CA.COST_PB
   
   Set TempSearch = GetCostAccumSearch(m_CostAccumSearchs, CA.GetKey1)
   If TempSearch Is Nothing Then
      Call m_CostAccumSearchs.Add(CA, CA.GetKey1)
   Else
      TempSearch.COST_RAW = TempSearch.COST_RAW + CA.COST_RAW
      TempSearch.COST_EXP = TempSearch.COST_EXP + CA.COST_EXP
      TempSearch.COST_PB = TempSearch.COST_PB + CA.COST_PB
      
      Print #FileID, "ใบโอนออก" & "              " & CA.GetKey1 & "/               " & TempSearch.COST_RAW & "                 " & TempSearch.COST_EXP & "                 " & TempSearch.COST_PB
      Print #FileID, "---------------------------"

   End If
   
   CB.AddEditMode = SHOW_ADD
   CB.DOCUMENT_DATE = EI.DOCUMENT_DATE
   CB.DOCUMENT_TYPE = EI.DOCUMENT_TYPE
   CB.LOCATION_ID = EI.LOCATION_ID
   CB.PART_ITEM_ID = EI.PART_ITEM_ID
   CB.BATCH_ID = BatchID
   CB.CUS_ID = EI.CUS_ID
   CB.ITEM_AMOUNT = (-1) * EI.EXPORT_AMOUNT
   
   Set TempSearchEx = GetCostAccumSearch(m_CostAccumSearchExs, CB.GetKey2)
   If TempSearchEx Is Nothing Then
      Call m_CostAccumSearchExs.Add(CB, CB.GetKey2)
   Else
      TempSearchEx.COST_RAW = TempSearchEx.COST_RAW + CB.COST_RAW
      TempSearchEx.COST_EXP = TempSearchEx.COST_EXP + CB.COST_EXP
      TempSearchEx.COST_PB = TempSearchEx.COST_PB + CB.COST_PB
      TempSearchEx.ITEM_AMOUNT = TempSearchEx.ITEM_AMOUNT + CB.ITEM_AMOUNT
   End If
   
   ExId.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID
   ExId.COST_RAW = Abs(CA.COST_RAW)
   ExId.COST_EXP = Abs(CA.COST_EXP)
   ExId.COST_PB = Abs(CA.COST_PB)
   Call m_ExportIDs.Add(ExId, ExId.GetKey1)
   
   Set ExId = Nothing
   Set CA = Nothing
   Set CB = Nothing
   
End Sub
Private Sub GeneratePigStatusChangeExp(EI As CExportItem, II As CImportItem, BatchID As Long)
Dim PigCount As Double
Dim Cm As CCapitalMovement
Dim Cl As CCapitalLoss
Dim Mi As CMovementItem
Dim Li As CLossItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1
Dim Ps As CProductStatus
Dim Amt As Double
Dim TempAmt As Double

   PigCount = GetPreviousAmount(EI)

   Set Cl = New CCapitalLoss
   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   Set Li = New CLossItem
   Set S = New CMovementItemSearch1

   Cm.AddEditMode = SHOW_ADD
   Cm.COMMIT_FLAG = EI.COMMIT_FLAG
   Cm.DOCUMENT_NO = EI.DOCUMENT_NO
   Cm.DOCUMENT_DATE = EI.DOCUMENT_DATE
   Cm.IVD_ID = EI.INVENTORY_DOC_ID
   Cm.DOCUMENT_CATEGORY = 1
   Cm.DOCUMENT_TYPE = EI.DOCUMENT_TYPE
   Cm.TX_AMOUNT = EI.EXPORT_AMOUNT
   Cm.TX_TYPE = EI.TX_TYPE
   Cm.FROM_HOUSE_ID = EI.LOCATION_ID
   Cm.TO_HOUSE_ID = II.LOCATION_ID
   Cm.PIG_ID = EI.PART_ITEM_ID
   Cm.TO_PIG_ID = II.PART_ITEM_ID
   Cm.PIG_STATUS = Minus2Zero(EI.PIG_STATUS)
   Cm.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID
   Cm.IMPORT_ITEM_ID = II.IMPORT_ITEM_ID
   Cm.TX_SEQ = EI.TRANSACTION_SEQ
   Cm.REPLACE_FLAG = EI.REPLACE_FLAG
   Cm.BATCH_ID = BatchID
   Call Cm.AddEditData

   Cl.AddEditMode = SHOW_ADD
   Call Cl.CopyFromCapitalMovenent(Cm)
   Call Cl.AddEditData
   
   If EI.PIG_STATUS > 0 Then
      Set Ps = m_ProductStatuss(Trim(Str(EI.PIG_STATUS)))
   Else
      Set Ps = New CProductStatus
   End If

   TempAmt = 0
   For Each S In m_MovementItemSearchs1
      If (S.PIG_ID = EI.PART_ITEM_ID) And (S.HOUSE_ID = EI.LOCATION_ID) Then
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.PART_ITEM_ID = S.PART_ITEM_ID
         Mi.EXPENSE_TYPE = 0

         If EI.PIG_STATUS > 0 Then
            'โอนเข้าเรือนขาย
            If Ps.CAPITAL_MOVE_FLAG = "Y" Then
               Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
            Else
               Amt = 0
            End If
         Else
            Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
         End If

         Mi.CAPITAL_AMOUNT = -1 * Amt
         Call Mi.AddEditData

         Li.AddEditMode = SHOW_ADD
         Li.CAPITAL_LOSS_ID = Cl.CAPITAL_LOSS_ID
         Call Li.CopyFromMovementItem(Mi)
         Call Li.AddEditData

         S.CAPITAL_AMOUNT = S.CAPITAL_AMOUNT + (-1 * Amt)
      End If
   Next S

   'ค่าใช้จ่าย
   For Each S In m_MovementItemSearchs2
      If (S.PIG_ID = EI.PART_ITEM_ID) And (S.HOUSE_ID = EI.LOCATION_ID) Then
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.EXPENSE_TYPE = S.EXPENSE_TYPE
         Mi.PART_ITEM_ID = 0
         Mi.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID
         
         If EI.PIG_STATUS > 0 Then
            'โอนเข้าเรือนขาย
            If Ps.CAPITAL_MOVE_FLAG = "Y" Then
               Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
            Else
               Amt = 0
            End If
         Else
            Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
         End If

         Mi.CAPITAL_AMOUNT = -1 * Amt
         Call Mi.AddEditData

         Li.AddEditMode = SHOW_ADD
         Li.CAPITAL_LOSS_ID = Cl.CAPITAL_LOSS_ID
         Call Li.CopyFromMovementItem(Mi)
         Call Li.AddEditData

         S.CAPITAL_AMOUNT = S.CAPITAL_AMOUNT + (-1 * Amt)
      End If
   Next S

   Set Ps = Nothing
   Set S = Nothing
   Set Mi = Nothing
   Set Cm = Nothing
   Set Li = Nothing
   Set Cl = Nothing
End Sub

Private Sub GeneratePigTransferMovementExpEx1(EI As CExportItem, II As CImportItem)
Dim PigCount As Double
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1
Dim Ps As CProductStatus
Dim Amt As Double
Dim TempAmt As Double

   PigCount = GetPreviousAmount(EI)

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   Set S = New CMovementItemSearch1

   Cm.AddEditMode = SHOW_ADD
   Cm.COMMIT_FLAG = EI.COMMIT_FLAG
   Cm.DOCUMENT_NO = EI.DOCUMENT_NO
   Cm.DOCUMENT_DATE = EI.DOCUMENT_DATE
   Cm.IVD_ID = EI.INVENTORY_DOC_ID
   Cm.DOCUMENT_CATEGORY = 1
   Cm.DOCUMENT_TYPE = EI.DOCUMENT_TYPE
   Cm.TX_AMOUNT = EI.EXPORT_AMOUNT
   Cm.TX_TYPE = EI.TX_TYPE
   Cm.FROM_HOUSE_ID = EI.LOCATION_ID
   Cm.TO_HOUSE_ID = II.LOCATION_ID
   Cm.PIG_ID = EI.PART_ITEM_ID
   Cm.PIG_STATUS = Minus2Zero(EI.PIG_STATUS)
   Cm.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID
   Cm.IMPORT_ITEM_ID = II.IMPORT_ITEM_ID
   Cm.TX_SEQ = EI.TRANSACTION_SEQ
   Cm.REPLACE_FLAG = EI.REPLACE_FLAG
   Call Cm.AddEditData

   If EI.PIG_STATUS > 0 Then
      Set Ps = m_ProductStatuss(Trim(Str(EI.PIG_STATUS)))
   Else
      Set Ps = New CProductStatus
   End If

   TempAmt = 0
   For Each S In m_MovementItemSearchs1
      If (S.PIG_ID = EI.PART_ITEM_ID) And (S.HOUSE_ID = EI.LOCATION_ID) Then
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.PART_ITEM_ID = S.PART_ITEM_ID
         Mi.EXPENSE_TYPE = 0
         
         If EI.PIG_STATUS > 0 Then
            'โอนเข้าเรือนขาย
            If Ps.CAPITAL_MOVE_FLAG = "Y" Then
               Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
            Else
               Amt = 0
            End If
         Else
            Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
         End If

         Mi.CAPITAL_AMOUNT = -1 * Amt
         Call Mi.AddEditData

         S.CAPITAL_AMOUNT = S.CAPITAL_AMOUNT + (-1 * Amt)
      End If
   Next S

   'ค่าใช้จ่าย
   For Each S In m_MovementItemSearchs2
      If (S.PIG_ID = EI.PART_ITEM_ID) And (S.HOUSE_ID = EI.LOCATION_ID) Then
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.EXPENSE_TYPE = S.EXPENSE_TYPE
         Mi.PART_ITEM_ID = 0
         Mi.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID
         
         If EI.PIG_STATUS > 0 Then
            'โอนเข้าเรือนขาย
            If Ps.CAPITAL_MOVE_FLAG = "Y" Then
               Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
            Else
               Amt = 0
            End If
         Else
            Amt = (MyDiffEx(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
         End If
         
         Mi.CAPITAL_AMOUNT = -1 * Amt
         Call Mi.AddEditData

         S.CAPITAL_AMOUNT = S.CAPITAL_AMOUNT + (-1 * Amt)
      End If
   Next S

   Call AddExportItem(EI)

   Set Ps = Nothing
   Set S = Nothing
   Set Mi = Nothing
   Set Cm = Nothing
End Sub

Private Sub GeneratePigStatusChangeImp(II As CImportItem, FromEi As CExportItem, BatchID As Long)
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1

   Set TempRs = New ADODB.Recordset
   Set Mi = New CMovementItem
   Set Cm = New CCapitalMovement
   Set S = New CMovementItemSearch1
   Set TempSearch = New CMovementItemSearch1

   Cm.AddEditMode = SHOW_ADD
   Cm.COMMIT_FLAG = II.COMMIT_FLAG
   Cm.DOCUMENT_NO = II.DOCUMENT_NO
   Cm.DOCUMENT_DATE = II.DOCUMENT_DATE
   Cm.IVD_ID = II.INVENTORY_DOC_ID
   Cm.DOCUMENT_CATEGORY = 1
   Cm.DOCUMENT_TYPE = II.DOCUMENT_TYPE
   Cm.TX_AMOUNT = II.IMPORT_AMOUNT
   Cm.TX_TYPE = II.TX_TYPE
   Cm.FROM_HOUSE_ID = II.LOCATION_ID
   Cm.TO_HOUSE_ID = FromEi.LOCATION_ID 'บอกว่ามาจากโรงเรือนไหน
   Cm.PIG_ID = II.PART_ITEM_ID
   Cm.TO_PIG_ID = FromEi.PART_ITEM_ID
   If II.DOCUMENT_TYPE = 12 Then 'ใบเปลี่ยนสถานะสุกร ให้นำสถานะสุกรมาจากตัวของมันเอง
      Cm.PIG_STATUS = Minus2Zero(II.PIG_STATUS)
   Else
      Cm.PIG_STATUS = Minus2Zero(FromEi.PIG_STATUS)
   End If
   Cm.IMPORT_ITEM_ID = II.IMPORT_ITEM_ID
   Cm.EXPORT_ITEM_ID = FromEi.EXPORT_ITEM_ID
   Cm.TX_SEQ = II.TRANSACTION_SEQ
   Cm.REPLACE_FLAG = II.REPLACE_FLAG
   Cm.BATCH_ID = BatchID
   Call Cm.AddEditData

   Mi.MOVEMENT_ITEM_ID = -1
   Mi.EXPORT_ITEM_ID = FromEi.EXPORT_ITEM_ID
   Call Mi.QueryData(1, TempRs, iCount)

   While Not TempRs.EOF
      Call Mi.PopulateFromRS(1, TempRs)

      Mi.AddEditMode = SHOW_ADD
      Mi.CAPITAL_AMOUNT = 0  'Abs(Mi.CAPITAL_AMOUNT) ไม่มีต้นทุนแล้ว
      Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
      Call Mi.AddEditData

      If Mi.EXPENSE_TYPE > 0 Then 'ค่าใช้จ่าย
         'ค่าใช้จ่าย
         Set S = New CMovementItemSearch1
         S.EXPENSE_TYPE = Mi.EXPENSE_TYPE
         S.PIG_ID = Cm.PIG_ID
         S.HOUSE_ID = Cm.FROM_HOUSE_ID
         S.CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
         Set TempSearch = GetMovementSearch1(m_MovementItemSearchs2, S.GetKey3)
         If TempSearch Is Nothing Then
            Call m_MovementItemSearchs2.Add(S, S.GetKey3)
         Else
            TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
         End If
         Set S = Nothing
      ElseIf Mi.PART_ITEM_ID > 0 Then 'อุปกรณ์ คงคลัง
         'วัตถุดิบ
         Set S = New CMovementItemSearch1
         S.PART_ITEM_ID = Mi.PART_ITEM_ID
         S.PIG_ID = Cm.PIG_ID
         S.HOUSE_ID = Cm.FROM_HOUSE_ID
         S.CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
         Set TempSearch = GetMovementSearch1(m_MovementItemSearchs1, S.GetKey1)
         If TempSearch Is Nothing Then
            Call m_MovementItemSearchs1.Add(S, S.GetKey1)
'Call S.PrintDebug(Cm.DOCUMENT_NO, "GeneratePigTransferMovementImp1", Ii.IMPORT_AMOUNT, Ii.TX_TYPE)
         Else
            TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
'Call TempSearch.PrintDebug(Cm.DOCUMENT_NO, "GeneratePigTransferMovementImp2", Ii.IMPORT_AMOUNT, Ii.TX_TYPE)
         End If
         Set S = Nothing
      End If

      TempRs.MoveNext
   Wend

   Set TempSearch = Nothing
   Set S = Nothing
   Set Cm = Nothing
   Set Mi = Nothing
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GeneratePigTransferMovementImp(II As CImportItem, FromEi As CExportItem, BatchID As Long, Optional FileID As Long)
Dim CA As CCost_Accum
Dim CB As CCost_Accum
Dim iCount As Long

Dim CostSearch As CCost_Accum
Dim CostSearchEx As CCost_Accum
Dim ExId As CExportId
Dim TempExId As CExportId
      
   Set CA = New CCost_Accum
   Set CB = New CCost_Accum
   Set TempExId = New CExportId
   
   CA.AddEditMode = SHOW_ADD
   CA.DOCUMENT_DATE = II.DOCUMENT_DATE
   CA.DOCUMENT_TYPE = II.DOCUMENT_TYPE
   CA.LOCATION_ID = II.LOCATION_ID
   CA.PART_ITEM_ID = II.PART_ITEM_ID
   CA.BATCH_ID = BatchID
   CA.CUS_ID = II.CUS_ID
   CA.DOCUMENT_CATEGORY = 1
   
   TempExId.EXPORT_ITEM_ID = FromEi.EXPORT_ITEM_ID
   Set ExId = GetCExportId(m_ExportIDs, TempExId.GetKey1)
   If Not (ExId Is Nothing) Then
      CA.COST_RAW = ExId.COST_RAW
      CA.COST_EXP = ExId.COST_EXP
      CA.COST_PB = ExId.COST_PB
      CB.COST_RAW = ExId.COST_RAW
      CB.COST_EXP = ExId.COST_EXP
      CB.COST_PB = ExId.COST_PB
   End If
   
   Print #FileID, "---------------------------"
   Print #FileID, "ใบโอนเข้า" & "              " & CA.GetKey1 & "/               " & CA.COST_RAW & "                 " & CA.COST_EXP & "                 " & CA.COST_PB
   
   Set CostSearch = GetCostAccumSearch(m_CostAccumSearchs, CA.GetKey1)
   If CostSearch Is Nothing Then
      Call m_CostAccumSearchs.Add(CA, CA.GetKey1)
   Else
      CostSearch.COST_RAW = CostSearch.COST_RAW + CA.COST_RAW
      CostSearch.COST_EXP = CostSearch.COST_EXP + CA.COST_EXP
      CostSearch.COST_PB = CostSearch.COST_PB + CA.COST_PB
      Print #FileID, "ใบโอนเข้า" & "              " & CA.GetKey1 & "/               " & CostSearch.COST_RAW & "                 " & CostSearch.COST_EXP & "                 " & CostSearch.COST_PB
      Print #FileID, "---------------------------"
   End If
   
   CB.AddEditMode = SHOW_ADD
   CB.DOCUMENT_DATE = II.DOCUMENT_DATE
   CB.DOCUMENT_TYPE = II.DOCUMENT_TYPE
   CB.LOCATION_ID = II.LOCATION_ID
   CB.PART_ITEM_ID = II.PART_ITEM_ID
   CB.BATCH_ID = BatchID
   CB.CUS_ID = II.CUS_ID
   CB.ITEM_AMOUNT = II.IMPORT_AMOUNT
   
   Set CostSearchEx = GetCostAccumSearch(m_CostAccumSearchExs, CB.GetKey2)
   If CostSearchEx Is Nothing Then
      Call m_CostAccumSearchExs.Add(CB, CB.GetKey2)
   Else
      CostSearchEx.COST_RAW = CostSearchEx.COST_RAW + CB.COST_RAW
      CostSearchEx.COST_EXP = CostSearchEx.COST_EXP + CB.COST_EXP
      CostSearchEx.COST_PB = CostSearchEx.COST_PB + CB.COST_PB
      CostSearchEx.ITEM_AMOUNT = CostSearchEx.ITEM_AMOUNT + CB.ITEM_AMOUNT
   End If
   
      
   Set TempExId = Nothing
   Set CA = Nothing
   Set CB = Nothing
   
End Sub

Private Sub GeneratePigTransferMovementImpEx1(II As CImportItem, FromEi As CExportItem)
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1

   Set TempRs = New ADODB.Recordset
   Set Mi = New CMovementItem
   Set Cm = New CCapitalMovement
   Set S = New CMovementItemSearch1
   Set TempSearch = New CMovementItemSearch1

   Cm.AddEditMode = SHOW_ADD
   Cm.COMMIT_FLAG = II.COMMIT_FLAG
   Cm.DOCUMENT_NO = II.DOCUMENT_NO
   Cm.DOCUMENT_DATE = II.DOCUMENT_DATE
   Cm.IVD_ID = II.INVENTORY_DOC_ID
   Cm.DOCUMENT_CATEGORY = 1
   Cm.DOCUMENT_TYPE = II.DOCUMENT_TYPE
   Cm.TX_AMOUNT = II.IMPORT_AMOUNT
   Cm.TX_TYPE = II.TX_TYPE
   Cm.FROM_HOUSE_ID = II.LOCATION_ID
   Cm.TO_HOUSE_ID = FromEi.LOCATION_ID 'บอกว่ามาจากโรงเรือนไหน
   Cm.PIG_ID = II.PART_ITEM_ID
   If II.DOCUMENT_TYPE = 12 Then 'ใบเปลี่ยนสถานะสุกร ให้นำสถานะสุกรมาจากตัวของมันเอง
      Cm.PIG_STATUS = Minus2Zero(II.PIG_STATUS)
   Else
      Cm.PIG_STATUS = Minus2Zero(FromEi.PIG_STATUS)
   End If
   Cm.IMPORT_ITEM_ID = II.IMPORT_ITEM_ID
   Cm.EXPORT_ITEM_ID = FromEi.EXPORT_ITEM_ID
   Cm.TX_SEQ = II.TRANSACTION_SEQ
   Cm.REPLACE_FLAG = II.REPLACE_FLAG
   Call Cm.AddEditData

   Mi.MOVEMENT_ITEM_ID = -1
   Mi.EXPORT_ITEM_ID = FromEi.EXPORT_ITEM_ID
   Call Mi.QueryData(1, TempRs, iCount)

   While Not TempRs.EOF
      Call Mi.PopulateFromRS(1, TempRs)

      Mi.AddEditMode = SHOW_ADD
      Mi.CAPITAL_AMOUNT = Abs(Mi.CAPITAL_AMOUNT)
      Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
      Call Mi.AddEditData

      If Mi.EXPENSE_TYPE > 0 Then 'ค่าใช้จ่าย
         'ค่าใช้จ่าย
         Set S = New CMovementItemSearch1
         S.EXPENSE_TYPE = Mi.EXPENSE_TYPE
         S.PIG_ID = Cm.PIG_ID
         S.HOUSE_ID = Cm.FROM_HOUSE_ID
         S.CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
         Set TempSearch = GetMovementSearch1(m_MovementItemSearchs2, S.GetKey3)
         If TempSearch Is Nothing Then
            Call m_MovementItemSearchs2.Add(S, S.GetKey3)
         Else
            TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
         End If
         Set S = Nothing
      ElseIf Mi.PART_ITEM_ID > 0 Then 'อุปกรณ์ คงคลัง
         'วัตถุดิบ
         Set S = New CMovementItemSearch1
         S.PART_ITEM_ID = Mi.PART_ITEM_ID
         S.PIG_ID = Cm.PIG_ID
         S.HOUSE_ID = Cm.FROM_HOUSE_ID
         S.CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
         Set TempSearch = GetMovementSearch1(m_MovementItemSearchs1, S.GetKey1)
         If TempSearch Is Nothing Then
            Call m_MovementItemSearchs1.Add(S, S.GetKey1)
         Else
            TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
         End If
         Set S = Nothing
      End If

      TempRs.MoveNext
   Wend

   Call AddImportItem(II)
   
   Set TempSearch = Nothing
   Set S = Nothing
   Set Cm = Nothing
   Set Mi = Nothing
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GetRelateItem1(II As CImportItem, EI As CExportItem)
Dim iCount As Long
   Set EI = Nothing
   Set EI = GetExportItem(m_ExportItems, Trim(Str(II.GUI_ID)))
End Sub

Private Sub GetRelateItem2(EI As CExportItem, II As CImportItem)
Dim iCount As Long
   Set II = Nothing
   Set II = GetImportItem(m_ImportItems, Trim(Str(EI.GUI_ID)))
End Sub
Private Sub UpdateCurrentAmount(O As Object)
On Error Resume Next
Dim EI As CExportItem
Dim II As CImportItem
Dim TempII As CImportItem
Dim Key As String
Dim Amt As Double
Dim TempO As Object
Dim Ba As CBalanceAccum
Dim iCount As Long

   Set Ba = New CBalanceAccum
   
   If O.TX_TYPE = "I" Then
      Set II = O
      Key = II.PART_ITEM_ID & "-" & II.LOCATION_ID
      Amt = II.IMPORT_AMOUNT
   ElseIf O.TX_TYPE = "E" Then
      Set EI = O
      Key = EI.PART_ITEM_ID & "-" & EI.LOCATION_ID
      Amt = -1 * EI.EXPORT_AMOUNT
   End If

   Set TempII = m_MovementItemSearchs3(Key)
   TempII.LOCATION_NAME = O.LOCATION_NAME
   TempII.PART_NO = O.PART_NO
   TempII.PIG_TYPE = O.PIG_TYPE
   
   If TempII Is Nothing Then
      Set TempII = New CImportItem
      TempII.CURRENT_AMOUNT = Amt
      Call m_MovementItemSearchs3.Add(TempII, Key)
      Set TempII = Nothing
   Else
      If Amt < 0 Then
         TempII.CURRENT_AMOUNT = TempII.CURRENT_AMOUNT + Amt
      Else
         TempII.CURRENT_AMOUNT = TempII.CURRENT_AMOUNT + Amt
      End If
   End If
   
   Set Ba = Nothing
End Sub

Private Function VerifyCapitalBalanceDate(D As Date, BalDate As Date) As Boolean
Dim Cm As CCapitalMovement
Dim iCount As Long

   Set Cm = New CCapitalMovement
   Cm.CAPITAL_MOVEMENT_ID = -1
   Cm.DOCUMENT_CATEGORY = 3
   Call Cm.QueryData(8, m_Rs, iCount)
   If Not m_Rs.EOF Then
      Call Cm.PopulateFromRS(8, m_Rs)
      BalDate = Cm.DOCUMENT_DATE
   End If
   
   If D > Cm.DOCUMENT_DATE Then
      VerifyCapitalBalanceDate = True
   Else
      VerifyCapitalBalanceDate = False
   End If
   
   Set Cm = Nothing
End Function

Private Sub AdjustPigBalance(Imports As Collection, Exports As Collection, Balances As Collection, Optional BatchID As Long = -1)
Dim Ba As CImportItem
Dim II As CImportItem
Dim EI As CExportItem

   For Each Ba In Balances
'If Ba.LOCATION_ID = 255 And Ba.PART_ITEM_ID = 8897 Then
'''debug.print
'End If
      
      Set II = GetImportItem(Imports, Ba.LOCATION_ID & "-" & Ba.PART_ITEM_ID)
      Set EI = GetExportItem(Exports, Ba.LOCATION_ID & "-" & Ba.PART_ITEM_ID)
'      If Ba.CURRENT_AMOUNT <> (Ii.IMPORT_AMOUNT - Ei.EXPORT_AMOUNT) Then
'         ''debug.print
'      Else
'         ''debug.print
'      End If
      Ba.CURRENT_AMOUNT = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
   Next Ba
End Sub
Private Sub cmdStart_Click()
Dim Ivd As CInventoryDoc
Dim IsOK As Boolean
Dim iCount As Long
Dim O As Object
Dim Percent As Double
Dim I As Long
Dim ItemCount As Long
Dim PrevDate As String
Dim NewDate As Date
Dim PrevDocNo As Long
Dim EI As CExportItem
Dim II As CImportItem
Dim D As Date
Dim BatchID As Long
Dim CA As CCost_Accum

   If Not VerifyDate(lblFileName, uctlFromDate) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblBatch, cboBatch, Not cboBatch.Enabled) Then
      Exit Sub
   End If
   BatchID = cboBatch.ItemData(Minus2Zero(cboBatch.ListIndex))
   
   If Not VerifyCapitalBalanceDate(uctlFromDate.ShowDate, D) Then
      glbErrorLog.LocalErrorMsg = "จากวันที่จะต้องมีค่ามากกว่าวันที่ของต้นทุนยกมา (" & DateToStringExtEx2(D) & ")"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set Ivd = New CInventoryDoc
   Ivd.INVENTORY_DOC_ID = -1
   Ivd.FROM_DATE = uctlFromDate.ShowDate
   Ivd.TO_DATE = uctlToDate.ShowDate
   Ivd.COMMIT_FLAG = CommitTypeToFlag(cboCommitType.ItemData(Minus2Zero(cboCommitType.ListIndex)))
   Ivd.DELETE_BALANCE_FLAG = Check2Flag(chkBalanceFlag.Value)
   
   Set CA = New CCost_Accum
   CA.COST_ACCUM_ID = -1
   CA.FROM_DATE = DateAdd("D", -1, uctlFromDate.ShowDate)
   CA.TO_DATE = uctlToDate.ShowDate
   CA.BATCH_ID = cboBatch.ItemData(Minus2Zero(cboBatch.ListIndex))
   
   I = 0

   prgProgress.MIN = 0
   prgProgress.MAX = 100
   prgProgress.Value = 0
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction
   Call glbDaily.ClearCapitalMovement(Ivd, IsOK, False, glbErrorLog)
   Call glbDaily.ClearCostAccum(CA, IsOK, False, glbErrorLog)
   
   NewDate = DateAdd("D", -1, uctlFromDate.ShowDate)
   If uctlFromDate.ShowDate > 0 Then
   
      Call LoadInitialCostAccum(Nothing, m_CostAccumSearchs, m_CostAccumSearchExs, -1, NewDate, , , BatchID) 'เดี่ยวจะต้องทำในส่วนของ load ยอดยกมาจาก COST_ACCUM
      
'      Call LoadInitialCapitalBalance1(Nothing, m_MovementItemSearchs1, -1, NewDate, , , BatchID) 'เดี่ยวจะต้องทำในส่วนของ load ยอดยกมาจาก COST_ACCUM
'      Call LoadInitialCapitalBalance2(Nothing, m_MovementItemSearchs2, -1, NewDate, , , , BatchID)
      
'      Call LoadHousePigImportAmount(Nothing, m_ImportPigs, -1, NewDate, , , , , BatchID)
'      Call LoadPigHouseExportAmountEx(Nothing, m_ExportPigs, -1, NewDate, , , , , , BatchID)
'
'      Call LoadInitialPigBalance(Nothing, m_MovementItemSearchs3, -1, NewDate, , , BatchID)
'      Call AdjustPigBalance(m_ImportPigs, m_ExportPigs, m_MovementItemSearchs3, BatchID)
   End If
   
   Call LoadRelatedImportItemEx(Nothing, m_ImportItems, uctlFromDate.ShowDate, uctlToDate.ShowDate, "", 1, BatchID)
   Call LoadRelatedExportItemEx(Nothing, m_ExportItems, uctlFromDate.ShowDate, uctlToDate.ShowDate, "", 1, BatchID)
   
   Set O = glbDaily.QueryAllTransaction(Ivd, IsOK, ItemCount, glbErrorLog, True, BatchID)
   
   Dim FileID As Long
   
   On Error GoTo XXX
      Call Kill("C:\Log.txt")
XXX:

   FileID = FreeFile
   Open "C:\Log.txt" For Append As #FileID
   
   While (I = 0) Or (Not (O Is Nothing))
      DoEvents
      Percent = MyDiffEx2(I, ItemCount) * 50
      prgProgress.Value = Percent
      txtPercent.Text = Percent
      Me.Refresh
      
      Set O = glbDaily.QueryAllTransaction(Ivd, IsOK, iCount, glbErrorLog, , BatchID)                     ' ทุกครั้งที่เข้ามานั้นจะมีการเรียงลำดับอยู่แล้ว
            
      If Not (O Is Nothing) Then
         If (O.TX_TYPE = "E") And (O.DOCUMENT_TYPE = 2) Then 'เบิกวัตถุดิบ
            Call GenerateExportMoveMent(O, BatchID, FileID)
         ElseIf (O.TX_TYPE = "I") And (O.DOCUMENT_TYPE = 11) Then 'ใบซื้อสุกร
            Call GeneratePigImportMoveMent(O, BatchID, FileID)
            Call UpdateCurrentAmount(O)
         ElseIf (O.TX_TYPE = "I") And ((O.DOCUMENT_TYPE = 6) Or (O.DOCUMENT_TYPE = 7) Or (O.DOCUMENT_TYPE = 8)) Then  'ใบโอนสุกรเข้าเรือนขาย
            Set EI = New CExportItem
            Call GetRelateItem1(O, EI)
            Call GeneratePigTransferMovementImp(O, EI, BatchID, FileID)
            Call UpdateCurrentAmount(O)
            Set EI = Nothing
         ElseIf (O.TX_TYPE = "E") And ((O.DOCUMENT_TYPE = 6) Or (O.DOCUMENT_TYPE = 7) Or (O.DOCUMENT_TYPE = 8) Or (O.DOCUMENT_TYPE = 9) Or (O.DOCUMENT_TYPE = 10) Or (O.DOCUMENT_TYPE = 13)) Then 'ใบโอนสุกรเข้าเรือนขาย
            Set II = New CImportItem
            Call GetRelateItem2(O, II)
            Call GeneratePigTransferMovementExp(O, II, BatchID, FileID)
            Call UpdateCurrentAmount(O)
            Set II = Nothing
         ElseIf (O.TX_TYPE = "I") And (O.DOCUMENT_TYPE = 5) Then 'ใบสุกรคลอด
            Call GeneratePigBirthMoveMent(O, BatchID, FileID)
            Call UpdateCurrentAmount(O)
         ElseIf (O.TX_TYPE = "X") And (O.DOCUMENT_TYPE = 5) Then 'ใบรับของ ระบบงานซื้อ
            Call GenerateExpenseMovement(O, BatchID, FileID)
         ElseIf (O.TX_TYPE = "I") And (O.DOCUMENT_TYPE = 12) Then 'ใบเปลี่ยนสถานะสุกรในเรือนขาย
            'Set EI = New CExportItem
            'Call GetRelateItem1(O, EI)
            'Call GeneratePigStatusChangeImp(O, EI, BatchID)
            'Call UpdateCurrentAmount(O)
            'Set EI = Nothing
         ElseIf (O.TX_TYPE = "E") And (O.DOCUMENT_TYPE = 12) Then 'ใบเปลี่ยนสถานะสุกรในเรือนขาย
            'Set II = New CImportItem
            'Call GetRelateItem2(O, II)
            'Call GeneratePigStatusChangeExp(O, II, BatchID)
            'Call UpdateCurrentAmount(O)
            'Set II = Nothing
         End If
      End If
      
      I = I + 1
   Wend
   
   Close #FileID
   
   I = 0
   For Each CA In m_CostAccumSearchs
      DoEvents
      I = I + 1
      Percent = MyDiffEx2(I, m_CostAccumSearchs.Count) * 50
      prgProgress.Value = Percent + 50
      txtPercent.Text = Percent + 50
      Me.Refresh
      CA.AddEditMode = SHOW_ADD
      Call ForAdd(CA)
   Next CA
   
'   If Check2Flag(chkLoss.Value) = "Y" Then
'      Call GenerateCapitalLoss
'   End If
   
   Call glbDaily.CommitTransaction
   Call EnableForm(Me, True)
   Set Ivd = Nothing
   
   pnlHeader.Caption = "สร้างข้อมูลการเคลื่อนไหวต้นทุน (กรุณารอซักครู่)"
   Me.Refresh
   
   Call cmdOK_Click
   
End Sub
Private Sub ForAdd(CA As CCost_Accum)
On Error GoTo ErrorHandler
   Call CA.AddEditData
   
   Exit Sub
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "ค่า COST_EXP  หรือ COST_RAW  มีปัญหา"
      
   CA.COST_EXP = 0
   CA.COST_RAW = 0
   CA.COST_PB = 0
   Call CA.AddEditData
End Sub
Private Sub GenerateCapitalLoss()
Dim II As CImportItem
Dim S As CMovementItemSearch1
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim Cl As CCapitalLoss
Dim Li As CLossItem

'glbErrorLog.LocalErrorMsg = "==== Generate capital loss = "
'glbErrorLog.ShowUserErrorEx
   For Each II In m_MovementItemSearchs3
'If (Ii.PART_ITEM_ID = 8897) And (Ii.LOCATION_ID = 255) Then
'glbErrorLog.LocalErrorMsg = "==== CURRENT AMOUNT = " & Ii.CURRENT_AMOUNT
'glbErrorLog.ShowUserErrorEx
'End If

      If (II.CURRENT_AMOUNT <= 0) And (II.LOCATION_ID > 0) And (II.PART_ITEM_ID > 0) Then
         Set Cl = New CCapitalLoss
         Set Cm = New CCapitalMovement
         Cm.AddEditMode = SHOW_ADD
         Cm.COMMIT_FLAG = "N"
         Cm.DOCUMENT_NO = "สูญเสีย"
         Cm.DOCUMENT_DATE = uctlToDate.ShowDate
         Cm.IVD_ID = -1
         Cm.DOCUMENT_CATEGORY = 1
         Cm.DOCUMENT_TYPE = 14 'โอนย้ายต้นทุน
         Cm.TX_AMOUNT = 0
         Cm.TX_TYPE = "E"
         Cm.FROM_HOUSE_ID = II.LOCATION_ID
         Cm.TO_HOUSE_ID = -1
         Cm.PIG_ID = II.PART_ITEM_ID
         Cm.TO_PIG_ID = -1
         Cm.PIG_STATUS = -1
         Cm.EXPORT_ITEM_ID = -1
         Cm.IMPORT_ITEM_ID = -1
         Cm.TX_SEQ = -1
         Cm.REPLACE_FLAG = "N"
         Call Cm.AddEditData
   
         Cl.AddEditMode = SHOW_ADD
         Call Cl.CopyFromCapitalMovenent(Cm)
         Call Cl.AddEditData
         
         For Each S In m_MovementItemSearchs1 'วัตถุดิบ
            If (S.PIG_ID = II.PART_ITEM_ID) And (S.HOUSE_ID = II.LOCATION_ID) Then 'II.PART_ITEM_ID is PIG_ID
               ''debug.print II.LOCATION_ID & "-" & II.PART_ITEM_ID & " ---> " & II.CURRENT_AMOUNT
               Set Mi = New CMovementItem
               Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
               Mi.AddEditMode = SHOW_ADD
               Mi.CAPITAL_AMOUNT = -1 * S.CAPITAL_AMOUNT
               Mi.PART_ITEM_ID = S.PART_ITEM_ID
               Mi.EXPENSE_TYPE = -1
               Call Mi.AddEditData
               
               Set Li = New CLossItem
               Li.AddEditMode = SHOW_ADD
               Call Li.CopyFromMovementItem(Mi)
               Li.CAPITAL_LOSS_ID = Cl.CAPITAL_LOSS_ID
               Call Li.AddEditData
               Set Li = Nothing
               
               Set Mi = Nothing
            End If
         Next S
         
         For Each S In m_MovementItemSearchs2 'ค่าใช้จ่าย
            If (S.PIG_ID = II.PART_ITEM_ID) And (S.HOUSE_ID = II.LOCATION_ID) Then 'II.PART_ITEM_ID is PIG_ID
               ''debug.print II.LOCATION_ID & "-" & II.PART_ITEM_ID & " ---> " & S.CAPITAL_AMOUNT
               Set Mi = New CMovementItem
               Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
               Mi.AddEditMode = SHOW_ADD
               Mi.CAPITAL_AMOUNT = -1 * S.CAPITAL_AMOUNT
               Mi.PART_ITEM_ID = -1
               Mi.EXPENSE_TYPE = S.EXPENSE_TYPE
               Call Mi.AddEditData
            
               Set Li = New CLossItem
               Li.AddEditMode = SHOW_ADD
               Call Li.CopyFromMovementItem(Mi)
               Li.CAPITAL_LOSS_ID = Cl.CAPITAL_LOSS_ID
               Call Li.AddEditData
               Set Li = Nothing
               
               Set Mi = Nothing
            End If
         Next S
         
         Set Cm = Nothing
         Set Cl = Nothing
      End If 'Current amount
   Next II
End Sub

Private Sub cmdStartNew_Click()
Dim Ivd As CInventoryDoc
Dim IsOK As Boolean
Dim iCount As Long
Dim O As Object
Dim Percent As Double
Dim I As Long
Dim ItemCount As Long
Dim PrevDate As String
Dim NewDate As Date
Dim PrevDocNo As Long
Dim EI As CExportItem
Dim II As CImportItem
Dim D As Date

   If Not VerifyDate(lblFileName, uctlFromDate) Then
      Exit Sub
   End If
   
   If Not VerifyCapitalBalanceDate(uctlFromDate.ShowDate, D) Then
      glbErrorLog.LocalErrorMsg = "จากวันที่จะต้องมีค่ามากกว่าวันที่ของต้นทุนยกมา (" & DateToStringExtEx2(D) & ")"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set Ivd = New CInventoryDoc
   Ivd.INVENTORY_DOC_ID = -1
   Ivd.FROM_DATE = uctlFromDate.ShowDate
   Ivd.TO_DATE = uctlToDate.ShowDate
   Ivd.COMMIT_FLAG = CommitTypeToFlag(cboCommitType.ItemData(Minus2Zero(cboCommitType.ListIndex)))
   Ivd.DELETE_BALANCE_FLAG = Check2Flag(chkBalanceFlag.Value)

   I = 0

   prgProgress.MIN = 0
   prgProgress.MAX = 100
   prgProgress.Value = 0
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction
   Call glbDaily.ClearCapitalMovement(Ivd, IsOK, False, glbErrorLog)

   NewDate = DateAdd("D", -1, uctlFromDate.ShowDate)
   If uctlFromDate.ShowDate > 0 Then
      Call LoadInitialCapitalBalance1(Nothing, m_MovementItemSearchs1, -1, NewDate)
      Call LoadInitialCapitalBalance2(Nothing, m_MovementItemSearchs2, -1, NewDate)
      Call LoadInitialPigBalance(Nothing, m_MovementItemSearchs3, -1, NewDate)
   End If
   
   Set O = glbDaily.QueryAllTransaction(Ivd, IsOK, ItemCount, glbErrorLog, True)
   While (I = 0) Or (Not (O Is Nothing))
      DoEvents
      Percent = MyDiff(I, ItemCount) * 100
      prgProgress.Value = Percent
      txtPercent.Text = Percent
      Me.Refresh
      
      Set O = glbDaily.QueryAllTransaction(Ivd, IsOK, iCount, glbErrorLog)
      If Not (O Is Nothing) Then
         If (O.TX_TYPE = "E") And (O.DOCUMENT_TYPE = 2) Then 'เบิกวัตถุดิบ
            Call GenerateExportMoveMentEx1(O)
         ElseIf (O.TX_TYPE = "I") And (O.DOCUMENT_TYPE = 11) Then 'ใบซื้อสุกร
            Call GeneratePigImportMoveMentEx1(O)
            Call UpdateCurrentAmount(O)
         ElseIf (O.TX_TYPE = "I") And ((O.DOCUMENT_TYPE = 6) Or (O.DOCUMENT_TYPE = 7) Or (O.DOCUMENT_TYPE = 8) Or (O.DOCUMENT_TYPE = 12)) Then 'ใบโอนสุกรเข้าเรือนขาย
            Set EI = New CExportItem
            Call GetRelateItem1(O, EI)
            Call GeneratePigTransferMovementImpEx1(O, EI)
            Call UpdateCurrentAmount(O)
            Set EI = Nothing
         ElseIf (O.TX_TYPE = "E") And ((O.DOCUMENT_TYPE = 6) Or (O.DOCUMENT_TYPE = 7) Or (O.DOCUMENT_TYPE = 8) Or (O.DOCUMENT_TYPE = 9) Or (O.DOCUMENT_TYPE = 10) Or (O.DOCUMENT_TYPE = 12)) Then 'ใบโอนสุกรเข้าเรือนขาย
            Set II = New CImportItem
            Call GetRelateItem2(O, II)
            Call GeneratePigTransferMovementExpEx1(O, II)
            Call UpdateCurrentAmount(O)
            Set II = Nothing
         ElseIf (O.TX_TYPE = "I") And (O.DOCUMENT_TYPE = 5) Then 'ใบสุกรคลอด
            Call GeneratePigBirthMoveMentEx(O, 0)
            Call UpdateCurrentAmount(O)
         ElseIf (O.TX_TYPE = "X") And (O.DOCUMENT_TYPE = 5) Then 'ใบรับของ ระบบงานซื้อ
            Call GenerateExpenseMovement(O, 0)
         End If
      End If
      
      I = I + 1
   Wend
   
   Call glbDaily.CommitTransaction
   Call EnableForm(Me, True)
   Set Ivd = Nothing
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
'      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlFromDate.ShowDate = DateSerial(Year(Now), 1, 1)
      uctlToDate.ShowDate = DateSerial(Year(Now), 12, 31)
      
      Call LoadBatch(cboBatch)
      
      Call LoadProductStatus(Nothing, m_ProductStatuss)
      Call InitCommitStatus(cboCommitType)
      
      Call LoadProductType(Nothing, m_PigTypes)
      
      Call LoadPartItem(Nothing, m_PartItems)
      
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

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "สร้างข้อมูลการเคลื่อนไหวต้นทุน"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "จากวันที่")
   Call InitNormalLabel(lblMasterName, "ถึงวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblCommitType, "สถานะรายการ")
   Call InitNormalLabel(lblBatch, "แบต")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   Call InitCheckBox(chkBalanceFlag, "ลบข้อมูลต้นทุนยกมา")
   Call InitCheckBox(chkLoss, "ย้ายต้นทุนคงเหลือ")
   chkLoss.Value = FlagToCheck("Y")
   
   chkBalanceFlag.Value = ssCBUnchecked
   Call InitCombo(cboCommitType)
   Call InitCombo(cboBatch)
   cboBatch.Enabled = (glbUser.SIMULATE_FLAG = "Y")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
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
   
   Set m_Employee = New CEmployee
   Set m_Rs = New ADODB.Recordset
   
   Set m_MovementItemSearchs1 = New Collection
   Set m_MovementItemSearchs2 = New Collection
   Set m_MovementItemSearchs3 = New Collection
   Set m_ProductStatuss = New Collection
   Set m_TempSearchs1 = New Collection
   Set m_PigBirthInMonthLocations = New Collection
   Set m_PigTypes = New Collection
   Set m_PartItems = New Collection
   Set m_ImportItems = New Collection
   Set m_ExportItems = New Collection
   Set m_ImportPigs = New Collection
   Set m_ExportPigs = New Collection
      
   Set m_CostAccumSearchs = New Collection
   Set m_CostAccumSearchExs = New Collection
   Set m_ExportIDs = New Collection
   
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

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_TempSearchs1 = Nothing
   Set m_MovementItemSearchs1 = Nothing
   Set m_MovementItemSearchs2 = Nothing
   Set m_MovementItemSearchs3 = Nothing
   
   Set m_CostAccumSearchs = Nothing
   Set m_CostAccumSearchExs = Nothing
   Set m_ExportIDs = Nothing
   
   Set m_ProductStatuss = Nothing
   Set m_PigBirthInMonthLocations = Nothing
   Set m_PigTypes = Nothing
   Set m_PartItems = Nothing
   Set m_ImportItems = Nothing
   Set m_ExportItems = Nothing
   Set m_ImportPigs = Nothing
   Set m_ExportPigs = Nothing
End Sub
