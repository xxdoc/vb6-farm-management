VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcssCommit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmProcessCommit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6376
      _Version        =   131073
      PictureBackgroundStyle=   2
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
         TabIndex        =   2
         Top             =   1920
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   8
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
         TabIndex        =   3
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   1530
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   4
         Top             =   2730
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessCommit.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   12
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   6
         Top             =   2730
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   5
         Top             =   2730
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessCommit.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmProcssCommit"
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

Private m_ProductStatuss As Collection

Public DocumentCategory As Long
Public DocumentType As Long

Private Sub cmdPasswd_Click()

End Sub


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
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = ID
      m_Employee.QueryFlag = 1
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
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

Private Sub GenerateExportMoveMent(EI As CExportItem)
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   
   Cm.AddEditMode = SHOW_ADD
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
   
   Set Mi = Nothing
   Set Cm = Nothing
End Sub

Private Sub GeneratePigImportMoveMent(II As CImportItem)
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   
   Cm.AddEditMode = SHOW_ADD
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
'If S.GetKey2 = "330-10295" Then
'   ''debug.print "Y: " & S.GetKey3
'End If
   Else
      TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
   End If
   Set S = Nothing
   
   Set Mi = Nothing
   Set Cm = Nothing
End Sub

Private Function GetPigBirthInMonth(FirstDate As Date, LastDate As Date) As Long
Static TempFirstDate As Date
Static TempLastDate As Date
Static TempCol As Collection

   If (FirstDate <> TempFirstDate) Or (LastDate <> TempLastDate) Then
      Set TempCol = New Collection
      
      TempFirstDate = FirstDate
      TempLastDate = LastDate
   End If
   
   If TempCol.Count > 0 Then
      GetPigBirthInMonth = TempCol(1).IMPORT_AMOUNT
   Else
      GetPigBirthInMonth = 0
   End If
End Function

Private Function GetPigBirthInMonthInHouse(FirstDate As Date, LastDate As Date, HouseId As Long, PigID As Long) As Long
Static TempFirstDate As Date
Static TempLastDate As Date
Static TempCol As Collection
Static TempHouseID As Long
Static TempPigID As Long

   If (FirstDate <> TempFirstDate) Or (LastDate <> TempLastDate) Or (HouseId <> TempHouseID) Or (PigID <> TempPigID) Then
      Set TempCol = New Collection

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

Private Sub GeneratePigBirthMoveMent(II As CImportItem)
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
Dim PigBirthInMonth As Long
Dim PigBirthInMonthInHouse As Long
Dim S As CMovementItemSearch1
Dim Ms As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   Set S = New CMovementItemSearch1
   Set Ms = New CMovementItemSearch1

   Call GetFirstLastDate(II.DOCUMENT_DATE, FirstDate, LastDate)
   FirstDateStr = DateToStringInt(FirstDate)
   LastDateStr = DateToStringInt(LastDate)

   If (TempFirstDateStr <> FirstDateStr) Or (TempLastDateStr <> LastDateStr) Then
      TempFirstDateStr = FirstDateStr
      TempLastDateStr = LastDateStr
      TempHouseID = 0

      Set TempCol = Nothing
      Set TempCol = New Collection
   End If

   Ms.HOUSE_ID = II.LOCATION_ID
   Ms.PIG_ID = II.PART_ITEM_ID
   Set TempSearch = GetMovementSearch1(m_TempSearchs1, Ms.GetKey2)
   If TempSearch Is Nothing Then
      Call m_TempSearchs1.Add(Ms, Ms.GetKey2)

      PigBirthInMonth = GetPigBirthInMonth(FirstDate, LastDate)
      PigBirthInMonthInHouse = GetPigBirthInMonthInHouse(FirstDate, LastDate, II.LOCATION_ID, II.PART_ITEM_ID)

      'ให้ใบแรกเป็นตัวแทนของสุกรคลอดทั้งเดือน
      Cm.AddEditMode = SHOW_ADD
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
      Call Cm.AddEditData

      For Each EI In TempCol
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.PART_ITEM_ID = EI.PART_ITEM_ID
         Mi.EXPENSE_TYPE = 0
         Mi.CAPITAL_AMOUNT = MyDiff(MyDiff(EI.EXPORT_TOTAL_PRICE, PigBirthInMonth), PigBirthInMonthInHouse)
         Call Mi.AddEditData

         'วัตถุดิบ
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
      Next EI
   End If

   Set Ms = Nothing
   Set S = Nothing
   Set Mi = Nothing
   Set Cm = Nothing
End Sub

Private Sub GenerateExpenseMovement(Ri As CROItem)
Dim ExpenseRatios As Collection
Dim PigInHouses As Collection
Dim EI As CExportItem
Dim Er As CExpenseRatio
Dim II As CPartItem
Dim PigInAllHouse As Long
Dim PigInHouse As Long
Dim ExpenseAmt As Double
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1
Dim FirstDate As Date
Dim LastDate As Date
Dim PigImportAmounts As Collection
Dim ImportItem As CImportItem

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   Set ExpenseRatios = New Collection
   Set PigInHouses = New Collection
   Set EI = New CExportItem
   Set TempSearch = New CMovementItemSearch1
   Set PigImportAmounts = New Collection
   
   Call GetFirstLastDate(Ri.DOCUMENT_DATE, FirstDate, LastDate)
   
   Call LoadExpenseRatio(Nothing, ExpenseRatios, Ri.RO_ITEM_ID)
   For Each Er In ExpenseRatios
      Call LoadImportPig(Nothing, PigInHouses, Er.LOCATION_ID)
      ExpenseAmt = (Er.RATIO / 100) * Ri.TOTAL_PRICE
      
      PigInAllHouse = 0
      
      For Each II In PigInHouses
         Set ImportItem = GetImportItem(PigImportAmounts, Trim(Str(II.PART_ITEM_ID)))

         PigInHouse = ImportItem.IMPORT_AMOUNT
         PigInAllHouse = PigInAllHouse + PigInHouse
      Next II

      For Each II In PigInHouses
         Cm.AddEditMode = SHOW_ADD
         Cm.DOCUMENT_NO = Ri.DOCUMENT_NO
         Cm.DOCUMENT_DATE = Ri.DOCUMENT_DATE
         Cm.BL_ID = Ri.BILLING_DOC_ID
         Cm.DOCUMENT_CATEGORY = 2
         Cm.DOCUMENT_TYPE = Ri.DOCUMENT_TYPE
         Cm.TX_TYPE = Ri.TX_TYPE
         Cm.TX_AMOUNT = 0
         Cm.FROM_HOUSE_ID = Er.LOCATION_ID
         Cm.TO_HOUSE_ID = 0
         Cm.PIG_ID = II.PART_ITEM_ID
         Cm.PIG_STATUS = 0
         Call Cm.AddEditData
   
         Set ImportItem = GetImportItem(PigImportAmounts, Trim(Str(II.PART_ITEM_ID)))
         PigInHouse = ImportItem.IMPORT_AMOUNT
         If PigInHouse > 0 Then
            Mi.AddEditMode = SHOW_ADD
            Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
            Mi.PART_ITEM_ID = 0
            Mi.EXPENSE_TYPE = Ri.EXPENSE_TYPE
            Mi.CAPITAL_AMOUNT = MyDiff(ExpenseAmt, PigInAllHouse) * PigInHouse
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
'If S.GetKey2 = "330-10295" Then
'   ''debug.print "Z: " & S.GetKey3
'End If
            Else
               TempSearch.CAPITAL_AMOUNT = TempSearch.CAPITAL_AMOUNT + S.CAPITAL_AMOUNT
            End If
            Set S = Nothing
         End If
      Next II
   Next Er

   Set PigImportAmounts = Nothing
   Set TempSearch = Nothing
   Set EI = Nothing
   Set PigInHouses = Nothing
   Set ExpenseRatios = Nothing
   Set Mi = Nothing
   Set Cm = Nothing
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
Dim II As CImportItem
Dim Key As String

   If O.TX_TYPE = "I" Then
      Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID
   ElseIf O.TX_TYPE = "E" Then
      Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID
   Else
      Key = ""
   End If
   Set II = GetImportItem(m_MovementItemSearchs3, Key)
   
   GetPreviousAmount = Minus2Zero(II.CURRENT_AMOUNT)
End Function

Private Sub GeneratePigTransferMovementExp(EI As CExportItem, II As CImportItem)
Dim PigCount As Long
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim S As CMovementItemSearch1
Dim TempSearch As CMovementItemSearch1
Dim Ps As CProductStatus
Dim Amt As Double

   PigCount = GetPreviousAmount(EI)

   Set Cm = New CCapitalMovement
   Set Mi = New CMovementItem
   Set S = New CMovementItemSearch1

   Cm.AddEditMode = SHOW_ADD
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
   Call Cm.AddEditData

   If EI.PIG_STATUS > 0 Then
      Set Ps = m_ProductStatuss(Trim(Str(EI.PIG_STATUS)))
   Else
      Set Ps = New CProductStatus
   End If

   For Each S In m_MovementItemSearchs1
      If (S.PIG_ID = EI.PART_ITEM_ID) And (S.HOUSE_ID = EI.LOCATION_ID) Then
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.PART_ITEM_ID = S.PART_ITEM_ID

         If EI.PIG_STATUS > 0 Then
            'โอนเข้าเรือนขาย
            If Ps.CAPITAL_MOVE_FLAG = "Y" Then
               Amt = (MyDiff(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
            Else
               Amt = 0
            End If
         Else
            Amt = (MyDiff(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
         End If

         Mi.CAPITAL_AMOUNT = -1 * Amt
         Call Mi.AddEditData

         S.CAPITAL_AMOUNT = S.CAPITAL_AMOUNT + (-1 * Amt)
      End If
   Next S

   'ค่าใช้จ่าย
   For Each S In m_MovementItemSearchs2
'''debug.print S.GetKey3
'If S.GetKey2 = "330-10295" Then
'   ''debug.print "1: " & S.GetKey3
'End If
      If (S.PIG_ID = EI.PART_ITEM_ID) And (S.HOUSE_ID = EI.LOCATION_ID) Then
         Mi.AddEditMode = SHOW_ADD
         Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
         Mi.EXPENSE_TYPE = S.EXPENSE_TYPE
         Mi.PART_ITEM_ID = 0
         Mi.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID
         
         If EI.PIG_STATUS > 0 Then
            'โอนเข้าเรือนขาย
            If Ps.CAPITAL_MOVE_FLAG = "Y" Then
               Amt = (MyDiff(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
            Else
               Amt = 0
            End If
         Else
            Amt = (MyDiff(S.CAPITAL_AMOUNT, PigCount)) * EI.EXPORT_AMOUNT
         End If
         
         Mi.CAPITAL_AMOUNT = -1 * Amt
         Call Mi.AddEditData

         S.CAPITAL_AMOUNT = S.CAPITAL_AMOUNT + (-1 * Amt)
      End If
   Next S

   Set Ps = Nothing
   Set S = Nothing
   Set Mi = Nothing
   Set Cm = Nothing
End Sub

Private Sub GeneratePigTransferMovementImp(II As CImportItem, FromEi As CExportItem)
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
'If S.GetKey2 = "330-10295" Then
'   ''debug.print "X: " & S.GetKey3
'End If
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
Dim TempRs As ADODB.Recordset

   Set TempRs = New ADODB.Recordset
   
   EI.EXPORT_ITEM_ID = -1
   EI.GUI_ID = II.GUI_ID
   Call EI.QueryData(1, TempRs, iCount)
   If Not TempRs.EOF Then
      Call EI.PopulateFromRS(1, TempRs)
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GetRelateItem2(EI As CExportItem, II As CImportItem)
Dim iCount As Long
Dim TempRs As ADODB.Recordset

   Set TempRs = New ADODB.Recordset
   
   II.IMPORT_ITEM_ID = -1
   II.GUI_ID = EI.GUI_ID
   Call II.QueryData(1, TempRs, iCount)
   If Not TempRs.EOF Then
      Call II.PopulateFromRS(1, TempRs)
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub UpdateCurrentAmount(O As Object)
On Error Resume Next
Dim EI As CExportItem
Dim II As CImportItem
Dim TempII As CImportItem
Dim Key As String
Dim Amt As Double

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
   
   If TempII Is Nothing Then
      Set TempII = New CImportItem
      TempII.CURRENT_AMOUNT = Amt
      Call m_MovementItemSearchs3.Add(TempII, Key)
      Set TempII = Nothing
   Else
      If Amt < 0 Then
'         Amt = 0
         TempII = TempII.CURRENT_AMOUNT + Amt
      Else
         TempII = TempII.CURRENT_AMOUNT + Amt
      End If
   End If
End Sub

Private Function DoCommitDoc(D As CInventoryDoc) As Boolean
On Error GoTo ErrorHandler
Dim Ivd As CInventoryDoc
Dim IsOK As Boolean
Dim iCount As Long
Dim Result As Boolean

   Result = True
   DoCommitDoc = False
   Set Ivd = New CInventoryDoc
   
   Ivd.INVENTORY_DOC_ID = D.INVENTORY_DOC_ID
   Ivd.COMMIT_FLAG = "Y"
   Call Ivd.UpdateCommitFlag
   
   Set Ivd = Nothing
   DoCommitDoc = True
   Exit Function
   
ErrorHandler:
   DoCommitDoc = False
End Function

Private Function DO2InventoryDoc(Bd As CBillingDoc, Ivd As CInventoryDoc) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDoItem
Dim EI As CExportItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If Bd.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = Bd.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = Bd.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = Bd.DOCUMENT_NO
   Ivd.COMMIT_FLAG = Bd.COMMIT_FLAG
   Ivd.DOCUMENT_TYPE = 10
   
   For Each Di In Bd.DoItems
      If Di.Flag = "A" Then
         Set EI = New CExportItem
         
         EI.TX_TYPE = "E"
         EI.Flag = "A"
         EI.PART_ITEM_ID = Di.PART_ITEM_ID
         EI.PIG_STATUS = Di.PIG_STATUS
         EI.LOCATION_ID = Di.LOCATION_ID
         EI.EXPORT_AMOUNT = Di.ITEM_AMOUNT
         EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         EI.TOTAL_PRICE = Di.TOTAL_PRICE
         EI.LINK_ID = Di.LINK_ID
         EI.CALCULATE_FLAG = "N"
         
         Call Ivd.ImportExports.Add(EI)
         Set EI = Nothing
      ElseIf Di.Flag = "E" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         
         EI.Flag = "E"
         EI.PART_ITEM_ID = Di.PART_ITEM_ID
         EI.PIG_STATUS = Di.PIG_STATUS
         EI.LOCATION_ID = Di.LOCATION_ID
         EI.EXPORT_AMOUNT = Di.ITEM_AMOUNT
         EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         EI.CALCULATE_FLAG = "N"
      ElseIf Di.Flag = "D" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         EI.Flag = "D"
      End If
   Next Di
End Function

Private Function DoCommitBillingDoc(D As CBillingDoc) As Boolean
On Error GoTo ErrorHandler
Dim Ivd As CInventoryDoc
Dim Bd As CBillingDoc
Dim IsOK As Boolean
Dim iCount As Long
Dim Result As Boolean

   Result = True
   DoCommitBillingDoc = False
   Set Ivd = New CInventoryDoc
   Set Bd = New CBillingDoc

   Bd.BILLING_DOC_ID = D.BILLING_DOC_ID
   Bd.COMMIT_FLAG = "Y"
   Call Bd.UpdateCommitFlag
   
   Ivd.INVENTORY_DOC_ID = D.INVENTORY_DOC_ID
   Ivd.COMMIT_FLAG = "Y"
   Call Ivd.UpdateCommitFlag

   Set Ivd = Nothing
   Set Bd = Nothing
   DoCommitBillingDoc = True
   Exit Function
   
ErrorHandler:
   DoCommitBillingDoc = False
End Function

Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim Bd As CBillingDoc
Dim Ivd As CInventoryDoc
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim Percent As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean

   If Not VerifyDate(lblFileName, uctlFromDate, False) Then
      Exit Sub
   End If

   HasBegin = False
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Set Ivd = New CInventoryDoc
   Set Bd = New CBillingDoc
   If DocumentCategory = 1 Then 'Inventory doc
      Call EnableForm(Me, False)
      Ivd.INVENTORY_DOC_ID = -1
      Ivd.FROM_DATE = uctlFromDate.ShowDate
      Ivd.TO_DATE = uctlToDate.ShowDate
      Ivd.DOCUMENT_TYPE = DocumentType
      Ivd.COMMIT_FLAG = "N"
      Call glbDaily.QueryInventoryDoc(Ivd, m_Rs, iCount, IsOK, glbErrorLog)
      RecordCount = iCount
      I = 0
      
      Call glbDaily.StartTransaction
      HasBegin = True
      While Not m_Rs.EOF
         I = I + 1
         Percent = MyDiff(I, RecordCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
         
         Call Ivd.PopulateFromRS(1, m_Rs)
         Result = DoCommitDoc(Ivd)
         
         Me.Refresh
         m_Rs.MoveNext
      Wend
      prgProgress.Value = 100
      Call glbDaily.CommitTransaction
      HasBegin = False
      Call EnableForm(Me, True)
   ElseIf DocumentCategory = 2 Then 'Billing doc
      Call EnableForm(Me, False)
      Bd.BILLING_DOC_ID = -1
      Bd.FROM_DATE = uctlFromDate.ShowDate
      Bd.TO_DATE = uctlToDate.ShowDate
      Bd.DOCUMENT_TYPE = DocumentType
      Bd.COMMIT_FLAG = "N"
      Call glbDaily.QueryBillingDoc(Bd, m_Rs, iCount, IsOK, glbErrorLog)
      RecordCount = iCount
      I = 0
      
      Call glbDaily.StartTransaction
      HasBegin = True
      While Not m_Rs.EOF
         I = I + 1
         Percent = MyDiff(I, RecordCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
         
         Call Bd.PopulateFromRS(1, m_Rs)
         Result = DoCommitBillingDoc(Bd)
         
         Me.Refresh
         m_Rs.MoveNext
      Wend
      prgProgress.Value = 100
      Call glbDaily.CommitTransaction
      HasBegin = False
      Call EnableForm(Me, True)
   End If
   Set Bd = Nothing
   Set Ivd = Nothing
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      Call glbDaily.RollbackTransaction
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadProductStatus(Nothing, m_ProductStatuss)
      
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
   pnlHeader.Caption = "ประมวลผลข้อมูล"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "จากวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
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
   Set m_ProductStatuss = Nothing
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
