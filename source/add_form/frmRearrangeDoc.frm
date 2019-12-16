VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReArrangeDoc 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmRearrangeDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   7646
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBatch 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3210
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
         TabIndex        =   3
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
         TabIndex        =   10
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
         TabIndex        =   4
         Top             =   2730
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
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   17
         Top             =   2280
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblPartItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblBatch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   3270
         Width           =   1575
      End
      Begin Threed.SSCheck chkBalanceFlag 
         Height          =   375
         Left            =   6450
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   6
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRearrangeDoc.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   15
         Top             =   2850
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   8
         Top             =   3600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   7
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRearrangeDoc.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmReArrangeDoc"
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

Private m_Balances As Collection
Private m_PartItemsDateLocations As Collection
Private m_PartItemsLocationMonthlies As Collection
Private m_Pigs As Collection
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
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = ID
      m_Employee.QueryFlag = 1
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
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

Private Function GetNextTransaction(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, II As CImportItem, EI As CExportItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
   
   'Export
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call EI.PopulateFromRS(13, Rs1)
   End If
   
   'Import
   EofFlag2 = Rs2.EOF
   If Not Rs2.EOF Then
      Call II.PopulateFromRS(7, Rs2)
   End If
   
   If (EofFlag1 And EofFlag2) Then
      GetNextTransaction = ""
   ElseIf (EofFlag1 And (Not EofFlag2)) Then
      GetNextTransaction = "I"
      Rs2.MoveNext
   ElseIf ((Not EofFlag1) And EofFlag2) Then
      GetNextTransaction = "E"
      Rs1.MoveNext
   Else
      '===
      'การเรียงลำดับมีผลอย่างมาก
      If DateToStringInt(EI.DOCUMENT_DATE) = DateToStringInt(II.DOCUMENT_DATE) Then
         If EI.PRIORITY1 = II.PRIORITY1 Then
            If EI.DOCUMENT_NO = II.DOCUMENT_NO Then
               If EI.TRANSACTION_SEQ < II.TRANSACTION_SEQ Then
                  GetNextTransaction = "E"
               Else
                  GetNextTransaction = "I"
               End If
            ElseIf EI.DOCUMENT_NO < II.DOCUMENT_NO Then
               GetNextTransaction = "E"
            Else
               GetNextTransaction = "I"
            End If
         ElseIf EI.PRIORITY1 < II.PRIORITY1 Then
            GetNextTransaction = "E"
         Else
            GetNextTransaction = "I"
         End If
      ElseIf DateToStringInt(EI.DOCUMENT_DATE) < DateToStringInt(II.DOCUMENT_DATE) Then
         GetNextTransaction = "E"
      Else
         GetNextTransaction = "I"
      End If 'Document date
      '===
      If GetNextTransaction = "I" Then
         Rs2.MoveNext
      ElseIf GetNextTransaction = "E" Then
         Rs1.MoveNext
      End If
   End If 'Eof flag
End Function

'Public Function GetBalanceAmount(PartItemID As Long, LocationID As Long, TxSeq As Long, DocDate As Date) As Object
'Dim EI As CExportItem
'Dim II As CImportItem
'Dim TempRs As ADODB.Recordset
'Dim iCount As Long
'
'   Set TempRs = New ADODB.Recordset
'
'   Set EI = New CExportItem
'   Set II = New CImportItem
'
'   EI.EXPORT_ITEM_ID = -1
'   EI.PIG_FLAG = "N"
'   EI.PART_ITEM_ID = PartItemID
'   EI.LOCATION_ID = LocationID
'   EI.FROM_TX_SEQ = -1
'   EI.TO_TX_SEQ = TxSeq
'   EI.FROM_DATE = -1
'   EI.TO_DATE = DocDate
'   EI.OrderBy = 11
'   EI.OrderType = 2
'   Call EI.QueryData(1, TempRs, iCount)
'   If Not TempRs.EOF Then
'      Call EI.PopulateFromRS(1, TempRs)
'   End If
'
'   II.IMPORT_ITEM_ID = -1
'   II.PIG_FLAG = "N"
'   II.PART_ITEM_ID = PartItemID
'   II.LOCATION_ID = LocationID
'   II.FROM_TX_SEQ = -1
'   II.TO_TX_SEQ = TxSeq
'   II.FROM_DATE = -1
'   II.TO_DATE = DocDate
'   II.OrderBy = 12
'   II.OrderType = 2
'   Call II.QueryData(1, TempRs, iCount)
'   If Not TempRs.EOF Then
'      Call II.PopulateFromRS(1, TempRs)
'   End If
'
'   If EI.TRANSACTION_SEQ > II.TRANSACTION_SEQ Then
'      Set GetBalanceAmount = EI
'   Else
'      Set GetBalanceAmount = II
'   End If
'
'   If TempRs.State = adStateOpen Then
'      Call TempRs.Close
'   End If
'   Set TempRs = Nothing
'   Set EI = Nothing
'   Set II = Nothing
'End Function

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

Private Sub GeneratePartItemLocationDate(O As Object, ImpI As CImportItem)
Dim Key As String
Dim Ba As CBalanceAccum
Dim II As CImportItem
Dim TempII As CImportItem
Dim AvgPrice As Double

'   If O.PIG_FLAG = "N" Then
'      Exit Sub
'   End If
   
   Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID & "-" & DateToStringInt(O.DOCUMENT_DATE)
   Set II = GetImportItem(m_PartItemsDateLocations, Key)
   If II.PART_ITEM_ID <= 0 Then
      Set TempII = New CImportItem
      TempII.PART_NO = O.PART_NO
      TempII.PIG_FLAG = O.PIG_FLAG
      TempII.LOCATION_ID = O.LOCATION_ID
      TempII.PART_ITEM_ID = O.PART_ITEM_ID
      TempII.DOCUMENT_DATE = O.DOCUMENT_DATE
      TempII.BALANCE_AMOUNT = ImpI.CURRENT_AMOUNT
      TempII.TOTAL_INCLUDE_PRICE = ImpI.TOTAL_INCLUDE_PRICE
      TempII.INCLUDE_UNIT_PRICE = MyDiffEx(ImpI.TOTAL_INCLUDE_PRICE, ImpI.CURRENT_AMOUNT)
      If O.TX_TYPE = "I" Then
         TempII.ALL_IMPORT_AMT = O.IMPORT_AMOUNT
      ElseIf O.TX_TYPE = "E" Then
         TempII.ALL_EXPORT_AMT = O.EXPORT_AMOUNT
      End If
      
      Call m_PartItemsDateLocations.Add(TempII, Key)
      Set TempII = Nothing
   Else
      If O.TX_TYPE = "I" Then
         II.INCLUDE_UNIT_PRICE = MyDiffEx(ImpI.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, ImpI.CURRENT_AMOUNT + O.IMPORT_AMOUNT)
         II.ALL_IMPORT_AMT = II.ALL_IMPORT_AMT + O.IMPORT_AMOUNT
         II.BALANCE_AMOUNT = II.BALANCE_AMOUNT + O.IMPORT_AMOUNT
         II.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
      ElseIf O.TX_TYPE = "E" Then
         II.ALL_EXPORT_AMT = II.ALL_EXPORT_AMT + O.EXPORT_AMOUNT
         II.BALANCE_AMOUNT = II.BALANCE_AMOUNT - O.EXPORT_AMOUNT
         II.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE - (O.EXPORT_AMOUNT * ImpI.INCLUDE_UNIT_PRICE)
      End If
   End If
End Sub
Private Sub GeneratePartItemLocationMonthly(O As Object, ImpI As CImportItem)
Dim Key As String
Dim Ba As CBalanceAccum
Dim II As CMonthlyAccum
Dim TempII As CMonthlyAccum
Dim AvgPrice As Double
   
   Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID & "-" & Mid(DateToStringInt(O.DOCUMENT_DATE), 1, 7)
   Set II = GetMonthlyAccum(m_PartItemsLocationMonthlies, Key)
   If II.PART_ITEM_ID <= 0 Then
      Set TempII = New CMonthlyAccum
      
      If O.TX_TYPE = "I" Then
         TempII.BALANCE_AMOUNT1 = ImpI.CURRENT_AMOUNT - O.IMPORT_AMOUNT
         If O.DOCUMENT_TYPE = 5 Then 'ใบเกิด
            TempII.BIRTH_AMOUNT = O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 11 Then 'ใบนำเข้าสุกร
            TempII.BUY_AMOUNT = O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 8 Then 'ใบขึ้นทดแทน
            TempII.STATUS_IN_AMOUNT = O.IMPORT_AMOUNT
         Else
            TempII.IMPORT_AMOUNT = O.IMPORT_AMOUNT
         End If
      ElseIf O.TX_TYPE = "E" Then
         TempII.BALANCE_AMOUNT1 = ImpI.CURRENT_AMOUNT + O.EXPORT_AMOUNT
         TempII.EXPORT_AMOUNT = O.EXPORT_AMOUNT
         If O.DOCUMENT_TYPE = 8 Then 'ใบขึ้นทดแทน
            TempII.STATUS_OUT_AMOUNT = O.EXPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 7 Then  'โอนไปเรือนขาย
            TempII.SELL_AMOUNT = O.EXPORT_AMOUNT
'''debug.print O.EXPORT_AMOUNT
         Else
            TempII.EXPORT_AMOUNT = O.EXPORT_AMOUNT
         End If
      End If
      
'      If O.TX_TYPE = "I" Then
'         'ต้องหักออกเพราะ CURRENT_AMOUNT ถูกบวกมาก่อนแล้ว
'         TempII.BALANCE_AMOUNT1 = ImpI.CURRENT_AMOUNT - O.IMPORT_AMOUNT
'         If O.DOCUMENT_TYPE = 5 Then
'            TempII.IMPORT_AMOUNT = 0
'            TempII.BIRTH_AMOUNT = O.IMPORT_AMOUNT
'         Else
'            TempII.IMPORT_AMOUNT = O.IMPORT_AMOUNT
'            TempII.BIRTH_AMOUNT = 0
'         End If
'      Else
'         ''ต้องบวกเพิ่มเพราะ CURRENT_AMOUNT ถูกหักออกมาก่อนแล้ว
'         TempII.BALANCE_AMOUNT1 = ImpI.CURRENT_AMOUNT + O.EXPORT_AMOUNT
'         TempII.EXPORT_AMOUNT = O.EXPORT_AMOUNT
'      End If
            
      TempII.LOCATION_ID = O.LOCATION_ID
      TempII.PART_ITEM_ID = O.PART_ITEM_ID
      TempII.DOCUMENT_DATE = O.DOCUMENT_DATE
      TempII.YYYYMM = Mid(DateToStringInt(O.DOCUMENT_DATE), 1, 7)
      TempII.BALANCE_AMOUNT2 = ImpI.CURRENT_AMOUNT
      Call m_PartItemsLocationMonthlies.Add(TempII, Key)
      Set TempII = Nothing
   Else
      If O.TX_TYPE = "I" Then
         If O.DOCUMENT_TYPE = 5 Then 'ใบเกิด
            II.BIRTH_AMOUNT = II.BIRTH_AMOUNT + O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 11 Then 'ใบนำเข้าสุกร
            II.BUY_AMOUNT = II.BUY_AMOUNT + O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 8 Then 'ใบขึ้นทดแทน
            II.STATUS_IN_AMOUNT = II.STATUS_IN_AMOUNT + O.IMPORT_AMOUNT
         Else
            II.IMPORT_AMOUNT = II.IMPORT_AMOUNT + O.IMPORT_AMOUNT
         End If
      ElseIf O.TX_TYPE = "E" Then
         If O.DOCUMENT_TYPE = 8 Then 'ใบขึ้นทดแทน
            II.STATUS_OUT_AMOUNT = II.STATUS_OUT_AMOUNT + O.EXPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 7 Then  'โอนไปเรือนขาย
            II.SELL_AMOUNT = II.SELL_AMOUNT + O.EXPORT_AMOUNT
'''debug.print O.EXPORT_AMOUNT
         Else
            II.EXPORT_AMOUNT = II.EXPORT_AMOUNT + O.EXPORT_AMOUNT
         End If
      End If
      II.BALANCE_AMOUNT2 = ImpI.CURRENT_AMOUNT
   End If
End Sub

Private Sub CopyBalanceAccum(Src As Collection, Dst As Collection)
Dim II As CImportItem
Dim Ba As CBalanceAccum
   
   For Each Ba In Src
      Set II = New CImportItem
      II.LOCATION_ID = Ba.LOCATION_ID
      II.PART_ITEM_ID = Ba.PART_ITEM_ID
      
'      If Ba.LOCATION_ID = 207 And Ba.PART_ITEM_ID = 12873 Then
'         ''debug.print
'      End If
      
      II.TOTAL_INCLUDE_PRICE = Ba.TOTAL_INCLUDE_PRICE
      II.CURRENT_AMOUNT = Ba.BALANCE_AMOUNT
      II.NEW_PRICE = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = Ba.AVG_PRICE
      II.TX_TYPE = "I"
      II.DOCUMENT_DATE = Ba.DOCUMENT_DATE
      Call Dst.Add(II, Ba.LOCATION_ID & "-" & Ba.PART_ITEM_ID)
      Set II = Nothing
   Next Ba
End Sub

Private Sub CalculateAdjustValue(II As CImportItem, Bals As Collection)
Dim TempII As CImportItem
Dim TempKey As String
Dim DifAmount As Double
Dim DifPrice As Double

   TempKey = II.LOCATION_ID & "-" & II.PART_ITEM_ID
   Set TempII = GetImportItem(Bals, TempKey)
   
   DifAmount = II.ACTUAL_AMOUNT - TempII.CURRENT_AMOUNT
   DifPrice = II.ACTUAL_PRICE - TempII.TOTAL_INCLUDE_PRICE
   
   II.IMPORT_AMOUNT = DifAmount
   II.TOTAL_INCLUDE_PRICE = DifPrice
      
   II.TOTAL_ACTUAL_PRICE = II.TOTAL_INCLUDE_PRICE
   II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.IMPORT_AMOUNT)
   II.ACTUAL_UNIT_PRICE = II.INCLUDE_UNIT_PRICE
   Call II.PatchAvgPrice(II.INCLUDE_UNIT_PRICE, 0, 0, 0, II.IMPORT_AMOUNT, 4, II.TOTAL_INCLUDE_PRICE)
End Sub
Private Sub cmdStart_Click()
'On Error GoTo ErrHandler
Dim Percent As Double
Dim MIN As Double
Dim MAX As Double
Dim RecordCount As Double
Dim O As Object
Dim TempO As Object
Dim InventoryBals As Collection
Dim RName As String
Dim cData As CPartLocation
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim IsOK As Boolean
Dim Amt As Double
Dim EI As CExportItem
Dim II As CImportItem
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim TxCode As String
Dim iCount As Long
Dim AvgPrice As Double
Dim PrevAmount As Double
Dim CurrentAmount As Double
Dim HasBegin As Boolean
Dim TempII As CImportItem
Dim TempKey As String
Dim Count1 As Long
Dim Count2 As Long
Dim TempCol As Collection
Dim TempEi As CExportItem
Dim ExportTotalPrice As Double
Dim NewDate As Date
Dim Ba As CBalanceAccum
Dim BalanceAccums As Collection
Dim BatchID As Long
Dim Ma As CMonthlyAccum
Dim PartItemID As Long

   If Not VerifyCombo(lblBatch, cboBatch, Not cboBatch.Enabled) Then
      Exit Sub
   End If
   BatchID = cboBatch.ItemData(Minus2Zero(cboBatch.ListIndex))
   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   
 '  PartItemID = 13844
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   HasBegin = False
   
   Call EnableForm(Me, False)
   
   Set Ba = New CBalanceAccum
   Ba.FROM_DATE = uctlFromDate.ShowDate
   Ba.TO_DATE = uctlToDate.ShowDate
   Ba.BATCH_ID = BatchID
   Ba.PART_ITEM_ID = PartItemID
   Call Ba.ClearData
   Set Ba = Nothing
   
   Set Ma = New CMonthlyAccum
   Ma.FROM_DATE = uctlFromDate.ShowDate
   Ma.TO_DATE = uctlToDate.ShowDate
   Ma.BATCH_ID = BatchID
   Ma.PART_ITEM_ID = PartItemID
   Call Ma.ClearData
   Set Ma = Nothing
   
   Set BalanceAccums = New Collection
   
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   
   Set TempCol = New Collection
   
   Set InventoryBals = New Collection
   Call LoadInventoryBalanceEx(Nothing, BalanceAccums, InternalDateToDate(DateToStringIntLow(uctlFromDate.ShowDate)), uctlToDate.ShowDate, "", , PartItemID, BatchID)
   
   Call CopyBalanceAccum(BalanceAccums, InventoryBals)
   
   Set TempEi = New CExportItem
   
   Set m_PartItemsDateLocations = Nothing
   Set m_PartItemsDateLocations = New Collection
   
   Set m_PartItemsLocationMonthlies = Nothing
   Set m_PartItemsLocationMonthlies = New Collection
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   NewDate = DateAdd("D", -1, uctlFromDate.ShowDate)
   
   '=== Detail
   Set EI = New CExportItem
   EI.EXPORT_ITEM_ID = -1
   EI.FROM_DATE = uctlFromDate.ShowDate
   EI.TO_DATE = uctlToDate.ShowDate
   EI.COMMIT_FLAG = ""
   EI.PIG_FLAG = ""
   EI.PART_ITEM_ID = -1
   EI.LOCATION_ID = -1
   EI.OrderBy = 2
   EI.OrderType = 1
   EI.BATCH_ID = BatchID
   EI.PART_ITEM_ID = PartItemID
'EI.LOCATION_ID = 254
'EI.DOCUMENT_NO = "B821/41038"
'EI.PIG_ID = 14475
'                           EI.PART_ITEM_ID = 14779
   Call EI.QueryData(13, Rs1, Count1)
   
   Set II = New CImportItem
   II.IMPORT_ITEM_ID = -1
   II.FROM_DATE = uctlFromDate.ShowDate
   II.TO_DATE = uctlToDate.ShowDate
   II.COMMIT_FLAG = ""
   II.PIG_FLAG = ""
   II.PART_ITEM_ID = -1
   II.LOCATION_ID = -1
   II.OrderBy = 2
   II.OrderType = 1
   II.BATCH_ID = BatchID
   II.PART_ITEM_ID = PartItemID
'                              II.PART_ITEM_ID = 14779
'II.LOCATION_ID = 236
'II.DOCUMENT_NO = "B821/41038"

   Call II.QueryData(7, Rs2, Count2)
   '== Detail
   
   Call glbDaily.StartTransaction
   
   MIN = 0
   MAX = 100
   Percent = 0
   RecordCount = 0
   prgProgress.MIN = MIN
   prgProgress.MAX = MAX
   
   TxCode = "X"
   While TxCode <> ""
      Percent = MyDiff(RecordCount, Count1 + Count2) * 100
      prgProgress.Value = Percent
      txtPercent.Text = Format(Percent, "0.00")
      
      TxCode = GetNextTransaction(Rs1, Rs2, II, EI)
      
'If II.DOCUMENT_NO = "B0987/098693" Then
'   Debug.Print ("")
'End If
'If Ei.DOCUMENT_NO = "B821/41038" Then
'   ''debug.print ("")
'End If
      

      
      If TxCode <> "" Then
         RecordCount = RecordCount + 1
         I = I + 1
         If TxCode = "I" Then
            '====
            Set O = II

'            If II.PART_ITEM_ID = 16234 And II.LOCATION_ID = 501 Then
'               Debug.Print "I " & O.DOCUMENT_NO & " " & O.DOCUMENT_DATE & " " & O.IMPORT_AMOUNT
'               Debug.Print
'            End If
            
            If II.DOCUMENT_TYPE = 3 Then 'ใบโอนวัตถุดิบ
               
               Set TempEi = New CExportItem
               Call GetRelateItem1(O, TempEi)
               II.INCLUDE_UNIT_PRICE = TempEi.EXPORT_AVG_PRICE
               II.TOTAL_INCLUDE_PRICE = TempEi.EXPORT_TOTAL_PRICE
               Set TempEi = Nothing
            ElseIf II.DOCUMENT_TYPE = 4 Then  'ใบปรับยอดวัตถุดิบ
               Call CalculateAdjustValue(II, InventoryBals)
            Else
               II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.IMPORT_AMOUNT)
            End If
            '====
         ElseIf TxCode = "E" Then
            Set O = EI
'''debug.print "E " & O.DOCUMENT_NO & " " & O.DOCUMENT_DATE & " " & O.EXPORT_AMOUNT
'            If EI.EXPORT_ITEM_ID = 817579 Then
'               ''debug.print
'            End If
         End If

         TempKey = O.LOCATION_ID & "-" & O.PART_ITEM_ID
         
            
                  
         Set TempII = GetImportItem(InventoryBals, TempKey)
         If TempII.PART_ITEM_ID <= 0 Then
            'Get balance item here
            Set TempO = GetImportItem(InventoryBals, TempKey)

            Set TempII = New CImportItem
            TempII.LOCATION_ID = O.LOCATION_ID
            TempII.PART_ITEM_ID = O.PART_ITEM_ID
            If O.TX_TYPE = "I" Then
               TempII.INCLUDE_UNIT_PRICE = MyDiffEx(O.TOTAL_INCLUDE_PRICE, O.IMPORT_AMOUNT)   'TempO.NEW_PRICE
               TempII.CURRENT_AMOUNT = O.IMPORT_AMOUNT  'TempO.CURRENT_AMOUNT
               TempII.TOTAL_INCLUDE_PRICE = O.TOTAL_INCLUDE_PRICE
            ElseIf O.TX_TYPE = "E" Then
               TempII.INCLUDE_UNIT_PRICE = O.EXPORT_AVG_PRICE
               TempII.CURRENT_AMOUNT = -1 * O.EXPORT_AMOUNT
               TempII.TOTAL_INCLUDE_PRICE = O.EXPORT_TOTAL_PRICE
            End If
            
            Call InventoryBals.Add(TempII, TempKey)
            Set TempII = Nothing
            Set TempII = GetImportItem(InventoryBals, TempKey)
         Else
            If O.TX_TYPE = "I" Then
               TempII.INCLUDE_UNIT_PRICE = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT)   'TempO.NEW_PRICE
               TempII.CURRENT_AMOUNT = TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT
               TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
            ElseIf O.TX_TYPE = "E" Then
               'TempII.INCLUDE_UNIT_PRICE =  TempO.EXPORT_AVG_PRICE 'ไม่เปลี่ยนแปลง
               TempII.CURRENT_AMOUNT = TempII.CURRENT_AMOUNT - O.EXPORT_AMOUNT
               TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE - (TempII.INCLUDE_UNIT_PRICE * O.EXPORT_AMOUNT)
            End If
         End If
         
         Call GeneratePartItemLocationDate(O, TempII)
         Call GeneratePartItemLocationMonthly(O, TempII)
         
         If TxCode = "I" Then
            PrevAmount = TempII.CURRENT_AMOUNT - O.IMPORT_AMOUNT
            CurrentAmount = PrevAmount + II.IMPORT_AMOUNT
            'AvgPrice = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE + II.TOTAL_INCLUDE_PRICE, CurrentAmount)
            AvgPrice = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, CurrentAmount)
                        
'            If CurrentAmount < 0 Then
'               Debug.Print
'            End If

            Call II.PatchAvgPrice(II.INCLUDE_UNIT_PRICE, PrevAmount, CurrentAmount, AvgPrice, II.IMPORT_AMOUNT, II.DOCUMENT_TYPE, II.TOTAL_INCLUDE_PRICE)
         ElseIf TxCode = "E" Then
            PrevAmount = TempII.CURRENT_AMOUNT + EI.EXPORT_AMOUNT
            CurrentAmount = PrevAmount - EI.EXPORT_AMOUNT
            AvgPrice = TempII.INCLUDE_UNIT_PRICE
               
            ExportTotalPrice = AvgPrice * EI.EXPORT_AMOUNT 'MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, PrevAmount) * EI.EXPORT_AMOUNT
'            If CurrentAmount < 0 Then
'               Debug.Print
'            End If
            Call EI.PatchAvgPrice(AvgPrice, PrevAmount, CurrentAmount, ExportTotalPrice)
         End If
      End If 'Tx code
      DoEvents
   Wend
   
   Call InsertBalanceAccum(BatchID)
   Call InsertMonthlyAccum(BatchID)

   txtPercent.Text = Format(100, "0.00")
   prgProgress.Value = 100
   Call glbDaily.CommitTransaction
   HasBegin = False
   
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
      
   If Rs2.State = adStateOpen Then
      Rs2.Close
   End If
   Set Rs2 = Nothing
      
   Set EI = Nothing
   Set II = Nothing
   Set TempEi = Nothing
   Set InventoryBals = Nothing
   Set BalanceAccums = Nothing
   Set TempCol = Nothing
   Call EnableForm(Me, True)
   
   Exit Sub
   
'ErrHandler:
'   If HasBegin Then
'      glbDaily.RollbackTransaction
'   End If
'   glbErrorLog.LocalErrorMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub InsertBalanceAccum(BatchID As Long)
Dim Ba As CBalanceAccum
Dim II As CImportItem
Dim iCount As Long

   For Each II In m_PartItemsDateLocations
      Set Ba = New CBalanceAccum
'If DateToStringInt(Ii.DOCUMENT_DATE) = "2005-03-31 00:00:00" Then
'''debug.print
'End If
'If (Ii.PART_ITEM_ID = 7366) And (Ii.LOCATION_ID = 254) And (DateToStringInt(Ii.DOCUMENT_DATE) = "2005-03-31 00:00:00") Then
'''debug.print
'End If

      Ba.PART_ITEM_ID = II.PART_ITEM_ID
      Ba.FROM_DATE = II.DOCUMENT_DATE
      Ba.TO_DATE = II.DOCUMENT_DATE
      Ba.LOCATION_ID = II.LOCATION_ID
      Call Ba.QueryData(1, m_Rs, iCount)
      If m_Rs.EOF Then
         Ba.AddEditMode = SHOW_ADD
      Else
         Call Ba.PopulateFromRS(1, m_Rs)
         Ba.AddEditMode = SHOW_EDIT
      End If
      Ba.DOCUMENT_DATE = II.DOCUMENT_DATE
      Ba.IMPORT_AMOUNT = II.ALL_IMPORT_AMT
      Ba.EXPORT_AMOUNT = II.ALL_EXPORT_AMT
      Ba.BALANCE_AMOUNT = II.BALANCE_AMOUNT
      Ba.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE
      Ba.AVG_PRICE = II.INCLUDE_UNIT_PRICE
      If II.PIG_FLAG = "Y" Then
         Ba.PIG_AGE = GetAge(II.PART_NO, Ba.DOCUMENT_DATE) 'ใช้สำหรับ report M212
      Else
         Ba.PIG_AGE = -1
      End If
      Ba.BATCH_ID = BatchID
      Call Ba.AddEditData
      
      Set Ba = Nothing
   Next II
End Sub

Private Sub InsertMonthlyAccum(BatchID As Long)
Dim Ba As CMonthlyAccum
Dim II As CMonthlyAccum
Dim iCount As Long

   For Each II In m_PartItemsLocationMonthlies
      II.AddEditMode = SHOW_ADD
      II.BATCH_ID = cboBatch.ItemData(Minus2Zero(cboBatch.ListIndex))
      Call II.AddEditData
   Next II
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call LoadBatch(cboBatch)
            
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Pigs)
      Set uctlPartLookup.MyCollection = m_Pigs
      
      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlFromDate.ShowDate = FromDate
      uctlToDate.ShowDate = ToDate
      
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
   pnlHeader.Caption = "ปรับราคาเฉลี่ย"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "จากวันที่")
   Call InitNormalLabel(lblMasterName, "ถึงวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblBatch, "แบต")
   Call InitNormalLabel(lblPartItem, "วัตถุดิบ")
   
'   Call InitCheckBox(chkBalanceFlag, "ลบยอดยกมา")
'   chkBalanceFlag.Value = ssCBUnchecked

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False

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
   Set m_Balances = New Collection
   Set m_PartItemsDateLocations = New Collection
   Set m_PartItemsLocationMonthlies = New Collection
   Set m_Pigs = New Collection
   
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
   Set m_Balances = Nothing
   Set m_PartItemsDateLocations = Nothing
   Set m_PartItemsLocationMonthlies = Nothing
   Set m_Pigs = Nothing
End Sub
Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub
