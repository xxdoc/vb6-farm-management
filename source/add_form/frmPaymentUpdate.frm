VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaymentUpdate 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmPaymentUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
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
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1050
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Top             =   2430
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
         TabIndex        =   5
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1500
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtCashBalance 
         Height          =   465
         Left            =   1860
         TabIndex        =   3
         Top             =   1950
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin VB.Label lblCashBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   16
         Top             =   2070
         Width           =   1275
      End
      Begin Threed.SSCheck chkBalanceFlag 
         Height          =   375
         Left            =   6450
         TabIndex        =   1
         Top             =   1110
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   6
         Top             =   3450
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPaymentUpdate.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   15
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   2490
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   2910
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1110
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   8
         Top             =   3450
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
         Top             =   3450
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPaymentUpdate.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmPaymentUpdate"
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
      TempII.LOCATION_ID = O.LOCATION_ID
      TempII.PART_ITEM_ID = O.PART_ITEM_ID
      TempII.DOCUMENT_DATE = O.DOCUMENT_DATE
      TempII.BALANCE_AMOUNT = ImpI.CURRENT_AMOUNT
      TempII.TOTAL_INCLUDE_PRICE = ImpI.TOTAL_INCLUDE_PRICE
'TempII.INCLUDE_UNIT_PRICE = MyDiffEx(ImpI.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, ImpI.CURRENT_AMOUNT + O.IMPORT_AMOUNT)
TempII.INCLUDE_UNIT_PRICE = MyDiffEx(ImpI.TOTAL_INCLUDE_PRICE, ImpI.CURRENT_AMOUNT)
If O.TX_TYPE = "I" Then
TempII.ALL_IMPORT_AMT = O.IMPORT_AMOUNT
ElseIf O.TX_TYPE = "E" Then
TempII.ALL_EXPORT_AMT = O.EXPORT_AMOUNT
End If
'TempII.BALANCE_AMOUNT = ImpI.IMPORT_AMOUNT  'TempII.BALANCE_AMOUNT + O.IMPORT_AMOUNT
'      If O.TX_TYPE = "I" Then
'TempII.INCLUDE_UNIT_PRICE = MyDiffEx(ImpI.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, ImpI.CURRENT_AMOUNT + O.IMPORT_AMOUNT)
'         TempII.ALL_IMPORT_AMT = O.IMPORT_AMOUNT
'         TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
'         TempII.BALANCE_AMOUNT = TempII.BALANCE_AMOUNT + O.IMPORT_AMOUNT
'      ElseIf O.TX_TYPE = "E" Then
'TempII.INCLUDE_UNIT_PRICE = 0
'         TempII.ALL_EXPORT_AMT = O.EXPORT_AMOUNT
'         TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE - O.EXPORT_TOTAL_PRICE
'         TempII.BALANCE_AMOUNT = TempII.BALANCE_AMOUNT - O.EXPORT_AMOUNT
'      End If
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

'Public Function GetBalanceItem(Col As Collection, PartItemID As Long, LocationID As Long, DocDate As Date) As Object
'Dim D As Object
'Dim Key As String
'Dim MaxSeq As Long
'Dim i As Long
'Dim MaxIndex As Long
'Static II As CImportItem
'Dim MaxDate As Date
'
'   MaxDate = -2
'   For Each D In Col
''Debug.Print D.TX_TYPE & ";" & D.PART_ITEM_ID & ";" & D.LOCATION_ID & ";" & DateToStringInt(D.DOCUMENT_DATE) & ";" & D.CURRENT_AMOUNT
'      If (DateToStringInt(D.DOCUMENT_DATE) < DateToStringInt(DocDate)) And (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) Then
'         If DateToStringInt(D.DOCUMENT_DATE) > DateToStringInt(MaxDate) Then
'            MaxDate = InternalDateToDate(DateToStringInt(D.DOCUMENT_DATE))
'         End If
'      End If
'   Next D
'
''If MaxDate <= 0 Then
''Debug.Print
''End If
'
'   i = 0
'   MaxSeq = -1
'   MaxIndex = -1
'   For Each D In Col
'      i = i + 1
'
'      If (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) And _
'         (DateToStringInt(D.DOCUMENT_DATE) = DateToStringInt(MaxDate)) Then
'            If D.TRANSACTION_SEQ > MaxSeq Then
'               MaxSeq = D.TRANSACTION_SEQ
'               MaxIndex = i
'            End If
'      End If
'   Next D
'
'   If MaxIndex > 0 Then
'      Set GetBalanceItem = Col(MaxIndex)
'   Else
'      If II Is Nothing Then
'         Set II = New CImportItem
'      End If
'      Set GetBalanceItem = II
'   End If
'End Function

Private Sub CopyBalanceAccum(Src As Collection, Dst As Collection)
Dim II As CImportItem
Dim Ba As CBalanceAccum
   
   For Each Ba In Src
      Set II = New CImportItem
      II.LOCATION_ID = Ba.LOCATION_ID
      II.PART_ITEM_ID = Ba.PART_ITEM_ID
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

Private Sub cmdStart_Click()
'On Error GoTo ErrHandler
Dim Percent As Double
Dim MIN As Double
Dim MAX As Double
Dim RecordCount As Double
Dim RName As String
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim IsOK As Boolean
Dim Amt As Double
Dim Rs1 As ADODB.Recordset
Dim TxCode As String
Dim iCount As Long
Dim HasBegin As Boolean
Dim Bd As CBillingDoc
Dim TempBd As CBillingDoc
Dim Count1 As Long
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim BalanceDate As Date

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   HasBegin = False
   
   Call EnableForm(Me, False)
      
   Set Rs1 = New ADODB.Recordset
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   Call glbDaily.StartTransaction
   
   Set Pm = New CPayment
   Call Pm.ClearPayment
   Set Pm = Nothing
   
   BalanceDate = DateAdd("D", -1, uctlFromDate.ShowDate)
   Set Pm = New CPayment
   Pm.AddEditMode = SHOW_ADD
   Pm.PAYMENT_NO = "เงินสดยกมา"
   Pm.PAYMENT_DATE = BalanceDate
   Pm.COMMIT_FLAG = "N"
   Pm.INTERNAL_FLAG = "Y"
   Pm.TX_TYPE = "I"
   Pm.TOTAL_AMOUNT = Val(txtCashBalance.Text)
   
   Set Pmi = New CPaymentItem
   Pmi.Flag = "A"
   Pmi.PAYMENT_TYPE = CASH_PMT
   Pmi.PAY_AMOUNT = Pm.TOTAL_AMOUNT
   Call Pm.PaymentItems.Add(Pmi)
   Set Pmi = Nothing
   
   Call glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog)
   Set Pm = Nothing

   '=== Detail
   Set Bd = New CBillingDoc
   Bd.BILLING_DOC_ID = -1
   Bd.FROM_DATE = uctlFromDate.ShowDate
   Bd.TO_DATE = uctlToDate.ShowDate
   Bd.DOCUMENT_TYPE = 2 'ใบเสร็จรับเงิน
   Bd.COMMIT_FLAG = ""
   Call glbDaily.QueryBillingDoc(Bd, Rs1, Count1, IsOK, glbErrorLog)

   MIN = 0
   MAX = 100
   Percent = 0
   RecordCount = 0
   prgProgress.MIN = MIN
   prgProgress.MAX = MAX

   While Not Rs1.EOF
      RecordCount = RecordCount + 1
      Call Bd.PopulateFromRS(1, Rs1)
      
      Set TempBd = New CBillingDoc
      Set Pm = New CPayment
      
      TempBd.BILLING_DOC_ID = Bd.BILLING_DOC_ID
      TempBd.QueryFlag = 1
      Call glbDaily.QueryBillingDoc(TempBd, m_Rs, iCount, IsOK, glbErrorLog)
      Call TempBd.PopulateFromRS(1, m_Rs)
      Call glbDaily.DO2Payment(TempBd, Pm)

      Pm.AddEditMode = SHOW_EDIT
      Call glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog)

      Set Pm = Nothing
      Set TempBd = Nothing

      Percent = MyDiff(RecordCount, Count1) * 100
      prgProgress.Value = Percent
      prgProgress.Refresh
      txtPercent.Text = Format(Percent, "0.00")
      txtPercent.Refresh

      DoEvents
      Rs1.MoveNext
   Wend

   Call glbDaily.CommitTransaction

   txtPercent.Text = Format(100, "0.00")
   prgProgress.Value = 100
   HasBegin = False

   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   Set Bd = Nothing
   
   Call EnableForm(Me, True)
   
   Exit Sub
   
'ErrHandler:
'   If HasBegin Then
'      glbDaily.RollbackTransaction
'   End If
'   glbErrorLog.LocalErrorMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub InsertBalanceAccum()
Dim Ba As CBalanceAccum
Dim II As CImportItem
Dim iCount As Long

   For Each II In m_PartItemsDateLocations
      Set Ba = New CBalanceAccum
'If DateToStringInt(Ii.DOCUMENT_DATE) = "2005-03-31 00:00:00" Then
'Debug.Print
'End If
'If (Ii.PART_ITEM_ID = 7366) And (Ii.LOCATION_ID = 254) And (DateToStringInt(Ii.DOCUMENT_DATE) = "2005-03-31 00:00:00") Then
'Debug.Print
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
      Call Ba.AddEditData
      
      Set Ba = Nothing
   Next II
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlFromDate.ShowDate = FromDate
      uctlToDate.ShowDate = ToDate
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
      uctlFromDate.ShowDate = InternalDateToDate("2005-10-01 00:00:00")
      
      If glbEnterPrise.BRANCH_CODE = "MA2" Then
         txtCashBalance.Text = 903465
      ElseIf glbEnterPrise.BRANCH_CODE = "MH" Then
         txtCashBalance.Text = 484175
      ElseIf glbEnterPrise.BRANCH_CODE = "DTS" Then
         txtCashBalance.Text = 114916
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
   pnlHeader.Caption = "สร้างข้อมูลการชำระเงิน"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   Call InitNormalLabel(lblFileName, "จากวันที่ใบเสร็จ")
   Call InitNormalLabel(lblMasterName, "ถึงวันที่ใบเสร็จ")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblCashBalance, "เงินสดยกมา")
   Call InitNormalLabel(Label2, "บาท")

   Call InitCheckBox(chkBalanceFlag, "เคลียร์ข้อมูลการชำระเงินทั้งหมด")

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtCashBalance.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   uctlToDate.Enable = False
   
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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
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
End Sub

Private Sub txtCashBalance_Change()
   m_HasModify = True
End Sub
