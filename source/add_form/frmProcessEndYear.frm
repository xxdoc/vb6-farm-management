VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProcessEndYear 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmProcessEndYear.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6555
      Left            =   -120
      TabIndex        =   5
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   11562
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   3
         Top             =   1560
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   2835
         Left            =   240
         TabIndex        =   12
         Top             =   3480
         Width           =   11655
      End
      Begin VB.Label lblProgressDesc 
         Caption         =   "Label1"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   2640
         Width           =   9255
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   1
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessEndYear.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   9
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   2
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmProcessEndYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKClick As Boolean
Public HeaderText As String

Private m_Conn As ADODB.Connection

Private Sub cmdStart_Click()
Dim Status As Boolean
Dim IsOK As Boolean

   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
   
   'Call glbDaily.StartTransaction
      
   Me.Enabled = False
   
   Status = AdjustStockCode
   
   Me.Enabled = True
   
   If Status Then
      'If ConfirmSave Then
         'Call glbDaily.CommitTransaction
         glbErrorLog.LocalErrorMsg = "การอัฟเดดเสร็จสมบูรณ์"
         glbErrorLog.ShowUserError
'      Else
'         Call glbDaily.RollbackTransaction
'         glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
'         glbErrorLog.ShowUserError
'      End If
   Else
'      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
      glbErrorLog.ShowUserError
   End If
   
   OKClick = True
   Unload Me
   Exit Sub
   
End Sub
Private Sub Form_Activate()
      Me.Refresh
      DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
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
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblToDate, "ถึงวันที่", RGB(255, 0, 0))
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
   Call InitNormalLabel(lblProgressDesc, "กรอกวันที่ จากนั้น กดเริ่ม")
   Call InitNormalLabel(lblNote, "ทำเสร็จวันที่ 07/01/2563 Test เรียบร้อย เริ่มใช้งาน จันทร์ 13/01/2563 จะใช้งาน เริ่มตั้งแต่  ถึงวันที่ 31/12/2559" & vbCrLf & "โดยการประมวลผลสิ้นปีนี้มีเงื่อนไขดังนี้" & vbCrLf & "1.ควร COPY ออกมา Test ในเครื่อง StandAlone ก่อน" & vbCrLf & "2.BACKUP / ตรวจสอบยอดหนี้ Stock ก่อนหลังการประมวลผล")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Call EnableForm(Me, False)
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Function AdjustStockCode() As Boolean
Dim I As Long
Dim IsOK As Boolean

Dim BalanceBeforeAdjustColl As Collection
Dim BalanceAferAdjustColl As Collection

Dim TempDate As String
Dim SQL1 As String

Dim Ivd As CInventoryDoc

Dim tempBalance As CBalanceAccum
Dim tempBalanceAfter As CBalanceAccum
Dim Amt As Double
Dim tempImport As CImportItem
Dim tempExport As CExportItem
   
   Set BalanceBeforeAdjustColl = New Collection
   Set BalanceAferAdjustColl = New Collection
   
   I = 0
   prgProgress.MIN = 0
   prgProgress.MAX = 100
   
   AdjustStockCode = False
   
   TempDate = DateToStringIntHi(Trim(uctlToDate.ShowDate))
   
   'Load ยกมา จาก Balance Accum
   Call LoadInventoryBalanceEx(Nothing, BalanceBeforeAdjustColl, DateAdd("D", 1, uctlToDate.ShowDate))
   
   prgProgress.Value = 1
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   '/*REMOVE CONSTRAINTS*/
   '/*CREATE CONSTRAINTS CASCADE OR SET NULL*/
   lblProgressDesc.Caption = "กำลัง REMOVE CONSTRAINTS AND CREATE CONSTRAINTS CASCADE OR SET NULL"
   Me.Refresh
   DoEvents
   
' ของเดิมไม่มี
'   SQL1 = "ALTER TABLE RECEIPT_ITEM DROP CONSTRAINT RECEIPT_ITEM_BILLING_DOC_ID_FK;"
'   Call m_Conn.Execute(SQL1)
' ของเดิมไม่มี
'   SQL1 = "ALTER TABLE RECEIPT_CNDN DROP CONSTRAINT RECEIPT_CNDN_BILLING_DOC_ID_FK;"
'   Call m_Conn.Execute(SQL1)
' ของเดิมไม่มี
'   SQL1 = "ALTER TABLE RECEIPT_CNDN DROP CONSTRAINT RECEIPT_CNDN_DO_ID_FK;"
'   Call m_Conn.Execute(SQL1)
' ของเดิมไม่มี
'   SQL1 = "ALTER TABLE EXPENSE_RATIO DROP CONSTRAINT EXPENSE_RATIO_RO_ITEM_ID_FK;"
'   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_BILLING_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_ITEM DROP CONSTRAINT RECEIPT_ITEM_DO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_DO_ID_FK FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_BILLING_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_DO_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE DO_ITEM DROP CONSTRAINT DO_ITEM_DO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_DO_ID_FK FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RO_ITEM DROP CONSTRAINT RO_ITEM_RO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RO_ITEM ADD CONSTRAINT RO_ITEM_RO_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE CASH_TRAN DROP CONSTRAINT CASH_TRAN_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE CASH_TRAN DROP CONSTRAINT CASH_TRAN_CASH_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CASH_DOC_FK FOREIGN KEY (CASH_DOC_ID) REFERENCES CASH_DOC(CASH_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE REVENUE_COST_ITEM DROP CONSTRAINT REVENUE_CTI_REVENUE_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE REVENUE_COST_ITEM ADD CONSTRAINT REVENUE_CTI_REVENUE_ID_FK FOREIGN KEY (REVENUE_COST_ID) REFERENCES REVENUE_COST(REVENUE_COST_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   
   prgProgress.Value = 3
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents

   
   SQL1 = "DELETE FROM EXPENSE_RATIO EPR"
   SQL1 = SQL1 & " " & "WHERE "
   SQL1 = SQL1 & " " & "EPR.RO_ITEM_ID NOT IN "
   SQL1 = SQL1 & " " & "("
   SQL1 = SQL1 & " " & "SELECT DISTINCT RI.RO_ITEM_ID FROM RO_ITEM RI "
   SQL1 = SQL1 & " " & ")"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE EXPENSE_RATIO ADD CONSTRAINT EXPENSE_RATIO_RO_ITEM_ID_FK FOREIGN KEY (RO_ITEM_ID) REFERENCES RO_ITEM(RO_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   
   prgProgress.Value = 5
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents

   SQL1 = "ALTER TABLE EXPORT_ITEM DROP CONSTRAINT EXPORT_ITEM_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE EXPORT_ITEM ADD CONSTRAINT EXPORT_ITEM_DOC_ID_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE IMPORT_ITEM DROP CONSTRAINT IMPORT_ITEM_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE IMPORT_ITEM ADD CONSTRAINT IMPORT_ITEM_DOC_ID_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   
   lblProgressDesc.Caption = "ลบเอกสารใบขายสด"
   prgProgress.Value = 15
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*KAY SOD*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 2 AND BD.RECEIPT_TYPE = 1) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
''''''   lblProgressDesc.Caption = "ลบเอกสารใบขายเชื่อ"
''''''   '/*INVOICE*/

Dim m_BillingDoc As CBillingDoc
Dim TempRs As ADODB.Recordset
Dim itemcount As Long
Dim m_PaidAmounts  As Collection
Dim m_DnAmounts As Collection
Dim m_CnAmounts As Collection
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem

   Set TempRs = New ADODB.Recordset
   Set m_DnAmounts = New Collection
   Set m_CnAmounts = New Collection
   Set m_PaidAmounts = New Collection
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.COMMIT_FLAG = ""
   m_BillingDoc.TO_DATE = uctlToDate.ShowDate
   m_BillingDoc.DOCUMENT_TYPE = 1
   'm_BillingDoc.VALID_DATE = DateAdd("D", 1, uctlToDate.ShowDate)
   m_BillingDoc.ItemSumFlag = True
   m_BillingDoc.OrderType = 1
   
   Call m_BillingDoc.SetFlag(False, True, False, False, False, False)
   
   Call glbDaily.QueryBillingDoc(m_BillingDoc, TempRs, itemcount, IsOK, glbErrorLog)
      
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, uctlToDate.ShowDate, , , , uctlToDate.ShowDate)
   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, -1, uctlToDate.ShowDate, 3, 2, uctlToDate.ShowDate)
   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, -1, uctlToDate.ShowDate, 4, 2, uctlToDate.ShowDate)
   
   Dim Percent As Double
   I = 0
   While Not TempRs.EOF
         I = I + 1
         Percent = 15 + MyDiff(I, itemcount) * 35
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
                  
         Call m_BillingDoc.PopulateFromRS(1, TempRs)
         
         Set Ri1_0 = GetReceiptItem(m_PaidAmounts, m_BillingDoc.BILLING_DOC_ID) 'รับชำระ
         Set Ri1_1 = GetReceiptItem(m_DnAmounts, m_BillingDoc.BILLING_DOC_ID) 'เพิ่มหนี้
         Set Ri1_2 = GetReceiptItem(m_CnAmounts, m_BillingDoc.BILLING_DOC_ID) 'ลดหนี้
         
         m_BillingDoc.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
         m_BillingDoc.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
         m_BillingDoc.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
        
         If (m_BillingDoc.DO_TOTAL_PRICE + m_BillingDoc.REVENUE_TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT + (m_BillingDoc.DEBIT_AMOUNT - m_BillingDoc.CREDIT_AMOUNT) - m_BillingDoc.PAID_AMOUNT) = 0 Then
            'หนี้เป็น 0 แล้ว ลบเลย
            SQL1 = "DELETE FROM BILLING_DOC BD"
            SQL1 = SQL1 & " " & "WHERE (BD.BILLING_DOC_ID = " & m_BillingDoc.BILLING_DOC_ID & ")"
            Call m_Conn.Execute(SQL1)
            
            SQL1 = "DELETE FROM INVENTORY_DOC IVD"
            SQL1 = SQL1 & " " & "WHERE (IVD.INVENTORY_DOC_ID = " & m_BillingDoc.INVENTORY_DOC_ID & ")"
            Call m_Conn.Execute(SQL1)
         End If
         Me.Refresh
         TempRs.MoveNext
      Wend


   lblProgressDesc.Caption = "ลบเอกสารใบเสร็จรับชำระ"
   prgProgress.Value = 55
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents

   '/*RECEIPT*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 2 OR BD.DOCUMENT_TYPE = 3 OR BD.DOCUMENT_TYPE = 4) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & "AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0);"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบปันต้นทุน"
   prgProgress.Value = 59
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 5 OR BD.DOCUMENT_TYPE = 6 OR BD.DOCUMENT_TYPE = 7) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบนำฝากธนาคาร"
   prgProgress.Value = 60
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   SQL1 = "DELETE FROM CASH_DOC CD "
   SQL1 = SQL1 & " " & "WHERE CD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบปันต้นทุนจากรายได้อื่นๆ"
   prgProgress.Value = 63
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   SQL1 = "DELETE FROM REVENUE_COST RVNC "
   SQL1 = SQL1 & " " & "WHERE RVNC.REVENUE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบ STOCK ที่อ้างอิงจากใบขายเชื่อ ขายสด"
   prgProgress.Value = 65
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents

   SQL1 = "DELETE FROM BALANCE_ACCUM BA"
   SQL1 = SQL1 & " " & "WHERE BA.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)

   SQL1 = "DELETE FROM MONTHLY_ACCUM BA"
   SQL1 = SQL1 & " " & "WHERE BA.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "เอกสารใบรับเข้า เบิกออก โอนย้าย ปรับยอด คลัง"
   prgProgress.Value = 67
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents

''''''   '/*IVTRD IMPORT EXPORT*/
   SQL1 = "DELETE FROM INVENTORY_DOC IVTRD"
   SQL1 = SQL1 & " " & "WHERE (IVTRD.DOCUMENT_TYPE = 1 Or IVTRD.DOCUMENT_TYPE = 2 Or IVTRD.DOCUMENT_TYPE = 3 Or IVTRD.DOCUMENT_TYPE = 4) AND IVTRD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)

   lblProgressDesc.Caption = "เอกสารใบรับเข้า เบิกออก โอนย้าย ปรับยอด สุกร"
   prgProgress.Value = 70
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*IVTRD IMPORT EXPORT*/
   SQL1 = "DELETE FROM INVENTORY_DOC IVTRD"
   SQL1 = SQL1 & " " & "WHERE (IVTRD.DOCUMENT_TYPE IN (5,9,11,6,7,8,12,888)) AND IVTRD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*REMOVE CONSTRAINTS*/
   '/*CREATE CONSTRAINTS*/
   lblProgressDesc.Caption = "กำลัง REMOVE CONSTRAINTS AND CREATE CONSTRAINTS"
   prgProgress.Value = 75
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE RECEIPT_ITEM DROP CONSTRAINT RECEIPT_ITEM_BILLING_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE RECEIPT_CNDN DROP CONSTRAINT RECEIPT_CNDN_BILLING_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE RECEIPT_CNDN DROP CONSTRAINT RECEIPT_CNDN_DO_ID_FK;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE EXPENSE_RATIO DROP CONSTRAINT EXPENSE_RATIO_RO_ITEM_ID_FK;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_BILLING_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_ITEM DROP CONSTRAINT RECEIPT_ITEM_DO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_DO_ID_FK FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_BILLING_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_DO_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE DO_ITEM DROP CONSTRAINT DO_ITEM_DO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_DO_ID_FK FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RO_ITEM DROP CONSTRAINT RO_ITEM_RO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RO_ITEM ADD CONSTRAINT RO_ITEM_RO_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE CASH_TRAN DROP CONSTRAINT CASH_TRAN_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE CASH_TRAN DROP CONSTRAINT CASH_TRAN_CASH_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CASH_DOC_FK FOREIGN KEY (CASH_DOC_ID) REFERENCES CASH_DOC(CASH_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE REVENUE_COST_ITEM DROP CONSTRAINT REVENUE_CTI_REVENUE_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE REVENUE_COST_ITEM ADD CONSTRAINT REVENUE_CTI_REVENUE_ID_FK FOREIGN KEY (REVENUE_COST_ID) REFERENCES REVENUE_COST(REVENUE_COST_ID);"
   Call m_Conn.Execute(SQL1)
      
   
   lblProgressDesc.Caption = "กำลัง ปรับยอด STOCK"
   prgProgress.Value = 80
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
''''''   'LOAD STOCK BALANCE AFTER ADJUST END YEAR
   Call LoadInventoryBalanceEx(Nothing, BalanceAferAdjustColl, DateAdd("D", 1, uctlToDate.ShowDate))
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
    Ivd.DOCUMENT_DATE = uctlToDate.ShowDate
   Ivd.DOCUMENT_NO = "***ENDYEAR_" & uctlToDate.ShowDate
   Ivd.DELIVERY_FEE = 0
   Ivd.DOCUMENT_TYPE = 4
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   
   I = 0
   For Each tempBalance In BalanceBeforeAdjustColl
      I = I + 1
      txtPercent.Text = 80 + (MyDiff(I, BalanceBeforeAdjustColl.Count) * 10)
      prgProgress.Value = Val(txtPercent.Text)
      Me.Refresh
      DoEvents
      
      Set tempBalanceAfter = GetObject("CLotItem", BalanceAferAdjustColl, Trim(tempBalance.LOCATION_ID & "-" & tempBalance.PART_ITEM_ID))
   
      If (tempBalance.BALANCE_AMOUNT > tempBalanceAfter.BALANCE_AMOUNT) Then
         Set tempImport = New CImportItem
         tempImport.Flag = "A"
         tempImport.PART_ITEM_ID = tempBalance.PART_ITEM_ID
         tempImport.LOCATION_ID = tempBalance.LOCATION_ID
         
         tempImport.ACTUAL_AMOUNT = tempBalance.BALANCE_AMOUNT - tempBalanceAfter.BALANCE_AMOUNT
         tempImport.ACTUAL_PRICE = tempBalance.TOTAL_INCLUDE_PRICE - tempBalanceAfter.TOTAL_INCLUDE_PRICE
   
      Else
            
      End If
      
''''''      TempLotItem.Flag = "A"
''''''      TempLotItem.PART_ITEM_ID = m_LotItem.PART_ITEM_ID
''''''      TempLotItem.LOCATION_ID = m_LotItem.LOCATION_ID
''''''
''''''      Set TempLotItemSearch = GetObject("CLotItem", InventoryBalAferAdjustColl, Trim(m_LotItem.LOCATION_ID & "-" & m_LotItem.PART_ITEM_ID))
''''''      If (m_LotItem.SUM_AMOUNT > TempLotItemSearch.SUM_AMOUNT) Then  'ก่อนปรับมากกว่าหลังปรับ ต้องปรับเพิ่ม
''''''         TempLotItem.TX_AMOUNT = m_LotItem.SUM_AMOUNT - TempLotItemSearch.SUM_AMOUNT
''''''         TempLotItem.MULTIPLIER = 1
''''''         TempLotItem.TX_TYPE = "I"
''''''      ElseIf (m_LotItem.SUM_AMOUNT < TempLotItemSearch.SUM_AMOUNT) Then
''''''         TempLotItem.TX_AMOUNT = TempLotItemSearch.SUM_AMOUNT - m_LotItem.SUM_AMOUNT
''''''         TempLotItem.MULTIPLIER = -1
''''''         TempLotItem.TX_TYPE = "E"
''''''      End If
''''''      Set Pi = GetObject("CStockCode", TempPartColl, Trim(Str(m_LotItem.PART_ITEM_ID)))
''''''      TempLotItem.UNIT_TRAN_ID = Pi.UNIT_CHANGE_ID
''''''      TempLotItem.UNIT_MULTIPLE = 1
''''''
''''''      Call Ivd.ImportExportItems.Add(TempLotItem)
''''''
''''''      Set TempLotItem = Nothing
   Next tempBalance
   
''''''   If Ivd.ImportExportItems.Count > 0 Then
''''''      If Not glbDaily.AddEditInventoryDoc(Ivd, True, False, glbErrorLog) Then
''''''         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
''''''         Exit Function
''''''      End If
''''''   End If
   
   Set Ivd = Nothing
   
   prgProgress.Value = prgProgress.MAX
   txtPercent.Text = 100
   Me.Refresh
   DoEvents
   
   Set BalanceBeforeAdjustColl = Nothing
   Set BalanceAferAdjustColl = Nothing
   
   AdjustStockCode = True
End Function
