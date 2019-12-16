VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmInitBalanceEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmInitBalanceEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9465
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3705
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6535
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtProcess 
         Height          =   465
         Left            =   1980
         TabIndex        =   1
         Top             =   1560
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   1980
         TabIndex        =   7
         Top             =   2070
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtProgress 
         Height          =   465
         Left            =   1980
         TabIndex        =   8
         Top             =   2310
         Width           =   1695
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1980
         TabIndex        =   0
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1230
         Width           =   1755
      End
      Begin VB.Label lblPercent 
         Caption         =   "1"
         Height          =   255
         Left            =   3780
         TabIndex        =   11
         Top             =   2430
         Width           =   225
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2430
         Width           =   1755
      End
      Begin VB.Label lblProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1740
         Width           =   1755
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   570
         TabIndex        =   2
         Top             =   2970
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7335
         TabIndex        =   5
         Top             =   2970
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5685
         TabIndex        =   3
         Top             =   2970
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInitBalanceEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Conn As ADODB.Connection

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset
Private m_Rs4 As ADODB.Recordset

Private m_InventoryBalances  As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
   ProgressBar1.MIN = 0
   ProgressBar1.MAX = 100
   ProgressBar1.Value = 0
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction
   
   Call GenerateBalance
   
   Call glbDaily.CommitTransaction
   
   Call EnableForm(Me, True)
End Sub
Private Sub GenerateBalance()
Dim I As Long
Dim IsOK As Boolean
Dim Percent As Double
Dim iCount2 As Long
   
Dim Bd As CBillingDoc
Dim Di As CDoItem
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim Rs As ADODB.Recordset
Dim m_PaidAmounts As Collection
Dim m_DnAmounts As Collection
Dim m_CnAmounts As Collection
Dim m_DiItems As Collection
Dim TempDate  As String

Dim Ivd As CInventoryDoc
Dim TotalPrice As Double
Dim SQL1 As String
   
   Dim FirstDate As Date
   Dim LastDate As Date
   
   Call GetFirstLastDate(Now, FirstDate, LastDate)
   
   txtProcess.Text = "สร้างข้อมูลยอดยกมาลูกหนี้"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "LOADING"
   txtProgress.Refresh
   
   Set Rs = New ADODB.Recordset
   Set Bd = New CBillingDoc
   Set m_PaidAmounts = New Collection
   Set m_DnAmounts = New Collection
   Set m_CnAmounts = New Collection
   Set m_DiItems = New Collection
   I = 0
   Bd.BILLING_DOC_ID = -1
   Bd.DOCUMENT_TYPE = 1
   'Bd.ItemSumFlag = True
   Bd.QueryFlag = -1
   Bd.OrderType = 1
   Bd.TO_DATE = DateAdd("D", -1, uctlFromDate.ShowDate)
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, , Bd.TO_DATE)
   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, , Bd.TO_DATE, 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, , Bd.TO_DATE, 4, 2)
   Call LoadTotalPriceByBill(Nothing, m_DiItems, , Bd.TO_DATE)
   
   Call glbDaily.QueryBillingDoc(Bd, Rs, iCount2, IsOK, glbErrorLog)
   
   txtProcess.Text = "สร้างข้อมูลยอดยกมาเงินสดในมือ"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "LOADING"
   txtProgress.Refresh
   
   Dim m_CashTranBlBalances As Collection
   Dim m_CashTranCdBalances As Collection
   Dim m_CashTranBlBalancesAfter As Collection
   Dim m_CashTranCdBalancesAfter As Collection
   
   Set m_CashTranBlBalances = New Collection
   Set m_CashTranCdBalances = New Collection
   Set m_CashTranBlBalancesAfter = New Collection
   Set m_CashTranCdBalancesAfter = New Collection
      
   Call LoadRemainMoneyBl(Nothing, m_CashTranBlBalances, , DateAdd("D", -1, uctlFromDate.ShowDate))
   Call LoadRemainMoneyCd(Nothing, m_CashTranCdBalances, , DateAdd("D", -1, uctlFromDate.ShowDate))
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   txtProcess.Text = "สร้างข้อมูลยอดยกมาลูกหนี้"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "PROCESSING"
   txtProgress.Refresh
   
   SQL1 = "DELETE FROM RECEIPT_CNDN"
   m_Conn.Execute (SQL1)
   
   I = 0
   Set Ivd = New CInventoryDoc
   While Not Rs.EOF
      Call Bd.PopulateFromRS(1, Rs)
      I = I + 1
      Percent = MyDiffEx(I, iCount2) * 100
      ProgressBar1.Value = Percent
      txtProgress.Text = FormatNumber(Percent)
      txtProgress.Refresh
      
      Set Di = GetDoItem(m_DiItems, Bd.BILLING_DOC_ID) 'ยอดเงิน
      Set Ri1_0 = GetReceiptItem(m_PaidAmounts, Bd.BILLING_DOC_ID) 'รับชำระ
      Set Ri1_1 = GetReceiptItem(m_DnAmounts, Bd.BILLING_DOC_ID) 'เพิ่มหนี้
      Set Ri1_2 = GetReceiptItem(m_CnAmounts, Bd.BILLING_DOC_ID) 'ลดหนี้
      
      Bd.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
      Bd.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
      Bd.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
'      ''debug.print (Bd.RECEIPT_PAID_AMOUNT)
'      ''debug.print (Bd.CNDN_TOTAL_PRICE)
'      ''debug.print (Bd.DO_TOTAL_PRICE)
      
      TotalPrice = Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT) - Bd.PAID_AMOUNT
      
      If Round(TotalPrice, 2) = 0 Then
         
         Call Ri1_0.DeleteDataFromDoID
         
         Call Bd.DeleteData
         
         Ivd.INVENTORY_DOC_ID = Bd.INVENTORY_DOC_ID
         If Ivd.INVENTORY_DOC_ID > 0 Then
            Call Ivd.DeleteData
         End If
         
      End If
      Rs.MoveNext
   Wend
   
   TempDate = DateToStringIntHi(Trim(DateAdd("D", -1, uctlFromDate.ShowDate)))
   
   SQL1 = "DELETE FROM DO_ITEM UG WHERE UG.DO_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 2 AND (BD.DOCUMENT_SUBTYPE IN (1,2)) AND DOCUMENT_DATE <= '" & TempDate & "') "
   m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
   
   SQL1 = "DELETE FROM CASH_TRAN CT WHERE CT.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 2 AND (BD.DOCUMENT_SUBTYPE IN (1,2)) AND DOCUMENT_DATE <= '" & TempDate & "') "
   m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
   
   SQL1 = "DELETE FROM BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 2 AND (BD.DOCUMENT_SUBTYPE IN (1,2)) AND DOCUMENT_DATE <= '" & TempDate & "' "
   m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
   
   SQL1 = "DELETE FROM CASH_TRAN CT WHERE CT.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 2 AND (DOCUMENT_SUBTYPE = 0 ) AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0)) "
   m_Conn.Execute (SQL1)               'ใบเสร็จขายเชื่อที่ไม่มี Item ลูกแล้ว
   
   SQL1 = "DELETE FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 3  AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0) "
   m_Conn.Execute (SQL1)               'ใบเพิ่มหนี้ที่ไม่มี Item ลูกแล้ว
   
   SQL1 = "DELETE FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 4  AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0) "
   m_Conn.Execute (SQL1)               'ใบลดหนี้ที่ไม่มี Item ลูกแล้ว
   
   SQL1 = "DELETE FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 2 AND (DOCUMENT_SUBTYPE = 0 ) AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0) "
   m_Conn.Execute (SQL1)               'ใบเสร็จขายเชื่อที่ไม่มี Item ลูกแล้ว
   
   SQL1 = "DELETE FROM CASH_TRAN CT WHERE CT.CASH_DOC_ID IN (SELECT BD.CASH_DOC_ID FROM  CASH_DOC BD WHERE DOCUMENT_DATE <= '" & TempDate & "') "
   m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
   
   SQL1 = "DELETE FROM CASH_DOC BD WHERE DOCUMENT_DATE <= '" & TempDate & "'"
   m_Conn.Execute (SQL1)               'TABLE CASH_DOC
   
   SQL1 = "DELETE FROM RO_ITEM RO WHERE RO.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE DOCUMENT_TYPE IN (5,6,7) AND DOCUMENT_DATE <= '" & TempDate & "') "
   m_Conn.Execute (SQL1)               'ใบปันต้นทุน
   
   SQL1 = "DELETE FROM BILLING_DOC BD WHERE DOCUMENT_TYPE IN (5,6,7) AND DOCUMENT_DATE <= '" & TempDate & "'"
   m_Conn.Execute (SQL1)         'ใบบันต้นทุน
   
   SQL1 = "DELETE FROM LOGIN_TRACKING"
   m_Conn.Execute (SQL1)         'LOGIN
   
   Set Bd = Nothing
   Set Ri1_0 = Nothing
   Set Ri1_1 = Nothing
   Set Ri1_2 = Nothing
   Set Rs = Nothing
   Set m_PaidAmounts = Nothing
   Set m_DnAmounts = Nothing
   Set m_CnAmounts = Nothing
   
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim Ct1 As CCashTran
   Dim Ct2 As CCashTran
   Dim Ct3 As CCashTran
   Dim Ct4 As CCashTran
   Dim Ct5 As CCashTran
   Dim Ct6 As CCashTran
   
   txtProcess.Text = "สร้างข้อมูลยอดยกมาเงินสดในมือ"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "PROCESSING"
   txtProgress.Refresh
      
   Call LoadRemainMoneyBl(Nothing, m_CashTranBlBalancesAfter, , DateAdd("D", -1, uctlFromDate.ShowDate))
   Call LoadRemainMoneyCd(Nothing, m_CashTranCdBalancesAfter, , DateAdd("D", -1, uctlFromDate.ShowDate))
   
   'FormatNumber(Ct1.GetFieldValue("NET_AMOUNT") + Ct2.GetFieldValue("NET_AMOUNT") - Ct3.GetFieldValue("AMOUNT"))
   
   Dim Cd As CCashDoc
   Dim Ct As CCashTran
   
   Set Cd = New CCashDoc
   
   Cd.ShowMode = SHOW_ADD
   Call Cd.SetFieldValue("CASH_DOC_ID", -1)
   Call Cd.SetFieldValue("DOCUMENT_DATE", DateAdd("D", -1, uctlFromDate.ShowDate))
   Call Cd.SetFieldValue("DOCUMENT_NO", "ยกมา" & DateAdd("D", -1, uctlFromDate.ShowDate))
   Call Cd.SetFieldValue("DOCUMENT_TYPE", CASH_DEPOSIT)
   Call Cd.SetFieldValue("BANK_ID", -1)
   Call Cd.SetFieldValue("BANK_BRANCH", -1)
   Call Cd.SetFieldValue("BANK_ACCOUNT", -1)
   Call Cd.SetFieldValue("EMP_ID", -1)
   Call Cd.SetFieldValue("CUSTOMER_ID", -1)
        
   For I = 1 To 5
      Set Ct = New CCashTran
      
      Set Ct1 = GetCashTran(m_CashTranBlBalances, Trim(I & "-1"))
      Set Ct2 = GetCashTran(m_CashTranBlBalances, Trim(I & "-3"))
      Set Ct3 = GetCashTran(m_CashTranCdBalances, Trim(Str(I)))
      Set Ct4 = GetCashTran(m_CashTranBlBalancesAfter, Trim(I & "-1"))
      Set Ct5 = GetCashTran(m_CashTranBlBalancesAfter, Trim(I & "-3"))
      Set Ct6 = GetCashTran(m_CashTranBlBalancesAfter, Trim(Str(I)))
      
      TotalPrice = (Ct1.GetFieldValue("NET_AMOUNT") + Ct2.GetFieldValue("NET_AMOUNT") - Ct3.GetFieldValue("AMOUNT")) - (Ct4.GetFieldValue("NET_AMOUNT") + Ct5.GetFieldValue("NET_AMOUNT") - Ct6.GetFieldValue("AMOUNT"))
      
      If Round(TotalPrice, 2) <> 0 Then
         Ct.Flag = "A"
         Call Ct.SetFieldValue("PAYMENT_TYPE", I) 'ออกเป็นเงินสด
         Call Ct.SetFieldValue("PAYMENT_TYPE_NAME", PaymentType2Text(I))
         Call Ct.SetFieldValue("AMOUNT", -TotalPrice)
         Call Ct.SetFieldValue("NET_AMOUNT", -TotalPrice)
         Call Ct.SetFieldValue("TX_TYPE", "E")
         Call Cd.CashTranItems.Add(Ct)
      End If
      Set Ct = Nothing
   Next I
   Call glbDaily.AddEditCashDoc(Cd, IsOK, False, glbErrorLog)
      
   Set Cd = Nothing
   Percent = 100
   ProgressBar1.Value = Percent
   txtProgress.Text = FormatNumber(Percent)
   txtProgress.Refresh
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
         
      ShowMode = SHOW_ADD
      ID = -1
      
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "ตั้งยอดยกมาระบบใหม่"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call txtProcess.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtProcess.Enabled = False
   Call txtProgress.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtProgress.Enabled = False
   
   Call InitNormalLabel(lblProcess, "โปรเซส")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "%")
   Call InitNormalLabel(lblFromDate, "วันที่ตั้งยอด")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
   Set m_Rs4 = New ADODB.Recordset
   
   Set m_InventoryBalances = New Collection
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_InventoryBalances = Nothing
   
   Set m_Rs1 = Nothing
   Set m_Rs2 = Nothing
   Set m_Rs3 = Nothing
   Set m_Rs4 = Nothing
End Sub
Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub
