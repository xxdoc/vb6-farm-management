VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportBillingDoc 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportBillingDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1440
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   767
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
         Left            =   0
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
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9780
         Top             =   750
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   8670
         TabIndex        =   1
         Top             =   1470
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportBillingDoc.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   4
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportBillingDoc.frx":2ABC
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
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1500
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   6
         Top             =   2910
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
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportBillingDoc.frx":2DD6
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportBillingDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public FK As Long

Private TempBillingDoc As Collection

Private AccountColls As Collection
Private EmpColls As Collection
Private BankBranchColls As Collection
Private RegionColls As Collection

Private PigStatusTypeColls As Collection
Private RevenueColls As Collection
Private PartColls As Collection
Private LocationColls As Collection

Private FromDate As Date
Private ToDate As Date
Private CountBill As Long
Private CountDown As Double

Private Bl As CBillingDoc
Private Ivd As CInventoryDoc
Private Pm As CPayment
Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.TXT,*.DAT)|*.txt;*.dat;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean

   If Not VerifyTextControl(lblMasterName, txtFileName) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   Call ImportBillingDoc
      
   Call EnableForm(Me, True)
End Sub
Private Sub ImportBillingDoc()
On Error GoTo ErrorHandler
Dim HasBegin As Boolean
Dim FileName As String
Dim F As Long
Dim TempStr As String
Dim LineCount As Long
Dim SuccessCount As Long
Dim ErrorCount As Long
Dim Sum As Long
Dim I As Long
Dim Bi As CBillingDoc
Set Bi = New CBillingDoc

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   FileName = txtFileName.Text
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   Sum = 0
   While Not EOF(F)
      Line Input #F, TempStr
      Sum = Sum + 1
   Wend
   
   HasBegin = True
      
   CountDown = Sum
      
   Call glbDatabaseMngr.DBConnection.BeginTrans
      
   SuccessCount = 0
   ErrorCount = 0
   F = FreeFile()
   Close #F
   Open FileName For Input As #F
   While Not EOF(F)
      I = I + 1
      Line Input #F, TempStr
      prgProgress.Value = MyDiffEx2(I, Sum) * 100
      txtPercent.Text = prgProgress.Value
      LineCount = LineCount + 1
      Me.Refresh
      DoEvents
      
      CountDown = CountDown - 1
      
      If I = 1 Then
         FromDate = DateSerial(Year(TempStr), Month(TempStr), Day(TempStr))
      ElseIf I = 2 Then
         ToDate = DateSerial(Year(TempStr), Month(TempStr), Day(TempStr))
         Bi.FROM_DATE = FromDate
         Bi.TO_DATE = ToDate
         Call LoadBillingDocCode(Bi, Nothing, TempBillingDoc)
         
      Else
         If ProcessLine(TempStr) Then
            SuccessCount = SuccessCount + 1
         Else
            ErrorCount = ErrorCount + 1
         End If
      
      End If
      
   Wend
   Close #F
   prgProgress.Value = 100
   txtPercent.Text = 100
   
   Call glbDatabaseMngr.DBConnection.CommitTrans
   
   glbErrorLog.LocalErrorMsg = "อิมพอร์ต สำเสร็จจำนวน " & SuccessCount & "และ ล้มเหลวจำนวน " & ErrorCount & " ข้อมูล"
   glbErrorLog.ShowUserError
   
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Function ProcessLine(LineStr As String) As Boolean
On Error GoTo ErrorHandler
Dim EmpCode As String
Dim TimeStamp As Date
Dim TimeStampStr As String
Dim TempAsc As Long
Dim OldTempAsc As Long
Dim FirstDate As Date
Dim LastDate As Date
Dim I As Long
Dim ItemCount As Long
Dim IsOK As Boolean
Dim BLD As CBillingDoc
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String

   If Left(LineStr, 2) = "BL" Then
      If CountBill > 0 Then
      
         Call DO2InventoryDoc(Bl, Ivd)
         
'         If Bl.DOCUMENT_TYPE = 2 Then
'            Call glbDaily.DO2Payment(Bl, Pm)
'         End If
         
         Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
      
'         If Bl.DOCUMENT_TYPE = 2 Then
'            Call glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog)
'
'            Bl.PAYMENT_ID = Pm.PAYMENT_ID
'         End If
         
         Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID

         
         Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
         
         Set Bl = New CBillingDoc
         Set Ivd = New CInventoryDoc
      End If
      CountBill = 1
      
      
      Bl.AddEditMode = SHOW_ADD
      
      TempAsc = 3
      OldTempAsc = TempAsc
      
      Dim Acc As CAccount
      TempAsc = InStr(4, LineStr, ";")
      Set Acc = GetAccount(AccountColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
      Bl.ACCOUNT_ID = Acc.ACCOUNT_ID                                                                                                                                 '1
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.CUSTOMER_CODE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '2
      OldTempAsc = TempAsc
      
      If Bl.CUSTOMER_CODE <> Acc.CUSTOMER_CODE Then
         Call MsgBox("ยังไม่มีรหัสลูกค้า " & Bl.CUSTOMER_CODE & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      
      Bl.CUSTOMER_ID = Acc.CUSTOMER_ID
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.DOCUMENT_NO = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '2
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.DOCUMENT_DATE = DateSerial(Left(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), 4), Mid(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), 6, 2), Mid(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1), 9, 2))                                              '3
      OldTempAsc = TempAsc
      
      Set BLD = GetBillingDoc(TempBillingDoc, Trim(Bl.DOCUMENT_NO & "-" & Bl.DOCUMENT_DATE))
      If BLD.BILLING_DOC_ID > 0 Then
         Call glbDaily.DeleteBillingDoc(BLD.BILLING_DOC_ID, IsOK, False, glbErrorLog)
      End If
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.DOCUMENT_TYPE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '4
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.DUE_DATE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '5
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.NOTE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '6
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.TOTAL_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '7
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.VAT_PERCENT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '8
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.VAT_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '9
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.WH_PERCENT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '10
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.WH_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '11
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.DISCOUNT_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '12
      OldTempAsc = TempAsc
      
      Dim Emp As CEmployee
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
      If Emp.EMP_CODE <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
         Call MsgBox("ยังไม่มีรหัสพนักงาน " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      Bl.ACCEPT_BY = Emp.EMP_ID                                                     '13
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set Emp = GetEmployee(EmpColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
      Bl.RECEIVE_BY = Emp.EMP_ID                                                     '14
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.EXCEPTION_FLAG = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '15
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.PAYEE_NAME = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '16
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.COMMIT_FLAG = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '17
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.RECEIPT_TYPE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '18
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.DOCUMENT_SUBTYPE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '19
      OldTempAsc = TempAsc
      
      Dim BankBranch As CBankBranch
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set BankBranch = GetBankBranch(BankBranch, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
      If BankBranch.BBRANCH_NO <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
         Call MsgBox("ยังไม่มีรหัสสาขา " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      Bl.BANK_BRANCH_ID = BankBranch.BBRANCH_ID                                                     '20
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.BANK_NOTE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '21
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.PAYMENT_TYPE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '22
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.CHECK_NO = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '23
      OldTempAsc = TempAsc
      
      Dim Region As CRegion
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set Region = GetRegion(RegionColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
      If Region.REGION_NO <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
         Call MsgBox("ยังไม่มีรหัสเขตการขาย " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      Bl.REGION_ID = Region.REGION_ID                                                     '24
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.PAID_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '25
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.DEBIT_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '26
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.CREDIT_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '27
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.SIMULATE_FLAG = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '28
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.YYYYMM = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '29
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Bl.YYYYMM2 = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)                                                     '30
      OldTempAsc = TempAsc
      
      
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      
      
      
      
         
      'FK = Bl.BILLING_DOC_ID
      
   End If
   
   
   If Left(LineStr, 2) = "DO" Then
      Dim Ti As CDoItem
      Set Ti = New CDoItem
      Ti.Flag = "A"

      'TI.DO_ID = FK

      TempAsc = 3
      OldTempAsc = TempAsc
      TempAsc = InStr(4, LineStr, ";")
      Ti.ITEM_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc

      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.AVG_WEIGHT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.TOTAL_PRICE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.AVG_PRICE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      Dim PigStatus As CProductStatus
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set PigStatus = GetProductStatus(PigStatusTypeColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
      If PigStatus.PRODUCT_STATUS_NO <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
         Call MsgBox("ยังไม่มีรหัสสถานะวัตถุดิบ " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      Ti.PIG_STATUS = PigStatus.PRODUCT_STATUS_ID                       'code
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.TOTAL_WEIGHT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.LINK_ID = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      
      Dim Revenue As CRevenueType
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Set Revenue = GetRevenueType(RevenueColls, Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)))
      If Revenue.REVENUE_NO <> Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) Then
         Call MsgBox("ยังไม่มีรหัสประเภทค่าใช้จ่าย " & Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)) & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      Ti.REVENUE_ID = Revenue.REVENUE_TYPE_ID                             'code
      OldTempAsc = TempAsc
      
      Dim Part As CPartItem
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Key1 = Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))                          'pig_type
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Key2 = Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))                          'pig_flag
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Key3 = Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))                          'Part_no
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Key4 = Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))                          'part_desc
      OldTempAsc = TempAsc
      
      Set Part = GetPartItem(PartColls, Trim(Key1 & "-" & Key2 & "-" & Key3 & "-" & Key4))
      If Part.PIG_TYPE <> Key1 Then
         Call MsgBox("ยังไม่มีประเภทหมู " & Key1 & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      If Part.PIG_FLAG <> Key2 Then
         Call MsgBox("ยังไม่มีแฟร็กหมู " & Key2 & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      If Part.PART_NO <> Key3 Then
         Call MsgBox("ยังไม่มีรหัสวัตถุดิบ " & Key3 & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      If Part.PART_DESC <> Key4 Then
         Call MsgBox("ยังไม่มีชื่อวัตถุดิบ " & Key4 & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      Ti.PART_ITEM_ID = Part.PART_ITEM_ID                                        'code
      Ti.PART_NO = Part.PART_NO
      
      Dim Lc As CLocation
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Key1 = Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))                          'location_no
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Key2 = Trim(Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))                          'locatin_type
      OldTempAsc = TempAsc
      Set Lc = GetLocation(LocationColls, Trim(Key1 & "-" & Key2))
      If Lc.LOCATION_NO <> Key1 Or Lc.LOCATION_TYPE <> Key2 And Ti.PART_ITEM_ID > 0 Then
         Call MsgBox("ยังไม่มีรหัสคลัง " & Key1 & " และประเภทคลัง " & Key2 & " นี้ในระบบ กรุณาใส่เพิ่มเติ่มในภายหลัง ", vbOKOnly, PROJECT_NAME)
      End If
      Ti.LOCATION_ID = Lc.LOCATION_ID                              'code
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.AGE_CODE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.PIG_AGE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.PKG_TYPE = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.PEDIGREE_COST = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.DISCOUNT_PERCENT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.DISCOUNT_AMOUNT = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.DISCOUNT_REASON = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Ti.SHOW_AVG = Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1)
      OldTempAsc = TempAsc
      If Ti.PART_ITEM_ID > 0 Then
         Call Bl.DoItems.Add(Ti)
      Else
         Call Bl.Revenues.Add(Ti)
      End If
   End If
   
   If Left(LineStr, 2) = "CT" Then
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Ct.Flag = "A"
   
      'Ct.DO_ID = FK
      
      TempAsc = 3
      OldTempAsc = TempAsc
      TempAsc = InStr(4, LineStr, ";")
      Call Ct.SetFieldValue("CHECK_ID", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("BANK_ID", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("BANK_BRANCH", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("TX_TYPE", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("AMOUNT", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("ENTERPRISE_ID", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("PAYMENT_TYPE", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("BANK_ACCOUNT", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("CASH_DOC_ID", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("FEE_AMOUNT", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("NET_AMOUNT", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("TX_NO", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("TX_DATE", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      TempAsc = InStr(TempAsc + 1, LineStr, ";")
      Call Ct.SetFieldValue("CUSTOMER_ID", Mid(LineStr, OldTempAsc + 1, TempAsc - OldTempAsc - 1))
      OldTempAsc = TempAsc
      
      Call Bl.Payments.Add(Ct)
      
   End If
   
   If CountDown = 0 Then
      Call DO2InventoryDoc(Bl, Ivd)
'      If Bl.DOCUMENT_TYPE = 2 Then
'         Call glbDaily.DO2Payment(Bl, Pm)
'      End If

      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   
'      If Bl.DOCUMENT_TYPE = 2 Then
'         Call glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog)
'
'         Bl.PAYMENT_ID = Pm.PAYMENT_ID
'      End If
   
      Bl.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID

      
      Call glbDaily.AddEditBillingDoc(Bl, IsOK, False, glbErrorLog)
      
      Set Bl = New CBillingDoc
      Set Ivd = New CInventoryDoc
   End If
   
   
   ProcessLine = True
   
   Exit Function
ErrorHandler:
   ProcessLine = False
End Function

Private Sub Form_Activate()
   
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadAccountEx(Nothing, AccountColls)
      Call LoadEmployeeCode(Nothing, EmpColls)
      Call LoadBankBranchEx(Nothing, BankBranchColls)
      Call LoadRegionEx(Nothing, RegionColls)
      
      Call LoadPartItem(Nothing, PartColls, , "", , , , 3)
      Call LoadLocationByCode(Nothing, LocationColls, , "", , , 2)
      Call LoadProductStatusCode(Nothing, PigStatusTypeColls)
      Call LoadRevenueTypeCode(Nothing, RevenueColls)
      
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
   pnlHeader.Caption = "อิมพอร์ตข้อมูลเวลาจากฟาร์ม"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblMasterName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
      
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   
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
   
   Set TempBillingDoc = New Collection
   
   Set m_Rs = New ADODB.Recordset
   Set AccountColls = New Collection
   Set EmpColls = New Collection
   Set BankBranchColls = New Collection
   Set RegionColls = New Collection
   
   Set PigStatusTypeColls = New Collection
   Set RevenueColls = New Collection
   Set PartColls = New Collection
   Set LocationColls = New Collection
   Set Bl = New CBillingDoc
   Set Ivd = New CInventoryDoc
   Set Pm = New CPayment
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set TempBillingDoc = Nothing
   Set AccountColls = Nothing
   Set EmpColls = Nothing
   Set BankBranchColls = Nothing
   Set RegionColls = Nothing
   
   Set PigStatusTypeColls = Nothing
   Set RevenueColls = Nothing
   Set PartColls = Nothing
   Set LocationColls = Nothing
   Set Bl = Nothing
   Set Ivd = Nothing
   Set Pm = Nothing
   
End Sub
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
   Ivd.EXCEPTION_FLAG = Bd.EXCEPTION_FLAG
   Ivd.CUS_ID = Bd.CUSTOMER_ID
   Ivd.DOCUMENT_TYPE = 10
   Ivd.DOCUMENT_SUBTYPE = Bd.DOCUMENT_SUBTYPE

   If Bd.DOCUMENT_SUBTYPE = 1 Then 'หมู
      Ivd.SALE_FLAG = "N"
   ElseIf Bd.DOCUMENT_SUBTYPE = 2 Then 'วัตถุดิบ
      Ivd.SALE_FLAG = "Y"
   End If

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
         EI.PIG_AGE = GetAge(Di.PART_NO, Bd.DOCUMENT_DATE)
         EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
         
         Di.PIG_AGE = EI.PIG_AGE
         Di.AGE_CODE = EI.AGE_CODE
                 
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
         EI.TOTAL_PRICE = Di.TOTAL_PRICE
         EI.CALCULATE_FLAG = "N"
         EI.PIG_AGE = GetAge(Di.PART_NO, Bd.DOCUMENT_DATE)
         EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
         
         Di.PIG_AGE = EI.PIG_AGE
         Di.AGE_CODE = EI.AGE_CODE
      ElseIf Di.Flag = "D" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         EI.Flag = "D"
      End If
   Next Di
End Function

