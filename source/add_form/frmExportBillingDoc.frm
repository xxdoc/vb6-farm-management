VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportBillingDoc 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   Icon            =   "frmExportBillingDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   11220
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   6429
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   3
         Top             =   1860
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   4
         Top             =   2190
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9840
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1410
         Width           =   7305
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   7200
         TabIndex        =   1
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   6120
         TabIndex        =   15
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   1110
         Width           =   945
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   5
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmExportBillingDoc.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   12
         Top             =   2310
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   2340
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   7
         Top             =   2820
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   6
         Top             =   2820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmExportBillingDoc"
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

Private EmpColls As Collection
Private BankBranchColls As Collection
Private RegionColls As Collection

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
   
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Sub
   End If
   
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   glbParameterObj.BLExportFileName = txtFileName.Text
   
   Call EnableForm(Me, False)
   Call ExportBillingDoc
   Call EnableForm(Me, True)
 
End Sub
Private Sub GenerateHeader(FileID As Long, O As CBillingDoc)
Dim TempStr As String
Dim Emp As CEmployee
Dim BankBranch As CBankBranch
Dim Region As CRegion
   
   TempStr = "BL;"
   TempStr = TempStr & O.ACCOUNT_NO & ";"                                                                                                                            '1
   TempStr = TempStr & O.CUSTOMER_CODE & ";"                                                                                                                            '1
   TempStr = TempStr & O.DOCUMENT_NO & ";"                                                                                                                            '2
   TempStr = TempStr & DateToStringInt(O.DOCUMENT_DATE) & ";"                                                                                                                            '3
   TempStr = TempStr & O.DOCUMENT_TYPE & ";"                                                                                                                            '4
   TempStr = TempStr & DateToStringInt(O.DUE_DATE) & ";"                                                                                                                            '5
   TempStr = TempStr & O.NOTE & ";"                                                                                                                            '6
   TempStr = TempStr & O.TOTAL_AMOUNT & ";"                                                                                                                            '7
   TempStr = TempStr & O.VAT_PERCENT & ";"                                                                                                                            '8
   TempStr = TempStr & O.VAT_AMOUNT & ";"                                                                                                                            '9
   TempStr = TempStr & O.WH_PERCENT & ";"                                                                                                                            '10
   TempStr = TempStr & O.WH_AMOUNT & ";"                                                                                                                            '11
   TempStr = TempStr & O.DISCOUNT_AMOUNT & ";"                                                                                                                            '12
      
   Set Emp = GetEmployee(EmpColls, Trim(Str(O.ACCEPT_BY)))
   TempStr = TempStr & Emp.EMP_CODE & ";"                    'code                                                                                                                            '13
   Set Emp = GetEmployee(EmpColls, Trim(Str(O.RECEIVE_BY)))
   TempStr = TempStr & Emp.EMP_CODE & ";"                    'code                                                                                                                            '14
   
   TempStr = TempStr & O.EXCEPTION_FLAG & ";"                                                                                                                            '15
   TempStr = TempStr & O.PAYEE_NAME & ";"                                                                                                                            '16
   TempStr = TempStr & O.COMMIT_FLAG & ";"                                                                                                                            '17
   TempStr = TempStr & O.RECEIPT_TYPE & ";"                                                                                                                            '18
   TempStr = TempStr & O.DOCUMENT_SUBTYPE & ";"                                                                                                                            '19
   
   Set BankBranch = GetBankBranch(BankBranchColls, Trim(Str(O.BANK_BRANCH_ID)))
   TempStr = TempStr & BankBranch.BBRANCH_NO & ";"         'code                                                                                                                            '20
   
   TempStr = TempStr & O.BANK_NOTE & ";"                                                                                                                            '21
   TempStr = TempStr & O.PAYMENT_TYPE & ";"                                                                                                                            '22
   TempStr = TempStr & O.CHECK_NO & ";"                                                                                                                            '23
   Set Region = GetRegion(RegionColls, Trim(Str(O.REGION_ID)))
   TempStr = TempStr & Region.REGION_NO & ";"         'code                                                                                                                            '24
   TempStr = TempStr & O.PAID_AMOUNT & ";"                                                                                                                            '25
   TempStr = TempStr & O.DEBIT_AMOUNT & ";"                                                                                                                            '26
   TempStr = TempStr & O.CREDIT_AMOUNT & ";"                                                                                                                            '27
   TempStr = TempStr & O.SIMULATE_FLAG & ";"                                                                                                                            '28
   TempStr = TempStr & O.YYYYMM & ";"                                                                                                                            '29
   TempStr = TempStr & O.YYYYMM2 & ";"                                                                                                                            '30
   
   
   Print #FileID, TempStr
End Sub

Private Sub GenerateDetail(FileID As Long, O As CDoItem)
Dim TempStr As String

   TempStr = "DO;"
   TempStr = TempStr & O.ITEM_AMOUNT & ";"
   TempStr = TempStr & O.AVG_WEIGHT & ";"
   TempStr = TempStr & O.TOTAL_PRICE & ";"
   TempStr = TempStr & O.AVG_PRICE & ";"
   TempStr = TempStr & O.PIG_STATUS_NO & ";"                                                 'code
   TempStr = TempStr & O.TOTAL_WEIGHT & ";"
   TempStr = TempStr & O.LINK_ID & ";"
   TempStr = TempStr & O.REVENUE_NO & ";"                                                       'code
   
   TempStr = TempStr & O.PIG_TYPE & ";"                                                                 'code
   TempStr = TempStr & O.PIG_FLAG & ";"                                                                 'code
   TempStr = TempStr & O.PART_NO & ";"                                                                 'code
   TempStr = TempStr & O.PART_DESC & ";"                                                                 'code
   
   TempStr = TempStr & O.LOCATION_NO & ";"                                                 'code
   TempStr = TempStr & O.LOCATION_TYPE & ";"                                                 'code
   
   TempStr = TempStr & O.AGE_CODE & ";"
   TempStr = TempStr & O.PIG_AGE & ";"
   
   TempStr = TempStr & O.PKG_TYPE & ";"
   TempStr = TempStr & O.PEDIGREE_COST & ";"
   TempStr = TempStr & O.DISCOUNT_PERCENT & ";"
   TempStr = TempStr & O.DISCOUNT_AMOUNT & ";"
   TempStr = TempStr & O.DISCOUNT_REASON & ";"
   TempStr = TempStr & O.SHOW_AVG & ";"
   
   
   Print #FileID, TempStr
End Sub
Private Sub GenerateDetailCt(FileID As Long, O As CCashTran)
Dim TempStr As String
   
   TempStr = "CT;"
   TempStr = TempStr & O.GetFieldValue("CHECK_ID") & ";"
   TempStr = TempStr & O.GetFieldValue("BANK_ID") & ";"
   TempStr = TempStr & O.GetFieldValue("BANK_BRANCH") & ";"
   TempStr = TempStr & O.GetFieldValue("TX_TYPE") & ";"
   TempStr = TempStr & O.GetFieldValue("AMOUNT") & ";"                                                 'code
   TempStr = TempStr & O.GetFieldValue("ENTERPRISE_ID") & ";"
   TempStr = TempStr & O.GetFieldValue("PAYMENT_TYPE") & ";"
   TempStr = TempStr & O.GetFieldValue("BANK_ACCOUNT") & ";"                                                       'code
   TempStr = TempStr & O.GetFieldValue("CASH_DOC_ID") & ";"                                                       'code
      
   TempStr = TempStr & O.GetFieldValue("FEE_AMOUNT") & ";"                                                       'code
   TempStr = TempStr & O.GetFieldValue("NET_AMOUNT") & ";"                                                       'code
   TempStr = TempStr & O.GetFieldValue("TX_NO") & ";"                                                       'code
   TempStr = TempStr & O.GetFieldValue("TX_DATE") & ";"                                                       'code
   TempStr = TempStr & O.GetFieldValue("CUSTOMER_ID") & ";"                                                       'code
   
   Print #FileID, TempStr
End Sub

Private Sub ExportBillingDoc()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim FromDate As Date
Dim ToDate As Date
Dim CB As CBillingDoc
Dim iCount As Long
Dim FileID As Long
Dim OldID As Long
Dim I As Long
Dim ItemCount As Long
Dim m_Rs2 As ADODB.Recordset
Dim m_Rs3 As ADODB.Recordset
Dim Doc As CDoItem
Dim Ct As CCashTran
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Set CB = New CBillingDoc
   CB.BILLING_DOC_ID = -1
   CB.FROM_DATE = uctlFromDate.ShowDate
   CB.TO_DATE = uctlToDate.ShowDate
   CB.OrderBy = 1
  Call CB.QueryData(1, m_Rs, iCount)

   On Error GoTo XXX
      Call Kill(txtFileName.Text)
XXX:

   FileID = FreeFile
   Open txtFileName.Text For Append As #FileID

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0

   I = 0
   
   Print #FileID, uctlFromDate.ShowDate
   
   Print #FileID, uctlToDate.ShowDate
   
   While Not m_Rs.EOF
      prgProgress.Value = MyDiff(I, iCount) * 100
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      Call CB.PopulateFromRS(1, m_Rs)
         
      If CB.DOCUMENT_TYPE = 1 Or (CB.DOCUMENT_TYPE = 2 And CB.RECEIPT_TYPE = 1) Then
         'm_Rs.MoveNext
         
         Call GenerateHeader(FileID, CB)
         
         Set m_Rs2 = New ADODB.Recordset
         Set m_Rs3 = New ADODB.Recordset
         Set Doc = New CDoItem
         
         Doc.DO_ITEM_ID = -1
         Doc.DO_ID = CB.BILLING_DOC_ID
         Call Doc.QueryData(1, m_Rs2, ItemCount)
         
   '      Generate detail here
         While Not m_Rs2.EOF
            Call Doc.PopulateFromRS(1, m_Rs2)
            Call GenerateDetail(FileID, Doc)
            
            m_Rs2.MoveNext
         Wend
         
         Set Ct = New CCashTran
         
         Call Ct.SetFieldValue("CASH_TRAN_ID", -1)
         Call Ct.SetFieldValue("BILLING_DOC_ID", CB.BILLING_DOC_ID)
         Call Ct.QueryData(1, m_Rs3, ItemCount)
         
   '      Generate detail here
         While Not m_Rs3.EOF
            Call Ct.PopulateFromRS(1, m_Rs3)
            Call GenerateDetailCt(FileID, Ct)
            
            m_Rs3.MoveNext
         Wend
         
      End If
      m_Rs.MoveNext
      I = I + 1
   Wend
   Close #FileID
   
   Set CB = Nothing
   prgProgress.Value = 100
   txtPercent.Text = 100
      
   Exit Sub
   
ErrorHandler:
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadEmployee(Nothing, EmpColls)
      Call LoadBankBranch(Nothing, BankBranchColls)
      Call LoadRegion(Nothing, RegionColls)
      
       txtFileName.Text = glbParameterObj.BLExportFileName
       
      If ShowMode = SHOW_EDIT Then
         'Call QueryData(True)
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
   pnlHeader.Caption = "Export บิลขาย"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFromDate, "จากวันที่")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFileName, "ชื่อไฟล์")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
      
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
   
   Set m_Rs = New ADODB.Recordset
   Set EmpColls = New Collection
   Set BankBranchColls = New Collection
   Set RegionColls = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set EmpColls = Nothing
   Set BankBranchColls = Nothing
   Set RegionColls = Nothing
End Sub
