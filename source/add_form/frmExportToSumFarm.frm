VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportToSumFarm 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   Icon            =   "frmExportToSumFarm.frx":0000
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
         MouseIcon       =   "frmExportToSumFarm.frx":27A2
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
Attribute VB_Name = "frmExportToSumFarm"
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
Private TestSum As Double


Private Sub cmdOK_Click()
   glbParameterObj.SumFarmExportFileName = txtFileName.Text
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long
Dim HasBegin As Boolean
Dim FirstDate  As Date
Dim LastDate  As Date
   
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Sub
   End If
   
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
   
   Call GetFirstLastDate(uctlFromDate.ShowDate, FirstDate, LastDate)
   If uctlFromDate.ShowDate <> FirstDate Or uctlToDate.ShowDate <> LastDate Then
      glbErrorLog.LocalErrorMsg = "กรุณาใส่ข้อมูลต้นเดือนปลายเดือน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   
   If Not VerifyTextControl(lblFileName, txtFileName) Then
      Exit Sub
   End If
   
   glbParameterObj.SumFarmExportFileName = txtFileName.Text
   
   
   Call EnableForm(Me, False)
   Call ExportToMain
   Call EnableForm(Me, True)
 
End Sub
Private Sub GenerateDetail(FileID As Long, O As CDoItem)
Dim TempStr As String
   
   TempStr = ""
   TempStr = TempStr & O.PART_NO & ";"
   TempStr = TempStr & O.PIG_TYPE & ";"
   TempStr = TempStr & O.PRODUCT_STATUS_NO & ";"
   TempStr = TempStr & O.PIG_AGE & ";"
   TempStr = TempStr & ChangeQuote(Trim(DateToStringIntToSumFarm(O.DOCUMENT_DATE))) & ";"
   TempStr = TempStr & O.REVENUE_NO & ";"
   
   TempStr = TempStr & O.TOTAL_PRICE - O.DISCOUNT_AMOUNT & ";"
   TempStr = TempStr & O.TOTAL_WEIGHT & ";"
   TempStr = TempStr & O.ITEM_AMOUNT & ";"
   TempStr = TempStr & O.EXPORT_TOTAL_PRICE & ";"
   
   TestSum = TestSum + O.ITEM_AMOUNT
   
   Print #FileID, TempStr
End Sub
Private Sub GenerateDetailEx(FileID As Long, O As CMovementItem)
Dim TempStr As String
   
   TempStr = ""
   TempStr = TempStr & O.PIG_NO & ";"
   TempStr = TempStr & O.PIG_TYPE & ";"
   TempStr = TempStr & O.PIG_STATUS_NO & ";"
   TempStr = TempStr & O.PIG_AGE & ";"
   TempStr = TempStr & O.TX_TYPE & ";"
   TempStr = TempStr & O.DOCUMENT_TYPE & ";"
   TempStr = TempStr & ChangeQuote(Trim(DateToStringIntToSumFarm(O.DOCUMENT_DATE))) & ";"
   TempStr = TempStr & O.TX_AMOUNT & ";"
   
   Print #FileID, TempStr
End Sub
Private Sub GenerateDetailEx2(FileID As Long, O As CMovementItem)
Dim TempStr As String
   
   TempStr = ""
   TempStr = TempStr & O.PIG_NO & ";"
   TempStr = TempStr & O.PIG_TYPE & ";"
   TempStr = TempStr & O.PIG_STATUS_NO & ";"
   TempStr = TempStr & O.PIG_AGE & ";"
   TempStr = TempStr & O.PART_TYPE_NO & ";"
   TempStr = TempStr & O.TX_TYPE & ";"
   TempStr = TempStr & O.EXPENSE_TYPE_NO & ";"
   TempStr = TempStr & O.DOCUMENT_TYPE & ";"
   TempStr = TempStr & ChangeQuote(Trim(DateToStringIntToSumFarm(O.DOCUMENT_DATE))) & ";"
   TempStr = TempStr & O.CAPITAL_AMOUNT & ";"
   
   Print #FileID, TempStr
End Sub
Private Sub GenerateDetailEx3(FileID As Long, O As CBalanceAccum)
Dim TempStr As String
   
   TempStr = ""
   TempStr = TempStr & O.PART_NO & ";"
   TempStr = TempStr & O.PIG_TYPE & ";"
   TempStr = TempStr & Left(O.YYYYMM, 4) & Right(O.YYYYMM, 2) & ";"
   TempStr = TempStr & O.BALANCE_AMOUNT & ";"
   
   Print #FileID, TempStr
End Sub
Private Sub GenerateDetailEx4(FileID As Long, O As CMovementItem)
Dim TempStr As String
   
   TempStr = ""
   TempStr = TempStr & O.PIG_NO & ";"
   TempStr = TempStr & O.PIG_TYPE & ";"
   TempStr = TempStr & Left(O.YYYYMM, 4) & Right(O.YYYYMM, 2) & ";"
   TempStr = TempStr & O.PART_GROUP_NO & ";"
   TempStr = TempStr & O.EXPENSE_TYPE_NO & ";"
   TempStr = TempStr & O.CAPITAL_AMOUNT & ";"
   TempStr = TempStr & ChangeQuote(Trim(DateToStringIntToSumFarm(O.DOCUMENT_DATE))) & ";"
   TempStr = TempStr & GetAge(O.PIG_NO, O.DOCUMENT_DATE) & ";"
   
   Print #FileID, TempStr
End Sub

Private Sub ExportToMain()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim FromDate As Date
Dim ToDate As Date
Dim iCount As Long
Dim FileID As Long
Dim OldID As Long
Dim I As Long
Dim ItemCount As Long
Dim m_Rs As ADODB.Recordset
Dim Doc As CDoItem
Dim Mi As CMovementItem
Dim Ba As CBalanceAccum

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
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
   Print #FileID, glbEnterPrise.SHORT_NAME
   Print #FileID, "TRANSACTION_SALEBUY"
   
   Set m_Rs = New ADODB.Recordset
   Set Doc = New CDoItem
   
   Doc.DO_ITEM_ID = -1
   Doc.FROM_DATE = uctlFromDate.ShowDate
   Doc.TO_DATE = uctlToDate.ShowDate
   Call Doc.QueryData(27, m_Rs, ItemCount)
   
'      Generate detail here
   While Not m_Rs.EOF
      
      I = I + 1
      
      prgProgress.Value = MyDiff(I, ItemCount) * 20
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      Call Doc.PopulateFromRS(27, m_Rs)
      Call GenerateDetail(FileID, Doc)
      
      m_Rs.MoveNext
   Wend
      
   '''debug.print (TestSum)
   
   Print #FileID, "COST_ITEM"
   
   prgProgress.Value = 20
   txtPercent.Text = prgProgress.Value
   Me.Refresh
      
   Set m_Rs = Nothing
   
   Set m_Rs = New ADODB.Recordset
   Set Mi = New CMovementItem
   
   Mi.MOVEMENT_ITEM_ID = -1
   Mi.FROM_DATE = uctlFromDate.ShowDate
   Mi.TO_DATE = uctlToDate.ShowDate
   Call Mi.QueryData(22, m_Rs, ItemCount)
   
   I = 0
'      Generate detail here
   While Not m_Rs.EOF
      
      I = I + 1
      
      prgProgress.Value = (MyDiff(I, ItemCount) * 20) + 20
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      Call Mi.PopulateFromRS(22, m_Rs)
      Call GenerateDetailEx2(FileID, Mi)
      
      m_Rs.MoveNext
   Wend
   '--------------------------------------------------------------------------------------------------------------------------------->
   
   Print #FileID, "COST"
   
   prgProgress.Value = 40
   txtPercent.Text = 40
   Me.Refresh
      
   Set m_Rs = Nothing
   
   Set m_Rs = New ADODB.Recordset
   Set Mi = New CMovementItem
   
   Mi.MOVEMENT_ITEM_ID = -1
   Mi.FROM_DATE = uctlFromDate.ShowDate
   Mi.TO_DATE = uctlToDate.ShowDate
   Call Mi.QueryData(23, m_Rs, ItemCount)
   
   I = 0
'      Generate detail here
   While Not m_Rs.EOF
      
      I = I + 1
      
      prgProgress.Value = (MyDiff(I, ItemCount) * 20) + 40
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      Call Mi.PopulateFromRS(23, m_Rs)
      Call GenerateDetailEx(FileID, Mi)
      
      m_Rs.MoveNext
   Wend
   
   '--------------------------------------------------------------------------------------------------------------------------------->
   
   Print #FileID, "BALANCE_ACCUM"
   
   prgProgress.Value = 60
   txtPercent.Text = 60
   Me.Refresh
      
   Set m_Rs = Nothing
   
   Set m_Rs = New ADODB.Recordset
   Set Ba = New CBalanceAccum
   
   Ba.BALANCE_ACCUM_ID = -1
   'Ba.FROM_DATE = uctlFromDate.ShowDate
   Call GetFirstLastDate(uctlToDate.ShowDate, FromDate, ToDate)
   Ba.TO_DATE = ToDate
   Ba.PIG_FLAG = "Y"
   Call Ba.QueryData(18, m_Rs, ItemCount)
   
   I = 0
'      Generate detail here
   While Not m_Rs.EOF
      
      I = I + 1
      
      prgProgress.Value = (MyDiff(I, ItemCount) * 20) + 60
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      Call Ba.PopulateFromRS(18, m_Rs)
      Call GenerateDetailEx3(FileID, Ba)
      
      m_Rs.MoveNext
   Wend
   
   
   '--------------------------------------------------------------------------------------------------------------------------------->
   
   Print #FileID, "COST_BALANCE"
   
   prgProgress.Value = 80
   txtPercent.Text = 80
   Me.Refresh
      
   Set m_Rs = Nothing
   
   Set m_Rs = New ADODB.Recordset
   Set Mi = New CMovementItem
   
   Mi.MOVEMENT_ITEM_ID = -1
   Mi.FROM_DATE = uctlFromDate.ShowDate
   Mi.TO_DATE = ToDate
   Call Mi.QueryData(24, m_Rs, ItemCount)
   
   I = 0
'      Generate detail here
   While Not m_Rs.EOF
      
      I = I + 1
      
      prgProgress.Value = (MyDiff(I, ItemCount) * 20) + 80
      txtPercent.Text = prgProgress.Value
      Me.Refresh
      
      Call Mi.PopulateFromRS(24, m_Rs)
      Call GenerateDetailEx4(FileID, Mi)
      
      m_Rs.MoveNext
   Wend
   
   '--------------------------------------------------------------------------------------------------------------------------------->
   
   Close #FileID
   
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
      
       txtFileName.Text = glbParameterObj.SumFarmExportFileName
       
      If ShowMode = SHOW_EDIT Then
         'Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
      Call GetFirstLastDate(Now, FromDate, ToDate)
      uctlFromDate.ShowDate = FromDate
      uctlToDate.ShowDate = ToDate
      
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
   pnlHeader.Caption = "Export ไปยัง ระบบรวมฟาร์ม"
   
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
