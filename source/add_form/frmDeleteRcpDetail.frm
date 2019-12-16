VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDeleteRcpDetail 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmDeleteRcpDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6915
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   12197
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.PictureBox Picture1 
         Height          =   5295
         Left            =   120
         Picture         =   "frmDeleteRcpDetail.frx":27A2
         ScaleHeight     =   5235
         ScaleWidth      =   2595
         TabIndex        =   13
         Top             =   1560
         Width           =   2655
      End
      Begin prjFarmManagement.uctlTextBox txtFileName1 
         Height          =   465
         Left            =   1890
         TabIndex        =   10
         Top             =   1020
         Width           =   6765
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   4500
         TabIndex        =   0
         Top             =   4980
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   4500
         TabIndex        =   1
         Top             =   5310
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9360
         Top             =   1020
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblNote 
         Caption         =   "Label2"
         Height          =   2175
         Left            =   3360
         TabIndex        =   14
         Top             =   2280
         Width           =   5895
      End
      Begin VB.Label lblFileName1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
      End
      Begin Threed.SSCommand cmdFileName1 
         Height          =   405
         Left            =   8670
         TabIndex        =   11
         Top             =   1020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDeleteRcpDetail.frx":7670
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   4530
         TabIndex        =   2
         Top             =   5940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDeleteRcpDetail.frx":798A
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   6240
         TabIndex        =   9
         Top             =   5430
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2850
         TabIndex        =   8
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2850
         TabIndex        =   7
         Top             =   5460
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7815
         TabIndex        =   4
         Top             =   5940
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6165
         TabIndex        =   3
         Top             =   5940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmDeleteRcpDetail.frx":7CA4
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmDeleteRcpDetail"
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

Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private Sub cmdFileName1_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName1.Text = dlgAdd.FileName
   m_HasModify = True
   
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub ImportBalance()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim HasBegin As Boolean
Dim Rcp As CReceiptItem
Dim TempRcp As CReceiptItem
Dim TempBD As CBillingDoc
Dim IsOK As Boolean
Dim ItemCount As Long
   
   HasBegin = False
   Set m_Rs = New ADODB.Recordset
   ID = 1
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
      
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count

   ReDim FieldNames(MaxCol)
   ReDim FieldTypes(MaxCol)
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdOK.Enabled = False
    
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 1
   prgProgress.MAX = (MaxRow) + 1
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   For Row = 2 To MaxRow
      DoEvents

      Set Rcp = New CReceiptItem
      
      Rcp.DOCUMENT_NO = Trim(m_ExcelSheet.Cells(Row, 1).Value)
      
      If Len(Rcp.DOCUMENT_NO) > 0 Then
         
         Set TempRcp = New CReceiptItem
         TempRcp.DOCUMENT_NO = Rcp.DOCUMENT_NO
         Call TempRcp.QueryData(16, m_Rs, ItemCount)
               
         'Call Rcp.DeleteDataFromReceiptNo
         While Not m_Rs.EOF
            Call TempRcp.PopulateFromRS(16, m_Rs)
            
            Call TempRcp.DeleteData
            
            m_Rs.MoveNext
         Wend
      End If
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
      
      Set Rcp = Nothing
   Next Row
   
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   HasBegin = False
   
   Set m_ExcelSheet = Nothing
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   Set m_ExcelSheet = Nothing
   Call EnableForm(Me, True)

   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True

   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub
'Private Sub ImportBalance()
''On Error GoTo ErrorHandler
'Dim MaxRow As Long
'Dim MaxCol As Long
'Dim ID As Long
'Dim FieldNames() As String
'Dim FieldTypes() As String
'Dim I As Long
'Dim TabField As String
'Dim StateMent As String
'Dim NewValue As String
'Dim Row As Long
'Dim Col As Long
'Dim ErrorCount As Long
'Dim SuccessCount As Long
'Dim ProgressCount As Long
'Dim ErrorFlag As Boolean
'Dim ServerDtm As String
'Dim HasBegin As Boolean
'Dim Bd As CBillingDoc
'Dim IsOK As Boolean
'Dim Accounts As Collection
'Dim Partitems As Collection
'Dim Di As CDoItem
'
'   HasBegin = False
'
'   ID = 1
'
'   Set Accounts = New Collection
'   Call LoadAccountEx(Nothing, Accounts)
'
'   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
'
'   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
'   MaxCol = m_ExcelSheet.UsedRange.Columns.Count
'
'   ReDim FieldNames(MaxCol)
'   ReDim FieldTypes(MaxCol)
'
'   Call EnableForm(Me, False)
'   cmdStart.Enabled = False
'   cmdExit.Enabled = False
'   cmdOK.Enabled = False
'
'   ProgressCount = 0
'   ErrorCount = 0
'   SuccessCount = 0
'
'   prgProgress.MIN = 1
'   prgProgress.MAX = (MaxRow) + 1
'
'   glbDatabaseMngr.DBConnection.BeginTrans
'   HasBegin = True
'
'   For Row = 2 To MaxRow
'      DoEvents
'
'      Set Bd = New CBillingDoc
'
'      Set Ac = GetAccount(Accounts, Trim(m_ExcelSheet.Cells(Row, 1).Value))
'      If Ac Is Nothing Then
'         glbErrorLog.LocalErrorMsg = Trim(m_ExcelSheet.Cells(Row, 1).Value) & " --> " & Trim(m_ExcelSheet.Cells(Row, 2).Value)
'         glbErrorLog.ShowUserError
'         Set Ac = GetAccount(Accounts, "C-0000")
'      End If
'
'      Bd.AddEditMode = SHOW_ADD
'      Bd.DOCUMENT_NO = Trim(m_ExcelSheet.Cells(Row, 2).Value) & "."
'      Bd.DOCUMENT_DATE = m_ExcelSheet.Cells(Row, 3).Value
'      Bd.ACCOUNT_ID = Ac.ACCOUNT_ID
'      Bd.DOCUMENT_TYPE = 1
'      Bd.DOCUMENT_SUBTYPE = 1
'      Bd.RECEIPT_TYPE = 0
'      Bd.COMMIT_FLAG = "N"
'      Bd.EXCEPTION_FLAG = "N"
'
'      Set Di = New CDoItem
'      Di.ITEM_AMOUNT = 1
'      Di.PART_ITEM_ID = glbDaily.LookupPigID("254400", "N")
'      Di.LOCATION_ID = glbDaily.LookupLocationID("00", "Y", 1)
'      Di.TOTAL_PRICE = Val(Trim(m_ExcelSheet.Cells(Row, 6).Value))
'
'      Di.Flag = "A"
'      Call Bd.DoItems.Add(Di)
'      Set Di = Nothing
'
'      Call glbDaily.AddEditBillingDoc(Bd, IsOK, False, glbErrorLog)
'
'      ProgressCount = ProgressCount + 1
'      prgProgress.Value = ProgressCount
'
'      Set Bd = Nothing
'   Next Row
'
'   prgProgress.Value = prgProgress.MAX
'
'   Call EnableForm(Me, True)
'   glbDatabaseMngr.DBConnection.CommitTrans
'   HasBegin = False
'
'   Set m_ExcelSheet = Nothing
'   Set Accounts = Nothing
'
'   cmdStart.Enabled = True
'   cmdExit.Enabled = True
'   cmdOK.Enabled = True
'   Exit Sub
'
''ErrorHandler:
''   If HasBegin Then
''      glbDatabaseMngr.DBConnection.RollbackTrans
''   End If
''
''   Call EnableForm(Me, True)
''
''   cmdStart.Enabled = True
''   cmdExit.Enabled = True
''   cmdOK.Enabled = True
''
''   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
''   glbErrorLog.ShowUserError
'End Sub

Private Sub cmdStart_Click()
   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Workbooks.Open (txtFileName1.Text)
      
   Call ImportBalance
      
   m_ExcelApp.Workbooks.Close
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If ShowMode = SHOW_EDIT Then
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
   pnlHeader.Caption = "ลบข้อมูลการตัดใบส่งสินค้าในใบเสร็จรับเงิน"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFileName1, "ไฟล์ Excel")
   Call InitNormalLabel(lblNote, "ขั้นตอนในการลบข้อมูลการตัดใบส่งสินค้าในใบเสร็จรับเงิน" & vbCrLf & "1. สร้าง Excel ตามรูป " & vbCrLf & "2. ปิดไฟล์ Excel ทั้งหมด " & vbCrLf & "3. Import เข้าที่โปรแกรม" & vbCrLf & "4.กดเริ่ม")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName1.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName1.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName1.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName1, MapText("..."))
   
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
   
   m_HasActivate = False
   Call InitFormLayout
   Set m_ExcelApp = CreateObject("Excel.application")
   
   Call EnableForm(Me, True)
End Sub
