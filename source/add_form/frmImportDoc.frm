VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmImportDoc 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5715
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   10081
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFileName1 
         Height          =   465
         Left            =   4290
         TabIndex        =   12
         Top             =   1020
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   3900
         Width           =   7305
         _ExtentX        =   12885
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
         Left            =   1860
         TabIndex        =   1
         Top             =   4230
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
      Begin prjFarmManagement.uctlTextBox txtFileName2 
         Height          =   465
         Left            =   4290
         TabIndex        =   14
         Top             =   1500
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtFileName3 
         Height          =   465
         Left            =   4290
         TabIndex        =   16
         Top             =   1980
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtFileName4 
         Height          =   465
         Left            =   4290
         TabIndex        =   19
         Top             =   2460
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtFileName5 
         Height          =   465
         Left            =   4290
         TabIndex        =   22
         Top             =   2940
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtFileName6 
         Height          =   465
         Left            =   4290
         TabIndex        =   26
         Top             =   3420
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   820
      End
      Begin Threed.SSOption radARCredit 
         Height          =   405
         Left            =   1830
         TabIndex        =   25
         Top             =   3480
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSCommand cmdFileName6 
         Height          =   405
         Left            =   8670
         TabIndex        =   27
         Top             =   3420
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName5 
         Height          =   405
         Left            =   8670
         TabIndex        =   24
         Top             =   2940
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSOption radBalance 
         Height          =   405
         Left            =   1830
         TabIndex        =   23
         Top             =   3000
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSCommand cmdFileName4 
         Height          =   405
         Left            =   8670
         TabIndex        =   21
         Top             =   2460
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":2DD6
         ButtonStyle     =   3
      End
      Begin Threed.SSOption radCustomer 
         Height          =   405
         Left            =   1830
         TabIndex        =   20
         Top             =   2520
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption radCapitalImport 
         Height          =   405
         Left            =   1830
         TabIndex        =   18
         Top             =   2040
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSCommand cmdFileName3 
         Height          =   405
         Left            =   8670
         TabIndex        =   17
         Top             =   1980
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":30F0
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName2 
         Height          =   405
         Left            =   8670
         TabIndex        =   15
         Top             =   1500
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":340A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName1 
         Height          =   405
         Left            =   8670
         TabIndex        =   13
         Top             =   1020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":3724
         ButtonStyle     =   3
      End
      Begin Threed.SSOption radPigImport 
         Height          =   405
         Left            =   1830
         TabIndex        =   11
         Top             =   1560
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption radStockImport 
         Height          =   405
         Left            =   1830
         TabIndex        =   10
         Top             =   1080
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   4860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":3A3E
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   9
         Top             =   4350
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   4380
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   4
         Top             =   4860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   3
         Top             =   4860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":3D58
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportDoc"
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

Private m_ExcelApp As Object
Private m_ExcelSheet As Object

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

Private Sub cmdFileName2_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName2.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdFileName3_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName3.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdFileName4_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName4.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdFileName5_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName5.Text = dlgAdd.FileName
   m_HasModify = True
End Sub

Private Sub cmdFileName6_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName6.Text = dlgAdd.FileName
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

Private Sub ImportStock()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim Ivd As CInventoryDoc
Dim II As CImportItem
Dim IsOK As Boolean

   HasBegin = False

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
   
   TabField = " ("
   For I = 1 To MaxCol
      FieldTypes(I - 1) = Trim(m_ExcelSheet.Cells(2, I))
      If I > 1 Then
         TabField = TabField & "," & Trim(m_ExcelSheet.Cells(1, I))
      Else
         TabField = TabField & Trim(m_ExcelSheet.Cells(1, I))
      End If
   Next I
   TabField = TabField & ")"
    
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.MIN = 1
   prgProgress.MAX = (MaxRow) + 1
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = "ยกมาวัตถุดิบ"
   Ivd.DOCUMENT_DATE = InternalDateToDateEx2(m_ExcelApp.Sheets(ID).Name)
   Ivd.COMMIT_FLAG = "N"
   Ivd.DOCUMENT_TYPE = 1
   Ivd.EXCEPTION_FLAG = "N"
   
   For Row = 2 To MaxRow
      DoEvents
      
      Set II = New CImportItem
      II.Flag = "A"
      II.TOTAL_INCLUDE_PRICE = Val(m_ExcelSheet.Cells(Row, 5).Value)
      II.TOTAL_ACTUAL_PRICE = II.TOTAL_INCLUDE_PRICE
      II.CALCULATE_FLAG = "Y"
      II.IMPORT_AMOUNT = Val(m_ExcelSheet.Cells(Row, 3).Value)
      II.INCLUDE_UNIT_PRICE = Minus2Zero(MyDiff(II.TOTAL_INCLUDE_PRICE, II.IMPORT_AMOUNT))
      II.ACTUAL_UNIT_PRICE = II.INCLUDE_UNIT_PRICE
      II.LOCATION_ID = glbDaily.LookupLocationID("00", "N", 2)
      II.PART_ITEM_ID = glbDaily.LookupPartItemIDFromName(Trim(m_ExcelSheet.Cells(Row, 2).Value), Trim(m_ExcelSheet.Cells(Row, 1).Value))
      If II.PART_ITEM_ID > 0 Then
         Call Ivd.ImportExports.Add(II)
      Else
'''debug.print "ไม่พบข้อมูลวัตถุดิบ " & Trim(m_ExcelSheet.Cells(Row, 2).Value) & " ประเภท " & Trim(m_ExcelSheet.Cells(Row, 1).Value) & " ในฐานข้อมูล"
         glbErrorLog.LocalErrorMsg = "ไม่พบข้อมูลวัตถุดิบ '" & Trim(m_ExcelSheet.Cells(Row, 2).Value) & "' ประเภท '" & Trim(m_ExcelSheet.Cells(Row, 1).Value) & "' ในฐานข้อมูล"
         glbErrorLog.ShowUserError
      End If
      Set II = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next Row
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   
   Set Ivd = Nothing
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
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub ImportPig()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim Ivd As CInventoryDoc
Dim II As CImportItem
Dim IsOK As Boolean
Dim OkFlag As Boolean
Dim SaleFlag As String

   HasBegin = False

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
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = "ยกมาสุกร"
   Ivd.DOCUMENT_DATE = InternalDateToDateEx2(m_ExcelApp.Sheets(ID).Name)
   Ivd.COMMIT_FLAG = "N"
   Ivd.DOCUMENT_TYPE = 11
   Ivd.EXCEPTION_FLAG = "N"
   
   For Row = 2 To MaxRow
      DoEvents
      OkFlag = True
      
      Set II = New CImportItem
      II.Flag = "A"
      II.TOTAL_INCLUDE_PRICE = 0
      II.TOTAL_ACTUAL_PRICE = 0
      II.CALCULATE_FLAG = "N"
      II.IMPORT_AMOUNT = Val(m_ExcelSheet.Cells(Row, 4).Value)
      II.INCLUDE_UNIT_PRICE = 0
      II.ACTUAL_UNIT_PRICE = 0
      II.LOCATION_ID = glbDaily.LookupLocationIDNameEx(Trim(m_ExcelSheet.Cells(Row, 1).Value), "", 1, SaleFlag)
      If II.LOCATION_ID <= 0 Then
         glbErrorLog.LocalErrorMsg = "ไม่พบโรงเรือน '" & Trim(m_ExcelSheet.Cells(Row, 1).Value) & "'"
         glbErrorLog.ShowUserError
         OkFlag = False
      End If
      
      II.PART_ITEM_ID = glbDaily.LookupPigIDEx(Trim(m_ExcelSheet.Cells(Row, 3).Value), Trim(m_ExcelSheet.Cells(Row, 2).Value))
      If II.PART_ITEM_ID <= 0 Then
         glbErrorLog.LocalErrorMsg = "ไม่พบสุกร '" & Trim(m_ExcelSheet.Cells(Row, 3).Value) & "'" & "ประเภท '" & Trim(m_ExcelSheet.Cells(Row, 2).Value) & "'"
         glbErrorLog.ShowUserError
         OkFlag = False
      End If

      If SaleFlag = "Y" Then
         II.PIG_STATUS = glbDaily.LookupPigStatusEx(Trim(m_ExcelSheet.Cells(Row, 5).Value))
         If II.PIG_STATUS <= 0 Then
            glbErrorLog.LocalErrorMsg = "ไม่พบสถานะสุกร '" & Trim(m_ExcelSheet.Cells(Row, 5).Value) & "'"
            glbErrorLog.ShowUserError
            OkFlag = False
         End If
      End If
      
      If OkFlag Then
         Call Ivd.ImportExports.Add(II)
      End If
      Set II = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
   Next Row
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   Set Ivd = Nothing
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
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub ImportCapital()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim Cm As CCapitalMovement
Dim Mi As CMovementItem
Dim IsOK As Boolean

   HasBegin = False

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

   Set Cm = New CCapitalMovement
   Call Cm.DeleteAllData
   Set Cm = Nothing
   
   For Row = 3 To MaxRow
      DoEvents

      Set Cm = New CCapitalMovement
      Cm.AddEditMode = SHOW_ADD
      Cm.DOCUMENT_NO = "ต้นทุนยกมา"
      Cm.DOCUMENT_DATE = InternalDateToDateEx2(m_ExcelApp.Sheets(ID).Name)
      Cm.DOCUMENT_TYPE = 0
      Cm.DOCUMENT_CATEGORY = 3 'ยอดยกมา
      Cm.FROM_HOUSE_ID = glbDaily.LookupLocationIDNameEx(Trim(m_ExcelSheet.Cells(Row, 1).Value), "", 1)
      If Cm.FROM_HOUSE_ID <= 0 Then
         glbErrorLog.LocalErrorMsg = "ไม่พบโรงเรือน '" & Trim(m_ExcelSheet.Cells(Row, 1).Value) & "'"
         glbErrorLog.ShowUserError
         
         glbDatabaseMngr.DBConnection.RollbackTrans
         cmdStart.Enabled = True
         cmdExit.Enabled = True
         cmdOK.Enabled = True
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Cm.PIG_ID = glbDaily.LookupPigIDEx(Trim(m_ExcelSheet.Cells(Row, 3).Value), Trim(m_ExcelSheet.Cells(Row, 2).Value))
      If Cm.PIG_ID <= 0 Then
         glbErrorLog.LocalErrorMsg = "ไม่พบสุกร '" & Trim(m_ExcelSheet.Cells(Row, 3).Value) & "' ประเภท '" & Trim(m_ExcelSheet.Cells(Row, 2).Value) & "'"
         glbErrorLog.ShowUserError
         
         glbDatabaseMngr.DBConnection.RollbackTrans
         cmdStart.Enabled = True
         cmdExit.Enabled = True
         cmdOK.Enabled = True
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Cm.PIG_STATUS = 0
      Cm.TO_HOUSE_ID = 0
      Cm.COMMIT_FLAG = "Y"
      Call Cm.AddEditData
      
      For Col = 5 To MaxCol
         Set Mi = New CMovementItem
         
         If Trim(m_ExcelSheet.Cells(1, Col).Value) = "EXP" Then
            Mi.AddEditMode = SHOW_ADD
            Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
            Mi.EXPENSE_TYPE = glbDaily.LookupExpenseIDName(Trim(m_ExcelSheet.Cells(2, Col).Value))
            If Mi.EXPENSE_TYPE <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบประเภทรายจ่าย '" & Trim(m_ExcelSheet.Cells(2, Col).Value) & "'"
               glbErrorLog.ShowUserError
               
               glbDatabaseMngr.DBConnection.RollbackTrans
               cmdStart.Enabled = True
               cmdExit.Enabled = True
               cmdOK.Enabled = True
               Call EnableForm(Me, True)
               Exit Sub
            End If
            Mi.PART_ITEM_ID = 0
            Mi.CAPITAL_AMOUNT = Val(m_ExcelSheet.Cells(Row, Col).Value)
            Call Mi.AddEditData
         ElseIf Trim(m_ExcelSheet.Cells(1, Col).Value) = "PG" Then
            Mi.AddEditMode = SHOW_ADD
            Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
            Mi.PART_ITEM_ID = glbDaily.LookupPartItemFromPartGroup(Trim(m_ExcelSheet.Cells(2, Col).Value))
            If Mi.PART_ITEM_ID <= 0 Then
               glbErrorLog.LocalErrorMsg = "ไม่พบกลุ่มวัตถุดิบ '" & Trim(m_ExcelSheet.Cells(2, Col).Value) & "'"
               glbErrorLog.ShowUserError
               
               glbDatabaseMngr.DBConnection.RollbackTrans
               cmdStart.Enabled = True
               cmdExit.Enabled = True
               cmdOK.Enabled = True
               Call EnableForm(Me, True)
               Exit Sub
            End If
            Mi.EXPENSE_TYPE = 0
            Mi.CAPITAL_AMOUNT = Val(m_ExcelSheet.Cells(Row, Col).Value)
            Call Mi.AddEditData
         End If
      Next Col

      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
      txtPercent.Text = MyDiffEx(ProgressCount * 100, MaxRow + 1)
      Set Cm = Nothing
   Next Row
   
   prgProgress.Value = prgProgress.MAX
   txtPercent.Text = "100"
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
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub ImportCustomer()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim Cm As CCustomer
Dim CstName As CCustomerName
Dim Name As cName
Dim Acc As CAccount
Dim IsOK As Boolean

   HasBegin = False

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

      If Trim(m_ExcelSheet.Cells(Row, 1).Value) <> "C-0000" Then
         Set Cm = New CCustomer
         
         Cm.AddEditMode = SHOW_ADD
         Cm.CUSTOMER_CODE = Trim(m_ExcelSheet.Cells(Row, 1).Value)
         Cm.CREDIT = Val(Trim(m_ExcelSheet.Cells(Row, 3).Value))
         Cm.CUSTOMER_NAME = Trim(m_ExcelSheet.Cells(Row, 2).Value)
         Cm.CUSTOMER_GRADE = -1
         Cm.CUSTOMER_TYPE = -1
         
         If Cm.CstNames.Count <= 0 Then
            Set CstName = New CCustomerName
            CstName.Flag = "A"
            
            Set Name = CstName.Name
            Name.LONG_NAME = Cm.CUSTOMER_NAME
            Name.SHORT_NAME = ""
            Name.Flag = "A"
            
            Call Cm.CstNames.Add(CstName)
         End If
         
         If Cm.CstAccounts.Count <= 0 Then
            Set Acc = New CAccount
            Acc.ACCOUNT_NO = Cm.CUSTOMER_CODE
            Acc.Flag = "A"
            
            Call Cm.CstAccounts.Add(Acc)
            Set Acc = Nothing
         End If
         
         Call glbDaily.AddEditCustomer(Cm, IsOK, False, glbErrorLog)
         
         ProgressCount = ProgressCount + 1
         prgProgress.Value = ProgressCount
         
         Set Cm = Nothing
      End If
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
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Function GetAccount(Col As Collection, AccNo As String) As CAccount
On Error Resume Next

   Set GetAccount = Col(AccNo)
End Function

Private Sub ImportBalance()
'On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim Bd As CBillingDoc
Dim IsOK As Boolean
Dim Accounts As Collection
Dim Partitems As Collection
Dim Ac As CAccount
Dim Di As CDoItem

   HasBegin = False

   ID = 1
   
   Set Accounts = New Collection
   Call LoadAccountEx(Nothing, Accounts)
   
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

      Set Bd = New CBillingDoc
      
      Set Ac = GetAccount(Accounts, Trim(m_ExcelSheet.Cells(Row, 1).Value))
      If Ac Is Nothing Then
         glbErrorLog.LocalErrorMsg = Trim(m_ExcelSheet.Cells(Row, 1).Value) & " --> " & Trim(m_ExcelSheet.Cells(Row, 2).Value)
         glbErrorLog.ShowUserError
         Set Ac = GetAccount(Accounts, "C-0000")
      End If
      
      Bd.AddEditMode = SHOW_ADD
      Bd.DOCUMENT_NO = Trim(m_ExcelSheet.Cells(Row, 2).Value) & "."
      Bd.DOCUMENT_DATE = m_ExcelSheet.Cells(Row, 3).Value
      Bd.ACCOUNT_ID = Ac.ACCOUNT_ID
      Bd.DOCUMENT_TYPE = 1
      Bd.DOCUMENT_SUBTYPE = 1
      Bd.RECEIPT_TYPE = 0
      Bd.COMMIT_FLAG = "N"
      Bd.EXCEPTION_FLAG = "N"
      
      Set Di = New CDoItem
      Di.ITEM_AMOUNT = 1
      Di.PART_ITEM_ID = glbDaily.LookupPigID("254400", "N")
      Di.LOCATION_ID = glbDaily.LookupLocationID("00", "Y", 1)
      Di.TOTAL_PRICE = Val(Trim(m_ExcelSheet.Cells(Row, 6).Value))
      
      Di.Flag = "A"
      Call Bd.DoItems.Add(Di)
      Set Di = Nothing
      
      Call glbDaily.AddEditBillingDoc(Bd, IsOK, False, glbErrorLog)
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
      
      Set Bd = Nothing
   Next Row
   
   prgProgress.Value = prgProgress.MAX
   
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   HasBegin = False
   
   Set m_ExcelSheet = Nothing
   Set Accounts = Nothing
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   Exit Sub
   
'ErrorHandler:
'   If HasBegin Then
'      glbDatabaseMngr.DBConnection.RollbackTrans
'   End If
'
'   Call EnableForm(Me, True)
'
'   cmdStart.Enabled = True
'   cmdExit.Enabled = True
'   cmdOK.Enabled = True
'
'   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowUserError
End Sub

Private Sub UpdateARCredit()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim Bd As CBillingDoc
Dim IsOK As Boolean
Dim Cm As CCustomer

   HasBegin = False

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
   
   For Row = 1 To MaxRow
      DoEvents

      Set Cm = New CCustomer
      Cm.CUSTOMER_CODE = Trim(m_ExcelSheet.Cells(Row, 1).Value)
      Cm.CREDIT = Val(Trim(m_ExcelSheet.Cells(Row, 2).Value))
      Call Cm.UpdateARCreditByCode
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
      prgProgress.Refresh
      
      Set Cm = Nothing
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
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdOK.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub cmdStart_Click()
   If radStockImport.Value Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName1.Text)
      
      Call ImportStock
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf radPigImport.Value Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName2.Text)
      
      Call ImportPig
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf radCapitalImport.Value Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName3.Text)
      
      Call ImportCapital
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf radCustomer.Value Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName4.Text)
      
      Call ImportCustomer
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf radBalance.Value Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName5.Text)
      
      Call ImportBalance
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf radARCredit.Value Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName6.Text)
      
      Call UpdateARCredit
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   End If
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
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
   pnlHeader.Caption = "นำเข้ายอดยกมา"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitOptionEx(radStockImport, "นำเข้ายอดยกมาวัตถุดิบ")
   Call InitOptionEx(radPigImport, "นำเข้ายอดยกมาสุกร")
   Call InitOptionEx(radCapitalImport, "นำเข้ายอดยต้นทุนยกมา")
   Call InitOptionEx(radCustomer, "นำเข้าลูกหนี้")
   Call InitOptionEx(radBalance, "นำเข้ายอดหนี้ยกมา")
   Call InitOptionEx(radARCredit, "อัพเดตเครดิตลูกหนี้")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName1.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName1.Enabled = False
   Call txtFileName2.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName2.Enabled = False
   Call txtFileName3.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName3.Enabled = False
   Call txtFileName4.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName4.Enabled = False
   Call txtFileName5.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName5.Enabled = False
   Call txtFileName6.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName6.Enabled = False

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName1.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName3.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName4.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName5.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName6.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName1, MapText("..."))
   Call InitMainButton(cmdFileName2, MapText("..."))
   Call InitMainButton(cmdFileName3, MapText("..."))
   Call InitMainButton(cmdFileName4, MapText("..."))
   Call InitMainButton(cmdFileName5, MapText("..."))
   Call InitMainButton(cmdFileName6, MapText("..."))
   
   radStockImport.Value = True
   
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
   
   m_HasActivate = False
   Call InitFormLayout
   Set m_ExcelApp = CreateObject("Excel.application")
   
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

Private Sub radARCredit_Click(Value As Integer)
   cmdFileName1.Enabled = Not Value
   cmdFileName2.Enabled = Not Value
   cmdFileName3.Enabled = Not Value
   cmdFileName4.Enabled = Not Value
   cmdFileName5.Enabled = Not Value
   cmdFileName6.Enabled = Value
End Sub

Private Sub radBalance_Click(Value As Integer)
   cmdFileName1.Enabled = Not Value
   cmdFileName2.Enabled = Not Value
   cmdFileName3.Enabled = Not Value
   cmdFileName4.Enabled = Not Value
   cmdFileName5.Enabled = Value
   cmdFileName6.Enabled = Not Value
End Sub

Private Sub radCapitalImport_Click(Value As Integer)
   cmdFileName1.Enabled = Not Value
   cmdFileName2.Enabled = Not Value
   cmdFileName3.Enabled = Value
   cmdFileName4.Enabled = Not Value
   cmdFileName5.Enabled = Not Value
   cmdFileName6.Enabled = Not Value
End Sub

Private Sub radCustomer_Click(Value As Integer)
   cmdFileName1.Enabled = Not Value
   cmdFileName2.Enabled = Not Value
   cmdFileName3.Enabled = Not Value
   cmdFileName4.Enabled = Value
   cmdFileName5.Enabled = Not Value
   cmdFileName6.Enabled = Not Value
End Sub

Private Sub radPigImport_Click(Value As Integer)
   cmdFileName1.Enabled = Not Value
   cmdFileName2.Enabled = Value
   cmdFileName3.Enabled = Not Value
   cmdFileName4.Enabled = Not Value
   cmdFileName5.Enabled = Not Value
   cmdFileName6.Enabled = Not Value
End Sub

Private Sub radStockImport_Click(Value As Integer)
   cmdFileName1.Enabled = Value
   cmdFileName2.Enabled = Not Value
   cmdFileName3.Enabled = Not Value
   cmdFileName4.Enabled = Not Value
   cmdFileName5.Enabled = Not Value
   cmdFileName6.Enabled = Not Value
End Sub
