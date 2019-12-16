VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLoaderMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmLoaderMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   4020
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   1590
      TabIndex        =   16
      Top             =   4020
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Height          =   495
      Left            =   5940
      TabIndex        =   1
      Top             =   4020
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7515
      Begin VB.CheckBox Check1 
         Height          =   315
         Left            =   1140
         TabIndex        =   19
         Top             =   1620
         Width           =   3735
      End
      Begin VB.TextBox txtFileName 
         Height          =   375
         Left            =   1110
         TabIndex        =   18
         Top             =   360
         Width           =   5505
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5160
         Top             =   1260
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chkFlag 
         Height          =   315
         Left            =   1140
         TabIndex        =   15
         Top             =   1260
         Width           =   3735
      End
      Begin VB.CommandButton cmdImport 
         Height          =   375
         Left            =   6690
         TabIndex        =   5
         Top             =   390
         Width           =   555
      End
      Begin VB.ComboBox cboTable 
         Height          =   330
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   3735
      End
      Begin VB.Label lblFile 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   675
      End
      Begin VB.Label lblTable 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   210
         TabIndex        =   3
         Top             =   870
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   0
      TabIndex        =   6
      Top             =   1860
      Width           =   7515
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   1110
         TabIndex        =   7
         Top             =   270
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblError 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         TabIndex        =   14
         Top             =   1530
         Width           =   765
      End
      Begin VB.Label lblSuccess 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   270
         TabIndex        =   13
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   240
         TabIndex        =   12
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblErrorCount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1110
         TabIndex        =   11
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label lblSuccessCount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1110
         TabIndex        =   10
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblProgressCount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1110
         TabIndex        =   9
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmLoaderMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ExcelApp As Object
Private m_ExcelSheet As Object

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdImport_Click()
Dim FileName As String
Dim MaxRow As Long
Dim MaxColumn As Long
Dim i As Long

   CommonDialog1.DefaultExt = "*.xls"
   CommonDialog1.Filter = "*.xls"
   CommonDialog1.FileName = ""
   CommonDialog1.ShowOpen
   FileName = CommonDialog1.FileName
   If FileName = "" Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_ExcelApp.Workbooks.Close
   
   txtFileName.Text = FileName
   m_ExcelApp.Workbooks.Open (txtFileName.Text)
   
   cboTable.Clear
   For i = 0 To m_ExcelApp.Worksheets.Count - 1
      cboTable.AddItem (m_ExcelApp.Worksheets.Item(i + 1).Name)
      cboTable.ItemData(i) = i
   Next i
   Call EnableForm(Me, True)
   
End Sub

Private Function InsertData(SQL As String, Conn As ADODB.Connection) As Boolean
On Error GoTo ErrorHandler
Dim RName As String

   InsertData = False
   RName = "InsertData"
   glbErrorLog.RoutineName = RName
   
   Conn.Execute (SQL)
   
   InsertData = True
   Exit Function
   
ErrorHandler:
   InsertData = False
   glbErrorLog.LocalErrorMsg = SQL
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
End Function

Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim i As Long
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

   HasBegin = False
   If (cboTable.ListCount <= 0) Or (cboTable.ListIndex < 0) Then
      glbErrorLog.LocalErrorMsg = "กรุณาเลือกตารางที่ต้องการก่อน"
      glbErrorLog.ShowUserError
      cboTable.SetFocus
      Exit Sub
   End If
   ID = cboTable.ListIndex + 1
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
      
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count

   If (MaxCol < 1) Or (MaxRow < 2) Then
      glbErrorLog.LocalErrorMsg = "รูปแบบไฟล์ไม่ถูกต้อง"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If MaxRow = 2 Then
      glbErrorLog.LocalErrorMsg = "รูปแบบไฟล์ไม่ถูกต้อง"
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   ReDim FieldNames(MaxCol)
   ReDim FieldTypes(MaxCol)
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   cmdImport.Enabled = False
   
   TabField = " ("
   For i = 1 To MaxCol
      FieldTypes(i - 1) = Trim(m_ExcelSheet.Cells(2, i))
      If i > 1 Then
         TabField = TabField & "," & Trim(m_ExcelSheet.Cells(1, i))
      Else
         TabField = TabField & Trim(m_ExcelSheet.Cells(1, i))
      End If
   Next i
   TabField = TabField & ")"

    
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   ProgressBar1.Min = 1
   ProgressBar1.Max = (MaxRow - 2) + 1
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   
   If Check1.Value = 1 Then
      StateMent = "DELETE FROM " & m_ExcelApp.Sheets(ID).Name
      glbDatabaseMngr.DBConnection.Execute (StateMent)
   End If
   
   For Row = 3 To MaxRow
      DoEvents
      StateMent = "INSERT INTO " & m_ExcelApp.Sheets(ID).Name & TabField & " VALUES ("
      For Col = 1 To MaxCol
         If UCase(FieldTypes(Col - 1)) = "S" Then
            NewValue = "'" & Replace(Trim(m_ExcelSheet.Cells(Row, Col).Value), "'", "''") & "'"
         ElseIf UCase(FieldTypes(Col - 1)) = "N" Then
            NewValue = "NULL"
         ElseIf UCase(FieldTypes(Col - 1)) = "I" Then
            NewValue = Trim(m_ExcelSheet.Cells(Row, Col).Value)
         ElseIf UCase(FieldTypes(Col - 1)) = "SD" Then 'sysdate
            Call glbDatabaseMngr.GetServerDateTime(ServerDtm, glbErrorLog)
            NewValue = "'" & ServerDtm & "'"
         ElseIf UCase(FieldTypes(Col - 1)) = "UID" Then 'user id
            NewValue = glbUser.USER_ID
         End If
         If Col > 1 Then
            StateMent = StateMent & "," & NewValue
         Else
            StateMent = StateMent & NewValue
         End If
      Next Col
      StateMent = StateMent & ")"
      
      ErrorFlag = False
      If Not InsertData(StateMent, glbDatabaseMngr.DBConnection) Then
         ErrorCount = ErrorCount + 1
         ErrorFlag = True
      Else
         SuccessCount = SuccessCount + 1
      End If
      ProgressCount = ProgressCount + 1
      ProgressBar1.Value = ProgressCount
      
      lblProgressCount.Caption = ProgressCount
      lblErrorCount.Caption = ErrorCount
      lblSuccessCount.Caption = SuccessCount
      
      If ErrorFlag Then
         If chkFlag.Value = 1 Then
            glbDatabaseMngr.DBConnection.RollbackTrans
            Call EnableForm(Me, True)
            
            cmdStart.Enabled = True
            cmdExit.Enabled = True
            cmdImport.Enabled = True
            Exit Sub
         End If
      End If
   Next Row
   ProgressBar1.Value = ProgressBar1.Max
   
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   HasBegin = False
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdImport.Enabled = True
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   cmdImport.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
End Sub

Private Sub Form_Load()
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   Frame2.BackColor = GLB_FORM_COLOR

   Call InitMainButtonOld(cmdStart, "เริ่ม")
   Call InitMainButtonOld(cmdCancel, "ยกเลิก")
   Call InitMainButtonOld(cmdExit, "ออก")
   Call InitMainButtonOld(cmdImport, "...")

   Call InitNormalLabel(lblFile, "ไฟล์")
   Call InitNormalLabel(lblTable, "ตาราง")
   Call InitNormalLabel(lblStatus, "สถานะ")
   Call InitNormalLabel(lblProgress, "คืบหน้า")
   Call InitNormalLabel(lblSuccess, "สำเร็จ")
   Call InitNormalLabel(lblError, "ผิดพลาด")
   
   Call InitTextBox(txtFileName, "")
   Call SetEnableDisableTextBox(txtFileName, False)
   
   Call InitCombo(cboTable)
   Call InitCheckBox(chkFlag, "จบการทำงานเมื่อผิดพลาด")
   Call InitCheckBox(Check1, "ลบข้อมูลเก่าก่อนทุกครั้ง")
   
   Set m_ExcelApp = CreateObject("Excel.application")
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_ExcelApp.Workbooks.Close
   m_ExcelApp.Quit
   Set m_ExcelApp = Nothing
End Sub
