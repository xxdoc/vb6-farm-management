VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBatch 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmBatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlJournalDate 
         Height          =   405
         Left            =   5940
         TabIndex        =   1
         Top             =   960
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1950
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1950
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   15
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5175
         Left            =   180
         TabIndex        =   7
         Top             =   2550
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9128
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmBatch.frx":27A2
         Column(2)       =   "frmBatch.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBatch.frx":290E
         FormatStyle(2)  =   "frmBatch.frx":2A6A
         FormatStyle(3)  =   "frmBatch.frx":2B1A
         FormatStyle(4)  =   "frmBatch.frx":2BCE
         FormatStyle(5)  =   "frmBatch.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBatch.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtJournalCode 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   960
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdRun 
         Height          =   525
         Left            =   6810
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBatch.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkPostFlag 
         Height          =   405
         Left            =   1560
         TabIndex        =   2
         Top             =   1470
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJournalDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4440
         TabIndex        =   19
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label lblJournalCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   17
         Top             =   2010
         Width           =   1185
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   2010
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   5
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBatch.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   1500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBatch.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBatch.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   9
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   13
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBatch.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Batch As CBatch
Private m_TempBatch As CBatch
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public OKClick As Boolean

Public HeaderText As String
Public ApArInd As Long
Private ApArText As String

Public ParamArea As Long
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   If Not VerifyAccessRight("SIMULATE_BATCH_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmAddEditBatch.ParamArea = ParamArea
   frmAddEditBatch.HeaderText = MapText("เพิ่มข้อมูลแบต Simulate")
   frmAddEditBatch.ShowMode = SHOW_ADD
   Load frmAddEditBatch
   frmAddEditBatch.Show 1

   OKClick = frmAddEditBatch.OKClick

   Unload frmAddEditBatch
   Set frmAddEditBatch = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtJournalCode.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyAccessRight("SIMULATE_BATCH_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   Call m_Batch.SetFieldValue("BATCH_ID", ID)
   If Not glbDaily.DeleteBatch(m_Batch, IsOK, True, glbErrorLog) Then
      Call m_Batch.SetFieldValue("BATCH_ID", -1)
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   
   frmAddEditBatch.ParamArea = ParamArea
   frmAddEditBatch.ID = ID
   frmAddEditBatch.HeaderText = MapText("แก้ไขข้อมูลแบต Simulate")
   frmAddEditBatch.ShowMode = SHOW_EDIT
   Load frmAddEditBatch
   frmAddEditBatch.Show 1

   OKClick = frmAddEditBatch.OKClick

   Unload frmAddEditBatch
   Set frmAddEditBatch = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdRun_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ID As Long
Dim IsOK As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("รัน SIMULATE", "-", "เคลียร์ SIMULATE", "-", "เคลียร์ SIMULATE ทั้งหมด")
   Set oMenu = Nothing
   
   If lMenuChosen <= 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      frmRunSimulate.ID = ID
      frmRunSimulate.ShowMode = SHOW_EDIT
      frmRunSimulate.HeaderText = "รัน SIMULATE"
      Load frmRunSimulate
      frmRunSimulate.Show 1
      
      Unload frmRunSimulate
      Set frmRunSimulate = Nothing
   ElseIf lMenuChosen = 3 Then
      glbErrorLog.LocalErrorMsg = "ต้องการที่จะทำการลบเอกสารทั้งหมดของ Batch นี้ใช่หรือไม่"
      If glbErrorLog.AskMessage = vbNo Then
         Exit Sub
      End If
      
      Call glbDaily.DeleteDocument(ID, IsOK, True, glbErrorLog)
   ElseIf lMenuChosen = 5 Then
      glbErrorLog.LocalErrorMsg = "ต้องการที่จะทำการลบเอกสารทั้งหมดของ ทุก Batch ใช่หรือไม่"
      If glbErrorLog.AskMessage = vbNo Then
         Exit Sub
      End If
      
      glbErrorLog.LocalErrorMsg = "กระบวนการลบเอกสารทั้งหมดนี้จะใช้เวลานานในการทำการ กรุณายืนยันที่จะทำการลบเอกสารทั้งหมด"
      If glbErrorLog.AskMessage = vbNo Then
         Exit Sub
      End If
      
      Call glbDaily.DeleteDocument(ID, IsOK, True, glbErrorLog, True)
   End If
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitBatchOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_Batch.SetFieldValue("BATCH_ID", -1)
      Call m_Batch.SetFieldValue("FROM_DATE", uctlJournalDate.ShowDate)
      Call m_Batch.SetFieldValue("TO_DATE", uctlJournalDate.ShowDate)
      Call m_Batch.SetFieldValue("BATCH_NO", txtJournalCode.Text)
      Call m_Batch.SetFieldValue("COMMIT_FLAG", Check2Flag(chkPostFlag.Value))
      Call m_Batch.SetFieldValue("ORDER_BY", cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex)))
      Call m_Batch.SetFieldValue("ORDER_TYPE", cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex)))
      If Not glbDaily.QueryBatch(m_Batch, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   If ItemCount > 0 Then
      cmdAdd.Enabled = False
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2460
   Col.Caption = MapText("เลขที่")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2850
   Col.Caption = MapText("วันที่")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 6165
   Col.Caption = MapText("รายละเอียด")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("COMMIT_FLAG")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
      
   Call InitGrid
   
   Call InitNormalLabel(lblJournalDate, MapText("วันที่"))
   Call InitNormalLabel(lblJournalCode, MapText("เลขที่"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Call InitCheckBox(chkPostFlag, "ห้ามแก้ไข")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdRun.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdRun, MapText("SIMULATE"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_Batch = New CBatch
   Set m_TempBatch = New CBatch
   Set m_Rs = New ADODB.Recordset
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(5)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempBatch.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempBatch.GetFieldValue("BATCH_ID")
   Values(2) = m_TempBatch.GetFieldValue("BATCH_NO")
   Values(3) = DateToStringExtEx2(m_TempBatch.GetFieldValue("BATCH_DATE"))
   Values(4) = m_TempBatch.GetFieldValue("BATCH_DESC")
   Values(5) = m_TempBatch.GetFieldValue("COMMIT_FLAG")
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim Pk As CBatch
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("COPY BATCH")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set Pk = New CBatch
      Call Pk.SetFieldValue("BATCH_ID", TempID1)
      Call glbDaily.CopyBatch(Pk, IsOK, True, glbErrorLog)
      Call QueryData(True)
      Set Pk = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub
