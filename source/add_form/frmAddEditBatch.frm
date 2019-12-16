VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditBatch 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditBatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlJournalDate 
         Height          =   405
         Left            =   7470
         TabIndex        =   2
         Top             =   570
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   7
         Top             =   2400
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtJournalCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   540
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtJournalDesc 
         Height          =   450
         Left            =   1860
         TabIndex        =   5
         Top             =   1410
         Width           =   8145
         _ExtentX        =   16907
         _ExtentY        =   794
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4815
         Left            =   150
         TabIndex        =   8
         Top             =   2940
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8493
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
         Column(1)       =   "frmAddEditBatch.frx":27A2
         Column(2)       =   "frmAddEditBatch.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditBatch.frx":290E
         FormatStyle(2)  =   "frmAddEditBatch.frx":2A6A
         FormatStyle(3)  =   "frmAddEditBatch.frx":2B1A
         FormatStyle(4)  =   "frmAddEditBatch.frx":2BCE
         FormatStyle(5)  =   "frmAddEditBatch.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditBatch.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   465
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   820
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   3
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   7470
         TabIndex        =   4
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtBalanceCash 
         Height          =   435
         Left            =   1860
         TabIndex        =   21
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCashFirstMonth 
         Height          =   435
         Left            =   4740
         TabIndex        =   23
         Top             =   1860
         Width           =   1305
         _ExtentX        =   4419
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtOtherFirstMonth 
         Height          =   435
         Left            =   10620
         TabIndex        =   25
         Top             =   1860
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMedicineFirstMonth 
         Height          =   435
         Left            =   7620
         TabIndex        =   27
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
      End
      Begin VB.Label lblMedicineFirstMonth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6120
         TabIndex        =   28
         Top             =   1995
         Width           =   1455
      End
      Begin VB.Label lblOtherFirstMonth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9000
         TabIndex        =   26
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblCashFirstMonth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3240
         TabIndex        =   24
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblBalanceCash 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5970
         TabIndex        =   20
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   1020
         Width           =   1365
      End
      Begin Threed.SSCheck chkPostFlag 
         Height          =   405
         Left            =   10170
         TabIndex        =   6
         Top             =   1410
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJournalDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6270
         TabIndex        =   18
         Top             =   570
         Width           =   1065
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4170
         TabIndex        =   1
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBatch.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBatch.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   13
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   10
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBatch.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditBatch.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblJournalDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   16
         Top             =   1530
         Width           =   1695
      End
      Begin VB.Label lblJournalCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   15
         Top             =   660
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAddEditBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Batch As CBatch
Private m_ApArMass As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ParamArea As Long

Private ApArText As String
Private FileName As String

Public JournalType As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_Batch.SetFieldValue("BATCH_ID", ID)
      If Not glbDaily.QueryBatch(m_Batch, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Batch.PopulateFromRS(1, m_Rs)
      txtJournalCode.Text = m_Batch.GetFieldValue("BATCH_NO")
      txtJournalDesc.Text = m_Batch.GetFieldValue("BATCH_DESC")
      chkPostFlag.Value = FlagToCheck(m_Batch.GetFieldValue("COMMIT_FLAG"))
      uctlJournalDate.ShowDate = m_Batch.GetFieldValue("BATCH_DATE")
      uctlFromDate.ShowDate = m_Batch.GetFieldValue("EXECUTE_FROM")
      uctlToDate.ShowDate = m_Batch.GetFieldValue("EXECUTE_TO")
      txtBalanceCash.Text = m_Batch.GetFieldValue("BALANCE_CASH")
      txtCashFirstMonth.Text = m_Batch.GetFieldValue("CASH_FIRST_MONTH")
      txtMedicineFirstMonth.Text = m_Batch.GetFieldValue("MEDICINE_FIRST_MONTH")
      txtOtherFirstMonth.Text = m_Batch.GetFieldValue("OTHER_FIRST_MONTH")
   Else
      ShowMode = SHOW_ADD
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
      If Not VerifyAccessRight("SIMULATE_BATCH_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("SIMULATE_BATCH_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblJournalCode, txtJournalCode, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblJournalDate, uctlJournalDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(CUSTCODE_UNIQUE, txtJournalCode.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJournalCode.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Batch.ShowMode = ShowMode
   Call m_Batch.SetFieldValue("BATCH_DATE", uctlJournalDate.ShowDate)
   Call m_Batch.SetFieldValue("COMMIT_FLAG", Check2Flag(chkPostFlag.Value))
   Call m_Batch.SetFieldValue("BATCH_NO", txtJournalCode.Text)
   Call m_Batch.SetFieldValue("BATCH_DESC", txtJournalDesc.Text)
   Call m_Batch.SetFieldValue("EXECUTE_FROM", uctlFromDate.ShowDate)
   Call m_Batch.SetFieldValue("EXECUTE_TO", uctlToDate.ShowDate)
   Call m_Batch.SetFieldValue("BALANCE_CASH", Val(txtBalanceCash.Text))
   Call m_Batch.SetFieldValue("CASH_FIRST_MONTH", Val(txtCashFirstMonth.Text))    ' ค่าอาหารที่ต้องจ่ายเดือน1
   Call m_Batch.SetFieldValue("MEDICINE_FIRST_MONTH", Val(txtMedicineFirstMonth.Text))    ' ค่ายาที่ต้องจ่ายเดือน1
   Call m_Batch.SetFieldValue("OTHER_FIRST_MONTH", Val(txtOtherFirstMonth.Text))    'อื่นๆที่ต้องจ่ายเดือน1
   
   Call MergeItem(m_Batch)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditBatch(m_Batch, IsOK, True, glbErrorLog) Then
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

Private Sub MergeItem(B As CBatch)
Dim Bi As CBatchItem

   Set B.BatchItems = Nothing
   Set B.BatchItems = New Collection
   
   For Each Bi In B.BirthItems
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.FoodItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.SaleItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.TransferItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.WeightItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.Feeds
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.Balances
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.Revenues
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.CustRatios
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.ChangePigTypes
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.BuyItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.ExpenseSharingItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.PigAdjItems
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.ManagementExpenses
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.Glages
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.GLbacks
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   
End Sub
Private Sub chkPostFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkPostFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   
   frmAddBatchItem.Area = TabStrip1.SelectedItem.Tag
   Set frmAddBatchItem.TempCollection = GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag)
   frmAddBatchItem.ShowMode = SHOW_ADD
   frmAddBatchItem.HeaderText = MapText("เลือกพารามิเตอร์")
   Load frmAddBatchItem
   frmAddBatchItem.Show 1

   OKClick = frmAddBatchItem.OKClick

   Unload frmAddBatchItem
   Set frmAddBatchItem = Nothing

   If OKClick Then
      GridEX1.ItemCount = CountItem(GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag))
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String

'   If Trim(txtJournalCode.Text) = "" Then
'      Call glbDatabaseMngr.GenerateNumber(CUSTOMER_NUMBER, No, glbErrorLog)
'      txtJournalCode.Text = No
'   End If
End Sub
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If ID1 <= 0 Then
      GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag).Remove (ID2)
   Else
      GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag).Item(ID2).Flag = "D"
   End If
   
   GridEX1.ItemCount = CountItem(GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag))
   GridEX1.Rebind
   m_HasModify = True
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim ParamID As Long
Dim OKClick As Boolean
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   ParamID = Val(GridEX1.Value(6))
   OKClick = False
   
   frmAddEditParameter.HeaderText = TabStrip1.SelectedItem.Caption
   frmAddEditParameter.ParamArea = TabStrip1.SelectedItem.Tag
   frmAddEditParameter.ShowMode = SHOW_VIEW_ONLY
   frmAddEditParameter.ID = ParamID
   Load frmAddEditParameter
   frmAddEditParameter.Show 1
   
   Unload frmAddEditParameter
   Set frmAddEditParameter = Nothing
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_Batch.GetFieldValue("BATCH_ID")
      m_Batch.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
   
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
                  
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Batch.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Batch.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Batch = Nothing
   Set m_ApArMass = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1(Ind As Long)
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   If (Ind = 1) Or (Ind = 2) Or (Ind = 3) Or (Ind = 4) Or (Ind = 5) Or (Ind = 6) Or (Ind = 7) Or (Ind = 9) Or (Ind = 10) Or (Ind = 11) Or (Ind = 12) Or (Ind = 13) Or (Ind = 14) Or (Ind = 15) Or (Ind = 16) Or (Ind = 17) Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2310
      Col.Caption = MapText("เลขที่")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2490
      Col.Caption = MapText("วันที่")
   
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 6465
      Col.Caption = MapText("รายละเอียด")
   
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 0
      Col.Visible = False
      Col.Caption = MapText("PARAM_ID")
   End If
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2925
   Col.Caption = MapText("เลขที่บัญชี")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 6270
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3240
   Col.Caption = MapText("แพคเกจ")
End Sub

Private Sub InitFormLayout()
Dim Obj As Object

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblJournalCode, MapText("เลขที่"))
   Call InitNormalLabel(lblJournalDate, MapText("วันที่"))
   Call InitNormalLabel(lblJournalDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblBalanceCash, MapText("ยกมาเงินสด"))
   Call InitNormalLabel(lblCashFirstMonth, MapText("จ.อาหารเดือน1"))
   Call InitNormalLabel(lblMedicineFirstMonth, MapText("จ.ยาเดือน1"))
   Call InitNormalLabel(lblOtherFirstMonth, MapText("จ.อื่นๆเดือน1"))
   
   Call txtJournalCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtJournalDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitCheckBox(chkPostFlag, "ห้ามแก้ไข")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   Call InitGrid1(1)
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("สุกรเกิด")
   Obj.Tag = 1
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("อาหาร/ยา")
   Obj.Tag = 2
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("การโอน (สูญเสีย)")
   Obj.Tag = 3
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("การขาย")
   Obj.Tag = 4
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("ราคาอาหาร/ยา")
   Obj.Tag = 6
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("ยอดยกมาสุกร")
   Obj.Tag = 7
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("ยกมา GL")
   Obj.Tag = 16
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("G กลับสัตว์")
   Obj.Tag = 17
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("ขายอื่น ๆ")
   Obj.Tag = 9

   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("% การขาย")
   Obj.Tag = 10

   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("เปลี่ยนประเภทสุกร")
   Obj.Tag = 11

   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("ซื้อสุกร")
   Obj.Tag = 12

   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("ปันค่าใช้จ่าย")
   Obj.Tag = 13

   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("คุมยอดสุกร")
   Obj.Tag = 14
   
   Set Obj = TabStrip1.Tabs.Add()
   Obj.Caption = MapText("คชจ ขายบริหาร")
   Obj.Tag = 15
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Batch = New CBatch
   Set m_ApArMass = New Collection
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Batch.BatchItems Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CBatchItem
   If GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag).Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag), RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.GetFieldValue("BATCH_ITEM_ID")
   Values(2) = RealIndex
   Values(3) = CR.GetFieldValue("PARAM_NO")
   Values(4) = DateToStringExtEx2(CR.GetFieldValue("PARAM_DATE"))
   Values(5) = CR.GetFieldValue("PARAM_DESC")
   Values(6) = CR.GetFieldValue("PARAM_ID")
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub RefreshGrid(Flag As Boolean)
   GridEX1.ItemCount = CountItem(GetCollectionFromType(m_Batch, TabStrip1.SelectedItem.Tag))
   GridEX1.Rebind
   
   If Flag Then
      m_HasModify = Flag
   End If
End Sub
Private Sub TabStrip1_Click()
   Call InitGrid1(TabStrip1.SelectedItem.Tag)
   Call RefreshGrid(False)
End Sub
Private Sub txtBalanceCash_Change()
   m_HasModify = True
End Sub

Private Sub txtCashFirstMonth_Change()
   m_HasModify = True
End Sub

Private Sub txtJournalDesc_Change()
   m_HasModify = True
End Sub
Private Sub txtJournalCode_Change()
   m_HasModify = True
End Sub

Private Sub txtMedicineFirstMonth_Change()
   m_HasModify = True
End Sub

Private Sub txtOtherFirstMonth_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlJournalDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
