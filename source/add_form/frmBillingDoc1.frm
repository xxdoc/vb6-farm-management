VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBillingDoc1 
   BackColor       =   &H80000000&
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11910
   Icon            =   "frmBillingDoc1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox CboPackageType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2250
         Width           =   6915
      End
      Begin VB.ComboBox cboBatch 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1770
         Width           =   2625
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1350
         Width           =   2535
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2790
         Width           =   2625
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2760
         Width           =   3015
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   19
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4125
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   7276
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
         Column(1)       =   "frmBillingDoc1.frx":27A2
         Column(2)       =   "frmBillingDoc1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBillingDoc1.frx":290E
         FormatStyle(2)  =   "frmBillingDoc1.frx":2A6A
         FormatStyle(3)  =   "frmBillingDoc1.frx":2B1A
         FormatStyle(4)  =   "frmBillingDoc1.frx":2BCE
         FormatStyle(5)  =   "frmBillingDoc1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBillingDoc1.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   840
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   6180
         TabIndex        =   1
         Top             =   840
         Width           =   2625
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAccountNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1770
         Width           =   2985
         _ExtentX        =   4630
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   3
         Top             =   1350
         Width           =   2535
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblPackageType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   30
         TabIndex        =   28
         Top             =   2370
         Width           =   1755
      End
      Begin VB.Label lblBatch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4950
         TabIndex        =   26
         Top             =   1830
         Width           =   1155
      End
      Begin Threed.SSCommand cmdProcess 
         Height          =   525
         Left            =   6810
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   8880
         TabIndex        =   6
         Top             =   1770
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   25
         Top             =   1950
         Width           =   1755
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   24
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4830
         TabIndex        =   23
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   2850
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10080
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   13
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":3B9E
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBillingDoc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_BillingDoc As CBillingDoc
Private m_TempBillingDoc As CBillingDoc
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean
Public DocumentType As Long
Public Area As Long

Private Sub cmdPasswd_Click()

End Sub
Private Sub cboOrderBy_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboOrderType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not VerifyAccessRight("LEDGER_SELL_" & DocumentType & "_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   If DocumentType = 1 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("ใบส่งของ(สุกร)", "-", "ใบส่งของ(วัตถุดิบ)")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      If lMenuChosen = 1 Then
         frmAddEditDO.DocumentSubType = 1
      ElseIf lMenuChosen = 3 Then
         frmAddEditDO.DocumentSubType = 2
      End If
         frmAddEditDO.DocumentType = DocumentType
      frmAddEditDO.HeaderText = MapText("เพิ่มข้อมูลใบส่งสินค้า")
      frmAddEditDO.ShowMode = SHOW_ADD
      Load frmAddEditDO
      frmAddEditDO.Show 1
      
      OKClick = frmAddEditDO.OKClick
      
      Unload frmAddEditDO
      Set frmAddEditDO = Nothing
   ElseIf DocumentType = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("สร้างรายการใหม่ (สุกร)", "สร้างรายการใหม่ (วัตถุดิบ)", "-", "สร้างจากใบส่งของ", "-", "สร้าจากใบรับเงินชั่วคราว")
      If lMenuChosen = 0 Then
         Exit Sub
      ElseIf lMenuChosen = 6 Then
         glbErrorLog.LocalErrorMsg = "ส่วนฟังก์ชันงานนี้ยังไม่เปิดให้ใช้งาน"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      If (lMenuChosen = 1) Then
         frmAddEditReceipt.DocumentSubType = 1
      ElseIf (lMenuChosen = 2) Then
         frmAddEditReceipt.DocumentSubType = 2
      End If
      
      If lMenuChosen = 4 Then
         frmAddEditReceipt.ReceiptType = 3
      ElseIf (lMenuChosen = 1) Or (lMenuChosen = 2) Then
         frmAddEditReceipt.ReceiptType = 1
      End If
      frmAddEditReceipt.Area = 1
      frmAddEditReceipt.HeaderText = MapText("เพิ่มข้อมูลใบเสร็จ")
      frmAddEditReceipt.ShowMode = SHOW_ADD
      Load frmAddEditReceipt
      frmAddEditReceipt.Show 1
      
      OKClick = frmAddEditReceipt.OKClick
      
      Unload frmAddEditReceipt
      Set frmAddEditReceipt = Nothing
   ElseIf DocumentType = 3 Then
      frmAddEditDebitCreditNote.Area = 1
      frmAddEditDebitCreditNote.DocumentType = DocumentType
      frmAddEditDebitCreditNote.HeaderText = MapText("เพิ่มข้อมูลใบเพิ่มหนี้")
      frmAddEditDebitCreditNote.DebitCreditType = 1
      frmAddEditDebitCreditNote.ShowMode = SHOW_ADD
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   ElseIf DocumentType = 4 Then
      frmAddEditDebitCreditNote.DocumentType = DocumentType
      frmAddEditDebitCreditNote.Area = 1
      frmAddEditDebitCreditNote.HeaderText = MapText("เพิ่มข้อมูลใบลดหนี้")
      frmAddEditDebitCreditNote.DebitCreditType = 2
      frmAddEditDebitCreditNote.ShowMode = SHOW_ADD
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   End If
   
   If OKClick Then
      Call QueryData(False)
   End If
End Sub

Private Sub cmdClear_Click()
   txtDocumentNo.Text = ""
   txtCustomerCode.Text = ""
   txtAccountNo.Text = ""
   
   uctlDocumentDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyAccessRight("LEDGER_SELL_" & DocumentType & "_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteBillingDoc(ID, IsOK, True, glbErrorLog) Then
      m_BillingDoc.BILLING_DOC_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim BatchID As Long
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Not VerifyDateInterval(GridEX1.Value(10)) Then
      Exit Sub
   End If
   
   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   BatchID = Val(GridEX1.Value(9))
   
   If DocumentType = 1 Then
      frmAddEditDO.BATCH_ID = BatchID
      frmAddEditDO.DocumentType = DocumentType
      frmAddEditDO.DocumentSubType = Val(GridEX1.Value(8))
      frmAddEditDO.ID = ID
      frmAddEditDO.HeaderText = MapText("แก้ไขข้อมูลใบส่งสินค้า")
      frmAddEditDO.ShowMode = SHOW_EDIT
      Load frmAddEditDO
      frmAddEditDO.Show 1
      
      OKClick = frmAddEditDO.OKClick
      
      Unload frmAddEditDO
      Set frmAddEditDO = Nothing
   ElseIf DocumentType = 2 Then
      frmAddEditReceipt.BATCH_ID = BatchID
      frmAddEditReceipt.ReceiptType = Val(GridEX1.Value(7))
      frmAddEditReceipt.DocumentSubType = Val(GridEX1.Value(8))
      frmAddEditReceipt.Area = 1
      frmAddEditReceipt.ID = ID
      frmAddEditReceipt.HeaderText = MapText("แก้ไขข้อมูลใบเสร็จ")
      frmAddEditReceipt.ShowMode = SHOW_EDIT
      Load frmAddEditReceipt
      frmAddEditReceipt.Show 1
      
      OKClick = frmAddEditReceipt.OKClick

      Unload frmAddEditReceipt
      Set frmAddEditReceipt = Nothing
   ElseIf DocumentType = 3 Then
      frmAddEditDebitCreditNote.BATCH_ID = BatchID
      frmAddEditDebitCreditNote.ID = ID
      frmAddEditDebitCreditNote.Area = 1
      frmAddEditDebitCreditNote.DocumentType = DocumentType
      frmAddEditDebitCreditNote.HeaderText = MapText("เพิ่มข้อมูลใบเพิ่มหนี้")
      frmAddEditDebitCreditNote.DebitCreditType = 1
      frmAddEditDebitCreditNote.ShowMode = SHOW_EDIT
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   ElseIf DocumentType = 4 Then
      frmAddEditDebitCreditNote.ID = ID
      frmAddEditDebitCreditNote.DocumentType = DocumentType
      frmAddEditDebitCreditNote.Area = 1
      frmAddEditDebitCreditNote.HeaderText = MapText("เพิ่มข้อมูลใบลดหนี้")
      frmAddEditDebitCreditNote.DebitCreditType = 2
      frmAddEditDebitCreditNote.ShowMode = SHOW_EDIT
      Load frmAddEditDebitCreditNote
      frmAddEditDebitCreditNote.Show 1
      
      OKClick = frmAddEditDebitCreditNote.OKClick
      
      Unload frmAddEditDebitCreditNote
      Set frmAddEditDebitCreditNote = Nothing
   End If

   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdProcess_Click()
Dim OKClick As Boolean

   frmProcssCommit.DocumentCategory = 2
   frmProcssCommit.DocumentType = DocumentType
   Load frmProcssCommit
   frmProcssCommit.Show 1
   
   OKClick = frmProcssCommit.OKClick
   
   Unload frmProcssCommit
   Set frmProcssCommit = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      If (glbUser.SIMULATE_FLAG = "Y") Then
         Call LoadBatch(cboBatch)
      End If
            
      Call InitBillingDocOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call InitPackageType(CboPackageType)
            
      Call QueryData(False)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
      m_BillingDoc.CUSTOMER_CODE = txtCustomerCode.Text
      m_BillingDoc.ACCOUNT_NO = txtAccountNo.Text
      m_BillingDoc.FROM_DATE = uctlDocumentDate.ShowDate
      m_BillingDoc.TO_DATE = uctlToDate.ShowDate
      m_BillingDoc.DOCUMENT_TYPE = DocumentType
      m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
      m_BillingDoc.PKG_TYPE = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
      
      m_BillingDoc.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_BillingDoc.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      
      If glbUser.SIMULATE_FLAG = "Y" Then
         m_BillingDoc.BATCH_ID = cboBatch.ItemData(Minus2Zero(cboBatch.ListIndex))
      End If
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      cmdDelete.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long

   If InUsed = 1 Then
      Exit Sub
   End If
   
   InUsed = 1
   
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
   
   InUsed = 0
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
   Col.Width = 2115
   Col.Caption = MapText("เลขที่เอกสาร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2055
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2305
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 4995
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("COMMIT FLAG")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("RECEIPT_TYPE")
   
   Set Col = GridEX1.Columns.Add '8
    If DocumentType = 1 Then
      Col.Width = 1000
   Else
      Col.Width = 0
   End If
   Col.Caption = MapText("ประเภทการขาย")
   
   Set Col = GridEX1.Columns.Add '9
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("BATCH_ID")
   
   Set Col = GridEX1.Columns.Add '10
   Col.Width = 0
   Col.Caption = MapText("วันที่เอกสารสำหรับกำหนดช่วง")
   
   If DocumentType = 1 Then
      Set Col = GridEX1.Columns.Add '11
      Col.Width = 1700
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ยอดเงิน")
   End If
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   If DocumentType = 1 Then
      Me.Caption = MapText("ใบส่งสินค้า")
   ElseIf DocumentType = 2 Then
      Me.Caption = MapText("ใบเสร็จรับเงิน")
   ElseIf DocumentType = 3 Then
      Me.Caption = MapText("ใบเพิ่มหนี้")
   ElseIf DocumentType = 4 Then
      Me.Caption = MapText("ใบลดหนี้")
   End If
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblCustomerCode, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   Call InitNormalLabel(lblBatch, MapText("แบต"))
   Call InitNormalLabel(lblPackageType, MapText("ประเภทราคา"))
   
   Call InitCheckBox(chkCommit, "คำนวณ")
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   Call InitCombo(cboBatch)
   Call InitCombo(CboPackageType)
   
   cboBatch.Enabled = (glbUser.SIMULATE_FLAG = "Y")
   
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
   cmdProcess.Picture = LoadPicture(glbParameterObj.NormalButton1)
'   cmdDelete.Enabled = False
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdProcess, MapText("ประมวลผล"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_BillingDoc = New CBillingDoc
   Set m_TempBillingDoc = New CBillingDoc
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
   RowBuffer.RowStyle = RowBuffer.Value(6)
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
   Call m_TempBillingDoc.PopulateFromRS(1, m_Rs)
   Values(1) = m_TempBillingDoc.BILLING_DOC_ID
   Values(2) = m_TempBillingDoc.DOCUMENT_NO
   Values(3) = DateToStringExt(m_TempBillingDoc.DOCUMENT_DATE)
   Values(4) = m_TempBillingDoc.CUSTOMER_CODE
   Values(5) = m_TempBillingDoc.CUSTOMER_NAME
   Values(6) = m_TempBillingDoc.COMMIT_FLAG
   Values(7) = m_TempBillingDoc.RECEIPT_TYPE
   Values(8) = m_TempBillingDoc.DOCUMENT_SUBTYPE
   Values(9) = m_TempBillingDoc.BATCH_ID
   Values(10) = m_TempBillingDoc.DOCUMENT_DATE
   If DocumentType = 1 Then
      Values(11) = FormatNumberToNull(m_TempBillingDoc.BILL_TOTAL_AMOUNT)
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Function GetTotalPrice() As Double
Dim II As CDoItem
'Dim Sum1 As Double
Dim Sum2 As Double
'Dim Sum3 As Double

'   Sum1 = 0
   Sum2 = 0
'   Sum3 = 0
   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
'         Sum1 = Sum1 + II.ITEM_AMOUNT
         Sum2 = Sum2 + II.TOTAL_PRICE
'         Sum3 = Sum3 + II.TOTAL_WEIGHT
      End If
   Next II

'   For Each II In m_BillingDoc.Revenues
'      If II.Flag <> "D" Then
'         Sum2 = Sum2 + II.TOTAL_PRICE
'      End If
'   Next II
   
'   txtTotalDiscount.Text = Format(Sum3, "0.00")
'   txtTotalAmount.Text = Format(Sum1, "0.00")
   GetTotalPrice = Format(Sum2, "0.00")
End Function
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
