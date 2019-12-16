VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddReceiptItem 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddReceiptItem.frx":0000
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
      TabIndex        =   8
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   1
         Top             =   1290
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   5865
         Left            =   6540
         TabIndex        =   5
         Top             =   1890
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   10345
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
         Column(1)       =   "frmAddReceiptItem.frx":27A2
         Column(2)       =   "frmAddReceiptItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddReceiptItem.frx":290E
         FormatStyle(2)  =   "frmAddReceiptItem.frx":2A6A
         FormatStyle(3)  =   "frmAddReceiptItem.frx":2B1A
         FormatStyle(4)  =   "frmAddReceiptItem.frx":2BCE
         FormatStyle(5)  =   "frmAddReceiptItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddReceiptItem.frx":2D5E
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5865
         Left            =   120
         TabIndex        =   14
         Top             =   1890
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   10345
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MultiSelect     =   -1  'True
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
         Column(1)       =   "frmAddReceiptItem.frx":2F36
         Column(2)       =   "frmAddReceiptItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddReceiptItem.frx":30A2
         FormatStyle(2)  =   "frmAddReceiptItem.frx":31FE
         FormatStyle(3)  =   "frmAddReceiptItem.frx":32AE
         FormatStyle(4)  =   "frmAddReceiptItem.frx":3362
         FormatStyle(5)  =   "frmAddReceiptItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddReceiptItem.frx":34F2
      End
      Begin Threed.SSCommand cmdSelectEx 
         Height          =   525
         Left            =   5640
         TabIndex        =   13
         Top             =   3090
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   5648
         TabIndex        =   4
         Top             =   5040
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   5648
         TabIndex        =   3
         Top             =   4470
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":3CFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10140
         TabIndex        =   2
         Top             =   870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":4018
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   12
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   11
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   10
         Top             =   1320
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4320
         TabIndex        =   6
         Top             =   7860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":4332
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5970
         TabIndex        =   7
         Top             =   7860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddReceiptItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Private m_TempCol2 As Collection

Public DocumentDate As Date
Public AccountID As Long
Public Area As Long
Public ReceiptType As Long
Public CnDnItems As Collection

Private Sub PopulateDestColl()
Dim Ri As CReceiptItem
Dim D As CBillingDoc

   For Each Ri In TempCollection
      Set D = New CBillingDoc
      
      If Ri.Flag <> "D" Then
         D.BILLING_DOC_ID = Ri.DO_ID
         D.DOCUMENT_DATE = Ri.DOCUMENT_DATE
         D.DOCUMENT_NO = Ri.DOCUMENT_NO
         Call m_TempCol2.Add(D)
      End If
      
      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CBillingDoc
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.BILLING_DOC_ID = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim Bd As CBillingDoc

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set Bd = New CBillingDoc
      Call Bd.PopulateFromRS(1, Rs)
      
'      ''debug.print Bd.DOCUMENT_NO
      '''debug.print Bd.DOCUMENT_DATE
'      If Bd.DOCUMENT_NO = "S5111/0109" Then
'         ''debug.print
'      End If
      If Round(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT) - Bd.PAID_AMOUNT, 2) > 0 Then   'ติดค้างมากกว่ายอดหนี้
         If Not IsIn(m_TempCol2, Bd.BILLING_DOC_ID) Then
            'Call TempCol.Add(Bd)
            Call TempCol.Add(Bd, Trim(Bd.DOCUMENT_NO))
         End If
      End If
      
      Set Bd = Nothing
      Rs.MoveNext
   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingDoc.COMMIT_FLAG = ""
      m_BillingDoc.FROM_DATE = uctlFromDate.ShowDate
      m_BillingDoc.TO_DATE = uctlDocumentDate.ShowDate
      m_BillingDoc.VALID_DATE = uctlDocumentDate.ShowDate
      m_BillingDoc.ACCOUNT_ID = AccountID
      If ReceiptType = 3 Then
         m_BillingDoc.DOCUMENT_TYPE = 1
      ElseIf ReceiptType = 5 Then
         m_BillingDoc.DOCUMENT_TYPE = 3
      End If
      m_BillingDoc.ItemSumFlag = True
      
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.itemcount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.itemcount = 0
      GridEX1.Rebind
   End If
   
   GridEX2.itemcount = TempCollection.Count
   GridEX2.Rebind
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

'   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("INVENTORY_EXPORT_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
'   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("INVENTORY_EXPORT_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
'   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call PopulateTempColl

   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkSaleFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkSaleFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub


Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

'Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
'   If ID > 0 Then
'      TempCol1(ID).Flag = "A"
'      TempCol1(ID).PAID_TYPE = 1
'      Call TempCol2.Add(TempCol1(ID))
'      TempCol1.Remove (ID)
'   End If
'End Sub
Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim Bl As CBillingDoc
   
   If ID > 0 Then
      Set Bl = TempCol1(ID)
      Bl.Flag = "A"
      Bl.PAID_TYPE = 1
      Call TempCol2.Add(Bl)
      TempCol1.Remove (ID)
   End If
End Sub
Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long

   For j = 1 To TempCol1.Count
      TempCol1(j).Flag = "A"
      TempCol1(j).PAID_TYPE = 1
      Call TempCol2.Add(TempCol1(j))
   Next j
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

'Private Sub cmdSelect_Click()
'Dim TempID As Long
'
'   m_HasModify = True
'
'   TempID = GridEX1.Row
'   Call CopyItem(m_TempCol1, m_TempCol2, TempID)
'
'   GridEX1.itemcount = m_TempCol1.Count
'   GridEX1.Rebind
'
'   GridEX2.itemcount = m_TempCol2.Count
'   GridEX2.Rebind
'End Sub
Private Sub cmdSelect_Click()
Dim TempID As Long
Dim check As CBillingDoc
Dim ID As Long
Dim I As Long
Dim Row As Long
Dim Total(100)  As Long
Dim RowSelected As Long
Dim Count As Long

   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   
   m_HasModify = True
   I = 0
   For Row = 1 To GridEX1.RowCount
      If GridEX1.RowSelected(Row) = True Then
         I = I + 1
         RowSelected = RowSelected + 1
         Total(I) = GridEX1.GetRowData(Row).Value(1)
      End If
   Next Row
   
   For Count = 1 To RowSelected
      I = 0
      For Each check In m_TempCol1
         I = I + 1
         If check.BILLING_DOC_ID = Total(Count) Then
            TempID = I
            Exit For
         End If
      Next check
      Call CopyItem(m_TempCol1, m_TempCol2, TempID)
   Next Count
   
'   For Row = 1 To GridEX1.RowCount
'      I = 0
'      If GridEX1.RowSelected(Row) = True Then
'         ID = GridEX1.GetRowData(Row).Value(1)
'         For Each check In m_TempCol1
'            I = I + 1
'            If check.BILLING_DOC_ID = ID Then
'               TempID = I
'               Exit For
'            End If
'         Next check
'         Call CopyItem(m_TempCol1, m_TempCol2, TempID)
'      End If
'   Next Row

   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
      
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)
   
   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Private Sub GenerateCnDnItem(DoID As Long)
Dim Ri As CReceiptItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim CnDn As CReceiptCnDn

   Set TempRs = New ADODB.Recordset
   Set Ri = New CReceiptItem
   
   Ri.RECEIPT_ITEM_ID = -1
   Ri.DO_ID = DoID
   Call Ri.QueryData(10, TempRs, iCount)
   While Not TempRs.EOF
      Call Ri.PopulateFromRS(10, TempRs)
      
      Set CnDn = New CReceiptCnDn
      CnDn.Flag = "A"
      CnDn.DO_ID = Ri.DO_ID
      CnDn.DO_NO = Ri.DO_NO
      CnDn.CNDN_ID = Ri.CNDN_ID
      CnDn.CNDN_NO = Ri.CNDN_NO
      CnDn.CNDN_DATE = Ri.CNDN_DATE
      If Ri.CNDN_TYPE = 3 Then 'เพิ่มหนี้
         CnDn.DN_AMOUNT = Ri.DEBIT_CREDIT_AMOUNT
         CnDn.CN_AMOUNT = 0
      Else 'ลดหนี้ = 4
         CnDn.DN_AMOUNT = 0
         CnDn.CN_AMOUNT = Ri.DEBIT_CREDIT_AMOUNT
      End If
      Call CnDnItems.Add(CnDn)
      Set CnDn = Nothing
      
      TempRs.MoveNext
   Wend
   
   Set Ri = Nothing
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
End Sub

Public Sub PopulateTempColl()
Dim D As CBillingDoc
Dim Ri As CReceiptItem

   For Each D In m_TempCol2
      Set Ri = New CReceiptItem
      
      If D.Flag = "A" Then
         Ri.Flag = "A"
         Ri.DO_ID = D.BILLING_DOC_ID
         Ri.DOCUMENT_DATE = D.DOCUMENT_DATE
         Ri.DOCUMENT_NO = D.DOCUMENT_NO
         Ri.RECEIPT_ITEM_AMOUNT = D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT) - D.DISCOUNT_AMOUNT
         If D.PAID_TYPE = 1 Then 'ชำระเต็มวงเงิน
            Ri.PAID_AMOUNT = (D.DO_TOTAL_PRICE + D.REVENUE_TOTAL_PRICE - D.DISCOUNT_AMOUNT + (D.DEBIT_AMOUNT - D.CREDIT_AMOUNT) - D.PAID_AMOUNT)
         Else
            Ri.PAID_AMOUNT = D.TEMP_PAID_AMOUNT
         End If
         
         Call GenerateCnDnItem(Ri.DO_ID)
         Call TempCollection.Add(Ri)
      End If
      
      Set Ri = Nothing
   Next D
End Sub

Public Sub CopyItemEx(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CBillingDoc

   If ID > 0 Then
      Set L = TempCol1(ID)
      
      Set frmAddEditReceiptItemAmount.BillingDoc = L
      frmAddEditReceiptItemAmount.ShowMode = SHOW_EDIT
      Load frmAddEditReceiptItemAmount
      frmAddEditReceiptItemAmount.Show 1
      
      OKClick = frmAddEditReceiptItemAmount.OKClick
      
      Unload frmAddEditReceiptItemAmount
      Set frmAddEditReceiptItemAmount = Nothing
            
      If OKClick Then
         L.Flag = "A"
         Call TempCol2.Add(L)
         TempCol1.Remove (ID)
      End If
'      TempCol1(ID).Flag = "A"
'      Call TempCol2.Add(TempCol1(ID))
'      TempCol1.Remove (ID)
   End If
End Sub

Private Sub cmdSelectEx_Click()
Dim TempID As Long

   m_HasModify = True
   
   TempID = GridEX1.Row
   Call CopyItemEx(m_TempCol1, m_TempCol2, TempID)

   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call EnableForm(Me, False)
      Call PopulateDestColl
      uctlDocumentDate.ShowDate = DocumentDate
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_BillingDoc.QueryFlag = 0
         Call QueryData(True)
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
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_Employees = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 2625
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2520
   Col.Caption = MapText("หมายเลขเอกสาร")
End Sub


Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.itemcount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.Name = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX2.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX2.Columns.Add '3
   Col.Width = 2625
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX2.Columns.Add '4
   Col.Width = 2520
   Col.Caption = MapText("หมายเลขเอกสาร")
End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_BillingDoc.ImportExports
'      If II.Flag <> "D" Then
'         Sum = Sum + CDbl(Format(II.EXPORT_AVG_PRICE, "0.00")) * CDbl(Format(II.EXPORT_AMOUNT, "0.00"))
'      End If
'   Next II
''
''   txtDeliveryFee.Text = Format(Sum, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectEx.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา"))
   Call InitMainButton(cmdSelect, MapText(">"))
   Call InitMainButton(cmdSelectAll, MapText(">>"))
   Call InitMainButton(cmdSelectEx, MapText("*>"))
   
   Call InitGrid1
   Call InitGrid2
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
   Set m_BillingDoc = New CBillingDoc
   Set m_Employees = New Collection
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CBillingDoc
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.BILLING_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CBillingDoc
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.BILLING_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
