VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddChequeItem 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddChequeItem.frx":0000
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
      TabIndex        =   9
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
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5865
         Left            =   150
         TabIndex        =   3
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
         Column(1)       =   "frmAddChequeItem.frx":27A2
         Column(2)       =   "frmAddChequeItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddChequeItem.frx":290E
         FormatStyle(2)  =   "frmAddChequeItem.frx":2A6A
         FormatStyle(3)  =   "frmAddChequeItem.frx":2B1A
         FormatStyle(4)  =   "frmAddChequeItem.frx":2BCE
         FormatStyle(5)  =   "frmAddChequeItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddChequeItem.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
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
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   5865
         Left            =   6540
         TabIndex        =   6
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
         Column(1)       =   "frmAddChequeItem.frx":2F36
         Column(2)       =   "frmAddChequeItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddChequeItem.frx":30A2
         FormatStyle(2)  =   "frmAddChequeItem.frx":31FE
         FormatStyle(3)  =   "frmAddChequeItem.frx":32AE
         FormatStyle(4)  =   "frmAddChequeItem.frx":3362
         FormatStyle(5)  =   "frmAddChequeItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddChequeItem.frx":34F2
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   5648
         TabIndex        =   5
         Top             =   5040
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   5648
         TabIndex        =   4
         Top             =   4470
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeItem.frx":39E4
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
         MouseIcon       =   "frmAddChequeItem.frx":3CFE
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   13
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   12
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   11
         Top             =   1320
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4320
         TabIndex        =   7
         Top             =   7860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddChequeItem.frx":4018
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5970
         TabIndex        =   8
         Top             =   7860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddChequeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Cheque As CCheque
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection
Public DocumentType As CASH_DOC_TYPE

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Private m_TempCol2 As Collection
Private m_RcpCndnItems As Collection

Public ApArID As Long
Public ReceiptType As Long
Public Area As Long
Public AccountID  As Long

Private Sub PopulateDestColl()
Dim Ri As CCashTransferItem
Dim D As CCheque

   For Each Ri In TempCollection
      Set D = New CCheque
      
      If (Ri.Flag <> "D") And (Ri.ExportItem.GetFieldValue("CHECK_ID") > 0) Then
         Call D.SetFieldValue("CHEQUE_ID", Ri.ExportItem.GetFieldValue("CHECK_ID"))
         Call D.SetFieldValue("CHEQUE_DATE", Ri.ExportItem.Cheque.GetFieldValue("CHEQUE_DATE"))
         Call D.SetFieldValue("CHEQUE_NO", Ri.ExportItem.Cheque.GetFieldValue("CHEQUE_NO"))
         Call D.SetFieldValue("CHEQUE_AMOUNT", Ri.ExportItem.Cheque.GetFieldValue("CHEQUE_AMOUNT"))
         Call m_TempCol2.Add(D)
      End If

      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CCheque
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.GetFieldValue("CHEQUE_ID") = TempID Then
         Found = True
      End If
   Next D

   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
On Error Resume Next
Dim Bd As CCheque
Dim X As Double

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set Bd = New CCheque
      Call Bd.PopulateFromRS(1, Rs)
      
      If Not IsIn(m_TempCol2, Bd.GetFieldValue("CHEQUE_ID")) Then
         Call TempCol.Add(Bd)
      End If
      
      Set Bd = Nothing
      Rs.MoveNext
   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      Call m_Cheque.SetFieldValue("FROM_DATE", uctlFromDate.ShowDate)
      Call m_Cheque.SetFieldValue("TO_DATE", uctlDocumentDate.ShowDate)
      Call m_Cheque.SetFieldValue("DIRECTION", 1)
      Call m_Cheque.SetFieldValue("BANK_FLAG", "N")
      Call m_Cheque.SetFieldValue("APAR_ID", ApArID)
      If Not glbDaily.QueryCheque(m_Cheque, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If ItemCount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
   End If

   GridEX2.ItemCount = m_TempCol2.Count
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

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CCheque

   If ID > 0 Then
      Set L = TempCol1(ID)

      frmAddEditChequeItemAmount.HeaderText = "��¡����"
      Set frmAddEditChequeItemAmount.BillingDoc = L
      frmAddEditChequeItemAmount.ShowMode = SHOW_EDIT
      Load frmAddEditChequeItemAmount
      frmAddEditChequeItemAmount.Show 1

      OKClick = frmAddEditChequeItemAmount.OKClick

      Unload frmAddEditChequeItemAmount
      Set frmAddEditChequeItemAmount = Nothing
   End If
      
   If OKClick Then
      L.Flag = "A"
      Call TempCol2.Add(L)
      TempCol1.Remove (ID)
   End If
End Sub

Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long
Dim D As CCheque

   For j = 1 To TempCol1.Count
      TempCol1(j).Flag = "A"
'      TempCol1(j).PAID_TYPE = 1
      Set D = TempCol1(j)
      Call TempCol1(j).SetFieldValue("CHEQUE_AMOUNT", D.GetFieldValue("CHEQUE_AMOUNT"))
      Call TempCol2.Add(TempCol1(j))
   Next j

   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

Private Sub cmdSelect_Click()
Dim TempID As Long

   m_HasModify = True
   
   TempID = GridEX1.Row
   Call CopyItem(m_TempCol1, m_TempCol2, TempID)

   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)
   
   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Public Sub PopulateTempColl()
Dim D As CCheque
Dim EnpAddress As CCashTransferItem
Dim EI As CCashTran
Dim II As CCashTran
Dim PaymentType As Long

   For Each D In m_TempCol2
      If D.Flag = "A" Then
         Set EI = New CCashTran
         Set II = New CCashTran
         Set EnpAddress = New CCashTransferItem
   
         EI.Flag = "A"
         II.Flag = "A"
         EnpAddress.Flag = "A"

         Set EnpAddress.ExportItem = EI
         Set EnpAddress.ImportItem = II
   
         Call TempCollection.Add(EnpAddress)
   
         '�ӽҡ�Թʴ����
         Call EnpAddress.ExportItem.SetFieldValue("PAYMENT_TYPE", 3) '�͡���Թʴ
         Call EnpAddress.ExportItem.SetFieldValue("PAYMENT_TYPE_NAME", PaymentType2Text(3))
         Call EnpAddress.ExportItem.SetFieldValue("AMOUNT", D.GetFieldValue("CHEQUE_AMOUNT"))
         Call EnpAddress.ExportItem.SetFieldValue("CHECK_ID", D.GetFieldValue("CHEQUE_ID"))
         Call EnpAddress.ExportItem.Cheque.SetFieldValue("CHEQUE_NO", D.GetFieldValue("CHEQUE_NO"))
         Call EnpAddress.ExportItem.Cheque.SetFieldValue("CHEQUE_DATE", D.GetFieldValue("CHEQUE_DATE"))
         Call EnpAddress.ExportItem.Cheque.SetFieldValue("EFFECTIVE_DATE", D.GetFieldValue("EFFECTIVE_DATE"))
         Call EnpAddress.ExportItem.SetFieldValue("NET_AMOUNT", D.GetFieldValue("CHEQUE_AMOUNT"))
         Call EnpAddress.ExportItem.SetFieldValue("TX_TYPE", "E")
         '�Թʴ����Ŵŧ
         Call EnpAddress.ExportItem.SetFieldValue("BANK_ID", -1)
         Call EnpAddress.ExportItem.SetFieldValue("BANK_BRANCH", -1)
         Call EnpAddress.ExportItem.SetFieldValue("BANK_ACCOUNT", -1)
         Call EnpAddress.ExportItem.SetFieldValue("BANK_NAME", "")
         Call EnpAddress.ExportItem.SetFieldValue("BRANCH_NAME", "")
         
         PaymentType = 3       '��� 1
         Call EnpAddress.ImportItem.SetFieldValue("PAYMENT_TYPE", PaymentType) '������Թʴ
         Call EnpAddress.ImportItem.SetFieldValue("PAYMENT_TYPE_NAME", PaymentType2Text(PaymentType))
         Call EnpAddress.ImportItem.SetFieldValue("AMOUNT", D.GetFieldValue("CHEQUE_AMOUNT"))
         Call EnpAddress.ImportItem.SetFieldValue("CHECK_ID", D.GetFieldValue("CHEQUE_ID"))
         Call EnpAddress.ImportItem.SetFieldValue("FEE_AMOUNT", D.GetFieldValue("TEMP_FEE_AMOUNT"))
         Call EnpAddress.ImportItem.SetFieldValue("NET_AMOUNT", D.GetFieldValue("CHEQUE_AMOUNT") - D.GetFieldValue("TEMP_FEE_AMOUNT"))
         Call EnpAddress.ImportItem.SetFieldValue("TX_TYPE", "I")
         
         If PaymentType = 1 Then
            Call EnpAddress.ImportItem.SetFieldValue("BANK_ID", -1)
            Call EnpAddress.ImportItem.SetFieldValue("BANK_BRANCH", -1)
            Call EnpAddress.ImportItem.SetFieldValue("BANK_ACCOUNT", -1)
            
            '���繤�����ǡѹ�Ѻ BANK_ID, BANK_BRANCH, BANK_ACCOUNT �ͧ CASH_DOC
            Call EnpAddress.ImportItem.SetFieldValue("BANK_ID", -1)
            Call EnpAddress.ImportItem.SetFieldValue("BANK_BRANCH", -1)
            Call EnpAddress.ImportItem.SetFieldValue("BANK_ACCOUNT", -1)
         End If
      End If
   Next D
End Sub

Private Sub Form_Activate()
Dim Ri As CReceiptItem

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call PopulateDestColl
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Cheque.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Cheque.QueryFlag = 0
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
   
   Set m_Cheque = Nothing
   Set m_Employees = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   Set m_RcpCndnItems = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
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
   Col.Width = 1650
   Col.Caption = MapText("�ѹ�����")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1830
   Col.Caption = MapText("�Ţ�����")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1665
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ӹǹ�Թ")
End Sub


Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.ItemCount = 0
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
   Col.Width = 1650
   Col.Caption = MapText("�ѹ�����")

   Set Col = GridEX2.Columns.Add '4
   Col.Width = 1830
   Col.Caption = MapText("�Ţ�����")
   
   Set Col = GridEX2.Columns.Add '5
   Col.Width = 1665
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("�ӹǹ�Թ")
End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_Cheque.ImportExports
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
   
   Call InitNormalLabel(lblDocumentDate, MapText("�֧�ѹ���"))
   Call InitNormalLabel(lblFromDate, MapText("�ҡ�ѹ���"))
   Call InitNormalLabel(Label4, MapText("�ҷ"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdSearch, MapText("����"))
   Call InitMainButton(cmdSelect, MapText(">"))
   Call InitMainButton(cmdSelectAll, MapText(">>"))
   
'   If (DocumentType = CN_DOCTYPE) Or (DocumentType = DN_DOCTYPE) Then
'      cmdSelectAll.Visible = False
'   End If
   
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
   Set m_Cheque = New CCheque
   Set m_Employees = New Collection
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   Set m_RcpCndnItems = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim X As Double

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CCheque
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.GetFieldValue("CHEQUE_ID")
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.GetFieldValue("CHEQUE_DATE"))
   Values(4) = CR.GetFieldValue("CHEQUE_NO")
   Values(5) = FormatNumber(CR.GetFieldValue("CHEQUE_AMOUNT"))
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

   Dim CR As CCheque
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.GetFieldValue("CHEQUE_ID")
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.GetFieldValue("CHEQUE_DATE"))
   Values(4) = CR.GetFieldValue("CHEQUE_NO")
   Values(5) = FormatNumber(CR.GetFieldValue("CHEQUE_AMOUNT"))
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
