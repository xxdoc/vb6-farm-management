VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddBatchItem 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddBatchItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtSource 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
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
         TabIndex        =   11
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         Enabled         =   0   'False
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlTravelDate 
         Height          =   405
         Left            =   5880
         TabIndex        =   1
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2385
         Left            =   180
         TabIndex        =   4
         Top             =   2040
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4207
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
         Column(1)       =   "frmAddBatchItem.frx":000C
         Column(2)       =   "frmAddBatchItem.frx":00D4
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddBatchItem.frx":0178
         FormatStyle(2)  =   "frmAddBatchItem.frx":02D4
         FormatStyle(3)  =   "frmAddBatchItem.frx":0384
         FormatStyle(4)  =   "frmAddBatchItem.frx":0438
         FormatStyle(5)  =   "frmAddBatchItem.frx":0510
         ImageCount      =   0
         PrinterProperties=   "frmAddBatchItem.frx":05C8
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   2385
         Left            =   180
         TabIndex        =   7
         Top             =   5310
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4207
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
         Column(1)       =   "frmAddBatchItem.frx":07A0
         Column(2)       =   "frmAddBatchItem.frx":0868
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddBatchItem.frx":090C
         FormatStyle(2)  =   "frmAddBatchItem.frx":0A68
         FormatStyle(3)  =   "frmAddBatchItem.frx":0B18
         FormatStyle(4)  =   "frmAddBatchItem.frx":0BCC
         FormatStyle(5)  =   "frmAddBatchItem.frx":0CA4
         ImageCount      =   0
         PrinterProperties=   "frmAddBatchItem.frx":0D5C
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblTravelDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   15
         Top             =   1080
         Width           =   1485
      End
      Begin Threed.SSCheck chkReference 
         Height          =   405
         Left            =   1590
         TabIndex        =   2
         Top             =   1440
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   6000
         TabIndex        =   6
         Top             =   4620
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   5280
         TabIndex        =   5
         Top             =   4620
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10140
         TabIndex        =   3
         Top             =   870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10920
         TabIndex        =   13
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4320
         TabIndex        =   8
         Top             =   7860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5970
         TabIndex        =   9
         Top             =   7860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddBatchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Parameters As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Private m_TempCol2 As Collection
Private m_Parameter As CParameter
Private m_Zones As Collection

Public AccountID As Long
Public ReceiptType As Long
Public InvoiceDOType As Long
Public Area As Long
Public TravelDate As Double
Public ReceiveNo As String
Public ZoneID As Long

Private Sub PopulateDestColl()
Dim Ri As CBatchItem
Dim D As CParameter

   For Each Ri In TempCollection
      Set D = New CParameter

      If Ri.Flag <> "D" Then
         Call D.SetFieldValue("PARAM_ID", Ri.GetFieldValue("PARAM_ID"))
         Call D.SetFieldValue("PARAM_NO", Ri.GetFieldValue("PARAM_NO"))
         Call D.SetFieldValue("PARAM_DESC", Ri.GetFieldValue("PARAM_DESC"))
         Call D.SetFieldValue("PARAM_DATE", Ri.GetFieldValue("PARAM_DATE"))
         
         Call m_TempCol2.Add(D)
      End If

      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CParameter
Dim Found As Boolean
Dim RefCount As Long
   
   Found = False
   For Each D In TempCol
      If D.GetFieldValue("PARAM_ID") = TempID Then
         Found = True
         Exit For
      End If
   Next D

   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim Bd As CParameter
Dim RefCount As Long
Dim Found As Boolean

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set Bd = New CParameter
      Call Bd.PopulateFromRS(1, Rs)

      If Not IsIn(m_TempCol2, Bd.GetFieldValue("PARAM_ID")) Then
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

      Call m_Parameter.SetFieldValue("PARAM_ID", -1)
      Call m_Parameter.SetFieldValue("PARAM_AREA", Area)
      Call m_Parameter.SetFieldValue("PARAM_NO", txtSource.Text)
      Call m_Parameter.SetFieldValue("FROM_DATE", uctlTravelDate.ShowDate)
      Call m_Parameter.SetFieldValue("TO_DATE", uctlTravelDate.ShowDate)
      If Not glbDaily.QueryParameter(m_Parameter, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
Dim L As CParameter

   If ID > 0 Then
      TempCol1(ID).Flag = "A"
      Call TempCol2.Add(TempCol1(ID))
      TempCol1.Remove (ID)
   End If
End Sub

Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long

   For j = 1 To TempCol1.Count
      TempCol1(j).Flag = "A"
      Call TempCol2.Add(TempCol1(j))
   Next j
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

Private Sub cmdSelect_Click()
Dim TempID As Long
Dim check As CParameter
Dim ID As Long
Dim I As Long
Dim Row As Long

   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If

   m_HasModify = True
   'Id = GridEX1.Value(1)
   For Row = 1 To GridEX1.RowCount
      I = 0
      If GridEX1.RowSelected(Row) = True Then
         ID = GridEX1.GetRowData(Row).Value(1)
         For Each check In m_TempCol1
            I = I + 1
            If check.GetFieldValue("PARAM_ID") = ID Then
               TempID = I
            End If
         Next check
         Call CopyItem(m_TempCol1, m_TempCol2, TempID)
       End If
       
   Next Row
        
   GridEX1.ItemCount = CountItem(m_TempCol1)
   GridEX1.Rebind
   
   GridEX2.ItemCount = CountItem(m_TempCol2)
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
Dim D As CParameter
Dim Ri As CBatchItem

   For Each D In m_TempCol2
      Set Ri = New CBatchItem

      If D.Flag = "A" Then
         Ri.Flag = "A"
         Call Ri.SetFieldValue("PARAM_ID", D.GetFieldValue("PARAM_ID"))
         Call Ri.SetFieldValue("PARAM_NO", D.GetFieldValue("PARAM_NO"))
         Call Ri.SetFieldValue("PARAM_DATE", D.GetFieldValue("PARAM_DATE"))
         Call Ri.SetFieldValue("PARAM_DESC", D.GetFieldValue("PARAM_DESC"))
         
         Call TempCollection.Add(Ri)
      End If

      Set Ri = Nothing
   Next D
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call PopulateDestColl
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Parameter.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Parameter.QueryFlag = 0
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

   Set m_Parameters = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   Set m_Parameter = Nothing
   Set m_Zones = Nothing
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
   
   '==
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2820
   Col.Caption = MapText("�Ţ���")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2910
   Col.Caption = MapText("�ѹ���")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 5760
   Col.Caption = MapText("��������´")
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

   '==
   Set Col = GridEX2.Columns.Add '3
   Col.Width = 2820
   Col.Caption = MapText("�Ţ���")
   
   Set Col = GridEX2.Columns.Add '4
   Col.Width = 2910
   Col.Caption = MapText("�ѹ���")

   Set Col = GridEX2.Columns.Add '5
   Col.Width = 5760
   Col.Caption = MapText("��������´")
End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_Parameter.ImportExports
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
   
   Call InitNormalLabel(lblTravelDate, MapText("�ѹ���"))
    Call InitNormalLabel(lblSource, MapText("�Ţ���"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
      
   uctlTravelDate.ShowDate = TravelDate
   txtSource.Text = ReceiveNo
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdSearch, MapText("���� (F5)"))
   Call InitMainButton(cmdSelect, MapText("V"))
   Call InitMainButton(cmdSelectAll, MapText("VVV"))
   
   Call InitCheckBox(chkReference, MapText("�������"))
   chkReference.Enabled = False
   
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
   Set m_Parameters = New Collection
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   Set m_Parameter = New CParameter
   Set m_Zones = New Collection
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

   Dim CR As CParameter
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

      Values(1) = CR.GetFieldValue("PARAM_ID")
      Values(2) = RealIndex
      Values(3) = CR.GetFieldValue("PARAM_NO")
      Values(4) = DateToStringExtEx2(CR.GetFieldValue("PARAM_DATE"))
      Values(5) = CR.GetFieldValue("PARAM_DESC")
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

   Dim CR As CParameter
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.GetFieldValue("PARAM_ID")
   Values(2) = RealIndex
   Values(3) = CR.GetFieldValue("PARAM_NO")
   Values(4) = DateToStringExtEx2(CR.GetFieldValue("PARAM_DATE"))
   Values(5) = CR.GetFieldValue("PARAM_DESC")
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

