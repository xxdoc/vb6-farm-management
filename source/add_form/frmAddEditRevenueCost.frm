VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditRevenueCost 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditRevenueCost.frx":0000
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
      Begin prjFarmManagement.uctlDate uctlRevenueDate 
         Height          =   405
         Left            =   6300
         TabIndex        =   1
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   2
         Top             =   2040
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
      Begin prjFarmManagement.uctlTextBox txtRevenueNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   900
         Width           =   2385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5175
         Left            =   150
         TabIndex        =   3
         Top             =   2550
         Width           =   11595
         _ExtentX        =   20452
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
         Column(1)       =   "frmAddEditRevenueCost.frx":27A2
         Column(2)       =   "frmAddEditRevenueCost.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditRevenueCost.frx":290E
         FormatStyle(2)  =   "frmAddEditRevenueCost.frx":2A6A
         FormatStyle(3)  =   "frmAddEditRevenueCost.frx":2B1A
         FormatStyle(4)  =   "frmAddEditRevenueCost.frx":2BCE
         FormatStyle(5)  =   "frmAddEditRevenueCost.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditRevenueCost.frx":2D5E
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
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtRevenueDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   13
         Top             =   1410
         Width           =   9825
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin VB.Label lblRevenueDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   14
         Top             =   1470
         Width           =   1395
      End
      Begin VB.Label lblRevenueDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   12
         Top             =   960
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditRevenueCost.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   8
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditRevenueCost.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditRevenueCost.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblRevenueNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditRevenueCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_RevenueCost As CRevenueCost

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_RevenueCost.SetFieldValue("REVENUE_COST_ID", ID)
      m_RevenueCost.QueryFlag = 1
      If Not glbDaily.QueryRevenueCost(m_RevenueCost, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_RevenueCost.PopulateFromRS(1, m_Rs)
      
      txtRevenueNo.Text = m_RevenueCost.GetFieldValue("REVENUE_NO")
      uctlRevenueDate.ShowDate = m_RevenueCost.GetFieldValue("REVENUE_DATE")
      txtRevenueDesc.Text = m_RevenueCost.GetFieldValue("REVENUE_DESC")
      
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("LEDGER_COST_REVENUE_EDIT") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
   End If

   If Not VerifyTextControl(lblRevenueNo, txtRevenueNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblRevenueDate, uctlRevenueDate, False) Then
      Exit Function
   End If
   If Not VerifyDateInterval(uctlRevenueDate.ShowDate) Then
      Exit Function
   End If
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_RevenueCost.ShowMode = ShowMode
   Call m_RevenueCost.SetFieldValue("REVENUE_COST_ID", ID)
    Call m_RevenueCost.SetFieldValue("REVENUE_DATE", uctlRevenueDate.ShowDate)
   Call m_RevenueCost.SetFieldValue("REVENUE_NO", txtRevenueNo.Text)
   Call m_RevenueCost.SetFieldValue("REVENUE_DESC", txtRevenueDesc.Text)
   
   Call EnableForm(Me, False)
   
   
   If Not glbDaily.AddEditRevenueCost(m_RevenueCost, IsOK, True, glbErrorLog) Then
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
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim oMenu As cPopupMenu
Dim lMenuChoosen As Long

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   
   Set frmAddEditRevenueCostItem.ParentForm = Me
   frmAddEditRevenueCostItem.HeaderText = "เพิ่มรายการ"
   frmAddEditRevenueCostItem.ShowMode = SHOW_ADD
   Set frmAddEditRevenueCostItem.TempCollection = m_RevenueCost.RevenueTypeItems
   Load frmAddEditRevenueCostItem
   frmAddEditRevenueCostItem.Show 1

   OKClick = frmAddEditRevenueCostItem.OKClick

   Unload frmAddEditRevenueCostItem
   Set frmAddEditRevenueCostItem = Nothing

   If OKClick Then
      m_HasModify = True
      GridEX1.ItemCount = CountItem(m_RevenueCost.RevenueTypeItems)
      GridEX1.Rebind
   End If
   
End Sub
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If ID1 <= 0 Then
      m_RevenueCost.RevenueTypeItems.Remove (ID2)
   Else
      m_RevenueCost.RevenueTypeItems.Item(ID2).Flag = "D"
   End If
   Call RefreshGrid
   
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim PaymentType As Long
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   Set frmAddEditRevenueCostItem.ParentForm = Me
   frmAddEditRevenueCostItem.ID = ID
   frmAddEditRevenueCostItem.HeaderText = "แก้ไขรายการ"
   frmAddEditRevenueCostItem.ShowMode = SHOW_EDIT
   Set frmAddEditRevenueCostItem.TempCollection = m_RevenueCost.RevenueTypeItems
   Load frmAddEditRevenueCostItem
   frmAddEditRevenueCostItem.Show 1

   OKClick = frmAddEditRevenueCostItem.OKClick

   Unload frmAddEditRevenueCostItem
   Set frmAddEditRevenueCostItem = Nothing

   If OKClick Then
      m_HasModify = True
      GridEX1.ItemCount = CountItem(m_RevenueCost.RevenueTypeItems)
      GridEX1.Rebind
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
      ID = m_RevenueCost.GetFieldValue("REVENUE_COST_ID")
      m_RevenueCost.QueryFlag = 1
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
         m_RevenueCost.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlRevenueDate.ShowDate = Now
         m_RevenueCost.QueryFlag = 0
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
   
   Set m_RevenueCost = Nothing
   
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
   Col.Width = 3000
   Col.Caption = MapText("ประเภทรายได้")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 3000
   Col.Caption = MapText("สถานะสุกร")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 3000
   Col.Caption = MapText("ประเภทสุกร")
   
      
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน")
   
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblRevenueNo, MapText("เลขที่"))
   Call InitNormalLabel(lblRevenueDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblRevenueDate, MapText("วันที่"))
   
   
   Call txtRevenueNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Dim T As Object
   TabStrip1.Tabs.Clear
   
   Set T = TabStrip1.Tabs.Add()
   T.Caption = MapText("รายละเอียด")
   T.Tag = "DESC"
   
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
   Set m_RevenueCost = New CRevenueCost
   
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim Pos As CRevenueCostItem

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_RevenueCost.RevenueTypeItems Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   If m_RevenueCost.RevenueTypeItems.Count <= 0 Then
      Exit Sub
   End If
   Set Pos = GetItem(m_RevenueCost.RevenueTypeItems, RowIndex, RealIndex)
   If Pos Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = Pos.GetFieldValue("REVENUE_COST_ITEM_ID")
   Values(2) = RealIndex
   Values(3) = Pos.GetFieldValue("REVENUE_TYPE_NAME")
   Values(4) = Pos.GetFieldValue("PIG_STATUS_NAME")
   Values(5) = Pos.GetFieldValue("PIG_TYPE_NAME") & "(" & Pos.GetFieldValue("PIG_TYPE_NO") & ")"
   Values(6) = FormatNumber(Pos.GetFieldValue("REVENUE_COST_ITEM_AMOUNT"))
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub RefreshGrid()
   If TabStrip1.SelectedItem.Tag = "DESC" Then
      GridEX1.ItemCount = CountItem(m_RevenueCost.RevenueTypeItems)
      GridEX1.Rebind
   End If
   
   m_HasModify = True
   
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   Call InitGrid1
   Call RefreshGrid
End Sub

Private Sub txtRevenueDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtRevenueNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlRevenueDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlRevenueDate_LostFocus()
   If ShowMode = SHOW_ADD And uctlRevenueDate.ShowDate > 0 Then
      If Not VerifyDateInterval(uctlRevenueDate.ShowDate) Then
         uctlRevenueDate.SetFocus
         Exit Sub
      End If
   End If
End Sub
