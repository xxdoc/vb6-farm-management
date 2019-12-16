VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditYearSeq 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditYearSeq.frx":0000
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
      TabIndex        =   9
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   2
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
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1530
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtYear 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   2025
         _ExtentX        =   3572
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
         Height          =   4755
         Left            =   150
         TabIndex        =   3
         Top             =   2940
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8387
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
         Column(1)       =   "frmAddEditYearSeq.frx":27A2
         Column(2)       =   "frmAddEditYearSeq.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditYearSeq.frx":290E
         FormatStyle(2)  =   "frmAddEditYearSeq.frx":2A6A
         FormatStyle(3)  =   "frmAddEditYearSeq.frx":2B1A
         FormatStyle(4)  =   "frmAddEditYearSeq.frx":2BCE
         FormatStyle(5)  =   "frmAddEditYearSeq.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditYearSeq.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
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
         MouseIcon       =   "frmAddEditYearSeq.frx":2F36
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
         MouseIcon       =   "frmAddEditYearSeq.frx":3250
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
         MouseIcon       =   "frmAddEditYearSeq.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   1620
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditYearSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_YearSeq As CYearSeq

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private FileName As String
Private m_ProductTypes As Collection

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_YearSeq.YEAR_SEQ_ID = ID
      m_YearSeq.QueryFlag = 1
      If Not glbDaily.QueryYearSeq(m_YearSeq, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_YearSeq.PopulateFromRS(1, m_Rs)
      
      txtYear.Text = m_YearSeq.YEAR_NO
      txtDesc.Text = m_YearSeq.YEAR_DESC
   Else
      ShowMode = SHOW_ADD
   End If
   
   If ShowMode = SHOW_ADD Then
   End If
   
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

Private Function CheckIntregrity() As Boolean
   CheckIntregrity = True
End Function

Private Function GetPartItemFromType(Yw As CYearWeek, Pt As CProductType) As CPrtItemMap
Dim Pi As CPartItem
Dim Pm As CPrtItemMap
   
   Set GetPartItemFromType = Nothing
   For Each Pm In Yw.PartItemMaps
      Set Pi = Pm.PartItem
      If Pi.PIG_TYPE = Pt.PRODUCT_TYPE_NO Then
         Set GetPartItemFromType = Pm
         Exit For
      End If
   Next Pm
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("PIG_WEEK_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("PIG_WEEK_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblYear, txtYear, False) Then
      Exit Function
   End If

   If Not CheckUniqueNs(YEAR_NO, txtYear.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtYear.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If Not CheckIntregrity Then
      Exit Function
   End If
   
   Dim Pi As CPartItem
   Dim Yw As CYearWeek
   Dim Pt As CProductType
   Dim Pm As CPrtItemMap
   
   For Each Yw In m_YearSeq.YearWeeks
      For Each Pt In m_ProductTypes
         Set Pm = GetPartItemFromType(Yw, Pt)
         If (Pm Is Nothing) Then 'ไม่พบ Part type ใน Yw.PartItems
            Yw.Flag = "A"
            Set Pm = New CPrtItemMap
            Set Pi = Pm.PartItem
            Pi.AddEditMode = SHOW_ADD
            Pm.Flag = "A"
            Pm.AddEditMode = SHOW_ADD
            Pm.PRODUCT_TYPE_ID = Pt.PRODUCT_TYPE_ID
            Call Yw.PartItemMaps.Add(Pm)
            Set Pm = Nothing
         Else
            If Yw.Flag = "D" Then
               Pm.Flag = "D"
            Else
               Yw.Flag = "E"
               Pm.Flag = "E"
               Pm.AddEditMode = SHOW_EDIT
               Pm.PRODUCT_TYPE_ID = Pt.PRODUCT_TYPE_ID
               Set Pi = Pm.PartItem
               Pi.AddEditMode = SHOW_EDIT
            End If
         End If
         
         If Yw.WEEK_NO = "01" Or Yw.WEEK_NO = "00" Then
            'debug.print (Pi.PART_NO & "-" & Pi.PIG_TYPE & "-" & Pi.PART_ITEM_ID)
         End If
         Pi.PART_NO = Trim(Str(txtYear.Text)) & Trim(Format(Yw.WEEK_NO, "00"))
         Pi.PART_DESC = Pt.PRODUCT_TYPE_NAME & " " & Pi.PART_NO
         Pi.PART_TYPE = -1
         Pi.PIG_FLAG = "Y"
         Pi.UNIT_COUNT = -1
         Pi.PIG_TYPE = Pt.PRODUCT_TYPE_NO
         
         
      Next Pt
   Next Yw
   
   m_YearSeq.AddEditMode = ShowMode
   m_YearSeq.YEAR_NO = Val(txtYear.Text)
   m_YearSeq.YEAR_DESC = txtDesc.Text
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditYearSeq(m_YearSeq, IsOK, True, glbErrorLog) Then
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

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set oMenu = New cPopupMenu
      
      lMenuChosen = oMenu.Popup("เพิ่มทีละรายการ", "-", "เพิ่มอัตโนมัติ")
      If lMenuChosen = 0 Then
         Exit Sub
      ElseIf lMenuChosen = 1 Then
         Set frmAddEditYearWeek.TempCollection = m_YearSeq.YearWeeks
         frmAddEditYearWeek.ShowMode = SHOW_ADD
         frmAddEditYearWeek.HeaderText = MapText("เพิ่มสัปดาห์เกิด")
         Load frmAddEditYearWeek
         frmAddEditYearWeek.Show 1
   
         OKClick = frmAddEditYearWeek.OKClick
   
         Unload frmAddEditYearWeek
         Set frmAddEditYearWeek = Nothing
      ElseIf lMenuChosen = 3 Then
         Set frmAddYearWeek.TempCollection = m_YearSeq.YearWeeks
         frmAddYearWeek.ShowMode = SHOW_ADD
         frmAddYearWeek.HeaderText = MapText("เพิ่มสัปดาห์เกิด")
         Load frmAddYearWeek
         frmAddYearWeek.Show 1
   
         OKClick = frmAddYearWeek.OKClick
   
         Unload frmAddYearWeek
         Set frmAddYearWeek = Nothing
      End If
      
      Set oMenu = Nothing
      If OKClick Then
         GridEX1.ItemCount = CountItem(m_YearSeq.YearWeeks)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
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
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_YearSeq.YearWeeks.Remove (ID2)
      Else
         m_YearSeq.YearWeeks.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_YearSeq.YearWeeks)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
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

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditYearWeek.ID = ID
      Set frmAddEditYearWeek.TempCollection = m_YearSeq.YearWeeks
      frmAddEditYearWeek.HeaderText = MapText("แก้ไขสัปดาห์เกิด")
      frmAddEditYearWeek.ShowMode = SHOW_EDIT
      Load frmAddEditYearWeek
      frmAddEditYearWeek.Show 1

      OKClick = frmAddEditYearWeek.OKClick

      Unload frmAddEditYearWeek
      Set frmAddEditYearWeek = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_YearSeq.YearWeeks)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)

      Call LoadProductType(Nothing, m_ProductTypes)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_YearSeq.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_YearSeq.QueryFlag = 0
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
   
   Set m_YearSeq = Nothing
   Set m_ProductTypes = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
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
   Col.Width = 1005
   Col.Caption = MapText("สัปดาห์ที่")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2130
   Col.Caption = MapText("จาก")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1995
   Col.Caption = MapText("ถึง")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 6435
   Col.Caption = MapText("รายละเอียด")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblYear, MapText("ปี"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
      
   Call txtYear.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
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
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("ที่อยู่")
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
   Set m_YearSeq = New CYearSeq
   Set m_ProductTypes = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

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

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_YearSeq.YearWeeks Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CYearWeek
      If m_YearSeq.YearWeeks.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_YearSeq.YearWeeks, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.YEAR_WEEK_ID
      Values(2) = RealIndex
      Values(3) = CR.WEEK_NO
      Values(4) = DateToStringExt(CR.FROM_DATE)
      Values(5) = DateToStringExt(CR.TO_DATE)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_YearSeq.YearWeeks)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub


Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtYear_Change()
   m_HasModify = True
End Sub

