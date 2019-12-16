VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditPackage 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditPackage.frx":0000
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
      Height          =   8535
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   6975
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   4
         Top             =   2520
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
      Begin prjFarmManagement.uctlTextBox txtPackageName 
         Height          =   435
         Left            =   2220
         TabIndex        =   2
         Top             =   1350
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPackageCode 
         Height          =   435
         Left            =   2220
         TabIndex        =   0
         Top             =   900
         Width           =   2835
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
         Height          =   4695
         Left            =   150
         TabIndex        =   5
         Top             =   3030
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8281
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
         Column(1)       =   "frmAddEditPackage.frx":27A2
         Column(2)       =   "frmAddEditPackage.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPackage.frx":290E
         FormatStyle(2)  =   "frmAddEditPackage.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPackage.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPackage.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPackage.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPackage.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSCheck chkPackageBasic 
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   900
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblArea 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7230
         TabIndex        =   17
         Top             =   3390
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackage.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   10
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1800
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackage.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackage.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblPackageCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblEnterpriseType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7230
         TabIndex        =   14
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label lblPackageType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   30
         TabIndex        =   13
         Top             =   1860
         Width           =   2085
      End
      Begin VB.Label lblPackageName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAddEditPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Package As CPackage

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
            
      m_Package.PKG_ID = ID
      
      If Not glbDaily.QueryPackage(m_Package, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
     Call m_Package.PopulateFromRS(1, m_Rs)
      
      txtPackageCode.Text = m_Package.PKG_CODE
      txtPackageName.Text = m_Package.PKG_NAME
      CboPackageType.ListIndex = IDToListIndex(CboPackageType, m_Package.PKG_TYPE)
      chkPackageBasic.Value = FlagToCheck(m_Package.PKG_BASIC_FLAG)
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
      If Not VerifyAccessRight("PACKAGE_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("PACKAGE_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblPackageCode, txtPackageCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPackageName, txtPackageName, False) Then
      Exit Function
   End If
   

   If Not CheckUniqueNs(PACKAGE_CODE, txtPackageCode.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPackageCode.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not CheckUniqueNs(PACKAGE_NAME, txtPackageName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPackageName.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
    If chkPackageBasic.Value = ssCBChecked Then
        If Not CheckUniqueNs(PACKAGE_BASIC, "Y", ID, , CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))) Then
            glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลการเซ็ต") & " (" & chkPackageBasic.Caption & ") และ ประเภทการตั้งราคาเป็น (" & CboPackageType.Text & ") " & MapText("อยู่ในระบบแล้ว")
            glbErrorLog.ShowUserError
            Exit Function
        End If
    End If
    
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Package.AddEditMode = ShowMode
   m_Package.PKG_CODE = txtPackageCode.Text
   m_Package.PKG_NAME = txtPackageName.Text
   m_Package.PKG_TYPE = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
   
   m_Package.PKG_BASIC_FLAG = Check2Flag(chkPackageBasic.Value)
   
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditPackage(m_Package, IsOK, True, glbErrorLog) Then
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

Private Sub chkPackageBasic_Click(Value As Integer)
    m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditPackageItem.TempCollection = m_Package.PackageDetail
      Set frmAddEditPackageItem.ParentForm = Me
      frmAddEditPackageItem.PackageType = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
      frmAddEditPackageItem.ShowMode = SHOW_ADD
      frmAddEditPackageItem.HeaderText = MapText("เพิ่มรายละเอียด")
      Load frmAddEditPackageItem
      frmAddEditPackageItem.Show 1

      OKClick = frmAddEditPackageItem.OKClick

      Unload frmAddEditPackageItem
      Set frmAddEditPackageItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Package.PackageDetail)
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
         m_Package.PackageDetail.Remove (ID2)
      Else
         m_Package.PackageDetail.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Package.PackageDetail)
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
      frmAddEditPackageItem.ID = ID
      Set frmAddEditPackageItem.TempCollection = m_Package.PackageDetail
      Set frmAddEditPackageItem.ParentForm = Me
      frmAddEditPackageItem.PackageType = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
      frmAddEditPackageItem.HeaderText = MapText("แก้ไขรายละเอียด")
      frmAddEditPackageItem.ShowMode = SHOW_EDIT
      Load frmAddEditPackageItem
      frmAddEditPackageItem.Show 1

      OKClick = frmAddEditPackageItem.OKClick

      Unload frmAddEditPackageItem
      Set frmAddEditPackageItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Package.PackageDetail)
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call InitPackageType(CboPackageType)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Package.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Package.QueryFlag = 0
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
   
   Set m_Package = Nothing
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
   Col.Width = 1905
   Col.Caption = MapText("สถานะ")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1425
   Col.Caption = MapText("จาก นน.")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1350
   Col.Caption = MapText("ถึง นน.")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 1185
   Col.Caption = MapText("ตัดที่นน.")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 1680
   Col.Caption = MapText("ราคา/ส่วนเกิน")
   Col.TextAlignment = jgexAlignRight
     
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1200
   Col.Caption = MapText("ราคา/นน.")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '9
   Col.Width = 1200
   Col.Caption = MapText("ค่าคงที่")
   Col.TextAlignment = jgexAlignRight
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblPackageCode, MapText("รหัสการตั้งราคา"))
   Call InitNormalLabel(lblPackageName, MapText("รายละเอียด"))
   Call InitNormalLabel(lblPackageType, MapText("ประเภทการตั้งราคา"))
   
   Call InitCombo(CboPackageType)
   
   Call InitCheckBox(chkPackageBasic, "ราคาพื้นฐาน")
   
   Call txtPackageCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtPackageName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
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
   TabStrip1.Tabs.Add().Caption = MapText("รายละเอียดการตั้งราคาสินค้าขาย")
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
   Set m_Package = New CPackage
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
      If m_Package.PackageDetail Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CPackageDetail
      
      If m_Package.PackageDetail.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Package.PackageDetail, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
    
      Values(1) = CR.PKG_DETAIL_ID
      Values(2) = RealIndex
      Values(3) = CR.PRODUCT_STATUS_NAME
      Values(4) = FormatNumber(CR.FROM_WEIGHT)
      Values(5) = FormatNumber(CR.TO_WEIGHT)
      Values(6) = FormatNumber(CR.CUT_WEIGHT)
      Values(7) = FormatNumber(CR.COST_PER_EXCEED)
      Values(8) = FormatNumber(CR.COST_PER_WEIGHT)
      Values(9) = FormatNumber(CR.COST_PER_UNIT)
      
      
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
      GridEX1.ItemCount = CountItem(m_Package.PackageDetail)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtPackageCode_Change()
    m_HasModify = True
End Sub

Private Sub txtPackageName_Change()
    m_HasModify = True
End Sub
 
Private Sub txtPackageType_Click()
    m_HasModify = True
End Sub
Public Sub ShowPackageItemGrid()
   GridEX1.ItemCount = CountItem(m_Package.PackageDetail)
   GridEX1.Rebind
   
   m_HasModify = True
End Sub

