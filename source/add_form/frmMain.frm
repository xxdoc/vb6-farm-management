VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.TreeView trvMain 
      Height          =   8085
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   14261
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Scroll          =   0   'False
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "AngsanaUPC"
         Size            =   15.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   885
      Left            =   3780
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   1561
      _Version        =   131073
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1290
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2358
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2640
         Top             =   7590
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   28
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2672
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":298C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3266
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5A18
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":62F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6BCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":74A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7D80
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":865A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8F34
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9386
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9C60
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A53A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":AE14
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B6EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":BB40
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":BF92
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":C0EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":C9C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D2A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":DB7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":DE94
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":E76E
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":F448
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":FD22
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":105FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10ED6
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":117B0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Threed.SSCommand cmdAll 
         Height          =   615
         Left            =   6870
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         _Version        =   131073
         PictureFrames   =   1
         Picture         =   "frmMain.frx":1208A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   615
         Left            =   7500
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         _Version        =   131073
         PictureFrames   =   1
         Picture         =   "frmMain.frx":12964
         ButtonStyle     =   3
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   8085
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4674
            MinWidth        =   4674
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "เวอร์ชัน : "
            TextSave        =   "เวอร์ชัน : "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "เวลา : "
            TextSave        =   "เวลา : "
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
   Begin GridEX20.GridEX GridEX1 
      Height          =   6675
      Left            =   3780
      TabIndex        =   2
      Top             =   870
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   11774
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      HeaderFontName  =   "JasmineUPC"
      HeaderFontSize  =   14.25
      FontName        =   "JasmineUPC"
      FontSize        =   14.25
      ColumnHeaderHeight=   390
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   5
      FormatStyle(1)  =   "frmMain.frx":1323E
      FormatStyle(2)  =   "frmMain.frx":13392
      FormatStyle(3)  =   "frmMain.frx":13442
      FormatStyle(4)  =   "frmMain.frx":134F6
      FormatStyle(5)  =   "frmMain.frx":135CE
      ImageCount      =   0
      PrinterProperties=   "frmMain.frx":13686
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   465
      Left            =   10200
      TabIndex        =   6
      Top             =   7590
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   820
      _Version        =   131073
      Caption         =   "SSCommand1"
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   465
      Left            =   7230
      TabIndex        =   5
      Top             =   7590
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   820
      _Version        =   131073
      Caption         =   "SSCommand1"
   End
   Begin Threed.SSCommand cmdEdit 
      Height          =   465
      Left            =   5520
      TabIndex        =   4
      Top             =   7590
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   820
      _Version        =   131073
      Caption         =   "SSCommand1"
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   525
      Left            =   3240
      TabIndex        =   10
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   926
      _Version        =   131073
      Caption         =   "SSCommand1"
   End
   Begin Threed.SSCommand cmdAdd 
      Height          =   465
      Left            =   3810
      TabIndex        =   3
      Top             =   7590
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   820
      _Version        =   131073
      Caption         =   "SSCommand1"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset
Private m_TableName As String

Public HeaderText As String
Private m_XCollection As CXCollection
Private m_YCollection As CYCollection
Private m_Formula As CFormula

Private Sub InitStatusBar()
   stbMain.Panels.Item(1).Text = "ผู้ใช้ : " & glbUser.USER_NAME
   stbMain.Panels.Item(2).Text = "กลุ่ม : " & glbUser.GROUP_NAME
   stbMain.Panels.Item(3).Text = "เวอร์ชัน : " & glbParameterObj.Version & " (Interbase) "
   stbMain.Panels.Item(3).Alignment = sbrLeft
   stbMain.Panels(4).Alignment = sbrCenter
   stbMain.Panels(4).Text = DateToStringExtEx(Now)
End Sub

Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String

   trvMain.Nodes.Clear
   trvMain.Font.NAME = GLB_FONT
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False
   
'   lsvMaster.Font.Name = GLB_FONT
'   lsvMaster.Font.Size = 14
'   lsvMaster.Font.Bold = False
   
   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE, "ระบบทำนายข้อมูลตัวเลข", 2)
   Node.Expanded = True
   Node.Selected = True
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", "ข้อมูลตัวเลข", 1, 1)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", "ข้อมูลกลุ่มตัวเลข", 3, 3)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", "ข้อมูลสูตร", 4, 4)
   Node.Expanded = False
End Sub

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2040
   Col.Caption = "ชื่อ"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 6015
   Col.Caption = "รายละเอียด"
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If trvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   OKClick = False
   If trvMain.SelectedItem.Key = (ROOT_TREE & " 1-0") Then
      Call EnableForm(Me, False)
      frmAddEditXCollection.HeaderText = "เพิ่มข้อมูลตัวเลข"
      frmAddEditXCollection.ShowMode = SHOW_ADD
      Load frmAddEditXCollection
      frmAddEditXCollection.Show 1
         
      OKClick = frmAddEditXCollection.OKClick
      
      Unload frmAddEditXCollection
      Set frmAddEditXCollection = Nothing
      Call EnableForm(Me, True)
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-1") Then
      Call EnableForm(Me, False)
      frmAddEditYCollection.HeaderText = "เพิ่มข้อมูลกลุ่มตัวเลข"
      frmAddEditYCollection.ShowMode = SHOW_ADD
      Load frmAddEditYCollection
      frmAddEditYCollection.Show 1
         
      OKClick = frmAddEditYCollection.OKClick
      
      Unload frmAddEditYCollection
      Set frmAddEditYCollection = Nothing
      Call EnableForm(Me, True)
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-2") Then
      Call EnableForm(Me, False)
      frmAddEditFormula.HeaderText = "เพิ่มข้อมูลสูตร"
      frmAddEditFormula.ShowMode = SHOW_ADD
      Load frmAddEditFormula
      frmAddEditFormula.Show 1
         
      OKClick = frmAddEditFormula.OKClick
      
      Unload frmAddEditFormula
      Set frmAddEditFormula = Nothing
      Call EnableForm(Me, True)
   End If
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdAll_Click()
   Set m_XCollection = Nothing
   Set m_XCollection = New CXCollection
   
   Set m_YCollection = Nothing
   Set m_YCollection = New CYCollection
   
   Set m_Formula = Nothing
   Set m_Formula = New CFormula
   
   Call QueryData(True)
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

'   If Not VerifyAccessRight("PACKAGE_QUOATAPLAN_DELETE") Then
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If

   If trvMain.SelectedItem Is Nothing Then
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
   If trvMain.SelectedItem.Key = (ROOT_TREE & " 1-0") Then
      If Not glbDaily.DeleteXCollection(ID, IsOK, glbErrorLog) Then
         m_XCollection.X_COLLECTION_ID = -1
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-1") Then
      If Not glbDaily.DeleteYCollection(ID, IsOK, glbErrorLog) Then
         m_YCollection.Y_COLLECTION_ID = -1
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-2") Then
      If Not glbDaily.DeleteFormula(ID, IsOK, glbErrorLog) Then
         m_Formula.FORMULA_ID = -1
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
         Call EnableForm(Me, True)
         Exit Sub
      End If
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

'   If Not VerifyAccessRight("PACKAGE_QUOATAPLAN_QUERY") Then
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If

   If trvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
            
   Call EnableForm(Me, False)
   If trvMain.SelectedItem.Key = (ROOT_TREE & " 1-0") Then
      frmAddEditXCollection.ID = ID
      frmAddEditXCollection.HeaderText = "แก้ไขข้อมูลตัวเลข"
      frmAddEditXCollection.ShowMode = SHOW_EDIT
      Load frmAddEditXCollection
      frmAddEditXCollection.Show 1
         
      OKClick = frmAddEditXCollection.OKClick
      
      Unload frmAddEditXCollection
      Set frmAddEditXCollection = Nothing
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-1") Then
      frmAddEditYCollection.ID = ID
      frmAddEditYCollection.HeaderText = "แก้ไขข้อมูลกลุ่มตัวเลข"
      frmAddEditYCollection.ShowMode = SHOW_EDIT
      Load frmAddEditYCollection
      frmAddEditYCollection.Show 1
         
      OKClick = frmAddEditYCollection.OKClick
      
      Unload frmAddEditYCollection
      Set frmAddEditYCollection = Nothing
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-2") Then
      frmAddEditFormula.ID = ID
      frmAddEditFormula.HeaderText = "แก้ไขข้อมูลสูตร"
      frmAddEditFormula.ShowMode = SHOW_EDIT
      Load frmAddEditFormula
      frmAddEditFormula.Show 1
         
      OKClick = frmAddEditFormula.OKClick
      
      Unload frmAddEditFormula
      Set frmAddEditFormula = Nothing
    End If
   Call EnableForm(Me, True)
            
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim OKClick As Boolean

   If trvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If trvMain.SelectedItem.Key = (ROOT_TREE & " 1-0") Then
      frmSearch.HeaderText = "ค้นหาข้อมูลลูกค้า"
      Set frmSearch.SearchRec = m_XCollection
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-1") Then
      frmSearch.HeaderText = "ค้นหาข้อมูลประวัติการใช้ยา"
      Set frmSearch.SearchRec = m_YCollection
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-2") Then
      frmSearch.HeaderText = "ค้นหาข้อมูลประวัติผลเลือด"
      Set frmSearch.SearchRec = m_Formula
   Else
      Exit Sub
   End If
   Call EnableForm(Me, False)
   Load frmSearch
   frmSearch.Show 1
      
   OKClick = frmSearch.OKClick
   
   Unload frmSearch
   Set frmSearch = Nothing
   
   If OKClick Then
      Me.Refresh
      QueryData (True)
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
'      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If MustAsk Then
      glbErrorLog.LocalErrorMsg = "ท่านต้องการออกจากโปรแกรมนี้ใช่หรือไม่"
      If glbErrorLog.AskMessage = vbNo Then
         Cancel = 1
         Exit Sub
      End If
   End If
   
   Cancel = 0
   Call ReleaseAll
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_XCollection = Nothing
   Set m_YCollection = Nothing
   Set m_Formula = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   Call InitHeaderFooter(pnlHeader, Nothing)
   
   cmdSearch.BackColor = GLB_FORM_COLOR
   cmdAll.BackColor = GLB_FORM_COLOR
   
   Call InitMainButton(cmdAdd, "เพิ่ม (F7)")
   Call InitMainButton(cmdEdit, "แก้ไข (F3)")
   Call InitMainButton(cmdDelete, "ลบ (F6)")
   Call InitMainButton(cmdExit, "ออก (ESC)")
   
   cmdSearch.ToolTipText = "ค้นหาตามเงื่อนไข"
   cmdAll.ToolTipText = "ดูข้อมูลทั้งหมด"
   
   Call InitMainTreeview
   Me.Caption = "ระบบทำนายข้อมูลตัวเลข"
   
   Call InitGrid
   Call trvMain_NodeClick(trvMain.SelectedItem)
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim ItemCount As Long

   If Not m_HasActivate Then
      m_HasActivate = True
   Else
      Exit Sub
   End If

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "Form_Activate"

'   Load frmLogin
'   frmLogin.Show 1
'
'   If Not frmLogin.OKClick Then
'      MustAsk = False
'      Unload Me
'      End
'   Else
'      Call InitFormLayout
'      Call InitStatusBar
'   End If
   
   Call InitStatusBar
'   Call PatchDB
'   Call CheckMemo(1)
'
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   Call EnableForm(Me, True)
'
'   Unload frmLogin
'   Set frmLogin = Nothing
   
   Exit Sub
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_Load()
   MustAsk = True
   Call InitFormLayout
   
   Set m_XCollection = New CXCollection
   Set m_YCollection = New CYCollection
   Set m_Formula = New CFormula
   
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
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
   If trvMain.SelectedItem.Key = (ROOT_TREE & " 1-0") Then
      Values(1) = NVLI(m_Rs("X_COLLECTION_ID"), -1)
      Values(2) = NVLS(m_Rs("X_COLLECTION_NAME"), "")
      Values(3) = NVLS(m_Rs("X_COLLECTION_DESC"), "")
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-1") Then
      Values(1) = NVLI(m_Rs("Y_COLLECTION_ID"), -1)
      Values(2) = NVLS(m_Rs("Y_COLLECTION_NAME"), "")
      Values(3) = NVLS(m_Rs("Y_COLLECTION_DESC"), "")
   ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-2") Then
      Values(1) = NVLI(m_Rs("FORMULA_ID"), -1)
      Values(2) = NVLS(m_Rs("FORMULA_NAME"), "")
      Values(3) = NVLS(m_Rs("FORMULA_DESC"), "")
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   
   stbMain.Panels(4).Alignment = sbrCenter
   stbMain.Panels(4).Text = DateToStringExtEx(Now)
   
   Timer1.Enabled = True
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
   
   If trvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
'      If Not VerifyAccessRight("INVENTORY_SUPPLIER_QUERY") Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
      If trvMain.SelectedItem.Key = (ROOT_TREE & " 1-0") Then
         m_XCollection.X_COLLECTION_ID = -1
         If Not glbDaily.QueryXCollection(m_XCollection, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-1") Then
         m_YCollection.Y_COLLECTION_ID = -1
         If Not glbDaily.QueryYCollection(m_YCollection, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf trvMain.SelectedItem.Key = (ROOT_TREE & " 1-2") Then
         m_Formula.FORMULA_ID = -1
         If Not glbDaily.QueryFormula(m_Formula, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
      End If
      
   End If

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind

'   Label1.Caption = ItemCount
   
   Call EnableForm(Me, True)
End Sub

Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
   pnlHeader.Caption = Node.Text
   Call QueryData(True)
End Sub


