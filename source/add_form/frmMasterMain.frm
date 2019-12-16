VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMasterMain 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMasterMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   2
         Top             =   7800
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8445
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10095
            TabIndex        =   7
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   525
            Left            =   1770
            TabIndex        =   6
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   525
            Left            =   150
            TabIndex        =   5
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2ABC
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   525
            Left            =   3420
            TabIndex        =   4
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2DD6
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1508
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":30F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":39CC
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2850
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":3CE8
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   6945
         Left            =   0
         TabIndex        =   3
         Top             =   870
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   12250
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   15.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6915
         Left            =   4500
         TabIndex        =   9
         Top             =   900
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   12197
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
         Column(1)       =   "frmMasterMain.frx":4002
         Column(2)       =   "frmMasterMain.frx":40CA
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMasterMain.frx":416E
         FormatStyle(2)  =   "frmMasterMain.frx":42CA
         FormatStyle(3)  =   "frmMasterMain.frx":437A
         FormatStyle(4)  =   "frmMasterMain.frx":442E
         FormatStyle(5)  =   "frmMasterMain.frx":4506
         ImageCount      =   0
         PrinterProperties=   "frmMasterMain.frx":45BE
      End
   End
End
Attribute VB_Name = "frmMasterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String
Private m_PartType As CPartType
Private m_Location As CLocation
Private m_ProductType As CProductType
Private m_ProductStatus As CProductStatus
Private m_House As CHouse
Private m_Country As CCountry
Private m_CustomerType As CCustomerType
Private m_CustomerGrade As CCustomerGrade
Private m_SupplierType As CSupplierType
Private m_SupplierGrade As CSupplierGrade
Private m_SupplierStatus As CSupplierStatus
Private m_Position As CEmpPosition
Private m_Unit As CUnit
Private m_PartGroup As CPartGroup
'Private m_ExposeType As CHGroupItem
Private m_HouseGroup As CHouseGroup
Private m_StatusGroup As CStatusGroup
Private m_AgeRange As CAgeRange
Private m_ExpenseType As CExpenseType
Private m_DocumentType As CDocumentType
Private m_Bank As CBank
Private m_BankBranch As CBankBranch
Private m_Region As CRegion
Private m_RevenueType As CRevenueType
Private m_CnDnReasons As CCnDnReason
Private m_BankAccount As CBankAccount
Private m_StatusType As CStatusType
Private m_PackageType As CPackageType

Private m_MasterRef As CMasterRef

Public HeaderText As String
Public MasterMode As Long

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_PIG_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_LEDGER_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 5 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If trvMaster.SelectedItem.Key = ROOT_TREE Then
      glbErrorLog.LocalErrorMsg = ""
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = "Root 2-4" Then
      frmAddEditHouseGroup.ShowMode = SHOW_ADD
      frmAddEditHouseGroup.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditHouseGroup
      frmAddEditHouseGroup.Show 1
      
      OKClick = frmAddEditHouseGroup.OKClick
      
      Unload frmAddEditHouseGroup
      Set frmAddEditHouseGroup = Nothing
   ElseIf trvMaster.SelectedItem.Key = "Root 2-5" Then
      frmAddEditStatusGroup.ShowMode = SHOW_ADD
      frmAddEditStatusGroup.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditStatusGroup
      frmAddEditStatusGroup.Show 1
      
      OKClick = frmAddEditStatusGroup.OKClick
      
      Unload frmAddEditStatusGroup
      Set frmAddEditStatusGroup = Nothing
   ElseIf trvMaster.SelectedItem.Key = "Root 2-6" Then
      frmAddEditAgeRange.ShowMode = SHOW_ADD
      frmAddEditAgeRange.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditAgeRange
      frmAddEditAgeRange.Show 1
      
      OKClick = frmAddEditAgeRange.OKClick
      
      Unload frmAddEditAgeRange
      Set frmAddEditAgeRange = Nothing
   ElseIf trvMaster.SelectedItem.Key = "Root 4-1" Then
      frmAddEditExpenseType.ShowMode = SHOW_ADD
      frmAddEditExpenseType.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditExpenseType
      frmAddEditExpenseType.Show 1
      
      OKClick = frmAddEditExpenseType.OKClick
      
      Unload frmAddEditExpenseType
      Set frmAddEditExpenseType = Nothing
   Else
      frmAddEditMaster1.MasterMode = MasterMode
      frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
      frmAddEditMaster1.ShowMode = SHOW_ADD
      frmAddEditMaster1.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditMaster1
      frmAddEditMaster1.Show 1
      
      OKClick = frmAddEditMaster1.OKClick
      
      Unload frmAddEditMaster1
      Set frmAddEditMaster1 = Nothing
   End If
   
   If OKClick Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   End If
End Sub


Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.Name = GLB_FONT
   trvMaster.Font.Size = 14
   
   If MasterMode = 1 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-4", MapText("กลุ่มวัตถุดิบ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("ประเภทวัตถุดิบ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("สถานที่จัดเก็บ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("หน่วยวัด"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-5", MapText("กลุ่มอาหาร"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-6", MapText("ประเภทการเบิก"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 2 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
'
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-1", MapText("ประเภทสุกร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-7", MapText("ประเภทสถานะสุกร"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-2", MapText("สถานะสุกร"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-3", MapText("โรงเรือนสุกร"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-4", MapText("กลุ่มโรงเรือนสุกร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-5", MapText("กลุ่มสถานะสุกร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-6", MapText("ช่วงอายุสุกร"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("ประเทศ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-2", MapText("ระดับลูกค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-3", MapText("ประเภทลูกค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-4", MapText("ระดับซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-5", MapText("ประเภทซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-6", MapText("สถานะซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-7", MapText("ตำแหน่ง"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-1", MapText("ประเภทรายจ่าย"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-2", MapText("ประเภทเอกสาร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-3", MapText("ธนาคาร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-4", MapText("สาขาธนาคาร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-8", MapText("เลขที่บัญชีธนาคาร"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5", MapText("เขตการค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-6", MapText("ประเภทรายรับอื่น ๆ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-7", MapText("สาเหตุการ เพิ่ม/ลด หนี้"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-9", MapText("ประเภทเช็ค"), 1, 2)
      Node.Expanded = False
      
      '4-8 Already used
   ElseIf MasterMode = 5 Then
   
    Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
    Node.Expanded = True
    Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-1", MapText("ประเภทการตั้งราคา"), 1, 2)
      Node.Expanded = False

   ElseIf MasterMode = 6 Then
   End If
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHandler
Dim Status As Boolean
Dim IsOK As Boolean
Dim TempID As Long

   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
      
   If MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_PIG_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_LEDGER_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 5 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-1" Then
      Status = glbMaster.DeletePartType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-2" Then
      Status = glbMaster.DeleteLocation(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-3" Then
      Status = glbMaster.DeleteUnit(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-4" Then
      Status = glbMaster.DeletePartGroup(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-5" Then
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 1-6" Then
      Status = glbMaster.DeleteExposeType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-1" Then
      Status = glbMaster.DeleteProductType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-2" Then
      Status = glbMaster.DeleteProductStatus(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-3" Then
      Status = glbMaster.DeleteLocation(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-4" Then
      Status = glbMaster.DeleteHouseGroup(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-5" Then
      Status = glbMaster.DeleteStatusGroup(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-6" Then
      Status = glbMaster.DeleteAgeRange(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 2-7" Then
      Status = glbMaster.DeleteStatusType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1" Then
      Status = glbMaster.DeleteCountry(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-2" Then
      Status = glbMaster.DeleteCustomerGrade(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-3" Then
      Status = glbMaster.DeleteCustomerType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-4" Then
      Status = glbMaster.DeleteSupplierGrade(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-5" Then
      Status = glbMaster.DeleteSupplierType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-6" Then
      Status = glbMaster.DeleteSupplierStatus(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 3-7" Then
      Status = glbMaster.DeletePosition(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-1" Then
      Status = glbMaster.DeleteExpenseType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-2" Then
      Status = glbMaster.DeleteDocumentType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-3" Then
      Status = glbMaster.DeleteBank(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-4" Then
      Status = glbMaster.DeleteBankBranch(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5" Then
      Status = glbMaster.DeleteRegion(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-6" Then
      Status = glbMaster.DeleteRevenueType(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-7" Then
      Status = glbMaster.DeleteCnDnReason(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-8" Then
      Status = glbMaster.DeleteBankAccount(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 4-9" Then
      Status = glbMaster.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   ElseIf trvMaster.SelectedItem.Key = ROOT_TREE & " 5-1" Then
      Status = glbMaster.DeletePackageType(TempID, IsOK, glbErrorLog)
   End If
   
   If Status Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   Else
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Exit Sub
   
ErrorHandler:
End Sub

Private Sub cmdEdit_Click()
Dim OKClick As Boolean
Dim TempID As Long

   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
   
   If MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_PIG_EDIT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_EDIT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_EDIT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_LEDGER_EDIT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 5 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_EDIT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If trvMaster.SelectedItem.Key = "Root 2-4" Then
      frmAddEditHouseGroup.ID = TempID
      frmAddEditHouseGroup.ShowMode = SHOW_EDIT
      frmAddEditHouseGroup.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditHouseGroup
      frmAddEditHouseGroup.Show 1
      
      OKClick = frmAddEditHouseGroup.OKClick
      
      Unload frmAddEditHouseGroup
      Set frmAddEditHouseGroup = Nothing
   ElseIf trvMaster.SelectedItem.Key = "Root 2-5" Then
      frmAddEditStatusGroup.ID = TempID
      frmAddEditStatusGroup.ShowMode = SHOW_EDIT
      frmAddEditStatusGroup.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditStatusGroup
      frmAddEditStatusGroup.Show 1
      
      OKClick = frmAddEditStatusGroup.OKClick
      
      Unload frmAddEditStatusGroup
      Set frmAddEditStatusGroup = Nothing
   ElseIf trvMaster.SelectedItem.Key = "Root 2-6" Then
      frmAddEditAgeRange.ID = TempID
      frmAddEditAgeRange.ShowMode = SHOW_EDIT
      frmAddEditAgeRange.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditAgeRange
      frmAddEditAgeRange.Show 1
      
      OKClick = frmAddEditAgeRange.OKClick
      
      Unload frmAddEditAgeRange
      Set frmAddEditAgeRange = Nothing
   ElseIf trvMaster.SelectedItem.Key = "Root 4-1" Then
      frmAddEditExpenseType.ID = TempID
      frmAddEditExpenseType.ShowMode = SHOW_EDIT
      frmAddEditExpenseType.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditExpenseType
      frmAddEditExpenseType.Show 1
      
      OKClick = frmAddEditExpenseType.OKClick
      
      Unload frmAddEditExpenseType
      Set frmAddEditExpenseType = Nothing
   Else
      frmAddEditMaster1.MasterMode = MasterMode
      frmAddEditMaster1.ID = TempID
      frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
      frmAddEditMaster1.ShowMode = SHOW_EDIT
      frmAddEditMaster1.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
      Load frmAddEditMaster1
      frmAddEditMaster1.Show 1
      
      OKClick = frmAddEditMaster1.OKClick
      
      Unload frmAddEditMaster1
      Set frmAddEditMaster1 = Nothing
   End If
   
   If OKClick Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   End If
End Sub

Private Sub Form_Activate()
Dim ItemCount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      
      m_HasActivate = True
   End If
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
'      Call cmdOK_Click
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PartType = Nothing
   Set m_Location = Nothing
   Set m_ProductType = Nothing
   Set m_ProductStatus = Nothing
   Set m_House = Nothing
   Set m_Country = Nothing
   Set m_CustomerGrade = Nothing
   Set m_CustomerType = Nothing
   Set m_SupplierGrade = Nothing
   Set m_SupplierType = Nothing
   Set m_SupplierStatus = Nothing
   Set m_Position = Nothing
   Set m_Unit = Nothing
   Set m_PartGroup = Nothing
 '  Set m_ExposeType = Nothing
   Set m_HouseGroup = Nothing
   Set m_StatusGroup = Nothing
   Set m_AgeRange = Nothing
   Set m_ExpenseType = Nothing
   Set m_DocumentType = Nothing
   Set m_Bank = Nothing
   Set m_BankBranch = Nothing
   Set m_Region = Nothing
   Set m_RevenueType = Nothing
   Set m_CnDnReasons = Nothing
   Set m_BankAccount = Nothing
   Set m_StatusType = Nothing
   Set m_PackageType = Nothing
   
   Set m_MasterRef = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid0()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1()
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
   Col.Width = 1110
   Col.Caption = MapText("หมายเลขวัตถุดิบ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 6225
   Col.Caption = MapText("วัตถุดิบ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสสถานที่จัดเก็บ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("สถานที่จัดเก็บ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_3()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสหน่วยวัด")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("หน่วยวัด")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_4()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสกลุ่มวัตถุดิบ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("กลุ่มวัตถุดิบ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_6()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเภท")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภทการโอน")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_1()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเภทสุกร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภทสุกร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_1()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเทศ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเทศ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_2()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสระดับลูกค้า")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ระดับลูกค้า")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_3()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเภทลูกค้า")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภทลูกค้า")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_4()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสระดับซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ระดับซับพลายเออร์")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_5()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสประเภทซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทซับพลายเออร์")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_6()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสสถานะซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("สถานะซับพลายเออร์")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid3_7()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสตำแหน่ง")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ตำแหน่ง")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_1()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสประเภทรายจ่าย")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทรายจ่าย")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_2()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสประเภทเอกสาร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทเอกสาร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_3()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสธนาคาร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ธนาคาร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_4()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสสาขาธนาคาร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("สาขาธนาคาร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_5()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสเขตการค้า")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("เขตการค้า")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_6()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสรายรับ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("รายรับ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_7()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสสาเหตุ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("สาเหตุการ เพิ่ม/ลด หนี้")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid4_8()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสบัญชี")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("เลขที่บัญชี")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid5_1()
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
   Col.Width = 2220
   Col.Caption = MapText("รหัสการตั้งราคา")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทการตั้งราคา")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_2()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสสถานะสุกร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("สถานะสุกร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_3()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสโรงเรือน")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("โรงเรือน")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_4()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสกลุ่มโรงเรือน")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ชื่อกลุ่มโรงเรือน")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_5()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสกลุ่มสถานะสุกร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ชื่อกลุ่มสถานะสุกร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_6()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสช่วงอายุสุกร")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ชื่อช่วงอายุสุกร")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid2_7()
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
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเภทสถานะ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ชื่อประเภทสถานะ")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdExit, MapText("ออก (ESC)"))
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitTreeView
   Call InitGrid0
   
'   lsvMaster.Font.NAME = GLB_FONT
'   lsvMaster.Font.Size = 14
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Call InitFormLayout
   
   m_HasActivate = False
   m_TableName = "SYSTEM_PARAM"
   Set m_Rs = New ADODB.Recordset
   
   Set m_PartType = New CPartType
   Set m_Location = New CLocation
   Set m_ProductType = New CProductType
   Set m_ProductStatus = New CProductStatus
   Set m_House = New CHouse
   Set m_Country = New CCountry
   Set m_CustomerGrade = New CCustomerGrade
   Set m_CustomerType = New CCustomerType
   Set m_SupplierGrade = New CSupplierGrade
   Set m_SupplierType = New CSupplierType
   Set m_SupplierStatus = New CSupplierStatus
   Set m_Position = New CEmpPosition
   Set m_Unit = New CUnit
   Set m_PartGroup = New CPartGroup
'   Set m_ExposeType = New CExposeType
   Set m_HouseGroup = New CHouseGroup
   Set m_StatusGroup = New CStatusGroup
   Set m_AgeRange = New CAgeRange
   Set m_ExpenseType = New CExpenseType
   Set m_DocumentType = New CDocumentType
   Set m_Bank = New CBank
   Set m_BankBranch = New CBankBranch
   Set m_Region = New CRegion
   Set m_RevenueType = New CRevenueType
   Set m_CnDnReasons = New CCnDnReason
   Set m_BankAccount = New CBankAccount
   Set m_StatusType = New CStatusType
   Set m_PackageType = New CPackageType
   
   Set m_MasterRef = New CMasterRef
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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
   
   If MasterMode = 1 Then
      If trvMaster.SelectedItem.Key = "Root 1-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_PartType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_PartType.PART_TYPE_ID
         Values(2) = m_PartType.PART_TYPE_NO
         Values(3) = m_PartType.PART_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Location.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Location.LOCATION_ID
         Values(2) = m_Location.LOCATION_NO
         Values(3) = m_Location.LOCATION_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Unit.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Unit.UNIT_ID
         Values(2) = m_Unit.UNIT_NO
         Values(3) = m_Unit.UNIT_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_PartGroup.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_PartGroup.PART_GROUP_ID
         Values(2) = m_PartGroup.PART_GROUP_NO
         Values(3) = m_PartGroup.PART_GROUP_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 1-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
'      ElseIf trvMaster.SelectedItem.Key = "Root 1-6" Then       'Ging
'         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
'         Call m_ExposeType.PopulateFromRS(1, m_Rs)
'
'         Values(1) = m_ExposeType.EXPOSE_TYPE_ID               ' m_ExposeType
'         Values(2) = m_ExposeType.EXPOSE_TYPE_NO
'         Values(3) = m_ExposeType.EXPOSE_TYPE_NAME
      End If
   ElseIf MasterMode = 2 Then
      If trvMaster.SelectedItem.Key = "Root 2-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_ProductType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_ProductType.PRODUCT_TYPE_ID
         Values(2) = m_ProductType.PRODUCT_TYPE_NO
         Values(3) = m_ProductType.PRODUCT_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_ProductStatus.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_ProductStatus.PRODUCT_STATUS_ID
         Values(2) = m_ProductStatus.PRODUCT_STATUS_NO
         Values(3) = m_ProductStatus.PRODUCT_STATUS_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Location.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Location.LOCATION_ID
         Values(2) = m_Location.LOCATION_NO
         Values(3) = m_Location.LOCATION_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_HouseGroup.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_HouseGroup.HOUSE_GROUP_ID
         Values(2) = m_HouseGroup.HOUSE_GROUP_NO
         Values(3) = m_HouseGroup.HOUSE_GROUP_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_StatusGroup.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_StatusGroup.STATUS_GROUP_ID
         Values(2) = m_StatusGroup.STATUS_GROUP_NO
         Values(3) = m_StatusGroup.STATUS_GROUP_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-6" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_AgeRange.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_AgeRange.AGE_RANGE_ID
         Values(2) = m_AgeRange.AGE_RANGE_NO
         Values(3) = m_AgeRange.AGE_RANGE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 2-7" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_StatusType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_StatusType.STATUS_TYPE_ID
         Values(2) = m_StatusType.STATUS_TYPE_NO
         Values(3) = m_StatusType.STATUS_TYPE_NAME
      End If
   ElseIf MasterMode = 3 Then
      If trvMaster.SelectedItem.Key = "Root 3-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Country.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Country.COUNTRY_ID
         Values(2) = m_Country.COUNTRY_NO
         Values(3) = m_Country.COUNTRY_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_CustomerGrade.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_CustomerGrade.CSTGRADE_ID
         Values(2) = m_CustomerGrade.CSTGRADE_NO
         Values(3) = m_CustomerGrade.CSTGRADE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_CustomerType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_CustomerType.CSTTYPE_ID
         Values(2) = m_CustomerType.CSTTYPE_NO
         Values(3) = m_CustomerType.CSTTYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_SupplierGrade.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_SupplierGrade.SUPPLIER_GRADE_ID
         Values(2) = m_SupplierGrade.SUPPLIER_GRADE_NO
         Values(3) = m_SupplierGrade.SUPPLIER_GRADE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_SupplierType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_SupplierType.SUPPLIER_TYPE_ID
         Values(2) = m_SupplierType.SUPPLIER_TYPE_NO
         Values(3) = m_SupplierType.SUPPLIER_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-6" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_SupplierStatus.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_SupplierStatus.SUPPLIER_STATUS_ID
         Values(2) = m_SupplierStatus.SUPPLIER_STATUS_NO
         Values(3) = m_SupplierStatus.SUPPLIER_STATUS_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 3-7" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Position.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Position.POSITION_ID
         Values(2) = m_Position.POSITION_NAME
         Values(3) = m_Position.POSITION_DESC
      End If
   ElseIf MasterMode = 4 Then
      If trvMaster.SelectedItem.Key = "Root 4-1" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_ExpenseType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_ExpenseType.EXPENSE_TYPE_ID
         Values(2) = m_ExpenseType.EXPENSE_TYPE_NO
         Values(3) = m_ExpenseType.EXPENSE_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-2" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_DocumentType.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_DocumentType.DOCUMENT_TYPE_ID
         Values(2) = m_DocumentType.DOCUMENT_TYPE_NO
         Values(3) = m_DocumentType.DOCUMENT_TYPE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-3" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Bank.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Bank.BANK_ID
         Values(2) = m_Bank.BANK_NO
         Values(3) = m_Bank.BANK_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-4" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_BankBranch.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_BankBranch.BBRANCH_ID
         Values(2) = m_BankBranch.BBRANCH_NO
         Values(3) = m_BankBranch.BBRANCH_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-5" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_Region.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_Region.REGION_ID
         Values(2) = m_Region.REGION_NO
         Values(3) = m_Region.REGION_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-6" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_RevenueType.PopulateFromRS(1, m_Rs)

         Values(1) = m_RevenueType.REVENUE_TYPE_ID
         Values(2) = m_RevenueType.REVENUE_NO
         Values(3) = m_RevenueType.REVENUE_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-7" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_CnDnReasons.PopulateFromRS(1, m_Rs)

         Values(1) = m_CnDnReasons.REASON_ID
         Values(2) = m_CnDnReasons.REASON_NO
         Values(3) = m_CnDnReasons.REASON_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-8" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_BankAccount.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_BankAccount.BANK_ACCOUNT_ID
         Values(2) = m_BankAccount.ACCOUNT_NO
         Values(3) = m_BankAccount.ACCOUNT_NAME
      ElseIf trvMaster.SelectedItem.Key = "Root 4-9" Then
         Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         
         Values(1) = m_MasterRef.KEY_ID
         Values(2) = m_MasterRef.KEY_CODE
         Values(3) = m_MasterRef.KEY_NAME
      End If
      
    ElseIf MasterMode = 5 Then
        If trvMaster.SelectedItem.Key = "Root 5-1" Then
            Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
            Call m_PackageType.PopulateFromRS(1, m_Rs)
         
            Values(1) = m_PackageType.PACKAGE_TYPE_ID
            Values(2) = m_PackageType.PACKAGE_TYPE_CODE
            Values(3) = m_PackageType.PACKAGE_TYPE_NAME
        End If
   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Private Sub LoadListView(Rs As ADODB.Recordset, FieldName As String, IDName As String)
'Dim Lst As ListItem
'
'   While Not Rs.EOF
'      Set Lst = lsvMaster.ListItems.Add(, , NVLS(Rs(FieldName), ""), 1, 1)
'      Lst.Tag = NVLI(Rs(IDName), 0)
'      Rs.MoveNext
'   Wend
'End Sub

Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim ItemCount As Long
Dim QueryFlag As Boolean

   If LastKey = Node.Key Then
      Exit Sub
   End If

   Status = True
   QueryFlag = False
   
   If Node.Key = ROOT_TREE & " 1-1" Then
      Call InitGrid1
      Dim a1_1 As CPartType

      Set a1_1 = New CPartType
      a1_1.PART_TYPE_ID = -1
      Status = a1_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitGrid2
      Dim a1_2 As CLocation

      Set a1_2 = New CLocation
      a1_2.LOCATION_ID = -1
      a1_2.LOCATION_TYPE = 2
      Status = a1_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitGrid1_3
      Dim a1_3 As CUnit

      Set a1_3 = New CUnit
      a1_3.UNIT_ID = -1
      Status = a1_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-4" Then
      Call InitGrid1_4
      Dim a1_4 As CPartGroup

      Set a1_4 = New CPartGroup
      a1_4.PART_GROUP_ID = -1
      Status = a1_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-5" Then
      Call InitGrid1_5
      Dim a1_5 As CMasterRef

      Set a1_5 = New CMasterRef
      a1_5.KEY_ID = -1
      a1_5.MASTER_AREA = FEED_GROUP
      Status = a1_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 1-6" Then       ' Ging
      Call InitGrid1_6
      Dim a1_6 As CExposeType

      Set a1_6 = New CExposeType
      a1_6.EXPOSE_TYPE_ID = -1
      Status = a1_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a1_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-1" Then
      Call InitGrid2_1
      Dim a2_1 As CProductType

      Set a2_1 = New CProductType
      a2_1.PRODUCT_TYPE_ID = -1
      Status = a2_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-2" Then
      Call InitGrid2_2
      Dim a2_2 As CProductStatus

      Set a2_2 = New CProductStatus
      a2_2.PRODUCT_STATUS_ID = -1
      Status = a2_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-3" Then
      Call InitGrid2_3
      Dim a2_3 As CLocation

      Set a2_3 = New CLocation
      a2_3.LOCATION_ID = -1
      a2_3.LOCATION_TYPE = 1
      Status = a2_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-4" Then
      Call InitGrid2_4
      Dim a2_4 As CHouseGroup

      Set a2_4 = New CHouseGroup
      a2_4.HOUSE_GROUP_ID = -1
      Status = a2_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-5" Then
      Call InitGrid2_5
      Dim a2_5 As CStatusGroup

      Set a2_5 = New CStatusGroup
      a2_5.STATUS_GROUP_ID = -1
      Status = a2_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-6" Then
      Call InitGrid2_6
      Dim a2_6 As CAgeRange

      Set a2_6 = New CAgeRange
      a2_6.AGE_RANGE_ID = -1
      Status = a2_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 2-7" Then
      Call InitGrid2_7
      Dim a2_7 As CStatusType

      Set a2_7 = New CStatusType
      a2_7.STATUS_TYPE_ID = -1
      Status = a2_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a2_7 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-1" Then
      Call InitGrid3_1
      Dim a3_1 As CCountry

      Set a3_1 = New CCountry
      a3_1.COUNTRY_ID = -1
      Status = a3_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-2" Then
      Call InitGrid3_2
      Dim a3_2 As CCustomerGrade

      Set a3_2 = New CCustomerGrade
      a3_2.CSTGRADE_ID = -1
      Status = a3_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-3" Then
      Call InitGrid3_3
      Dim a3_3 As CCustomerType

      Set a3_3 = New CCustomerType
      a3_3.CSTTYPE_ID = -1
      Status = a3_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-4" Then
      Call InitGrid3_4
      Dim a3_4 As CSupplierGrade

      Set a3_4 = New CSupplierGrade
      a3_4.SUPPLIER_GRADE_ID = -1
      Status = a3_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-5" Then
      Call InitGrid3_5
      Dim a3_5 As CSupplierType

      Set a3_5 = New CSupplierType
      a3_5.SUPPLIER_TYPE_ID = -1
      Status = a3_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-6" Then
      Call InitGrid3_6
      Dim a3_6 As CSupplierStatus

      Set a3_6 = New CSupplierStatus
      a3_6.SUPPLIER_STATUS_ID = -1
      Status = a3_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 3-7" Then
      Call InitGrid3_7
      Dim a3_7 As CEmpPosition

      Set a3_7 = New CEmpPosition
      a3_7.POSITION_ID = -1
      Status = a3_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a3_7 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-1" Then
      Call InitGrid4_1
      Dim a4_1 As CExpenseType
      
      Set a4_1 = New CExpenseType
      a4_1.EXPENSE_TYPE_ID = -1
      Status = a4_1.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_1 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-2" Then
      Call InitGrid4_2
      Dim a4_2 As CDocumentType

      Set a4_2 = New CDocumentType
      a4_2.DOCUMENT_TYPE_ID = -1
      Status = a4_2.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_2 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-3" Then
      Call InitGrid4_3
      Dim a4_3 As CBank

      Set a4_3 = New CBank
      a4_3.BANK_ID = -1
      Status = a4_3.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_3 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-4" Then
      Call InitGrid4_4
      Dim a4_4 As CBankBranch

      Set a4_4 = New CBankBranch
      a4_4.BBRANCH_ID = -1
      Status = a4_4.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_4 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-5" Then
      Call InitGrid4_5
      Dim a4_5 As CRegion

      Set a4_5 = New CRegion
      a4_5.REGION_ID = -1
      Status = a4_5.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_5 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-6" Then
      Call InitGrid4_6
      Dim a4_6 As CRevenueType

      Set a4_6 = New CRevenueType
      a4_6.REVENUE_TYPE_ID = -1
      Status = a4_6.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_6 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-7" Then
      Call InitGrid4_7
      Dim a4_7 As CCnDnReason

      Set a4_7 = New CCnDnReason
      a4_7.REASON_ID = -1
      Status = a4_7.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_7 = Nothing
   ElseIf Node.Key = ROOT_TREE & " 4-8" Then
      Call InitGrid4_8
      Dim a4_8 As CBankAccount

      Set a4_8 = New CBankAccount
      a4_8.BANK_ACCOUNT_ID = -1
      Status = a4_8.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_8 = Nothing
    ElseIf Node.Key = ROOT_TREE & " 5-1" Then
      Call InitGrid5_1
      Dim a5_1 As CPackageType

      Set a5_1 = New CPackageType
      a5_1.PACKAGE_TYPE_ID = -1
      Status = a5_1.QueryData(1, m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a5_1 = Nothing
      
   ElseIf Node.Key = ROOT_TREE & " 4-9" Then
      Call InitGrid4_9
      Dim a4_9 As CMasterRef

      Set a4_9 = New CMasterRef
      a4_9.KEY_ID = -1
      a4_9.MASTER_AREA = CHEQUE_TYPE
      Status = a4_9.QueryData(m_Rs, ItemCount)
      GridEX1.ItemCount = ItemCount
      GridEX1.Rebind
      
      Set a4_9 = Nothing
   Else
      Call InitGrid0
   End If
End Sub

Private Sub InitGrid4_9()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสประเภทเช็ค")
  
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทเช็ค")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitGrid1_5()
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
   Col.Width = 2220
  Col.Caption = MapText("รหัสกลุ่มอาหาร")
  
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("กลุ่มอาหาร")

   GridEX1.ItemCount = 0
End Sub

Private Sub Form_Resize()
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   pnlHeader.Width = ScaleWidth
   trvMaster.Width = (1 / 3) * ScaleWidth
   trvMaster.HEIGHT = ScaleHeight - pnlHeader.HEIGHT - pnlFooter.HEIGHT
   GridEX1.Left = trvMaster.Width
   GridEX1.Width = ScaleWidth - trvMaster.Width
   GridEX1.HEIGHT = trvMaster.HEIGHT
   pnlFooter.Width = ScaleWidth
   pnlFooter.Top = ScaleHeight - pnlFooter.HEIGHT
   
   cmdExit.Left = ScaleWidth - cmdExit.Width - 20
   
End Sub
