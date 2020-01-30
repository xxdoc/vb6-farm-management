VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSummaryReport 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmSummaryReport.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10995
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   19394
      _Version        =   131073
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   4
         Top             =   7800
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8460
            TabIndex        =   15
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmSummaryReport.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10110
            TabIndex        =   14
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "JasmineUPC"
               Size            =   24
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   2145
         End
         Begin Threed.SSCommand cmdConfig 
            Height          =   525
            Left            =   6810
            TabIndex        =   13
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   615
            Left            =   2160
            TabIndex        =   0
            Top             =   60
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   615
            Left            =   2610
            TabIndex        =   1
            Top             =   60
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11865
         _ExtentX        =   20929
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
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":2ABC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":3398
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":36B4
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
                  Picture         =   "frmSummaryReport.frx":3F8E
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   6915
         Left            =   0
         TabIndex        =   6
         Top             =   870
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   12197
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
      Begin Threed.SSFrame SSFrame2 
         Height          =   6975
         Left            =   4560
         TabIndex        =   7
         Top             =   960
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   12303
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1275
            Left            =   0
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   12
            Top             =   2820
            Visible         =   0   'False
            Width           =   1635
         End
         Begin prjFarmManagement.uctlTextBox txtGeneric 
            Height          =   435
            Index           =   0
            Left            =   3210
            TabIndex        =   9
            Top             =   870
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin VB.ComboBox cboGeneric 
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
            Index           =   0
            Left            =   3210
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   510
            Visible         =   0   'False
            Width           =   3855
         End
         Begin prjFarmManagement.uctlDate uctlGenericDate 
            Height          =   435
            Index           =   0
            Left            =   3210
            TabIndex        =   8
            Top             =   90
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin Threed.SSCheck chkGeneric 
            Height          =   465
            Index           =   0
            Left            =   3210
            TabIndex        =   17
            Top             =   1860
            Visible         =   0   'False
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   820
            _Version        =   131073
            Caption         =   "SSCheck1"
         End
         Begin Threed.SSCommand cmdEntry 
            Height          =   525
            Left            =   3210
            TabIndex        =   16
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin VB.Label lblGeneric 
            Alignment       =   1  'Right Justify
            Height          =   375
            Index           =   0
            Left            =   330
            TabIndex        =   10
            Top             =   210
            Visible         =   0   'False
            Width           =   2805
         End
      End
   End
End
Attribute VB_Name = "frmSummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String

Public HeaderText As String
Public MasterMode As Long

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_Dates As Collection
Private m_CheckBoxes As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_TextLookups As Collection
Private m_CyclePerMonth As Long
Private m_ReportParams As Collection
Private m_FromDate As Date
Private m_ToDate As Date
Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.Name = GLB_FONT
   trvMaster.Font.Size = 14
   If MasterMode = 1 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("รายงานข้อมูลกลุ่มผู้ใช้งาน"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("รายงานข้อมูลผู้ใช้งาน"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("รายงานการล็อคอินสู่ระบบ"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 2 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("รายงานข้อมูลลูกค้า"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1-1", MapText("รายงานข้อมูลลูกค้า เรียงตามจังหวัด"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1-2", MapText("รายงานข้อมูลลูกค้า ที่เคลื่อนไหว"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-2", MapText("รายงานข้อมูลซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-3", MapText("รายงานข้อมูลพนักงาน"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-1", MapText("รายงานใบตรวจสอบวัตถุดิบ (T102)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-2", MapText("รายงานตรวจสอบบิลรับของ (T103)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-3", MapText("รายงานการโอนวัตถุดิบ (T202)"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-4", MapText("รายงานข้อมูลใบโอนวัตถุดิบ (T203)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5", MapText("รายงานยอดคงเหลือ (M401)"), 1, 2)
      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-6", MapText("รายงานยอดคงเหลือรวมทุกคลัง"), 1, 2)
'      Node.Expanded = False
   
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-9", MapText("รายงานสรุปยอดคงเหลือรวม (M408)"), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-11", MapText("รายงาน STOCK CARD วัตถุดิบ (M405.1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-12", MapText("รายงานสรุป STOCK CARD วัตถุดิบ (M405.2)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-11-1", MapText("รายงาน STOCK CARD วัตถุดิบ (M405.3)"), 1, 2)
      Node.Expanded = False
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-12-1", MapText("รายงานสรุป STOCK CARD วัตถุดิบ (M405.4)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-10", MapText("รายงานสรุปยอดคงเหลือวัตถุดิบ (M408)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-13", MapText("รายงานสรุปยอดคงเหลือวัตถุดิบ (M409)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-7", MapText("รายงานการนำเข้าคลัง"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-8", MapText("รายงานการเบิกจากคลัง"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-14", MapText("รายงานยอดคงเหลือรายคลัง (ST001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-15", MapText("บันทึกจำนวน และการเบิกยา วัคซีน สารเคมี (ST002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-7-1", MapText("บันทึกการนำเข้า ยา วัคซีน  อาหาร ประจำฟาร์มแยกตามบริษัทจัดจำหน่าย(ST003)"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-7-2", MapText("บันทึกการนำเข้า ยา วัคซีน ประจำฟาร์มแยกตามบริษัทจัดจำหน่ายและทะเบียนรถ(ST004)"), 1, 2)
'      Node.Expanded = False
      
   ElseIf MasterMode = 5 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-15", MapText("Management Report อาหาร (C212)"), 1, 2)
      Node.Expanded = False
            
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-16", MapText("Management Report ประชากรสุกร (C212)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-16-1", MapText("Management Report ประชากรสุกรตามบิลขาย (C212.1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-16-2", MapText("Management Report ประชากรสุกร ตามเดือน (C212.2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-15-1", MapText("รายงานการใช้อาหารตามกลุ่มอาหาร (C213)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-1", MapText("รายงานค่าอาหาร ยา วัคซีน (T302)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-2", MapText("รายงานการเบิกใช้ตามเรือน (T308)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-5", MapText("รายงานการเบิกใช้ตามเรือน-สัปดาห์เกิด (T308)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-3", MapText("รายงานสรุปมูลค่าการเบิกใช้ตามเรือน (T308)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-4", MapText("รายงานสรุปมูลค่าการเบิกใช้ตามกลุ่มโรงเรือน (T308)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-8", MapText("รายงานยอดใช้วัตถุดิบแยกสัปดาห์เกิด (T311)"), 1, 2)
      Node.Expanded = False
      '5-5
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-13", MapText("รายงานสรุปสุกรคลอดรวม (T403)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-6", MapText("รายงานประชากรรายวัน ณ วันที่ (M210)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-14", MapText("รายงานสรุปประชากรรายวัน ณ วันที่ (M210.1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-9", MapText("รายงาน STOCK CARD สุกร (M203)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-10", MapText("รายงานสรุป STOCK CARD สุกร (M203)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-11", MapText("รายงานการเคลื่อนไหวสุกรตามเรือน (M208)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-12", MapText("รายงานการเคลื่อนไหวสุกรทั้งฟาร์ม (M208.1)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-17", MapText("รายงานการเคลื่อนไหวพ่อแม่สุกร (M209)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-18", MapText("รายงาน INTAKE ตามเบอร์อาหาร (M210)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-19", MapText("รายงาน INTAKE ตามสัปดาห์เกิด/เบอร์อาหาร (M211)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-20", MapText("รายงาน INTAKE ตามสัปดาห์เกิด/เบอร์อาหาร ตามปีเกิด (M212)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-21", MapText("รายงานสรุปการใช้รายวัน (M213)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 6 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " A-1", MapText("รายงานการขาย"), 3, 3)
      Node.Expanded = False

         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-1", MapText("รายงานยอดขายสุกรฟาร์ม (T706)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-2", MapText("รายงานการขายสุกรตามวัน (T709)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-2-1", MapText("รายงานการขายสุกรตามวัน (T709-1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3", MapText("รายงานการขายสุกรตามประเภท (T709.1)"), 1, 2)
         Node.Expanded = False
            
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-1", MapText("รายงานการขายสุกรตามวันแสดงอายุ (T710)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-2", MapText("รายงานการขายสุกรตามอายุ (T711)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-3", MapText("รายงานการขายสุกรบันทึกมือ (T712)"), 1, 2)
         Node.Expanded = False

         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-4", MapText("รายงานการขายสินค้าอื่น ๆ ตามวัน (T713)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-5", MapText("รายงานการขายสินค้าอื่น ๆ ตามประเภท (T714)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-6", MapText("รายงานรายรับอย่างอื่น (T715)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-7", MapText("รายงานสรุปขายสินค้าอื่น ๆ รายลูกค้า (T716)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-19", MapText("รายงานสรุปขายสินค้าอื่น ๆ รายสินค้า (T716-1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-8", MapText("รายงานรายรับอื่น ๆ ตามวัน (T717)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-9", MapText("รายงานรายรับอื่น ๆ ตามประเภท (T718)"), 1, 2)
         Node.Expanded = False

         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-10", MapText("รายงานสรุปรายรับอื่น ๆ รายลูกค้า (T719)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-11", MapText("รายงานสรุปขายสุกรแยกตามสัปดาห์เกิด (T720)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-12", MapText("รายงานสรุปขายสุกรแยกตามรุ่น (T721)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-12-1", MapText("รายงานสรุปขายสุกรแยกตามรุ่น (เอกสาร H) (T721-1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-13", MapText("รายงานการขายแยกตามลูกค้า (T722)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-15", MapText("รายงานการขายแยกตามลูกค้า/ราคา (ละเอียด) (T723)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-16", MapText("รายงานการขายแยกตามลูกค้า/ราคาขาย (สรุป) (T723-1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-26", MapText("รายงานการขายแยกตามลูกค้า KeyAccount/ราคาขาย (สรุป) (T723-1-1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-16-1", MapText("รายงานการขายแยกตามลูกค้า/ราคาขายกรุ๊ปเป็นช่วง (สรุป) (T723-2)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-16-2", MapText("รายงานการขายแยกตาม พนักงานขาย ลูกค้า (T723-3)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-14", MapText("รายงานการขายตามสัปดาห์เกิด-อายุ (T724)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-17", MapText("รายงานการขายตามอายุ ช่วงน้ำหนัก (T725)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-17-1", MapText("รายงานสรุปจำนวนและอายุสุกรที่ขาย (T725-1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-18", MapText("รายงานการขายตามช่วง นน สถานะ แยกตามลูกค้า (T726)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-21", MapText("รายงานขาย เปรียบเทียบ รับชำระ กลุ่มสินค้า (T728)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-22", MapText("รายงานสรุปขายสินค้าทั้งหมด ประเภทสินค้า เลขที่เอกสาร(T729)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-23", MapText("รายงานสรุปขายสินค้าทั้งหมด ประเภทสินค้า เดือนปี(T730)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-24", MapText("รายงานรายละเอียดการขายสินค้าอื่น ๆ รายบิล(T731)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-1", tvwChild, ROOT_TREE & " 6-3-25", MapText("รายงานยอดขายสุกร สถานะ ลูกค้า แยกตาม เดือนปี(T732)"), 1, 2)
         Node.Expanded = False
         
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " A-2", MapText("รายงานต้นทุน"), 3, 3)
      Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-4", MapText("รายงานการเคลื่อนไหวต้นทุน (C205)"), 1, 2)
         Node.Expanded = False

         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-4-1", MapText("รายงานสรุปการเคลื่อนไหวต้นทุน (C205.1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-5", MapText("รายงานต้นทุนคงเหลือเรือนขาย (C206)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-6", MapText("รายงานต้นทุนคงเหลือทั่วไป (C207)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-7", MapText("รายงานต้นทุนคงเหลือแสดงโรงเรือน (C208)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-8", MapText("รายงานต้นทุนขาย (C209)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-9", MapText("รายงานต้นทุนขายตามสถานะ (C210)"), 1, 2)
         Node.Expanded = False
      
   '      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-12", MapText("รายงานต้นทุนขายแยกตามสถานะ (C211)"), 1, 2)
   '      Node.Expanded = False
   '
   '      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-10", MapText("รายงานต้นทุนคงเหลือตามประเภทสุกร (C212)"), 1, 2)
   '      Node.Expanded = False
   
   '      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-11", MapText("รายงานต้นทุนคงเหลือตามอายุสุกร (C213)"), 1, 2)
   '      Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-14", MapText("รายงานต้นทุนลูกสุกร ณ วันที่ (C215)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-15", MapText("รายงานสรุปการปันค่าใช้จ่าย (C216)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-16", MapText("รายงานสรุปต้นทุนลูกเกิด (C217)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-16-1", MapText("รายงานสรุปต้นทุนลูกเกิด แจกแจงค่าใช้จ่าย(C217.1)"), 1, 2)
         Node.Expanded = False
        
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-17", MapText("รายงานสรุปซื้อสุกร (C218)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-18", MapText("รายงานต้นทุนสุกรทดแทน 1 (C219.1)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-18-1", MapText("รายงานต้นทุนสุกรทดแทน 2 (C219.2)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-19", MapText("รายงานต้นทุนสุกรหมด (C220)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-19-1", MapText("รายงานต้นทุนสุกรคงค้าง (C220.1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-20", MapText("รายงานต้นทุนคงเหลือเลือกกลุ่มโรงเรือน (C221.1)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-2", tvwChild, ROOT_TREE & " 6-20-1", MapText("รายงานต้นทุนคงเหลือ/ตัวเลือกกลุ่มโรงเรือน (C221.2)"), 1, 2)
         Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " A-3", MapText("รายงาน GP"), 3, 3)
      Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-3", tvwChild, ROOT_TREE & " 6-13", MapText("รายงานวิเคราะห์ GP แยกตามสถานะ (C214)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-3", tvwChild, ROOT_TREE & " 6-21", MapText("รายงานวิเคราะห์ GP ของสินค้าแต่ละชนิด1  (C222.1)"), 1, 2)
         Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-3", tvwChild, ROOT_TREE & " 6-21-1", MapText("รายงานวิเคราะห์ GP ของสินค้าและรายได้อื่นๆ  (C222.1.1)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-3", tvwChild, ROOT_TREE & " 6-22", MapText("รายงานวิเคราะห์ GP ของสินค้าแต่ละชนิด 2 (C222.2)"), 1, 2)
         Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " A-4", MapText("รายงานลูกหนี้"), 3, 3)
      Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-23", MapText("รายงานลูกหนี้ค้างชำระรายตัว (AR001)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-24", MapText("รายงานการขายเชื่อสด (AR002)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-26", MapText("รายงานการนำฝากธนาคาร (AR003)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-27", MapText("รายงานการ์ดลูกหนี้ (AR004)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-37", MapText("รายงานสรุปการ์ดลูกหนี้ (AR004-1)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-28", MapText("รายงานวิเคราะห์อายุลูกหนี้ (AR005)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-36", MapText("รายงานสรุปวิเคราะห์อายุลูกหนี้ (AR005-1)"), 1, 2)
         Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-29", MapText("รายงานสรุปลูกหนี้ค้างชำระ (AR006)"), 1, 2)
         Node.Expanded = False
   
'         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-30", MapText("รายงานเงินคงเหลือ (AR007)"), 1, 2)
'         Node.Expanded = False
         
'         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-31", MapText("รายงานการการเคลื่อนไหวหนี้ (AR008)"), 1, 2)
'         Node.Expanded = False

         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-32", MapText("รายงานการเคลื่อนไหวเงินสด (AR009)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-33", MapText("รายงานส่วนต่างการรับชำระ (AR011)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-34", MapText("รายงานเงินคงเหลือ(ใหม่) (AR012)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-35", MapText("รายงานการนำฝากธนาคาร(ใหม่) (AR013)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-38", MapText("รายงานการรับชำระตามประเภทลูกค้า (AR014)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-39", MapText("รายงานการรับชำระเงิน(AR015)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-39-1", MapText("รายงานการรับชำระเงิน+หนี้คงเหลือ(AR015-1)"), 1, 2)
         Node.Expanded = False
         
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-40", MapText("รายงานวิเคาระห์ดิวเช็ค(AR016)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-41", MapText("รายงานการตามหนี้ระหว่างกัน(AR017)"), 1, 2)
         Node.Expanded = False
         
         Set Node = trvMaster.Nodes.Add(ROOT_TREE & " A-4", tvwChild, ROOT_TREE & " 6-42", MapText("รายงานเพิ่มหนี้/ลดหนี้(AR018)"), 1, 2)
         Node.Expanded = False
         
         
   ElseIf MasterMode = 8 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-A", MapText("รายงานอาหาร"), 3, 3)
         Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-A", tvwChild, ROOT_TREE & " 8-5", MapText("รายงานจำนวนการใช้อาหาร (FEED001)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-A", tvwChild, ROOT_TREE & " 8-6", MapText("รายงานมูลค่าการใช้อาหาร (FEED002)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-A", tvwChild, ROOT_TREE & " 8-11", MapText("รายงานราคาอาหาร/ยา (FEED003)"), 1, 2)
      Node.Expanded = False
   
         Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-B", MapText("รายงานการผลิต"), 3, 3)
         Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-8", MapText("รายงานการโอนไปเรือนขายตามสถานะ/สัปดาห์เกิด (PD001)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-2", MapText("รายงานสรุปการเคลื่อนไหวสุกร (PD002)"), 1, 2)
      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-3", MapText("รายงาน INTAKE ตามเบอร์อาหาร (PD003)"), 1, 2)
'      Node.Expanded = False
   
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-4", MapText("รายงาน INTAKE ตามสัปดาห์เกิด (PD004)"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-5", MapText("รายงาน INTAKE ตามเบอร์อาหาร/สัปดาห์เกิด (PD005)"), 1, 2)
'      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-8", MapText("รายงานการใช้อาหาร ตามเบอร์อาหาร/สัปดาห์เกิด (PD005.1)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-8-1", MapText("รายงานการมูลค่าอาหาร ตามเบอร์อาหาร/สัปดาห์เกิด (PD005.2)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-6", MapText("รายงานการเกิดตามสัปดาห์เกิด (PD006)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-7", MapText("รายงานอายุสุกรตามสัปดาห์เกิด (PD007)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-9", MapText("รายงานน้ำหนักสุกรตามอายุ(PD008)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-B", tvwChild, ROOT_TREE & " 8-B-10", MapText("รายงานน้ำหนักสุกรเบอร์สุกร(PD009)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-C", MapText("รายงานขาย/รายรับ/รายจ่าย"), 3, 3)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-7", MapText("รายงานค่าใช้จ่ายปันให้กับสุกร (BG003)"), 1, 2)
      Node.Expanded = False
   
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-9", MapText("รายงานน้ำหนักคาดว่าจะขาย (BG005)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-10", MapText("รายงานรายได้จากการขายอื่น ๆ (BG006)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-12", MapText("รายงานจำนวนขายสุกรตามสถานะ (BG008)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-13", MapText("รายงานมูลค่าขายสุกรตามสถานะ (BG009)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-14", MapText("รายงานจำนวนขายสุกรตามสถานะ/สัปดาห์เกิด (BG010)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-15", MapText("รายงานมูลค่าขายสุกรตามสถานะ/สัปดาห์เกิด (BG011)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-16", MapText("รายงานน้ำหนักขายสุกรตามสถานะ/สัปดาห์เกิด (BG012)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-17", MapText("รายงานน้ำหนักขายเฉลี่ยตามสถานะ/สัปดาห์เกิด (BG013)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-C", tvwChild, ROOT_TREE & " 8-18", MapText("รายงาน จำนวน น้ำหนัก มูลค่า (BG014)"), 1, 2)
      Node.Expanded = False
      
         Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-D", MapText("รายงานงบกำไรขาดทุน"), 3, 3)
         Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-D", tvwChild, ROOT_TREE & " 8-D-1", MapText("รายงานงบกำไรขาดทุน (PR001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-D", tvwChild, ROOT_TREE & " 8-D-2", MapText("รายงานเงินสดรับจ่าย (PR002)"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-D", tvwChild, ROOT_TREE & " 8-D-3", MapText("รายงาน Cash flow (PR003)"), 1, 2)
'      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-D", tvwChild, ROOT_TREE & " 8-D-4", MapText("รายงานวิเคราะห์ GP แยกตามสถานะ (PR004)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-D", tvwChild, ROOT_TREE & " 8-D-5", MapText("รายงานวิเคราะห์ GP แยกตามสถานะ สัปดาห์เกิด (PR005)"), 1, 2)
      Node.Expanded = False
      
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-1", MapText("รายงานงบกำไรขาดทุน (BG001)"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-2", MapText("รายงานเงินสดรับ (BG002)"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-3", MapText("รายงานราคาขายสุกรตามสถานะ (BG003)"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-4", MapText("รายงานรายรับแยกตามตามสถานะ (BG004)"), 1, 2)
'      Node.Expanded = False


      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-P", MapText("รายงานข้อมูลแบต"), 3, 3)
      Node.Expanded = False
         
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-1", MapText("รายงานแบตสุกรเกิด (ฺBIRTH001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-2", MapText("รายงานแบตอาหาร/ยา (FOOD001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-3", MapText("รายงานแบตการโอน(สูญเสีย) (TRANSFER001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-4", MapText("รายงานแบตการขาย (SALE001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-5", MapText("รายงานแบตราคาอาหาร/ยา (FOOD002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-6", MapText("รายงานแบตยอดยกมา (BALANCE001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-7", MapText("รายงานแบตขายอื่นๆ (REVENUE001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-8", MapText("รายงานแบต%การขาย (PARAM001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-9", MapText("รายงานแบตเปลี่ยนประเภทสุกร (CHANGESTATUS001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-10", MapText("รายงานแบตซื้อสุกร (BUYPIG001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-11", MapText("รายงานแบตปันค่าใช้จ่าย (SHARINGEXPENSE001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-12", MapText("รายงานแบตคุมยอดสุกร (PIGADJITEM001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-13", MapText("รายงานแบต คชจ ขายบริหาร (MANAGEMENTEXPENSE001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-14", MapText("รายงานแบต ยกมา GL (GLE001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 8-P", tvwChild, ROOT_TREE & " 8-P-15", MapText("รายงานแบต G กลับสัตว์ (GLE002)"), 1, 2)
      Node.Expanded = False
      
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
   End If

   Label1.Caption = itemcount
   
   Call EnableForm(Me, True)
End Sub

Private Sub FillReportInput(R As CReportInterface)
Dim C As CReportControl

   Call R.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
      End If
   
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
      End If
   
      If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param2)
         End If
      End If
      
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               End If
            End If
            If C.Param2 = "FROM_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            End If
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
   
   Next C
End Sub

Private Function VerifyReportInput() As Boolean
Dim C As CReportControl
   
   VerifyReportInput = False
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "T") Then
         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
            
            Exit Function
         End If
      End If
   
      If (C.ControlType = "D") Then
         If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
         If trvMaster.SelectedItem.Parent.Key = "Root A-1" Then
            If Not VerifyDateInterval(m_Dates(C.ControlIndex).ShowDate) Then
               Exit Function
            End If
         End If
      End If
   Next C
   VerifyReportInput = True
End Function

Private Sub cboGeneric_Click(Index As Integer)
Dim TempID As Long

   If ((trvMaster.SelectedItem.Key = ROOT_TREE & " 5-18") And (Index = 1)) Or _
      ((trvMaster.SelectedItem.Key = ROOT_TREE & " 5-19") And (Index = 1)) Then
      TempID = cboGeneric(Index).ItemData(Minus2Zero(cboGeneric(Index).ListIndex))
      If TempID > 0 Then
         Call LoadPartType(cboGeneric(Index + 1), , TempID)
      End If
   End If
End Sub

Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long

   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
      
   ReportKey = trvMaster.SelectedItem.Key
   
   Set Rc = New CReportConfig
   Rc.REPORT_KEY = ReportKey
   Call Rc.QueryData(m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If
   
   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = trvMaster.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdEntry_Click()
Dim Key As String
Dim D As CManagementConfig
Dim iCount As Long

   If (trvMaster.SelectedItem Is Nothing) Then
      Exit Sub
   End If
   
   Key = trvMaster.SelectedItem.Key
   
   Set D = New CManagementConfig
   If Key = ROOT_TREE & " 5-15" Then
      D.MANAGEMENT_CONFIG_ID = -1
      Call D.QueryData(m_Rs, iCount)
      
      If m_Rs.EOF Then
         frmMangementConfig.ShowMode = SHOW_ADD
      Else
         Call D.PopulateFromRS(1, m_Rs)
         frmMangementConfig.ShowMode = SHOW_EDIT
         frmMangementConfig.ID = D.MANAGEMENT_CONFIG_ID
      End If
      Set m_ReportParams = Nothing
      Set m_ReportParams = New Collection
      Set frmMangementConfig.ReportParams = m_ReportParams
      frmMangementConfig.HeaderText = trvMaster.SelectedItem.Text
      Load frmMangementConfig
      frmMangementConfig.Show 1
      
      Unload frmMangementConfig
      Set frmMangementConfig = Nothing
   End If
   Set D = Nothing
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String
Dim ClassName As String

   Key = trvMaster.SelectedItem.Key
   Name = trvMaster.SelectedItem.Text
      
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
   If Not (trvMaster.SelectedItem Is Nothing) Then
      Call Report.AddParam(trvMaster.SelectedItem.Text, "REPORT_TEXT")
   End If
   If Key = ROOT_TREE & " 1-1" Then
      Set Report = New CReportAdmin001
      ClassName = "CReportAdmin001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 1-2" Then
      Set Report = New CReportAdmin002
      ClassName = "CReportAdmin002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 1-3" Then
      Set Report = New CReportAdmin003
      ClassName = "CReportAdmin003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-1" Then
      Set Report = New CReportMain001
      ClassName = "CReportMain001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-1-1" Then
      Set Report = New CReportMain001_1
      ClassName = "CReportMain001_1"
      SelectFlag = True
    ElseIf Key = ROOT_TREE & " 3-1-2" Then
      Set Report = New CReportMain001_2
      ClassName = "CReportMain001_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-2" Then
      Set Report = New CReportMain002
      ClassName = "CReportMain002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 3-3" Then
      Set Report = New CReportMain003
      ClassName = "CReportMain003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-1" Then
      Set Report = New CReportInventory001
      ClassName = "CReportInventory001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-2" Then
      Set Report = New CReportInventory002
      ClassName = "CReportInventory002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-3" Then
      Set Report = New CReportInventory003
      ClassName = "CReportInventory003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-4" Then
      Set Report = New CReportInventory004
      ClassName = "CReportInventory004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-5" Then
      Set Report = New CReportInventory019
      ClassName = "CReportInventory019"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-6" Then
      Set Report = New CReportInventory006
      ClassName = "CReportInventory006"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-7" Then
      Set Report = New CReportInventory007
      ClassName = "CReportInventory007"
      SelectFlag = True
      
      ElseIf Key = ROOT_TREE & " 4-7-1" Then
      Set Report = New CReportInventory007_1
      ClassName = "CReportInventory007_1"
      SelectFlag = True
      
       ElseIf Key = ROOT_TREE & " 4-7-2" Then
      Set Report = New CReportInventory007_2
      ClassName = "CReportInventory007_2"
      SelectFlag = True
      
   ElseIf Key = ROOT_TREE & " 4-8" Then
      Set Report = New CReportInventory008
      ClassName = "CReportInventory008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-9" Then
      Set Report = New CReportInventory010
      ClassName = "CReportInventory010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-10" Then
      Set Report = New CReportInventory016_2 'CReportInventory016
      ClassName = "CReportInventory016_2"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-11" Then
      Set Report = New CReportInventory017_8
      ClassName = "CReportInventory017_8"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-11-1" Then
      Set Report = New CReportInventory017_6
      ClassName = "CReportInventory017_6"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12" Then
      Set Report = New CReportInventory017_7
      ClassName = "CReportInventory017_7"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-12-1" Then
      Set Report = New CReportInventory018_2
      ClassName = "CReportInventory018_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-13" Then
      Set Report = New CReportInventory016_2
      ClassName = "CReportInventory016_2"
      Call Report.AddParam(2, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 4-14" Then
      Set Report = New CReportInventory051
      ClassName = "CReportInventory051"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 4-15" Then
      Set Report = New CReportInventory069
      ClassName = "CReportInventory069"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-1" Then
      Set Report = New CReportInventory009
      ClassName = "CReportInventory009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-2" Then
      Set Report = New CReportInventory013
      ClassName = "CReportInventory013"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-3" Then
      Set Report = New CReportInventory012
      ClassName = "CReportInventory012"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-4" Then
      Set Report = New CReportInventory011
      ClassName = "CReportInventory011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-5" Then
      Set Report = New CReportInventory014
      ClassName = "CReportInventory014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-6" Then
      Set Report = New CReportInventory026
      ClassName = "CReportInventory026"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-8" Then
      Set Report = New CReportInventory020
      ClassName = "CReportInventory020"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-9" Then
      Set Report = New CReportInventory021
      ClassName = "CReportInventory021"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-10" Then
      Set Report = New CReportInventory022
      ClassName = "CReportInventory022"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-11" Then
      Set Report = New CReportInventory023
      ClassName = "CReportInventory023"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-12" Then
      Set Report = New CReportInventory024
      ClassName = "CReportInventory024"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-13" Then
      Set Report = New CReportInventory025
      ClassName = "CReportInventory025"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-14" Then
      Set Report = New CReportInventory027
      ClassName = "CReportInventory027"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-15" Then
      If m_ReportParams.Count <= 0 Then
         glbErrorLog.LocalErrorMsg = "กรุณาทำการป้อนค่าให้เรียบร้อยก่อน"
         Call glbErrorLog.ShowUserError

         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Set Report = New CReportInventory028
      ClassName = "CReportInventory028"
      Dim Yg As CYGroup
      For Each Yg In m_ReportParams
         Call Report.AddParam(Yg.Value, Yg.Key)
      Next Yg
      
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-15-1" Then
      Set Report = New CReportInventory065
      ClassName = "CReportInventory065"
      
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-16" Then
      Set Report = New CReportInventory029
      ClassName = "CReportInventory029"
      SelectFlag = True
      Call Report.AddParam(1, "MODE")
   ElseIf Key = ROOT_TREE & " 5-16-1" Then
      Set Report = New CReportInventory029
      ClassName = "CReportInventory029"
      SelectFlag = True
      Call Report.AddParam(2, "MODE")
   ElseIf Key = ROOT_TREE & " 5-16-2" Then
      Set Report = New CReportInventory029_1
      ClassName = "CReportInventory029_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-17" Then
      Set Report = New CReportInventory049
      ClassName = "CReportInventory049"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-18" Then
      Set Report = New CReportInventory061
      ClassName = "CReportInventory061"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-19" Then
      Set Report = New CReportInventory063
      ClassName = "CReportInventory063"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-20" Then
      Set Report = New CReportInventory064
      ClassName = "CReportInventory064"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 5-21" Then
      Set Report = New CReportInventory066
      ClassName = "CReportInventory066"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-1" Then
      Set Report = New CReportInventory030
      ClassName = "CReportInventory030"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-2" Then
      Set Report = New CReportInventory031
      ClassName = "CReportInventory031"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-2-1" Then
      Set Report = New CReportInventory031_1
      ClassName = "CReportInventory031_1"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " 6-3" Then
      Set Report = New CReportInventory032
      ClassName = "CReportInventory032"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-1" Then
      Set Report = New CReportInventory032_1
      ClassName = "CReportInventory032_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-2" Then
      Set Report = New CReportInventory032_2
      ClassName = "CReportInventory032_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-3" Then
      Set Report = New CReportSell001
      ClassName = "CReportSell001"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-4" Then
      Set Report = New CReportInventory054
      ClassName = "CReportInventory054"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-5" Then
      Set Report = New CReportInventory055
      ClassName = "CReportInventory055"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-6" Then
      Set Report = New CReportSell002
      ClassName = "CReportSell002"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-7" Then
      Set Report = New CReportSell003
      ClassName = "CReportSell003"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-8" Then
      Set Report = New CReportSell004
      ClassName = "CReportSell004"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-9" Then
      Set Report = New CReportSell005
      ClassName = "CReportSell005"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-10" Then
      Set Report = New CReportSell007
      ClassName = "CReportSell007"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-11" Then
      Set Report = New CReportInventory057
      ClassName = "CReportInventory057"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-12" Then
      Set Report = New CReportInventory058
      ClassName = "CReportInventory058"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-12-1" Then
      Set Report = New CReportInventory058_1
      ClassName = "CReportInventory058_1"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-13" Then
      Set Report = New CReportInventory059
      ClassName = "CReportInventory059"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-14" Then
      Set Report = New CReportInventory060
      ClassName = "CReportInventory060"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-15" Then
      Set Report = New CReportInventory062
      ClassName = "CReportInventory062"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-16" Then
      Set Report = New CReportInventory062_1
      ClassName = "CReportInventory062_1"
      SelectFlag = True
      ElseIf Key = ROOT_TREE & " 6-3-16-1" Then
      Set Report = New CReportInventory062_2
      ClassName = "CReportInventory062_2"
      SelectFlag = True
  ElseIf Key = ROOT_TREE & " 6-3-16-2" Then
      Set Report = New CReportInventory062_3
      ClassName = "CReportInventory062_3"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " 6-3-17" Then
      Set Report = New CReportInventory067
      ClassName = "CReportInventory067"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-17-1" Then
      Set Report = New CReportInventory067_1
      ClassName = "CReportInventory067_1"
      SelectFlag = True
    ElseIf Key = ROOT_TREE & " 6-3-18" Then
      Set Report = New CReportInventory068
      ClassName = "CReportInventory068"
      SelectFlag = True
      
    ElseIf Key = ROOT_TREE & " 6-3-19" Then
      Set Report = New CReportSell008
      ClassName = "CReportSell008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-21" Then
      Set Report = New CReportSell010
      ClassName = "CReportSell010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-22" Then
      Set Report = New CReportSell011
      ClassName = "CReportSell011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-23" Then
      Set Report = New CReportSell014
      ClassName = "CReportSell014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-24" Then
      Set Report = New CReportSell003_1
      ClassName = "CReportSell003_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-25" Then
      Set Report = New CReportInventory030_2
      ClassName = "CReportInventory030_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-3-26" Then
      Set Report = New CReportInventory070
      ClassName = "CReportInventory070"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-4" Then
      Set Report = New CReportInventory033
      ClassName = "CReportInventory033"
      Call Report.AddParam("N", "SUMMARY_FLAG")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-4-1" Then
      Set Report = New CReportInventory033_1
      ClassName = "CReportInventory033_1"
      Call Report.AddParam("Y", "SUMMARY_FLAG")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-5" Then
      Set Report = New CReportInventory034
      ClassName = "CReportInventory034"
      Call Report.AddParam("Y", "SALE_FLAG")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-6" Then
      Set Report = New CReportInventory034
      ClassName = "CReportInventory034"
      Call Report.AddParam("N", "SALE_FLAG")
      Call Report.AddParam("-1", "STATUS_ID")
      Call Report.AddParam("", "STATUS_NAME")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-7" Then
      Set Report = New CReportInventory035
      ClassName = "CReportInventory035"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-8" Then
      Set Report = New CReportInventory036
      ClassName = "CReportInventory036"
      Call Report.AddParam(-1, "STATUS_ID")
      Call Report.AddParam("", "STATUS_NAME")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-9" Then
      Set Report = New CReportInventory036_1
      ClassName = "CReportInventory036_1"
      Call Report.AddParam(-1, "STATUS_ID")
      Call Report.AddParam("", "STATUS_NAME")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-10" Then
      Set Report = New CReportInventory037
      ClassName = "CReportInventory037"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-11" Then
      Set Report = New CReportInventory038
      ClassName = "CReportInventory038"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-12" Then
      Set Report = New CReportInventory039
      ClassName = "CReportInventory039"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-13" Then
      Set Report = New CReportInventory040
      ClassName = "CReportInventory040"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-14" Then
      Set Report = New CReportInventory041
      ClassName = "CReportInventory041"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-15" Then
      Set Report = New CReportInventory042
      ClassName = "CReportInventory042"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-16" Then
      Set Report = New CReportInventory043
      ClassName = "CReportInventory043"
      SelectFlag = True
    ElseIf Key = ROOT_TREE & " 6-16-1" Then
      Set Report = New CReportInventory043_1
      ClassName = "CReportInventory043_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-17" Then
      Set Report = New CReportInventory044
      ClassName = "CReportInventory044"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-18" Then
      Set Report = New CReportInventory045
      ClassName = "CReportInventory045"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-18-1" Then
      Set Report = New CReportInventory045_1
      ClassName = "CReportInventory045_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-19" Then
      Set Report = New CReportInventory046
      ClassName = "CReportInventory046"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-19-1" Then
      Set Report = New CReportInventory046_1
      ClassName = "CReportInventory046_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-20" Then
      Set Report = New CReportInventory047
      ClassName = "CReportInventory047"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-20-1" Then
      Set Report = New CReportInventory047_1
      ClassName = "CReportInventory047_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-21" Then
      Set Report = New CReportInventory048
      ClassName = "CReportInventory048"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-21-1" Then
      Set Report = New CReportInventory048_2
      ClassName = "CReportInventory048_2"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-22" Then
      Set Report = New CReportInventory048_1
      ClassName = "CReportInventory048_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-23" Then
      Set Report = New CReportAR001
      ClassName = "CReportAR001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-24" Then
      Set Report = New CReportAR008
      ClassName = "CReportAR008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-26" Then
      Set Report = New CReportAR003
      ClassName = "CReportAR003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-27" Then
      Set Report = New CReportAR004
      ClassName = "CReportAR004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-28" Then
      Set Report = New CReportAR005
      ClassName = "CReportAR005"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-29" Then
      Set Report = New CReportAR006
      ClassName = "CReportAR006"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-30" Then
'      Set Report = New CReportAR007
'      ClassName = "CReportAR007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-31" Then
      Set Report = New CReportAR008
      ClassName = "CReportAR008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-32" Then
      Set Report = New CReportCash001
      ClassName = "CReportCash001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-33" Then
      Set Report = New CReportAR011
      ClassName = "CReportAR011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-34" Then
      Set Report = New CReportAR012
      ClassName = "CReportAR012"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-35" Then
      Set Report = New CReportAR013
      ClassName = "CReportAR013"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-36" Then
      Set Report = New CReportAR005_1
      ClassName = "CReportAR005_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-37" Then
      Set Report = New CReportAR004_1
      ClassName = "CReportAR004_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-38" Then
      Set Report = New CReportAR014
      ClassName = "CReportAR014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-39" Then
      Set Report = New CReportAR015
      ClassName = "CReportAR015"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-39-1" Then
      Set Report = New CReportAR015_1
      ClassName = "CReportAR015_1"
      SelectFlag = True
   
   ElseIf Key = ROOT_TREE & " 6-40" Then
      Set Report = New CReportAR016
      ClassName = "CReportAR016"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-41" Then
      Set Report = New CReportAR017
      ClassName = "CReportAR017"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 6-42" Then
      Set Report = New CReportAR018
      ClassName = "CReportAR018"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-1" Then
      Set Report = New CReportBudget001
      ClassName = "CReportBudget001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-2" Then
      Set Report = New CReportBudget002
      ClassName = "CReportBudget002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-3" Then
      Set Report = New CReportBudget003
      ClassName = "CReportBudget003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-4" Then
      Set Report = New CReportBudget004
      ClassName = "CReportBudget004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-5" Then
      Set Report = New CReportBudget005
      ClassName = "CReportBudget005"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-6" Then
      Set Report = New CReportBudget005
      ClassName = "CReportBudget005"
      Call Report.AddParam(2, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-7" Then
      Set Report = New CReportBudget007
      ClassName = "CReportBudget007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-8" Then
      Set Report = New CReportBudget008
      ClassName = "CReportBudget008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-2" Then
      Set Report = New CReportBudget013
      ClassName = "CReportBudget013"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-3" Then
      Set Report = New CReportBudget014
      ClassName = "CReportBudget014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-4" Then
      Set Report = New CReportBudget015
      ClassName = "CReportBudget015"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-5" Then
      Set Report = New CReportBudget017
      ClassName = "CReportBudget017"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-6" Then
      Set Report = New CReportBudget016
      ClassName = "CReportBudget016"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-7" Then
      Set Report = New CReportBudget019
      ClassName = "CReportBudget019"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-8" Then
      Set Report = New CReportBudget017
      ClassName = "CReportBudget017"
      Call Report.AddParam(2, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-8-1" Then
      Set Report = New CReportBudget017
      ClassName = "CReportBudget017"
      Call Report.AddParam(3, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-9" Then
      Set Report = New CReportBudget009
      ClassName = "CReportBudget009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-10" Then
      Set Report = New CReportBudget010
      ClassName = "CReportBudget010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-11" Then
      Set Report = New CReportBudget011
      ClassName = "CReportBudget011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-12" Then
      Set Report = New CReportBudget012
      ClassName = "CReportBudget012"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-13" Then
      Set Report = New CReportBudget012
      ClassName = "CReportBudget012"
      Call Report.AddParam(2, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-14" Then
      Set Report = New CReportBudget018
      ClassName = "CReportBudget018"
      Call Report.AddParam(1, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-15" Then
      Set Report = New CReportBudget018
      ClassName = "CReportBudget018"
      Call Report.AddParam(2, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-16" Then
      Set Report = New CReportBudget018
      ClassName = "CReportBudget018"
      Call Report.AddParam(3, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-17" Then
      Set Report = New CReportBudget018
      ClassName = "CReportBudget018"
      Call Report.AddParam(4, "MODE")
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-18" Then
      Set Report = New CReportBudget022
      ClassName = "CReportBudget022"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-D-1" Then
      Set Report = New CReportBudget001
      ClassName = "CReportBudget001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-D-2" Then
      Set Report = New CReportBudget002
      ClassName = "CReportBudget002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-D-3" Then
      Set Report = New CReportBudget001_1
      ClassName = "CReportBudget001_1"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-D-4" Then
      Set Report = New CReportBudget023
      ClassName = "CReportBudget023"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-D-5" Then
      Set Report = New CReportBudget024
      ClassName = "CReportBudget024"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-9" Then
      Set Report = New CReportBudget020
      ClassName = "CReportBudget020"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-B-10" Then
      Set Report = New CReportBudget021
      ClassName = "CReportBudget021"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-1" Then
      Set Report = New CReportParameter001
      ClassName = "CReportParameter001"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-2" Then
      Set Report = New CReportParameter002
      ClassName = "CReportParameter002"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-3" Then
      Set Report = New CReportParameter003
      ClassName = "CReportParameter003"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-4" Then
      Set Report = New CReportParameter004
      ClassName = "CReportParameter004"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-5" Then
      Set Report = New CReportParameter005
      ClassName = "CReportParameter005"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-6" Then
      Set Report = New CReportParameter006
      ClassName = "CReportParameter006"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-7" Then
      Set Report = New CReportParameter007
      ClassName = "CReportParameter007"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-8" Then
      Set Report = New CReportParameter008
      ClassName = "CReportParameter008"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-9" Then
      Set Report = New CReportParameter009
      ClassName = "CReportParameter009"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-10" Then
      Set Report = New CReportParameter010
      ClassName = "CReportParameter010"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-11" Then
      Set Report = New CReportParameter011
      ClassName = "CReportParameter011"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-12" Then
      Set Report = New CReportParameter012
      ClassName = "CReportParameter012"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-13" Then
      Set Report = New CReportParameter013
      ClassName = "CReportParameter013"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-14" Then
      Set Report = New CReportParameter014
      ClassName = "CReportParameter014"
      SelectFlag = True
   ElseIf Key = ROOT_TREE & " 8-P-15" Then
      Set Report = New CReportParameter015
      ClassName = "CReportParameter015"
      SelectFlag = True
   
   End If

   If SelectFlag Then
      If glbParameterObj.Temp = 0 Then
         glbParameterObj.UsedCount = glbParameterObj.UsedCount + 1
         glbParameterObj.Temp = 1
      End If
      
      Call FillReportInput(Report)
      Call Report.AddParam(Name, "REPORT_NAME")
      Call Report.AddParam(Key, "REPORT_KEY")
      
      Set frmReport.ReportObject = Report
      frmReport.ClassName = ClassName
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
   End If
End Sub

Private Sub Form_Activate()
Dim itemcount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
            
      Call QueryData(True)
      m_HasActivate = True
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
'      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
   Set m_ReportParams = Nothing
   Set m_CheckBoxes = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
  ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   SSFrame2.BackColor = GLB_FORM_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdConfig, MapText("ปรับค่า"))
   Call InitMainButton(cmdEntry, MapText("ป้อนค่า"))
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEntry.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitTreeView
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Call InitFormLayout
   
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
   
   Set m_ReportControls = New Collection
   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
   Set m_Combos = New Collection
   Set m_TextLookups = New Collection
   Set m_ReportParams = New Collection
   Set m_CheckBoxes = New Collection
End Sub

Private Sub UnloadAllControl()
Dim I As Long
Dim j As Long

   I = m_Labels.Count
   While I > 0
      Call Unload(m_Labels(I))
      Call m_Labels.Remove(I)
      I = I - 1
   Wend
   
   I = m_Texts.Count
   While I > 0
      Call Unload(m_Texts(I))
      Call m_Texts.Remove(I)
      I = I - 1
   Wend

   I = m_Dates.Count
   While I > 0
      Call Unload(m_Dates(I))
      Call m_Dates.Remove(I)
      I = I - 1
   Wend

   I = m_Combos.Count
   While I > 0
      Call Unload(m_Combos(I))
      Call m_Combos.Remove(I)
      I = I - 1
   Wend
   
   I = m_TextLookups.Count
   While I > 0
      Call Unload(m_TextLookups(I))
      Call m_TextLookups.Remove(I)
      I = I - 1
   Wend
   
   I = m_CheckBoxes.Count
   While I > 0
      Call Unload(m_CheckBoxes(I))
      Call m_CheckBoxes.Remove(I)
      I = I - 1
   Wend
   
   Set m_ReportControls = Nothing
   Set m_ReportControls = New Collection
End Sub

Private Sub ShowControl()
Dim PrevTop As Long
Dim PrevLeft As Long
Dim PrevWidth As Long
Dim CurTop As Long
Dim CurLeft As Long
Dim CurWidth As Long
Dim C As CReportControl

   cmdEntry.Visible = False
   
   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Then
            m_Combos(C.ControlIndex).Left = PrevLeft
            m_Combos(C.ControlIndex).Top = PrevTop
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).HEIGHT
            PrevLeft = m_Combos(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "D" Then
            m_Dates(C.ControlIndex).Left = PrevLeft
            m_Dates(C.ControlIndex).Top = PrevTop
            m_Dates(C.ControlIndex).Width = C.Width
            m_Dates(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).HEIGHT
            PrevLeft = m_Dates(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "T" Then
            m_Texts(C.ControlIndex).Left = PrevLeft
            m_Texts(C.ControlIndex).Left = PrevLeft
            m_Texts(C.ControlIndex).Top = PrevTop
            m_Texts(C.ControlIndex).Width = C.Width
            Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
            m_Texts(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).HEIGHT
            PrevLeft = m_Texts(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "LU" Then
            m_TextLookups(C.ControlIndex).Left = PrevLeft
            m_TextLookups(C.ControlIndex).Top = PrevTop
            m_TextLookups(C.ControlIndex).Width = C.Width
            m_TextLookups(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).HEIGHT
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "CH" Then
            m_CheckBoxes(C.ControlIndex).Left = PrevLeft
            m_CheckBoxes(C.ControlIndex).Top = PrevTop
            m_CheckBoxes(C.ControlIndex).Width = C.Width
            m_CheckBoxes(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_CheckBoxes(C.ControlIndex).Top + m_CheckBoxes(C.ControlIndex).HEIGHT
            PrevLeft = m_CheckBoxes(C.ControlIndex).Left
            PrevWidth = C.Width
            Call InitCheckBox(m_CheckBoxes(C.ControlIndex), C.TextMsg)
         End If
      Else 'Label
            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
            m_Labels(C.ControlIndex).Top = CurTop
            m_Labels(C.ControlIndex).Width = C.Width
            Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
            m_Labels(C.ControlIndex).Visible = True
      End If
   Next C
   
   If cboGeneric.UBound > 1 Then
      cmdEntry.Top = CurTop + cboGeneric.Item(1).HEIGHT
   End If
End Sub

Private Sub LoadComboData()
Dim C As CReportControl

   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-1" Then
            If C.ComboLoadID = 1 Then
               Call InitUserGroupOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadUserGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitUserOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-3" Then
            If C.ComboLoadID = 1 Then
               Call InitLoginOrderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPosition(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-2" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitInventoryDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport4_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-4" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitInventoryDocOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-6" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2, "")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitDocumentType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport4_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
      
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2, "")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitDocumentType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport4_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-9" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-10" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-11" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-11-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-12-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-13" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-14" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2, "")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport4_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-15" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
               cboGeneric(C.ControlIndex).ListIndex = Month(Now)
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport4_15Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex), , "")
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex), , "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-4" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport5_4Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex), , "")
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadYearSeq(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_6Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1, "")
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_6Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1, "")
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport5_6Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call LoadBatch(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-10" Then
            If C.ComboLoadID = 1 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1, "")
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport5_6Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-11" Then
            If C.ComboLoadID = 1 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1, "")
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-12" Then
            If C.ComboLoadID = 1 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1, "")
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_12Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-13" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport5_12Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-14" Then
            If C.ComboLoadID = 1 Then
               Call LoadYearSeq(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport5_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-15" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_15Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-15-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport5_15Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
   
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-16" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 5-16-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport5_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-16-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-17" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport5_17Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-18" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1)
            ElseIf C.ComboLoadID = 5 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport5_18Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-19" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1)
            ElseIf C.ComboLoadID = 5 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport5_18Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-20" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 1)
            ElseIf C.ComboLoadID = 5 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitReport5_18Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 8 Then
               Call LoadYearSeq(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 9 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_3Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2, "")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 2, "")
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadRevenueType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_3_6Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-7" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-19" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_3_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadRevenueType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadRevenueType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-10" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_3_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-11" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_3_11Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-12" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-12-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_3_12Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call LoadBatch(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-13" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-15" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-16" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-26" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-16-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-4" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "") 'N
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadYearSeq(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call LoadBatch(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-4-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "") 'N
            End If
         End If
                  
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-6" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "N")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-7" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-8" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-9" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               'Call LoadProductStatus(m_Combos(C.ControlIndex))
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-10" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_7Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-11" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-12" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitReport6_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-15" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-16" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-16-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-17" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_5Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-18" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-18-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-19" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-19-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "")
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call LoadYearSeq(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-20" Then
            If C.ComboLoadID = 1 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-20-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadHouseGroup(m_Combos(C.ControlIndex), Nothing)
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-21" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-21-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-13" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-22" Then
            If C.ComboLoadID = 1 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadStatusGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_22Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-23" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_23Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-24" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_24Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-26" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-35" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_24Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-27" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-37" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_27Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-28" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-36" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_24Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-29" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_24Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-30" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-34" Then
            If C.ComboLoadID = 1 Then
               Call InitPaymentType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport6_30Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-31" Then
            If C.ComboLoadID = 1 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_24Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-32" Then
            If C.ComboLoadID = 1 Then
               Call LoadBankAccount(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitReportCashTx(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-33" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-39" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-39-1" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-38" Then
            If C.ComboLoadID = 1 Then
               Call InitReport6_33Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-40" Then
            If C.ComboLoadID = 1 Then
               Call InitReportCashTx(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-41" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-42" Then
            If C.ComboLoadID = 1 Then
               Call InitReport6_41Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-1" Then
            If C.ComboLoadID = 1 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-2" Then
            If C.ComboLoadID = 1 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-3" Then
            If C.ComboLoadID = 1 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-4" Then
         If C.ComboLoadID = 1 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      End If
      
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-6" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-7" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-8" Then
         If C.ComboLoadID = 1 Then
            Call LoadLocation(m_Combos(C.ControlIndex), , 1, "N")
         ElseIf C.ComboLoadID = 2 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-2" Then
         If C.ComboLoadID = 1 Then
            Call LoadLocation(m_Combos(C.ControlIndex), , 1, "N")
         ElseIf C.ComboLoadID = 2 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-3" Then
         If C.ComboLoadID = 1 Then
            Call LoadLocation(m_Combos(C.ControlIndex), , 1, "N")
         ElseIf C.ComboLoadID = 2 Then
            Call LoadProductType(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 5 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-4" Then
         If C.ComboLoadID = 1 Then
            Call LoadLocation(m_Combos(C.ControlIndex), , 1, "N")
         ElseIf C.ComboLoadID = 2 Then
            Call LoadProductType(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 5 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-5" Then
         If C.ComboLoadID = 1 Then
            Call LoadLocation(m_Combos(C.ControlIndex), , 1, "N")
         ElseIf C.ComboLoadID = 2 Then
            Call LoadProductType(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 5 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-6" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-7" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-8" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-8-1" Then
         If C.ComboLoadID = 1 Then
            Call LoadLocation(m_Combos(C.ControlIndex), , 1, "N")
         ElseIf C.ComboLoadID = 2 Then
            Call LoadProductType(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 5 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-9" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-10" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-11" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-12" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-13" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-14" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-15" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-16" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-17" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-18" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
            Call LoadProductStatus(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
         
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-D-1" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-D-2" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
   
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-D-3" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
      
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-9" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
            Call LoadProductType(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
      
      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-B-10" Then
         If C.ComboLoadID = 1 Then
            Call LoadBatch(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 2 Then
            Call LoadProductType(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 3 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
         ElseIf C.ComboLoadID = 4 Then
            Call InitOrderType(m_Combos(C.ControlIndex))
         End If
      End If
      
'      If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-1" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-2" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-3" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-4" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-5" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-6" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-7" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-8" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-9" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-10" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-11" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-13" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-14" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-15" Or _
'         trvMaster.SelectedItem.Key = ROOT_TREE & " 8-P-12" Then
'         If C.ComboLoadID = 1 Then
'            Call LoadBatch(m_Combos(C.ControlIndex))
'         ElseIf C.ComboLoadID = 2 Then
'               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
'         ElseIf C.ComboLoadID = 3 Then
'            Call InitOrderType(m_Combos(C.ControlIndex))
'         End If
'      End If
   Next C
   Call EnableForm(Me, True)
End Sub
Private Sub LoadComboDataEx1()
Dim C As CReportControl

   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-7-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
             ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2, "")
             End If
         End If
         
          If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-7-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
             ElseIf C.ComboLoadID = 2 Then
               Call LoadLocation(m_Combos(C.ControlIndex), , 2, "")
             End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-D-4" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-D-5" Then
            If C.ComboLoadID = 1 Then
               Call LoadBatch(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 8-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-3" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-4" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-5" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-6" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-7" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-9" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-10" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-11" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-12" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-13" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-14" _
         Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-15" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-16" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 8-17" Then
            If C.ComboLoadID = 1 Then
               Call LoadBatch(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
                  Call InitReport8_1Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call LoadProductType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 5-21" Then
            If C.ComboLoadID = 1 Then
               Call LoadPartGroup(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-14" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_14Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
               Call LoadYearSeq(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 8 Then
               Call LoadLocation(m_Combos(C.ControlIndex), Nothing, 1, "Y")
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-17" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
          If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-17-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-18" Then
            If C.ComboLoadID = 1 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
                Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
                Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
                Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
                Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 7 Then
                Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 8 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-16-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadCustomerType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitBillingBillSubType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitCommitStatus(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 5 Then
               Call InitReport6_3_13Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 6 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3-25" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call LoadProductStatus(m_Combos(C.ControlIndex))
            End If
         End If
      End If
   Next C
   Call EnableForm(Me, True)
End Sub

Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "")
Dim CboIdx As Long
Dim TxtIdx As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long
Dim C As CReportControl
Dim ChkIdx As Long

   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   LkupIdx = m_TextLookups.Count + 1
   ChkIdx = m_CheckBoxes.Count + 1
   
   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.Add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
   ElseIf ControlType = "C" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.Add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
   ElseIf ControlType = "T" Then
      Load txtGeneric(TxtIdx)
      Call m_Texts.Add(txtGeneric(TxtIdx))
      C.ControlIndex = TxtIdx
   ElseIf ControlType = "D" Then
      Load uctlGenericDate(DateIdx)
      Call m_Dates.Add(uctlGenericDate(DateIdx))
      C.ControlIndex = DateIdx
   
      If DateIdx = 1 Then
         uctlGenericDate(DateIdx).ShowDate = m_FromDate
      ElseIf DateIdx = 2 Then
         uctlGenericDate(DateIdx).ShowDate = m_ToDate
      End If
   ElseIf ControlType = "CH" Then
      Load chkGeneric(ChkIdx)
      Call m_CheckBoxes.Add(chkGeneric(ChkIdx))
      C.ControlIndex = ChkIdx
   End If
   
   C.AllowNull = NullAllow
   C.ControlType = ControlType
   C.Width = Width
   C.TextMsg = TextMsg
   C.Param1 = Param2
   C.Param2 = Param1
   C.ComboLoadID = ComboLoadID
   Call m_ReportControls.Add(C)
   Set C = Nothing
End Sub

Private Sub InitReport1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))
 
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "GROUP_ID", "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "บันทึกลง File", , "PRINT_TO_FILE")
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport3_1_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
    '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "บันทึกลง File", , "PRINT_TO_FILE")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสซัพพลายเออร์"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อซัพพลายเออร์"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทซัพ ฯ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงาน"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อพนักงาน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("นามสกุลพนักงาน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ตำแหน่ง"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
'   Call LoadControl("C", cboGeneric(0).WIDTH, True, "", 1, "GROUP_ID", "GROUP_NAME")
'   Call LoadControl("L", lblGeneric(0).WIDTH, True, GetTextMessage("TEXT-KEY71"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '6 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่เอกสาร"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสซัพพลายเออร์"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DOCUMENT_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เลขที่เอกสาร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))


   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "DOCUMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทเอกสาร"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_7_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   '   '3 =============================
    'Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "SUPPLIER_NAME")
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
    Call LoadControl("L", lblGeneric(0).Width, True, MapText("บริษัทจัดจำหน่าย"))

   '4 =============================

   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))
  
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "LOCATION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))
  
   '5 =============================\
  Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))
   
   
'   6 =============================

   Call LoadControl("CH", chkGeneric(0).Width, True, "แยกบริษัทจัดจำหน่าย ", , "SUPPLIER_NAME_FLAG")
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "DOCUMENT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทเอกสาร"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
'
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
'
'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport4_7_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   '   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
    Call LoadControl("L", lblGeneric(0).Width, True, MapText("บริษัทจัดจำหน่าย"))

   '4 =============================

   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))
  
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "LOCATION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))
  
   '5 =============================\
  Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))
   
'   6 =============================
   Call LoadControl("CH", chkGeneric(0).Width, True, "แยกบริษัทจัดจำหน่าย ", , "SUPPLIER_NAME_FLAG")
   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport4_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
  Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))
         
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "DOCUMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทเอกสาร"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport4_14()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ดึงยอดยกมาจากตารางปรับราคาเฉลี่ย", , "STKCARD_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport4_15()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))

   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))
      
   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
 '  Call LoadControl("CH", chkGeneric(0).Width, True, "ดึงยอดยกมาจาก stock card", , "STKCARD_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_11_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport4_12_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   '3 =============================
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะพ่อแม่", , "PARENT_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP", "PART_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_GROUP", "LOCATION_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมนม", , "INTAKE_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "HOUSE_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "HOUSE_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ณ วันที่"))

   '2 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
'   uctlGenericDate(0).Enable = False

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "YEAR_SEQ_ID", "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีเกิดสุกร"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_14()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ณ วันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ณ วันที่"))
   uctlGenericDate(0).Enable = False

''   3 =============================
'   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "YEAR_SEQ_ID", "YEAR_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีเกิดสุกร"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_15()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "PART_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "HOUSE_GROUP_ID", "HOUSE_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
'   Call LoadControl("CH", chkGeneric(0).Width, True, "คิด intake จากอาหารทั้งหมด", , "INTAKE_FLAG")
   
   Call ShowControl
   cmdEntry.Visible = True
   Call LoadComboData
End Sub

Private Sub InitReport5_15_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "PART_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_16()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "STATUS_GROUP_ID1", "STATUS_GROUP_NAME1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะตาย"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "STATUS_GROUP_ID2", "STATUS_GROUP_NAME2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะคัดทิ้ง"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "PART_GROUP_ID", "PART_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "HOUSE_GROUP_ID", "HOUSE_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
'   Call LoadControl("CH", chkGeneric(0).Width, True, "คิด intake จากอาหารทั้งหมด", , "INTAKE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport5_16_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
        
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
   
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_17()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

    '3 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "NO_STATUS_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไม่รวมสถานะ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_25()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
   

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   Call ShowControl
   Call LoadComboDataEx1
End Sub

Private Sub InitReport6_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือนขาย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "GROUP_STATUS_ID", "GROUP_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือนขาย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานที่จัดเก็บ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))
'
'   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "REVENUE_ID", "REVENUE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทรายรับ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "REVENUE_TYPE", "REVENUE_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทรายรับ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call LoadControl("CH", chkGeneric(0).Width, True, "ยอดยกมานับจากต้นเดือน", , "BALANCE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
'   From-To Sale =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงาน"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงาน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE", "CUSTOMER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงสินค้า", , "NOT_PRODUCT_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงลูกค้า", , "NOT_CUS_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE", "CUSTOMER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
         
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดง GP", , "SHOW_GP")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_12_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
      
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงาน"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงาน"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", , "SUMMARY_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_14()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
        
    '3 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 7, "YEAR_SEQ_ID", "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีเกิดสุกร"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 4, True, "", , "WEEK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สัปดาห์เกิด"))
      
   Call LoadControl("C", cboGeneric(0).Width, True, "", 8, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือนขาย"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงต้นทุน", , "SHOW_COST")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดง GP", , "SHOW_GP")
   
   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport6_3_15()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "GROUP_STATUS_ID", "GROUP_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะสุกร"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "เฉพาะลูกค้าหน้าฟาร์ม", , "TAKE_AWAY_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเวลา", , "SHOW_TIME")
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_16_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", , "SUMMARY_FLAG")
   
   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport6_3_16_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงาน"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงาน"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงสินค้า", , "NOT_PRODUCT_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_19()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
   
   '   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

'   From-To Sale =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงาน"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงาน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE", "CUSTOMER_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดง ลูกค้า", , "NOT_CUS_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดง สินค้า", , "NOT_PART_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดง กลุ่มสินค้า", , "PART_TYPE")

   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_24()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
   
   '   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขวัตถุดิบ"))
   
'   From-To Sale =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "FROM_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสพนักงาน"))
   
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "TO_SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสพนักงาน"))

   
   Call ShowControl
   Call LoadComboData
End Sub


Private Sub InitReport6_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือนขาย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3_21()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
      
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_22()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EXCEPT_PIG_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ยกเว้นสถานะสุกร (03,09,12)"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_3_23()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
   
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EXCEPT_PIG_STATUS")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ยกเว้นสถานะสุกร (03,09,12)"))
      
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุปแนว Column", , "SUMMARY_COLUMN")
         
         
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
         
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 5, "YEAR_SEQ_ID", "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีเกิดสุกร"))
   
   '   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 4, True, "", , "WEEK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สัปดาห์เกิด"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_4_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดงคงเหลือ 0 และ ต้นทุน 0", , "NOT_SHOW_ZERO")
   
   Dim m_ExpenseTypes1 As Collection
   Set m_ExpenseTypes1 = New Collection
   
   Call LoadExpenseType(Nothing, m_ExpenseTypes1)
   Dim TempData As CExpenseType
   For Each TempData In m_ExpenseTypes1
      Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดง " & TempData.EXPENSE_TYPE_NAME, , "EXP" & "-" & TempData.EXPENSE_TYPE_ID)
   Next TempData
   Set m_ExpenseTypes1 = Nothing
   
   Dim m_PartGroup  As Collection
   Dim TempGroup As CPartGroup
   Set m_PartGroup = New Collection
   Call LoadPartGroup(Nothing, m_PartGroup)
   For Each TempGroup In m_PartGroup
      Call LoadControl("CH", chkGeneric(0).Width, True, "ไม่แสดง " & TempGroup.PART_GROUP_NAME, , "PGP" & "-" & TempGroup.PART_GROUP_ID)
   Next TempGroup
   Set m_PartGroup = Nothing
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_6()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

'   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_7()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_GROUP_ID", "STATUS_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   '4 =============================
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมเนื้อสุกร", , "EXCEPTION_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมทิ้งซาก", , "CAPITAL_MOVE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   '4 =============================
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมเนื้อสุกร", , "EXCEPTION_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมทิ้งซาก", , "CAPITAL_MOVE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
            
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "GROUP_STATUS_ID", "GROUP_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไม่รวมสถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_14()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
      
'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_15()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_16()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_17()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_19()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_20()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "HOUSE_GROUP_ID", "HOUSE_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมพ่อแม่", , "PARENT_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_21()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "GROUP_STATUS_ID", "GROUP_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไม่รวมสถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_21_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "GROUP_STATUS_ID", "GROUP_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไม่รวมสถานะสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "GROUP_STATUS_ID_EX", "GROUP_STATUS_NAME_EX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ไม่แสดงรายละเอียด"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ต้นทุนตายขายไม่ได้คิดรวมพ่อแม่", , "PARENT_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "สรุป", , "SUMMARY_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_23()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100

'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   uctlGenericDate(0).Enable = False
   
   '   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมยอดคงเหลือเป็น 0", , "INCLUDE_FLAG")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงวันที่ค้าง", , "SHOW_LEFT_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_26()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100

'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_24()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

'   2 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
'   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   '4 =============================
'   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมยอดคงเหลือเป็น 0", , "INCLUDE_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_31()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   '4 =============================
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมยอดคงเหลือเป็น 0", , "INCLUDE_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
   
   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport8_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PIG_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เบอร์หมู"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
         
   '   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
            
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เบอร์หมู"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8B10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เบอร์หมู"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_2_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 3, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport8_B_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))
            
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 2, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_27()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
    
'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   uctlGenericDate(0).Enable = False
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PIG_TYPE_ID", "PIG_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STATUS_ID", "STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะสุกร"))

   'ไม่สามารถเลือกสถานะรายการได้
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_8()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "PART_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_9()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "STATUS_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะ"))
   
'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, glbUser.SIMULATE_FLAG = "N", "", 7, "BATCH_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
   
   '4 =============================
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงโรงเรือนขาย", , "SALE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "STATUS_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะ"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   '4 =============================
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงโรงเรือนขาย", , "SALE_FLAG")
   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงเฉพาะสรุปโรงเรือน", , "SUMMARY_HOUSE")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_11()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "STATUS_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะ"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงโรงเรือนขาย", , "SALE_FLAG")
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_12()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "STATUS_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสถานะ"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_13()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

''   3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "HOUSE_ID", "HOUSE_NAME")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport5_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

'   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "HOUSE_GROUP_ID")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_28()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งของ"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CREDIT_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่ CREDIT"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_29()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่ส่งของ"))
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DOC_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ส่งของ"))
      
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_PAY_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่ชำระหนี้"))
   
'   '1 =============================
'   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_30()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PAYMENT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("การชำระเงิน"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "ยอดยกมานับจากต้นเดือน", , "BALANCE_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim itemcount As Long
Dim QueryFlag As Boolean
   
   If Node.Key = ROOT_TREE Then
      Exit Sub
   End If
   If LastKey = Node.Key Then
      Exit Sub
   End If
   
   Status = True
   QueryFlag = False
   
   Call UnloadAllControl
   
   cmdOK.Enabled = True
   
   If MasterMode = 1 Then
      If Not VerifyAccessRight("ADMIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MAIN_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
   ElseIf MasterMode = 5 Then
      If Not VerifyAccessRight("PIG_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & trvMaster.SelectedItem.Text, trvMaster.SelectedItem.Text) Then
         cmdOK.Enabled = False                                                                                                                                                               '''''''''
         Exit Sub
      End If
   End If
      
   If Node.Key = ROOT_TREE & " 1-1" Then
      Call InitReport1_1
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitReport1_2
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitReport1_3
   ElseIf Node.Key = ROOT_TREE & " 3-1" Then
      Call InitReport3_1
   ElseIf Node.Key = ROOT_TREE & " 3-1-1" Then
      Call InitReport3_1
    ElseIf Node.Key = ROOT_TREE & " 3-1-2" Then
      Call InitReport3_1_2
   ElseIf Node.Key = ROOT_TREE & " 3-2" Then
      Call InitReport3_2
   ElseIf Node.Key = ROOT_TREE & " 3-3" Then
      Call InitReport3_3
   ElseIf Node.Key = ROOT_TREE & " 4-1" Then
      Call InitReport4_1
   ElseIf Node.Key = ROOT_TREE & " 4-2" Then
      Call InitReport4_2
   ElseIf Node.Key = ROOT_TREE & " 4-3" Then
      Call InitReport4_3
   ElseIf Node.Key = ROOT_TREE & " 4-4" Then
      Call InitReport4_4
   ElseIf Node.Key = ROOT_TREE & " 4-5" Then
      Call InitReport4_5
   ElseIf Node.Key = ROOT_TREE & " 4-6" Then
      Call InitReport4_6
   ElseIf Node.Key = ROOT_TREE & " 4-7" Then
      Call InitReport4_7
  ElseIf Node.Key = ROOT_TREE & " 4-7-1" Then
      Call InitReport4_7_1
   ElseIf Node.Key = ROOT_TREE & " 4-7-2" Then
      Call InitReport4_7_2
   ElseIf Node.Key = ROOT_TREE & " 4-8" Then
      Call InitReport4_8
   ElseIf Node.Key = ROOT_TREE & " 4-9" Then
      Call InitReport4_9
   ElseIf Node.Key = ROOT_TREE & " 4-10" Then
      Call InitReport4_10
   ElseIf Node.Key = ROOT_TREE & " 4-11" Then
      Call InitReport4_11
   ElseIf Node.Key = ROOT_TREE & " 4-12" Then
      Call InitReport4_12
   ElseIf Node.Key = ROOT_TREE & " 4-11-1" Then
      Call InitReport4_11_1
   ElseIf Node.Key = ROOT_TREE & " 4-12-1" Then
      Call InitReport4_12_1
   ElseIf Node.Key = ROOT_TREE & " 4-13" Then
      Call InitReport4_13
   ElseIf Node.Key = ROOT_TREE & " 4-14" Then
      Call InitReport4_14
   ElseIf Node.Key = ROOT_TREE & " 4-15" Then
      Call InitReport4_15
   ElseIf Node.Key = ROOT_TREE & " 5-1" Then
      Call InitReport5_1
   ElseIf Node.Key = ROOT_TREE & " 5-2" Then
      Call InitReport5_2
   ElseIf Node.Key = ROOT_TREE & " 5-3" Then
      Call InitReport5_3
   ElseIf Node.Key = ROOT_TREE & " 5-4" Then
      Call InitReport5_4
   ElseIf Node.Key = ROOT_TREE & " 5-5" Then
      Call InitReport5_5
   ElseIf Node.Key = ROOT_TREE & " 5-6" Then
      Call InitReport5_6
   ElseIf Node.Key = ROOT_TREE & " 5-8" Then
      Call InitReport5_8
   ElseIf Node.Key = ROOT_TREE & " 5-9" Then
      Call InitReport5_9
   ElseIf Node.Key = ROOT_TREE & " 5-10" Then
      Call InitReport5_10
   ElseIf Node.Key = ROOT_TREE & " 5-11" Then
      Call InitReport5_11
   ElseIf Node.Key = ROOT_TREE & " 5-12" Then
      Call InitReport5_12
   ElseIf Node.Key = ROOT_TREE & " 5-13" Then
      Call InitReport5_13
   ElseIf Node.Key = ROOT_TREE & " 5-14" Then
      Call InitReport5_14
   ElseIf Node.Key = ROOT_TREE & " 5-15" Then
      Call InitReport5_15
   ElseIf Node.Key = ROOT_TREE & " 5-15-1" Then
      Call InitReport5_15_1
   ElseIf Node.Key = ROOT_TREE & " 5-16" Or Node.Key = ROOT_TREE & " 5-16-1" Then
      Call InitReport5_16
   ElseIf Node.Key = ROOT_TREE & " 5-16-2" Then
      Call InitReport5_16_2
   ElseIf Node.Key = ROOT_TREE & " 5-17" Then
      Call InitReport5_17
   ElseIf Node.Key = ROOT_TREE & " 5-18" Then
      Call InitReport5_18
   ElseIf Node.Key = ROOT_TREE & " 5-19" Then
      Call InitReport5_18
   ElseIf Node.Key = ROOT_TREE & " 5-20" Then
      Call InitReport5_20
   ElseIf Node.Key = ROOT_TREE & " 5-21" Then
      Call InitReport5_21
   ElseIf Node.Key = ROOT_TREE & " 6-1" Then
      Call InitReport6_1
   ElseIf Node.Key = ROOT_TREE & " 6-2" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-2-1" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-3" Then
      Call InitReport6_3
   ElseIf Node.Key = ROOT_TREE & " 6-3-1" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-3-2" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-3-3" Then
      Call InitReport6_3_3
   ElseIf Node.Key = ROOT_TREE & " 6-3-4" Then
      Call InitReport6_3_4
   ElseIf Node.Key = ROOT_TREE & " 6-3-5" Then
      Call InitReport6_3_4
   ElseIf Node.Key = ROOT_TREE & " 6-3-6" Then
      Call InitReport6_3_6
   ElseIf Node.Key = ROOT_TREE & " 6-3-7" Then
      Call InitReport6_3_7
   ElseIf Node.Key = ROOT_TREE & " 6-3-8" Then
      Call InitReport6_3_8
   ElseIf Node.Key = ROOT_TREE & " 6-3-9" Then
      Call InitReport6_3_8
   ElseIf Node.Key = ROOT_TREE & " 6-3-10" Then
      Call InitReport6_3_10
   ElseIf Node.Key = ROOT_TREE & " 6-3-11" Then
      Call InitReport6_3_11
   ElseIf Node.Key = ROOT_TREE & " 6-3-12" Then
      Call InitReport6_3_12
   ElseIf Node.Key = ROOT_TREE & " 6-3-12-1" Then
      Call InitReport6_3_12_1
   ElseIf Node.Key = ROOT_TREE & " 6-3-13" Then
      Call InitReport6_3_13
   ElseIf Node.Key = ROOT_TREE & " 6-3-14" Then
      Call InitReport6_3_14
   ElseIf Node.Key = ROOT_TREE & " 6-3-15" Then
      Call InitReport6_3_15
   ElseIf Node.Key = ROOT_TREE & " 6-3-16" Then
      Call InitReport6_3_15
   ElseIf Node.Key = ROOT_TREE & " 6-3-16-1" Then
      Call InitReport6_3_16_1
   ElseIf Node.Key = ROOT_TREE & " 6-3-16-2" Then
      Call InitReport6_3_16_2
   ElseIf Node.Key = ROOT_TREE & " 6-3-17" Then
      Call InitReport6_3_17
   ElseIf Node.Key = ROOT_TREE & " 6-3-17-1" Then
      Call InitReport6_3_17_1
    ElseIf Node.Key = ROOT_TREE & " 6-3-18" Then
      Call InitReport6_3_18
  ElseIf Node.Key = ROOT_TREE & " 6-3-19" Then
      Call InitReport6_3_19
   ElseIf Node.Key = ROOT_TREE & " 6-3-21" Then
      Call InitReport6_3_21
   ElseIf Node.Key = ROOT_TREE & " 6-3-22" Then
      Call InitReport6_3_22
   ElseIf Node.Key = ROOT_TREE & " 6-3-23" Then
      Call InitReport6_3_23
   ElseIf Node.Key = ROOT_TREE & " 6-3-24" Then
      Call InitReport6_3_24
   ElseIf Node.Key = ROOT_TREE & " 6-3-25" Then
      Call InitReport6_3_25
   ElseIf Node.Key = ROOT_TREE & " 6-3-26" Then
      Call InitReport6_3_26
   ElseIf Node.Key = ROOT_TREE & " 6-4" Then
      Call InitReport6_4
   ElseIf Node.Key = ROOT_TREE & " 6-4-1" Then
      Call InitReport6_4_1
   ElseIf Node.Key = ROOT_TREE & " 6-5" Then
      Call InitReport6_5
   ElseIf Node.Key = ROOT_TREE & " 6-6" Then
      Call InitReport6_6
   ElseIf Node.Key = ROOT_TREE & " 6-7" Then
      Call InitReport6_7
   ElseIf Node.Key = ROOT_TREE & " 6-8" Then
      Call InitReport6_8
   ElseIf Node.Key = ROOT_TREE & " 6-9" Then
      Call InitReport6_8
   ElseIf Node.Key = ROOT_TREE & " 6-10" Then
      Call InitReport6_10
   ElseIf Node.Key = ROOT_TREE & " 6-11" Then
      Call InitReport6_11
   ElseIf Node.Key = ROOT_TREE & " 6-12" Then
      Call InitReport6_12
   ElseIf Node.Key = ROOT_TREE & " 6-13" Then
      Call InitReport6_21_1_1
   ElseIf Node.Key = ROOT_TREE & " 6-14" Then
      Call InitReport6_14
   ElseIf Node.Key = ROOT_TREE & " 6-15" Then
      Call InitReport6_15
   ElseIf Node.Key = ROOT_TREE & " 6-16" Then
      Call InitReport6_16
    ElseIf Node.Key = ROOT_TREE & " 6-16-1" Then
      Call InitReport6_16
   ElseIf Node.Key = ROOT_TREE & " 6-17" Then
      Call InitReport6_17
   ElseIf Node.Key = ROOT_TREE & " 6-18" Then
      Call InitReport6_18
   ElseIf Node.Key = ROOT_TREE & " 6-18-1" Then
      Call InitReport6_18
   ElseIf Node.Key = ROOT_TREE & " 6-19" Then
      Call InitReport6_19
   ElseIf Node.Key = ROOT_TREE & " 6-19-1" Then
      Call InitReport6_4
   ElseIf Node.Key = ROOT_TREE & " 6-20" Then
      Call InitReport6_20
   ElseIf Node.Key = ROOT_TREE & " 6-20-1" Then
      Call InitReport6_20
   ElseIf Node.Key = ROOT_TREE & " 6-21" Then
      Call InitReport6_21
   ElseIf Node.Key = ROOT_TREE & " 6-21-1" Then
      Call InitReport6_21_1_1
   ElseIf Node.Key = ROOT_TREE & " 6-22" Then
      Call InitReport6_21
   ElseIf Node.Key = ROOT_TREE & " 6-23" Then
      Call InitReport6_23
   ElseIf Node.Key = ROOT_TREE & " 6-24" Then
      Call InitReport6_24
   ElseIf Node.Key = ROOT_TREE & " 6-26" Then
      Call InitReport6_26
   ElseIf Node.Key = ROOT_TREE & " 6-27" Then
      Call InitReport6_27
   ElseIf Node.Key = ROOT_TREE & " 6-28" Then
      Call InitReport6_28
   ElseIf Node.Key = ROOT_TREE & " 6-29" Then
      Call InitReport6_29
   ElseIf Node.Key = ROOT_TREE & " 6-30" Then
      Call InitReport6_30
   ElseIf Node.Key = ROOT_TREE & " 6-31" Then
      Call InitReport6_31
   ElseIf Node.Key = ROOT_TREE & " 6-32" Then
      Call InitReport6_32
   ElseIf Node.Key = ROOT_TREE & " 6-33" Then
      Call InitReport6_33
   ElseIf Node.Key = ROOT_TREE & " 6-34" Then
      Call InitReport6_30
   ElseIf Node.Key = ROOT_TREE & " 6-35" Then
      Call InitReport6_26
   ElseIf Node.Key = ROOT_TREE & " 6-36" Then
      Call InitReport6_28
   ElseIf Node.Key = ROOT_TREE & " 6-37" Then
      Call InitReport6_27
   ElseIf Node.Key = ROOT_TREE & " 6-38" Then
      Call InitReport6_33
   ElseIf Node.Key = ROOT_TREE & " 6-39" Then
      Call InitReport6_33
   ElseIf Node.Key = ROOT_TREE & " 6-39-1" Then
      Call InitReport6_33
   ElseIf Node.Key = ROOT_TREE & " 6-40" Then
      Call InitReport6_40
   ElseIf Node.Key = ROOT_TREE & " 6-41" Then
      Call InitReport6_41
   ElseIf Node.Key = ROOT_TREE & " 6-42" Then
      Call InitReport6_33
   ElseIf Node.Key = ROOT_TREE & " 8-1" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-2" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-3" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-4" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-5" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-6" Then
      Call InitReport8_1
   ElseIf Node.Key = ROOT_TREE & " 8-7" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-8" Then
      Call InitReport8_2
   ElseIf Node.Key = ROOT_TREE & " 8-B-2" Then
      Call InitReport8_B_2
   ElseIf Node.Key = ROOT_TREE & " 8-B-3" Then
      Call InitReport8_2_1
   ElseIf Node.Key = ROOT_TREE & " 8-B-4" Then
      Call InitReport8_2_1
   ElseIf Node.Key = ROOT_TREE & " 8-B-5" Then
      Call InitReport8_2_1
   ElseIf Node.Key = ROOT_TREE & " 8-B-6" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-B-7" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-B-8" Then
      Call InitReport8_2_1
   ElseIf Node.Key = ROOT_TREE & " 8-B-8-1" Then
      Call InitReport8_2_1
   ElseIf Node.Key = ROOT_TREE & " 8-B-9" Then
      Call InitReport8B10
   ElseIf Node.Key = ROOT_TREE & " 8-B-10" Then
      Call InitReport8B10
   ElseIf Node.Key = ROOT_TREE & " 8-9" Then
      Call InitReport8_9
   ElseIf Node.Key = ROOT_TREE & " 8-10" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-11" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-12" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-13" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-14" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-15" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-16" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-17" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-18" Then
      Call InitReport8_18
   ElseIf Node.Key = ROOT_TREE & " 8-D-1" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-D-2" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-D-3" Then
      Call InitReport8_11
   ElseIf Node.Key = ROOT_TREE & " 8-D-4" Then
      Call InitReport8_D4
   ElseIf Node.Key = ROOT_TREE & " 8-D-5" Then
      Call InitReport8_D4
   ElseIf Node.Key = ROOT_TREE & " 8-P-1" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-2" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-3" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-4" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-5" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-6" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-7" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-8" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-9" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-10" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-11" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-12" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-13" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-14" Then
      Call InitReport8_P_1
   ElseIf Node.Key = ROOT_TREE & " 8-P-15" Then
      Call InitReport8_P_1
   End If
End Sub
Private Sub InitReport8_P_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_32()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "BANK_ACCOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สมุดบัญชี"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_33()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   uctlGenericDate(0).Enable = False
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_40()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "PRINT_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่พิมพ์"))

   '8 =============================
   Call LoadControl("T", cboGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_41()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
'   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   uctlGenericDate(0).Enable = False
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport5_20()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP", "PART_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PART_TYPE", "PART_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทวัตถุดิบ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "LOCATION_GROUP", "LOCATION_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มโรงเรือน"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("โรงเรือน"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, True, "", 8, "YEAR_SEQ_ID", "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปีเกิดสุกร"))
   
   '   3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 9, "PIG_TYPE_ID", "PIG_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
      
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 4, True, "", , "WEEK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สัปดาห์เกิด"))
   
   Call LoadControl("CH", chkGeneric(0).Width, True, "รวมนม", , "INTAKE_FLAG")

   Call ShowControl
   Call LoadComboData
End Sub
Private Sub Form_Resize()
   pnlHeader.Width = ScaleWidth
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   trvMaster.Width = ScaleWidth - SSFrame2.Width
   SSFrame2.Left = trvMaster.Width
   trvMaster.HEIGHT = ScaleHeight - pnlHeader.HEIGHT - pnlFooter.HEIGHT
   SSFrame2.HEIGHT = trvMaster.HEIGHT
   pnlFooter.Width = ScaleWidth
   pnlFooter.Top = ScaleHeight - pnlFooter.HEIGHT
   
   cmdExit.Left = ScaleWidth - cmdExit.Width - 20
   cmdOK.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20
   cmdConfig.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20 - cmdConfig.Width - 20
End Sub
Private Sub InitReport8_D4()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ปี"))
      
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "BATCH_ID", "BATCH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("แบต"))
         
   '   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
            
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะขาย"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "PART_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เบอร์หมู"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport5_21()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
      
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PART_GROUP", "PART_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มวัตถุดิบ"))
   
   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport6_3_17()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากน้ำหนัก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงน้ำหนัก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "PERIOD_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงห่างน้ำหนัก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_AGE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากอายุ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_AGE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงอายุ"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))
   
'   '3 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))
'
'   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
'
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงยอดขาย", , "SHOW_PRICE")
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดงต้นทุน", , "SHOW_COST")
'   Call LoadControl("CH", chkGeneric(0).Width, True, "แสดง GP", , "SHOW_GP")
   
   Call ShowControl
   Call LoadComboDataEx1
End Sub
Private Sub InitReport6_3_17_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกร"))
   
   Call ShowControl
   Call LoadComboDataEx1
End Sub

Private Sub InitReport6_3_18()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "PIG_STATUS1", "PIG_STATUS_NAME1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกรตามนน1"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "PIG_STATUS2", "PIG_STATUS_NAME2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกรตามนน2"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_STATUS3", "PIG_STATUS_NAME3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสุกรตามนน2"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "LESS_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("น้ำหนัก น้อยกว่าเท่ากับ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากน้ำหนัก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงน้ำหนัก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "MORE_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("น้ำหนัก มากกว่าเท่ากับ"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "PERIOD_WEIGHT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ช่วงห่างน้ำหนัก"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "PIG_STATUS_IN1", "PIG_STATUS_IN_NAME1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมประเภทสุกร1"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "PIG_STATUS_IN2", "PIG_STATUS_IN_NAME2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมประเภทสุกร2"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "PIG_STATUS_IN3", "PIG_STATUS_IN_NAME3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมประเภทสุกร3"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "PIG_STATUS_IN4", "PIG_STATUS_IN_NAME4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมประเภทสุกร4"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 8, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   Call ShowControl
   Call LoadComboDataEx1
End Sub

Private Sub InitReport6_3_26()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

'   2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   uctlGenericDate(0).Enable = False

   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width \ 2, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "BILL_SUBTYPE", "BILL_SUBTYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "PIG_STATUS", "PIG_STATUS_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะการขาย"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "COMMIT_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("สถานะรายการ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
