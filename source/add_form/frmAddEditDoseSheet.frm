VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditDoseSheet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   Icon            =   "frmAddEditDoseSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15002
      _Version        =   131073
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   18
         Top             =   7770
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         Begin Threed.SSCommand cmdDelete 
            Height          =   615
            Left            =   4200
            TabIndex        =   15
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   615
            Left            =   30
            TabIndex        =   13
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   615
            Left            =   2115
            TabIndex        =   14
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdCancel 
            Cancel          =   -1  'True
            Height          =   615
            Left            =   9765
            TabIndex        =   17
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdOK 
            Height          =   615
            Left            =   7680
            TabIndex        =   16
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin prjBoonmeeGraph.uctlDate uctlRecordDate 
         Height          =   435
         Left            =   1320
         TabIndex        =   1
         Top             =   1110
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjBoonmeeGraph.uctlTextBox txtPatient 
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   30
         TabIndex        =   11
         Top             =   4080
         Width           =   11805
         _ExtentX        =   20823
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
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   1244
         _Version        =   131073
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
                  Picture         =   "frmAddEditDoseSheet.frx":014A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":0464
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":0D3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":34F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":3DCA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":46A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":4F7E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":5858
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":6132
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":6A0C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":6E5E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":7738
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":8012
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":88EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":91C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":9618
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":9A6A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":9BC4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":A49E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":AD78
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":B652
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":B96C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":C246
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":CF20
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":D7FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":E0D4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":E9AE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditDoseSheet.frx":F288
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin Threed.SSFrame fraDrug 
         Height          =   3165
         Left            =   0
         TabIndex        =   22
         Top             =   4620
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   5583
         _Version        =   131073
         Begin GridEX20.GridEX GridEX1 
            Height          =   3135
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   11865
            _ExtentX        =   20929
            _ExtentY        =   5530
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
            FormatStyle(1)  =   "frmAddEditDoseSheet.frx":FB62
            FormatStyle(2)  =   "frmAddEditDoseSheet.frx":FCB6
            FormatStyle(3)  =   "frmAddEditDoseSheet.frx":FD66
            FormatStyle(4)  =   "frmAddEditDoseSheet.frx":FE1A
            FormatStyle(5)  =   "frmAddEditDoseSheet.frx":FEF2
            ImageCount      =   0
            PrinterProperties=   "frmAddEditDoseSheet.frx":FFAA
         End
      End
      Begin prjBoonmeeGraph.uctlTextBox txtWeight 
         Height          =   405
         Left            =   1320
         TabIndex        =   4
         Top             =   1980
         Width           =   1395
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjBoonmeeGraph.uctlTextBox txtKS 
         Height          =   405
         Left            =   4440
         TabIndex        =   5
         Top             =   1980
         Width           =   1395
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjBoonmeeGraph.uctlTextBox txtCD4 
         Height          =   405
         Left            =   7740
         TabIndex        =   6
         Top             =   1980
         Width           =   1395
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjBoonmeeGraph.uctlTextBox txtVL 
         Height          =   405
         Left            =   1320
         TabIndex        =   7
         Top             =   2400
         Width           =   1395
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjBoonmeeGraph.uctlTextBox txtLog10 
         Height          =   405
         Left            =   4440
         TabIndex        =   8
         Top             =   2400
         Width           =   1395
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjBoonmeeGraph.uctlTextBox txtNote1 
         Height          =   405
         Left            =   1320
         TabIndex        =   9
         Top             =   2820
         Width           =   10185
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjBoonmeeGraph.uctlTextBox txtNote2 
         Height          =   405
         Left            =   1320
         TabIndex        =   10
         Top             =   3240
         Width           =   10185
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin VB.Label lblNote2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   3300
         Width           =   1125
      End
      Begin VB.Label lblNote1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   2910
         Width           =   1125
      End
      Begin VB.Label lblLog10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   3270
         TabIndex        =   31
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblVL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   150
         TabIndex        =   30
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblCD4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   6570
         TabIndex        =   29
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label lblUnit2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   5910
         TabIndex        =   28
         Top             =   2100
         Width           =   315
      End
      Begin VB.Label lblKS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   3270
         TabIndex        =   27
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label lblUnit1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   2790
         TabIndex        =   26
         Top             =   2100
         Width           =   315
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   150
         TabIndex        =   25
         Top             =   2100
         Width           =   1095
      End
      Begin Threed.SSCommand cmdCustomer 
         Height          =   435
         Left            =   8610
         TabIndex        =   3
         Top             =   1560
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   767
         _Version        =   131073
      End
      Begin VB.Label lblPatient 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   150
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblRecordDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1230
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmAddEditDoseSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
'Private m_Customer As CCustomer

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private m_DoseSheet As CDoseSheet

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2550
   Col.Caption = "ชื่อยา"

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1695
   Col.Caption = "จำนวน"
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2145
   Col.Caption = "ครั้งละ (เม็ด)"
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2160
   Col.Caption = "จำนวนครั้ง/วัน"
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 2760
   Col.Caption = "เวลารับประทาน"
End Sub

Private Sub InitFormLayout()
   pnlHeader.Caption = HeaderText
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlFooter.BackColor = GLB_FORM_COLOR
   
   Call InitGrid
   
   Call InitNormalLabel(lblRecordDate, "วันที่บันทึก")
   Call InitNormalLabel(lblPatient, "ลูกค้า")
   Call txtPatient.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtPatient.Enabled = False
   
   Call InitNormalLabel(lblWeight, "น้ำหนัก")
   Call InitNormalLabel(lblUnit1, "ก.ก.")
   Call InitNormalLabel(lblKS, "KS")
   Call InitNormalLabel(lblUnit2, "%")
   Call InitNormalLabel(lblCD4, "CD4")
   Call InitNormalLabel(lblVL, "VL")
   Call InitNormalLabel(lblLog10, "Log 10")
   Call InitNormalLabel(lblNote1, "อื่น ๆ 1")
   Call InitNormalLabel(lblNote2, "อื่น ๆ 2")

   Call txtWeight.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtKS.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtCD4.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtVL.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtLog10.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtNote1.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtNote2.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

   Call InitMainButton(cmdAdd, "เพิ่ม (F7)")
   Call InitMainButton(cmdEdit, "แก้ไข (F3)")
   Call InitMainButton(cmdDelete, "ลบ (F6)")
   
   Call InitMainButton(cmdOK, "ตกลง (F2)")
   Call InitMainButton(cmdCancel, "ยกเลิก (ESC)")
   Call InitMainButton(cmdCustomer, "...")
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = "การใช้ยา"
'   TabStrip1.Tabs.Add().Caption = "อื่น ๆ"
End Sub

Private Sub cboStatus_Click()
   m_HasModify = True
End Sub

Private Sub Check1_Click()
   m_HasModify = True
End Sub

Private Sub Check10_Click()
   m_HasModify = True
End Sub

Private Sub Check11_Click()
   m_HasModify = True
End Sub

Private Sub Check12_Click()
   m_HasModify = True
End Sub

Private Sub Check13_Click()
   m_HasModify = True
End Sub

Private Sub Check14_Click()
   m_HasModify = True
End Sub

Private Sub Check15_Click()
   m_HasModify = True
End Sub

Private Sub Check16_Click()
   m_HasModify = True
End Sub

Private Sub Check17_Click()
   m_HasModify = True
End Sub

Private Sub Check18_Click()
   m_HasModify = True
End Sub

Private Sub Check19_Click()
   m_HasModify = True
End Sub

Private Sub Check2_Click()
   m_HasModify = True
End Sub

Private Sub Check20_Click()
   m_HasModify = True
End Sub

Private Sub Check21_Click()
   m_HasModify = True
End Sub

Private Sub Check22_Click()
   m_HasModify = True
End Sub

Private Sub Check23_Click()
   m_HasModify = True
End Sub

Private Sub Check24_Click()
   m_HasModify = True
End Sub

Private Sub Check25_Click()
   m_HasModify = True
End Sub

Private Sub Check26_Click()
   m_HasModify = True
End Sub

Private Sub Check27_Click()
   m_HasModify = True
End Sub

Private Sub Check28_Click()
   m_HasModify = True
End Sub

Private Sub Check29_Click()
   m_HasModify = True
End Sub

Private Sub Check3_Click()
   m_HasModify = True
End Sub

Private Sub Check30_Click()
   m_HasModify = True
End Sub

Private Sub Check31_Click()
   m_HasModify = True
End Sub

Private Sub Check32_Click()
   m_HasModify = True
End Sub

Private Sub Check33_Click()
   m_HasModify = True
End Sub

Private Sub Check34_Click()
   m_HasModify = True
End Sub

Private Sub Check35_Click()
   m_HasModify = True
End Sub

Private Sub Check36_Click()
   m_HasModify = True
End Sub

Private Sub Check4_Click()
   m_HasModify = True
End Sub

Private Sub Check5_Click()
   m_HasModify = True
End Sub

Private Sub Check6_Click()
   m_HasModify = True
End Sub

Private Sub Check7_Click()
   m_HasModify = True
End Sub

Private Sub Check8_Click()
   m_HasModify = True
End Sub

Private Sub Check9_Click()
   m_HasModify = True
End Sub

Private Sub chkBerk_Click()
   m_HasModify = True
End Sub

Private Sub chkChild_Click()
   m_HasModify = True
End Sub

Private Sub chkHusband_Click()
   m_HasModify = True
End Sub

Private Sub chkNoJob_Click()
   m_HasModify = True
End Sub

Private Sub chkPay_Click()
   m_HasModify = True
End Sub

Private Sub chkWife_Click()
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditDoseItem.TempCollection = m_DoseSheet.DoseSheetItems
      frmAddEditDoseItem.ShowMode = SHOW_ADD
      frmAddEditDoseItem.HeaderText = "เพิ่มรายการใช้ยา"
      Load frmAddEditDoseItem
      frmAddEditDoseItem.Show 1

      OKClick = frmAddEditDoseItem.OKClick

      Unload frmAddEditDoseItem
      Set frmAddEditDoseItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_DoseSheet.DoseSheetItems)
         GridEX1.Rebind
         m_HasModify = True
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Function VerifyControl() As Boolean
   VerifyControl = False
   
   If Not VerifyDate(lblRecordDate, uctlRecordDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPatient, txtPatient, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblWeight, txtWeight, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblKS, txtKS, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCD4, txtCD4, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblVL, txtVL, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblLog10, txtLog10, True) Then
      Exit Function
   End If
   
   
   VerifyControl = True
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("DAILY_DAILY_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("DAILY_DAILY_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
      
   If Not VerifyControl Then
      Exit Function
   End If
               
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_DoseSheet.DOSE_SHEET_ID = ID
   m_DoseSheet.AddEditMode = ShowMode
   m_DoseSheet.RECORD_DATE = uctlRecordDate.ShowDate
   m_DoseSheet.PATIENT_ID = Val(cmdCustomer.Tag)
   m_DoseSheet.WEIGHT = Val(txtWeight.Text)
   m_DoseSheet.KS_PERCENT = Val(txtKS.Text)
   m_DoseSheet.CD4_VALUE = Val(txtCD4.Text)
   m_DoseSheet.VL_VALUE = Val(txtVL.Text)
   m_DoseSheet.LOG10_VALUE = Val(txtLog10.Text)
    m_DoseSheet.NOTE1 = txtNote1.Text
    m_DoseSheet.NOTE2 = txtNote2.Text
    
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditDoseSheet(m_DoseSheet, IsOK, glbErrorLog) Then
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

Private Sub cmdCustomer_Click()
Dim TempID As Long
Dim TempStr As String
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("เลือกข้อมูลลูกค้า", "-", "ดูข้อมูลลูกค้า")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 3 Then
      If Val(cmdCustomer.Tag) <= 0 Then
         glbErrorLog.LocalErrorMsg = "กรุณาทำการเลือกข้อมูลลูกค้าก่อน"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      Call EnableForm(Me, False)
      frmAddEditCustomer.ID = Val(cmdCustomer.Tag)
      frmAddEditCustomer.HeaderText = "ตรวจสอบข้อมูลลูกค้า"
      frmAddEditCustomer.ShowMode = SHOW_VIEW_ONLY
      Load frmAddEditCustomer
      frmAddEditCustomer.Show 1
         
      OKClick = frmAddEditCustomer.OKClick
      
      Unload frmAddEditCustomer
      Set frmAddEditCustomer = Nothing
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmCustomerSearch
   frmCustomerSearch.Show 1
   
   If frmCustomerSearch.OKClick Then
      TempID = frmCustomerSearch.PersonID
      TempStr = frmCustomerSearch.PersonName
   Else
      TempID = Val(cmdCustomer.Tag)
      TempStr = txtPatient.Text
   End If
   
   Unload frmCustomerSearch
   Set frmCustomerSearch = Nothing
   
   cmdCustomer.Tag = TempID
   txtPatient.Text = TempStr
   m_HasModify = True
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
         m_DoseSheet.DoseSheetItems.Remove (ID2)
      Else
         m_DoseSheet.DoseSheetItems.Item(ID2).Flag = "D"
      End If
      GridEX1.ItemCount = CountItem(m_DoseSheet.DoseSheetItems)
      GridEX1.Rebind
      
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim OKClick As Boolean
Dim ID As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditDoseItem.TempCollection = m_DoseSheet.DoseSheetItems
      frmAddEditDoseItem.ID = ID
      frmAddEditDoseItem.ShowMode = SHOW_EDIT
      frmAddEditDoseItem.HeaderText = "แก้ไขรายการใช้ยา"
      Load frmAddEditDoseItem
      frmAddEditDoseItem.Show 1

      OKClick = frmAddEditDoseItem.OKClick

      Unload frmAddEditDoseItem
      Set frmAddEditDoseItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_DoseSheet.DoseSheetItems)
         GridEX1.Rebind
         m_HasModify = True
      End If
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
            
      m_DoseSheet.DOSE_SHEET_ID = ID
      m_DoseSheet.QueryFlag = 1
      If Not glbDaily.QueryDoseSheet(m_DoseSheet, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   Else
      IsOK = True
   End If
   
   If ItemCount > 0 Then
      uctlRecordDate.ShowDate = InternalDateToDate(NVLS(m_Rs("RECORD_DATE"), ""))
      txtPatient.Text = NVLS(m_Rs("NAME"), "") & " " & NVLS(m_Rs("LAST_NAME"), "")
      cmdCustomer.Tag = NVLI(m_Rs("PATIENT_ID"), -1)
      If NVLD(m_Rs("RECORD_WEIGHT"), 0) > 0 Then
         txtWeight.Text = NVLD(m_Rs("RECORD_WEIGHT"), 0)
      End If
      If NVLD(m_Rs("KS_PERCENT"), 0) > 0 Then
         txtKS.Text = NVLD(m_Rs("KS_PERCENT"), 0)
      End If
      If NVLD(m_Rs("RECORD_CD4"), 0) > 0 Then
         txtCD4.Text = NVLD(m_Rs("RECORD_CD4"), 0)
      End If
      If NVLD(m_Rs("VL_VALUE"), 0) > 0 Then
         txtVL.Text = NVLD(m_Rs("VL_VALUE"), 0)
      End If
      If NVLD(m_Rs("LOG10_VALUE"), 0) > 0 Then
         txtLog10.Text = NVLD(m_Rs("LOG10_VALUE"), 0)
      End If
      txtNote1.Text = NVLS(m_Rs("RECORD_NOTE1"), "")
      txtNote2.Text = NVLS(m_Rs("RECORD_NOTE2"), "")
      
      GridEX1.ItemCount = CountItem(m_DoseSheet.DoseSheetItems)
      GridEX1.Rebind
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_DoseSheet.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlRecordDate.ShowDate = Now
      End If
      
      TabStrip1_Click
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_Load()
   Set m_DoseSheet = New CDoseSheet
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_DoseSheet = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub radAllow_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub radUnAllow_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_DoseSheet.DoseSheetItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CDSheetItem
      If m_DoseSheet.DoseSheetItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_DoseSheet.DoseSheetItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.DSHEET_ITEM_ID
      Values(2) = RealIndex
      If CR.DRUG_ID <= 0 Then
         Values(3) = CR.EXTRA_DRUG_NAME
      Else
         Values(3) = CR.DRUG_NAME
      End If
      Values(4) = FormatNumber(CR.TOTAL_AMOUNT)
      Values(5) = FormatNumber(CR.DOSE_AMOUNT)
      Values(6) = FormatNumber(CR.PERIOD_AMOUNT)
      Values(7) = CR.PERIOD_DESC
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   fraDrug.Visible = False
   fraDrug.BackColor = GLB_FORM_COLOR
   
   If TabStrip1.SelectedItem.Index = 1 Then
      fraDrug.Left = 0
      fraDrug.Top = 4620
      fraDrug.Visible = True
   End If
End Sub

Private Sub txtAge_Change()
   m_HasModify = True
End Sub

Private Sub txtCardNo_Change()
   m_HasModify = True
End Sub

Private Sub txtCD4_Change()
   m_HasModify = True
End Sub

Private Sub txtChannel_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtEquivalence_Change()
   m_HasModify = True
End Sub

Private Sub txtExpense1_Change()
   m_HasModify = True
End Sub

Private Sub txtGender_Change()
   m_HasModify = True
End Sub

Private Sub txtHeight_Change()
   m_HasModify = True
End Sub

Private Sub txtHome_Change()
   m_HasModify = True
End Sub

Private Sub txtJob_Change()
   m_HasModify = True
End Sub

Private Sub txtKhate_Change()
   m_HasModify = True
End Sub

Private Sub txtKwang_Change()
   m_HasModify = True
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub txtMoo_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtOther1_Change()
   m_HasModify = True
End Sub

Private Sub txtOther2_Change()
   m_HasModify = True
End Sub

Private Sub txtOther3_Change()
   m_HasModify = True
End Sub

Private Sub txtOther4_Change()
   m_HasModify = True
End Sub

Private Sub txtOther5_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone2_Change()
   m_HasModify = True
End Sub

Private Sub txtPreWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtReason_Change()
   m_HasModify = True
End Sub

Private Sub txtReference_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSalary_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtViral_Change()
   m_HasModify = True
End Sub

Private Sub txtKS_Change()
   m_HasModify = True
End Sub

Private Sub txtLog10_Change()
   m_HasModify = True
End Sub

Private Sub txtNote1_Change()
   m_HasModify = True
End Sub

Private Sub txtNote2_Change()
   m_HasModify = True
End Sub

Private Sub txtVL_Change()
   m_HasModify = True
End Sub

Private Sub txtWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtYearKnow_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlDate1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDate2_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlRegisterDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox10_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox11_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox12_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox13_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox14_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox15_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox16_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox17_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox18_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox19_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox2_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox3_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox4_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox5_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox6_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox7_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox8_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox9_Change()
   m_HasModify = True
End Sub

Private Sub txtPatient_Change()
   m_HasModify = True
End Sub

Private Sub uctlRecordDate_HasChange()
   m_HasModify = True
End Sub
