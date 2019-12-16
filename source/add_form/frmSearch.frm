VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   Icon            =   "frmSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   9975
      _Version        =   131073
      Begin Threed.SSFrame SSFrame2 
         Height          =   2595
         Left            =   -30
         TabIndex        =   21
         Top             =   2370
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   4577
         _Version        =   131073
         Begin prjFarmManagement.uctlTextBox txtHomeNo 
            Height          =   405
            Left            =   2070
            TabIndex        =   6
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtMoo 
            Height          =   405
            Left            =   4650
            TabIndex        =   7
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtSoi 
            Height          =   405
            Left            =   7170
            TabIndex        =   8
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtRoad 
            Height          =   405
            Left            =   2070
            TabIndex        =   9
            Top             =   630
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtKwang 
            Height          =   405
            Left            =   2070
            TabIndex        =   10
            Top             =   1050
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtKhate 
            Height          =   405
            Left            =   2070
            TabIndex        =   11
            Top             =   1470
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtProvince 
            Height          =   405
            Left            =   2070
            TabIndex        =   12
            Top             =   1890
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   714
         End
         Begin VB.Label lblProvince 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   315
            Left            =   90
            TabIndex        =   28
            Top             =   2010
            Width           =   1845
         End
         Begin VB.Label lblKhate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   315
            Left            =   90
            TabIndex        =   27
            Top             =   1590
            Width           =   1845
         End
         Begin VB.Label lblKwang 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   315
            Left            =   90
            TabIndex        =   26
            Top             =   1170
            Width           =   1845
         End
         Begin VB.Label lblRoad 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   315
            Left            =   90
            TabIndex        =   25
            Top             =   750
            Width           =   1845
         End
         Begin VB.Label lblSoi 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   315
            Left            =   6360
            TabIndex        =   24
            Top             =   330
            Width           =   675
         End
         Begin VB.Label lblMoo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   315
            Left            =   3840
            TabIndex        =   23
            Top             =   330
            Width           =   675
         End
         Begin VB.Label lblHomeNo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   330
            Width           =   1845
         End
      End
      Begin prjFarmManagement.uctlTextBox txtName 
         Height          =   405
         Left            =   2040
         TabIndex        =   3
         Top             =   1350
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   714
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   0
         TabIndex        =   15
         Top             =   4920
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   1244
         _Version        =   131073
         Begin Threed.SSCommand cmdOK 
            Height          =   615
            Left            =   2505
            TabIndex        =   13
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
            Left            =   4590
            TabIndex        =   14
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   9135
         _ExtentX        =   16113
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
                  Picture         =   "frmSearch.frx":08CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":0BE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":14BE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":3C70
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":454A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":4E24
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":56FE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":5FD8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":68B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":718C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":75DE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":7EB8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":8792
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":906C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":9946
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":9D98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":A1EA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":A344
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":AC1E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":B4F8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":BDD2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":C0EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":C9C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":D6A0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":DF7A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":E854
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":F12E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSearch.frx":FA08
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin prjFarmManagement.uctlDate uctlDate 
         Height          =   435
         Left            =   2040
         TabIndex        =   1
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLastName 
         Height          =   405
         Left            =   6270
         TabIndex        =   4
         Top             =   1350
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtCardNo 
         Height          =   405
         Left            =   2040
         TabIndex        =   5
         Top             =   1770
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtCode 
         Height          =   405
         Left            =   7110
         TabIndex        =   2
         Top             =   930
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   5940
         TabIndex        =   29
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label lblCardNo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   1890
         Width           =   1845
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   4800
         TabIndex        =   19
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   60
         TabIndex        =   18
         Top             =   1470
         Width           =   1845
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   1050
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKClick As Boolean
Public HeaderText As String
Public SearchRec As Object

Private Sub InitFormLayout()
   pnlHeader.Caption = HeaderText
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   SSFrame1.BackColor = GLB_FORM_COLOR
   SSFrame2.BackColor = GLB_FORM_COLOR
   pnlFooter.BackColor = GLB_FORM_COLOR
   
   Call InitNormalLabel(lblCode, "รหัส")
   Call InitNormalLabel(lblDate, "วันที่")
   Call InitNormalLabel(lblName, "ชื่อ")
   Call InitNormalLabel(lblLastName, "นามสกุล")
   Call InitNormalLabel(lblCardNo, "บัตรประชาชน")
   Call InitNormalLabel(lblHomeNo, "บ้านเลขที่")
   Call InitNormalLabel(lblMoo, "หมู่")
   Call InitNormalLabel(lblSoi, "ซอย")
   Call InitNormalLabel(lblRoad, "ถนน")
   Call InitNormalLabel(lblKwang, "แขวง/ตำบล")
   Call InitNormalLabel(lblKhate, "เขต/อำเภอ")
   Call InitNormalLabel(lblProvince, "จังหวัด")
   
   Call InitMainButton(cmdOK, "ตกลง (F2)")
   Call InitMainButton(cmdCancel, "ยกเลิก (ESC)")
End Sub

Private Sub cmdCancel_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()

   SearchRec.PATIENT_CODE = txtCode.Text
    SearchRec.REGISTER_DATE = uctlDate.ShowDate
   SearchRec.Name = txtName.Text
   SearchRec.LAST_NAME = txtLastName.Text
   SearchRec.HOME_NO1 = txtHomeNo.Text
   SearchRec.MOO1 = txtMoo.Text
   SearchRec.SOI1 = txtSoi.Text
   SearchRec.ROAD1 = txtRoad.Text
   SearchRec.KWANG1 = txtKwang.Text
   SearchRec.KHATE1 = txtKhate.Text
   SearchRec.PROVINCE = txtProvince.Text
   SearchRec.CARD_NO = txtCardNo.Text
   
   OKClick = True
   Unload Me
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

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
