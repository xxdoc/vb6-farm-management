VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmWinPricingMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   1305
   ClientTop       =   1065
   ClientWidth     =   11910
   Icon            =   "frmWinPricingMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fraLedger 
      Height          =   4875
      Left            =   4140
      TabIndex        =   30
      Top             =   2340
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8599
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdGeneric 
         Height          =   765
         Left            =   900
         TabIndex        =   63
         Top             =   1240
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCash 
         Height          =   765
         Left            =   900
         TabIndex        =   55
         Top             =   2850
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCapital 
         Height          =   765
         Left            =   900
         TabIndex        =   54
         Top             =   2060
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdLedgerReport 
         Height          =   765
         Left            =   900
         TabIndex        =   32
         Top             =   3645
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSell 
         Height          =   765
         Left            =   900
         TabIndex        =   31
         Top             =   450
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "��Ѻ���§�͡���"
      Height          =   405
      Left            =   3570
      TabIndex        =   53
      Top             =   6900
      Visible         =   0   'False
      Width           =   1305
   End
   Begin Threed.SSFrame fraPig 
      Height          =   5355
      Left            =   11610
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9446
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdPigImport 
         Height          =   765
         Left            =   900
         TabIndex        =   52
         Top             =   1890
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigAdjustment 
         Height          =   765
         Left            =   900
         TabIndex        =   42
         Top             =   3450
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigWeek 
         Height          =   765
         Left            =   900
         TabIndex        =   40
         Top             =   330
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigBirth 
         Height          =   765
         Left            =   900
         TabIndex        =   29
         Top             =   1110
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigTransfer 
         Height          =   765
         Left            =   900
         TabIndex        =   28
         Top             =   2670
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPigReport 
         Height          =   765
         Left            =   900
         TabIndex        =   27
         Top             =   4230
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "��͹�ء�������͹���"
      Height          =   495
      Left            =   3480
      TabIndex        =   51
      Top             =   6300
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command8 
      Caption         =   "�����¹ʶҹ��ء�"
      Height          =   495
      Left            =   3480
      TabIndex        =   50
      Top             =   5790
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��͹�ء�"
      Height          =   495
      Left            =   3480
      TabIndex        =   49
      Top             =   5280
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��͹�ѵ�شԺ"
      Height          =   495
      Left            =   3480
      TabIndex        =   48
      Top             =   4770
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��ԡ�ѵ�شԺ"
      Height          =   495
      Left            =   3480
      TabIndex        =   47
      Top             =   4260
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��ءä�ʹ"
      Height          =   495
      Left            =   3480
      TabIndex        =   46
      Top             =   3750
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command3 
      Caption         =   "㺹�����ѵ�شԺ"
      Height          =   495
      Left            =   3480
      TabIndex        =   45
      Top             =   3240
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "������ Interim"
      Height          =   495
      Left            =   3480
      TabIndex        =   44
      Top             =   2730
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ҧ Interim"
      Height          =   495
      Left            =   3480
      TabIndex        =   43
      Top             =   2220
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":24B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3980
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":425A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":4B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   795
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1402
      _Version        =   131073
      BackStyle       =   1
      Begin VB.Label lblRegPath 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   345
         Left            =   1080
         TabIndex        =   62
         Top             =   480
         Width           =   10005
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   555
         Left            =   9660
         TabIndex        =   37
         Top             =   6390
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   979
         _Version        =   131073
         PictureFrames   =   1
         Picture         =   "frmWinPricingMain.frx":540E
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin VB.Label lblDateTime 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   315
         Left            =   9390
         TabIndex        =   36
         Top             =   30
         Width           =   2505
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   13679
      _Version        =   131073
      BackStyle       =   1
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   4185
         Left            =   240
         TabIndex        =   1
         Top             =   1230
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   7382
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdPasswd 
         Height          =   465
         Left            =   330
         TabIndex        =   39
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   1920
         TabIndex        =   38
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblVersion 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   35
         Top             =   6600
         Width           =   3045
      End
      Begin VB.Label lblUserGroup 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   34
         Top             =   6090
         Width           =   3045
      End
      Begin VB.Label lblUserName 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   33
         Top             =   5580
         Width           =   3045
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   735
      Left            =   3450
      TabIndex        =   3
      Top             =   810
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   1296
      _Version        =   131073
      BackStyle       =   1
   End
   Begin Threed.SSFrame fraMain 
      Height          =   4875
      Left            =   3720
      TabIndex        =   4
      Top             =   7410
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8599
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdMainEmployee 
         Height          =   765
         Left            =   900
         TabIndex        =   9
         Top             =   2820
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainReport 
         Height          =   765
         Left            =   900
         TabIndex        =   8
         Top             =   3600
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainEnterprise 
         Height          =   765
         Left            =   900
         TabIndex        =   7
         Top             =   480
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainCustomer 
         Height          =   765
         Left            =   900
         TabIndex        =   6
         Top             =   1260
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainSupplier 
         Height          =   765
         Left            =   900
         TabIndex        =   5
         Top             =   2040
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraAdmin 
      Height          =   3615
      Left            =   6120
      TabIndex        =   10
      Top             =   7740
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdAdminReport 
         Height          =   765
         Left            =   900
         TabIndex        =   13
         Top             =   2190
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUser 
         Height          =   765
         Left            =   900
         TabIndex        =   12
         Top             =   1410
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUserGroup 
         Height          =   765
         Left            =   900
         TabIndex        =   11
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraMaster 
      Height          =   4875
      Left            =   11520
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8599
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdMaster3 
         Height          =   765
         Left            =   1020
         TabIndex        =   19
         Top             =   2130
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster2 
         Height          =   765
         Left            =   1020
         TabIndex        =   18
         Top             =   1350
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster1 
         Height          =   765
         Left            =   1020
         TabIndex        =   17
         Top             =   570
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster5 
         Height          =   765
         Left            =   1020
         TabIndex        =   16
         Top             =   3660
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMaster4 
         Height          =   765
         Left            =   1020
         TabIndex        =   15
         Top             =   2910
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraInventory 
      Height          =   5625
      Left            =   4290
      TabIndex        =   20
      Top             =   7530
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9922
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdAdjust 
         Height          =   765
         Left            =   900
         TabIndex        =   41
         Top             =   3600
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdTransfer 
         Height          =   765
         Left            =   900
         TabIndex        =   25
         Top             =   2820
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdInventoryReport 
         Height          =   765
         Left            =   900
         TabIndex        =   24
         Top             =   4380
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdRawMatterial 
         Height          =   765
         Left            =   900
         TabIndex        =   23
         Top             =   480
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdImport 
         Height          =   765
         Left            =   900
         TabIndex        =   22
         Top             =   1260
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExport 
         Height          =   765
         Left            =   900
         TabIndex        =   21
         Top             =   2040
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraSimulate 
      Height          =   3615
      Left            =   5040
      TabIndex        =   56
      Top             =   7560
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdParameter 
         Height          =   765
         Left            =   900
         TabIndex        =   59
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdBatch 
         Height          =   765
         Left            =   900
         TabIndex        =   58
         Top             =   1410
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSimulateReport 
         Height          =   765
         Left            =   900
         TabIndex        =   57
         Top             =   2190
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame FraPackage 
      Height          =   2055
      Left            =   5640
      TabIndex        =   60
      Top             =   1920
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdSetCost 
         Height          =   765
         Left            =   900
         TabIndex        =   61
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmWinPricingMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"

Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset
Private m_TableName As String

Public HeaderText As String
Private m_XCollection As CXCollection
Private m_Formula As CFormula
Private m_MustAsk As Boolean
Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String
   
   trvMain.Nodes.Clear
   trvMain.Font.Name = GLB_FONT
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False
   
   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE, MapText("�к������çҹ ������ء�"), 1)
   Node.Expanded = True
   Node.Selected = True
   
   '==
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", MapText("�к������ż����ҹ"), 4, 4)
   Node.Expanded = False
   '==
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("�к���������ѡ"), 2, 2)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("�к���������ǹ��ҧ"), 6, 6)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("�к������ä�ѧ"), 3, 3)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-4", MapText("�к��������ء�"), 9, 9)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-7", MapText("�к���õ���Ҥ��Թ��Ң��"), 10, 10)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-5", MapText("�к������úѭ��"), 8, 8)
   Node.Expanded = False
   
End Sub

Private Sub AddSimulateMenuItem()
Dim Node As Node

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-6", MapText("�к� Simulate"), 10, 10)
   Node.Expanded = False
End Sub

Private Sub InitFormLayout()
   Call InitNormalLabel(lblUsername, MapText("����� : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblUserGroup, MapText("���������� : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblVersion, MapText("�����ѹ : ") & glbParameterObj.Version & " (Interbase) ", RGB(0, 0, 255))
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))

   If Command = "1" Or Command = "" Then
      Call InitNormalLabel(lblRegPath, glbParameterObj.DBFile, RGB(0, 0, 255))
   Else
      Call InitNormalLabel(lblRegPath, glbParameterObj.DBFileAPX, RGB(0, 0, 255))
   End If
   
   lblDateTime.BackStyle = 1
   lblDateTime.BackColor = RGB(255, 255, 255)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdUserGroup.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdUser.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdAdminReport.Picture = LoadPicture(glbParameterObj.MainButton)
      
   cmdMaster1.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster2.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster3.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster4.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMaster5.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdMainEnterprise.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainCustomer.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainSupplier.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainReport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainEmployee.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdRawMatterial.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdImport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdExport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdTransfer.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdAdjust.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdInventoryReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdPigWeek.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigBirth.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigImport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigTransfer.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigAdjustment.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdPigReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
'   cmdBuy.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdSell.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdGeneric.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdCapital.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdCash.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdLedgerReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdParameter.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdBatch.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdSimulateReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   cmdSetCost.Picture = LoadPicture(glbParameterObj.MainButton)
   
   Me.Caption = MapText("�к������çҹ ������ء�")
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitMainButton(cmdUserGroup, MapText("�����š���������ҹ"))
   Call InitMainButton(cmdUser, MapText("�����ż����ҹ"))
   Call InitMainButton(cmdAdminReport, MapText("��§ҹ�����ż����ҹ"))

   Call InitMainButton(cmdMaster1, MapText("��������ѡ��ǹ��ҧ"))
   Call InitMainButton(cmdMaster2, MapText("��������ѡ�к���ѧ"))
   Call InitMainButton(cmdMaster3, MapText("��������ѡ�к��������ء�"))
   Call InitMainButton(cmdMaster4, MapText("��������ѡ�к������úѭ��"))
   Call InitMainButton(cmdMaster5, MapText("��������ѡ�к���õ���Ҥ�"))

   Call InitMainButton(cmdMainEnterprise, MapText("������ͧ��� (�����)"))
   Call InitMainButton(cmdMainCustomer, MapText("�������١���"))
   Call InitMainButton(cmdMainSupplier, MapText("�����ūѾ���������"))
   Call InitMainButton(cmdMainEmployee, MapText("�����ž�ѡ�ҹ"))
   Call InitMainButton(cmdMainReport, MapText("��§ҹ�����š�ҧ"))
   
   Call InitMainButton(cmdRawMatterial, MapText("�������ѵ�شԺ"))
   Call InitMainButton(cmdImport, MapText("�����š���Ѻ����ѵ�شԺ"))
   Call InitMainButton(cmdExport, MapText("�����š���ԡ�ѵ�شԺ"))
   Call InitMainButton(cmdTransfer, MapText("�����š���͹�����ѵ�شԺ"))
   Call InitMainButton(cmdAdjust, MapText("�����š�û�Ѻ�ʹ��ѧ"))
   Call InitMainButton(cmdInventoryReport, MapText("��§ҹ�к���ѧ"))
   
   Call InitMainButton(cmdPigWeek, MapText("�����������ѻ�����Դ�ء�"))
   Call InitMainButton(cmdPigBirth, MapText("�������ءä�ʹ"))
   Call InitMainButton(cmdPigImport, MapText("�����š�ù�����ء�"))
   Call InitMainButton(cmdPigTransfer, MapText("�����š���͹�����ء�"))
   Call InitMainButton(cmdPigAdjustment, MapText("�����š�û�Ѻ�ʹ�ء�"))
   Call InitMainButton(cmdPigReport, MapText("��§ҹ�к��������ء�"))
   
'   Call InitMainButton(cmdBuy, MapText("�к��ҹ���� (��¨���)"))
   Call InitMainButton(cmdSell, MapText("�к��ҹ���"))
   Call InitMainButton(cmdGeneric, MapText("�к��������"))
   Call InitMainButton(cmdCapital, MapText("�к��ҹ�鹷ع"))
   Call InitMainButton(cmdCash, MapText("�Թʴ˹�ҿ����"))
   Call InitMainButton(cmdLedgerReport, MapText("��§ҹ�к��ѭ��"))
   
   Call InitMainButton(cmdParameter, MapText("�����ž���������"))
   Call InitMainButton(cmdBatch, MapText("������ẵ Simulate"))
   Call InitMainButton(cmdSimulateReport, MapText("��§ҹ�к� Simulate"))
   
   Call InitMainButton(cmdSetCost, MapText("�����š�õ���ҤҢ��"))
   
   Call InitMainButton(cmdExit, MapText("�͡"))
   Call InitMainButton(cmdPasswd, MapText("�����"))
   
   Call InitMainTreeview
End Sub

Private Sub cmdAdjust_Click()
   If Not VerifyAccessRight("INVENTORY_ADJUST") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmInventoryDoc4
   frmInventoryDoc4.Show 1

   Unload frmInventoryDoc4
   Set frmInventoryDoc4 = Nothing
End Sub

Private Sub cmdAdminReport_Click()

   If Not VerifyAccessRight("ADMIN_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdAdminReport.Caption
   frmSummaryReport.MasterMode = 1
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdBatch_Click()
   frmBatch.HeaderText = cmdBatch.Caption
   Load frmBatch
   frmBatch.Show 1

   Unload frmBatch
   Set frmBatch = Nothing
End Sub

Private Sub cmdBuy_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("��Ѻ�ͧ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      frmBillingDoc2.DocumentType = 5
   ElseIf lMenuChosen = 3 Then
   ElseIf lMenuChosen = 5 Then
   End If
   Load frmBillingDoc2
   frmBillingDoc2.Show 1

   Unload frmBillingDoc2
   Set frmBillingDoc2 = Nothing
End Sub

Private Sub cmdCapital_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
   
   If Not VerifyAccessRight("LEDGER_COST") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("�ѹ�鹷ع �.�.�.", "-", "������鹷ع", "-", "�Ŵ�鹷ع", "-", "�ѹ�鹷ع�ҡ���������")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not VerifyAccessRight("LEDGER_COST_5", "�ѹ�鹷ع �.�.�.") Then                                                                                     '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmBillingDoc2.DocumentType = 5
      Load frmBillingDoc2
      frmBillingDoc2.Show 1
   
      Unload frmBillingDoc2
      Set frmBillingDoc2 = Nothing
   ElseIf lMenuChosen = 3 Then
      If Not VerifyAccessRight("LEDGER_COST_6", "������鹷ع") Then                                                                                     '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmBillingDoc3.DocumentType = 6
      Load frmBillingDoc3
      frmBillingDoc3.Show 1
   
      Unload frmBillingDoc3
      Set frmBillingDoc3 = Nothing
   ElseIf lMenuChosen = 5 Then
      If Not VerifyAccessRight("LEDGER_COST_7", "�Ŵ�鹷ع") Then                                                                                     '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmBillingDoc3.DocumentType = 7
      Load frmBillingDoc3
      frmBillingDoc3.Show 1
      
      Unload frmBillingDoc3
      Set frmBillingDoc3 = Nothing
   ElseIf lMenuChosen = 7 Then
      If Not VerifyAccessRight("LEDGER_COST_REVENUE", "�ѹ�鹷ع�ҡ���������") Then                                                                                     '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Load frmRevenueCost
      frmRevenueCost.Show 1
      
      Unload frmRevenueCost
      Set frmRevenueCost = Nothing
   End If
End Sub
Private Sub cmdCash_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
   
   If Not VerifyAccessRight("LEDGER_CASH") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("�Թ�ҡ��Ҥ��", "-", "��׹�ѹ��� clearing ��")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   
   
   If lMenuChosen = 1 Then
      frmCashDoc.Area = CASH_DEPOSIT                   '㺹ӽҡ �Թʴ/������
      frmCashDoc.DocumentType = CASH_DEPOSIT
      frmCashDoc.HeaderText = CashDocType2Text(CASH_DEPOSIT)
      Load frmCashDoc
      frmCashDoc.Show 1
         
      Unload frmCashDoc
      Set frmCashDoc = Nothing
      Exit Sub
   ElseIf (lMenuChosen = 3) Then
      frmCashDoc.Area = POST_CHEQUE                   '��׹�ѹ������Ѻ�Թ
      frmCashDoc.DocumentType = POST_CHEQUE
      frmCashDoc.HeaderText = CashDocType2Text(POST_CHEQUE)
      Load frmCashDoc
      frmCashDoc.Show 1
      
      Unload frmCashDoc
      Set frmCashDoc = Nothing
      Exit Sub
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdExport_Click()
   If Not VerifyAccessRight("INVENTORY_EXPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmInventoryDoc2
   frmInventoryDoc2.Show 1
   
   Unload frmInventoryDoc2
   Set frmInventoryDoc2 = Nothing
End Sub


Private Sub cmdImport_Click()
   If Not VerifyAccessRight("INVENTORY_IMPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmInventoryDoc1
   frmInventoryDoc1.Show 1
   
   Unload frmInventoryDoc1
   Set frmInventoryDoc1 = Nothing
End Sub
Private Sub cmdInventoryReport_Click()
   If Not VerifyAccessRight("INVENTORY_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdInventoryReport.Caption
   frmSummaryReport.MasterMode = 4
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdLedgerReport_Click()
   frmSummaryReport.HeaderText = cmdLedgerReport.Caption
   frmSummaryReport.MasterMode = 6
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
   Set frmSummaryReport = Nothing
End Sub

Private Sub cmdMainCustomer_Click()
   If Not VerifyAccessRight("MAIN_CUSTOMER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmCustomer
   frmCustomer.Show 1
   
   Unload frmCustomer
   Set frmCustomer = Nothing
End Sub

Private Sub cmdMainEmployee_Click()
   If Not VerifyAccessRight("MAIN_EMPLOYEE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmEmployee
   frmEmployee.Show 1
   
   Unload frmEmployee
   Set frmEmployee = Nothing
End Sub

Private Sub cmdMainEnterprise_Click()
   If Not VerifyAccessRight("MAIN_ENTERPRISE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmAddEditEnterprise.ShowMode = SHOW_ADD
   frmAddEditEnterprise.HeaderText = cmdMainEnterprise.Caption
   Load frmAddEditEnterprise
   frmAddEditEnterprise.Show 1
   
   Unload frmAddEditEnterprise
   Set frmAddEditEnterprise = Nothing
End Sub

Private Sub cmdMainReport_Click()
   If Not VerifyAccessRight("MAIN_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdMainReport.Caption
   frmSummaryReport.MasterMode = 3
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
   Set frmSummaryReport = Nothing
End Sub

Private Sub cmdMainSupplier_Click()
   If Not VerifyAccessRight("MAIN_SUPPLIER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmSupplier
   frmSupplier.Show 1
   
   Unload frmSupplier
   Set frmSupplier = Nothing
End Sub

Private Sub cmdMaster1_Click()
   If Not VerifyAccessRight("MASTER_MAIN") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMaster1.Caption
   frmMasterMain.MasterMode = 3
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMaster2_Click()
   If Not VerifyAccessRight("MASTER_INVENTORY") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMaster2.Caption
   frmMasterMain.MasterMode = 1
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMaster3_Click()
   If Not VerifyAccessRight("MASTER_PIG") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMaster3.Caption
   frmMasterMain.MasterMode = 2
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMaster4_Click()
   If Not VerifyAccessRight("MASTER_LEDGER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMaster3.Caption
   frmMasterMain.MasterMode = 4
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMaster5_Click()
   If Not VerifyAccessRight("MASTER_PACKAGE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMaster5.Caption
   frmMasterMain.MasterMode = 5
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing

End Sub

Private Sub cmdParameter_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("�������������Դ", "-", "�������������������", "-", "�������������͹ (�٭����)", "-", "�����������â��", "-", "�����������Ҥ������/��", "-", "�ʹ¡���ء�", "-", "��������������� �", "-", "���������� % ��â��", "-", "����������������¹�������ء�", "-", "�����������ë����ء�", "-", "�����������ûѹ��������", "-", "�����������ä���ʹ�ء�", "-", "����������������¢��/������", "-", "¡����� GL", "-", "G ��Ѻ�ѵ��")
   Set oMenu = Nothing
   If lMenuChosen <= 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      frmParameter.ParamArea = 1
      frmParameter.HeaderText = "�������������Դ"
   ElseIf lMenuChosen = 3 Then
      frmParameter.ParamArea = 2
      frmParameter.HeaderText = "�������������������"
   ElseIf lMenuChosen = 5 Then
      frmParameter.ParamArea = 3
      frmParameter.HeaderText = "�������������͹ (�٭����)"
   ElseIf lMenuChosen = 7 Then
      frmParameter.ParamArea = 4
      frmParameter.HeaderText = "�����������â��"
'   ElseIf lMenuChosen = 9 Then
'      frmParameter.ParamArea = 5
'      frmParameter.HeaderText = "�������������������˹ѡ"
   ElseIf lMenuChosen = 9 Then
      frmParameter.ParamArea = 6
      frmParameter.HeaderText = "�����������Ҥ������/��"
   ElseIf lMenuChosen = 11 Then
      frmParameter.ParamArea = 7
      frmParameter.HeaderText = "�ʹ¡���ء�"
   ElseIf lMenuChosen = 13 Then
      frmParameter.ParamArea = 9
      frmParameter.HeaderText = "��������������� �"
   ElseIf lMenuChosen = 15 Then
      frmParameter.ParamArea = 10
      frmParameter.HeaderText = "���������� % ��â��"
   ElseIf lMenuChosen = 17 Then
      frmParameter.ParamArea = 11
      frmParameter.HeaderText = "����������������¹�������ء�"
   ElseIf lMenuChosen = 19 Then
      frmParameter.ParamArea = 12
      frmParameter.HeaderText = "�����������ë����ء�"
   ElseIf lMenuChosen = 21 Then
      frmParameter.ParamArea = 13
      frmParameter.HeaderText = "�����������ûѹ��������"
   ElseIf lMenuChosen = 23 Then
      frmParameter.ParamArea = 14
      frmParameter.HeaderText = "�����������ä���ʹ�ء�"
   ElseIf lMenuChosen = 25 Then
      frmParameter.ParamArea = 15
      frmParameter.HeaderText = "����������������¢�º�����"
   ElseIf lMenuChosen = 27 Then
      frmParameter.ParamArea = 16
      frmParameter.HeaderText = "����������¡����� GL"
   ElseIf lMenuChosen = 29 Then
      frmParameter.ParamArea = 17
      frmParameter.HeaderText = "����������G ��Ѻ�ѵ��"
   End If
   
   Load frmParameter
   frmParameter.Show 1
   
   Unload frmParameter
   Set frmParameter = Nothing
End Sub

Private Sub cmdPasswd_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim Cm As CCapitalMovement
Dim p As CPatch

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("����¹���ʼ�ҹ", "-", "����Ң������ʹ¡��", "-", "��� �ӹǹ/��Ť�� �ѵ�شԺ", "-", "��Ѻ��ا�Ҥ������", "-", "���ҧ�����š������͹��ǵ鹷ع", "-", "ź�����š������͹��ǵ鹷ع", "-", "Export ��Ţ��", "-", "Import ��Ţ��", "-", "��˹���ǧ�ѹ����͡���", "-", "����ʹ¡���к�����(�ء� + �ѵ�شԺ)", "-", "����ʹ¡���к�����(�١˹�� + �Թʴ)", "-", "�͹�Ԥ�Ţ����͡���", "-", "EXPORT ��������ѧ �к���������", "-", "���ҧ����Ҥ�", "-", "UPDATE �Ҥ�", "-", "���¢����ŵ��ҧ�ʹ�����������㹵��ҧ���ͧ", "-", "���ͺ������", "-", "UPDATE �����ź�ŷ������������", "-", "�����§ҹ�ç�ʹ����ءÿ����", "-", "ź��¡���������Ѻ�Թ", "-", "�Ѿഴ�ҤҺ�Ţ���ء� �ҡ Excel", "-", "�к���駤���Է���͹��ѵ����觫���")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      Load frmChangePassword
      frmChangePassword.Show 1
      
      Unload frmChangePassword
      Set frmChangePassword = Nothing
   ElseIf lMenuChosen = 3 Then
      Load frmImportDoc
      frmImportDoc.Show 1
      
      Unload frmImportDoc
      Set frmImportDoc = Nothing
   ElseIf lMenuChosen = 5 Then
      frmInventoryBalance.ShowMode = SHOW_ADD
      Load frmInventoryBalance
      frmInventoryBalance.Show 1
      
      Unload frmInventoryBalance
      Set frmInventoryBalance = Nothing
   ElseIf lMenuChosen = 7 Then
      Load frmReArrangeDoc
      frmReArrangeDoc.Show 1
      
      Unload frmReArrangeDoc
      Set frmReArrangeDoc = Nothing
   ElseIf lMenuChosen = 9 Then
      Load frmCapitalMovement
      frmCapitalMovement.Show 1
   
      Unload frmCapitalMovement
      Set frmCapitalMovement = Nothing
   ElseIf lMenuChosen = 11 Then
      glbErrorLog.LocalErrorMsg = "������зӡ��ź�����š������͹��ǵ鹷ع ��е鹷ع¡�� �����Թ��õ���������"
      If glbErrorLog.AskMessage = vbYes Then
         Call EnableForm(Me, False)
         Set Cm = New CCapitalMovement
         Call glbDaily.StartTransaction
         Call Cm.DeleteAllData
         Call glbDaily.CommitTransaction
         Set Cm = Nothing
         Call EnableForm(Me, True)
      End If
   ElseIf lMenuChosen = 13 Then
      Load frmExportBillingDoc
      frmExportBillingDoc.Show 1
      
      Unload frmExportBillingDoc
      Set frmExportBillingDoc = Nothing
   ElseIf lMenuChosen = 15 Then
      Load frmImportBillingDoc
      frmImportBillingDoc.Show 1
      
      Unload frmImportBillingDoc
      Set frmImportBillingDoc = Nothing
   ElseIf lMenuChosen = 17 Then
      Load frmLockDocDate
      frmLockDocDate.Show 1
      
      Unload frmLockDocDate
      Set frmLockDocDate = Nothing
   ElseIf lMenuChosen = 19 Then
      Load frmInitBalance
      frmInitBalance.Show 1
      
      Unload frmInitBalance
      Set frmInitBalance = Nothing
   ElseIf lMenuChosen = 21 Then
      Load frmInitBalanceEx
      frmInitBalanceEx.Show 1
      
      Unload frmInitBalanceEx
      Set frmInitBalanceEx = Nothing
   ElseIf lMenuChosen = 23 Then
      frmConfigDoc.HeaderText = "�͹�Ԥ�Ţ����͡���"
      Load frmConfigDoc
      frmConfigDoc.Show 1
      
      Unload frmConfigDoc
      Set frmConfigDoc = Nothing
   ElseIf lMenuChosen = 25 Then
      Load frmExportToSumFarm
      frmExportToSumFarm.Show 1
      
      Unload frmExportToSumFarm
      Set frmExportToSumFarm = Nothing
   ElseIf lMenuChosen = 27 Then
      Load frmPriceAdjust
      frmPriceAdjust.Show 1
      
      Unload frmPriceAdjust
      Set frmPriceAdjust = Nothing
   ElseIf lMenuChosen = 29 Then
      Load frmUpdateAvgPrice
      frmUpdateAvgPrice.Show 1
      
      Unload frmUpdateAvgPrice
      Set frmUpdateAvgPrice = Nothing
   ElseIf lMenuChosen = 31 Then
      Load frmUpdateAvgPrice
      frmUpdateAvgPrice.Show 1
      
      Unload frmUpdateAvgPrice
      Set frmUpdateAvgPrice = Nothing
   ElseIf lMenuChosen = 33 Then
      Load frmCheckError
      frmCheckError.Show 1
      
      Unload frmCheckError
      Set frmCheckError = Nothing
   ElseIf lMenuChosen = 35 Then
      Load frmUpdateCloseBilling
      frmUpdateCloseBilling.Show 1
      
      Unload frmUpdateCloseBilling
      Set frmUpdateCloseBilling = Nothing
   ElseIf lMenuChosen = 37 Then
      Load frmUpdateT706
      frmUpdateT706.Show 1
      
      Unload frmUpdateT706
      Set frmUpdateT706 = Nothing
   ElseIf lMenuChosen = 39 Then
      Load frmDeleteRcpDetail
      frmDeleteRcpDetail.Show 1

      Unload frmDeleteRcpDetail
      Set frmDeleteRcpDetail = Nothing
   ElseIf lMenuChosen = 41 Then
      Load frmUpdateDoExcel
      frmUpdateDoExcel.Show 1

      Unload frmUpdateDoExcel
      Set frmUpdateDoExcel = Nothing
   ElseIf lMenuChosen = 43 Then
      If Not VerifyAccessRight("PROGRAM_APPROVE-PO", "�к���駤���Է���͹��ѵ����觫���") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
     
     frmAuthenPO.HeaderText = "�Է�����͹��ѵ����觫���"
     frmAuthenPO.ShowMode = SHOW_VIEW_ONLY
      Load frmAuthenPO
      frmAuthenPO.Show 1

      Unload frmAuthenPO
      Set frmAuthenPO = Nothing
   End If
End Sub

Private Sub cmdPigAdjustment_Click()
   If Not VerifyAccessRight("PIG_ADJUST") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmPigDoc3
   frmPigDoc3.Show 1
   
   Unload frmPigDoc3
   Set frmPigDoc3 = Nothing
End Sub

Private Sub cmdPigBirth_Click()
   If Not VerifyAccessRight("PIG_BIRTH") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmPigDoc1
   frmPigDoc1.Show 1
   
   Unload frmPigDoc1
   Set frmPigDoc1 = Nothing
End Sub

Private Sub cmdPigImport_Click()
   If Not VerifyAccessRight("PIG_IMPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmPigDoc4
   frmPigDoc4.Show 1
   
   Unload frmPigDoc4
   Set frmPigDoc4 = Nothing
End Sub

Private Sub cmdPigReport_Click()
   If Not VerifyAccessRight("PIG_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdPigReport.Caption
   frmSummaryReport.MasterMode = 5
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
   Set frmSummaryReport = Nothing
End Sub

Private Sub cmdPigTransfer_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not VerifyAccessRight("PIG_TRANSFER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("�͹�����ء�", "-", "�͹�ء�������͹���", "-", "�͹�ء��繾����� (����¹�������ء�)", "-", "�͹����¹ʶҹ��ء�����͹���", "-", "�͹ G �� L")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      frmPigDoc2.DocumentType = 6
   ElseIf lMenuChosen = 3 Then
      frmPigDoc2.DocumentType = 7
   ElseIf lMenuChosen = 5 Then
      frmPigDoc2.DocumentType = 8
   ElseIf lMenuChosen = 7 Then
      frmPigDoc2.DocumentType = 12
   ElseIf lMenuChosen = 9 Then
      frmPigDoc2.DocumentType = 888
   End If
   Load frmPigDoc2
   frmPigDoc2.Show 1
   
   Unload frmPigDoc2
   Set frmPigDoc2 = Nothing
End Sub

Private Sub cmdPigWeek_Click()
   If Not VerifyAccessRight("PIG_WEEK") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmPigWeek
   frmPigWeek.Show 1
   
   Unload frmPigWeek
   Set frmPigWeek = Nothing
End Sub

Private Sub cmdRawMatterial_Click()
   If Not VerifyAccessRight("INVENTORY_PART") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmPartItem
   frmPartItem.Show 1
   
   Unload frmPartItem
   Set frmPartItem = Nothing
End Sub

Private Sub cmdSell_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
   
   If Not VerifyAccessRight("LEDGER_SELL") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("��觢ͧ", "-", "��Ѻ�Թ���Ǥ���", "-", "������Ѻ�Թ", "-", "�����˹��", "-", "�Ŵ˹��")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      If Not VerifyAccessRight("LEDGER_SELL_1", "��觢ͧ") Then                                                                                     '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmBillingDoc1.DocumentType = 1
   ElseIf lMenuChosen = 3 Then
      glbErrorLog.LocalErrorMsg = "��ǹ�ѧ��ѹ�ҹ����ѧ����Դ�����ҹ"
      glbErrorLog.ShowUserError
      Exit Sub
      
   ElseIf lMenuChosen = 5 Then
      If Not VerifyAccessRight("LEDGER_SELL_2", "������Ѻ�Թ") Then                                                                               '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 2
   ElseIf lMenuChosen = 7 Then
      If Not VerifyAccessRight("LEDGER_SELL_3", "�����˹��") Then                                                                               '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 3
   ElseIf lMenuChosen = 9 Then
      If Not VerifyAccessRight("LEDGER_SELL_4", "�Ŵ˹��") Then                                                                                          '''''''''
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmBillingDoc1.DocumentType = 4
   End If
   
   Load frmBillingDoc1
   frmBillingDoc1.Show 1

   Unload frmBillingDoc1
   Set frmBillingDoc1 = Nothing

End Sub

Private Sub cmdSetCost_Click()
   If Not VerifyAccessRight("PACKAGE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmPackage
   frmPackage.Show 1

   Unload frmPackage
   Set frmPackage = Nothing

End Sub

Private Sub cmdSimulateReport_Click()
   If Not VerifyAccessRight("SIMULATE_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdSimulateReport.Caption
   frmSummaryReport.MasterMode = 8
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdTransfer_Click()
   If Not VerifyAccessRight("INVENTORY_TRANSFER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmInventoryDoc3
   frmInventoryDoc3.Show 1

   Unload frmInventoryDoc3
   Set frmInventoryDoc3 = Nothing
End Sub

Private Sub cmdUser_Click()
   If Not VerifyAccessRight("ADMIN_USER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmUser
   frmUser.Show 1
   
   Unload frmUser
   Set frmUser = Nothing
End Sub

Private Sub cmdUserGroup_Click()

   If Not VerifyAccessRight("ADMIN_GROUP") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmUserGroup
   frmUserGroup.Show 1
   
   Unload frmUserGroup
   Set frmUserGroup = Nothing
End Sub

Private Sub Command1_Click()
Dim II As CImportItem
Dim IsOK As Boolean

   Set II = New CImportItem
   Call glbDaily.PatchPigBirthPartID(II, IsOK, True, glbErrorLog)
   Set II = Nothing
End Sub

Private Sub Command10_Click()
   Call EnableForm(Me, False)
   Call glbDaily.PatchBirthItemParam(True, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub Command2_Click()
   Call glbDaily.ClearInterim(True)
   glbErrorLog.LocalErrorMsg = "Clear successfully"
   glbErrorLog.ShowUserError
End Sub

Private Sub Command3_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 2
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc2(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Command4_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 4
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc4(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Command5_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 1
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc1(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Command6_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 3
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc3(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Command7_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 7
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc6(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Command8_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 6
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc8(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Command9_Click()
Dim D As CLegacy_h
Dim iCount As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim Dt As CLegacy_d
Dim Ivd As CInventoryDoc

   Call EnableForm(Me, False)
   
   Set D = New CLegacy_h
   D.DOCUMENT_ID = 5
   Call glbDaily.QueryLegacyH(D, m_Rs, iCount, IsOK, glbErrorLog)
   Set D = Nothing
   
   Set TempRs = New ADODB.Recordset
   While Not m_Rs.EOF
      Set D = New CLegacy_h
      D.LEGACY_H_ID = NVLI(m_Rs("LEGACY_H_ID"), -1)
      D.QueryFlag = 1
      Call glbDaily.QueryLegacyH(D, TempRs, iCount, IsOK, glbErrorLog)
      Call D.PopulateFromRS(1, TempRs)
      
      Set Ivd = New CInventoryDoc
      Call glbDaily.CreateInventoryDoc7(D, Ivd, "Y")
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
      Set Ivd = Nothing
      
      m_Rs.MoveNext
   Wend
   
   Set TempRs = Nothing
   Set D = Nothing
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
Dim OKClick As Boolean
Dim iCount As Long

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      
      DoEvents
      Call EnableForm(Me, False)
      
      glbEnterPrise.ENTERPRISE_ID = -1
      Call glbEnterPrise.QueryData(m_Rs, iCount)
      If Not m_Rs.EOF Then
         Call glbEnterPrise.PopulateFromRS(1, m_Rs)
      End If
      
      Call PatchDB
      Call EnableForm(Me, True)
      
      Call LoadCustomerPackage(Nothing, CustomerPackage)
      Call LoadPackageDetail(Nothing, PackageDetail)
      
      trvMain.Refresh
      Load frmLogin
      frmLogin.Show 1
      
      OKClick = frmLogin.OKClick
      
      Unload frmLogin
      Set frmLogin = Nothing
      
      If Not OKClick Then
         m_MustAsk = False
         Unload Me
      Else
         If glbUser.SIMULATE_FLAG = "Y" Then
            Call AddSimulateMenuItem
         End If
         Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
      End If
   End If
End Sub

Private Sub Form_Load()
   m_MustAsk = True
   Call InitFormLayout
   Set m_Rs = New ADODB.Recordset
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   If m_MustAsk Then
      glbErrorLog.LocalErrorMsg = MapText("��ҹ��ͧ����͡�ҡ��������������")
      If glbErrorLog.AskMessage = vbYes Then
         Cancel = False
      Else
         Cancel = True
      End If
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Call ReleaseAll
   Set m_Rs = Nothing
End Sub
Private Sub Timer1_Timer()
   Timer1.Enabled = False
   
   lblDateTime.Caption = "                                                    "
   lblDateTime.Caption = DateToStringExtEx3(Now)
   lblUsername.Caption = MapText("����� : ") & " " & glbUser.USER_NAME
   lblUserGroup.Caption = MapText("���������� : ") & " " & glbUser.GROUP_NAME
   
  Timer1.Enabled = True
End Sub
Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node Is Nothing Then
      Exit Sub
   End If
   
   fraAdmin.Visible = False
   fraMaster.Visible = False
   fraMain.Visible = False
   fraInventory.Visible = False
   fraPig.Visible = False
   fraLedger.Visible = False
   fraSimulate.Visible = False
   FraPackage.Visible = False
   
   pnlHeader.Caption = Node.Text
   If Node.Key = ROOT_TREE & " 1-0" Then
        fraAdmin.Left = 4710
        fraAdmin.Top = 2190
        fraAdmin.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-1" Then
        fraMaster.Left = 4710
        fraMaster.Top = 2190
        fraMaster.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
        fraMain.Left = 4710
        fraMain.Top = 2190
        fraMain.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
        fraInventory.Left = 4710
        fraInventory.Top = 2190
        fraInventory.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-4" Then
        fraPig.Left = 4710
        fraPig.Top = 2190
        fraPig.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-5" Then
        fraLedger.Left = 4710
        fraLedger.Top = 2190
        fraLedger.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 1-6" Then
        fraSimulate.Left = 4710
        fraSimulate.Top = 2190
        fraSimulate.Visible = True
    ElseIf Node.Key = ROOT_TREE & " 1-7" Then
        FraPackage.Left = 4710
        FraPackage.Top = 2190
        FraPackage.Visible = True
   End If
End Sub