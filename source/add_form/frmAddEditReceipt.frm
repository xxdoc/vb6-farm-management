VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditReceipt 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditReceipt.frx":0000
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
      TabIndex        =   20
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboArea 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2730
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.ComboBox cboEnpAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2310
         Width           =   9585
      End
      Begin VB.ComboBox cboCustomerAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1860
         Width           =   9585
      End
      Begin VB.ComboBox cboAccount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1410
         Width           =   2925
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1410
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   2
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   13
         Top             =   4560
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
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   2360
         TabIndex        =   1
         Top             =   960
         Width           =   2235
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
         Height          =   2655
         Left            =   210
         TabIndex        =   14
         Top             =   5100
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4683
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
         FontSize        =   12
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditReceipt.frx":27A2
         Column(2)       =   "frmAddEditReceipt.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditReceipt.frx":290E
         FormatStyle(2)  =   "frmAddEditReceipt.frx":2A6A
         FormatStyle(3)  =   "frmAddEditReceipt.frx":2B1A
         FormatStyle(4)  =   "frmAddEditReceipt.frx":2BCE
         FormatStyle(5)  =   "frmAddEditReceipt.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditReceipt.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1650
            ScaleHeight     =   435
            ScaleWidth      =   495
            TabIndex        =   32
            Top             =   90
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotal 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   3210
         Width           =   1785
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   10
         Top             =   3660
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalRcp 
         Height          =   435
         Left            =   5460
         TabIndex        =   34
         Top             =   3210
         Width           =   1785
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDipRcp 
         Height          =   435
         Left            =   9300
         TabIndex        =   37
         Top             =   3210
         Width           =   1785
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   39
         Top             =   2760
         Width           =   1785
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalWeight 
         Height          =   435
         Left            =   5460
         TabIndex        =   40
         Top             =   2760
         Width           =   1785
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1860
         TabIndex        =   45
         Top             =   4080
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   767
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   46
         Top             =   4200
         Width           =   1485
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   44
         Top             =   2790
         Width           =   1545
      End
      Begin VB.Label Label8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3660
         TabIndex        =   43
         Top             =   2820
         Width           =   585
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   42
         Top             =   2790
         Width           =   585
      End
      Begin VB.Label lblTotalweight 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   41
         Top             =   2790
         Width           =   1185
      End
      Begin VB.Label lblDipRcp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   33
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11160
         TabIndex        =   38
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label lblTotalRcp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   36
         Top             =   3240
         Width           =   1185
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   35
         Top             =   3240
         Width           =   585
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceipt.frx":2F36
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin VB.Label lblArea 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8820
         TabIndex        =   31
         Top             =   2820
         Visible         =   0   'False
         Width           =   915
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10500
         TabIndex        =   3
         Top             =   960
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblSellBy 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   30
         Top             =   3720
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8250
         TabIndex        =   11
         Top             =   3690
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceipt.frx":3250
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   9870
         TabIndex        =   12
         Top             =   3690
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin VB.Label lblEnpAddress 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1950
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   27
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   26
         Top             =   1470
         Width           =   1635
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3660
         TabIndex        =   25
         Top             =   3270
         Width           =   585
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   1545
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   23
         Top             =   1020
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   18
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceipt.frx":356A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   19
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   16
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReceipt.frx":3884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   17
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   21
         Top             =   1020
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_DateHasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_Customers As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ReceiptType As Long
Public DocumentSubType As Long
Public Area As Long
Public BATCH_ID As Long

Private FileName As String
Private m_SumUnit As Double

Private m_Cd As Collection
Private DocAdd As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingDoc.BILLING_DOC_ID = ID
      m_BillingDoc.BATCH_ID = BATCH_ID
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If itemcount > 0 Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)

      uctlDocumentDate.ShowDate = m_BillingDoc.DOCUMENT_DATE
      txtDocumentNo.Text = m_BillingDoc.DOCUMENT_NO
      If Area = 1 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.CUSTOMER_ID)
         cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.ACCOUNT_ID)
      ElseIf Area = 2 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.SUPPLIER_ID)
         cboAccount.ListIndex = -1
      End If
      cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_BillingDoc.BILLING_ADDRESS_ID)
      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)

      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
      cboArea.ListIndex = IDToListIndex(cboArea, m_BillingDoc.REGION_ID)
      chkCommit.Value = FlagToCheck(m_BillingDoc.COMMIT_FLAG)
      chkCommit.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      txtNote.Text = m_BillingDoc.NOTE
      Call ShowButton(1)
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
      
   Call EnableForm(Me, True)
End Sub

Private Sub ShowButton(Ind As Long)
   If ShowMode = SHOW_ADD Then
      Exit Sub
   End If
   
   If Ind = 1 Then
      cmdAdd.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      If ReceiptType = 1 Then
         cmdEdit.Enabled = True
      Else
         cmdEdit.Enabled = False
      End If
      
   ElseIf Ind = 2 Then
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False

      cmdDelete.Enabled = False
   End If
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Sub PopulateGuiID(Bd As CBillingDoc)
Dim Di As CDoItem

   For Each Di In Bd.DoItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(Bd As CBillingDoc) As Long
Dim Di As CDoItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In Bd.DoItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long) As CExportItem
Dim EI As CExportItem

      For Each EI In Ivd.ImportExports
         If EI.LINK_ID = GuiID Then
            Set GetExportItem = EI
            Exit Function
         End If
      Next EI
End Function

Private Function DO2InventoryDoc(Bd As CBillingDoc, Ivd As CInventoryDoc) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDoItem
Dim EI As CExportItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If Bd.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = Bd.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If

   Ivd.DOCUMENT_DATE = Bd.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = Bd.DOCUMENT_NO
   Ivd.COMMIT_FLAG = Bd.COMMIT_FLAG
   Ivd.EXCEPTION_FLAG = Bd.EXCEPTION_FLAG
   Ivd.DOCUMENT_TYPE = 13
   Ivd.DOCUMENT_SUBTYPE = Bd.DOCUMENT_SUBTYPE
   If Bd.DOCUMENT_SUBTYPE = 1 Then 'หมู
      Ivd.SALE_FLAG = "N"
   ElseIf Bd.DOCUMENT_SUBTYPE = 2 Then 'วัตถุดิบ
      Ivd.SALE_FLAG = "Y"
   End If

   For Each Di In Bd.DoItems
      If Di.Flag = "A" Then
         Set EI = New CExportItem
         
         EI.TX_TYPE = "E"
         EI.Flag = "A"
         EI.PART_ITEM_ID = Di.PART_ITEM_ID
         EI.PIG_STATUS = Di.PIG_STATUS
         EI.LOCATION_ID = Di.LOCATION_ID
         EI.EXPORT_AMOUNT = Di.ITEM_AMOUNT
         EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         EI.TOTAL_PRICE = Di.TOTAL_PRICE
         EI.LINK_ID = Di.LINK_ID
         EI.CALCULATE_FLAG = "N"
         EI.PIG_AGE = GetAge(Di.PART_NO, Bd.DOCUMENT_DATE)
         EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
         
         Di.PIG_AGE = EI.PIG_AGE
         Di.AGE_CODE = EI.AGE_CODE
         Call Ivd.ImportExports.Add(EI)
         Set EI = Nothing
      ElseIf Di.Flag = "E" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         
         EI.Flag = "E"
         EI.PART_ITEM_ID = Di.PART_ITEM_ID
         EI.PIG_STATUS = Di.PIG_STATUS
         EI.LOCATION_ID = Di.LOCATION_ID
         EI.EXPORT_AMOUNT = Di.ITEM_AMOUNT
         EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         EI.TOTAL_PRICE = Di.TOTAL_PRICE
         EI.CALCULATE_FLAG = "N"
         EI.PIG_AGE = GetAge(Di.PART_NO, Bd.DOCUMENT_DATE)
         EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
      
         Di.PIG_AGE = EI.PIG_AGE
         Di.AGE_CODE = EI.AGE_CODE
      ElseIf Di.Flag = "D" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         EI.Flag = "D"
      End If
   Next Di
End Function

Private Sub UpdatePigAge()
Dim Di As CDoItem
Dim OldPigAge As Long

   For Each Di In m_BillingDoc.DoItems
      If (Di.Flag <> "A") And (Di.Flag <> "D") Then
         Di.Flag = "E"
      End If
   Next Di
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment
   
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("LEDGER_SELL_2_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
      uctlDocumentDate.SetFocus
      Exit Function
   End If
   
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   If CountItem(m_BillingDoc.Payments) <= 0 Then
      glbErrorLog.LocalErrorMsg = "กรุณาใส่การชำระเงินใหถูกต้องครบถ้วน"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_BillingDoc.AddEditMode = ShowMode
   m_BillingDoc.BILLING_DOC_ID = ID
    m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   If Area = 1 Then
      m_BillingDoc.DOCUMENT_TYPE = 2 'ใบเสร็จรับเงิน
      m_BillingDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_BillingDoc.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
   ElseIf Area = 2 Then
      m_BillingDoc.DOCUMENT_TYPE = 8 'ใบเสร็จรับเงิน
      m_BillingDoc.SUPPLIER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_BillingDoc.ACCOUNT_ID = -1
   End If
   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.EXCEPTION_FLAG = "N"
   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_BillingDoc.RECEIPT_TYPE = ReceiptType
   m_BillingDoc.DOCUMENT_SUBTYPE = DocumentSubType
   m_BillingDoc.EXCEPTION_FLAG = "N"
   m_BillingDoc.REGION_ID = cboArea.ItemData(Minus2Zero(cboArea.ListIndex))
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.NOTE = txtNote.Text
   
   Call PopulateGuiID(m_BillingDoc)

   Call EnableForm(Me, False)
   
   If m_DateHasModify Then
      Call UpdatePigAge
   End If
   Call DO2InventoryDoc(m_BillingDoc, Ivd)
   
   If ReceiptType <> 5 Then
      'Call glbDaily.DO2Payment(m_BillingDoc, Pm)
   End If

   If (m_BillingDoc.COMMIT_FLAG = "Y") Then
      If m_BillingDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(Ivd.ImportExports)
         If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
            m_BillingDoc.COMMIT_FLAG = "N"
            Call EnableForm(Me, True)
            Exit Function
         End If
         
         If Not glbDaily.VerifyStockBalanceEx(Ivd.ImportExports, glbErrorLog) Then
            m_BillingDoc.COMMIT_FLAG = "N"
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
   End If
   
   Call glbDaily.StartTransaction
   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If

   'ถ้ารายการเป็นใบเสร็จรับเงินชั่วคราวแสดงว่ามี PAYMENT เกิดขึ้นแล้ว
   If ReceiptType <> 5 Then
'      If Not glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         SaveData = False
'         Call glbDaily.RollbackTransaction
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
'      m_BillingDoc.PAYMENT_ID = Pm.PAYMENT_ID
   End If

   m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If

   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If

   Call glbDaily.CommitTransaction
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

Private Sub cboAccount_Click()
   m_HasModify = True
End Sub

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub
Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub
Private Sub cboPaymentType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If Area = 1 Then
      If Not VerifyCombo(lblAccountNo, cboAccount) Then
         Exit Sub
      End If
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      If (ReceiptType = 3) Or (ReceiptType = 5) Then
         frmAddReceiptItem.Area = Area
         frmAddReceiptItem.ReceiptType = ReceiptType
         If Area = 1 Then
            frmAddReceiptItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
         ElseIf Area = 2 Then
            frmAddReceiptItem.AccountID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         End If
         frmAddReceiptItem.DocumentDate = uctlDocumentDate.ShowDate
         Set frmAddReceiptItem.TempCollection = m_BillingDoc.ReceiptItems
         Set frmAddReceiptItem.CnDnItems = m_BillingDoc.ReceiptCnDns
         frmAddReceiptItem.ShowMode = SHOW_ADD
         frmAddReceiptItem.HeaderText = MapText("เพิ่มรายการใบเสร็จ")
         Load frmAddReceiptItem
         frmAddReceiptItem.Show 1
   
         OKClick = frmAddReceiptItem.OKClick
   
         Unload frmAddReceiptItem
         Set frmAddReceiptItem = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.itemcount = CountItem(m_BillingDoc.ReceiptItems)
            GridEX1.Rebind
         End If
      ElseIf ReceiptType = 1 Then
         If DocumentSubType = 1 Then
            If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
               Exit Sub
            End If
            Set frmAddEditDoItem.ParentForm = Me
            frmAddEditDoItem.Area = Area
            frmAddEditDoItem.CusId = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
            frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
            frmAddEditDoItem.ParentShowMode = ShowMode
            frmAddEditDoItem.ShowMode = SHOW_ADD
            frmAddEditDoItem.HeaderText = MapText("เพิ่มรายการใบเสร็จ")
            Load frmAddEditDoItem
            frmAddEditDoItem.Show 1
      
            OKClick = frmAddEditDoItem.OKClick
      
            Unload frmAddEditDoItem
            Set frmAddEditDoItem = Nothing
         ElseIf DocumentSubType = 2 Then
            frmAddEditDoItem2.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditDoItem2.TempCollection = m_BillingDoc.DoItems
            frmAddEditDoItem2.ParentShowMode = ShowMode
            frmAddEditDoItem2.ShowMode = SHOW_ADD
            frmAddEditDoItem2.HeaderText = MapText("เพิ่มรายการใบส่งสินค้า")
            Load frmAddEditDoItem2
            frmAddEditDoItem2.Show 1
      
            OKClick = frmAddEditDoItem2.OKClick
      
            Unload frmAddEditDoItem2
            Set frmAddEditDoItem2 = Nothing
         End If
         
         If OKClick Then
            Call GetTotalPriceEx
   
            GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditRevenueItem.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditRevenueItem.TempCollection = m_BillingDoc.Revenues
      frmAddEditRevenueItem.ParentShowMode = ShowMode
      frmAddEditRevenueItem.ShowMode = SHOW_ADD
      frmAddEditRevenueItem.HeaderText = MapText("เพิ่มรายการรายรับอื่น ๆ")
      Load frmAddEditRevenueItem
      frmAddEditRevenueItem.Show 1

      OKClick = frmAddEditRevenueItem.OKClick

      Unload frmAddEditRevenueItem
      Set frmAddEditRevenueItem = Nothing
   
      If OKClick Then
         Call GetTotalPriceEx
         GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      frmAddEditCashTran.Area = Area
      Set frmAddEditCashTran.ParentForm = Me
      frmAddEditCashTran.HeaderText = "เพิ่มรายการการชำระเงิน"
      frmAddEditCashTran.ShowMode = SHOW_ADD
      Set frmAddEditCashTran.TempCollection = m_BillingDoc.Payments
      Load frmAddEditCashTran
      frmAddEditCashTran.Show 1
      
      OKClick = frmAddEditCashTran.OKClick

      Unload frmAddEditCashTran
      Set frmAddEditCashTran = Nothing

      If OKClick Then
         m_HasModify = True

         GridEX1.itemcount = CountItem(m_BillingDoc.Payments)
         Call GridEX1.Rebind

         Call GetTotalRcp
      End If

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
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If (ReceiptType = 3) Or (ReceiptType = 5) Then
         If ID1 <= 0 Then
            m_BillingDoc.ReceiptItems.Remove (ID2)
         Else
            m_BillingDoc.ReceiptItems.Item(ID2).Flag = "D"
         End If
   
         Call GetTotalPrice
         GridEX1.itemcount = CountItem(m_BillingDoc.ReceiptItems)
         GridEX1.Rebind
         m_HasModify = True
      ElseIf ReceiptType = 1 Then
         If ID1 <= 0 Then
            m_BillingDoc.DoItems.Remove (ID2)
         Else
            m_BillingDoc.DoItems.Item(ID2).Flag = "D"
         End If
   
         Call GetTotalPriceEx
         GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
         m_HasModify = True
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_BillingDoc.Revenues.Remove (ID2)
      Else
         m_BillingDoc.Revenues.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If ID1 <= 0 Then
         m_BillingDoc.Payments.Remove (ID2)
      Else
         m_BillingDoc.Payments.Item(ID2).Flag = "D"
      End If

      Call GetTotalRcp
      GridEX1.itemcount = CountItem(m_BillingDoc.Payments)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If ID1 <= 0 Then
         m_BillingDoc.ReceiptCnDns.Remove (ID2)
      Else
         m_BillingDoc.ReceiptCnDns.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_BillingDoc.ReceiptCnDns)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ReceiptType = 1 Then
         If DocumentSubType = 1 Then
            If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
               Exit Sub
            End If
            frmAddEditDoItem.ID = ID
            Set frmAddEditDoItem.ParentForm = Me
            frmAddEditDoItem.CusId = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
            frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
            frmAddEditDoItem.HeaderText = MapText("แก้ไขรายการใบเสร็จ")
            frmAddEditDoItem.ParentShowMode = ShowMode
            frmAddEditDoItem.ShowMode = SHOW_EDIT
            Load frmAddEditDoItem
            frmAddEditDoItem.Show 1
      
            OKClick = frmAddEditDoItem.OKClick
      
            Unload frmAddEditDoItem
            Set frmAddEditDoItem = Nothing
         ElseIf DocumentSubType = 2 Then
            frmAddEditDoItem2.ID = ID
            frmAddEditDoItem2.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
            Set frmAddEditDoItem2.TempCollection = m_BillingDoc.DoItems
            frmAddEditDoItem2.HeaderText = MapText("แก้ไขรายการใบส่งสินค้า")
            frmAddEditDoItem2.ParentShowMode = ShowMode
            frmAddEditDoItem2.ShowMode = SHOW_EDIT
            Load frmAddEditDoItem2
            frmAddEditDoItem2.Show 1
      
            OKClick = frmAddEditDoItem2.OKClick
      
            Unload frmAddEditDoItem2
            Set frmAddEditDoItem2 = Nothing
         End If
         If OKClick Then
            Call GetTotalPrice
            GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditRevenueItem.ID = ID
      frmAddEditRevenueItem.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditRevenueItem.TempCollection = m_BillingDoc.Revenues
      frmAddEditRevenueItem.ParentShowMode = ShowMode
      frmAddEditRevenueItem.ShowMode = SHOW_EDIT
      frmAddEditRevenueItem.HeaderText = MapText("แก้ไขรายการรายรับอื่น ๆ")
      Load frmAddEditRevenueItem
      frmAddEditRevenueItem.Show 1

      OKClick = frmAddEditRevenueItem.OKClick

      Unload frmAddEditRevenueItem
      Set frmAddEditRevenueItem = Nothing
   
      If OKClick Then
         GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      frmAddEditCashTran.Area = Area
      Set frmAddEditCashTran.ParentForm = Me
      frmAddEditCashTran.ID = ID
      frmAddEditCashTran.HeaderText = "แก้ไขรายการการชำระเงิน"
      frmAddEditCashTran.ShowMode = SHOW_EDIT
      Set frmAddEditCashTran.TempCollection = m_BillingDoc.Payments
      Load frmAddEditCashTran
      frmAddEditCashTran.Show 1
      
      OKClick = frmAddEditCashTran.OKClick
      
      Unload frmAddEditCashTran
      Set frmAddEditCashTran = Nothing
   
      If OKClick Then
         m_HasModify = True
         
         GridEX1.itemcount = CountItem(m_BillingDoc.Payments)
         Call GridEX1.Rebind
         
         Call GetTotalRcp
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Public Sub ShowDoItemGrid()
   GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
   GridEX1.Rebind
   
   m_HasModify = True
End Sub

Private Sub CalculateIncludePrice()
Dim II As CImportItem
Dim AvgFee As Double

'   If m_SumUnit > 0 Then
'      AvgFee = Val(txtTotalAmount.Text) / m_SumUnit
'   Else
'      AvgFee = 0
'   End If
'
'   For Each II In m_BillingDoc.DoItems
'      If II.Flag <> "D" Then
'         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE + AvgFee
'      End If
'   Next II
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

Private Function Menu2Flag(MenuID As Long) As String
   If MenuID = 1 Then
      Menu2Flag = "Y"
   Else
      Menu2Flag = "N"
   End If
End Function

Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long
Dim Report As CReportInterface
Dim ReportFlag As Boolean
Dim EditMode As SHOW_MODE_TYPE
Dim HeaderText As String

   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ReportFlag = False
   
   If (ReceiptType = 3) Or (ReceiptType = 5) Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("พิมพ์ใบเสร็จรับเงิน", "ปรับค่าหน้าใบเสร็จ", "-", "พิมพ์ใบเสร็จรับเงิน บนกระดาษเปล่า (2 ภาษา)", "ปรับค่าหน้าใบเสร็จ", "-", "พิมพ์ใบสำคัญรับ", "ปรับค่าหน้าใบสำคัญรับ")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
   Else
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("พิมพ์ใบเสร็จรับเงิน", "ปรับค่าหน้าใบเสร็จรับเงิน", "-", "พิมพ์ใบส่งสินค้า", "ปรับค่าหน้าใบส่งสินค้า", "-", "พิมพ์ใบส่งสินค้า(รวมตามประเภท)", "ปรับค่าหน้ากระดาษ", "พิมพ์ใบส่งสินค้า(ขุน+สายพันธุ์)", "ปรับค่าหน้ากระดาษ")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
   End If
   
   If (ReceiptType = 3) Or (ReceiptType = 5) Then
      If (lMenuChosen = 1) Then
         ReportKey = "CReportNormalReceipt001"
         
         Set Report = New CReportNormalRcp001
         
         Call Report.AddParam(Menu2Flag(lMenuChosen), "PICTURE_FLAG")
         Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
         Call Report.AddParam(ReportKey, "REPORT_KEY")
         Call Report.AddParam(MapText("ใบเสร็จรับเงิน"), "REPORT_HEADER")
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         
         ReportFlag = True
      ElseIf (lMenuChosen = 2) Then
         ReportKey = "CReportNormalReceipt001"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("ใบเสร็จรับเงิน")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf (lMenuChosen = 4) Then
         ReportKey = "CReportNormalReceipt002"
         
         Set Report = New CReportNormalRcp002
         
         Call Report.AddParam(Menu2Flag(lMenuChosen), "PICTURE_FLAG")
         Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
         Call Report.AddParam(ReportKey, "REPORT_KEY")
         Call Report.AddParam(MapText("ใบเสร็จรับเงิน"), "REPORT_HEADER")
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         
         ReportFlag = True
      ElseIf (lMenuChosen = 5) Then
         ReportKey = "CReportNormalReceipt002"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("ใบเสร็จรับเงิน")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf (lMenuChosen = 7) Then
         ReportKey = "CReportVocherRcp001"
         
         Set Report = New CReportVocherRcp001
         
         Call Report.AddParam(Menu2Flag(lMenuChosen), "PICTURE_FLAG")
         Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
         Call Report.AddParam(ReportKey, "REPORT_KEY")
         Call Report.AddParam(MapText("ใบเสร็จรับเงิน"), "REPORT_HEADER")
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         
         ReportFlag = True
      ElseIf lMenuChosen = 8 Then
         ReportKey = "CReportVocherRcp001"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("ใบสำคัญรับเงิน")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      Else
         Exit Sub
      End If
   Else
      If (lMenuChosen = 1) Then
         ReportKey = "CReportNormalReceipt001"
         
         Set Report = New CReportNormalRcp001
         
         Call Report.AddParam(Menu2Flag(lMenuChosen), "PICTURE_FLAG")
         Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
         Call Report.AddParam(ReportKey, "REPORT_KEY")
         Call Report.AddParam(MapText("ใบเสร็จรับเงิน"), "REPORT_HEADER")
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         
         ReportFlag = True
      ElseIf (lMenuChosen = 2) Then
         ReportKey = "CReportNormalReceipt001"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("ใบเสร็จรับเงิน")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf (lMenuChosen = 4) Then
         ReportKey = "CReportNormalDO001"
         Set Report = New CReportNormalDO001
         
         Call Report.AddParam(0, "OPTION_MODE")
         Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
         Call Report.AddParam(ReportKey, "REPORT_KEY")
         Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         
         ReportFlag = True
      ElseIf lMenuChosen = 5 Then
         ReportKey = "CReportNormalDO001"
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Call Rc.QueryData(m_Rs, iCount)
         HeaderText = MapText("ใบส่งสินค้า")
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            EditMode = SHOW_EDIT
         Else
            EditMode = SHOW_ADD
         End If
      ElseIf lMenuChosen = 7 Or lMenuChosen = 9 Then
         ReportKey = "CReportNormalDO002"
         Set Report = New CReportNormalDO002
      
         If lMenuChosen = 7 Then
            Call Report.AddParam(0, "OPTION_MODE")
         ElseIf lMenuChosen = 9 Then
            Call Report.AddParam(1, "OPTION_MODE")
         End If
      
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      
      ReportFlag = True
   ElseIf lMenuChosen = 8 Or lMenuChosen = 10 Then
      ReportKey = "CReportNormalDO002"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบส่งสินค้า")
         
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
      Else
         Exit Sub
      End If
   End If
   
   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1
         
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   ID = m_BillingDoc.BILLING_DOC_ID
   m_BillingDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Click()
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
      Call LoadRegion(cboArea)
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If
      
      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_BillingDoc.QueryFlag = 0
         Call QueryData(False)
         uctlDocumentDate.ShowDate = Now
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
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   
   Set m_Cd = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 1305
   Col.Caption = MapText("สัปดาห์เกิด")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1230
   Col.Caption = MapText("ประเภทสุกร")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2175
   Col.Caption = MapText("สถานะสุกร")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 855
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 930
   Col.Caption = MapText("น้ำหนัก")
   
   Set Col = GridEX1.Columns.Add '9
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1290
   Col.Caption = MapText("ส่วนลด")
   
   Set Col = GridEX1.Columns.Add '8
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1575
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.Add '9
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1290
   Col.Caption = MapText("ราคา/หน่วย")

   Set Col = GridEX1.Columns.Add '10
   Col.Width = 2235
   Col.Caption = MapText("โรงเรือน")
End Sub

Private Sub InitGrid2_1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 2400
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1725
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3420
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 1770
   Col.Caption = MapText("จำนวน")
      
   Set Col = GridEX1.Columns.Add '7
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1890
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.Add '8
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1290
   Col.Caption = MapText("ราคา/หน่วย")

   Set Col = GridEX1.Columns.Add '9
   Col.Width = 2235
   Col.Caption = MapText("สถานที่จัดเก็บ")
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 3195
   Col.Caption = MapText("เลขที่เอกสาร")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2565
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2460
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน (ตามบิล)")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2460
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ชำระ")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 3345
   Col.Caption = MapText("ประเภทเอกสาร")
   
   Set Col = GridEX1.Columns.Add '8
   Col.Visible = False
   Col.Caption = MapText("DO_ID")
End Sub

Private Sub InitGrid3()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 1650
   Col.Caption = MapText("รหัสรายได้อื่น ๆ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 5625
   Col.Caption = MapText("รายได้อื่น ๆ")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2175
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2115
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")
End Sub

Private Sub InitGrid4()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 2220
   Col.Caption = MapText("เลขที่ใบลด/เพิ่มหนี้")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2205
   Col.Caption = MapText("วันที่ลด/เพิ่มหนี้")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2190
   Col.Caption = MapText("ใบส่งของ")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2175
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ลดหนี้")

   Set Col = GridEX1.Columns.Add '7
   Col.Width = 2550
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("เพิ่มหนี้")
End Sub
Private Sub GetTotalPrice()
Dim II As CReceiptItem
Dim Sum1 As Double
Dim Sum2 As Double
   Sum1 = 0
   Sum2 = 0
   For Each II In m_BillingDoc.ReceiptItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.PAID_AMOUNT
         Sum2 = Sum2 + 0
      End If
   Next II

   txtNetTotal.Text = Format(Sum1, "0.00")
   txtDipRcp.Text = Format(Val(txtNetTotal.Text) - Val(txtTotalRcp.Text), "0.00")
End Sub
Private Sub GetTotalRcp()
Dim Pm As CCashTran
Dim Sum7 As Double
   Sum7 = 0
   For Each Pm In m_BillingDoc.Payments
      If Pm.Flag <> "D" Then
         Sum7 = Sum7 + Pm.GetFieldValue("AMOUNT")
      End If
   Next Pm
   
   txtTotalRcp.Text = Format(Sum7, "0.00")
   txtDipRcp.Text = Format(Val(txtNetTotal.Text) - Val(txtTotalRcp.Text), "0.00")

End Sub
Private Sub GetTotalPriceEx()
Dim II As CDoItem
Dim Pm As CCashTran
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum7 As Double

   Sum7 = 0
   Sum2 = 0
   Sum1 = 0
   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.TOTAL_PRICE
         Sum2 = Sum2 + II.ITEM_AMOUNT
         Sum3 = Sum3 + II.TOTAL_WEIGHT
      End If
   Next II

   For Each II In m_BillingDoc.Revenues
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.TOTAL_PRICE
      End If
   Next II
   
   For Each Pm In m_BillingDoc.Payments
      If Pm.Flag <> "D" Then
         Sum7 = Sum7 + Pm.GetFieldValue("AMOUNT")
      End If
   Next Pm
   
   txtNetTotal.Text = Format(Sum1, "0.00")
   txtTotalRcp.Text = Format(Sum7, "0.00")
   txtTotalAmount.Text = Format(Sum2, "0.00")
   txtTotalWeight.Text = Format(Sum3, "0.00")
   txtDipRcp.Text = Format(Val(txtNetTotal.Text) - Val(txtTotalRcp.Text), "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   GridEX1.Left = 150
   GridEX1.Top = 4980
   GridEX1.Visible = True
   GridEX1.itemcount = 0
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption

   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบเสร็จ"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   If Area = 1 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
      Call InitNormalLabel(lblSellBy, MapText("ผู้ออกใบเสร็จ"))
   ElseIf Area = 2 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ซัพ ฯ"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสซัพ ฯ"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่รับเอกสาร"))
      Call InitNormalLabel(lblSellBy, MapText("ผู้รับเอกสาร"))
   End If
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblNetTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblArea, MapText("เขตการขาย"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label8, MapText("ตัว"))
   Call InitNormalLabel(Label7, MapText("กก"))
   Call InitNormalLabel(lblTotalRcp, MapText("ยอดชำะจริง"))
   Call InitNormalLabel(lblDipRcp, MapText("ส่วนต่างชำระ"))
   Call InitNormalLabel(lblTotalAmount, MapText("จำนวนรวม"))
   Call InitNormalLabel(lblTotalweight, MapText("นน รวม"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   
   Call InitCheckBox(chkCommit, "คำนวณ")
   
   
   If Area = 1 Then
      lblAccountNo.Visible = True
      cboAccount.Visible = True
   ElseIf Area = 2 Then
      lblAccountNo.Visible = False
      cboAccount.Visible = False
   End If
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTotalRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDipRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTotalWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   txtTotalWeight.Enabled = False
   txtTotalRcp.Enabled = False
   txtDipRcp.Enabled = False
   txtNetTotal.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitCombo(cboAccount)
   Call InitCombo(cboCustomerAddress)
   Call InitCombo(cboEnpAddress)
   Call InitCombo(cboArea)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
      
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   If DocumentSubType = 1 Then 'หมู
      If (ReceiptType = 3) Or (ReceiptType = 5) Then 'อ้างใบส่งของ
         Call InitGrid1
      ElseIf ReceiptType = 1 Then 'สร้างรายการใหม่
         Call InitGrid2
      End If
   Else ' วัตถุดิบ
      If (ReceiptType = 3) Or (ReceiptType = 5) Then
         Call InitGrid1
      ElseIf ReceiptType = 1 Then
         Call InitGrid2_1
      End If
   End If
      
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("รายการใบเสร็จ")
   TabStrip1.Tabs.Add().Caption = MapText("รายการรายรับอื่น ๆ")
   If ReceiptType <> 5 Then
      TabStrip1.Tabs.Add().Caption = MapText("การชำระเงิน")
   End If
   
   If ReceiptType = 3 Then
      TabStrip1.Tabs.Add().Caption = MapText("ส่วนลด/เพิ่มหนี้")
   End If
   
   If (ReceiptType = 3) Or (ReceiptType = 5) Then
      cmdEdit.Enabled = False
   End If
   
   Call LoadPictureFromFile(glbParameterObj.ReceiptVoccherPic1, Picture1)
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
   
   m_DateHasModify = False
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_BillingDoc = New CBillingDoc
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   
   Set m_Cd = New Collection
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
      If m_BillingDoc.ReceiptItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If (ReceiptType = 3) Or (ReceiptType = 5) Then
         Dim CR As CReceiptItem
         If m_BillingDoc.ReceiptItems.Count <= 0 Then
            Exit Sub
         End If
         Set CR = GetItem(m_BillingDoc.ReceiptItems, RowIndex, RealIndex)
         If CR Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = CR.RECEIPT_ITEM_ID
         Values(2) = RealIndex
         Values(3) = CR.DOCUMENT_NO
         Values(4) = DateToStringExtEx2(CR.DOCUMENT_DATE)
         Values(5) = FormatNumber(CR.RECEIPT_ITEM_AMOUNT)
         Values(6) = FormatNumber(CR.PAID_AMOUNT)
         If ReceiptType = 3 Then
            Values(7) = "ใบส่งสินค้า"
         ElseIf ReceiptType = 5 Then
            Values(7) = "ใบรับเงินชั่วคราว"
         End If
         Values(8) = CR.DO_ID
      ElseIf ReceiptType = 1 Then
         Dim Di As CDoItem
         If m_BillingDoc.DoItems.Count <= 0 Then
            Exit Sub
         End If
         Set Di = GetItem(m_BillingDoc.DoItems, RowIndex, RealIndex)
         If Di Is Nothing Then
            Exit Sub
         End If
   
         If DocumentSubType = 1 Then
            Values(1) = Di.DO_ITEM_ID
            Values(2) = RealIndex
            Values(3) = Di.PART_NO
            Values(4) = Di.PIG_TYPE
            Values(5) = Di.PIG_STATUS_NAME
            Values(6) = FormatNumber(Di.ITEM_AMOUNT)
            Values(7) = FormatNumber(Di.TOTAL_WEIGHT)
            Values(8) = FormatNumber(Di.DISCOUNT_AMOUNT)
            Values(9) = FormatNumber(Di.TOTAL_PRICE)
            Values(10) = FormatNumber(Di.AVG_PRICE)
            Values(11) = Di.LOCATION_NAME
         ElseIf DocumentSubType = 2 Then
            Values(1) = Di.DO_ITEM_ID
            Values(2) = RealIndex
            Values(3) = Di.PART_TYPE_NAME
            Values(4) = Di.PART_NO
            Values(5) = Di.PART_DESC
            Values(6) = FormatNumber(Di.ITEM_AMOUNT)
            Values(7) = FormatNumber(Di.TOTAL_PRICE)
            Values(8) = FormatNumber(Di.AVG_PRICE)
            Values(9) = Di.LOCATION_NAME
         End If
      End If 'ReceiptType
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_BillingDoc.Revenues Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Rv As CDoItem
      If m_BillingDoc.Revenues.Count <= 0 Then
         Exit Sub
      End If
      Set Rv = GetItem(m_BillingDoc.Revenues, RowIndex, RealIndex)
      If Rv Is Nothing Then
         Exit Sub
      End If

      Values(1) = Rv.DO_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Rv.REVENUE_NO
      Values(4) = Rv.REVENUE_NAME
      Values(5) = FormatNumber(Rv.ITEM_AMOUNT)
      Values(6) = FormatNumber(Rv.TOTAL_PRICE)
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If m_BillingDoc.Payments Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ct As CCashTran
      If m_BillingDoc.Payments.Count <= 0 Then
         Exit Sub
      End If
      Set Ct = GetItem(m_BillingDoc.Payments, RowIndex, RealIndex)
      If Ct Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ct.GetFieldValue("CASH_TRAN_ID")
      Values(2) = RealIndex
      Values(3) = Ct.GetFieldValue("PAYMENT_TYPE_NAME")
      If Ct.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Or Ct.GetFieldValue("PAYMENT_TYPE") = CASHRET_PMT Then
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = BANKTRF_PMT Then
         Values(4) = Ct.GetFieldValue("ACCOUNT_NAME")
         Values(5) = Ct.GetFieldValue("BANK_NAME")
         Values(6) = Ct.GetFieldValue("BRANCH_NAME")
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = CHECK_PMT Then
         Values(4) = Ct.Cheque.GetFieldValue("CHEQUE_NO")
         Values(5) = Ct.Cheque.GetFieldValue("BANK_NAME")
         Values(6) = Ct.Cheque.GetFieldValue("BRANCH_NAME")
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      End If

   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If m_BillingDoc.ReceiptCnDns Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CnDn As CReceiptCnDn
      If m_BillingDoc.ReceiptCnDns.Count <= 0 Then
         Exit Sub
      End If
      Set CnDn = GetItem(m_BillingDoc.ReceiptCnDns, RowIndex, RealIndex)
      If CnDn Is Nothing Then
         Exit Sub
      End If
   
      Values(1) = CnDn.CNDN_ID
      Values(2) = RealIndex
      Values(3) = CnDn.CNDN_NO
      Values(4) = DateToStringExtEx2(CnDn.CNDN_DATE)
      Values(5) = CnDn.DO_NO
      Values(6) = FormatNumber(CnDn.CN_AMOUNT)
      Values(7) = FormatNumber(CnDn.DN_AMOUNT)
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   GridEX1.Top = 5050
   GridEX1.Left = 150
   GridEX1.Visible = False
   
   Call SetEnableDisableButton(cmdAdd, True)
   Call SetEnableDisableButton(cmdEdit, True)
   Call SetEnableDisableButton(cmdDelete, True)
   
   cmdAdd.Visible = True
   cmdEdit.Visible = True
   cmdDelete.Visible = True
      
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentSubType = 1 Then 'หมู
         If (ReceiptType = 3) Or (ReceiptType = 5) Then
            Call GetTotalPrice
            Call InitGrid1
            GridEX1.itemcount = CountItem(m_BillingDoc.ReceiptItems)
            GridEX1.Rebind
         ElseIf ReceiptType = 1 Then
            Call GetTotalPriceEx
            Call InitGrid2
            GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      Else
         If (ReceiptType = 3) Or (ReceiptType = 5) Then
            Call GetTotalPrice
            Call InitGrid1
            GridEX1.itemcount = CountItem(m_BillingDoc.ReceiptItems)
            GridEX1.Rebind
         ElseIf ReceiptType = 1 Then
            Call GetTotalPriceEx
            Call InitGrid2_1
            GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      End If
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call SetEnableDisableButton(cmdAdd, ReceiptType = 1)
      Call SetEnableDisableButton(cmdEdit, ReceiptType = 1)
      Call SetEnableDisableButton(cmdDelete, ReceiptType = 1)
   
      Call InitGrid3
      GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
      GridEX1.Rebind
      GridEX1.Visible = True
      
'      Call ShowButton(TabStrip1.SelectedItem.Index)
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Call InitGrid5
      GridEX1.itemcount = CountItem(m_BillingDoc.Payments)
      GridEX1.Rebind
      GridEX1.Visible = True
      Call GetTotalRcp
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Call SetEnableDisableButton(cmdAdd, False)
      Call SetEnableDisableButton(cmdEdit, False)
      Call SetEnableDisableButton(cmdDelete, True)
   
      Call InitGrid4
      GridEX1.itemcount = CountItem(m_BillingDoc.ReceiptCnDns)
      GridEX1.Rebind
      GridEX1.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
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

Private Sub txtCheckNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDipRcp_Change()
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

Private Sub txtDocumentNo_LostFocus()
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Sub
   End If
End Sub

Private Sub txtIncludeVat_Change()
   m_HasModify = True
End Sub

Private Sub txtIncludeWH_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   m_HasModify = True
End Sub
Private Sub txtTotalRcp_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
   m_DateHasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
Dim CustomerID As Long
Dim C As CCustomer

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CustomerID > 0 Then
      Set C = m_Customers(Trim(Str(CustomerID)))
      If Area = 1 Then
         Call LoadAccount(cboAccount, , CustomerID)
         cboAccount.ListIndex = 1
         
         Call LoadCustomerAddress(cboCustomerAddress, , CustomerID, True)
         cboArea.ListIndex = IDToListIndex(cboArea, C.REGION_ID)
      ElseIf Area = 2 Then
         cboAccount.ListIndex = -1
         
         Call LoadSupplierAddress(cboCustomerAddress, , CustomerID, True)
      End If
   Else
      cboAccount.ListIndex = -1
      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_LostFocus()
   If ShowMode = SHOW_ADD And uctlDocumentDate.ShowDate > 0 Then
      If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
         uctlDocumentDate.SetFocus
         Exit Sub
      End If
   ElseIf Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      txtDocumentNo.SetFocus
      Exit Sub
   ElseIf Not (uctlDocumentDate.ShowDate > 0) Then
      uctlDocumentDate.SetFocus
      Exit Sub
   End If
End Sub
Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
Private Sub cmdAuto_Click()
Dim ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
   
   ID = ConvertDocToConfigNo(1, -1, DocumentSubType, ReceiptType)
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         Dim TempCd As CConfigDoc
         If ShowMode = SHOW_ADD Then
            If Cd.GetFieldValue("UPDATE_MONTH_FLAG") = "Y" Then
               If Not (Right(Trim(Str(Cd.GetFieldValue("MM"))), 2) = Format(Month(uctlDocumentDate.ShowDate), "00")) Then
                  Set TempCd = New CConfigDoc
                  Call TempCd.SetFieldValue("RUNNING_NO", 0)
                  Call TempCd.SetFieldValue("MM", Right(Format(Year(Now), "00") & Format(Month(uctlDocumentDate.ShowDate), "00"), 4))
                  Call TempCd.SetFieldValue("CONFIG_DOC_TYPE", ID)
                  Call TempCd.UpdateYearMonthRunningNo
                  Set Cd = Nothing
                  Set m_Cd = Nothing
                  Set m_Cd = New Collection
                  Call LoadConfigDoc(Nothing, m_Cd)
                  Call cmdAuto_Click
                  Exit Sub
               End If
            ElseIf Cd.GetFieldValue("UPDATE_YEAR_FLAG") = "Y" Then
               If Not (Left(Cd.GetFieldValue("MM"), 2) = Right(Format(Year(uctlDocumentDate.ShowDate), "00"), 2)) Then
                  Set TempCd = New CConfigDoc
                  Call TempCd.SetFieldValue("RUNNING_NO", 0)
                  Call TempCd.SetFieldValue("MM", Right(Format(Year(uctlDocumentDate.ShowDate), "00") & Format(Month(Now), "00"), 4))
                  Call TempCd.SetFieldValue("CONFIG_DOC_TYPE", ID)
                  Call TempCd.UpdateYearMonthRunningNo
                  Set Cd = Nothing
                  Set m_Cd = Nothing
                  Set m_Cd = New Collection
                  Call LoadConfigDoc(Nothing, m_Cd)
                  Call cmdAuto_Click
                  Exit Sub
               End If
            End If
            Set TempCd = Nothing
         End If
         
         txtDocumentNo.Text = Cd.GetFieldValue("PREFIX") & Cd.GetFieldValue("CODE1")
         TempStr = ""
         If Cd.GetFieldValue("YEAR_TYPE") = 1 Then
            TempStr = Right(Format(Year(Now) + 543, "0000"), 2)
         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 2 Then
            TempStr = Format(Year(Now) + 543, "0000")
         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 3 Then
            TempStr = Right(Format(Year(Now), "0000"), 2)
         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 4 Then
            TempStr = Format(Year(Now), "0000")
         End If
         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE2")
         TempStr = ""
         If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
            TempStr = Format(Month(Now), "00")
         End If
         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE3")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         txtDocumentNo.Text = txtDocumentNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
         m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
         m_BillingDoc.CONFIG_DOC_TYPE = ID
      Else
         txtDocumentNo.Text = ""
      End If
      txtDocumentNo.SetFocus
   End If
End Sub
Private Sub InitGrid5()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
      Col.Width = 1965
      Col.Caption = MapText("ประเภทการชำระเงิน")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2625
      Col.Caption = MapText("เลขที่เช็ค/บัญชี")
   
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2160
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("ธนาคาร")
   
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 2565
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("สาขาธนาคาร")
   
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนเงิน")

End Sub
Public Sub RefreshGrid()
   Call GetTotalPrice

   GridEX1.itemcount = CountItem(m_BillingDoc.Payments)
   GridEX1.Rebind
End Sub


