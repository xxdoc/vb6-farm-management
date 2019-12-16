VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "JasmineUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboOrderBy 
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
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3930
      Width           =   2685
   End
   Begin VB.ComboBox cboOrderType 
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
      Left            =   6870
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3930
      Width           =   2685
   End
   Begin VB.ComboBox cboAccountType 
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
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3480
      Width           =   2685
   End
   Begin VB.ComboBox cboAccountStatus 
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
      Left            =   6870
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3480
      Width           =   2685
   End
   Begin VB.ComboBox cboDocument 
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
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3030
      Width           =   3495
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   2370
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   979
      _Version        =   393216
   End
   Begin projBase.uctlTextBox txtGeneral 
      Height          =   435
      Index           =   0
      Left            =   810
      TabIndex        =   1
      Top             =   1500
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   767
   End
   Begin VB.CommandButton Command1 
      Height          =   525
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   270
      Width           =   1815
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_TempControl As Collection
Private m_TextControls As Collection
Private m_ComboControls As Collection

Private Sub Form_Load()
Dim Ctl As Control
Dim C As CReportControl

   Set m_TempControl = New Collection
   Set m_TextControls = New Collection
   
  Load Command1(1)
  Command1(1).Visible = True
  Command1(1).Top = 0
  Command1(1).Left = 0
   Set C = New CReportControl
   C.ControlIndex = 1
   C.ControlType = "B"
  Call m_TempControl.Add(C)
  Set C = Nothing
  
  Load Command1(2)
  Command1(2).Visible = True
  Command1(2).Top = 100
  Command1(2).Left = 200
   Set C = New CReportControl
   C.ControlIndex = 2
   C.ControlType = "B"
  Call m_TempControl.Add(C)
  Set C = Nothing
  
  Load txtGeneral(1)
  txtGeneral(1).Visible = True
  txtGeneral(1).Top = 100
  txtGeneral(1).Left = 1800
   Set C = New CReportControl
   C.ControlIndex = 2
   C.ControlType = "T"
  Call m_TempControl.Add(C)
  Call m_TextControls.Add(txtGeneral(1))
  m_TextControls(1).Text = "DDDDDD"
  Set C = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_TempControl = Nothing
   Set m_TextControls = Nothing
End Sub
