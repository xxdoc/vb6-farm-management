VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSFrame fraRepair 
      Height          =   3165
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5583
      _Version        =   131073
   End
   Begin Threed.SSFrame fraPhyco 
      Height          =   3165
      Left            =   30
      TabIndex        =   1
      Top             =   3600
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   5583
      _Version        =   131073
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Left            =   2460
         TabIndex        =   13
         Top             =   330
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         Height          =   375
         Left            =   4260
         TabIndex        =   12
         Top             =   330
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check1"
         Height          =   375
         Left            =   6030
         TabIndex        =   11
         Top             =   330
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check1"
         Height          =   375
         Left            =   2460
         TabIndex        =   10
         Top             =   750
         Width           =   1575
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check1"
         Height          =   375
         Left            =   4260
         TabIndex        =   9
         Top             =   750
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check1"
         Height          =   375
         Left            =   6030
         TabIndex        =   8
         Top             =   750
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check1"
         Height          =   375
         Left            =   2460
         TabIndex        =   7
         Top             =   1170
         Width           =   1575
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check1"
         Height          =   375
         Left            =   4260
         TabIndex        =   6
         Top             =   1170
         Width           =   1575
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Check1"
         Height          =   375
         Left            =   6030
         TabIndex        =   5
         Top             =   1170
         Width           =   1575
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Check1"
         Height          =   375
         Left            =   2460
         TabIndex        =   4
         Top             =   1590
         Width           =   1575
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Check1"
         Height          =   375
         Left            =   4260
         TabIndex        =   3
         Top             =   1590
         Width           =   1575
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Check1"
         Height          =   375
         Left            =   6030
         TabIndex        =   2
         Top             =   1590
         Width           =   1575
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox1 
         Height          =   405
         Left            =   7770
         TabIndex        =   14
         Top             =   330
         Width           =   3645
         _ExtentX        =   2619
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox2 
         Height          =   405
         Left            =   7770
         TabIndex        =   15
         Top             =   750
         Width           =   3645
         _ExtentX        =   2619
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox3 
         Height          =   405
         Left            =   7770
         TabIndex        =   16
         Top             =   1170
         Width           =   3645
         _ExtentX        =   2619
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox uctlTextBox4 
         Height          =   405
         Left            =   7770
         TabIndex        =   17
         Top             =   1590
         Width           =   3645
         _ExtentX        =   2619
         _ExtentY        =   714
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   345
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   2085
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   2085
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   1620
         Width           =   2085
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

