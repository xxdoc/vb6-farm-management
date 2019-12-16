VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddEditYCollection 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   Icon            =   "frmAddEditYCollection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
         TabIndex        =   14
         Top             =   7770
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         Begin Threed.SSCommand cmdCancel 
            Cancel          =   -1  'True
            Height          =   615
            Left            =   5955
            TabIndex        =   110
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdOK 
            Height          =   615
            Left            =   3870
            TabIndex        =   108
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   30
         TabIndex        =   3
         Top             =   2160
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
         TabIndex        =   15
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
                  Picture         =   "frmAddEditYCollection.frx":014A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":0464
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":0D3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":34F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":3DCA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":46A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":4F7E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":5858
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":6132
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":6A0C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":6E5E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":7738
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":8012
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":88EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":91C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":9618
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":9A6A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":9BC4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":A49E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":AD78
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":B652
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":B96C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":C246
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":CF20
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":D7FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":E0D4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":E9AE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditYCollection.frx":F288
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin Threed.SSFrame fraDrug 
         Height          =   5115
         Left            =   0
         TabIndex        =   16
         Top             =   2700
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   9022
         _Version        =   131073
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   99
            Left            =   8580
            TabIndex        =   109
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   98
            Left            =   7770
            TabIndex        =   107
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   97
            Left            =   7050
            TabIndex        =   106
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   96
            Left            =   6390
            TabIndex        =   105
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   95
            Left            =   5730
            TabIndex        =   104
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   94
            Left            =   4980
            TabIndex        =   103
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   93
            Left            =   4230
            TabIndex        =   102
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   92
            Left            =   3510
            TabIndex        =   101
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   91
            Left            =   2790
            TabIndex        =   100
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   90
            Left            =   2070
            TabIndex        =   99
            Top             =   4440
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   89
            Left            =   8580
            TabIndex        =   98
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   88
            Left            =   7770
            TabIndex        =   97
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   87
            Left            =   7050
            TabIndex        =   96
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   86
            Left            =   6390
            TabIndex        =   95
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   85
            Left            =   5730
            TabIndex        =   94
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   84
            Left            =   4980
            TabIndex        =   93
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   83
            Left            =   4230
            TabIndex        =   92
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   82
            Left            =   3510
            TabIndex        =   91
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   81
            Left            =   2790
            TabIndex        =   90
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   80
            Left            =   2070
            TabIndex        =   89
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   79
            Left            =   8580
            TabIndex        =   88
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   78
            Left            =   7770
            TabIndex        =   87
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   77
            Left            =   7050
            TabIndex        =   86
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   76
            Left            =   6390
            TabIndex        =   85
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   75
            Left            =   5730
            TabIndex        =   84
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   74
            Left            =   4980
            TabIndex        =   83
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   73
            Left            =   4230
            TabIndex        =   82
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   72
            Left            =   3510
            TabIndex        =   81
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   71
            Left            =   2790
            TabIndex        =   80
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   70
            Left            =   2070
            TabIndex        =   79
            Top             =   3480
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   69
            Left            =   8580
            TabIndex        =   78
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   68
            Left            =   7770
            TabIndex        =   77
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   67
            Left            =   7050
            TabIndex        =   76
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   66
            Left            =   6390
            TabIndex        =   75
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   65
            Left            =   5730
            TabIndex        =   74
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   64
            Left            =   4980
            TabIndex        =   73
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   63
            Left            =   4230
            TabIndex        =   72
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   62
            Left            =   3510
            TabIndex        =   71
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   61
            Left            =   2790
            TabIndex        =   70
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   60
            Left            =   2070
            TabIndex        =   69
            Top             =   3030
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   59
            Left            =   8580
            TabIndex        =   68
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   58
            Left            =   7770
            TabIndex        =   67
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   57
            Left            =   7050
            TabIndex        =   66
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   56
            Left            =   6390
            TabIndex        =   65
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   55
            Left            =   5730
            TabIndex        =   64
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   54
            Left            =   4980
            TabIndex        =   63
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   53
            Left            =   4230
            TabIndex        =   62
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   52
            Left            =   3510
            TabIndex        =   61
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   51
            Left            =   2790
            TabIndex        =   60
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   50
            Left            =   2070
            TabIndex        =   59
            Top             =   2610
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   49
            Left            =   8580
            TabIndex        =   58
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   48
            Left            =   7770
            TabIndex        =   57
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   47
            Left            =   7050
            TabIndex        =   56
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   46
            Left            =   6390
            TabIndex        =   55
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   45
            Left            =   5730
            TabIndex        =   54
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   44
            Left            =   4980
            TabIndex        =   53
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   43
            Left            =   4230
            TabIndex        =   52
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   42
            Left            =   3510
            TabIndex        =   51
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   41
            Left            =   2790
            TabIndex        =   50
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   40
            Left            =   2070
            TabIndex        =   49
            Top             =   2190
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   39
            Left            =   8580
            TabIndex        =   48
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   38
            Left            =   7770
            TabIndex        =   47
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   37
            Left            =   7050
            TabIndex        =   46
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   36
            Left            =   6390
            TabIndex        =   45
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   35
            Left            =   5730
            TabIndex        =   44
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   34
            Left            =   4980
            TabIndex        =   43
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   33
            Left            =   4230
            TabIndex        =   42
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   32
            Left            =   3510
            TabIndex        =   41
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   31
            Left            =   2790
            TabIndex        =   40
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   30
            Left            =   2070
            TabIndex        =   39
            Top             =   1710
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   29
            Left            =   8580
            TabIndex        =   38
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   28
            Left            =   7770
            TabIndex        =   37
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   27
            Left            =   7050
            TabIndex        =   36
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   26
            Left            =   6390
            TabIndex        =   35
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   25
            Left            =   5730
            TabIndex        =   34
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   24
            Left            =   4980
            TabIndex        =   33
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   23
            Left            =   4230
            TabIndex        =   32
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   22
            Left            =   3510
            TabIndex        =   31
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   21
            Left            =   2790
            TabIndex        =   30
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   20
            Left            =   2070
            TabIndex        =   29
            Top             =   1230
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   19
            Left            =   8580
            TabIndex        =   28
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   18
            Left            =   7770
            TabIndex        =   27
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   17
            Left            =   7050
            TabIndex        =   26
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   16
            Left            =   6390
            TabIndex        =   25
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   15
            Left            =   5730
            TabIndex        =   24
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   14
            Left            =   4980
            TabIndex        =   23
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   13
            Left            =   4230
            TabIndex        =   22
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   12
            Left            =   3510
            TabIndex        =   21
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   11
            Left            =   2790
            TabIndex        =   20
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   10
            Left            =   2070
            TabIndex        =   19
            Top             =   750
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   9
            Left            =   8580
            TabIndex        =   13
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   8
            Left            =   7770
            TabIndex        =   12
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   7
            Left            =   7050
            TabIndex        =   11
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   6
            Left            =   6390
            TabIndex        =   10
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   5
            Left            =   5730
            TabIndex        =   9
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   4
            Left            =   4980
            TabIndex        =   8
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   3
            Left            =   4230
            TabIndex        =   7
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   2
            Left            =   3510
            TabIndex        =   6
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   5
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkMask 
            Height          =   315
            Index           =   0
            Left            =   2070
            TabIndex        =   4
            Top             =   270
            Width           =   615
         End
      End
      Begin prjFarmManagement.uctlTextBox txtNote1 
         Height          =   405
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   2475
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtNote2 
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Top             =   1380
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin VB.Label lblNote2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label lblNote1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1050
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmAddEditYCollection"
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

Private m_YCollection As CYCollection


Private Sub InitFormLayout()
Dim i As Long

   pnlHeader.Caption = HeaderText
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlFooter.BackColor = GLB_FORM_COLOR
      
   Call InitNormalLabel(lblNote1, "ชื่อ")
   Call InitNormalLabel(lblNote2, "รายละเอียด")

   Call txtNote1.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtNote2.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)

'   Call InitMainButton(cmdAdd, "เพิ่ม (F7)")
'   Call InitMainButton(cmdEdit, "แก้ไข (F3)")
'   Call InitMainButton(cmdDelete, "ลบ (F6)")
   
   Call InitMainButton(cmdOK, "ตกลง (F2)")
   Call InitMainButton(cmdCancel, "ยกเลิก (ESC)")
   
   For i = 0 To 99
      Call InitCheckBox(chkMask(i), i Mod 10)
   Next i
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = "รายการกลุ่มตัวเลข"
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

Private Sub chkMask_Click(Index As Integer)
Static InUsed As Long
Dim i As Long
Dim Flag As Boolean

   If InUsed Then
      Exit Sub
   End If
   
   InUsed = 1
   
   If Check2Flag(chkMask(Index).Value) = "Y" Then
      Flag = True
   Else
      Flag = False
   End If
   
   For i = 0 To 99
      If (i Mod 10) = (Index Mod 10) Then
         If i = Index Then
            If Flag Then
               chkMask(i).Enabled = Flag
            End If
         Else
            chkMask(i).Enabled = Not Flag
         End If
      End If
   Next i
   
   m_HasModify = True
   InUsed = 0
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
   
   If Not VerifyTextControl(lblNote1, txtNote1, False) Then
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
   
   m_YCollection.Y_COLLECTION_ID = ID
   m_YCollection.AddEditMode = ShowMode
    m_YCollection.Y_COLLECTION_NAME = txtNote1.Text
    m_YCollection.Y_COLLECTION_DESC = txtNote2.Text
    m_YCollection.MASK1 = CreateMask(1)
    m_YCollection.MASK2 = CreateMask(2)
    m_YCollection.MASK3 = CreateMask(3)
    m_YCollection.MASK4 = CreateMask(4)
    m_YCollection.MASK5 = CreateMask(5)
    m_YCollection.MASK6 = CreateMask(6)
    m_YCollection.MASK7 = CreateMask(7)
    m_YCollection.MASK8 = CreateMask(8)
    m_YCollection.MASK9 = CreateMask(9)
    m_YCollection.MASK10 = CreateMask(10)
    
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditYCollection(m_YCollection, IsOK, glbErrorLog) Then
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

Private Function CreateMask(Row As Long) As String
Dim j As Long
Dim i As Long
Dim TempStr As String

   TempStr = ""
   For j = 0 To 9
      i = (Row - 1) * 10 + j
      TempStr = TempStr & Check2Flag(chkMask(i))
   Next j
   CreateMask = TempStr
End Function

Private Sub ShowMask(Row As Long, Mask As String)
Dim j As Long
Dim i As Long
Dim TempStr As String

   TempStr = ""
   For j = 0 To 9
      i = (Row - 1) * 10 + j
      chkMask(i).Value = FlagToCheck(Mid(Mask, j + 1, 1))
   Next j
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
            
      m_YCollection.Y_COLLECTION_ID = ID
      m_YCollection.QueryFlag = 1
      If Not glbDaily.QueryYCollection(m_YCollection, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   Else
      IsOK = True
   End If
   
   If ItemCount > 0 Then
      txtNote1.Text = NVLS(m_Rs("Y_COLLECTION_NAME"), "")
      txtNote2.Text = NVLS(m_Rs("Y_COLLECTION_DESC"), "")
   
      Call ShowMask(1, NVLS(m_Rs("MASK1"), ""))
      Call ShowMask(2, NVLS(m_Rs("MASK2"), ""))
      Call ShowMask(3, NVLS(m_Rs("MASK3"), ""))
      Call ShowMask(4, NVLS(m_Rs("MASK4"), ""))
      Call ShowMask(5, NVLS(m_Rs("MASK5"), ""))
      Call ShowMask(6, NVLS(m_Rs("MASK6"), ""))
      Call ShowMask(7, NVLS(m_Rs("MASK7"), ""))
      Call ShowMask(8, NVLS(m_Rs("MASK8"), ""))
      Call ShowMask(9, NVLS(m_Rs("MASK9"), ""))
      Call ShowMask(10, NVLS(m_Rs("MASK10"), ""))
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
         m_YCollection.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
      End If
      
      TabStrip1_Click
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
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
   Set m_YCollection = New CYCollection
   Set m_Rs = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_YCollection = Nothing
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

Private Sub TabStrip1_Click()
   fraDrug.Visible = False
   fraDrug.BackColor = GLB_FORM_COLOR
   
   If TabStrip1.SelectedItem.Index = 1 Then
      fraDrug.Left = 0
      fraDrug.Top = 2700
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
