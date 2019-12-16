VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmUpdateT706 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   Icon            =   "frmUpdateT706.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11820
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   465
         Left            =   30
         TabIndex        =   12
         Top             =   0
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   820
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPrice1 
         Height          =   465
         Left            =   4800
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAmount1 
         Height          =   465
         Left            =   6600
         TabIndex        =   5
         Top             =   1320
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumWeight1 
         Height          =   465
         Left            =   7440
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
         _ExtentX        =   1720
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAvg1 
         Height          =   465
         Left            =   9120
         TabIndex        =   7
         Top             =   1320
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumPrice1 
         Height          =   465
         Left            =   10080
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtStatus1 
         Height          =   465
         Left            =   120
         TabIndex        =   0
         Top             =   1320
         Width           =   735
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAmount1 
         Height          =   465
         Left            =   1200
         TabIndex        =   1
         Top             =   1320
         Width           =   855
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtWeight1 
         Height          =   465
         Left            =   2160
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAvg1 
         Height          =   465
         Left            =   3480
         TabIndex        =   3
         Top             =   1320
         Width           =   975
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtPrice2 
         Height          =   465
         Left            =   4800
         TabIndex        =   27
         Top             =   2040
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAmount2 
         Height          =   465
         Left            =   6600
         TabIndex        =   28
         Top             =   2040
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumWeight2 
         Height          =   465
         Left            =   7440
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
         _ExtentX        =   1720
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAvg2 
         Height          =   465
         Left            =   9120
         TabIndex        =   30
         Top             =   2040
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumPrice2 
         Height          =   465
         Left            =   10080
         TabIndex        =   31
         Top             =   2040
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtStatus2 
         Height          =   465
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Width           =   735
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAmount2 
         Height          =   465
         Left            =   1200
         TabIndex        =   33
         Top             =   2040
         Width           =   855
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtWeight2 
         Height          =   465
         Left            =   2160
         TabIndex        =   34
         Top             =   2040
         Width           =   1215
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAvg2 
         Height          =   465
         Left            =   3480
         TabIndex        =   35
         Top             =   2040
         Width           =   975
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtPrice3 
         Height          =   465
         Left            =   4800
         TabIndex        =   36
         Top             =   2760
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAmount3 
         Height          =   465
         Left            =   6600
         TabIndex        =   37
         Top             =   2760
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumWeight3 
         Height          =   465
         Left            =   7440
         TabIndex        =   38
         Top             =   2760
         Width           =   1215
         _ExtentX        =   1720
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAvg3 
         Height          =   465
         Left            =   9120
         TabIndex        =   39
         Top             =   2760
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumPrice3 
         Height          =   465
         Left            =   10080
         TabIndex        =   40
         Top             =   2760
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtStatus3 
         Height          =   465
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         Width           =   735
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAmount3 
         Height          =   465
         Left            =   1200
         TabIndex        =   42
         Top             =   2760
         Width           =   855
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtWeight3 
         Height          =   465
         Left            =   2160
         TabIndex        =   43
         Top             =   2760
         Width           =   1215
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAvg3 
         Height          =   465
         Left            =   3480
         TabIndex        =   44
         Top             =   2760
         Width           =   975
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtPrice4 
         Height          =   465
         Left            =   4800
         TabIndex        =   45
         Top             =   3360
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAmount4 
         Height          =   465
         Left            =   6600
         TabIndex        =   46
         Top             =   3360
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumWeight4 
         Height          =   465
         Left            =   7440
         TabIndex        =   47
         Top             =   3360
         Width           =   1215
         _ExtentX        =   1720
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAvg4 
         Height          =   465
         Left            =   9120
         TabIndex        =   48
         Top             =   3360
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumPrice4 
         Height          =   465
         Left            =   10080
         TabIndex        =   49
         Top             =   3360
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtStatus4 
         Height          =   465
         Left            =   120
         TabIndex        =   50
         Top             =   3360
         Width           =   735
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAmount4 
         Height          =   465
         Left            =   1200
         TabIndex        =   51
         Top             =   3360
         Width           =   855
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtWeight4 
         Height          =   465
         Left            =   2160
         TabIndex        =   52
         Top             =   3360
         Width           =   1215
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAvg4 
         Height          =   465
         Left            =   3480
         TabIndex        =   53
         Top             =   3360
         Width           =   975
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtPrice5 
         Height          =   465
         Left            =   4800
         TabIndex        =   54
         Top             =   3960
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAmount5 
         Height          =   465
         Left            =   6600
         TabIndex        =   55
         Top             =   3960
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumWeight5 
         Height          =   465
         Left            =   7440
         TabIndex        =   56
         Top             =   3960
         Width           =   1215
         _ExtentX        =   1720
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumAvg5 
         Height          =   465
         Left            =   9120
         TabIndex        =   57
         Top             =   3960
         Width           =   855
         _ExtentX        =   1085
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtSumPrice5 
         Height          =   465
         Left            =   10080
         TabIndex        =   58
         Top             =   3960
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtStatus5 
         Height          =   465
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   735
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAmount5 
         Height          =   465
         Left            =   1200
         TabIndex        =   24
         Top             =   3960
         Width           =   855
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtWeight5 
         Height          =   465
         Left            =   2160
         TabIndex        =   25
         Top             =   3960
         Width           =   1215
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtAvg5 
         Height          =   465
         Left            =   3480
         TabIndex        =   26
         Top             =   3960
         Width           =   975
         _ExtentX        =   1931
         _ExtentY        =   820
      End
      Begin VB.Label lblSumPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   10320
         TabIndex        =   22
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblSumAvg 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   9240
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblSumWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7800
         TabIndex        =   20
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblSumAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6840
         TabIndex        =   19
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblAvg 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3720
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2280
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   1635
         Left            =   120
         TabIndex        =   14
         Top             =   5280
         Width           =   11535
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   4320
         TabIndex        =   9
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmUpdateT706.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6120
         TabIndex        =   10
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmUpdateT706"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cmdStart_Click()
Dim IsOK As Boolean
   'Save To T706Collection1
   'Save To T706Collection2
   
   Set T706Collection1 = Nothing
   Set T706Collection2 = Nothing
   Set T706Collection1 = New Collection
   Set T706Collection2 = New Collection
   '1
   '-------------------------------------------------------------------------------------------------------------------------------------
   Dim TempExport As CExportItem
   If Len(txtStatus1.Text) > 0 Then
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus1.Text
      TempExport.EXPORT_AMOUNT = Val(txtAmount1.Text)
      TempExport.TOTAL_WEIGHT = Val(txtWeight1.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtAvg1.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtPrice1.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight1.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection1.Add(TempExport, Trim(txtStatus1.Text))
      
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus1.Text
      TempExport.EXPORT_AMOUNT = Val(txtSumAmount1.Text)
      TempExport.TOTAL_WEIGHT = Val(txtSumWeight1.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtSumAvg1.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtSumPrice1.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight1.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection2.Add(TempExport, Trim(txtStatus1.Text))
   End If
   '-------------------------------------------------------------------------------------------------------------------------------------
   
   '2
   '-------------------------------------------------------------------------------------------------------------------------------------
   If Len(txtStatus2.Text) > 0 Then
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus2.Text
      TempExport.EXPORT_AMOUNT = Val(txtAmount2.Text)
      TempExport.TOTAL_WEIGHT = Val(txtWeight2.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtAvg2.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtPrice2.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight2.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection1.Add(TempExport, Trim(txtStatus2.Text))
      
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus2.Text
      TempExport.EXPORT_AMOUNT = Val(txtSumAmount2.Text)
      TempExport.TOTAL_WEIGHT = Val(txtSumWeight2.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtSumAvg2.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtSumPrice2.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight2.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection2.Add(TempExport, Trim(txtStatus2.Text))
   End If
   '-------------------------------------------------------------------------------------------------------------------------------------
   
   '3
   '-------------------------------------------------------------------------------------------------------------------------------------
   If Len(txtStatus3.Text) > 0 Then
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus3.Text
      TempExport.EXPORT_AMOUNT = Val(txtAmount3.Text)
      TempExport.TOTAL_WEIGHT = Val(txtWeight3.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtAvg3.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtPrice3.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight3.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection1.Add(TempExport, Trim(txtStatus3.Text))
      
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus3.Text
      TempExport.EXPORT_AMOUNT = Val(txtSumAmount3.Text)
      TempExport.TOTAL_WEIGHT = Val(txtSumWeight3.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtSumAvg3.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtSumPrice3.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight3.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection2.Add(TempExport, Trim(txtStatus3.Text))
   End If
   '-------------------------------------------------------------------------------------------------------------------------------------
   
   '4
   '-------------------------------------------------------------------------------------------------------------------------------------
   If Len(txtStatus4.Text) > 0 Then
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus4.Text
      TempExport.EXPORT_AMOUNT = Val(txtAmount4.Text)
      TempExport.TOTAL_WEIGHT = Val(txtWeight4.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtAvg4.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtPrice4.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight4.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection1.Add(TempExport, Trim(txtStatus4.Text))
      
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus4.Text
      TempExport.EXPORT_AMOUNT = Val(txtSumAmount4.Text)
      TempExport.TOTAL_WEIGHT = Val(txtSumWeight4.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtSumAvg4.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtSumPrice4.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight4.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection2.Add(TempExport, Trim(txtStatus4.Text))
   End If
   '-------------------------------------------------------------------------------------------------------------------------------------
   
   '5
   '-------------------------------------------------------------------------------------------------------------------------------------
   If Len(txtStatus5.Text) > 0 Then
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus5.Text
      TempExport.EXPORT_AMOUNT = Val(txtAmount5.Text)
      TempExport.TOTAL_WEIGHT = Val(txtWeight5.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtAvg5.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtPrice5.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight5.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection1.Add(TempExport, Trim(txtStatus5.Text))
      
      Set TempExport = New CExportItem
      TempExport.PIG_STATUS_NO = txtStatus5.Text
      TempExport.EXPORT_AMOUNT = Val(txtSumAmount5.Text)
      TempExport.TOTAL_WEIGHT = Val(txtSumWeight5.Text)
      TempExport.EXPORT_AVG_PRICE = Val(txtSumAvg5.Text)
      TempExport.EXPORT_TOTAL_PRICE = Val(txtSumPrice5.Text)
      'นน ต่อตัว = TempExport.TOTAL_WEIGHT = txtWeight5.Text  / TempExport.EXPORT_AMOUNT
      Call T706Collection2.Add(TempExport, Trim(txtStatus5.Text))
   End If
   '-------------------------------------------------------------------------------------------------------------------------------------
   Set TempExport = Nothing
   
   Unload Me
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call QueryData
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   End If
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "แก้ไขรายงานตรงยอดขายสุกรฟาร์ม"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblStatus, "สถ")
   
   Call InitNormalLabel(lblAmount, "จน")
   Call InitNormalLabel(lblWeight, "นน")
   Call InitNormalLabel(lblAvg, "เฉลี่ย")
   Call InitNormalLabel(lblPrice, "มูลค่า")
   
   Call InitNormalLabel(lblSumAmount, "จน")
   Call InitNormalLabel(lblSumWeight, "นน")
   Call InitNormalLabel(lblSumAvg, "เฉลี่ย")
   Call InitNormalLabel(lblSumPrice, "มูลค่า")
   
   Call InitNormalLabel(lblNote, "เป็นส่วนแก้ไขชั่วคราวเพื่อออกรายงาน T706โดยเมื่อปิดโปรแกรมสิ่งที่ Key ไว้จากหายไป")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("ตกลง"))
   
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub QueryData()
Dim TempExport As CExportItem
Dim I As Integer
   I = 0
   For Each TempExport In T706Collection1
      I = I + 1
      If I = 1 Then
         txtStatus1.Text = TempExport.PIG_STATUS_NO
         txtAmount1.Text = TempExport.EXPORT_AMOUNT
         txtWeight1.Text = TempExport.TOTAL_WEIGHT
         txtAvg1.Text = TempExport.EXPORT_AVG_PRICE
         txtPrice1.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 2 Then
         txtStatus2.Text = TempExport.PIG_STATUS_NO
         txtAmount2.Text = TempExport.EXPORT_AMOUNT
         txtWeight2.Text = TempExport.TOTAL_WEIGHT
         txtAvg2.Text = TempExport.EXPORT_AVG_PRICE
         txtPrice2.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 3 Then
         txtStatus3.Text = TempExport.PIG_STATUS_NO
         txtAmount3.Text = TempExport.EXPORT_AMOUNT
         txtWeight3.Text = TempExport.TOTAL_WEIGHT
         txtAvg3.Text = TempExport.EXPORT_AVG_PRICE
         txtPrice3.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 4 Then
         txtStatus4.Text = TempExport.PIG_STATUS_NO
         txtAmount4.Text = TempExport.EXPORT_AMOUNT
         txtWeight4.Text = TempExport.TOTAL_WEIGHT
         txtAvg4.Text = TempExport.EXPORT_AVG_PRICE
         txtPrice4.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 5 Then
         txtStatus5.Text = TempExport.PIG_STATUS_NO
         txtAmount5.Text = TempExport.EXPORT_AMOUNT
         txtWeight5.Text = TempExport.TOTAL_WEIGHT
         txtAvg5.Text = TempExport.EXPORT_AVG_PRICE
         txtPrice5.Text = TempExport.EXPORT_TOTAL_PRICE
      End If
   Next TempExport
   
   I = 0
   For Each TempExport In T706Collection2
      I = I + 1
      If I = 1 Then
         txtSumAmount1.Text = TempExport.EXPORT_AMOUNT
         txtSumWeight1.Text = TempExport.TOTAL_WEIGHT
         txtSumAvg1.Text = TempExport.EXPORT_AVG_PRICE
         txtSumPrice1.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 2 Then
         txtSumAmount2.Text = TempExport.EXPORT_AMOUNT
         txtSumWeight2.Text = TempExport.TOTAL_WEIGHT
         txtSumAvg2.Text = TempExport.EXPORT_AVG_PRICE
         txtSumPrice2.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 3 Then
         txtSumAmount3.Text = TempExport.EXPORT_AMOUNT
         txtSumWeight3.Text = TempExport.TOTAL_WEIGHT
         txtSumAvg3.Text = TempExport.EXPORT_AVG_PRICE
         txtSumPrice3.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 4 Then
         txtSumAmount4.Text = TempExport.EXPORT_AMOUNT
         txtSumWeight4.Text = TempExport.TOTAL_WEIGHT
         txtSumAvg4.Text = TempExport.EXPORT_AVG_PRICE
         txtSumPrice4.Text = TempExport.EXPORT_TOTAL_PRICE
      ElseIf I = 5 Then
         txtSumAmount5.Text = TempExport.EXPORT_AMOUNT
         txtSumWeight5.Text = TempExport.TOTAL_WEIGHT
         txtSumAvg5.Text = TempExport.EXPORT_AVG_PRICE
         txtSumPrice5.Text = TempExport.EXPORT_TOTAL_PRICE
      End If
   Next TempExport
End Sub
