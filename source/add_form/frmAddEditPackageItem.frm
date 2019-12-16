VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPackageItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPackageItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6588
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtFromWeight 
         Height          =   435
         Left            =   2880
         TabIndex        =   1
         Top             =   750
         Width           =   2025
         _extentx        =   3572
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlStatusBuy 
         Height          =   435
         Left            =   2880
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _extentx        =   9446
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtToWeight 
         Height          =   435
         Left            =   7320
         TabIndex        =   2
         Top             =   750
         Width           =   1995
         _extentx        =   3519
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCostPerWeight 
         Height          =   435
         Left            =   2880
         TabIndex        =   5
         Top             =   1680
         Width           =   2025
         _extentx        =   3572
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCutWeight 
         Height          =   435
         Left            =   2880
         TabIndex        =   3
         Top             =   1230
         Width           =   2025
         _extentx        =   3572
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCostPerExceed 
         Height          =   435
         Left            =   7320
         TabIndex        =   4
         Top             =   1230
         Width           =   2025
         _extentx        =   3572
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCostPerUnit 
         Height          =   435
         Left            =   2880
         TabIndex        =   6
         Top             =   2160
         Width           =   2025
         _extentx        =   3572
         _extenty        =   767
      End
      Begin VB.Label Label10 
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label Label8 
         Height          =   375
         Left            =   4920
         TabIndex        =   23
         Top             =   1830
         Width           =   765
      End
      Begin VB.Label Label7 
         Height          =   375
         Left            =   8280
         TabIndex        =   22
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label6 
         Height          =   375
         Left            =   8280
         TabIndex        =   21
         Top             =   870
         Width           =   525
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   4950
         TabIndex        =   19
         Top             =   870
         Width           =   525
      End
      Begin VB.Label lblCostPerExceed 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5715
         TabIndex        =   18
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblCostPerUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   2280
         Width           =   2325
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   3345
         TabIndex        =   7
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackageItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblCutWeight 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1335
         TabIndex        =   16
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label lblCostPerWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Label lblToWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5730
         TabIndex        =   14
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label lblStatusBuy 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1290
         TabIndex        =   13
         Top             =   360
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4995
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackageItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6645
         TabIndex        =   9
         Top             =   2880
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1275
         TabIndex        =   12
         Top             =   840
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditPackageItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public PackageType  As Long

Private m_StatusBuy As Collection

Public ParentForm As Form

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblStatusBuy, MapText("สถานะ"))
   Call InitNormalLabel(lblFromWeight, MapText("จาก นน."))
   Call InitNormalLabel(lblToWeight, MapText("ถึง นน."))
   Call InitNormalLabel(lblCutWeight, MapText("ตัดที่ นน."))
   Call InitNormalLabel(lblCostPerExceed, MapText("ราคา/ส่วนเกิน"))
   Call InitNormalLabel(lblCostPerWeight, MapText("ราคา/นน."))
   
   Call InitNormalLabel(lblCostPerUnit, MapText("ค่าคงที่"))
  
   Call InitNormalLabel(Label3, MapText("กก."))
   Call InitNormalLabel(Label4, MapText("กก."))
   Call InitNormalLabel(Label6, MapText("กก."))
   Call InitNormalLabel(Label7, MapText("บาท/กก."))
   Call InitNormalLabel(Label8, MapText("บาท"))
   Call InitNormalLabel(Label10, MapText("บาท"))
   
   Call txtFromWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtToWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtCutWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
    Call txtCostPerExceed.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
    Call txtCostPerWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtCostPerUnit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   
   Call SetEnableDisableTextBox
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOk As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CPackageDetail
         
         Set Di = TempCollection.Item(ID)
         
         uctlStatusBuy.MyCombo.ListIndex = IDToListIndex(uctlStatusBuy.MyCombo, Di.STATUS_BUY_ID)
         txtFromWeight.Text = Di.FROM_WEIGHT
         txtToWeight.Text = Di.TO_WEIGHT
         txtCutWeight.Text = Di.CUT_WEIGHT
         txtCostPerExceed.Text = Di.COST_PER_EXCEED
         txtCostPerWeight.Text = Di.COST_PER_WEIGHT
         txtCostPerUnit.Text = Di.COST_PER_UNIT
         
         
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim I As Long

   I = 0
   For Each O In Col
      I = I + 1
      If (I > OldID) And (O.Flag <> "D") Then
         GetNextID = I
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function


Private Sub cmdNext_Click()
Dim NewID As Long
   If Not SaveData Then
      Exit Sub
   End If
   
   Call ParentForm.ShowPackageItemGrid
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.ShowPackageItemGrid
         Exit Sub
      End If

      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
        uctlStatusBuy.MyCombo.ListIndex = -1
        txtFromWeight.Text = ""
        txtToWeight.Text = ""
        txtCutWeight.Text = ""
        txtCostPerExceed.Text = ""
        txtCostPerWeight.Text = ""
        txtCostPerUnit.Text = ""
   End If
   Call QueryData(True)
   Call ParentForm.ShowPackageItemGrid
   uctlStatusBuy.SetFocus
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOk As Boolean
Dim RealIndex As Long
Dim I As Long
   If Not VerifyCombo(lblStatusBuy, uctlStatusBuy.MyCombo, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   
   Dim Di As CPackageDetail
   Dim CheckDetail As CPackageDetail
   I = 0
'   If Not (PackageType = 5) And Not (PackageType = 6) Then
'      For Each CheckDetail In TempCollection
'      I = I + 1
'
'      If CheckDetail.STATUS_BUY_ID = uctlStatusBuy.MyCombo.ItemData(Minus2Zero(uctlStatusBuy.MyCombo.ListIndex)) And ID <> I Then
'         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & uctlStatusBuy.MyCombo.Text & " " & MapText("อยู่ในระบบแล้ว")
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If
'      Next
'   End If

   If ShowMode = SHOW_ADD Then
      Set Di = New CPackageDetail
      
      Di.Flag = "A"
      Call TempCollection.Add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If


   Di.STATUS_BUY_ID = uctlStatusBuy.MyCombo.ItemData(Minus2Zero(uctlStatusBuy.MyCombo.ListIndex))
   Di.PRODUCT_STATUS_NAME = uctlStatusBuy.MyCombo.Text
   Di.FROM_WEIGHT = Val(txtFromWeight.Text)
   Di.TO_WEIGHT = Val(txtToWeight.Text)
   Di.CUT_WEIGHT = Val(txtCutWeight.Text)
   Di.COST_PER_EXCEED = Val(txtCostPerExceed.Text)
   Di.PEDIGREE_COST = 0
   Di.COST_PER_WEIGHT = Val(txtCostPerWeight.Text)
   Di.COST_PER_UNIT = Val(txtCostPerUnit.Text)
   
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadProductStatus(uctlStatusBuy.MyCombo, m_StatusBuy)
      Set uctlStatusBuy.MyCollection = m_StatusBuy
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
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
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_StatusBuy = New Collection
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_StatusBuy = Nothing
   
End Sub

Private Sub txtCostPerExceed_Change()
   m_HasModify = True
End Sub

Private Sub txtCostPerUnit_Change()
   m_HasModify = True
End Sub

Private Sub txtCostPerWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtCutWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtFromWeight_Change()
   m_HasModify = True
End Sub
Private Sub txtToWeight_Change()
   m_HasModify = True
End Sub

Private Sub uctlStatusBuy_Change()
   m_HasModify = True
End Sub

Private Sub SetEnableDisableTextBox()
   txtFromWeight.Enabled = False
   txtToWeight.Enabled = False
   txtCutWeight.Enabled = False
   txtCostPerExceed.Enabled = False
   txtCostPerWeight.Enabled = False
   txtCostPerUnit.Enabled = False
   
   
   If PackageType = 4 Then
      txtFromWeight.Enabled = True
      txtToWeight.Enabled = True
      txtCostPerUnit.Enabled = True
      txtCostPerWeight.Enabled = True
      txtCutWeight.Enabled = True
      txtCostPerExceed.Enabled = True
   ElseIf PackageType = 5 Then
      txtFromWeight.Enabled = True
      txtToWeight.Enabled = True
      txtCostPerWeight.Enabled = True
   End If
   
End Sub
