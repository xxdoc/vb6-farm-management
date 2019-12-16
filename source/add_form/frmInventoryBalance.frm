VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmInventoryBalance 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmInventoryBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4305
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   7594
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlDate1 
         Height          =   405
         Left            =   1890
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlTextLookup1 
         Height          =   435
         Left            =   1890
         TabIndex        =   1
         Top             =   1470
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   8
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1890
         TabIndex        =   3
         Top             =   2310
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextLookup uctlTextLookup2 
         Height          =   435
         Left            =   1890
         TabIndex        =   2
         Top             =   1890
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice 
         Height          =   465
         Left            =   1890
         TabIndex        =   4
         Top             =   2790
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   820
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   2910
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4080
         TabIndex        =   14
         Top             =   2940
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4080
         TabIndex        =   13
         Top             =   2460
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1110
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5235
         TabIndex        =   6
         Top             =   3480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3585
         TabIndex        =   5
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInventoryBalance.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInventoryBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BalanceAccum As CBalanceAccum

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_Locations As Collection
Private m_PartItems As Collection
Private m_InventoryBalances As Collection
Private m_BalanceAccumID As Long

Private Sub cmdPasswd_Click()

End Sub


Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboPosition_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
   End If
   
   If ItemCount > 0 Then
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("ADMIN_GROUP_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("ADMIN_GROUP_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
   
'   If Not VerifyCombo(lblPartType, cboPartType, False) Then
'      Exit Function
'   End If
   
   If Not VerifyDate(lblFileName, uctlDate1, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblMasterName, uctlTextLookup1.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblProgress, uctlTextLookup2.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPercent, txtPercent, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAvgPrice, txtAvgPrice, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_BalanceAccum.BALANCE_ACCUM_ID = m_BalanceAccumID
   m_BalanceAccum.AddEditMode = SHOW_EDIT
   m_BalanceAccum.BALANCE_AMOUNT = Val(txtPercent.Text)
   m_BalanceAccum.AVG_PRICE = Val(txtAvgPrice.Text)
         
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditBalanceAccum(m_BalanceAccum, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadLocation(uctlTextLookup1.MyCombo, m_Locations, 2, "N")
      Set uctlTextLookup1.MyCollection = m_Locations
      Call LoadPartItem(uctlTextLookup2.MyCombo, m_PartItems, , "N")
      Set uctlTextLookup2.MyCollection = m_PartItems
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
      End If
      
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "แก้ไขข้อมูลยอดวัตถุดิบ"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "ณ วันที่")
   Call InitNormalLabel(lblMasterName, "สถานที่จัดเก็บ")
   Call InitNormalLabel(lblProgress, "วัตถุดิบ")
   Call InitNormalLabel(lblPercent, "ปริมาณ")
   Call InitNormalLabel(Label1, "")
   Call InitNormalLabel(lblAvgPrice, "มูลค่า")
   Call InitNormalLabel(Label2, "บาท")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtAvgPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
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
   
   Set m_BalanceAccum = New CBalanceAccum
   Set m_Rs = New ADODB.Recordset
   
   Set m_Locations = New Collection
   Set m_PartItems = New Collection
   Set m_InventoryBalances = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Locations = Nothing
   Set m_PartItems = Nothing
   Set m_InventoryBalances = Nothing
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtPercent_GotFocus()
Dim NewDate As Date
Dim PartItemID As Long
Dim LocationID As Long
Dim Ba As CBalanceAccum

   If Not VerifyCombo(lblMasterName, uctlTextLookup1.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblProgress, uctlTextLookup2.MyCombo, False) Then
      Exit Sub
   End If

   LocationID = uctlTextLookup1.MyCombo.ItemData(Minus2Zero(uctlTextLookup1.MyCombo.ListIndex))
   PartItemID = uctlTextLookup2.MyCombo.ItemData(Minus2Zero(uctlTextLookup2.MyCombo.ListIndex))
   If uctlDate1.ShowDate <= 0 Then
      NewDate = DateAdd("D", 1, Now)
   Else
      NewDate = DateAdd("D", 1, uctlDate1.ShowDate)
   End If
   Call LoadInventoryBalanceEx(Nothing, m_InventoryBalances, NewDate, -1, "", LocationID, PartItemID)
   
   Set Ba = GetBalanceAccum(m_InventoryBalances, LocationID & "-" & PartItemID)
   txtPercent.Text = Ba.BALANCE_AMOUNT
   txtAvgPrice.Text = Ba.AVG_PRICE
   m_BalanceAccumID = Ba.BALANCE_ACCUM_ID
End Sub

Private Sub uctlDate1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextLookup2_Change()
   m_HasModify = True
End Sub
