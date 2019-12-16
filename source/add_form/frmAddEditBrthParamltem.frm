VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBrthParamItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditBrthParamltem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7125
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   12568
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPigGLookup 
         Height          =   405
         Left            =   1710
         TabIndex        =   2
         Top             =   1170
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFromBreed 
         Height          =   405
         Left            =   1710
         TabIndex        =   0
         Top             =   270
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtBreedAmount 
         Height          =   435
         Left            =   1710
         TabIndex        =   6
         Top             =   2970
         Width           =   1935
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlToBreed 
         Height          =   405
         Left            =   1710
         TabIndex        =   1
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtBreedRate 
         Height          =   435
         Left            =   6270
         TabIndex        =   7
         Top             =   2970
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtChildRate 
         Height          =   435
         Left            =   1710
         TabIndex        =   8
         Top             =   3420
         Width           =   1935
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBirthAmount 
         Height          =   435
         Left            =   6270
         TabIndex        =   9
         Top             =   3420
         Width           =   1935
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFromBirth 
         Height          =   405
         Left            =   1710
         TabIndex        =   10
         Top             =   3870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToBirth 
         Height          =   405
         Left            =   1710
         TabIndex        =   11
         Top             =   4290
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDayCount 
         Height          =   435
         Left            =   1710
         TabIndex        =   12
         Top             =   4710
         Width           =   1935
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBirthRate 
         Height          =   435
         Left            =   6270
         TabIndex        =   13
         Top             =   4710
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgWeight 
         Height          =   435
         Left            =   1710
         TabIndex        =   14
         Top             =   5160
         Width           =   1935
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigLLookup 
         Height          =   405
         Left            =   1710
         TabIndex        =   3
         Top             =   1620
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPigGLAmount 
         Height          =   435
         Left            =   1710
         TabIndex        =   4
         Top             =   2070
         Width           =   1485
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBreedPercent 
         Height          =   435
         Left            =   1710
         TabIndex        =   5
         Top             =   2520
         Width           =   1485
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPigGAmount 
         Height          =   435
         Left            =   1710
         TabIndex        =   15
         Top             =   5610
         Width           =   1485
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPigLAmount 
         Height          =   435
         Left            =   6270
         TabIndex        =   16
         Top             =   5580
         Width           =   1485
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtBirthCost 
         Height          =   435
         Left            =   6270
         TabIndex        =   39
         Top             =   5160
         Width           =   1485
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin VB.Label lblBirthCost 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4680
         TabIndex        =   40
         Top             =   5190
         Width           =   1575
      End
      Begin VB.Label lblPigLAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4680
         TabIndex        =   38
         Top             =   5610
         Width           =   1575
      End
      Begin VB.Label lblPigGAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label lblBreedPercent 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   30
         TabIndex        =   36
         Top             =   2550
         Width           =   1575
      End
      Begin VB.Label lblPigGLAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Label lblPigG 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   34
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblPigL 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   30
         TabIndex        =   33
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label lblAvgWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   5220
         Width           =   1485
      End
      Begin VB.Label lblBirthRate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4410
         TabIndex        =   31
         Top             =   4770
         Width           =   1785
      End
      Begin VB.Label lblDayCount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   4770
         Width           =   1485
      End
      Begin VB.Label lblToBirth 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   4290
         Width           =   1485
      End
      Begin VB.Label lblFromBirth 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   3840
         Width           =   1485
      End
      Begin VB.Label lblBirthAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4380
         TabIndex        =   27
         Top             =   3480
         Width           =   1785
      End
      Begin VB.Label lblChildRate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label lblBreedRate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4410
         TabIndex        =   25
         Top             =   3030
         Width           =   1785
      End
      Begin VB.Label lblFromBreed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblToBreed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2040
         TabIndex        =   17
         Top             =   6270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3690
         TabIndex        =   18
         Top             =   6270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5340
         TabIndex        =   19
         Top             =   6270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblBreedAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3030
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditBrthParamItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private m_PartTypes As Collection
Private m_PartItems As Collection
Private m_pigGs As Collection
Private m_pigLs As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboDrCr_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboDrCr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

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
   
   Call InitNormalLabel(lblFromBreed, MapText("จากวันที่ผสม"))
   Call InitNormalLabel(lblToBreed, MapText("ถึงวันที่ผสม"))
   Call InitNormalLabel(lblBreedAmount, MapText("จำนวนผสม"))
   Call InitNormalLabel(lblBreedRate, MapText("% การเข้าคลอด"))
   Call InitNormalLabel(lblChildRate, MapText("อัตรส่วนลูก/แม่"))
   Call InitNormalLabel(lblBirthAmount, MapText("จำนวนลูกเกิด"))
   Call InitNormalLabel(lblFromBirth, MapText("จากวันที่เกิด"))
   Call InitNormalLabel(lblToBirth, MapText("ถึงวันที่เกิด"))
   Call InitNormalLabel(lblDayCount, MapText("จำนวนวัน"))
   Call InitNormalLabel(lblBirthRate, MapText("ลูกเกิด/วัน"))
   Call InitNormalLabel(lblAvgWeight, MapText("น้ำหนักเฉลี่ย"))
   Call InitNormalLabel(lblPigG, MapText("สุกร G"))
   Call InitNormalLabel(lblPigL, MapText("สุกร L"))
   Call InitNormalLabel(lblPigGLAmount, MapText("จำนวน G+L"))
   Call InitNormalLabel(lblBreedPercent, MapText("% ผสม"))
   Call InitNormalLabel(lblPigGAmount, MapText("จำนวน G"))
   Call InitNormalLabel(lblPigLAmount, MapText("จำนวน L"))
   Call InitNormalLabel(lblBirthCost, MapText("ต้นทุนเกิด/ตัว"))
   
   Call txtBreedAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtBreedRate.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtChildRate.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtBirthAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   txtBirthAmount.Enabled = False
   Call txtDayCount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtDayCount.Enabled = False
   Call txtBirthRate.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
'   txtBirthRate.Enabled = False
   Call txtAvgWeight.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtPigGLAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtBreedPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtPigGAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtPigLAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   uctlFromBirth.Enable = False
   uctlToBirth.Enable = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Ji As CBrtPrmItem

         Set Ji = TempCollection.Item(ID)
         uctlFromBreed.ShowDate = Ji.GetFieldValue("FROM_BREED")
         uctlToBreed.ShowDate = Ji.GetFieldValue("TO_BREED")
         uctlFromBirth.ShowDate = Ji.GetFieldValue("FROM_BIRTH")
         uctlToBirth.ShowDate = Ji.GetFieldValue("TO_BIRTH")
         txtBreedAmount.Text = Ji.GetFieldValue("BREED_AMOUNT")
          txtBreedRate.Text = Ji.GetFieldValue("BREED_RATE")
         txtChildRate.Text = Ji.GetFieldValue("CHILD_RATE")
         txtBirthAmount.Text = Ji.GetFieldValue("BIRTH_AMOUNT")
         txtBirthRate.Text = Ji.GetFieldValue("BIRTH_RATE")
         txtDayCount.Text = Ji.GetFieldValue("DAY_COUNT")
         txtAvgWeight.Text = Ji.GetFieldValue("AVG_WEIGHT")
         
         uctlPigLLookup.MyCombo.ListIndex = IDToListIndex(uctlPigLLookup.MyCombo, Ji.GetFieldValue("PIGL_ID"))
         uctlPigGLookup.MyCombo.ListIndex = IDToListIndex(uctlPigGLookup.MyCombo, Ji.GetFieldValue("PIGG_ID"))
         txtPigGLAmount.Text = Ji.GetFieldValue("PIGGL_AMOUNT")
         txtBreedPercent.Text = Ji.GetFieldValue("BREED_PERCENT")
         txtPigGAmount.Text = Ji.GetFieldValue("PIGG_AMOUNT")
         txtPigLAmount.Text = Ji.GetFieldValue("PIGL_AMOUNT")
         txtBirthCost.Text = Ji.GetFieldValue("BIRTH_COST")
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If

   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid(True)
         Exit Sub
      End If

      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      uctlFromBreed.ShowDate = -1
      uctlToBreed.ShowDate = -1
      txtBreedAmount.Text = ""
      txtBreedRate.Text = ""
      txtChildRate.Text = ""
      txtBirthAmount.Text = ""
      uctlFromBirth.ShowDate = -1
      uctlToBirth.ShowDate = -1
      txtDayCount.Text = ""
      txtBirthRate.Text = ""
      txtAvgWeight.Text = ""
   End If
   Call QueryData(True)
   Call ParentForm.RefreshGrid(True)
   
   uctlFromBreed.SetFocus
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyDate(lblFromBreed, uctlFromBreed, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblToBreed, uctlToBreed, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigG, uctlPigGLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigL, uctlPigLLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPigGLAmount, txtPigGLAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBreedPercent, txtBreedPercent, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBreedAmount, txtBreedAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBreedRate, txtBreedRate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblChildRate, txtChildRate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBirthAmount, txtBirthAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPigGAmount, txtPigGAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPigLAmount, txtPigLAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ji As CBrtPrmItem
   If ShowMode = SHOW_ADD Then
      Set Ji = New CBrtPrmItem
      Ji.Flag = "A"
      Call TempCollection.Add(Ji)
   Else
      Set Ji = TempCollection.Item(ID)
      If Ji.Flag <> "A" Then
         Ji.Flag = "E"
      End If
   End If

   Call Ji.SetFieldValue("FROM_BREED", uctlFromBreed.ShowDate)
   Call Ji.SetFieldValue("TO_BREED", uctlToBreed.ShowDate)
   Call Ji.SetFieldValue("BREED_AMOUNT", txtBreedAmount.Text)
   Call Ji.SetFieldValue("BREED_RATE", txtBreedRate.Text)
   Call Ji.SetFieldValue("CHILD_RATE", txtChildRate.Text)
   Call Ji.SetFieldValue("BIRTH_AMOUNT", txtBirthAmount.Text)
   Call Ji.SetFieldValue("FROM_BIRTH", uctlFromBirth.ShowDate)
   Call Ji.SetFieldValue("TO_BIRTH", uctlToBirth.ShowDate)
   Call Ji.SetFieldValue("BIRTH_RATE", txtBirthRate.Text)
   Call Ji.SetFieldValue("DAY_COUNT", txtDayCount.Text)
   Call Ji.SetFieldValue("AVG_WEIGHT", Val(txtAvgWeight.Text))
   Call Ji.SetFieldValue("PIGG_ID", uctlPigGLookup.MyCombo.ItemData(Minus2Zero(uctlPigGLookup.MyCombo.ListIndex)))
   Call Ji.SetFieldValue("PIGL_ID", uctlPigLLookup.MyCombo.ItemData(Minus2Zero(uctlPigLLookup.MyCombo.ListIndex)))
   Call Ji.SetFieldValue("PIGG_AMOUNT", Val(txtPigGAmount.Text))
   Call Ji.SetFieldValue("PIGL_AMOUNT", Val(txtPigLAmount.Text))
   Call Ji.SetFieldValue("PIGGL_AMOUNT", Val(txtPigGLAmount.Text))
   Call Ji.SetFieldValue("BREED_PERCENT", Val(txtBreedPercent.Text))
   Call Ji.SetFieldValue("BIRTH_COST", Val(txtBirthCost.Text))
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartItem(uctlPigGLookup.MyCombo, m_pigGs, , "Y", "G")
      Set uctlPigGLookup.MyCollection = m_pigGs
      Call LoadPartItem(uctlPigLLookup.MyCombo, m_pigLs, , "Y", "L")
      Set uctlPigLLookup.MyCollection = m_pigLs
      
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
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_PartItems = New Collection
   Set m_pigGs = New Collection
   Set m_pigLs = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_pigGs = Nothing
   Set m_pigLs = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txtAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtMoo_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtVillage_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlAccountLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtAvgWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtBirthAmount_Change()
   m_HasModify = True
   txtBirthRate.Text = MyDiffEx(Val(txtBirthAmount.Text), Val(txtDayCount.Text))
End Sub

Private Sub txtBirthCost_Change()
   m_HasModify = True
End Sub

Private Sub txtBirthRate_Change()
   m_HasModify = True
End Sub

Private Sub txtBreedAmount_Change()
   m_HasModify = True
   
   txtBirthAmount.Text = Val(txtChildRate.Text) * Val(txtBreedAmount.Text) * (Val(txtBreedRate.Text) / 100)
   txtPigLAmount.Text = Val(txtBreedAmount.Text) * (Val(txtBreedRate.Text) / 100)
End Sub

Private Sub txtBreedPercent_Change()
   m_HasModify = True
   txtBreedAmount.Text = Val(txtPigGLAmount.Text) * (Val(txtBreedPercent.Text) / 100)
End Sub

Private Sub txtBreedRate_Change()
   m_HasModify = True
   txtBirthAmount.Text = Val(txtChildRate.Text) * Val(txtBreedAmount.Text) * (Val(txtBreedRate.Text) / 100)
   txtPigLAmount.Text = Val(txtBreedAmount.Text) * (Val(txtBreedRate.Text) / 100)
   txtPigGAmount.Text = Val(txtPigGLAmount.Text) - Val(txtBreedAmount.Text) * (Val(txtBreedRate.Text) / 100)
End Sub

Private Sub txtChildRate_Change()
   m_HasModify = True
   txtBirthAmount.Text = Val(txtChildRate.Text) * Val(txtBreedAmount.Text) * (Val(txtBreedRate.Text) / 100)
End Sub

Private Sub txtDayCount_Change()
   m_HasModify = True
   txtBirthRate.Text = MyDiffEx(Val(txtBirthAmount.Text), Val(txtDayCount.Text))
End Sub

Private Sub txtPigGAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtPigGLAmount_Change()
   m_HasModify = True
   txtBreedAmount.Text = Val(txtPigGLAmount.Text) * (Val(txtBreedPercent.Text) / 100)
   txtPigGAmount.Text = Val(txtPigGLAmount.Text) - Val(txtBreedAmount.Text) * (Val(txtBreedRate.Text) / 100)
End Sub

Private Sub txtPigLAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromBirth_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromBreed_HasChange()
   m_HasModify = True
   
   txtDayCount.Text = DateDiff("D", uctlFromBreed.ShowDate, uctlToBreed.ShowDate) + 1
   uctlFromBirth.ShowDate = DateAdd("D", 115, uctlFromBreed.ShowDate)
End Sub

Private Sub uctlPigGLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPigLLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlToBirth_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToBreed_HasChange()
   m_HasModify = True
   
   txtDayCount.Text = DateDiff("D", uctlFromBreed.ShowDate, uctlToBreed.ShowDate) + 1
   uctlToBirth.ShowDate = DateAdd("D", 115, uctlToBreed.ShowDate)
End Sub
