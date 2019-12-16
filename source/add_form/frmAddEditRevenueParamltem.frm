VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditRevenueParamltem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditRevenueParamltem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4155
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   7329
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlRevenueTypeLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   2
         Top             =   1230
         Width           =   5355
         _extentx        =   9446
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFromBreed 
         Height          =   405
         Left            =   1710
         TabIndex        =   0
         Top             =   360
         Width           =   2535
         _extentx        =   6800
         _extenty        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtBreedAmount 
         Height          =   435
         Left            =   1710
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
         _extentx        =   9763
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlDate uctlToBreed 
         Height          =   405
         Left            =   1710
         TabIndex        =   1
         Top             =   780
         Width           =   2535
         _extentx        =   6800
         _extenty        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtChildRate 
         Height          =   435
         Left            =   1710
         TabIndex        =   4
         Top             =   2130
         Width           =   1935
         _extentx        =   9763
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgWeight 
         Height          =   435
         Left            =   1710
         TabIndex        =   5
         Top             =   2580
         Width           =   1935
         _extentx        =   9763
         _extenty        =   767
      End
      Begin VB.Label lblRevenueType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblAvgWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   1485
      End
      Begin VB.Label lblChildRate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2190
         Width           =   1485
      End
      Begin VB.Label lblFromBreed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblToBreed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1605
         TabIndex        =   6
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3255
         TabIndex        =   7
         Top             =   3300
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
         Left            =   4905
         TabIndex        =   8
         Top             =   3300
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
         TabIndex        =   11
         Top             =   1740
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditRevenueParamltem"
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

Private m_RevenueTypes As Collection
Private m_PartItems As Collection

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
   
   Call InitNormalLabel(lblFromBreed, MapText("จากวันขาย"))
   Call InitNormalLabel(lblToBreed, MapText("ถึงวันที่ขาย"))
   Call InitNormalLabel(lblBreedAmount, MapText("จำนวนที่ขาย"))
   Call InitNormalLabel(lblChildRate, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblAvgWeight, MapText("ราคารวม"))
   Call InitNormalLabel(lblRevenueType, MapText("สินค้า"))
   
   Call txtBreedAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtChildRate.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtAvgWeight.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtAvgWeight.Enabled = False
   
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
         Dim Ji As CRvnPrmItem

         Set Ji = TempCollection.Item(ID)
         uctlFromBreed.ShowDate = Ji.GetFieldValue("FROM_SALE")
         uctlToBreed.ShowDate = Ji.GetFieldValue("TO_SALE")
         txtBreedAmount.Text = Ji.GetFieldValue("SALE_AMOUNT")
         txtChildRate.Text = Ji.GetFieldValue("UNIT_PRICE")
         txtAvgWeight.Text = Ji.GetFieldValue("TOTAL_PRICE")
         uctlRevenueTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlRevenueTypeLookup.MyCombo, Ji.GetFieldValue("REVENUE_ID"))
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
      txtChildRate.Text = ""
      txtAvgWeight.Text = ""
      uctlRevenueTypeLookup.MyCombo.ListIndex = -1
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
   If Not VerifyCombo(lblRevenueType, uctlRevenueTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblBreedAmount, txtBreedAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblChildRate, txtChildRate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ji As CRvnPrmItem
   If ShowMode = SHOW_ADD Then
      Set Ji = New CRvnPrmItem
      Ji.Flag = "A"
      Call TempCollection.Add(Ji)
   Else
      Set Ji = TempCollection.Item(ID)
      If Ji.Flag <> "A" Then
         Ji.Flag = "E"
      End If
   End If

   Call Ji.SetFieldValue("FROM_SALE", uctlFromBreed.ShowDate)
   Call Ji.SetFieldValue("TO_SALE", uctlToBreed.ShowDate)
   Call Ji.SetFieldValue("SALE_AMOUNT", txtBreedAmount.Text)
   Call Ji.SetFieldValue("UNIT_PRICE", txtChildRate.Text)
   Call Ji.SetFieldValue("TOTAL_PRICE", Val(txtAvgWeight.Text))
   Call Ji.SetFieldValue("REVENUE_ID", uctlRevenueTypeLookup.MyCombo.ItemData(Minus2Zero(uctlRevenueTypeLookup.MyCombo.ListIndex)))
   Call Ji.SetFieldValue("REVENUE_NO", uctlRevenueTypeLookup.MyTextBox.Text)
   Call Ji.SetFieldValue("REVENUE_NAME", uctlRevenueTypeLookup.MyCombo.Text)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadRevenueType(uctlRevenueTypeLookup.MyCombo, m_RevenueTypes)
      Set uctlRevenueTypeLookup.MyCollection = m_RevenueTypes
      
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
   Set m_RevenueTypes = New Collection
   Set m_PartItems = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_RevenueTypes = Nothing
   Set m_PartItems = Nothing
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

Private Sub txtBirthRate_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromBirth_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToBirth_HasChange()
   m_HasModify = True
End Sub

Private Sub txtBreedAmount_Change()
   m_HasModify = True
   txtAvgWeight.Text = Val(txtBreedAmount.Text) * Val(txtChildRate.Text)
End Sub

Private Sub txtChildRate_Change()
   m_HasModify = True
   txtAvgWeight.Text = Val(txtBreedAmount.Text) * Val(txtChildRate.Text)
End Sub

Private Sub uctlFromBreed_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlRevenueTypeLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlToBreed_HasChange()
   m_HasModify = True
End Sub
