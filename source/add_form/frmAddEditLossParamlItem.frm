VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditLossParamItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3405
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
   Icon            =   "frmAddEditLossParamlItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2835
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   5001
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboLossType 
         Height          =   510
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3855
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1980
         TabIndex        =   2
         Top             =   1170
         Width           =   1935
         _extentx        =   9763
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlParttypeLookup 
         Height          =   405
         Left            =   1980
         TabIndex        =   0
         Top             =   300
         Width           =   5385
         _extentx        =   9499
         _extenty        =   714
      End
      Begin VB.Label lblLossType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   10
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   390
         TabIndex        =   9
         Top             =   330
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2310
         TabIndex        =   3
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3960
         TabIndex        =   4
         Top             =   1980
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
         Left            =   5610
         TabIndex        =   5
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   390
         TabIndex        =   8
         Top             =   1230
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditLossParamItem"
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

Private m_ProductStatus As Collection
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

Private Sub cboPartType_Click()
   m_HasModify = True
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
   
   Call InitNormalLabel(lblAmount, MapText("% สูญเสีย/วัน"))
   Call InitNormalLabel(lblPartType, MapText("สถานะ"))
   Call InitNormalLabel(lblLossType, MapText("จำนวนสูญเสีย"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboLossType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Ji As CTrnPrmItem

         Set Ji = TempCollection.Item(ID)

         txtAmount.Text = Ji.GetFieldValue("LOSS_RATE")
         uctlParttypeLookup.MyCombo.ListIndex = IDToListIndex(uctlParttypeLookup.MyCombo, Ji.GetFieldValue("PIG_STATUS"))
         cboLossType.ListIndex = IDToListIndex(cboLossType, Ji.GetFieldValue("LOSS_TYPE"))
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
      uctlParttypeLookup.MyCombo.ListIndex = -1
      txtAmount.Text = ""
      Call uctlParttypeLookup.SetFocus
   End If
   Call QueryData(True)
   Call ParentForm.RefreshGrid(True)
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

   If Not VerifyCombo(lblPartType, uctlParttypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLossType, cboLossType, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ji As CTrnPrmItem
   If ShowMode = SHOW_ADD Then
      Set Ji = New CTrnPrmItem
      Ji.Flag = "A"
      Call TempCollection.Add(Ji)
   Else
      Set Ji = TempCollection.Item(ID)
      If Ji.Flag <> "A" Then
         Ji.Flag = "E"
      End If
   End If

   Call Ji.SetFieldValue("LOSS_RATE", txtAmount.Text)
   Call Ji.SetFieldValue("PIG_STATUS", uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex)))
   Call Ji.SetFieldValue("STATUS_NAME", uctlParttypeLookup.MyCombo.Text)
   Call Ji.SetFieldValue("STATUS_NO", uctlParttypeLookup.MyTextBox.Text)
Call Ji.SetFieldValue("LOSS_TYPE", cboLossType.ItemData(Minus2Zero(cboLossType.ListIndex)))

   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadProductStatus(uctlParttypeLookup.MyCombo, m_ProductStatus)
      Set uctlParttypeLookup.MyCollection = m_ProductStatus
      
      Call InitLossType(cboLossType)
      
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
   Set m_ProductStatus = New Collection
   Set m_PartItems = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_ProductStatus = Nothing
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
