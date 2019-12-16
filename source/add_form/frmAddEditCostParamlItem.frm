VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCostParamItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
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
   Icon            =   "frmAddEditCostParamlItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
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
      Height          =   2715
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   4789
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlAccountLookup 
         Height          =   405
         Left            =   1710
         TabIndex        =   1
         Top             =   750
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtAmount 
         Height          =   435
         Left            =   1710
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlParttypeLookup 
         Height          =   405
         Left            =   1710
         TabIndex        =   0
         Top             =   300
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblAccount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2033
         TabIndex        =   3
         Top             =   1860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3683
         TabIndex        =   4
         Top             =   1860
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
         Left            =   5333
         TabIndex        =   5
         Top             =   1860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1260
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditCostParamItem"
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
   
   Call InitNormalLabel(lblAmount, MapText("ราคา"))
   Call InitNormalLabel(lblAccount, MapText("อาหาร/ยา"))
   Call InitNormalLabel(lblPartType, MapText("ประเภท"))
   
   Call txtAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
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
         Dim Ji As CCostPrmItem

         Set Ji = TempCollection.Item(ID)

         txtAmount.Text = Ji.GetFieldValue("COST_RATE")
         uctlParttypeLookup.MyCombo.ListIndex = IDToListIndex(uctlParttypeLookup.MyCombo, Ji.GetFieldValue("PART_TYPE"))
         uctlAccountLookup.MyCombo.ListIndex = IDToListIndex(uctlAccountLookup.MyCombo, Ji.GetFieldValue("FOOD_ID"))
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
      uctlAccountLookup.MyCombo.ListIndex = -1
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

   If Not VerifyCombo(lblAccount, uctlAccountLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAmount, txtAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ji As CCostPrmItem
   If ShowMode = SHOW_ADD Then
      Set Ji = New CCostPrmItem
      Ji.Flag = "A"
      Call TempCollection.Add(Ji)
   Else
      Set Ji = TempCollection.Item(ID)
      If Ji.Flag <> "A" Then
         Ji.Flag = "E"
      End If
   End If

   Call Ji.SetFieldValue("COST_RATE", txtAmount.Text)
   Call Ji.SetFieldValue("FOOD_ID", uctlAccountLookup.MyCombo.ItemData(Minus2Zero(uctlAccountLookup.MyCombo.ListIndex)))
   Call Ji.SetFieldValue("FOOD_NO", uctlAccountLookup.MyTextBox.Text)
   Call Ji.SetFieldValue("FOOD_NAME", uctlAccountLookup.MyCombo.Text)
   Call Ji.SetFieldValue("PART_TYPE", uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex)))
   Call Ji.SetFieldValue("PART_TYPE_NAME", uctlParttypeLookup.MyCombo.Text)

   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPartType(uctlParttypeLookup.MyCombo, m_PartTypes)
      Set uctlParttypeLookup.MyCollection = m_PartTypes
      
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
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
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

Private Sub uctlParttypeLookup_Change()
Dim PartTypeID As Long

   PartTypeID = uctlParttypeLookup.MyCombo.ItemData(Minus2Zero(uctlParttypeLookup.MyCombo.ListIndex))
   If PartTypeID > 0 Then
      Call LoadPartItem(uctlAccountLookup.MyCombo, m_PartItems, PartTypeID, "N")
      Set uctlAccountLookup.MyCollection = m_PartItems
   End If
   m_HasModify = True
End Sub
