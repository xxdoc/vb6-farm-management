VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditBloodItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditBloodItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   1085
      _Version        =   131073
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   675
      Left            =   0
      TabIndex        =   7
      Top             =   2460
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   1191
      _Version        =   131073
      Begin Threed.SSCommand cmdOK2 
         Height          =   615
         Left            =   3225
         TabIndex        =   4
         Top             =   30
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   5310
         TabIndex        =   6
         Top             =   30
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   615
         Index           =   0
         Left            =   11130
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   13230
         TabIndex        =   8
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1845
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   3254
      _Version        =   131073
      Begin VB.ComboBox cboBloodSpec 
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
         Left            =   1950
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   3885
      End
      Begin prjBoonmeeGraph.uctlTextBox txtZipcode 
         Height          =   435
         Left            =   12450
         TabIndex        =   5
         Top             =   3270
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   767
      End
      Begin prjBoonmeeGraph.uctlTextBox txtMaleStd 
         Height          =   435
         Left            =   1950
         TabIndex        =   1
         Top             =   720
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
      End
      Begin prjBoonmeeGraph.uctlTextBox txtFemaleStd 
         Height          =   435
         Left            =   5730
         TabIndex        =   2
         Top             =   720
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
      End
      Begin prjBoonmeeGraph.uctlTextBox txtResult 
         Height          =   435
         Left            =   1950
         TabIndex        =   3
         Top             =   1170
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
      End
      Begin VB.Label lblResult 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   16
         Top             =   1230
         Width           =   1815
      End
      Begin VB.Label lblFemaleStd 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label lblUnit1 
         Height          =   345
         Left            =   3870
         TabIndex        =   14
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label lblMaleStd 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   13
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label lblBloodSpec 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAddEditBloodItem"
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
Public ResourceTypeID As Long
Private m_BloodSpecs As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboDocumentType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub cboBloodSpec_Click()
Dim TempID As Long
Dim Bp As CBloodSpec

   m_HasModify = True
   TempID = cboBloodSpec.ItemData(Minus2Zero(cboBloodSpec.ListIndex))
   If TempID > 0 Then
      Set Bp = m_BloodSpecs.Item(Str(TempID))
      txtMaleStd.Text = Bp.MALE_STD
      txtFemaleStd.Text = Bp.FEMALE_STD
   End If
End Sub

Private Sub cboPeriodDesc_Change()
   m_HasModify = True
End Sub

Private Sub cboPeriodDesc_Click()
   m_HasModify = True
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblBloodSpec, "รายการตรวจ")
   Call InitNormalLabel(lblMaleStd, "ค่ามาตรฐานชาย")
   Call InitNormalLabel(lblFemaleStd, "ค่ามาตรฐานหญิง")
   Call InitNormalLabel(lblResult, "ผลการตรวจ")
   Call InitNormalLabel(lblUnit1, "")

   Call InitCombo(cboBloodSpec)

   Call txtMaleStd.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtMaleStd.Enabled = False
   Call txtFemaleStd.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFemaleStd.Enabled = False
   Call txtResult.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitMainButton(cmdOK2, "ตกลง (F2)")
   Call InitMainButton(cmdCancel, "ยกเลิก (ESC)")
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
            
      If ShowMode = SHOW_EDIT Then
         Dim D As CBSheetItem
         Set D = TempCollection.Item(ID)

         cboBloodSpec.ListIndex = IDToListIndex(cboBloodSpec, D.BLOOD_SPEC_ID)
         txtResult.Text = D.SPEC_VALUE
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK2_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
   
   SaveData = False
   If Not VerifyCombo(lblBloodSpec, cboBloodSpec) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblResult, txtResult) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim D As CBSheetItem
   If ShowMode = SHOW_ADD Then
      Set D = New CBSheetItem
      D.Flag = "A"

      Call TempCollection.Add(D)
   Else
      Set D = TempCollection.Item(ID)
      D.Flag = "E"
   End If
   
   D.SPEC_NAME = cboBloodSpec.Text
   D.BLOOD_SPEC_ID = cboBloodSpec.ItemData(Minus2Zero(cboBloodSpec.ListIndex))
   D.SPEC_VALUE = Val(txtResult.Text)
   D.MALE_STD = txtMaleStd.Text
   D.FEMALE_STD = txtFemaleStd.Text
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadBloodSpec(cboBloodSpec, m_BloodSpecs)
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
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
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK2_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   
   Set m_BloodSpecs = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_BloodSpecs = Nothing
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

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtFemaleStd_Change()
   m_HasModify = True
End Sub

Private Sub txtDrugName_Change()
   m_HasModify = True
End Sub

Private Sub txtMoo_Change()
   m_HasModify = True
End Sub

Private Sub txtIssuePlace_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtVillage_Change()
   m_HasModify = True
End Sub

Private Sub SSOption2_Click(Value As Integer)

End Sub

Private Sub radResourceType1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub radResourceType2_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub Label1_Click()

End Sub

Private Sub txtResult_Change()
   m_HasModify = True
End Sub

Private Sub txtMaleStd_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlExpireDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlIssueDate_HasChange()
   m_HasModify = True
End Sub
