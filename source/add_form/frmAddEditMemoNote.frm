VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMemoNote 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAddEditMemoNote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8565
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   15108
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.TextBox txtResolution 
         Height          =   1575
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   6000
         Width           =   9975
      End
      Begin VB.TextBox txtDescription 
         Height          =   1695
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   4080
         Width           =   9975
      End
      Begin VB.ComboBox cboMemoType 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   2595
      End
      Begin VB.ComboBox cboMemoStatus 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2760
         Width           =   2595
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtSubject 
         Height          =   435
         Left            =   1560
         TabIndex        =   8
         Top             =   3390
         Width           =   9915
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlFromDateCreate 
         Height          =   405
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlFromDateFinish 
         Height          =   405
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlFromDateFinishReal 
         Height          =   405
         Left            =   1560
         TabIndex        =   3
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlCreateBy 
         Height          =   435
         Left            =   6165
         TabIndex        =   5
         Top             =   2280
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlCreateTo 
         Height          =   435
         Left            =   6165
         TabIndex        =   7
         Top             =   2760
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblResolution 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   24
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   23
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label lblMemoType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   22
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label lblCreateDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblFinishDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   20
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label lblFinishReal 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   19
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lblMemoStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label lblCreateBy 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Label lblCreateTo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   2790
         Width           =   1485
      End
      Begin Threed.SSCheck chkWarn 
         Height          =   435
         Left            =   6120
         TabIndex        =   1
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   15
         Top             =   3480
         Width           =   1335
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5715
         TabIndex        =   12
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4065
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMemoNote.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMemoNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_MemoNote As CMemoNote

Private m_Employees As Collection
Private m_Employee1s As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cboMemoStatus_Click()
   m_HasModify = True
End Sub

Private Sub cboMemoStatus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboMemoType_Click()
   m_HasModify = True
End Sub

Private Sub cboMemoType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkWarn_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkWarn_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_MemoNote.SetFieldValue("MEMO_NOTE_ID", ID)
      m_MemoNote.QueryFlag = 1
      If Not glbDaily.QueryMemoNote(m_MemoNote, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_MemoNote.PopulateFromRS(1, m_Rs)
      
      uctlFromDateCreate.ShowDate = m_MemoNote.GetFieldValue("MEMO_NOTE_DATE_CREATE")
      uctlFromDateFinish.ShowDate = m_MemoNote.GetFieldValue("MEMO_NOTE_DATE_FINISH")
      uctlFromDateFinishReal.ShowDate = m_MemoNote.GetFieldValue("MEMO_NOTE_DATE_FINISH_REAL")
      chkWarn.Value = FlagToCheck(m_MemoNote.GetFieldValue("MEMO_NOTE_WARN"))
      
      cboMemoType.ListIndex = IDToListIndex(cboMemoType, m_MemoNote.GetFieldValue("MEMO_NOTE_TYPE"))
      cboMemoStatus.ListIndex = IDToListIndex(cboMemoStatus, m_MemoNote.GetFieldValue("MEMO_NOTE_STATUS"))
      uctlCreateBy.MyCombo.ListIndex = IDToListIndex(uctlCreateBy.MyCombo, m_MemoNote.GetFieldValue("MEMO_NOTE_CREATE_BY"))
      uctlCreateTo.MyCombo.ListIndex = IDToListIndex(uctlCreateTo.MyCombo, m_MemoNote.GetFieldValue("MEMO_NOTE_CREATE_TO"))
      
      txtSubject.Text = m_MemoNote.GetFieldValue("MEMO_NOTE_SUBJECT")
      txtDescription.Text = m_MemoNote.GetFieldValue("MEMO_NOTE_DESCRIPTION")
      txtResolution.Text = m_MemoNote.GetFieldValue("MEMO_NOTE_RESOLUTION")
      
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
   
   If Not VerifyDate(lblCreateDate, uctlFromDateCreate, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSubject, txtSubject, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_MemoNote.ShowMode = ShowMode
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_ID", ID)
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_DATE_CREATE", uctlFromDateCreate.ShowDate)
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_DATE_FINISH", uctlFromDateFinish.ShowDate)
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_DATE_FINISH_REAL", uctlFromDateFinishReal.ShowDate)
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_TYPE", cboMemoType.ItemData(Minus2Zero(cboMemoType.ListIndex)))
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_STATUS", cboMemoStatus.ItemData(Minus2Zero(cboMemoStatus.ListIndex)))
   
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_CREATE_BY", uctlCreateBy.MyCombo.ItemData(Minus2Zero(uctlCreateBy.MyCombo.ListIndex)))
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_CREATE_TO", uctlCreateTo.MyCombo.ItemData(Minus2Zero(uctlCreateTo.MyCombo.ListIndex)))
   
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_SUBJECT", txtSubject.Text)
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_DESCRIPTION", txtDescription.Text)
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_RESOLUTION", txtResolution.Text)
   
   Call m_MemoNote.SetFieldValue("MEMO_NOTE_WARN", Check2Flag(chkWarn.Value))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditMemoNote(m_MemoNote, IsOK, True, glbErrorLog) Then
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
      
      Call LoadMaster(cboMemoType, , MEMO_TYPE)
      Call LoadMaster(cboMemoStatus, , MEMO_STATUS)
            
      Call LoadUserAccount(uctlCreateBy.MyCombo, m_Employees)
      Set uctlCreateBy.MyCollection = m_Employees
      
      Call LoadUserAccount(uctlCreateTo.MyCombo, m_Employee1s)
      Set uctlCreateTo.MyCollection = m_Employee1s
            
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         uctlFromDateCreate.ShowDate = Now
         chkWarn.Value = ssCBChecked
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
   ElseIf Shift = 0 And KeyCode = 123 Then
      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblCreateDate, MapText("วันที่สร้าง"))
   Call InitNormalLabel(lblFinishDate, MapText("วันกำหนดเสร็จ"))
   Call InitNormalLabel(lblFinishReal, MapText("วันที่เสร็จ"))
   Call InitNormalLabel(lblMemoType, MapText("ประเภท"))
   Call InitNormalLabel(lblMemoStatus, MapText("สถานะ"))
   Call InitNormalLabel(lblCreateBy, MapText("สร้างโดย"))
   Call InitNormalLabel(lblCreateTo, MapText("มอบหมายให้"))
   
   Call InitNormalLabel(lblSubject, MapText("หัวข้อ"))
   Call InitNormalLabel(lblDescription, MapText("รายละเอียด"))
   Call InitNormalLabel(lblResolution, MapText("วิธีแก้ปัญหา"))
   
   Call InitCheckBox(chkWarn, "เตือน")
   
   Call txtSubject.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitTextBox(txtDescription, "")
   Call InitTextBox(txtResolution, "")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboMemoType)
   Call InitCombo(cboMemoStatus)
   
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
   
   Set m_MemoNote = New CMemoNote
   Set m_Employees = New Collection
   Set m_Employee1s = New Collection
   
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_MemoNote = Nothing
   Set m_Employees = Nothing
   Set m_Employee1s = Nothing
   If m_Rs.State = adStateOpen Then
      Call m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub txtDescription_Change()
   m_HasModify = True
End Sub

Private Sub txtResolution_Change()
   m_HasModify = True
End Sub

Private Sub txtSubject_Change()
   m_HasModify = True
End Sub

Private Sub uctlCreateBy_Change()
   m_HasModify = True
End Sub

Private Sub uctlCreateTo_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromDateCreate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromDateFinish_HasChange()
   m_HasModify = True
End Sub
