VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmConfigDoc 
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3480
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6138
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDocumentType 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ComboBox cboMonthType 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox cboYearType 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
      End
      Begin prjFarmManagement.uctlTextBox txtCode1 
         Height          =   405
         Left            =   1440
         TabIndex        =   3
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtCode2 
         Height          =   405
         Left            =   3840
         TabIndex        =   5
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDigitAmount 
         Height          =   405
         Left            =   7080
         TabIndex        =   8
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtPreFix 
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtRunningNo 
         Height          =   405
         Left            =   7800
         TabIndex        =   9
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtLastNo 
         Height          =   405
         Left            =   3840
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtCode3 
         Height          =   405
         Left            =   6360
         TabIndex        =   7
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin Threed.SSCheck ChkAutoYear 
         Height          =   375
         Left            =   8280
         TabIndex        =   25
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck ChkAutoMonth 
         Height          =   375
         Left            =   6600
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblCode3 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   6360
         TabIndex        =   23
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblLastNo 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   840
         Width           =   2505
      End
      Begin VB.Label lblRunningNo 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   7800
         TabIndex        =   21
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblMonthType 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lblYearType 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label lblPreFix 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDigitAmount 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   1680
         Width           =   585
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4290
         TabIndex        =   12
         Top             =   2580
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2640
         TabIndex        =   11
         Top             =   2580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentType 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblCode1 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblCode2 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   1680
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmConfigDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Cd As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public AllDocType As Collection
Private Sub cboDocumentType_Click()
Dim ID As Long
Dim Cd As CConfigDoc
   
   ID = cboDocumentType.ItemData(Minus2Zero(cboDocumentType.ListIndex))
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         txtLastNo.Text = Cd.GetFieldValue("LAST_NO")
         txtPreFix.Text = Cd.GetFieldValue("PREFIX")
         txtCode1.Text = Cd.GetFieldValue("CODE1")
         cboYearType.ListIndex = IDToListIndex(cboYearType, Cd.GetFieldValue("YEAR_TYPE"))
         txtCode2.Text = Cd.GetFieldValue("CODE2")
         cboMonthType.ListIndex = IDToListIndex(cboMonthType, Cd.GetFieldValue("MONTH_TYPE"))
         txtCode3.Text = Cd.GetFieldValue("CODE3")
         txtDigitAmount.Text = Cd.GetFieldValue("DIGIT_AMOUNT")
         txtRunningNo.Text = Cd.GetFieldValue("RUNNING_NO")
         
         ChkAutoMonth.Value = FlagToCheck(Cd.GetFieldValue("UPDATE_MONTH_FLAG"))
         ChkAutoYear.Value = FlagToCheck(Cd.GetFieldValue("UPDATE_YEAR_FLAG"))
      Else
         txtLastNo.Text = ""
         txtPreFix.Text = ""
         txtCode1.Text = ""
         cboYearType.ListIndex = -1
         txtCode2.Text = ""
         cboMonthType.ListIndex = -1
         txtCode3.Text = ""
         txtDigitAmount.Text = ""
         txtRunningNo.Text = ""
      End If
   End If

   m_HasModify = True
End Sub

Private Sub cboDocumentType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboMonthType_Click()
   m_HasModify = True
End Sub

Private Sub cboMonthType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboYearType_Click()
   m_HasModify = True
End Sub

Private Sub cboYearType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadConfigDoc(Nothing, m_Cd)
      Call GenerateAllConfigDoc
      Call LoadDocType
      Call InitYearType(cboYearType)
      Call InitMonthType(cboMonthType)
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub
Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Cd = New Collection
   Set AllDocType = New Collection
   
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   Call InitNormalLabel(lblDocumentType, MapText("�������͡���"))
   Call InitNormalLabel(lblLastNo, MapText("�����Ţ�ش����"))
   Call InitNormalLabel(lblPreFix, MapText("Prefix"))
   Call InitNormalLabel(lblCode1, MapText("-"))
   Call InitNormalLabel(lblYearType, MapText("��������"))
   Call InitNormalLabel(lblCode2, MapText("-"))
   Call InitNormalLabel(lblMonthType, MapText("��������͹"))
   Call InitNormalLabel(lblCode3, MapText("-"))
   Call InitNormalLabel(lblDigitAmount, MapText("��ѡ"))
   Call InitNormalLabel(lblRunningNo, MapText("RunNo"))
   
   Call txtDigitAmount.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtRunningNo.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   txtLastNo.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   Call InitCombo(cboDocumentType)
   Call InitCombo(cboYearType)
   Call InitCombo(cboMonthType)
   
   Call InitCheckBox(ChkAutoMonth, "��૵�ء��͹")
   Call InitCheckBox(ChkAutoYear, "��૵�ء��")
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))

End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

   OKClick = False
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim I As Long
Dim ID As Long
Dim Cd As CConfigDoc
   
   If Not VerifyCombo(lblDocumentType, cboDocumentType, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblDigitAmount, txtDigitAmount, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblRunningNo, txtRunningNo, True) Then
      Exit Function
   End If
   
   ID = cboDocumentType.ItemData(Minus2Zero(cboDocumentType.ListIndex))
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Cd Is Nothing Then
         Set Cd = New CConfigDoc
         Cd.Flag = "A"
         Call Cd.SetFieldValue("CONFIG_DOC_TYPE", ID)
      Else
         Cd.Flag = "E"
      End If
   End If
      
      
   If Cd.Flag = "A" Then
      Cd.ShowMode = SHOW_ADD
   ElseIf Cd.Flag = "E" Then
      Cd.ShowMode = SHOW_EDIT
   End If
   Call Cd.SetFieldValue("ENTERPRISE_ID", 99999)
   Call Cd.SetFieldValue("PREFIX", txtPreFix.Text)
   Call Cd.SetFieldValue("CODE1", txtCode1.Text)
   Call Cd.SetFieldValue("YEAR_TYPE", cboYearType.ItemData(Minus2Zero(cboYearType.ListIndex)))
   Call Cd.SetFieldValue("CODE2", txtCode2.Text)
   Call Cd.SetFieldValue("MONTH_TYPE", cboMonthType.ItemData(Minus2Zero(cboMonthType.ListIndex)))
   Call Cd.SetFieldValue("CODE3", txtCode3.Text)
   Call Cd.SetFieldValue("DIGIT_AMOUNT", txtDigitAmount.Text)
   Call Cd.SetFieldValue("RUNNING_NO", txtRunningNo.Text)
   
   Call Cd.SetFieldValue("MM", Right(Format(Year(Now), "00") & Format(Month(Now), "00"), 4))
   
   Call Cd.SetFieldValue("UPDATE_MONTH_FLAG", Check2Flag(ChkAutoMonth.Value))
   Call Cd.SetFieldValue("UPDATE_YEAR_FLAG", Check2Flag(ChkAutoYear.Value))
   
   Call EnableForm(Me, False)

   Call Cd.AddEditData
  
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cmdOK_Click()
   If cmdOK.Enabled = False Then
      Exit Sub
   End If
   Call SaveData
   
   OKClick = True
   Unload Me
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
      'Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub
Private Sub GenerateAllConfigDoc()
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
Dim MenuMask As String
   
   MenuMask = "YYYYYY"
   '1
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("� QUATATION (���)")
'   D.KEY_ID = 1
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.add(D)
'   Set D = Nothing
   
   '2
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("� PO (���)")
'   D.KEY_ID = 2
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("� �觢ͧ")
   D.KEY_ID = 3
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.Add(D)
   Set D = Nothing
   
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("�����")
'   D.KEY_ID = 4
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
   
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("����� (�Ѻ����)")
'   D.KEY_ID = 4
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("����˹�� (���)")
'   D.KEY_ID = 5
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("Ŵ˹�� (���)")
'   D.KEY_ID = 6
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("Ŵ˹�� �Ѻ�׹�Թ��� (���)")
'   D.KEY_ID = 7
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("��ҧ��� (���)")
'   D.KEY_ID = 8
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
'
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("㺹����")
   D.KEY_ID = 50
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.Add(D)
   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("��ԡ")
'   D.KEY_ID = 51
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
'   '===
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("��͹")
'   D.KEY_ID = 52
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("㺻�Ѻ�ʹ")
'   D.KEY_ID = 53
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.Add(D)
'   Set D = Nothing
   
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("�����١���/�١˹��")
'   D.KEY_ID = 70
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.add(D)
'   Set D = Nothing
'
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("���ʫѾ���������/���˹��")
'   D.KEY_ID = 71
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.add(D)
'   Set D = Nothing
   
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("���ʾ�ѡ�ҹ")
'   D.KEY_ID = 72
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.add(D)
'   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("㺹ӽҡ��Ҥ��")
   D.KEY_ID = 81
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.Add(D)
   Set D = Nothing
   
   '====
   TempCount = AllDocType.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call AllDocType.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Cd = Nothing
   Set AllDocType = Nothing
   If m_Rs.State = adStateOpen Then
      Call m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub LoadDocType()
Dim Mu As CMenuItem
Dim I As Long
   I = 0
   cboDocumentType.Clear
   cboDocumentType.AddItem ("")
   
   For Each Mu In AllDocType
      I = I + 1
      cboDocumentType.AddItem (Mu.MENU_TEXT)
      cboDocumentType.ItemData(I) = Mu.KEY_ID
   Next
End Sub


Private Sub txtCode1_Change()
   m_HasModify = True
End Sub

Private Sub txtCode2_Change()
   m_HasModify = True
End Sub
Private Sub txtDigitAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtLastNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPreFix_Change()
   m_HasModify = True
End Sub

Private Sub txtRunningNo_Change()
   m_HasModify = True
End Sub
Private Sub InitYearType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("�� 2 ��ѡ")
   C.ItemData(1) = 1

   C.AddItem ("�� 4 ��ѡ")
   C.ItemData(2) = 2
   
   C.AddItem ("�� 2 ��ѡ")
   C.ItemData(3) = 3

   C.AddItem ("�� 4 ��ѡ")
   C.ItemData(4) = 4
End Sub
Private Sub InitMonthType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("��")
   C.ItemData(1) = 1
   
End Sub

