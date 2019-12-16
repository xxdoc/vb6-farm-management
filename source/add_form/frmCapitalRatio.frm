VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmCapitalRatio 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmCapitalRatio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2805
      Left            =   -30
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   4948
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboHouse 
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   990
         Width           =   4395
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtRatioAmount 
         Height          =   435
         Left            =   1830
         TabIndex        =   1
         Top             =   1440
         Width           =   1875
         _extentx        =   2355
         _extenty        =   767
      End
      Begin VB.Label lblRatioAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   8
         Top             =   1500
         Width           =   1665
      End
      Begin VB.Label lblBath 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3750
         TabIndex        =   7
         Top             =   1530
         Width           =   675
      End
      Begin VB.Label lblUsername 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   1020
         Width           =   1665
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3495
         TabIndex        =   3
         Top             =   2010
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1845
         TabIndex        =   2
         Top             =   2010
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCapitalRatio.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCapitalRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OKClick As Boolean
Public TempCollection As Collection
Public ID As Long

Public m_HasModify As Boolean
Public m_HasActivate As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public TotalAmount As Double
Private m_GuiIndex As Long
Public ExpensePartMode As Long

Private Sub cmdPasswd_Click()

End Sub

Private Sub cboHouse_Click()
   m_HasModify = True
End Sub

Private Sub cboHouse_KeyPress(KeyAscii As Integer)
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

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If

   If Not VerifyCombo(lblUsername, cboHouse, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblRatioAmount, txtRatioAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CMovementItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CMovementItem
      EnpAddress.Flag = "A"
      Call TempCollection.Add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If
   
   If ExpensePartMode = 1 Then
      EnpAddress.EXPENSE_TYPE = cboHouse.ItemData(Minus2Zero(cboHouse.ListIndex))
      EnpAddress.PART_GROUP_ID = -1
      EnpAddress.PART_ITEM_ID = -1
      EnpAddress.EXPENSE_TYPE_NAME = cboHouse.Text
   Else
      EnpAddress.EXPENSE_TYPE = -1
      EnpAddress.PART_GROUP_ID = cboHouse.ItemData(Minus2Zero(cboHouse.ListIndex))
      EnpAddress.PART_ITEM_ID = glbDaily.LookupPartItemFromPartGroupID(EnpAddress.PART_GROUP_ID)
      EnpAddress.PART_GROUP_NAME = cboHouse.Text
   End If
   EnpAddress.CAPITAL_AMOUNT = txtRatioAmount.Text
   
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
Dim Er As CMovementItem

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call EnableForm(Me, False)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Set Er = TempCollection(ID)
         If Er.EXPENSE_TYPE > 0 Then
            ExpensePartMode = 1
         Else
            ExpensePartMode = 2
         End If
      End If
      
      If ExpensePartMode = 1 Then
         Call LoadExpenseType(cboHouse, Nothing)
      Else
         Call LoadPartGroup(cboHouse, Nothing)
      End If
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Set Er = TempCollection(ID)
         If ExpensePartMode = 1 Then
            cboHouse.ListIndex = IDToListIndex(cboHouse, Er.EXPENSE_TYPE)
         Else
            cboHouse.ListIndex = IDToListIndex(cboHouse, Er.PART_GROUP_ID)
         End If
         txtRatioAmount.Text = Er.CAPITAL_AMOUNT
      End If
      
      Call EnableForm(Me, True)
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
   pnlHeader.Caption = Me.Caption
   
   If ExpensePartMode = 1 Then
      Call InitNormalLabel(lblUsername, "ประเภทรายจ่าย ")
   Else
      Call InitNormalLabel(lblUsername, "กลุ่มวัตถุดิบ ")
   End If
   Call InitNormalLabel(lblRatioAmount, "มูลค่า")
   Call InitNormalLabel(lblBath, "บาท")
   Call InitCombo(cboHouse)
   
   Call txtRatioAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_HasModify = False
   m_HasActivate = False
   Call InitFormLayout
End Sub

Private Sub txtUsername_Change()
   m_HasModify = True
End Sub

Private Sub txtOldPassword_GotFocus()
   m_GuiIndex = 1
End Sub

Private Sub txtRatioAmount_GotFocus()
   m_GuiIndex = 2
End Sub
