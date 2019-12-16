VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditRevenueCostItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmAddEditRevenueCostItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4185
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   7382
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPigType 
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1920
         Width           =   5295
      End
      Begin VB.ComboBox cboPigStatus 
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   5295
      End
      Begin VB.ComboBox cboRevenueType 
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
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1020
         Width           =   5295
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtRevenueCostItemAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   2880
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtRevenueCostItemSell 
         Height          =   435
         Left            =   1890
         TabIndex        =   12
         Top             =   2400
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1980
         Width           =   1485
      End
      Begin VB.Label lblRevenueCostItemSell 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   14
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4080
         TabIndex        =   13
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblPigStatus 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1500
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2205
         TabIndex        =   2
         Top             =   3540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditRevenueCostItem.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblRevenueType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   9
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4050
         TabIndex        =   8
         Top             =   2940
         Width           =   1575
      End
      Begin VB.Label lblRevenueCostItemAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   2940
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5505
         TabIndex        =   4
         Top             =   3540
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3855
         TabIndex        =   3
         Top             =   3540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditRevenueCostItem.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditRevenueCostItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object
Public TempCollection As Collection
Private Sub cboPigStatus_Click()
   m_HasModify = True
End Sub
Private Sub cboPigType_Click()
   m_HasModify = True
End Sub

Private Sub cboRevenueType_Change()
   m_HasModify = True
End Sub

Private Sub cboRevenueType_Click()
   m_HasModify = True
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
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      cboRevenueType.ListIndex = -1
      txtRevenueCostItemAmount.Text = ""
      
   End If
   
   cboRevenueType.SetFocus
   Call ParentForm.RefreshGrid
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

   If Flag Then
      Call EnableForm(Me, False)
      
      Dim Ji As CRevenueCostItem
      Set Ji = TempCollection.Item(ID)
      
      cboRevenueType.ListIndex = IDToListIndex(cboRevenueType, Ji.GetFieldValue("REVENUE_TYPE_ID"))
      txtRevenueCostItemAmount.Text = Ji.GetFieldValue("REVENUE_COST_ITEM_AMOUNT")
      cboPigStatus.ListIndex = IDToListIndex(cboPigStatus, Ji.GetFieldValue("PIG_STATUS"))
      txtRevenueCostItemSell.Text = Ji.GetFieldValue("REVENUE_COST_ITEM_SELL")
      cboPigType.ListIndex = IDToListIndex(cboPigType, Ji.GetFieldValue("PIG_TYPE"))
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim PaymentType As Long

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("LEDGER_CASH_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("LEDGER_CASH_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
   
   If Not VerifyCombo(lblRevenueType, cboRevenueType, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(USERNAME_UNIQUE, txtRevenueCostItemNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtRevenueCostItemNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CRevenueCostItem
   
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CRevenueCostItem
      
      EnpAddress.Flag = "A"
      
      Call TempCollection.Add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If

   'นำฝากเงินสดในมือ
   Call EnpAddress.SetFieldValue("REVENUE_COST_ITEM_AMOUNT", Val(txtRevenueCostItemAmount.Text))
   Call EnpAddress.SetFieldValue("REVENUE_TYPE_ID", cboRevenueType.ItemData(Minus2Zero(cboRevenueType.ListIndex)))
   Call EnpAddress.SetFieldValue("REVENUE_TYPE_NAME", cboRevenueType.Text)
   Call EnpAddress.SetFieldValue("REVENUE_COST_ITEM_SELL", Val(txtRevenueCostItemSell.Text))
   Call EnpAddress.SetFieldValue("PIG_STATUS", cboPigStatus.ItemData(Minus2Zero(cboPigStatus.ListIndex)))
   Call EnpAddress.SetFieldValue("PIG_STATUS_NAME", cboPigStatus.Text)
   Call EnpAddress.SetFieldValue("PIG_TYPE", cboPigType.ItemData(Minus2Zero(cboPigType.ListIndex)))
   Call EnpAddress.SetFieldValue("PIG_TYPE_NAME", cboPigType.Text)
   
   Set EnpAddress = Nothing

   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
                              
      Call LoadRevenueType(cboRevenueType)
      
      Call LoadProductStatus(cboPigStatus)
      
      Call LoadProductType(cboPigType)
      
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
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblRevenueCostItemAmount, MapText("จำนวนต้นทุน"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblRevenueType, MapText("ประเภทรายรับ"))
   Call InitNormalLabel(lblPigStatus, MapText("สถานะสุกร"))
   Call InitNormalLabel(lblRevenueCostItemSell, MapText("มูลค่ารายรับ"))
   Call InitNormalLabel(lblPigType, MapText("ประเภทสุกร"))
   
   Call txtRevenueCostItemAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboRevenueType)
   Call InitCombo(cboPigType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboPigStatus)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   
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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub txtRevenueCostItemAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtRevenueCostItemSell_Change()
   m_HasModify = True
End Sub
