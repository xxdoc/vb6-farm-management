VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmMangementConfig 
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmMangementConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4770
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8414
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtTarget 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1800
         Width           =   1365
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   20
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtDiff 
         Height          =   435
         Left            =   4650
         TabIndex        =   3
         Top             =   1800
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvg 
         Height          =   435
         Left            =   6180
         TabIndex        =   4
         Top             =   1800
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtActualBirth 
         Height          =   435
         Left            =   3240
         TabIndex        =   2
         Top             =   1800
         Width           =   1395
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlBirthDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtMonth1 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2340
         Width           =   1365
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMonth3 
         Height          =   435
         Left            =   4650
         TabIndex        =   7
         Top             =   2340
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMonth4 
         Height          =   435
         Left            =   6180
         TabIndex        =   8
         Top             =   2340
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMonth2 
         Height          =   435
         Left            =   3240
         TabIndex        =   6
         Top             =   2340
         Width           =   1395
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft1 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   2790
         Width           =   1365
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft3 
         Height          =   435
         Left            =   4650
         TabIndex        =   11
         Top             =   2790
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft4 
         Height          =   435
         Left            =   6180
         TabIndex        =   12
         Top             =   2790
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft2 
         Height          =   435
         Left            =   3240
         TabIndex        =   10
         Top             =   2790
         Width           =   1395
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMix1 
         Height          =   435
         Left            =   1860
         TabIndex        =   13
         Top             =   3240
         Width           =   1365
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMix3 
         Height          =   435
         Left            =   4650
         TabIndex        =   15
         Top             =   3240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMix4 
         Height          =   435
         Left            =   6180
         TabIndex        =   16
         Top             =   3240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMix2 
         Height          =   435
         Left            =   3240
         TabIndex        =   14
         Top             =   3240
         Width           =   1395
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin VB.Label lblMix 
         Alignment       =   1  'Right Justify
         Caption         =   "ผสมไว้"
         Height          =   315
         Left            =   180
         TabIndex        =   28
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "คงเหลือ"
         Height          =   315
         Left            =   180
         TabIndex        =   27
         Top             =   2910
         Width           =   1575
      End
      Begin VB.Label lblMonth 
         Alignment       =   1  'Right Justify
         Caption         =   "เดือน"
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblBirthDate 
         Alignment       =   1  'Right Justify
         Caption         =   "กำหนดคลอด"
         Height          =   315
         Left            =   180
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4275
         TabIndex        =   18
         Top             =   3960
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2625
         TabIndex        =   17
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmMangementConfig.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblTarget 
         Alignment       =   2  'Center
         Caption         =   "เป้าหมาย"
         Height          =   315
         Left            =   1920
         TabIndex        =   24
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label lblActualBirth 
         Alignment       =   2  'Center
         Caption         =   "เกิดจริง"
         Height          =   315
         Left            =   3330
         TabIndex        =   23
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label lblDiff 
         Alignment       =   2  'Center
         Caption         =   "ผลต่าง"
         Height          =   315
         Left            =   4770
         TabIndex        =   22
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label lblAvg 
         Alignment       =   2  'Center
         Caption         =   "เฉลี่ย"
         Height          =   315
         Left            =   6330
         TabIndex        =   21
         Top             =   1500
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMangementConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_ManagementConfig As CManagementConfig
Private m_Houses As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ReportKey As String
Public ReportParams As Collection

Private FileName As String
Private m_SumUnit As Double
Private m_OldPartItemID As Long
Private m_PigStatus As Collection

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_ManagementConfig.MANAGEMENT_CONFIG_ID = ID
      If Not glbDaily.QueryManagementConfig(m_ManagementConfig, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_ManagementConfig.PopulateFromRS(1, m_Rs)
      txtTarget.Text = m_ManagementConfig.TARGET
      txtActualBirth.Text = m_ManagementConfig.ACTUAL_BIRTH
      txtDiff.Text = m_ManagementConfig.Diff
      txtAvg.Text = m_ManagementConfig.AVERAGE
      uctlBirthDate.ShowDate = m_ManagementConfig.BIRTH_DATE
      txtMonth1.Text = m_ManagementConfig.MONTH1
      txtLeft1.Text = m_ManagementConfig.LEFT1
      txtMix1.Text = m_ManagementConfig.MIX1
      txtMonth2.Text = m_ManagementConfig.MONTH2
      txtLeft2.Text = m_ManagementConfig.LEFT2
      txtMix2.Text = m_ManagementConfig.MIX2
      txtMonth3.Text = m_ManagementConfig.MONTH3
      txtLeft3.Text = m_ManagementConfig.LEFT3
      txtMix3.Text = m_ManagementConfig.MIX3
      txtMonth4.Text = m_ManagementConfig.MONTH4
      txtLeft4.Text = m_ManagementConfig.LEFT4
      txtMix4.Text = m_ManagementConfig.MIX4
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Sub AddParam(ParamName As String, ParamValue As Variant)
Dim Yg As CYGroup

   Set Yg = New CYGroup
   Yg.Value = ParamValue
   Yg.Key = ParamName
   Call ReportParams.Add(Yg, ParamName)
   Set Yg = Nothing
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Pi As CPartItem
Dim Yg As CYGroup

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

   If Not VerifyTextControl(lblTarget, txtTarget, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblActualBirth, txtActualBirth, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDiff, txtDiff, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAvg, txtAvg, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDiff, uctlBirthDate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMonth, txtMonth1, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMonth, txtMonth2, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMonth, txtMonth3, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMonth, txtMonth4, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblLeft, txtLeft1, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblLeft, txtLeft2, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblLeft, txtLeft3, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblLeft, txtLeft4, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMix, txtMix1, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMix, txtMix2, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMix, txtMix3, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMix, txtMix4, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
   Call EnableForm(Me, False)
   
'   Set ReportParams = Nothing
'   Set ReportParams = New Collection
   
   Call AddParam("TARGET", Val(txtTarget.Text))
   Call AddParam("ACTUAL_BIRTH", Val(txtActualBirth.Text))
   Call AddParam("DIFF", Val(txtDiff.Text))
   Call AddParam("AVG", Val(txtAvg.Text))
   Call AddParam("BIRTH_DATE", uctlBirthDate.ShowDate)
   Call AddParam("MONTH1", Val(txtMonth1.Text))
   Call AddParam("MONTH2", Val(txtMonth2.Text))
   Call AddParam("MONTH3", Val(txtMonth3.Text))
   Call AddParam("MONTH4", Val(txtMonth4.Text))
   Call AddParam("LEFT1", Val(txtLeft1.Text))
   Call AddParam("LEFT2", Val(txtLeft2.Text))
   Call AddParam("LEFT3", Val(txtLeft3.Text))
   Call AddParam("LEFT4", Val(txtLeft4.Text))
   Call AddParam("MIX1", Val(txtMix1.Text))
   Call AddParam("MIX2", Val(txtMix2.Text))
   Call AddParam("MIX3", Val(txtMix3.Text))
   Call AddParam("MIX4", Val(txtMix4.Text))

   m_ManagementConfig.MANAGEMENT_CONFIG_ID = ID
   m_ManagementConfig.AddEditMode = ShowMode
   m_ManagementConfig.TARGET = txtTarget.Text
   m_ManagementConfig.ACTUAL_BIRTH = txtActualBirth.Text
   m_ManagementConfig.Diff = txtDiff.Text
   m_ManagementConfig.AVERAGE = txtAvg.Text
   m_ManagementConfig.BIRTH_DATE = uctlBirthDate.ShowDate
   m_ManagementConfig.MONTH1 = txtMonth1.Text
   m_ManagementConfig.LEFT1 = txtLeft1.Text
   m_ManagementConfig.MIX1 = txtMix1.Text
   m_ManagementConfig.MONTH2 = txtMonth2.Text
   m_ManagementConfig.LEFT2 = txtLeft2.Text
    m_ManagementConfig.MIX2 = txtMix2.Text
   m_ManagementConfig.MONTH3 = txtMonth3.Text
   m_ManagementConfig.LEFT3 = txtLeft3.Text
   m_ManagementConfig.MIX3 = txtMix3.Text
   m_ManagementConfig.MONTH4 = txtMonth4.Text
   m_ManagementConfig.LEFT4 = txtLeft4.Text
   m_ManagementConfig.MIX4 = txtMix4.Text

   If Not glbDaily.AddEditManagementConfig(m_ManagementConfig, IsOK, True, glbErrorLog) Then
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

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub
Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkExtraFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboOrientation_Click()
   m_HasModify = True
End Sub

Private Sub cboPaperSize_Click()
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_ManagementConfig.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_ManagementConfig.QueryFlag = 0
         Call QueryData(False)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_ManagementConfig = Nothing
   Set m_Houses = Nothing
   Set m_Employees = Nothing
   Set m_PigStatus = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
      
   Call InitNormalLabel(lblTarget, MapText("เป้าหมาย"))
   Call InitNormalLabel(lblActualBirth, MapText("เกิดจริง"))
   Call InitNormalLabel(lblDiff, MapText("ผลต่าง"))
   Call InitNormalLabel(lblAvg, MapText("เฉลี่ย"))
   Call InitNormalLabel(lblBirthDate, MapText("กำหนดคลอด"))
   Call InitNormalLabel(lblMonth, MapText("เดือน"))
   Call InitNormalLabel(lblLeft, MapText("คงเหลือ"))
   Call InitNormalLabel(lblMix, MapText("ผสมไว้"))

   Call txtTarget.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtActualBirth.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtDiff.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtAvg.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
            
   Call txtMonth1.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtMonth2.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtMonth3.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtMonth4.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   
   Call txtLeft1.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtLeft2.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtLeft3.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtLeft4.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   
   Call txtMix1.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtMix2.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtMix3.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtMix4.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
      
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
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_ManagementConfig = New CManagementConfig
   Set m_Houses = New Collection
   Set m_Employees = New Collection
   Set m_PigStatus = New Collection
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtParentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
End Sub

Private Sub txtPaperSize_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub lblCm3_Click()

End Sub

Private Sub txtMarginBottom_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginFooter_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginHeader_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginRight_Change()
   m_HasModify = True
End Sub

Private Sub txtMarginTop_Change()
   m_HasModify = True
End Sub

Private Sub txtPaperHeight_Change()
   m_HasModify = True
End Sub

Private Sub txtPaperWidth_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlHouseLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlResponseByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox10_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox8_Change()
   m_HasModify = True
End Sub

Private Sub txtActualBirth_Change()
Dim DayCount As Long

   DayCount = Day(uctlBirthDate.ShowDate)
   
   m_HasModify = True
   txtDiff.Text = Val(txtTarget.Text) - Val(txtActualBirth.Text)
   txtAvg.Text = MyDiff(txtActualBirth.Text, DayCount)
End Sub

Private Sub txtAvg_Change()
   m_HasModify = True
End Sub

Private Sub txtDiff_Change()
   m_HasModify = True
End Sub

Private Sub txtLeft1_Change()
   m_HasModify = True
End Sub

Private Sub txtLeft2_Change()
   m_HasModify = True
End Sub

Private Sub txtLeft3_Change()
   m_HasModify = True
End Sub

Private Sub txtLeft4_Change()
   m_HasModify = True
End Sub

Private Sub txtMix1_Change()
   m_HasModify = True
End Sub

Private Sub txtMix2_Change()
   m_HasModify = True
End Sub

Private Sub txtMix3_Change()
   m_HasModify = True
End Sub

Private Sub txtMix4_Change()
   m_HasModify = True
End Sub

Private Sub txtMonth1_Change()
   m_HasModify = True
End Sub

Private Sub txtMonth2_Change()
   m_HasModify = True
End Sub

Private Sub txtMonth3_Change()
   m_HasModify = True
End Sub

Private Sub txtMonth4_Change()
   m_HasModify = True
End Sub

Private Sub txtTarget_Change()
   m_HasModify = True
   txtDiff.Text = Val(txtTarget.Text) - Val(txtActualBirth.Text)
End Sub

Private Sub uctlBirthDate_HasChange()
   m_HasModify = True
End Sub
