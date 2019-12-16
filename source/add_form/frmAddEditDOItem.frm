VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form frmAddEditDoItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditDOItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7800
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6405
      Left            =   0
      TabIndex        =   20
      Top             =   600
      Width           =   18975
      _ExtentX        =   33470
      _ExtentY        =   11298
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox CboPackageType 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2940
         Width           =   5355
      End
      Begin prjFarmManagement.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   5
         Top             =   2480
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1620
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigStatusLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   2040
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeight 
         Height          =   435
         Left            =   5190
         TabIndex        =   6
         Top             =   2480
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1170
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3420
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   1800
         TabIndex        =   9
         Top             =   3900
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalPrice2 
         Height          =   435
         Left            =   7830
         TabIndex        =   17
         Top             =   3420
         Width           =   1530
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice2 
         Height          =   435
         Left            =   7830
         TabIndex        =   18
         Top             =   3900
         Width           =   1530
         _ExtentX        =   3572
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountAmount 
         Height          =   435
         Left            =   2520
         TabIndex        =   11
         Top             =   4380
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountReason 
         Height          =   435
         Left            =   1800
         TabIndex        =   12
         Top             =   4860
         Width           =   7305
         _ExtentX        =   3572
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   1800
         TabIndex        =   10
         Top             =   4380
         Width           =   705
         _ExtentX        =   3572
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPigNo 
         Height          =   435
         Left            =   1800
         TabIndex        =   0
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtK 
         Height          =   435
         Left            =   7830
         TabIndex        =   16
         Top             =   2940
         Width           =   1530
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdGetWeight2 
         Height          =   525
         Left            =   7320
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2370
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblK 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7200
         TabIndex        =   38
         Top             =   2940
         Width           =   525
      End
      Begin VB.Label lblPigNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   37
         Top             =   330
         Width           =   1485
      End
      Begin Threed.SSCheck chkShowAvg 
         Height          =   375
         Left            =   3840
         TabIndex        =   36
         Top             =   3540
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDiscountReason 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   240
         TabIndex        =   35
         Top             =   4920
         Width           =   1485
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   600
         TabIndex        =   34
         Top             =   4440
         Width           =   1125
      End
      Begin VB.Label lblPackageType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3030
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2280
         TabIndex        =   13
         Top             =   5550
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin VB.Label Label5 
         Height          =   345
         Left            =   8325
         TabIndex        =   32
         Top             =   3510
         Width           =   855
      End
      Begin VB.Label lblAvgPrice2 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6720
         TabIndex        =   31
         Top             =   4020
         Width           =   1125
      End
      Begin VB.Label lblTotalPrice2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6360
         TabIndex        =   30
         Top             =   3540
         Width           =   1485
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   29
         Top             =   3480
         Width           =   1485
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   600
         TabIndex        =   28
         Top             =   4020
         Width           =   1125
      End
      Begin VB.Label Label2 
         Height          =   345
         Left            =   3825
         TabIndex        =   27
         Top             =   3900
         Width           =   855
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   26
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label lblWeight 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3960
         TabIndex        =   25
         Top             =   2480
         Width           =   1125
      End
      Begin VB.Label lblPigStatus 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   24
         Top             =   2100
         Width           =   1485
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   23
         Top             =   780
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3960
         TabIndex        =   14
         Top             =   5550
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5565
         TabIndex        =   15
         Top             =   5550
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   22
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   21
         Top             =   2480
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditDoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection2 As Collection
Public COMMIT_FLAG As String

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Houses As Collection
Private m_Pigs As Collection
Private m_PigStatuss As Collection
Private m_PigTypes As Collection
Private m_PigNo As Long

Public ExtraFlag As String
Public ParentForm As Form

Public PEDIGREE_COST As Double

Public Area As Long
Public CusId As Long

Private intPortID As Integer ' Ex. 1, 2, 3, 4 for COM1 - COM4
Private time_Delay As Long
Private lngStatus As Long
Private strError  As String
Private strData   As String
   
Private LocationLockFlag As Boolean
Private ProductTypeLockFlag As Boolean
Private ProductStatusLockFlag As Boolean
Private OldWeightEn As Boolean
Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub CboPackageType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkShowAvg_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   If Not OKClick Then
      OKClick = False
   End If
   
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
   
   Call InitNormalLabel(lblPart, MapText("สัปดาห์เกิด"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblToLocation, MapText("จากโรงเรือน"))
   Call InitNormalLabel(lblPigStatus, MapText("สถานะหมู"))
   Call InitNormalLabel(lblWeight, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(Label2, MapText("บาท/ก.ก."))
   Call InitNormalLabel(Label5, MapText("บาท/ก.ก."))
   Call InitNormalLabel(lblPigType, MapText("ประเภทสุกร"))
   If ExtraFlag = "Y" Then
      Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม 1"))
      Call InitNormalLabel(lblAvgPrice, MapText("ราคา/ก.ก. 1"))
      Call InitNormalLabel(lblTotalPrice2, MapText("ราคารวม 2"))
      Call InitNormalLabel(lblAvgPrice2, MapText("ราคา/ก.ก. 2"))
   Else
      Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
      Call InitNormalLabel(lblAvgPrice, MapText("ราคา/ก.ก."))
      lblTotalPrice2.Visible = False
      lblAvgPrice2.Visible = False
   End If
   Call InitNormalLabel(lblPackageType, MapText("ประเภทราคา"))
   Call InitNormalLabel(lblDiscountReason, MapText("สาเหตุการลด"))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(lblPigNo, MapText("ลำดับที่"))
   Call InitNormalLabel(lblK, MapText("K"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtAvgPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   txtAvgPrice.Enabled = False
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtDiscountAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtDiscountPercent.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtPigNo.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   txtPigNo.Enabled = False
   Call txtK.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   
   Call InitCombo(CboPackageType)
   Call InitCheckBox(chkShowAvg, "แสดงราคาเฉลี่ย")
   
   If ExtraFlag = "Y" Then
      txtAvgPrice2.Visible = True
      txtTotalPrice2.Visible = True
      Label5.Visible = True
      
      Call txtAvgPrice2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
      txtAvgPrice2.Enabled = False
      Call txtTotalPrice2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Else
      Label5.Visible = False
      txtAvgPrice2.Visible = False
      txtTotalPrice2.Visible = False
   End If
      
   If Not VerifyAccessRight("NOT-LOCK-WEIGHT", "ไม่ LOCK น้ำหนัก", 2) Then
      txtWeight.Enabled = False
   Else
      txtWeight.Enabled = True
   End If
   
   cmdGetWeight2.Picture = LoadPicture(glbParameterObj.NormalButton1)
      
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   'cmdNext.Enabled = (ShowMode = SHOW_ADD)
   
   Call InitMainButton(cmdGetWeight2, MapText("W"))
   
'   txtPort.Text = Val(glbParameterObj.Rs232PortNo)
'   txtDelay.Text = Val(glbParameterObj.TimeDelay)

   intPortID = Val(glbParameterObj.Rs232PortNo)
   time_Delay = Val(glbParameterObj.TimeDelay)
   
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
         Dim Di As CDoItem
         
         Set Di = TempCollection.Item(ID)
         
         uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, PigCodeToID(Di.PIG_TYPE))
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Di.PART_ITEM_ID)
         txtQuantity.Text = Di.ITEM_AMOUNT
         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, Di.LOCATION_ID)
          uctlPigStatusLookup.MyCombo.ListIndex = IDToListIndex(uctlPigStatusLookup.MyCombo, Di.PIG_STATUS)
          CboPackageType.ListIndex = IDToListIndex(CboPackageType, Di.PKG_TYPE)
         txtWeight.Text = Di.TOTAL_WEIGHT
         txtTotalPrice.Text = Di.TOTAL_PRICE
         txtAvgPrice.Text = Di.AVG_PRICE
         txtK.Text = Di.PEDIGREE_COST
         txtDiscountPercent.Text = Di.DISCOUNT_PERCENT
         txtDiscountAmount.Text = Di.DISCOUNT_AMOUNT
         txtDiscountReason.Text = Di.DISCOUNT_REASON
         chkShowAvg.Value = FlagToCheck(Di.SHOW_AVG)
         txtPigNo.Text = m_PigNo
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub
Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim I As Long

   I = 0
   For Each O In Col
      I = I + 1
      If (I > OldID) And (O.Flag <> "D") Then
         GetNextID = I
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function
'Private Sub cmdGetWeight_Click()
'On Error Resume Next
'Dim I As Long
'Dim FirstWeight As Double
'Dim SecondWeight As Double
'
'   intPortID = Val(txtPort.Text)
'   time_Delay = Val(txtDelay.Text)
'
'   ' Initialize Communications
'   Call CommClose(intPortID)
'
'
'
'   lngStatus = CommOpen(intPortID, "COM" & CStr(intPortID), "baud=9600 parity=N data=8 stop=1")
'
'   If lngStatus <> 0 Then
'      ' Handle error.
'      lngStatus = CommGetError(strError)
'      MsgBox "COM Error: " & strError
'      cmdGetWeight.Enabled = True
'      Exit Sub
'   End If
'
'   ' Set modem control lines.
'   lngStatus = CommSetLine(intPortID, LINE_RTS, True)
'   lngStatus = CommSetLine(intPortID, LINE_DTR, True)
'
'   Sleep time_Delay
'
'   ' Read maximum of 64 bytes from serial port.
'   lngStatus = CommRead(intPortID, strData, 64)
'   If lngStatus > 0 Then
'   ' Process data.
'   ElseIf lngStatus < 0 Then
'   ' Handle error.
'   End If
'
'   ' Reset modem control lines.
'   lngStatus = CommSetLine(intPortID, LINE_RTS, False)
'   lngStatus = CommSetLine(intPortID, LINE_DTR, False)
'
'   txtLog.Text = strData
'
'   FirstWeight = GetAbsoluteWeight(strData)
'   ' Close communications.
''   Call CommClose(intPortID)
'
'   If FirstWeight <= 0 Then
'      cmdGetWeight.Enabled = True
'      Exit Sub
'   End If
'
'   '-------------------------------------> รอบสอง
'   ' Initialize Communications
''''   Call CommClose(intPortID)
''''
''''   lngStatus = CommOpen(intPortID, "COM" & CStr(intPortID), "baud=9600 parity=N data=8 stop=1")
'
'   If lngStatus <> 0 Then
'   ' Handle error.
'      lngStatus = CommGetError(strError)
'      MsgBox "COM Error: " & strError
'      cmdGetWeight.Enabled = True
'      Exit Sub
'   End If
'
'   ' Set modem control lines.
'   lngStatus = CommSetLine(intPortID, LINE_RTS, True)
'   lngStatus = CommSetLine(intPortID, LINE_DTR, True)
'
'   Sleep time_Delay
'
'   ' Read maximum of 64 bytes from serial port.
'   lngStatus = CommRead(intPortID, strData, 64)
'   If lngStatus > 0 Then
'   ' Process data.
'   ElseIf lngStatus < 0 Then
'   ' Handle error.
'   End If
'
'   ' Reset modem control lines.
'   lngStatus = CommSetLine(intPortID, LINE_RTS, False)
'   lngStatus = CommSetLine(intPortID, LINE_DTR, False)
'
'   txtLog.Text = txtLog.Text & "-" & strData
'   SecondWeight = GetAbsoluteWeight(strData)
'   ' Close communications.
'   Call CommClose(intPortID)
'
'   If FirstWeight = SecondWeight Then
'      txtWeight.Text = FirstWeight
'   End If
'
'   cmdGetWeight.Enabled = True
'
'
'   glbParameterObj.Rs232PortNo = txtPort.Text
'   glbParameterObj.TimeDelay = txtDelay.Text
'
'   Me.Refresh
'End Sub
Private Function GetAbsoluteWeight(strData As String) As Double
Dim TempStr As String
Dim FirstWightPosition As Long
Dim FirstWight As Double
Dim SecondWightPosition As Long
   
   If Len(strData) > 0 Then
      FirstWightPosition = InStr(1, strData, ",+")
      SecondWightPosition = InStr(1, strData, "kg")
      FirstWight = Val(Mid(strData, FirstWightPosition + 1, SecondWightPosition - FirstWightPosition))
      
      GetAbsoluteWeight = FirstWight
   Else
      GetAbsoluteWeight = 0
   End If
End Function


Private Sub cmdGetWeight2_Click()
On Error GoTo handler
 If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
   txtWeight.Text = ""
   cmdGetWeight2.Enabled = False
   
   MSComm1.CommPort = intPortID 'Val(txtPort.Text)
   MSComm1.Settings = "9600,n,8,1"
   MSComm1.PortOpen = True
   MSComm1.RThreshold = 1
   
   Exit Sub
handler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   cmdGetWeight2.Enabled = True
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long
Dim Di As CDoItem
Dim SumAmount As Long
Dim SumWeight As Double
   
   If Not SaveData Then
      Exit Sub
   End If
   
      If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.ShowDoItemGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      txtWeight.Text = ""
      txtTotalPrice.Text = ""
      uctlPartLookup.SetFocus
      m_PigNo = m_PigNo + 1
      txtPigNo.Text = m_PigNo
   End If
   Call QueryData(True)
   Call ParentForm.ShowDoItemGrid
   If txtWeight.Enabled Then
      txtWeight.SetFocus
   End If
   
   SumAmount = 0
   SumWeight = 0
   For Each Di In TempCollection
      If Di.Flag <> "D" Then
         SumWeight = SumWeight + Di.TOTAL_WEIGHT
         SumAmount = SumAmount + 1
      End If
   Next Di
         
   Me.Caption = "จำนวนตัว " & FormatNumber(SumAmount) & " น้ำหนักรวม " & FormatNumber(SumWeight) & " น้ำหนักเฉลี่ย " & FormatNumber(MyDiff(SumWeight, SumAmount))
   
   OKClick = True
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigStatus, uctlPigStatusLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblWeight, txtWeight, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Di As CDoItem
   If ShowMode = SHOW_ADD Then
      Set Di = New CDoItem
      
      Di.Flag = "A"
      Call TempCollection.Add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If

   Di.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   Di.PART_NO = uctlPartLookup.MyTextBox.Text
   Di.ITEM_AMOUNT = txtQuantity.Text
   Di.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   Di.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
   Di.PIG_TYPE = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
   Di.PIG_STATUS = uctlPigStatusLookup.MyCombo.ItemData(Minus2Zero(uctlPigStatusLookup.MyCombo.ListIndex))
   Di.PKG_TYPE = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
   Di.PIG_STATUS_NAME = uctlPigStatusLookup.MyCombo.Text
   Di.TOTAL_WEIGHT = Val(txtWeight.Text)
   Di.TOTAL_PRICE = Format(Val(txtTotalPrice.Text), "0.00")
   Di.AVG_PRICE = Val(txtAvgPrice.Text)
   Di.AVG_WEIGHT = 0
   Di.PEDIGREE_COST = Val(txtK.Text)
   Di.DISCOUNT_PERCENT = Val(txtDiscountPercent.Text)
   Di.DISCOUNT_AMOUNT = Val(txtDiscountAmount.Text)
   Di.DISCOUNT_REASON = txtDiscountReason.Text
   Di.SHOW_AVG = Check2Flag(chkShowAvg.Value)
      
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadProductType(uctlPigTypeLookup.MyCombo, m_PigTypes)
      Set uctlPigTypeLookup.MyCollection = m_PigTypes
      
      Call LoadLocation(uctlToLocationLookup.MyCombo, m_Houses, 1, "Y")
      Set uctlToLocationLookup.MyCollection = m_Houses

      Call LoadProductStatus(uctlPigStatusLookup.MyCombo, m_PigStatuss)
      Set uctlPigStatusLookup.MyCollection = m_PigStatuss

      Call InitPackageType(CboPackageType)

      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         m_PigNo = ID
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         m_PigNo = CountItem(TempCollection) + 1
         Call QueryData(True)
         chkShowAvg.Value = ssCBChecked
      End If
      
      txtPigNo.Text = m_PigNo
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Houses = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatuss = New Collection
   Set m_PigTypes = New Collection
   m_PigNo = 0
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_Houses = Nothing
   Set m_Pigs = Nothing
   Set m_PigStatuss = Nothing
   Set m_PigTypes = Nothing
   
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

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub Label4_Click()

End Sub


Private Sub MSComm1_OnComm()
On Error GoTo handler
Dim strInput As String
Dim FirstData As String
Dim SecondData As String
With MSComm1
   Select Case .CommEvent
      Case comEvReceive
      
      Sleep time_Delay
      'อ่านรอบแรก
      strInput = .Input
      FirstData = GetAbsoluteWeight(strInput)
      strInput = ""
      
      Sleep 1000
      'อ่านรอบสอง
      strInput = .Input
      SecondData = GetAbsoluteWeight(strInput)
      
      If FirstData = SecondData Then
         txtWeight.Text = FirstData
         cmdGetWeight2.Enabled = True
      Else
         cmdGetWeight2.Enabled = True
      End If
       If .PortOpen = True Then .PortOpen = False
   End Select
End With 'MSComm1
Exit Sub
handler:
   glbErrorLog.LocalErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowUserError
   cmdGetWeight2.Enabled = True
End Sub

Private Sub txtAvgPrice_Change()
   m_HasModify = True
End Sub



Private Sub txtDiscountAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountPercent_Change()
   txtDiscountAmount.Text = Format(Val(txtTotalPrice.Text) * Val(txtDiscountPercent.Text) / 100, "0.00")
   m_HasModify = True
End Sub

Private Sub txtDiscountReason_Change()
   m_HasModify = True
End Sub

Private Sub txtK_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
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

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalPrice_Change()
   m_HasModify = True
   If Val(txtWeight.Text) > 0 Then
      txtAvgPrice.Text = Format(Val(txtTotalPrice.Text) / Val(txtWeight.Text), "0.00")
   Else
      txtAvgPrice.Text = "0.00"
   End If
End Sub

Private Sub txtWeight_Change()
   m_HasModify = True
   If Val(txtWeight.Text) > 0 Then
      txtAvgPrice.Text = Format(Val(txtTotalPrice.Text) / Val(txtWeight.Text), "0.00")
      Call CboPackageType_Click
   Else
      txtAvgPrice.Text = "0.00"
   End If
End Sub

Private Sub CboPackageType_Click()
On Error Resume Next
Dim D As CCustomerPackage
Dim PkgDetail As CPackageDetail
Dim ID1 As Long
Dim ID2 As Long
Dim ID3 As Long
   
   ID1 = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
   ID2 = uctlPigStatusLookup.MyCombo.ItemData(Minus2Zero(uctlPigStatusLookup.MyCombo.ListIndex))
   
    If ID1 <= 0 Or ID2 <= 0 Then
         txtTotalPrice.Text = 0
         PEDIGREE_COST = 0
         txtK.Text = PEDIGREE_COST
        Exit Sub
    End If
    
    Set D = CustomerPackage(CusId & "-" & ID1)
    
    For Each PkgDetail In PackageDetail
      If D Is Nothing Then
         If PkgDetail.PKG_BASIC_FLAG = "Y" And PkgDetail.STATUS_BUY_ID = ID2 And PkgDetail.PKG_TYPE = ID1 Then
            If MyDiff(Val(txtWeight.Text), Val(txtQuantity.Text)) >= PkgDetail.FROM_WEIGHT And MyDiff(Val(txtWeight.Text), Val(txtQuantity.Text)) <= PkgDetail.TO_WEIGHT Then
               Exit For
            End If
         End If
      Else
         If PkgDetail.PKG_ID = D.PKG_ID And PkgDetail.STATUS_BUY_ID = ID2 Then
            If MyDiff(Val(txtWeight.Text), Val(txtQuantity.Text)) >= PkgDetail.FROM_WEIGHT And MyDiff(Val(txtWeight.Text), Val(txtQuantity.Text)) <= PkgDetail.TO_WEIGHT Then
               Exit For
            End If
         End If
      End If
   Next PkgDetail
   
   If Not (PkgDetail Is Nothing) Then
      If ID1 = 4 Then
         txtTotalPrice.Text = Format(Val(txtQuantity.Text) * (PkgDetail.COST_PER_UNIT + (MyDiffEx(Val(txtWeight.Text), Val(txtQuantity.Text)) * PkgDetail.COST_PER_WEIGHT) + ((MyDiffEx(Val(txtWeight.Text), Val(txtQuantity.Text)) - PkgDetail.CUT_WEIGHT) * PkgDetail.COST_PER_EXCEED)), "0.00")
         PEDIGREE_COST = Val(txtQuantity.Text) * PkgDetail.COST_PER_UNIT
         txtK.Text = PEDIGREE_COST
      ElseIf ID1 = 5 Then
         txtTotalPrice.Text = Format(Val(txtQuantity.Text) * (MyDiff(Val(txtWeight.Text), Val(txtQuantity.Text)) * PkgDetail.COST_PER_WEIGHT), "0.00")
         PEDIGREE_COST = 0
         txtK.Text = PEDIGREE_COST
      End If
   Else
      txtTotalPrice.Text = 0
      PEDIGREE_COST = 0
      txtK.Text = PEDIGREE_COST
   End If
    
    m_HasModify = True
End Sub

Private Sub uctlPigStatusLookup_Change()
   Call CheckEnableWeight
   
    Call CboPackageType_Click
   m_HasModify = True
End Sub

Private Sub uctlPigTypeLookup_Change()
Dim PigTypeCode As String
   
   m_HasModify = True
   
   PigTypeCode = PigTypeToCode(uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
   If PigTypeCode <> "" Then
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Pigs, -1, "Y", PigTypeCode)
      Set uctlPartLookup.MyCollection = m_Pigs
   End If
   
   Call CheckEnableWeight
End Sub

Private Sub uctlToLocationLookup_Change()
   Call CheckEnableWeight
   m_HasModify = True
End Sub
Private Sub CheckEnableWeight()
Dim LocationID As Long
Dim ProductStatusID As Long
Dim ProductTypeID As Long

Dim TempLocation As CLocation
Dim TempProductStatus As CProductStatus
Dim TempProductType As CProductType
   
   If Not VerifyAccessRight("NOT-LOCK-WEIGHT", "ไม่ LOCK น้ำหนัก", 2) Then
      txtWeight.Enabled = False
   Else
      txtWeight.Enabled = True
      Exit Sub
   End If
   
   LocationID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   ProductTypeID = uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex))
   ProductStatusID = uctlPigStatusLookup.MyCombo.ItemData(Minus2Zero(uctlPigStatusLookup.MyCombo.ListIndex))
   
   If Not (LocationID > 0 And ProductTypeID > 0 And ProductStatusID > 0) Then
      txtWeight.Enabled = False
      Exit Sub
   End If
   
   Set TempProductStatus = GetObject("", m_PigStatuss, Trim(Str(ProductStatusID)))
   Set TempLocation = GetObject("", m_Houses, Trim(Str(LocationID)))
   Set TempProductType = GetObject("", m_PigTypes, Trim(Str(ProductTypeID)))
   
   If TempLocation.LOCK_WEIGHT_FLAG = "Y" And TempProductStatus.LOCK_WEIGHT_FLAG = "Y" And TempProductType.LOCK_WEIGHT_FLAG = "Y" Then
      If OldWeightEn Then
         txtWeight.Text = ""
      End If
      txtWeight.Enabled = False
   Else
      txtWeight.Enabled = True
   End If
   
   OldWeightEn = txtWeight.Enabled
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlParttypeLookup_Change()
Dim PartTypeID As Long
   
   Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
   Set uctlPartLookup.MyCollection = m_Parts
   
   m_HasModify = True
End Sub

Private Sub uctlPigWeekLookup_Change()
   m_HasModify = True
End Sub
