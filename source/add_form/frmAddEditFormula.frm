VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddEditFormula 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   Icon            =   "frmAddEditFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   6324
      _Version        =   131073
      Begin VB.ComboBox cboYCollection 
         Height          =   315
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   3105
      End
      Begin VB.ComboBox cboXCollection 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   2895
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   9
         Top             =   2880
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   1244
         _Version        =   131073
         Begin Threed.SSCommand cmdCancel 
            Cancel          =   -1  'True
            Height          =   615
            Left            =   4620
            TabIndex        =   8
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdOK 
            Height          =   615
            Left            =   2535
            TabIndex        =   7
            Top             =   60
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   1244
         _Version        =   131073
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2640
            Top             =   7590
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   28
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":014A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":0464
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":0D3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":34F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":3DCA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":46A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":4F7E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":5858
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":6132
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":6A0C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":6E5E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":7738
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":8012
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":88EC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":91C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":9618
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":9A6A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":9BC4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":A49E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":AD78
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":B652
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":B96C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":C246
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":CF20
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":D7FA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":E0D4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":E9AE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAddEditFormula.frx":F288
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin prjFarmManagement.uctlTextBox txtNote1 
         Height          =   405
         Left            =   1620
         TabIndex        =   1
         Top             =   960
         Width           =   2475
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtNote2 
         Height          =   405
         Left            =   1620
         TabIndex        =   3
         Top             =   1380
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtXStart 
         Height          =   435
         Left            =   1620
         TabIndex        =   6
         Top             =   2220
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   767
      End
      Begin VB.Label lblXStart 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   2340
         Width           =   1425
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   8280
         TabIndex        =   2
         Top             =   840
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         PictureFrames   =   1
         BackStyle       =   1
         Picture         =   "frmAddEditFormula.frx":FB62
         ButtonStyle     =   3
      End
      Begin VB.Label lblYCollection 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   4560
         TabIndex        =   14
         Top             =   1890
         Width           =   1155
      End
      Begin VB.Label lblXCollection 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1890
         Width           =   1425
      End
      Begin VB.Label lblNote2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label lblNote1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1050
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmAddEditFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
'Private m_Customer As CCustomer

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private m_Formula As CFormula
Private m_Points As Collection

Private Sub InitFormLayout()
Dim i As Long

   pnlHeader.Caption = HeaderText
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlFooter.BackColor = GLB_FORM_COLOR
      
   Call InitNormalLabel(lblNote1, "ชื่อ")
   Call InitNormalLabel(lblNote2, "รายละเอียด")
   Call InitNormalLabel(lblXCollection, "ตัวเลข")
   Call InitNormalLabel(lblYCollection, "กลุ่มตัวเลข")
   Call InitNormalLabel(lblXStart, "เริ่มจาก")
   
   Call txtNote1.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtNote2.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtXStart.SetTextLenType(TEXT_INTEGER, glbSetting.CODE_TYPE)

   Call InitCombo(cboXCollection)
   Call InitCombo(cboYCollection)
'   Call InitMainButton(cmdAdd, "เพิ่ม (F7)")
'   Call InitMainButton(cmdEdit, "แก้ไข (F3)")
'   Call InitMainButton(cmdDelete, "ลบ (F6)")
   
   Call InitMainButton(cmdOK, "ตกลง (F2)")
   Call InitMainButton(cmdCancel, "ยกเลิก (ESC)")
End Sub

Private Sub cboStatus_Click()
   m_HasModify = True
End Sub

Private Sub Check1_Click()
   m_HasModify = True
End Sub

Private Sub Check10_Click()
   m_HasModify = True
End Sub

Private Sub Check11_Click()
   m_HasModify = True
End Sub

Private Sub Check12_Click()
   m_HasModify = True
End Sub

Private Sub Check13_Click()
   m_HasModify = True
End Sub

Private Sub Check14_Click()
   m_HasModify = True
End Sub

Private Sub Check15_Click()
   m_HasModify = True
End Sub

Private Sub Check16_Click()
   m_HasModify = True
End Sub

Private Sub Check17_Click()
   m_HasModify = True
End Sub

Private Sub Check18_Click()
   m_HasModify = True
End Sub

Private Sub Check19_Click()
   m_HasModify = True
End Sub

Private Sub Check2_Click()
   m_HasModify = True
End Sub

Private Sub Check20_Click()
   m_HasModify = True
End Sub

Private Sub Check21_Click()
   m_HasModify = True
End Sub

Private Sub Check22_Click()
   m_HasModify = True
End Sub

Private Sub Check23_Click()
   m_HasModify = True
End Sub

Private Sub Check24_Click()
   m_HasModify = True
End Sub

Private Sub Check25_Click()
   m_HasModify = True
End Sub

Private Sub Check26_Click()
   m_HasModify = True
End Sub

Private Sub Check27_Click()
   m_HasModify = True
End Sub

Private Sub Check28_Click()
   m_HasModify = True
End Sub

Private Sub Check29_Click()
   m_HasModify = True
End Sub

Private Sub Check3_Click()
   m_HasModify = True
End Sub

Private Sub Check30_Click()
   m_HasModify = True
End Sub

Private Sub Check31_Click()
   m_HasModify = True
End Sub

Private Sub Check32_Click()
   m_HasModify = True
End Sub

Private Sub Check33_Click()
   m_HasModify = True
End Sub

Private Sub Check34_Click()
   m_HasModify = True
End Sub

Private Sub Check35_Click()
   m_HasModify = True
End Sub

Private Sub Check36_Click()
   m_HasModify = True
End Sub

Private Sub Check4_Click()
   m_HasModify = True
End Sub

Private Sub Check5_Click()
   m_HasModify = True
End Sub

Private Sub Check6_Click()
   m_HasModify = True
End Sub

Private Sub Check7_Click()
   m_HasModify = True
End Sub

Private Sub Check8_Click()
   m_HasModify = True
End Sub

Private Sub Check9_Click()
   m_HasModify = True
End Sub

Private Sub chkBerk_Click()
   m_HasModify = True
End Sub

Private Sub chkChild_Click()
   m_HasModify = True
End Sub

Private Sub chkHusband_Click()
   m_HasModify = True
End Sub

Private Sub chkNoJob_Click()
   m_HasModify = True
End Sub

Private Sub chkPay_Click()
   m_HasModify = True
End Sub

Private Sub chkWife_Click()
   m_HasModify = True
End Sub

Private Sub cboXCollection_Click()
   m_HasModify = True
End Sub

Private Sub cboYCollection_Click()
   m_HasModify = True
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Function VerifyControl() As Boolean
   VerifyControl = False
   
   If Not VerifyTextControl(lblNote1, txtNote1, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblXCollection, cboXCollection, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblYCollection, cboYCollection, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblXStart, txtXStart, False) Then
      Exit Function
   End If
   
   VerifyControl = True
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("DAILY_DAILY_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("DAILY_DAILY_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
      
   If Not VerifyControl Then
      Exit Function
   End If
               
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Formula.FORMULA_ID = ID
   m_Formula.AddEditMode = ShowMode
    m_Formula.FORMULA_NAME = txtNote1.Text
    m_Formula.FORMULA_DESC = txtNote2.Text
    m_Formula.X_COLLECTION_ID = cboXCollection.ItemData(Minus2Zero(cboXCollection.ListIndex))
    m_Formula.Y_COLLECTION_ID = cboYCollection.ItemData(Minus2Zero(cboYCollection.ListIndex))
    m_Formula.X_START = Val(txtXStart.Text)
    
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditFormula(m_Formula, IsOK, glbErrorLog) Then
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

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
            
      m_Formula.FORMULA_ID = ID
      m_Formula.QueryFlag = 1
      If Not glbDaily.QueryFormula(m_Formula, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   Else
      IsOK = True
   End If
   
   If ItemCount > 0 Then
      txtNote1.Text = NVLS(m_Rs("FORMULA_NAME"), "")
      txtNote2.Text = NVLS(m_Rs("FORMULA_DESC"), "")
      cboXCollection.ListIndex = IDToListIndex(cboXCollection, NVLI(m_Rs("X_COLLECTION_ID"), -1))
      cboYCollection.ListIndex = IDToListIndex(cboYCollection, NVLI(m_Rs("Y_COLLECTION_ID"), -1))
      txtXStart.Text = NVLI(m_Rs("X_START"), 0)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function Value2Group(Value As Double) As Long
Dim YCollection As CYCollection
Dim j As Long
Dim i As Long
Dim Temp(10) As String

   Set YCollection = m_Formula.YCollection
   Temp(1) = YCollection.MASK1
   Temp(2) = YCollection.MASK2
   Temp(3) = YCollection.MASK3
   Temp(4) = YCollection.MASK4
   Temp(5) = YCollection.MASK5
   Temp(6) = YCollection.MASK6
   Temp(7) = YCollection.MASK7
   Temp(8) = YCollection.MASK8
   Temp(9) = YCollection.MASK9
   Temp(10) = YCollection.MASK10
   
   j = 0
   For i = 1 To 10
      If Mid(Temp(i), Value + 1, 1) = "Y" Then
         j = j + 1
         Value2Group = j
         Exit Function
      Else
         If (Temp(i) <> "NNNNNNNNNN") Then
            j = j + 1
         End If
      End If
   Next i
End Function

Private Sub GeneratePoint()
Dim p As CPoint
Dim XItem As CXItem
Dim i As Long

   Set m_Points = Nothing
   Set m_Points = New Collection
   i = 0
   
   For Each XItem In m_Formula.XCollection.XItems
      i = i + 1
      
      Set p = New CPoint
      
      p.X = i
      p.Y = Value2Group(XItem.ITEM_VALUE)
      p.Value = XItem.ITEM_VALUE
      p.ValueDate = XItem.ITEM_DATE
      
      Call m_Points.Add(p)
      Set p = Nothing
   Next XItem
End Sub

Private Sub GenerateGroupCollection(Col As Collection)
Dim YCollection As CYCollection
Dim Temp(10) As String
Dim i As Long
Dim TempStr As String
Dim j As Long
Dim TempTok As String
Dim Cg As CYGroup

   Set YCollection = m_Formula.YCollection
   Temp(1) = YCollection.MASK1
   Temp(2) = YCollection.MASK2
   Temp(3) = YCollection.MASK3
   Temp(4) = YCollection.MASK4
   Temp(5) = YCollection.MASK5
   Temp(6) = YCollection.MASK6
   Temp(7) = YCollection.MASK7
   Temp(8) = YCollection.MASK8
   Temp(9) = YCollection.MASK9
   Temp(10) = YCollection.MASK10
   
   Set Col = Nothing
   Set Col = New Collection
   
   For i = 1 To 10
      TempStr = Temp(i)
      TempTok = ""
      For j = 1 To 10
         If Mid(TempStr, j, 1) = "Y" Then
            TempTok = TempTok & Trim(Str(j - 1)) & ","
         End If
      Next j
      
      If (TempTok <> "") Then
         Mid(TempTok, Len(TempTok), 1) = ")"
         TempTok = "(" & TempTok
      End If
      
      If TempTok <> "" Then
         Set Cg = New CYGroup
         Cg.Y_GROUP = TempTok
         Call Col.Add(Cg)
         Set Cg = Nothing
      End If
   Next i
End Sub

Private Sub cmdPrint_Click()
Dim p As CPoint
Dim Report As CReportInterface
Dim TempCol As Collection

   Call GeneratePoint
   Call GenerateGroupCollection(TempCol)
   
   For Each p In m_Points
      Debug.Print "(" & p.X & "," & p.Y & "," & p.Value & ")"
   Next p
             
   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณากดปุ่ม ตกลง เพื่อนบันทึกข้อมูลให้เรียบร้อยก่อนพิมพ์กราฟ"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
             
   Call EnableForm(Me, False)
   Set Report = New CReportGraph001
   Call Report.AddParam(ID, "FORMULA_ID")
   Call Report.AddParam(txtNote1.Text, "FORMULA_NAME")
   Call Report.AddParam(txtNote2.Text, "FORMULA_DESC")
   Call Report.AddParam(cboXCollection.Text, "X_COLLECTION_NAME")
   Call Report.AddParam(cboYCollection.Text, "Y_COLLECTION_NAME")
   Call Report.AddParam(Val(txtXStart.Text), "X_START")
   Call Report.AddParam(50, "X_WINDOW")
   Call Report.AddParam(m_Points, "POINTS")
   Call Report.AddParam(TempCol, "GROUPS")
   
   Set frmReport.ReportObject = Report
   frmReport.HeaderText = "พิมพ์ " & pnlHeader.Caption
   Load frmReport
   frmReport.Show 1
      
   Unload frmReport
   Set frmReport = Nothing
   Set Report = Nothing
      
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadXCollection(cboXCollection)
      Call LoadYCollection(cboYCollection)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Formula.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_Load()
   Set m_Formula = New CFormula
   Set m_Rs = New ADODB.Recordset
   Set m_Points = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Formula = Nothing
   Set m_Points = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub radAllow_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub radUnAllow_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub txtAge_Change()
   m_HasModify = True
End Sub

Private Sub txtCardNo_Change()
   m_HasModify = True
End Sub

Private Sub txtCD4_Change()
   m_HasModify = True
End Sub

Private Sub txtChannel_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtEquivalence_Change()
   m_HasModify = True
End Sub

Private Sub txtExpense1_Change()
   m_HasModify = True
End Sub

Private Sub txtGender_Change()
   m_HasModify = True
End Sub

Private Sub txtHeight_Change()
   m_HasModify = True
End Sub

Private Sub txtHome_Change()
   m_HasModify = True
End Sub

Private Sub txtJob_Change()
   m_HasModify = True
End Sub

Private Sub txtKhate_Change()
   m_HasModify = True
End Sub

Private Sub txtKwang_Change()
   m_HasModify = True
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub txtMoo_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtOther1_Change()
   m_HasModify = True
End Sub

Private Sub txtOther2_Change()
   m_HasModify = True
End Sub

Private Sub txtOther3_Change()
   m_HasModify = True
End Sub

Private Sub txtOther4_Change()
   m_HasModify = True
End Sub

Private Sub txtOther5_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone2_Change()
   m_HasModify = True
End Sub

Private Sub txtPreWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtReason_Change()
   m_HasModify = True
End Sub

Private Sub txtReference_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSalary_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtViral_Change()
   m_HasModify = True
End Sub

Private Sub txtKS_Change()
   m_HasModify = True
End Sub

Private Sub txtLog10_Change()
   m_HasModify = True
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txtNote1_Change()
   m_HasModify = True
End Sub

Private Sub txtNote2_Change()
   m_HasModify = True
End Sub

Private Sub txtVL_Change()
   m_HasModify = True
End Sub

Private Sub txtWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtYearKnow_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlDate1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDate2_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlRegisterDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox10_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox11_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox12_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox13_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox14_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox15_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox16_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox17_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox18_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox19_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox2_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox3_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox4_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox5_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox6_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox7_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox8_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox9_Change()
   m_HasModify = True
End Sub

Private Sub txtPatient_Change()
   m_HasModify = True
End Sub

Private Sub uctlRecordDate_HasChange()
   m_HasModify = True
End Sub

Private Sub txtXStart_Change()
   m_HasModify = True
End Sub
