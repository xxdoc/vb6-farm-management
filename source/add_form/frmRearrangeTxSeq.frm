VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReArrangeTxSeq 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmRearrangeTxSeq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   12
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   1920
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   1
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   14
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSCheck chkBalanceFlag 
         Height          =   375
         Left            =   6450
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRearrangeTxSeq.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   11
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   4
         Top             =   2910
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   3
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRearrangeTxSeq.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmReArrangeTxSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_Balances As Collection

Private Sub cmdPasswd_Click()

End Sub


Private Sub cboPartType_Click()
   m_HasModify = True
End Sub

Private Sub cboPosition_Click()
   m_HasModify = True
End Sub

Private Sub chkPigFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkBalanceFlag_Click(Value As Integer)
   m_HasModify = True
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
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = ID
      m_Employee.QueryFlag = 1
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Employee.PopulateFromRS(1, m_Rs)
      
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

   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("ADMIN_GROUP_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("ADMIN_GROUP_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
   
'   If Not VerifyCombo(lblPartType, cboPartType, False) Then
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Employee.EMP_ID = ID
   m_Employee.AddEditMode = ShowMode
   m_Employee.PASS_STATUS = "Y"
   
   m_Employee.EmpName.AddEditMode = ShowMode
   m_Employee.EName.AddEditMode = ShowMode
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEmployee(m_Employee, IsOK, True, glbErrorLog) Then
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

Private Function GetNextTransaction(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, II As CImportItem, EI As CExportItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
   
   'Export
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call EI.PopulateFromRS(13, Rs1)
   End If
   
   'Import
   EofFlag2 = Rs2.EOF
   If Not Rs2.EOF Then
      Call II.PopulateFromRS(7, Rs2)
   End If
   
   If (EofFlag1 And EofFlag2) Then
      GetNextTransaction = ""
   ElseIf (EofFlag1 And (Not EofFlag2)) Then
      GetNextTransaction = "I"
      Rs2.MoveNext
   ElseIf ((Not EofFlag1) And EofFlag2) Then
      GetNextTransaction = "E"
      Rs1.MoveNext
   Else
      '===
      '������§�ӴѺ�ռ����ҧ�ҡ
      If DateToStringInt(EI.DOCUMENT_DATE) = DateToStringInt(II.DOCUMENT_DATE) Then
         If EI.PRIORITY1 = II.PRIORITY1 Then
            If EI.DOCUMENT_NO = II.DOCUMENT_NO Then
               If EI.TRANSACTION_SEQ < II.TRANSACTION_SEQ Then
                  GetNextTransaction = "E"
               Else
                  GetNextTransaction = "I"
               End If
            ElseIf EI.DOCUMENT_NO < II.DOCUMENT_NO Then
               GetNextTransaction = "E"
            Else
               GetNextTransaction = "I"
            End If
         ElseIf EI.PRIORITY1 < II.PRIORITY1 Then
            GetNextTransaction = "E"
         Else
            GetNextTransaction = "I"
         End If
      ElseIf DateToStringInt(EI.DOCUMENT_DATE) < DateToStringInt(II.DOCUMENT_DATE) Then
         GetNextTransaction = "E"
      Else
         GetNextTransaction = "I"
      End If 'Document date
      '===
      If GetNextTransaction = "I" Then
         Rs2.MoveNext
      ElseIf GetNextTransaction = "E" Then
         Rs1.MoveNext
      End If
   End If 'Eof flag
End Function

'Public Function GetBalanceAmount(PartItemID As Long, LocationID As Long, TxSeq As Long, DocDate As Date) As Object
'Dim EI As CExportItem
'Dim II As CImportItem
'Dim TempRs As ADODB.Recordset
'Dim iCount As Long
'
'   Set TempRs = New ADODB.Recordset
'
'   Set EI = New CExportItem
'   Set II = New CImportItem
'
'   EI.EXPORT_ITEM_ID = -1
'   EI.PIG_FLAG = "N"
'   EI.PART_ITEM_ID = PartItemID
'   EI.LOCATION_ID = LocationID
'   EI.FROM_TX_SEQ = -1
'   EI.TO_TX_SEQ = TxSeq
'   EI.FROM_DATE = -1
'   EI.TO_DATE = DocDate
'   EI.OrderBy = 11
'   EI.OrderType = 2
'   Call EI.QueryData(1, TempRs, iCount)
'   If Not TempRs.EOF Then
'      Call EI.PopulateFromRS(1, TempRs)
'   End If
'
'   II.IMPORT_ITEM_ID = -1
'   II.PIG_FLAG = "N"
'   II.PART_ITEM_ID = PartItemID
'   II.LOCATION_ID = LocationID
'   II.FROM_TX_SEQ = -1
'   II.TO_TX_SEQ = TxSeq
'   II.FROM_DATE = -1
'   II.TO_DATE = DocDate
'   II.OrderBy = 12
'   II.OrderType = 2
'   Call II.QueryData(1, TempRs, iCount)
'   If Not TempRs.EOF Then
'      Call II.PopulateFromRS(1, TempRs)
'   End If
'
'   If EI.TRANSACTION_SEQ > II.TRANSACTION_SEQ Then
'      Set GetBalanceAmount = EI
'   Else
'      Set GetBalanceAmount = II
'   End If
'
'   If TempRs.State = adStateOpen Then
'      Call TempRs.Close
'   End If
'   Set TempRs = Nothing
'   Set EI = Nothing
'   Set II = Nothing
'End Function

Private Sub GetRelateItem1(II As CImportItem, EI As CExportItem)
Dim iCount As Long
Dim TempRs As ADODB.Recordset

   Set TempRs = New ADODB.Recordset
   
   EI.EXPORT_ITEM_ID = -1
   EI.GUI_ID = II.GUI_ID
   Call EI.QueryData(1, TempRs, iCount)
   If Not TempRs.EOF Then
      Call EI.PopulateFromRS(1, TempRs)
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub cmdStart_Click()
Dim IsOK As Boolean
Dim Yw As CImportItem
Dim Ivd As CInventoryDoc

   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Call EnableForm(Me, False)
   
   Set Ivd = New CInventoryDoc
   Call glbDaily.RearrangeTxSeq(Ivd, IsOK, True, glbErrorLog, prgProgress, txtPercent)
   Set Ivd = Nothing

   Set Yw = New CImportItem
   Yw.IMPORT_ITEM_ID = -1
   Call glbDaily.PatchTxSeq(Yw, IsOK, True, glbErrorLog)
   Set Yw = Nothing
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
'      Call GetFirstLastDate(Now, FromDate, ToDate)
'      uctlFromDate.ShowDate = FromDate
'      uctlToDate.ShowDate = ToDate
      
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

Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "�Ѵ���§�͡���"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "�ҡ�ѹ���")
   Call InitNormalLabel(lblMasterName, "�֧�ѹ���")
   Call InitNormalLabel(lblProgress, "�����׺˹��")
   Call InitNormalLabel(lblPercent, "����ૹ��")
   Call InitNormalLabel(Label1, "%")

   uctlFromDate.Enable = False
   uctlToDate.Enable = False
   
'   Call InitCheckBox(chkBalanceFlag, "ź�ʹ¡��")
'   chkBalanceFlag.Value = ssCBUnchecked

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdStart, MapText("�����"))
   
   Call ResetStatus
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
   
   Set m_Employee = New CEmployee
   Set m_Rs = New ADODB.Recordset
   Set m_Balances = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub txtLastName_Change()
   m_HasModify = True
End Sub

Private Sub txtCode_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtMaxUser_Change()
   m_HasModify = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Balances = Nothing
End Sub
