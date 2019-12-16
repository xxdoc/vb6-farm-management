VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmLockDocDate 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmLockDocDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2985
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   5265
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   1
         Top             =   1410
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   780
         TabIndex        =   7
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1110
         Width           =   945
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3795
         TabIndex        =   3
         Top             =   2160
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2145
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmLockDocDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private EmpColls As Collection
Private BankBranchColls As Collection
Private SystemParams As Collection
Private m_SP As CSystemParam

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Sp As CSystemParam

   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_SP.ShowMode = ShowMode
   Call m_SP.SetFieldValue("PARAM_ID", ID)
   Call m_SP.SetFieldValue("PARAM_NAME", "DOC_LOCKDATE")
   Call m_SP.SetFieldValue("FROM_LOCK_DATE", uctlFromDate.ShowDate)
   Call m_SP.SetFieldValue("TO_LOCK_DATE", uctlToDate.ShowDate)
   Call m_SP.SetFieldValue("PARAM_TYPE", "S")
   Call m_SP.SetFieldValue("PARAM_DESC", "Valud date interval for all documents")
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction
   Call m_SP.AddEditData
   Call glbDaily.CommitTransaction
   
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

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
       
       Call LoadSystemParam(Nothing, SystemParams)
      Set m_SP = GetSystemParam(SystemParams, "DOC_LOCKDATE")
      If m_SP.GetFieldValue("PARAM_ID") > 0 Then
         uctlFromDate.ShowDate = m_SP.GetFieldValue("FROM_LOCK_DATE")
         uctlToDate.ShowDate = m_SP.GetFieldValue("TO_LOCK_DATE")
         ShowMode = SHOW_EDIT
         ID = m_SP.GetFieldValue("PARAM_ID")
      Else
         Call GetFirstLastDate(Now, FromDate, ToDate)
         uctlFromDate.ShowDate = FromDate
         uctlToDate.ShowDate = ToDate
         ShowMode = SHOW_ADD
         ID = -1
         m_HasModify = True
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "กำหนดช่วงวันที่เอกสาร"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFromDate, "จากวันที่")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
         
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
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   Set EmpColls = New Collection
   Set BankBranchColls = New Collection
   Set SystemParams = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set EmpColls = Nothing
   Set BankBranchColls = Nothing
   Set SystemParams = Nothing
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
