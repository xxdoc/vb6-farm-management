VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckError 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmCheckError.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6376
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   1800
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
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
         Top             =   2130
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Threed.SSOption SSOption2 
         Height          =   495
         Left            =   3720
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin VB.Label lblSupplierNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1500
         Width           =   1455
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   2
         Top             =   2730
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCheckError.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   8
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   3
         Top             =   2730
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCheckError"
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

Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim Percent As Double
Dim I As Long
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim Ivd As CInventoryDocSearch
Dim II As CImportItemSearch
Dim EI As CExportItemSearch

Dim SearchCollection As Collection
   
   Set SearchCollection = New Collection
   
   Call LoadInventoryDocSearch(Nothing, SearchCollection)
   
   HasBegin = False
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Call EnableForm(Me, False)
   If SSOption1.Value = True Then
      Set II = New CImportItemSearch
      II.INVENTORY_DOC_ID = -1
      Call II.QueryData(1, m_Rs, ItemCount)
      
   
      HasBegin = True
      While Not m_Rs.EOF
         I = I + 1
         Percent = MyDiff(I, ItemCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
         
         Call II.PopulateFromRS(1, m_Rs)
         
         Set Ivd = GetObject("CInventoryDocSearch", SearchCollection, Trim(Str(II.INVENTORY_DOC_ID)), False)
         If Ivd Is Nothing Then
            Call MsgBox("Error : IMPORT_ITEM_ID = " & II.IMPORT_ITEM_ID, vbOKOnly, "Error")
         End If
         
          Me.Refresh
         m_Rs.MoveNext
      Wend
   ElseIf SSOption2.Value = True Then
      Set EI = New CExportItemSearch
      EI.INVENTORY_DOC_ID = -1
      Call EI.QueryData(1, m_Rs, ItemCount)
      
   
      HasBegin = True
      While Not m_Rs.EOF
         I = I + 1
         Percent = MyDiff(I, ItemCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
         
         Call EI.PopulateFromRS(1, m_Rs)
         
         Set Ivd = GetObject("CInventoryDocSearch", SearchCollection, Trim(Str(EI.INVENTORY_DOC_ID)), False)
         If Ivd Is Nothing Then
            Call MsgBox("Error : EXPORT_ITEM_ID = " & EI.EXPORT_ITEM_ID, vbOKOnly, "Error")
         End If
         
          Me.Refresh
         m_Rs.MoveNext
      Wend
   
   End If
   prgProgress.Value = 100
   HasBegin = False
   Call EnableForm(Me, True)
   
   Set Ivd = Nothing
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      Call glbDaily.RollbackTransaction
   End If
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
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
   pnlHeader.Caption = "ประมวลผลข้อมูล"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   Call InitOptionEx(SSOption1, "นำเข้า")
   SSOption1.Value = True
   Call InitOptionEx(SSOption2, "จ่ายออก")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
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
   
   Set m_Rs = New ADODB.Recordset
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
