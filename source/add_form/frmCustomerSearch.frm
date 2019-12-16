VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCustomerSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Icon            =   "frmCustomerSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4425
      Left            =   0
      TabIndex        =   4
      Top             =   510
      Width           =   8145
      Begin prjBoonmeeGraph.uctlTextBox txtCode 
         Height          =   435
         Left            =   1890
         TabIndex        =   8
         Top             =   240
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3165
         Left            =   0
         TabIndex        =   1
         Top             =   1260
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   5583
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   12
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmCustomerSearch.frx":08CA
         Column(2)       =   "frmCustomerSearch.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmCustomerSearch.frx":0A36
         FormatStyle(2)  =   "frmCustomerSearch.frx":0B92
         FormatStyle(3)  =   "frmCustomerSearch.frx":0C42
         FormatStyle(4)  =   "frmCustomerSearch.frx":0CF6
         FormatStyle(5)  =   "frmCustomerSearch.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmCustomerSearch.frx":0E86
      End
      Begin prjBoonmeeGraph.uctlTextBox txtName 
         Height          =   435
         Left            =   1890
         TabIndex        =   9
         Top             =   690
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   555
         Left            =   6300
         TabIndex        =   0
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   979
         _Version        =   131073
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   690
         Width           =   1605
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   1605
      End
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   555
      Left            =   2123
      TabIndex        =   2
      Top             =   4950
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   979
      _Version        =   131073
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   4103
      TabIndex        =   3
      Top             =   4950
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   979
      _Version        =   131073
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   8145
   End
End
Attribute VB_Name = "frmCustomerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const MODULE_NAME = "frmEmployeeSearch"

Private m_HasActivate As Boolean
Public OKClick As Boolean
Public PersonID As Long
Public PersonName As String
Public PersonLastName As String
Public DefaultPositionID As Long
Public HeaderText As String
Public CREDIT As Long

Private m_Rs As ADODB.Recordset
Private m_Patient As CPatient
Private m_TableName As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Patient.PATIENT_ID = -1
      m_Patient.Name = txtName.Text
      m_Patient.PATIENT_CODE = txtCode.Text
      m_Patient.OrderBy = -1
      m_Patient.OrderType = -1
      If Not glbDaily.QueryPatient(m_Patient, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind

'   Label1.Caption = ItemCount
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
Dim ItemCount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      
      Call QueryData(True)
      m_HasActivate = True
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub cmdCancel_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   PersonID = GridEX1.Value(1)
   PersonName = GridEX1.Value(3) & " " & GridEX1.Value(4)
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_Patient = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1320
   Col.Caption = "รหัส"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 1890
   Col.Caption = "ชื่อ"

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2205
   Col.Caption = "นามสกุล"
      
   Set Col = GridEX1.Columns.Add '5
   Col.Width = 4200
   Col.Caption = "ที่อยู่"
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   lblCaption.Font.Name = GLB_FONT
   lblCaption.Font.Bold = True
   lblCaption.Font.Size = 19
   lblCaption.Caption = HeaderText
   
   Me.BackColor = GLB_FORM_COLOR
   lblCaption.BackColor = GLB_HEAD_COLOR
   
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR

   m_HasActivate = False
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"
   
   Call InitMainButton(cmdSearch, "ค้นหา (F5)")
   Call InitMainButton(cmdCancel, "ยกเลิก (ESC)")
   Call InitMainButton(cmdOK, "ตกลง (F2)")
      
   Call InitNormalLabel(lblName, " ชื่อลูกค้า")
   Call InitNormalLabel(lblCode, "รหัสลูกค้า")
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitGrid
   Call EnableForm(Me, True)
   
   m_HasActivate = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

   Call InitFormLayout
   
   m_HasActivate = False
   m_TableName = "CUSTOMER"
   
   Set m_Rs = New ADODB.Recordset
   Set m_Patient = New CPatient
   
   Call EnableForm(Me, True)
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
   Call cmdOK_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler

   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "GridEX1_UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Values(1) = NVLI(m_Rs("PATIENT_ID"), -1)
   Values(2) = NVLS(m_Rs("PATIENT_CODE"), "")
   Values(3) = NVLS(m_Rs("NAME"), "")
   Values(4) = NVLS(m_Rs("LAST_NAME"), "")
   Values(5) = PackAddress(m_Rs)
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

