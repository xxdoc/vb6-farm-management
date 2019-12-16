VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateAvgPrice 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmUpdateAvgPrice.frx":0000
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
      TabIndex        =   6
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6376
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2535
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
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
         TabIndex        =   7
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
         TabIndex        =   3
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   6660
         TabIndex        =   1
         Top             =   990
         Width           =   2535
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextLookup uctlSupplierLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   13
         Top             =   1440
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblSupplierNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5010
         TabIndex        =   12
         Top             =   1050
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   4
         Top             =   2730
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmUpdateAvgPrice.frx":27A2
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
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   5
         Top             =   2730
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmUpdateAvgPrice"
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
Private m_Suppliers As Collection
Private m_PriceUpdate As Collection

Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim Ivd As CInventoryDoc
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim Percent As Double
Dim I As Long
Dim HasBegin As Boolean
Dim TempIvd As CInventoryDoc
Dim II As CImportItem
Dim TempRs As ADODB.Recordset
Dim SumAmount  As Double
Dim TempPriceAdjust  As CPriceAdjust
Dim AvgFee As Double
   Set TempRs = New ADODB.Recordset
   If Not VerifyDate(lblFileName, uctlFromDate, False) Then
      Exit Sub
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
   
   HasBegin = False
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Set Ivd = New CInventoryDoc
   
   Call EnableForm(Me, False)
   Ivd.INVENTORY_DOC_ID = -1
   Ivd.FROM_DATE = uctlFromDate.ShowDate
   Ivd.TO_DATE = uctlToDate.ShowDate
   Ivd.DOCUMENT_TYPE = 1
   Ivd.SUPPLIER_ID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   Ivd.QueryFlag = -1
   Call glbDaily.QueryInventoryDoc(Ivd, m_Rs, iCount, IsOK, glbErrorLog)
   RecordCount = iCount
   I = 0
   
   Call glbDaily.StartTransaction
   HasBegin = True
   While Not m_Rs.EOF
      I = I + 1
      Percent = MyDiff(I, RecordCount) * 100
      prgProgress.Value = Percent
      txtPercent.Text = FormatNumber(Percent)
         
      Call Ivd.PopulateFromRS(1, m_Rs)
            
      Set TempIvd = New CInventoryDoc
      TempIvd.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
      TempIvd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(TempIvd, TempRs, iCount, IsOK, glbErrorLog)
      
      SumAmount = 0
      Call TempIvd.PopulateFromRS(1, TempRs)
      For Each II In TempIvd.ImportExports
         II.Flag = "E"
         Set TempPriceAdjust = GetObject("CPriceAdjust", m_PriceUpdate, Trim(Str(II.PART_ITEM_ID)), False)
         If Not (TempPriceAdjust Is Nothing) Then
            II.ACTUAL_UNIT_PRICE = TempPriceAdjust.AVG_PRICE
            II.TOTAL_ACTUAL_PRICE = II.ACTUAL_UNIT_PRICE * II.IMPORT_AMOUNT
         End If
         SumAmount = SumAmount + II.IMPORT_AMOUNT
      Next II
      
      If SumAmount > 0 Then
         AvgFee = MyDiffEx(TempIvd.DELIVERY_FEE, SumAmount)
      Else
         AvgFee = 0
      End If
      
      For Each II In TempIvd.ImportExports
         If II.Flag <> "D" Then
            II.TOTAL_INCLUDE_PRICE = II.TOTAL_ACTUAL_PRICE + (AvgFee * II.IMPORT_AMOUNT)
            II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.IMPORT_AMOUNT)
         End If
      Next II
       
       Call glbDaily.AddEditInventoryDoc(TempIvd, IsOK, False, glbErrorLog)
      Me.Refresh
      m_Rs.MoveNext
   Wend
   prgProgress.Value = 100
   Call glbDaily.CommitTransaction
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
      
      Call LoadSupplier(uctlSupplierLookup.MyCombo, m_Suppliers)
      Set uctlSupplierLookup.MyCollection = m_Suppliers
      
      Call LoadSumUpdateAvg(Nothing, m_PriceUpdate)
      
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
   
   Call InitNormalLabel(lblFileName, "จากวันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   Call InitNormalLabel(lblSupplierNo, MapText("รหัสซัพ ฯ"))
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
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
   Set m_Suppliers = New Collection
   Set m_PriceUpdate = New Collection
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Suppliers = Nothing
   Set m_PriceUpdate = Nothing
End Sub
