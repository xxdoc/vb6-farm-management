VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmUpdateCloseBilling 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmUpdateCloseBilling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   11456
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   1440
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
         TabIndex        =   2
         Top             =   1770
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   990
         Width           =   2535
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   3315
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   9975
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1050
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   3
         Top             =   2250
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmUpdateCloseBilling.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   9
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   4
         Top             =   2250
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmUpdateCloseBilling"
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
Private Sub cmdStart_Click()
On Error GoTo ErrorHandler
Dim IsOK As Boolean

Dim itemcount As Long
Dim RecordCount As Long
Dim Percent As Double
Dim I As Long
Dim HasBegin As Boolean

Dim m_BillingDoc As CBillingDoc
Dim TempRs As ADODB.Recordset
Dim BalanceAmount  As Double
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim TempRcp As CReceiptItem
Dim m_PaidAmounts  As Collection
Dim m_DnAmounts As Collection
Dim m_CnAmounts As Collection
Dim CountEx As Long
   Set m_DnAmounts = New Collection
   Set m_CnAmounts = New Collection
   Set m_PaidAmounts = New Collection
   
   Set TempRs = New ADODB.Recordset
   
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
   
   HasBegin = False
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   Call EnableForm(Me, False)
   
   Set TempRcp = New CReceiptItem
   Set m_BillingDoc = New CBillingDoc
   m_BillingDoc.COMMIT_FLAG = ""
   m_BillingDoc.TO_DATE = uctlToDate.ShowDate
   m_BillingDoc.DOCUMENT_TYPE = 1
   m_BillingDoc.VALID_DATE = DateAdd("D", 1, uctlToDate.ShowDate)
   m_BillingDoc.ItemSumFlag = True
   m_BillingDoc.OrderType = 1
   
   Call m_BillingDoc.SetFlag(False, True, False, False, False, False)
   If Not glbDaily.QueryBillingDoc(m_BillingDoc, TempRs, itemcount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, -1, uctlToDate.ShowDate, , , , uctlToDate.ShowDate)
   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, -1, uctlToDate.ShowDate, 3, 2, uctlToDate.ShowDate)
   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, -1, uctlToDate.ShowDate, 4, 2, uctlToDate.ShowDate)
   
   Call glbDaily.StartTransaction
   
   If itemcount > 0 Then
      CountEx = 0
      RecordCount = itemcount
      I = 0
      HasBegin = True
      While Not TempRs.EOF
         I = I + 1
         Percent = MyDiff(I, RecordCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
                  
         Call m_BillingDoc.PopulateFromRS(1, TempRs)
         
         Set Ri1_0 = GetReceiptItem(m_PaidAmounts, m_BillingDoc.BILLING_DOC_ID) 'รับชำระ
         Set Ri1_1 = GetReceiptItem(m_DnAmounts, m_BillingDoc.BILLING_DOC_ID) 'เพิ่มหนี้
         Set Ri1_2 = GetReceiptItem(m_CnAmounts, m_BillingDoc.BILLING_DOC_ID) 'ลดหนี้
         
         m_BillingDoc.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
         m_BillingDoc.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
         m_BillingDoc.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
        
         If (m_BillingDoc.DO_TOTAL_PRICE + m_BillingDoc.REVENUE_TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT + (m_BillingDoc.DEBIT_AMOUNT - m_BillingDoc.CREDIT_AMOUNT) - m_BillingDoc.PAID_AMOUNT) = 0 Then
            'ยอดหนี้ของบิลนั้นเป็น 0 ให้ Update CLOSE_FLAG เป็น Y เลย
            CountEx = CountEx + 1
            m_BillingDoc.VALID_DATE = uctlToDate.ShowDate
            ''debug.print m_BillingDoc.BILLING_DOC_ID
            Call m_BillingDoc.UpdateValidDate
            
            TempRcp.VALID_DATE = uctlToDate.ShowDate
            TempRcp.DO_ID = m_BillingDoc.BILLING_DOC_ID
            Call TempRcp.UpdateValidDate
         End If
         
         Me.Refresh
         TempRs.MoveNext
      Wend
   End If
      
   prgProgress.Value = 100
   Call glbDaily.CommitTransaction
   HasBegin = False
   Call EnableForm(Me, True)
       
    glbErrorLog.LocalErrorMsg = "สำเร็จ  จำนวน " & CountEx & " รายการ"
    Call glbErrorLog.ShowUserError
    OKClick = False
   Unload Me
         
   Set m_BillingDoc = Nothing
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
   pnlHeader.Caption = "อัพเดดข้อมูลยอดหนี้ เพื่อปิดบิลที่ชำระแล้ว"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   
   Call InitNormalLabel(lblNote, "ส่วนนี้เป็นส่วนที่จะคอยอัพเดด บิลที่จ่ายหมดแล้ว ซึ่งจะไม่ได้เอาไปรวมกับการออกรายงานหนี้คงค้างรายตัว (ลดเวลาลงได้มาก) โดยวิธีการอัพเดดนั้นต้องอัพเดดเฉพาะบิลที่เราไม่ได้ไปดูย้อนหลังแล้วเท่านั้น เช่น ตอนนี้วันที่ 18/06/2552 เราคงไม่ได้ออกรายงานค้างถึงวันที่ 31/12/2551 เป็นแน่ดังนั้นก็ให้อัพพเดด ถึงวันที่ 31/12/2551 เลย หรือกรณีที่ปัจจุบันเป็นวันที่ 28/10/2552 และไม่มีความจำเป็นต้องออกรายงานหนี้คงค้างย้อนหลังมาก เราก็สามารถปรับถึงวันที่ 30/06/2552 ก็ได้ มีผลโดยตรงกับรายงานลูกหนี้ค้างชำระรายรายตัว AR001")
   
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
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
