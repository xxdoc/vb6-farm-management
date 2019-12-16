VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmInitBalance 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmInitBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9465
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4155
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7329
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextBox txtProcess 
         Height          =   465
         Left            =   1980
         TabIndex        =   7
         Top             =   1710
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   1980
         TabIndex        =   6
         Top             =   2220
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1980
         TabIndex        =   0
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtProgress 
         Height          =   465
         Left            =   1980
         TabIndex        =   8
         Top             =   2460
         Width           =   1695
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin VB.Label lblPercent 
         Caption         =   "1"
         Height          =   255
         Left            =   3780
         TabIndex        =   12
         Top             =   2580
         Width           =   225
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2580
         Width           =   1755
      End
      Begin VB.Label lblProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1890
         Width           =   1755
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   570
         TabIndex        =   9
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1110
         Width           =   1755
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7335
         TabIndex        =   2
         Top             =   3240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5685
         TabIndex        =   1
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInitBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Conn As ADODB.Connection

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset
Private m_Rs4 As ADODB.Recordset

Private m_InventoryBalances  As Collection

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
   ProgressBar1.MIN = 0
   ProgressBar1.MAX = 100
   ProgressBar1.Value = 0
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction
   
   Call GenerateBalance
   
   Call glbDaily.CommitTransaction
   
   Call EnableForm(Me, True)
End Sub
Private Sub GenerateBalance()
Dim Ba As CBalanceAccum
Dim Ivd As CInventoryDoc
Dim II As CImportItem

Dim NewDate As Date
Dim I As Long
Dim iCount As Long
Dim IsOK As Boolean
Dim Percent As Double
Dim iCount2 As Long
   txtProcess.Text = "สร้างข้อมูลยอดยกมาหมู"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "LOADING"
   txtProgress.Refresh
         
   Dim m_InventoryBalancePiqs As Collection
   Set m_InventoryBalancePiqs = New Collection
   Call LoadInventoryBalanceForBalance(Nothing, m_InventoryBalancePiqs, uctlFromDate.ShowDate, , , , , , "Y")
   
   txtProcess.Text = "สร้างข้อมูลยอดยกมาวัตถุดิบ"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "LOADING"
   txtProgress.Refresh
   
   Dim m_InventoryBalanceItems As Collection
   Set m_InventoryBalanceItems = New Collection
   Call LoadInventoryBalanceForBalance(Nothing, m_InventoryBalanceItems, uctlFromDate.ShowDate, , , , , , "N")
   
   'Delete All Data
   txtProcess.Text = "ลบข้อมูลเดิมในระบบ"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "PROCESSING"
   txtProgress.Refresh
   Call DeleteBalance(DateAdd("D", -1, uctlFromDate.ShowDate))
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   txtProcess.Text = "สร้างข้อมูลยอดยกมาหมู"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "PROCESSING"
   txtProgress.Refresh
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_DATE = DateAdd("D", -1, uctlFromDate.ShowDate)
   Ivd.DOCUMENT_NO = "ตั้งยอดใหม่หมู"
   Ivd.DOCUMENT_TYPE = 11
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   
   I = 0
   iCount = m_InventoryBalancePiqs.Count
   For Each Ba In m_InventoryBalancePiqs
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      ProgressBar1.Value = Percent
      txtProgress.Text = FormatNumber(Percent)
      txtProgress.Refresh
      
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I"
      II.PART_ITEM_ID = Ba.PART_ITEM_ID
      II.LOCATION_ID = Ba.LOCATION_ID
      II.IMPORT_AMOUNT = Ba.BALANCE_AMOUNT
      II.ACTUAL_UNIT_PRICE = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE
      II.CALCULATE_FLAG = "Y"
      II.TOTAL_ACTUAL_PRICE = Ba.TOTAL_INCLUDE_PRICE
      II.TOTAL_INCLUDE_PRICE = II.TOTAL_ACTUAL_PRICE
      Call Ivd.ImportExports.Add(II)
      Set II = Nothing
   Next Ba
   
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   Set Ivd = Nothing
   Set Ba = Nothing
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   txtProcess.Text = "สร้างข้อมูลยอดยกมาสินค้า"
   txtProcess.Refresh
   ProgressBar1.Value = 0
   txtProgress.Text = "PROCESSING"
   txtProgress.Refresh
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_DATE = DateAdd("D", -1, uctlFromDate.ShowDate)
   Ivd.DOCUMENT_NO = "ตั้งยอดใหม่สินค้า"
   Ivd.DOCUMENT_TYPE = 1
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   
   I = 0
   iCount = m_InventoryBalanceItems.Count
   For Each Ba In m_InventoryBalanceItems
   
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      ProgressBar1.Value = Percent
      txtProgress.Text = FormatNumber(Percent)
      txtProgress.Refresh
      
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I"
      II.PART_ITEM_ID = Ba.PART_ITEM_ID
      II.LOCATION_ID = Ba.LOCATION_ID
      II.IMPORT_AMOUNT = Ba.BALANCE_AMOUNT
      II.ACTUAL_UNIT_PRICE = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
      II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE
      II.CALCULATE_FLAG = "Y"
      II.TOTAL_ACTUAL_PRICE = Ba.TOTAL_INCLUDE_PRICE
      II.TOTAL_INCLUDE_PRICE = II.TOTAL_ACTUAL_PRICE
      Call Ivd.ImportExports.Add(II)
      Set II = Nothing
   Next Ba
   
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   Set Ivd = Nothing
   Set Ba = Nothing
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
         
      Call GetFirstLastDate(Now, FromDate, ToDate)
      ShowMode = SHOW_ADD
      ID = -1
      
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
   pnlHeader.Caption = "ตั้งยอดยกมาระบบใหม่"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call txtProcess.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtProcess.Enabled = False
   Call txtProgress.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtProgress.Enabled = False
   
   Call InitNormalLabel(lblFromDate, "ยอดยกมา ณ วันที่")
   Call InitNormalLabel(lblProcess, "โปรเซส")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "%")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
   Set m_Rs4 = New ADODB.Recordset
   
   Set m_InventoryBalances = New Collection
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_InventoryBalances = Nothing
   
   Set m_Rs1 = Nothing
   Set m_Rs2 = Nothing
   Set m_Rs3 = Nothing
   Set m_Rs4 = Nothing
End Sub
Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub
Private Sub DeleteBalance(ToDate As Date)
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
   
   WhereStr = ""
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr = "" Then
         WhereStr = " WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = "DELETE FROM BALANCE_ACCUM " & WhereStr
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM MOVEMENT_ITEM "
   m_Conn.Execute (SQL1)

   SQL1 = "DELETE FROM CAPITAL_MOVEMENT "
   m_Conn.Execute (SQL1)

   SQL1 = "DELETE FROM LOSS_ITEM "
   m_Conn.Execute (SQL1)

   SQL1 = "DELETE FROM CAPITAL_LOSS "
   m_Conn.Execute (SQL1)
   
   
   WhereStr = ""
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr = "" Then
         WhereStr = " WHERE (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
'   Dim II As CImportItem
'   Dim TempRs As ADODB.Recordset
'   Dim ItemCount As Long
'
'   Set II = New CImportItem
'   Set TempRs = New ADODB.Recordset
'
'   II.FROM_DATE = uctlFromDate.ShowDate
'   Call II.QueryData(1, TempRs, ItemCount)
'
'   Set II = Nothing
   
   SQL1 = "DELETE FROM IMPORT_ITEM II WHERE II.INVENTORY_DOC_ID IN "
   SQL1 = SQL1 & "(SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD " & WhereStr & ")"
   m_Conn.Execute (SQL1)
   
'   While Not TempRs.EOF
'      Set II = New CImportItem
'      Call II.PopulateFromRS(1, TempRs)
'      II.AddEditMode = SHOW_ADD
'      Call II.AddEditData(False)
'      Set II = Nothing
'      TempRs.MoveNext
'   Wend
'
'   If TempRs.State = adStateOpen Then
'      TempRs.Close
'   End If
'   Set TempRs = Nothing
'
'
'   Dim EI As CExportItem
'   Set EI = New CExportItem
'   Set TempRs = New ADODB.Recordset
'
'   EI.FROM_DATE = uctlFromDate.ShowDate
'   Call EI.QueryData(1, TempRs, ItemCount)
'
'   Set EI = Nothing
   
   SQL1 = "DELETE FROM EXPORT_ITEM EI WHERE EI.INVENTORY_DOC_ID IN "
   SQL1 = SQL1 & "(SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD " & WhereStr & ")"
   m_Conn.Execute (SQL1)
   
'   While Not TempRs.EOF
'      Set EI = New CExportItem
'      Call EI.PopulateFromRS(1, TempRs)
'      EI.AddEditMode = SHOW_ADD
'      Call EI.AddEditData(False)
'      Set EI = Nothing
'      TempRs.MoveNext
'   Wend
'
'   If TempRs.State = adStateOpen Then
'      TempRs.Close
'   End If
'   Set TempRs = Nothing
   
   SQL1 = "UPDATE BILLING_DOC BD SET BD.COMMIT_FLAG = 'Y',BD.INVENTORY_DOC_ID = NULL WHERE BD.INVENTORY_DOC_ID IN "
   SQL1 = SQL1 & "(SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD " & WhereStr & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM INVENTORY_DOC IVD " & WhereStr
   m_Conn.Execute (SQL1)
   
End Sub
