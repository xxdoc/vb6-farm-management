VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditDebitCreditNote 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditDebitCreditNote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboEnpAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2310
         Width           =   9585
      End
      Begin VB.ComboBox cboCustomerAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1860
         Width           =   9585
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1410
         Width           =   2325
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1410
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   1
         Top             =   990
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   12
         Top             =   3990
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3225
         Left            =   150
         TabIndex        =   13
         Top             =   4530
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5689
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditDebitCreditNote.frx":27A2
         Column(2)       =   "frmAddEditDebitCreditNote.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditDebitCreditNote.frx":290E
         FormatStyle(2)  =   "frmAddEditDebitCreditNote.frx":2A6A
         FormatStyle(3)  =   "frmAddEditDebitCreditNote.frx":2B1A
         FormatStyle(4)  =   "frmAddEditDebitCreditNote.frx":2BCE
         FormatStyle(5)  =   "frmAddEditDebitCreditNote.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditDebitCreditNote.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000009&
            Height          =   1275
            Left            =   0
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   1635
         End
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotal 
         Height          =   435
         Left            =   1860
         TabIndex        =   7
         Top             =   2730
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   3180
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDueDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   8
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblDueDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   32
         Top             =   2790
         Width           =   1395
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4410
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDebitCreditNote.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10500
         TabIndex        =   2
         Top             =   960
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblSellBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   29
         Top             =   3240
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8250
         TabIndex        =   10
         Top             =   3210
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDebitCreditNote.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   9870
         TabIndex        =   11
         Top             =   3210
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblEnpAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1950
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7890
         TabIndex        =   26
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   1470
         Width           =   1635
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   4140
         TabIndex        =   24
         Top             =   2790
         Width           =   585
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   2820
         Width           =   1545
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   22
         Top             =   1020
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   17
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDebitCreditNote.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   18
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   15
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDebitCreditNote.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDebitCreditNote.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   20
         Top             =   1020
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditDebitCreditNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_Customers As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ReceiptType As Long
Public DebitCreditType As Long
Public Area As Long
Public DocumentType As Long
Public BATCH_ID As Long

Private FileName As String
Private m_SumUnit As Double

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_BillingDoc.BILLING_DOC_ID = ID
      m_BillingDoc.BATCH_ID = BATCH_ID
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_BillingDoc.DOCUMENT_DATE
      txtDocumentNo.Text = m_BillingDoc.DOCUMENT_NO
      If Area = 1 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.CUSTOMER_ID)
         cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.ACCOUNT_ID)
      ElseIf Area = 2 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.SUPPLIER_ID)
         cboAccount.ListIndex = -1
      End If
      cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_BillingDoc.BILLING_ADDRESS_ID)
      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)

      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
      
      chkCommit.Value = FlagToCheck(m_BillingDoc.COMMIT_FLAG)
      cmdAdd.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      chkCommit.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
      
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Sub PopulateGuiID(Bd As CBillingDoc)
Dim Di As CDoItem

   For Each Di In Bd.DoItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(Bd)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(Bd As CBillingDoc) As Long
Dim Di As CDoItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In Bd.DoItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

'Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long) As CLotItem
'Dim Ei As CLotItem
'
'      For Each Ei In Ivd.ImportExports
'         If Ei.LINK_ID = GuiID Then
'            Set GetExportItem = Ei
'            Exit Function
'         End If
'      Next Ei
'End Function
'
'Private Function DO2InventoryDoc(Bd As CBillingDoc, Ivd As CInventoryDoc) As Boolean
'Dim TempRs As ADODB.Recordset
'Dim iCount As Long
'Dim IsOK As Boolean
'Dim Di As CDoItem
'Dim Ei As CLotItem
'
'   Set Ivd = Nothing
'   Set Ivd = New CInventoryDoc
'
'   If Bd.INVENTORY_DOC_ID > 0 Then
'      Set TempRs = New ADODB.Recordset
'
'      Ivd.INVENTORY_DOC_ID = Bd.INVENTORY_DOC_ID
'      Ivd.QueryFlag = 1
'      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
'
'      If TempRs.State = adStateOpen Then
'         TempRs.Close
'      End If
'      Set TempRs = Nothing
'
'      Ivd.AddEditMode = SHOW_EDIT
'   Else
'      Ivd.AddEditMode = SHOW_ADD
'   End If
'
'   Ivd.DOCUMENT_DATE = Bd.DOCUMENT_DATE
'   Ivd.DOCUMENT_NO = Bd.DOCUMENT_NO
'   Ivd.COMMIT_FLAG = Bd.COMMIT_FLAG
'   Ivd.DOCUMENT_TYPE = 10
'
'   For Each Di In Bd.DoItems
'      If Di.Flag = "A" Then
'         Set Ei = New CLotItem
'
'         Ei.TX_TYPE = "E"
'         Ei.Flag = "A"
'         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
'         Ei.LOCATION_ID = Di.LOCATION_ID
'         Ei.TX_AMOUNT = Di.ITEM_AMOUNT
'         Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
'         Ei.TOTAL_INCLUDE_PRICE = Di.TOTAL_PRICE
'         Ei.INCLUDE_UNIT_PRICE = MyDiff(Ei.TOTAL_INCLUDE_PRICE, Ei.TX_AMOUNT)
'         Ei.LINK_ID = Di.LINK_ID
'         Ei.CALCULATE_FLAG = "N"
'
'         Set Ei.SubLotItems = Di.SubLotItems
'
'         Call Ivd.ImportExports.Add(Ei)
'         Set Ei = Nothing
'      ElseIf Di.Flag = "E" Then
'         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
'
'         Ei.Flag = "E"
'         Ei.PART_ITEM_ID = Di.PART_ITEM_ID
'         Ei.LOCATION_ID = Di.LOCATION_ID
'         Ei.TX_AMOUNT = Di.ITEM_AMOUNT
'         Ei.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
'         Ei.CALCULATE_FLAG = "N"
'
'         Set Ei.SubLotItems = Di.SubLotItems
'      ElseIf Di.Flag = "D" Then
'         Set Ei = GetExportItem(Ivd, Di.LINK_ID)
'         If Not (Ei Is Nothing) Then
'            Ei.Flag = "D"
'         End If
'      End If
'   Next Di
'End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc

   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("LEDGER_SELL_" & DocumentType & "_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, True) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_BillingDoc.AddEditMode = ShowMode
   m_BillingDoc.BILLING_DOC_ID = ID
    m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   If Area = 1 Then
      m_BillingDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_BillingDoc.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
      m_BillingDoc.DOCUMENT_TYPE = DocumentType
   ElseIf Area = 2 Then
'      m_BillingDoc.SUPPLIER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'      m_BillingDoc.ACCOUNT_ID = -1
'      If DebitCreditType = 1 Then
'         m_BillingDoc.DOCUMENT_TYPE = 10 'ใบเพิ่มหนี้
'      ElseIf DebitCreditType = 2 Then
'         m_BillingDoc.DOCUMENT_TYPE = 9 'ใบลดหนี้
'      End If
   End If
   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.EXCEPTION_FLAG = "N"
   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_BillingDoc.RECEIPT_TYPE = ReceiptType
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.DUE_DATE = uctlDueDate.ShowDate

   Call EnableForm(Me, False)

   Call glbDaily.StartTransaction
   If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If
   
   Call glbDaily.CommitTransaction
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

Private Sub cboAccount_Click()
   m_HasModify = True
End Sub

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
      KeyAscii = 0
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If Area = 1 Then
      If Not VerifyCombo(lblAccountNo, cboAccount) Then
         Exit Sub
      End If
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = 1
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      
      frmAddDebitCreditNoteItem.Area = Area
      If lMenuChosen = 1 Then
         frmAddDebitCreditNoteItem.InvoiceDOType = 1 'DO
      ElseIf lMenuChosen = 3 Then
         frmAddDebitCreditNoteItem.InvoiceDOType = 2 'Invoice
      End If
      If Area = 1 Then
         frmAddDebitCreditNoteItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
      ElseIf Area = 2 Then
         frmAddDebitCreditNoteItem.AccountID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      End If
      Set frmAddDebitCreditNoteItem.TempCollection = m_BillingDoc.ReceiptItems
      frmAddDebitCreditNoteItem.ShowMode = SHOW_ADD
      If DebitCreditType = 1 Then
         frmAddDebitCreditNoteItem.HeaderText = MapText("เพิ่มรายการใบเพิ่มหนี้")
      ElseIf DebitCreditType = 2 Then
         frmAddDebitCreditNoteItem.HeaderText = MapText("เพิ่มรายการใบลดหนี้")
      End If
      Load frmAddDebitCreditNoteItem
      frmAddDebitCreditNoteItem.Show 1

      OKClick = frmAddDebitCreditNoteItem.OKClick

      Unload frmAddDebitCreditNoteItem
      Set frmAddDebitCreditNoteItem = Nothing

      If OKClick Then
         Call GetTotalPrice

         GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtDocumentNo.Text) = "" Then
      If DebitCreditType = 1 Then
         Call glbDatabaseMngr.GenerateNumber(DBN_NUMBER, No, glbErrorLog)
      ElseIf DebitCreditType = 2 Then
         Call glbDatabaseMngr.GenerateNumber(CDN_NUMBER, No, glbErrorLog)
      End If
      txtDocumentNo.Text = No
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_BillingDoc.ReceiptItems.Remove (ID2)
      Else
         m_BillingDoc.ReceiptItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub CalculateIncludePrice()
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Function VerifyOnwerVersionMenu(Menu As Long, Owner As String) As Boolean

End Function


Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   ID = m_BillingDoc.BILLING_DOC_ID
   m_BillingDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If
      
      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         m_BillingDoc.QueryFlag = 0
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
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
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 1785
   Col.Caption = MapText("เลขที่เอกสาร")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1905
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2460
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดลดหนี้")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 1920
   Col.Caption = MapText("ประเภทเอกสาร")

   Set Col = GridEX1.Columns.Add '7
   Col.Visible = False
   Col.Caption = MapText("DO_ID")

   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1785
   Col.Caption = MapText("สาเหตุ")

   Set Col = GridEX1.Columns.Add '9
   Col.Width = 1710
   Col.Caption = MapText("สถานะ")
   
   Set Col = GridEX1.Columns.Add '10
   Col.Width = 3600
   Col.Caption = MapText("รายละเอียดการลดหนี้")
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 1785
   Col.Caption = MapText("เลขที่เอกสาร")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1905
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2460
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดเพิ่มหนี้")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 1920
   Col.Caption = MapText("ประเภทเอกสาร")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Visible = False
   Col.Caption = MapText("DO_ID")
   
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1785
   Col.Caption = MapText("สาเหตุ")

   Set Col = GridEX1.Columns.Add '9
   Col.Width = 1710
   Col.Caption = MapText("สถานะ")
   
   Set Col = GridEX1.Columns.Add '10
   Col.Width = 3600
   Col.Caption = MapText("รายละเอียดการลดหนี้")
   
End Sub

Private Sub GetTotalPrice()
Dim II As CReceiptItem
Dim Sum1 As Double
Dim Sum2 As Double

   Sum1 = 0
   Sum2 = 0
   For Each II In m_BillingDoc.ReceiptItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.DEBIT_CREDIT_AMOUNT
      End If
   Next II

   txtNetTotal.Text = Format(Sum1, "0.00")
End Sub

Private Sub GetTotalPriceEx()
Dim II As CDoItem
Dim Sum1 As Double

   Sum1 = 0
   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.TOTAL_PRICE
      End If
   Next II

   txtNetTotal.Text = Format(Sum1, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblDueDate, MapText("วันที่ครบดิว"))
   If Area = 1 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
      Call InitNormalLabel(lblSellBy, MapText("ผู้ออกเอกสาร"))
   ElseIf Area = 2 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ซัพ ฯ"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสซัพ ฯ"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่รับเอกสาร"))
      Call InitNormalLabel(lblSellBy, MapText("ผู้รับเอกสาร"))
      cmdAuto.Visible = False
   End If
   Call InitNormalLabel(Label4, MapText("บาท"))
   If DebitCreditType = 1 Then
      Call InitNormalLabel(lblNetTotal, MapText("ยอดเพิ่มหนี้รวม"))
      Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบเพิ่มหนี้"))
   ElseIf DebitCreditType = 2 Then
      Call InitNormalLabel(lblNetTotal, MapText("ยอดลดหนี้รวม"))
      Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบลดหนี้"))
   End If
   
   Call InitCheckBox(chkCommit, "คำนวณ")
   If Area = 1 Then
      lblAccountNo.Visible = True
      cboAccount.Visible = True
   ElseIf Area = 2 Then
      lblAccountNo.Visible = False
      cboAccount.Visible = False
   End If
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetTotal.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitCombo(cboAccount)
   Call InitCombo(cboCustomerAddress)
   Call InitCombo(cboEnpAddress)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   If Area = 2 Then
   cmdPrint.Enabled = False
   End If
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Enabled = False
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   If DebitCreditType = 1 Then
      Call InitGrid2
   ElseIf DebitCreditType = 2 Then
      Call InitGrid1
   End If
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   If DebitCreditType = 1 Then
      TabStrip1.Tabs.Add().Caption = MapText("รายการใบเพิ่มหนี้")
   ElseIf DebitCreditType = 2 Then
      TabStrip1.Tabs.Add().Caption = MapText("รายการใบลดหนี้")
   End If
   cmdEdit.Enabled = False
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
   Set m_BillingDoc = New CBillingDoc
   Set m_Customers = New Collection
   Set m_Employees = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_BillingDoc.ReceiptItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CReceiptItem
      If m_BillingDoc.ReceiptItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.ReceiptItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.RECEIPT_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.DOCUMENT_NO
      Values(4) = DateToStringExtEx2(CR.DOCUMENT_DATE)
      Values(5) = FormatNumber(CR.DEBIT_CREDIT_AMOUNT)
      Values(6) = "ใบส่งสินค้า"
      Values(7) = CR.DO_ID
      Values(8) = CR.REASON_NAME
      Values(9) = CR.PRODUCT_STATUS_NAME
      Values(10) = CR.DESCRIPTION_DETAIL
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Label7_Click()

End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      If DebitCreditType = 1 Then
         Call InitGrid2
      ElseIf DebitCreditType = 2 Then
         Call InitGrid1
      End If
   
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
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

Private Sub txtTruckNo_Change()
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

Private Sub txtIncludeVat_Change()
   m_HasModify = True
End Sub

Private Sub txtIncludeWH_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtVatAmount_Change()
   m_HasModify = True
End Sub

Private Sub CalculateAmount()

End Sub

Private Sub txtVatPercent_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtWH_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtWHAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlDate1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
Dim CustomerID As Long
Dim C As CCustomer

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CustomerID > 0 Then
      Set C = m_Customers(Trim(Str(CustomerID)))
      Call LoadAccount(cboAccount, , CustomerID)
      cboAccount.ListIndex = 1
      
      Call LoadCustomerAddress(cboCustomerAddress, , CustomerID, True)
      uctlDueDate.ShowDate = DateAdd("D", C.CREDIT, uctlDocumentDate.ShowDate)
   Else
      cboAccount.ListIndex = -1
      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlDueDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
