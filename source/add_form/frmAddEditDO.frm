VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditDO 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditDO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      TabIndex        =   24
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboArea 
         Height          =   315
         Left            =   9210
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3120
         Width           =   2265
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2805
         Left            =   2400
         TabIndex        =   18
         Top             =   4920
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4948
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
         Column(1)       =   "frmAddEditDO.frx":27A2
         Column(2)       =   "frmAddEditDO.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditDO.frx":290E
         FormatStyle(2)  =   "frmAddEditDO.frx":2A6A
         FormatStyle(3)  =   "frmAddEditDO.frx":2B1A
         FormatStyle(4)  =   "frmAddEditDO.frx":2BCE
         FormatStyle(5)  =   "frmAddEditDO.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditDO.frx":2D5E
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2835
         Left            =   120
         TabIndex        =   39
         Top             =   4920
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   5001
         _Version        =   131073
         PictureBackgroundStyle=   1
         Begin VB.ComboBox cboBankBranch 
            Height          =   315
            Left            =   6420
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   750
            Width           =   4905
         End
         Begin VB.ComboBox cboBank 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   750
            Width           =   2925
         End
         Begin prjFarmManagement.uctlTextBox txtCheckNo 
            Height          =   435
            Left            =   6420
            TabIndex        =   41
            Top             =   300
            Width           =   3375
            _ExtentX        =   2831
            _ExtentY        =   767
         End
         Begin VB.ComboBox cboPaymentType 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   300
            Width           =   2925
         End
         Begin VB.Label lblBankBranch 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4830
            TabIndex        =   47
            Top             =   870
            Width           =   1485
         End
         Begin VB.Label lblCheckNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4590
            TabIndex        =   45
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label lblBankName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   46
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label lblPaymentType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            TabIndex        =   44
            Top             =   390
            Width           =   1395
         End
      End
      Begin VB.ComboBox cboEnpAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2250
         Width           =   9585
      End
      Begin VB.ComboBox cboCustomerAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   9585
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1350
         Width           =   2925
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1350
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   2
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   17
         Top             =   4440
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
         Left            =   2340
         TabIndex        =   1
         Top             =   900
         Width           =   2235
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture1 
            Height          =   495
            Left            =   0
            ScaleHeight     =   435
            ScaleWidth      =   495
            TabIndex        =   50
            Top             =   0
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   2670
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalDiscount 
         Height          =   435
         Left            =   5910
         TabIndex        =   9
         Top             =   2670
         Width           =   1335
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotal 
         Height          =   435
         Left            =   9210
         TabIndex        =   10
         Top             =   2670
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   11
         Top             =   3120
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDueDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   13
         Top             =   3570
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   5160
         TabIndex        =   53
         Top             =   7900
         Width           =   1335
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNote 
         Height          =   435
         Left            =   1860
         TabIndex        =   54
         Top             =   3960
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   767
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   55
         Top             =   4080
         Width           =   1485
      End
      Begin Threed.SSCommand cmdPrintOption 
         Height          =   525
         Left            =   11160
         TabIndex        =   52
         Top             =   3600
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUpdateAvgPrice 
         Height          =   525
         Left            =   6840
         TabIndex        =   51
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1845
         TabIndex        =   0
         Top             =   900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblDueDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   49
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblArea 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7740
         TabIndex        =   48
         Top             =   3210
         Width           =   1395
      End
      Begin Threed.SSCheck chkExtraFlag 
         Height          =   435
         Left            =   240
         TabIndex        =   16
         Top             =   390
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10500
         TabIndex        =   3
         Top             =   900
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
         TabIndex        =   38
         Top             =   3180
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   7890
         TabIndex        =   14
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   9510
         TabIndex        =   15
         Top             =   3600
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
         TabIndex        =   37
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   34
         Top             =   1410
         Width           =   1635
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   10860
         TabIndex        =   33
         Top             =   2730
         Width           =   585
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8190
         TabIndex        =   32
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label lblTotalDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4080
         TabIndex        =   31
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   7350
         TabIndex        =   30
         Top             =   2730
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3540
         TabIndex        =   29
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   28
         Top             =   960
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   22
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   23
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   21
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   26
         Top             =   2790
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   25
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_DateHasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_Customers As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public BATCH_ID As Long

Private FileName As String
Private m_SumUnit As Double
Public DocumentSubType As Long
Public DocumentType As Long

Private m_Cd As Collection
Private DocAdd As Long

Private Sub ShowButton(Ind As Long)
   If ShowMode = SHOW_ADD Then
      Exit Sub
   End If
   
   If Ind = 1 Then
      cmdAdd.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      cmdEdit.Enabled = True
   ElseIf Ind = 2 Then
      cmdAdd.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      cmdEdit.Enabled = True
   ElseIf Ind = 3 Then
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False

      cmdDelete.Enabled = False
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      m_BillingDoc.BILLING_DOC_ID = ID
      m_BillingDoc.BATCH_ID = BATCH_ID
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_BillingDoc.DOCUMENT_DATE
      txtDocumentNo.Text = m_BillingDoc.DOCUMENT_NO
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.CUSTOMER_ID)
      cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.ACCOUNT_ID)
      cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_BillingDoc.BILLING_ADDRESS_ID)
      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)
      txtTotalAmount.Text = Format(m_BillingDoc.TOTAL_AMOUNT, "0.00")
      txtTotalDiscount.Text = Format(m_BillingDoc.DISCOUNT_AMOUNT, "0.00")
      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
      cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, m_BillingDoc.PAYMENT_TYPE)
      txtCheckNo.Text = m_BillingDoc.CHECK_NO
      cboBank.ListIndex = IDToListIndex(cboBank, m_BillingDoc.BANK_ID)
      cboBankBranch.ListIndex = IDToListIndex(cboBankBranch, m_BillingDoc.BANK_BRANCH_ID)
      txtNote.Text = m_BillingDoc.NOTE
      cboArea.ListIndex = IDToListIndex(cboArea, m_BillingDoc.REGION_ID)
      uctlDueDate.ShowDate = m_BillingDoc.DUE_DATE
      
      chkCommit.Value = FlagToCheck(m_BillingDoc.COMMIT_FLAG)
      chkExtraFlag.Value = FlagToCheck(m_BillingDoc.EXCEPTION_FLAG)
      chkCommit.Enabled = (m_BillingDoc.COMMIT_FLAG = "N")
      Call ShowButton(1)
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

Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long) As CExportItem
Dim EI As CExportItem

      For Each EI In Ivd.ImportExports
         If EI.LINK_ID = GuiID Then
            Set GetExportItem = EI
            Exit Function
         End If
      Next EI
End Function

Private Function DO2InventoryDoc(Bd As CBillingDoc, Ivd As CInventoryDoc) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDoItem
Dim EI As CExportItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If Bd.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = Bd.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
   
   Ivd.DOCUMENT_DATE = Bd.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = Bd.DOCUMENT_NO
   Ivd.COMMIT_FLAG = Bd.COMMIT_FLAG
   Ivd.EXCEPTION_FLAG = Bd.EXCEPTION_FLAG
   Ivd.CUS_ID = Bd.CUSTOMER_ID
   Ivd.DOCUMENT_TYPE = 10
   Ivd.DOCUMENT_SUBTYPE = Bd.DOCUMENT_SUBTYPE
   
   If Bd.DOCUMENT_SUBTYPE = 1 Then 'หมู
      Ivd.SALE_FLAG = "N"
   ElseIf Bd.DOCUMENT_SUBTYPE = 2 Then 'วัตถุดิบ
      Ivd.SALE_FLAG = "Y"
   End If
   
   For Each Di In Bd.DoItems
      If Di.Flag = "A" Then
         Set EI = New CExportItem
         
         EI.TX_TYPE = "E"
         EI.Flag = "A"
         EI.PART_ITEM_ID = Di.PART_ITEM_ID
         EI.PIG_STATUS = Di.PIG_STATUS
         EI.LOCATION_ID = Di.LOCATION_ID
         EI.EXPORT_AMOUNT = Di.ITEM_AMOUNT
         EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         EI.TOTAL_PRICE = Di.TOTAL_PRICE
         EI.DISCOUNT_AMOUNT = Di.DISCOUNT_AMOUNT
         EI.LINK_ID = Di.LINK_ID
         EI.CALCULATE_FLAG = "N"
         EI.PIG_AGE = GetAge(Di.PART_NO, Bd.DOCUMENT_DATE)
         EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
         
         Di.PIG_AGE = EI.PIG_AGE
         Di.AGE_CODE = EI.AGE_CODE
                 
         Call Ivd.ImportExports.Add(EI)
         Set EI = Nothing
      ElseIf Di.Flag = "E" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         
         EI.Flag = "E"
         EI.PART_ITEM_ID = Di.PART_ITEM_ID
         EI.PIG_STATUS = Di.PIG_STATUS
         EI.LOCATION_ID = Di.LOCATION_ID
         EI.EXPORT_AMOUNT = Di.ITEM_AMOUNT
         EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         EI.TOTAL_PRICE = Di.TOTAL_PRICE
         EI.DISCOUNT_AMOUNT = Di.DISCOUNT_AMOUNT
         EI.CALCULATE_FLAG = "N"
         EI.PIG_AGE = GetAge(Di.PART_NO, Bd.DOCUMENT_DATE)
         EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
         
         Di.PIG_AGE = EI.PIG_AGE
         Di.AGE_CODE = EI.AGE_CODE
      ElseIf Di.Flag = "D" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         EI.Flag = "D"
      End If
   Next Di
End Function

Private Sub UpdatePigAge()
Dim Di As CDoItem
Dim OldPigAge As Long

   For Each Di In m_BillingDoc.DoItems
      If (Di.Flag <> "A") And (Di.Flag <> "D") Then
         Di.Flag = "E"
      End If
   Next Di
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment


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
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalAmount, txtTotalAmount, True) Then
      Exit Function
   End If
   If Not VerifyDate(lblDueDate, uctlDueDate, False) Then
      Exit Function
   End If
   If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
      uctlDocumentDate.SetFocus
      Exit Function
   End If
   
   If DocumentType = 3 Then
      If Not VerifyCombo(lblPaymentType, cboPaymentType, False) Then
         Exit Function
      End If
   End If
   
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_BillingDoc.AddEditMode = ShowMode
   m_BillingDoc.BILLING_DOC_ID = ID
    m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
    m_BillingDoc.DUE_DATE = uctlDueDate.ShowDate
   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_BillingDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_BillingDoc.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.DOCUMENT_TYPE = DocumentType
   m_BillingDoc.DOCUMENT_SUBTYPE = DocumentSubType
   m_BillingDoc.EXCEPTION_FLAG = Check2Flag(chkExtraFlag.Value)
   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.PAYMENT_TYPE = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   m_BillingDoc.CHECK_NO = txtCheckNo.Text
   m_BillingDoc.NOTE = txtNote.Text
   m_BillingDoc.REGION_ID = cboArea.ItemData(Minus2Zero(cboArea.ListIndex))
   m_BillingDoc.TOTAL_AMOUNT = Val(txtNetTotal.Text)
   'txtNetTotal
   If cboBankBranch.ListIndex > 0 Then
      m_BillingDoc.BANK_BRANCH_ID = cboBankBranch.ItemData(Minus2Zero(cboBankBranch.ListIndex))
   Else
      m_BillingDoc.BANK_BRANCH_ID = -1
   End If
   Call PopulateGuiID(m_BillingDoc)
   
   Call EnableForm(Me, False)
   
   If m_DateHasModify Then
      Call UpdatePigAge
   End If

   Call DO2InventoryDoc(m_BillingDoc, Ivd)

   If DocumentType = 3 Then
'      Call glbDaily.DO2Payment(m_BillingDoc, Pm)
   End If
   
   If (m_BillingDoc.COMMIT_FLAG = "Y") Then
      If m_BillingDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(Ivd.ImportExports)
         If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
            m_BillingDoc.COMMIT_FLAG = "N"
            Call EnableForm(Me, True)
            Exit Function
         End If
         
         If Not glbDaily.VerifyStockBalanceEx(Ivd.ImportExports, glbErrorLog) Then
            m_BillingDoc.COMMIT_FLAG = "N"
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
   End If
   
   Call glbDaily.StartTransaction
   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If DocumentType = 3 Then
'      If Not glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         SaveData = False
'         Call glbDaily.RollbackTransaction
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
'      m_BillingDoc.PAYMENT_ID = Pm.PAYMENT_ID
   End If
   
   m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
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
   End If
End Sub

Private Sub cboArea_Click()
   m_HasModify = True
End Sub

Private Sub cboBank_Click()
Dim BankID As Long

   BankID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
   If BankID > 0 Then
      Call LoadBankBranch(cboBankBranch, , BankID)
   End If
   m_HasModify = True
End Sub

Private Sub cboBank_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboBankBranch_Click()
   m_HasModify = True
End Sub

Private Sub cboBankBranch_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboPaymentType_Click()
Dim PaymentTypeID As PAYMENT_TYPE

   m_HasModify = True
   
   PaymentTypeID = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   If PaymentTypeID > 0 Then
      If PaymentTypeID = CASH_PMT Then
         txtCheckNo.Enabled = False
         cboBank.Enabled = False
         cboBankBranch.Enabled = False
      ElseIf PaymentTypeID = CHECK_PMT Then
         txtCheckNo.Enabled = True
         cboBank.Enabled = True
         cboBankBranch.Enabled = True
      End If
   End If
End Sub

Private Sub cboPaymentType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkExtraFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentSubType = 1 Then
        
        If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
            Exit Sub
        End If

         Set frmAddEditDoItem.ParentForm = Me
         frmAddEditDoItem.ExtraFlag = Check2Flag(chkExtraFlag.Value)
         frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         frmAddEditDoItem.CusId = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItem.ParentShowMode = ShowMode
         frmAddEditDoItem.ShowMode = SHOW_ADD
         frmAddEditDoItem.HeaderText = MapText("เพิ่มรายการใบส่งสินค้า")
         Load frmAddEditDoItem
         frmAddEditDoItem.Show 1
   
         OKClick = frmAddEditDoItem.OKClick
   
         Unload frmAddEditDoItem
         Set frmAddEditDoItem = Nothing
      ElseIf DocumentSubType = 2 Then
         frmAddEditDoItem2.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditDoItem2.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItem2.ParentShowMode = ShowMode
         frmAddEditDoItem2.ShowMode = SHOW_ADD
         frmAddEditDoItem2.HeaderText = MapText("เพิ่มรายการใบส่งสินค้า")
         Load frmAddEditDoItem2
         frmAddEditDoItem2.Show 1
   
         OKClick = frmAddEditDoItem2.OKClick
   
         Unload frmAddEditDoItem2
         Set frmAddEditDoItem2 = Nothing
      End If
      
      If OKClick Then
         Call GetTotalPrice

         GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditRevenueItem.ExtraFlag = Check2Flag(chkExtraFlag.Value)
      frmAddEditRevenueItem.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditRevenueItem.TempCollection = m_BillingDoc.Revenues
      frmAddEditRevenueItem.ParentShowMode = ShowMode
      frmAddEditRevenueItem.ShowMode = SHOW_ADD
      frmAddEditRevenueItem.HeaderText = MapText("เพิ่มรายการรายรับอื่น ๆ")
      Load frmAddEditRevenueItem
      frmAddEditRevenueItem.Show 1

      OKClick = frmAddEditRevenueItem.OKClick

      Unload frmAddEditRevenueItem
      Set frmAddEditRevenueItem = Nothing
   
      If OKClick Then
         Call GetTotalPrice
         
         GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
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
         m_BillingDoc.DoItems.Remove (ID2)
      Else
         m_BillingDoc.DoItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_BillingDoc.Revenues.Remove (ID2)
      Else
         m_BillingDoc.Revenues.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Public Sub ShowDoItemGrid()
   GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
   GridEX1.Rebind
   
   m_HasModify = True
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentSubType = 1 Then
        If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
            Exit Sub
        End If

         Set frmAddEditDoItem.ParentForm = Me
         frmAddEditDoItem.ExtraFlag = Check2Flag(chkExtraFlag.Value)
         frmAddEditDoItem.ID = ID
         frmAddEditDoItem.CusId = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItem.HeaderText = MapText("แก้ไขรายการใบส่งสินค้า")
         frmAddEditDoItem.ParentShowMode = ShowMode
         frmAddEditDoItem.ShowMode = SHOW_EDIT
         Load frmAddEditDoItem
         frmAddEditDoItem.Show 1
   
         OKClick = frmAddEditDoItem.OKClick
   
         Unload frmAddEditDoItem
         Set frmAddEditDoItem = Nothing
      ElseIf DocumentSubType = 2 Then
         frmAddEditDoItem2.ID = ID
         frmAddEditDoItem2.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditDoItem2.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItem2.HeaderText = MapText("แก้ไขรายการใบส่งสินค้า")
         frmAddEditDoItem2.ParentShowMode = ShowMode
         frmAddEditDoItem2.ShowMode = SHOW_EDIT
         Load frmAddEditDoItem2
         frmAddEditDoItem2.Show 1
   
         OKClick = frmAddEditDoItem2.OKClick
   
         Unload frmAddEditDoItem2
         Set frmAddEditDoItem2 = Nothing
      End If
      
      If OKClick Then
         Call GetTotalPrice
         GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditRevenueItem.ID = ID
      frmAddEditRevenueItem.ExtraFlag = Check2Flag(chkExtraFlag.Value)
      frmAddEditRevenueItem.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditRevenueItem.TempCollection = m_BillingDoc.Revenues
      frmAddEditRevenueItem.ParentShowMode = ShowMode
      frmAddEditRevenueItem.ShowMode = SHOW_EDIT
      frmAddEditRevenueItem.HeaderText = MapText("แก้ไขรายการรายรับอื่น ๆ")
      Load frmAddEditRevenueItem
      frmAddEditRevenueItem.Show 1

      OKClick = frmAddEditRevenueItem.OKClick

      Unload frmAddEditRevenueItem
      Set frmAddEditRevenueItem = Nothing
   
      If OKClick Then
         Call GetTotalPrice
         
         GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub CalculateIncludePrice()
Dim II As CImportItem
Dim AvgFee As Double

'   If m_SumUnit > 0 Then
'      AvgFee = Val(txtTotalAmount.Text) / m_SumUnit
'   Else
'      AvgFee = 0
'   End If
'
'   For Each II In m_BillingDoc.DoItems
'      If II.Flag <> "D" Then
'         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE + AvgFee
'      End If
'   Next II
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

Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long
Dim Report As CReportInterface
Dim ReportFlag As Boolean
Dim EditMode As SHOW_MODE_TYPE
Dim HeaderText As String

   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ReportFlag = False
   
   Set oMenu = New cPopupMenu
   If DocumentType = 1 Then
      lMenuChosen = oMenu.Popup("พิมพ์ใบส่งสินค้า", "ปรับค่าหน้ากระดาษ", "พิมพ์ใบส่งสินค้า(รวมตามประเภท)", "ปรับค่าหน้ากระดาษ", "พิมพ์ใบส่งสินค้าบนกระดาษเปล่า", "ปรับค่าหน้ากระดาษ", "พิมพ์ใบส่งสินค้า บนกระดาษเปล่า (แบบเต็ม)", "ปรับค่าหน้ากระดาษ", "พิมพ์ใบส่งสินค้า บนกระดาษเปล่า (2 ภาษา)", "ปรับค่าหน้ากระดาษ", "พิมพ์ใบส่งสินค้า(ขุน+สายพันธุ์)", "ปรับค่าหน้ากระดาษ")
   ElseIf DocumentType = 3 Then
      lMenuChosen = oMenu.Popup("พิมพ์ใบรับเงินชั่วคราว", "ปรับค่าหน้ากระดาษ")
   End If
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportNormalDO001"
      
      Set Report = New CReportNormalDO001
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(0, "OPTION_MODE")
      If DocumentType = 1 Then
         Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      ElseIf DocumentType = 3 Then
         Call Report.AddParam(MapText("ใบรับเงินชั่วคราว"), "REPORT_HEADER")
      End If
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")

      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportNormalDO001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      If DocumentType = 1 Then
         HeaderText = MapText("ใบส่งสินค้า")
      ElseIf DocumentType = 3 Then
         HeaderText = MapText("ใบรับเงินชั่วคราว")
      End If
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 3 Or lMenuChosen = 11 Then
      ReportKey = "CReportNormalDO002"
      
      Set Report = New CReportNormalDO002
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      
      If lMenuChosen = 3 Then
         Call Report.AddParam(0, "OPTION_MODE")
      ElseIf lMenuChosen = 11 Then
         Call Report.AddParam(1, "OPTION_MODE")
      End If
      
      If DocumentType = 1 Then
         Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      ElseIf DocumentType = 3 Then
         Call Report.AddParam(MapText("ใบรับเงินชั่วคราว"), "REPORT_HEADER")
      End If
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")

      ReportFlag = True
   ElseIf lMenuChosen = 4 Or lMenuChosen = 12 Then
      ReportKey = "CReportNormalDO002"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      If DocumentType = 1 Then
         HeaderText = MapText("ใบส่งสินค้า")
      ElseIf DocumentType = 3 Then
         HeaderText = MapText("ใบรับเงินชั่วคราว")
      End If
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportNormalDO003"
      
      Set Report = New CReportNormalDO003
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      If DocumentType = 1 Then
         Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      ElseIf DocumentType = 3 Then
         Call Report.AddParam(MapText("ใบรับเงินชั่วคราว"), "REPORT_HEADER")
      End If
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")

      ReportFlag = True
   ElseIf lMenuChosen = 6 Then
      ReportKey = "CReportNormalDO003"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      If DocumentType = 1 Then
         HeaderText = MapText("ใบส่งสินค้า")
     End If
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 7 Then
      ReportKey = "CReportNormalDO004"
      
      Set Report = New CReportNormalDO004
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      If DocumentType = 1 Then
         Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      ElseIf DocumentType = 3 Then
         Call Report.AddParam(MapText("ใบรับเงินชั่วคราว"), "REPORT_HEADER")
      End If
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")

      ReportFlag = True
   ElseIf lMenuChosen = 8 Then
      ReportKey = "CReportNormalDO004"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      If DocumentType = 1 Then
         HeaderText = MapText("ใบส่งสินค้า")
     End If
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 9 Then
      ReportKey = "CReportNormalDO006"
      
      Set Report = New CReportNormalDO006
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      
      Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")

      ReportFlag = True
   ElseIf lMenuChosen = 10 Then
      ReportKey = "CReportNormalDO006"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      If DocumentType = 1 Then
         HeaderText = MapText("ใบส่งสินค้า")
     End If
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
   End If
   
   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1
         
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
     
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub
Private Sub cmdPrintOption_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long
Dim Report As CReportInterface
Dim ReportFlag As Boolean
Dim EditMode As SHOW_MODE_TYPE
Dim HeaderText As String

   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If Not VerifyAccessRight("CAN-PRINT-OPTION", "สามารถพิมพ์บิลพิเศษได้", 2) Then
      frmVerifyAccRight.AccName = "CAN-PRINT-OPTION"
      frmVerifyAccRight.AccDesc = "สามารถพิมพ์บิลพิเศษได้"
      Load frmVerifyAccRight
      frmVerifyAccRight.Show 1

      If frmVerifyAccRight.GrantRight Then
         Unload frmVerifyAccRight
         Set frmVerifyAccRight = Nothing
      Else
         Unload frmVerifyAccRight
         Set frmVerifyAccRight = Nothing
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   ReportFlag = False
   
   Set oMenu = New cPopupMenu
   If DocumentType = 1 Then
      lMenuChosen = oMenu.Popup("พิมพ์ใบส่งสินค้า แบบที่ 1 (ดึงเฉลี่ย)", "-", "ปรับค่าหน้ากระดาษ")
   End If
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen <= 1 Then
      ReportKey = "CReportNormalDO001"
      
      Set Report = New CReportNormalDO001
      Call Report.AddParam(lMenuChosen, "OPTION_MODE")
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      If DocumentType = 1 Then
         Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      ElseIf DocumentType = 3 Then
         Call Report.AddParam(MapText("ใบรับเงินชั่วคราว"), "REPORT_HEADER")
      End If
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      
      ReportFlag = True
   ElseIf lMenuChosen = 3 Then
      ReportKey = "CReportNormalDO005"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      If DocumentType = 1 Then
         HeaderText = MapText("ใบส่งสินค้า")
      ElseIf DocumentType = 3 Then
         HeaderText = MapText("ใบรับเงินชั่วคราว")
      End If
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
    ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportNormalDO005"
      
      Set Report = New CReportNormalDO005
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      If DocumentType = 1 Then
         Call Report.AddParam(MapText("ใบส่งสินค้า"), "REPORT_HEADER")
      ElseIf DocumentType = 3 Then
         Call Report.AddParam(MapText("ใบรับเงินชั่วคราว"), "REPORT_HEADER")
      End If
      Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
      
      ReportFlag = True
   ElseIf lMenuChosen = 7 Then
      ReportKey = "CReportNormalDO005"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      If DocumentType = 1 Then
         HeaderText = MapText("ใบส่งสินค้า")
      ElseIf DocumentType = 3 Then
         HeaderText = MapText("ใบรับเงินชั่วคราว")
      End If
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If

   End If
   
   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1
         
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
     
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub

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

Private Sub cmdUpdateAvgPrice_Click()
   Set frmUpdatePigPrice.TempCollection = m_BillingDoc.DoItems
   frmUpdatePigPrice.HeaderText = MapText("Update ราคาสุกร")
   Load frmUpdatePigPrice
   frmUpdatePigPrice.Show 1
   
   OKClick = frmUpdatePigPrice.OKClick
   
   Unload frmUpdatePigPrice
   Set frmUpdatePigPrice = Nothing
   
   Call GetTotalPrice
   GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
   GridEX1.Rebind
   GridEX1.Visible = True
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees
      
      Call InitPaymentType(cboPaymentType)
      
      Call LoadBank(cboBank)
      Call LoadRegion(cboArea)
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_BillingDoc.QueryFlag = 0
         Call QueryData(False)
         uctlDocumentDate.ShowDate = Now
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long

   If InUsed = 1 Then
      Exit Sub
   End If
   
   InUsed = 1
   
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
   
   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   
   Set m_Cd = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 1305
   Col.Caption = MapText("สัปดาห์เกิด")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1230
   Col.Caption = MapText("ประเภทสุกร")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2175
   Col.Caption = MapText("สถานะสุกร")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 855
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 930
   Col.Caption = MapText("น้ำหนัก")
   
   Set Col = GridEX1.Columns.Add '9
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1290
   Col.Caption = MapText("ส่วนลด")
   
   Set Col = GridEX1.Columns.Add '8
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1575
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.Add '9
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1290
   Col.Caption = MapText("ราคา/หน่วย")

   Set Col = GridEX1.Columns.Add '10
   Col.Width = 2235
   Col.Caption = MapText("โรงเรือน")
End Sub

Private Sub InitGrid3()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 1650
   Col.Caption = MapText("รหัสรายได้อื่น ๆ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 5625
   Col.Caption = MapText("รายได้อื่น ๆ")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2175
   Col.Caption = MapText("จำนวน")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2115
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
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
   Col.Width = 2400
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1725
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3420
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 1770
   Col.Caption = MapText("จำนวน")
      
   Set Col = GridEX1.Columns.Add '7
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1890
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.Add '8
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1290
   Col.Caption = MapText("ราคา/หน่วย")

   Set Col = GridEX1.Columns.Add '9
   Col.Width = 2235
   Col.Caption = MapText("สถานที่จัดเก็บ")
End Sub

Private Sub GetTotalPrice()
Dim II As CDoItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.ITEM_AMOUNT
         Sum2 = Sum2 + II.TOTAL_PRICE
         Sum3 = Sum3 + II.TOTAL_WEIGHT
      End If
   Next II

   For Each II In m_BillingDoc.Revenues
      If II.Flag <> "D" Then
         Sum2 = Sum2 + II.TOTAL_PRICE
      End If
   Next II
   
   txtTotalDiscount.Text = Format(Sum3, "0.00")
   txtTotalAmount.Text = Format(Sum1, "0.00")
   txtNetTotal.Text = Format(Sum2, "0.00")
   txtAvgPrice.Text = MyDiff(Val(txtTotalDiscount.Text), Val(txtTotalAmount.Text))
      
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   GridEX1.Left = 150
   GridEX1.Top = 4920
   GridEX1.Visible = True
   GridEX1.itemcount = 0
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบส่งสินค้า"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
   Call InitNormalLabel(lblTotalAmount, MapText("จำนวนรวม"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่ส่งสินค้า"))
   Call InitNormalLabel(lblTotalDiscount, MapText("น้ำหนักรวม"))
   If DocumentSubType = 1 Then
      Label1.Visible = True
   Else
      Label1.Visible = False
   End If
   Call InitNormalLabel(Label1, MapText("ตัว"))
   Call InitNormalLabel(Label2, MapText("ก.ก."))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblNetTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
   Call InitNormalLabel(lblSellBy, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblPaymentType, MapText("การชำระเงิน"))
   Call InitNormalLabel(lblBankName, MapText("ธนาคาร"))
   Call InitNormalLabel(lblBankBranch, MapText("สาขา"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblCheckNo, MapText("เลขที่เช็ค"))
   Call InitNormalLabel(lblArea, MapText("เขตการขาย"))
   Call InitNormalLabel(lblDueDate, MapText("วันที่ครบดิว"))
   
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitCheckBox(chkExtraFlag, "ขายพิเศษ")
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   Call txtTotalDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalDiscount.Enabled = False
   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetTotal.Enabled = False
   Call txtCheckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitCombo(cboAccount)
   Call InitCombo(cboCustomerAddress)
   Call InitCombo(cboEnpAddress)
   Call InitCombo(cboPaymentType)
   Call InitCombo(cboBank)
   Call InitCombo(cboBankBranch)
   Call InitCombo(cboArea)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdPrintOption.Enabled = True
   cmdPrintOption.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdUpdateAvgPrice.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdPrintOption, MapText("P"))
   
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdUpdateAvgPrice, MapText("Updateราคา"))
   
   If DocumentSubType = 1 Then
      Call InitGrid1
   ElseIf DocumentSubType = 2 Then
      Call InitGrid2
   End If
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("รายการใบส่งสินค้า")
   TabStrip1.Tabs.Add().Caption = MapText("รายการรายรับอื่น ๆ")
   If DocumentType = 3 Then
      TabStrip1.Tabs.Add().Caption = MapText("การชำระเงิน")
   End If

   txtAvgPrice.Enabled = False
'   Call LoadPictureFromFile(glbParameterObj.DOFormPic1, Picture1)
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
   m_DateHasModify = False

   Set m_Rs = New ADODB.Recordset
   Set m_BillingDoc = New CBillingDoc
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   
   Set m_Cd = New Collection
   
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
      If m_BillingDoc.DoItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CDoItem
      If m_BillingDoc.DoItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.DoItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      If DocumentSubType = 1 Then
         Values(1) = CR.DO_ITEM_ID
         Values(2) = RealIndex
         Values(3) = CR.PART_NO
         Values(4) = CR.PIG_TYPE
         Values(5) = CR.PIG_STATUS_NAME
         Values(6) = FormatNumber(CR.ITEM_AMOUNT)
         Values(7) = FormatNumber(CR.TOTAL_WEIGHT)
         Values(8) = FormatNumber(CR.DISCOUNT_AMOUNT)
         Values(9) = FormatNumber(CR.TOTAL_PRICE)
         Values(10) = FormatNumber(CR.AVG_PRICE)
         Values(11) = CR.LOCATION_NAME
      ElseIf DocumentSubType = 2 Then
         Values(1) = CR.DO_ITEM_ID
         Values(2) = RealIndex
         Values(3) = CR.PART_TYPE_NAME
         Values(4) = CR.PART_NO
         Values(5) = CR.PART_DESC
         Values(6) = FormatNumber(CR.ITEM_AMOUNT)
         Values(7) = FormatNumber(CR.TOTAL_PRICE)
         Values(8) = FormatNumber(CR.AVG_PRICE)
         Values(9) = CR.LOCATION_NAME
         
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_BillingDoc.Revenues Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Rv As CDoItem
      If m_BillingDoc.Revenues.Count <= 0 Then
         Exit Sub
      End If
      Set Rv = GetItem(m_BillingDoc.Revenues, RowIndex, RealIndex)
      If Rv Is Nothing Then
         Exit Sub
      End If

      Values(1) = Rv.DO_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Rv.REVENUE_NO
      Values(4) = Rv.REVENUE_NAME
      Values(5) = FormatNumber(Rv.ITEM_AMOUNT)
      Values(6) = FormatNumber(Rv.TOTAL_PRICE)
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()
   GridEX1.Left = 150
   GridEX1.Top = 4920
   SSFrame2.Left = 150
   SSFrame2.Top = 4920
   
   GridEX1.Visible = False
   SSFrame2.Visible = False
   If TabStrip1.SelectedItem.Index = 1 Then
      If DocumentSubType = 1 Then
         Call InitGrid1
      ElseIf DocumentSubType = 2 Then
         Call InitGrid2
      End If
      
      Call GetTotalPrice
      GridEX1.itemcount = CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
      GridEX1.Visible = True
      
      Call ShowButton(TabStrip1.SelectedItem.Index)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid3
      GridEX1.itemcount = CountItem(m_BillingDoc.Revenues)
      GridEX1.Rebind
      GridEX1.Visible = True
      
      Call ShowButton(TabStrip1.SelectedItem.Index)
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      SSFrame2.Visible = True
      Call ShowButton(TabStrip1.SelectedItem.Index)
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtCheckNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_LostFocus()
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Sub
   End If
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalAmount_Change()
   m_HasModify = True
   txtNetTotal.Text = Format(Val(txtTotalAmount.Text) + Val(txtTotalDiscount.Text), "0.00")
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

Private Sub txtTotalDiscount_Change()
   m_HasModify = True
   txtNetTotal.Text = Format(Val(txtTotalAmount.Text) + Val(txtTotalDiscount.Text), "0.00")
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
   m_DateHasModify = True
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
      cboArea.ListIndex = IDToListIndex(cboArea, C.REGION_ID)
      uctlDueDate.ShowDate = DateAdd("D", C.Credit, uctlDocumentDate.ShowDate)
   Else
      cboAccount.ListIndex = -1
      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_LostFocus()
   If ShowMode = SHOW_ADD And uctlDocumentDate.ShowDate > 0 Then
      If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
         uctlDocumentDate.SetFocus
         Exit Sub
      End If
   ElseIf Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      txtDocumentNo.SetFocus
      Exit Sub
   ElseIf Not (uctlDocumentDate.ShowDate > 0) Then
      uctlDocumentDate.SetFocus
      Exit Sub
   End If
End Sub

Private Sub uctlDueDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub
Private Sub cmdAuto_Click()
Dim ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
   
   ID = ConvertDocToConfigNo(1, DocumentType, DocumentSubType)
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         Dim TempCd As CConfigDoc
         If ShowMode = SHOW_ADD Then
            If Cd.GetFieldValue("UPDATE_MONTH_FLAG") = "Y" Then
               If Not (Right(Trim(Str(Cd.GetFieldValue("MM"))), 2) = Format(Month(uctlDocumentDate.ShowDate), "00")) Then
                  Set TempCd = New CConfigDoc
                  Call TempCd.SetFieldValue("RUNNING_NO", 0)
                  Call TempCd.SetFieldValue("MM", Right(Format(Year(Now), "00") & Format(Month(uctlDocumentDate.ShowDate), "00"), 4))
                  Call TempCd.SetFieldValue("CONFIG_DOC_TYPE", ID)
                  Call TempCd.UpdateYearMonthRunningNo
                  Set Cd = Nothing
                  Set m_Cd = Nothing
                  Set m_Cd = New Collection
                  Call LoadConfigDoc(Nothing, m_Cd)
                  Call cmdAuto_Click
                  Exit Sub
               End If
            ElseIf Cd.GetFieldValue("UPDATE_YEAR_FLAG") = "Y" Then
               If Not (Left(Cd.GetFieldValue("MM"), 2) = Right(Format(Year(uctlDocumentDate.ShowDate), "00"), 2)) Then
                  Set TempCd = New CConfigDoc
                  Call TempCd.SetFieldValue("RUNNING_NO", 0)
                  Call TempCd.SetFieldValue("MM", Right(Format(Year(uctlDocumentDate.ShowDate), "00") & Format(Month(Now), "00"), 4))
                  Call TempCd.SetFieldValue("CONFIG_DOC_TYPE", ID)
                  Call TempCd.UpdateYearMonthRunningNo
                  Set Cd = Nothing
                  Set m_Cd = Nothing
                  Set m_Cd = New Collection
                  Call LoadConfigDoc(Nothing, m_Cd)
                  Call cmdAuto_Click
                  Exit Sub
               End If
            End If
            Set TempCd = Nothing
         End If
         
         txtDocumentNo.Text = Cd.GetFieldValue("PREFIX") & Cd.GetFieldValue("CODE1")
         TempStr = ""
         If Cd.GetFieldValue("YEAR_TYPE") = 1 Then
            TempStr = Right(Format(Year(Now) + 543, "0000"), 2)
         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 2 Then
            TempStr = Format(Year(Now) + 543, "0000")
         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 3 Then
            TempStr = Right(Format(Year(Now), "0000"), 2)
         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 4 Then
            TempStr = Format(Year(Now), "0000")
         End If
         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE2")
         TempStr = ""
         If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
            TempStr = Format(Month(Now), "00")
         End If
         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE3")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         txtDocumentNo.Text = txtDocumentNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
         m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
         m_BillingDoc.CONFIG_DOC_TYPE = ID
      Else
         txtDocumentNo.Text = ""
      End If
   txtDocumentNo.SetFocus
   End If
End Sub

