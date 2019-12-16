VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditInventoryDoc1 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditInventoryDoc1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlSupplierLookup 
         Height          =   435
         Left            =   6300
         TabIndex        =   4
         Top             =   1500
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6300
         TabIndex        =   2
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   15
         Top             =   4170
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
      Begin prjFarmManagement.uctlTextBox txtDoNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1530
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   2460
         TabIndex        =   1
         Top             =   1080
         Width           =   2235
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDeliveryNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   1980
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   7
         Top             =   2430
         Width           =   2835
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   3015
         Left            =   150
         TabIndex        =   16
         Top             =   4710
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5318
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
         Column(1)       =   "frmAddEditInventoryDoc1.frx":27A2
         Column(2)       =   "frmAddEditInventoryDoc1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDoc1.frx":290E
         FormatStyle(2)  =   "frmAddEditInventoryDoc1.frx":2A6A
         FormatStyle(3)  =   "frmAddEditInventoryDoc1.frx":2B1A
         FormatStyle(4)  =   "frmAddEditInventoryDoc1.frx":2BCE
         FormatStyle(5)  =   "frmAddEditInventoryDoc1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDoc1.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtSender 
         Height          =   435
         Left            =   1860
         TabIndex        =   10
         Top             =   2880
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtReceiver 
         Height          =   435
         Left            =   6300
         TabIndex        =   11
         Top             =   2880
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDeliveryFee 
         Height          =   435
         Left            =   1860
         TabIndex        =   12
         Top             =   3330
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtMaterialPrice 
         Height          =   435
         Left            =   6300
         TabIndex        =   13
         Top             =   3330
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotal 
         Height          =   435
         Left            =   9810
         TabIndex        =   14
         Top             =   3330
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlDeliveryLookup 
         Height          =   435
         Left            =   6300
         TabIndex        =   6
         Top             =   1950
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPoNo 
         Height          =   435
         Left            =   6300
         TabIndex        =   8
         Top             =   2400
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin VB.Label lblPoNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4710
         TabIndex        =   41
         Top             =   2460
         Width           =   1485
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   5160
         TabIndex        =   40
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   6840
         TabIndex        =   20
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   9600
         TabIndex        =   9
         Top             =   2400
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblDeliveryCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4770
         TabIndex        =   39
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lblSupplierNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4770
         TabIndex        =   38
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   37
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8580
         TabIndex        =   36
         Top             =   3450
         Width           =   1125
      End
      Begin VB.Label lblMaterialPrice 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   35
         Top             =   3450
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   7740
         TabIndex        =   34
         Top             =   3420
         Width           =   765
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3300
         TabIndex        =   33
         Top             =   3420
         Width           =   765
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   32
         Top             =   1110
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   21
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   22
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   19
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDoc1.frx":3EB8
         ButtonStyle     =   3
      End
      Begin VB.Label lblDeliveryFee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   30
         Top             =   3450
         Width           =   1695
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblDeliveryNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   28
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   27
         Top             =   1140
         Width           =   1665
      End
      Begin VB.Label lblReceiver 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4710
         TabIndex        =   26
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label lblSender 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   25
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label lblDoNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   1620
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDoc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryDoc As CInventoryDoc
Private m_Suppliers As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public BATCH_ID As Long

Private FileName As String
Private m_SumUnit As Double
Private m_Cd As Collection
Private DocAdd As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_InventoryDoc.BATCH_ID = BATCH_ID
      m_InventoryDoc.INVENTORY_DOC_ID = ID
      m_InventoryDoc.COMMIT_FLAG = ""
      If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_InventoryDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_InventoryDoc.DOCUMENT_DATE
      txtDoNo.Text = m_InventoryDoc.DO_NO
      txtTruckNo.Text = m_InventoryDoc.TRUCK_NO
      txtDocumentNo.Text = m_InventoryDoc.DOCUMENT_NO
      txtDeliveryFee.Text = Format(m_InventoryDoc.DELIVERY_FEE, "0.00")
      txtDeliveryNo.Text = m_InventoryDoc.BILL_NO
      txtSender.Text = m_InventoryDoc.SENDER_NAME
      txtReceiver.Text = m_InventoryDoc.RECEIVE_NAME
      uctlSupplierLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierLookup.MyCombo, m_InventoryDoc.SUPPLIER_ID)
      uctlDeliveryLookup.MyCombo.ListIndex = IDToListIndex(uctlDeliveryLookup.MyCombo, m_InventoryDoc.DELIVERY_ID)
      chkCommit.Value = FlagToCheck(m_InventoryDoc.COMMIT_FLAG)
      txtPoNo.Text = m_InventoryDoc.PO_NO
      
      cmdAdd.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      cmdDelete.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      chkCommit.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
      txtDeliveryFee.Enabled = (m_InventoryDoc.COMMIT_FLAG = "N")
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

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("INVENTORY_IMPORT_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("INVENTORY_IMPORT_EDIT") Then
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
   If Not VerifyCombo(lblSupplierNo, uctlSupplierLookup.MyCombo, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDeliveryFee, txtDeliveryFee, True) Then
      Exit Function
   End If
   If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
      uctlDocumentDate.SetFocus
      Exit Function
   End If
   
   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_InventoryDoc.AddEditMode = ShowMode
   m_InventoryDoc.INVENTORY_DOC_ID = ID
    m_InventoryDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_InventoryDoc.DO_NO = txtDoNo.Text
   m_InventoryDoc.TRUCK_NO = txtTruckNo.Text
   m_InventoryDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_InventoryDoc.DELIVERY_FEE = Val(txtDeliveryFee.Text)
   m_InventoryDoc.BILL_NO = txtDeliveryNo.Text
   m_InventoryDoc.SENDER_NAME = txtSender.Text
   m_InventoryDoc.RECEIVE_NAME = txtReceiver.Text
   m_InventoryDoc.SUPPLIER_ID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   m_InventoryDoc.DELIVERY_ID = uctlDeliveryLookup.MyCombo.ItemData(Minus2Zero(uctlDeliveryLookup.MyCombo.ListIndex))
   m_InventoryDoc.DOCUMENT_TYPE = 1
   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_InventoryDoc.EXCEPTION_FLAG = "N"
   m_InventoryDoc.PO_NO = txtPoNo.Text
   
   If m_InventoryDoc.COMMIT_FLAG = "Y" Then
      If m_InventoryDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(m_InventoryDoc.ImportExports)
      End If
   End If
   Call CalculateIncludePrice(m_InventoryDoc, m_SumUnit, Val(txtDeliveryFee.Text))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
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
Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditImportItem.DefaultLocationID = glbDaily.GetPreviousLocationID(m_InventoryDoc.ImportExports, 1)
      frmAddEditImportItem.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditImportItem.TempCollection = m_InventoryDoc.ImportExports
      frmAddEditImportItem.ParentShowMode = ShowMode
      frmAddEditImportItem.ShowMode = SHOW_ADD
      frmAddEditImportItem.HeaderText = MapText("เพิ่มรายการนำเข้า")
      Load frmAddEditImportItem
      frmAddEditImportItem.Show 1
      
      OKClick = frmAddEditImportItem.OKClick

      Unload frmAddEditImportItem
      Set frmAddEditImportItem = Nothing

      If OKClick Then
         Call GetTotalPrice
         
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
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
Dim ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
   
   ID = ConvertDocToConfigNo(3, 1, 1)
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
         m_InventoryDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
         m_InventoryDoc.CONFIG_DOC_TYPE = ID
      Else
         txtDocumentNo.Text = ""
      End If
   txtDocumentNo.SetFocus
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
         m_InventoryDoc.ImportExports.Remove (ID2)
      Else
         m_InventoryDoc.ImportExports.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
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
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditImportItem.ID = ID
      frmAddEditImportItem.COMMIT_FLAG = m_InventoryDoc.COMMIT_FLAG
      Set frmAddEditImportItem.TempCollection = m_InventoryDoc.ImportExports
      frmAddEditImportItem.HeaderText = MapText("แก้ไขรายการนำเข้า")
      frmAddEditImportItem.ParentShowMode = ShowMode
      frmAddEditImportItem.ShowMode = SHOW_EDIT
      Load frmAddEditImportItem
      frmAddEditImportItem.Show 1

      OKClick = frmAddEditImportItem.OKClick

      Unload frmAddEditImportItem
      Set frmAddEditImportItem = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
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
Dim ClassName As String


   If m_HasModify Or (m_InventoryDoc.INVENTORY_DOC_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ReportFlag = False
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("พิมพ์ใบรายงานรับของ", "พิมพ์ใบรายงานรับของ (ฟอร์มเปล่า)", "พิมพ์ใบรายงานรับของ (ฟอร์มเปล่า-กำหนดเลขที่เอง)", "ปรับค่าหน้ากระดาษ")
   
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportNormalInv001"
      
      Set Report = New CReportNormalInv001
      ClassName = "CReportNormalInv001"
      Call Report.AddParam(m_InventoryDoc.INVENTORY_DOC_ID, "INVENTORY_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(1, "REPORT_TYPE")
      
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportNormalInv001"
      
      Set Report = New CReportNormalInv001
      Call Report.AddParam(m_InventoryDoc.INVENTORY_DOC_ID, "INVENTORY_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(2, "REPORT_TYPE")
      
      ReportFlag = True
   ElseIf lMenuChosen = 3 Then
      Dim DocumentRunStart As String
      Dim DocumentRunNo As Double
      
      Load frmSetFormPrint
      frmSetFormPrint.Show 1
   
      Unload frmSetFormPrint
      DocumentRunStart = frmSetFormPrint.DocumentRunStart
      DocumentRunNo = frmSetFormPrint.DocumentRunNo
      
      Set frmSetFormPrint = Nothing
      
      If Len(DocumentRunStart) <= 0 Or DocumentRunNo < 0 Then
         glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการใส่ข้อมูลฟอร์มพิมพ์ให้ครบ")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      
      ReportKey = "CReportNormalInv002"
      
      Set Report = New CReportNormalInv002
      Call Report.AddParam(m_InventoryDoc.INVENTORY_DOC_ID, "INVENTORY_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      
      Call Report.AddParam(DocumentRunStart, "REPORT_DOC_NO")
      Call Report.AddParam(DocumentRunNo, "REPORT_RUN_NO")
      
      ReportFlag = True
   ElseIf lMenuChosen = 4 Then
      ReportKey = "CReportNormalInv001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบรายงานรับของ")
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
   ID = m_InventoryDoc.INVENTORY_DOC_ID
   m_InventoryDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      Call LoadSupplier(uctlSupplierLookup.MyCombo, m_Suppliers)
      Call LoadSupplier(uctlDeliveryLookup.MyCombo, m_Suppliers)
      Set uctlSupplierLookup.MyCollection = m_Suppliers
      Set uctlDeliveryLookup.MyCollection = m_Suppliers
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_InventoryDoc.QueryFlag = 0
         Call QueryData(False)
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
      Call cmdSave_Click
      KeyCode = 0
   End If
   
   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_InventoryDoc = Nothing
   Set m_Suppliers = Nothing
   Set m_Cd = Nothing
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
   Col.Width = 2100
   Col.Caption = MapText("หมายเลขวัตถุดิบ")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 4425
   Col.Caption = MapText("วัตถุดิบ")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 1785
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ปริมาณ")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 1620
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1980
   Col.Caption = MapText("สถานที่จัดเก็บ")
End Sub

Private Sub GetTotalPrice()
Dim II As CImportItem
Dim Sum As Double

   Sum = 0
   m_SumUnit = 0
   For Each II In m_InventoryDoc.ImportExports
      If II.Flag <> "D" Then
         Sum = Sum + CDbl(Format(II.TOTAL_ACTUAL_PRICE, "0.00"))
         m_SumUnit = m_SumUnit + II.IMPORT_AMOUNT
      End If
   Next II
   
   txtMaterialPrice.Text = Format(Sum, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblTruckNo, MapText("ทะเบียนรถ"))
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่บิลรับของ"))
   Call InitNormalLabel(lblReceiver, MapText("ชื่อผู้รับ"))
   Call InitNormalLabel(lblDoNo, MapText("เลขที่ใบส่งของ"))
   Call InitNormalLabel(lblDeliveryNo, MapText("เลขที่บิลค่าขนส่ง"))
   Call InitNormalLabel(lblSender, MapText("ชื่อผู้ส่ง"))
   Call InitNormalLabel(lblDeliveryFee, MapText("ค่าขนส่ง"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblMaterialPrice, MapText("ราคาวัตถุดิบ"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblSupplierNo, MapText("รหัสซัพ ฯ"))
   Call InitNormalLabel(lblDeliveryCode, MapText("รหัสผู้ขนส่ง"))
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitNormalLabel(lblPoNo, MapText("เลขที่ใบสั่งซื้อ"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtDeliveryNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDoNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTruckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDeliveryFee.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtMaterialPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtMaterialPrice.Enabled = False
   Call txtReceiver.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotal.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก (F10)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("รายการรับวัตถุดิบ")
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
   Set m_InventoryDoc = New CInventoryDoc
   Set m_Suppliers = New Collection
   Set m_Cd = New Collection
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
      If m_InventoryDoc.ImportExports Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CImportItem
      If m_InventoryDoc.ImportExports.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_InventoryDoc.ImportExports, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.IMPORT_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PART_NO
      If Len(CR.IMPORT_DESC) > 0 Then
         Values(4) = CR.PART_DESC & " (" & CR.IMPORT_DESC & ")"
      Else
         Values(4) = CR.PART_DESC
      End If
      Values(5) = FormatNumber(CR.IMPORT_AMOUNT)
      Values(6) = FormatNumber(CR.ACTUAL_UNIT_PRICE)
      Values(7) = FormatNumber(CR.TOTAL_ACTUAL_PRICE)
      Values(8) = CR.LOCATION_NAME
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

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryDoc.ImportExports)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtDeliveryFee_Change()
   m_HasModify = True
   txtTotal.Text = Format(Val(txtDeliveryFee.Text) + Val(txtMaterialPrice.Text), "0.00")
End Sub

Private Sub txtDocumentNo_LostFocus()
   
   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      
      'txtDocumentNo.SetFocus
      txtDocumentNo.Text = ""
      Exit Sub
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

Private Sub txtMaterialPrice_Change()
   m_HasModify = True
   txtTotal.Text = Format(Val(txtDeliveryFee.Text) + Val(txtMaterialPrice.Text), "0.00")
End Sub

Private Sub txtPoNo_Change()
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlDocumentDate_LostFocus()
   If ShowMode = SHOW_ADD And uctlDocumentDate.ShowDate > 0 Then
      If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
         uctlDocumentDate.SetFocus
         Exit Sub
      End If
   ElseIf Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
      txtDocumentNo.SetFocus
      Exit Sub
   ElseIf Not (uctlDocumentDate.ShowDate > 0) Then
      uctlDocumentDate.SetFocus
      Exit Sub
   End If
End Sub
Private Sub uctlSupplierLookup_Change()
   m_HasModify = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   pnlHeader.Width = ScaleWidth
   
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = cmdAdd.Top
   cmdDelete.Top = cmdAdd.Top
   cmdOK.Top = cmdAdd.Top
   cmdExit.Top = cmdAdd.Top
   cmdSave.Top = cmdAdd.Top
   cmdPrint.Top = cmdAdd.Top
   
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdSave.Left = cmdOK.Left - cmdSave.Width - 50
   cmdPrint.Left = cmdSave.Left - cmdPrint.Width - 50
End Sub
