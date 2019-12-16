VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditParameter 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditParameter.frx":0000
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
      TabIndex        =   17
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlPigTypeLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   7
         Top             =   2430
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlJournalDate 
         Height          =   405
         Left            =   6270
         TabIndex        =   2
         Top             =   1050
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   10
         Top             =   3480
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
      Begin prjFarmManagement.uctlTextBox txtJournalCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtJournalDesc 
         Height          =   450
         Left            =   1860
         TabIndex        =   3
         Top             =   1500
         Width           =   8265
         _ExtentX        =   16907
         _ExtentY        =   794
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3705
         Left            =   150
         TabIndex        =   11
         Top             =   4020
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   6535
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
         Column(1)       =   "frmAddEditParameter.frx":27A2
         Column(2)       =   "frmAddEditParameter.frx":286A
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAddEditParameter.frx":290E
         FormatStyle(2)  =   "frmAddEditParameter.frx":29EE
         FormatStyle(3)  =   "frmAddEditParameter.frx":2B4A
         FormatStyle(4)  =   "frmAddEditParameter.frx":2BFA
         FormatStyle(5)  =   "frmAddEditParameter.frx":2CAE
         FormatStyle(6)  =   "frmAddEditParameter.frx":2D86
         ImageCount      =   0
         PrinterProperties=   "frmAddEditParameter.frx":2E3E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtFromAge 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   1980
         Width           =   1665
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtToAge 
         Height          =   435
         Left            =   6270
         TabIndex        =   6
         Top             =   1980
         Width           =   1665
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPigStatusLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   9
         Top             =   2850
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtADG 
         Height          =   435
         Left            =   8460
         TabIndex        =   8
         Top             =   2430
         Width           =   1665
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin VB.Label lblADG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7290
         TabIndex        =   28
         Top             =   2550
         Width           =   1065
      End
      Begin VB.Label lblPigStatus 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   27
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblPigType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   26
         Top             =   2490
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   8040
         TabIndex        =   25
         Top             =   2070
         Width           =   825
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3630
         TabIndex        =   24
         Top             =   2070
         Width           =   825
      End
      Begin VB.Label lblToAge 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5100
         TabIndex        =   23
         Top             =   2100
         Width           =   1065
      End
      Begin VB.Label lblFromAge 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   690
         TabIndex        =   22
         Top             =   2100
         Width           =   1065
      End
      Begin Threed.SSCheck chkPostFlag 
         Height          =   405
         Left            =   10170
         TabIndex        =   4
         Top             =   1500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJournalDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5070
         TabIndex        =   21
         Top             =   1050
         Width           =   1065
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4170
         TabIndex        =   1
         Top             =   1020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditParameter.frx":3016
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditParameter.frx":3330
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   16
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditParameter.frx":364A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditParameter.frx":3964
         ButtonStyle     =   3
      End
      Begin VB.Label lblJournalDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   19
         Top             =   1620
         Width           =   1695
      End
      Begin VB.Label lblJournalCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   18
         Top             =   1140
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAddEditParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Parameter As CParameter
Private m_ApArMass As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ParamArea As Long

Private ApArText As String
Private FileName As String
Private m_PigTypes As Collection
Private m_PigStatus As Collection

Public JournalType As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_Parameter.SetFieldValue("PARAM_ID", ID)
      If Not glbDaily.QueryParameter(m_Parameter, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Parameter.PopulateFromRS(1, m_Rs)
      txtJournalCode.Text = m_Parameter.GetFieldValue("PARAM_NO")
      txtJournalDesc.Text = m_Parameter.GetFieldValue("PARAM_DESC")
      chkPostFlag.Value = FlagToCheck(m_Parameter.GetFieldValue("COMMIT_FLAG"))
      uctlJournalDate.ShowDate = m_Parameter.GetFieldValue("PARAM_DATE")
      txtFromAge.Text = m_Parameter.GetFieldValue("FROM_AGE")
      txtToAge.Text = m_Parameter.GetFieldValue("TO_AGE")
      uctlPigTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPigTypeLookup.MyCombo, m_Parameter.GetFieldValue("PIG_TYPE"))
      uctlPigStatusLookup.MyCombo.ListIndex = IDToListIndex(uctlPigStatusLookup.MyCombo, m_Parameter.GetFieldValue("PIG_STATUS"))
      txtADG.Text = m_Parameter.GetFieldValue("ADG")
   Else
      ShowMode = SHOW_ADD
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
      If Not VerifyAccessRight("SIMULATE_PARAMETER_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("SIMULATE_PARAMETER_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblJournalCode, txtJournalCode, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblJournalDate, uctlJournalDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigType, uctlPigTypeLookup.MyCombo, Not uctlPigTypeLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPigStatus, uctlPigStatusLookup.MyCombo, Not uctlPigStatusLookup.Enabled) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblFromAge, txtFromAge, Not txtFromAge.Enabled) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblToAge, txtToAge, Not txtToAge.Enabled) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(CUSTCODE_UNIQUE, txtJournalCode.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJournalCode.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Parameter.ShowMode = ShowMode
   Call m_Parameter.SetFieldValue("PARAM_DATE", uctlJournalDate.ShowDate)
   Call m_Parameter.SetFieldValue("COMMIT_FLAG", Check2Flag(chkPostFlag.Value))
   Call m_Parameter.SetFieldValue("PARAM_NO", txtJournalCode.Text)
   Call m_Parameter.SetFieldValue("PARAM_DESC", txtJournalDesc.Text)
   Call m_Parameter.SetFieldValue("PARAM_AREA", ParamArea)
   Call m_Parameter.SetFieldValue("FROM_AGE", Val(txtFromAge.Text))
   Call m_Parameter.SetFieldValue("TO_AGE", Val(txtToAge.Text))
   Call m_Parameter.SetFieldValue("FROM_SALE_DATE", -1)
   Call m_Parameter.SetFieldValue("TO_SALE_DATE", -1)
   Call m_Parameter.SetFieldValue("PIG_TYPE", uctlPigTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPigTypeLookup.MyCombo.ListIndex)))
   Call m_Parameter.SetFieldValue("PIG_STATUS", uctlPigStatusLookup.MyCombo.ItemData(Minus2Zero(uctlPigStatusLookup.MyCombo.ListIndex)))
   Call m_Parameter.SetFieldValue("ADG", Val(txtADG.Text))

   Call EnableForm(Me, False)
   If Not glbDaily.AddEditParameter(m_Parameter, IsOK, True, glbErrorLog) Then
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

Private Sub chkPostFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkPostFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      If ParamArea = 1 Then
         Set frmAddEditBrthParamItem.ParentForm = Me
         Set frmAddEditBrthParamItem.TempCollection = m_Parameter.BrtPrmItems
         frmAddEditBrthParamItem.ShowMode = SHOW_ADD
         frmAddEditBrthParamItem.HeaderText = MapText("เพิ่มการผสม")
         Load frmAddEditBrthParamItem
         frmAddEditBrthParamItem.Show 1

         OKClick = frmAddEditBrthParamItem.OKClick

         Unload frmAddEditBrthParamItem
         Set frmAddEditBrthParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.BrtPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 2 Then
         Set frmAddEditFoodParamItem.ParentForm = Me
         Set frmAddEditFoodParamItem.TempCollection = m_Parameter.UsedPrmItems
         frmAddEditFoodParamItem.ShowMode = SHOW_ADD
         frmAddEditFoodParamItem.HeaderText = MapText("เพิ่มเบอร์อาหาร/ยา")
         Load frmAddEditFoodParamItem
         frmAddEditFoodParamItem.Show 1
   
         OKClick = frmAddEditFoodParamItem.OKClick
   
         Unload frmAddEditFoodParamItem
         Set frmAddEditFoodParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.UsedPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 3 Then
         Set frmAddEditLossParamItem.ParentForm = Me
         Set frmAddEditLossParamItem.TempCollection = m_Parameter.TrnPrmItems
         frmAddEditLossParamItem.ShowMode = SHOW_ADD
         frmAddEditLossParamItem.HeaderText = MapText("เพิ่มการโอนตามสถานะ")
         Load frmAddEditLossParamItem
         frmAddEditLossParamItem.Show 1
   
         OKClick = frmAddEditLossParamItem.OKClick
   
         Unload frmAddEditLossParamItem
         Set frmAddEditLossParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.TrnPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 4 Then
         Set frmAddEditSaleParamItem.ParentForm = Me
         Set frmAddEditSaleParamItem.TempCollection = m_Parameter.SalePrmItems
         frmAddEditSaleParamItem.ShowMode = SHOW_ADD
         frmAddEditSaleParamItem.HeaderText = MapText("เพิ่มราคาขายตามสถานะ")
         Load frmAddEditSaleParamItem
         frmAddEditSaleParamItem.Show 1
   
         OKClick = frmAddEditSaleParamItem.OKClick
   
         Unload frmAddEditSaleParamItem
         Set frmAddEditSaleParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.SalePrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 5 Then
         Set frmAddEditWeightParamItem.ParentForm = Me
         Set frmAddEditWeightParamItem.TempCollection = m_Parameter.WeightPrmItems
         frmAddEditWeightParamItem.ShowMode = SHOW_ADD
         frmAddEditWeightParamItem.HeaderText = MapText("เพิ่มน้ำหนักตามช่วงอายุ")
         Load frmAddEditWeightParamItem
         frmAddEditWeightParamItem.Show 1
   
         OKClick = frmAddEditWeightParamItem.OKClick
   
         Unload frmAddEditWeightParamItem
         Set frmAddEditWeightParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.WeightPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 6 Then
         Set frmAddEditCostParamItem.ParentForm = Me
         Set frmAddEditCostParamItem.TempCollection = m_Parameter.CostPrmItems
         frmAddEditCostParamItem.ShowMode = SHOW_ADD
         frmAddEditCostParamItem.HeaderText = MapText("เพิ่มราคาอาหาร/ยา")
         Load frmAddEditCostParamItem
         frmAddEditCostParamItem.Show 1
   
         OKClick = frmAddEditCostParamItem.OKClick
   
         Unload frmAddEditCostParamItem
         Set frmAddEditCostParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.CostPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 7 Then
         Set frmAddEditAdjParamItem.ParentForm = Me
         Set frmAddEditAdjParamItem.TempCollection = m_Parameter.AdjPrmItems
         frmAddEditAdjParamItem.ShowMode = SHOW_ADD
         frmAddEditAdjParamItem.HeaderText = MapText("เพิ่มยอดหมูยกมา")
         Load frmAddEditAdjParamItem
         frmAddEditAdjParamItem.Show 1

         OKClick = frmAddEditAdjParamItem.OKClick

         Unload frmAddEditAdjParamItem
         Set frmAddEditAdjParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.AdjPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 9 Then
         Set frmAddEditRevenueParamltem.ParentForm = Me
         Set frmAddEditRevenueParamltem.TempCollection = m_Parameter.RvnPrmItems
         frmAddEditRevenueParamltem.ShowMode = SHOW_ADD
         frmAddEditRevenueParamltem.HeaderText = MapText("เพิ่มรายการขายอื่น ๆ")
         Load frmAddEditRevenueParamltem
         frmAddEditRevenueParamltem.Show 1

         OKClick = frmAddEditRevenueParamltem.OKClick

         Unload frmAddEditRevenueParamltem
         Set frmAddEditRevenueParamltem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.RvnPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 10 Then
         Set frmAddEditCustRatioParamlItem.ParentForm = Me
         Set frmAddEditCustRatioParamlItem.TempCollection = m_Parameter.CustRatioItems
         frmAddEditCustRatioParamlItem.ShowMode = SHOW_ADD
         frmAddEditCustRatioParamlItem.HeaderText = MapText("เพิ่ม % การขาย")
         Load frmAddEditCustRatioParamlItem
         frmAddEditCustRatioParamlItem.Show 1

         OKClick = frmAddEditCustRatioParamlItem.OKClick

         Unload frmAddEditCustRatioParamlItem
         Set frmAddEditCustRatioParamlItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.CustRatioItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 11 Then
         Set frmAddEditPigTypeChangeParamlItem.ParentForm = Me
         Set frmAddEditPigTypeChangeParamlItem.TempCollection = m_Parameter.PigStatusChangeItems
         frmAddEditPigTypeChangeParamlItem.ShowMode = SHOW_ADD
         frmAddEditPigTypeChangeParamlItem.HeaderText = MapText("เพิ่มการเปลี่ยนสถานะ")
         Load frmAddEditPigTypeChangeParamlItem
         frmAddEditPigTypeChangeParamlItem.Show 1

         OKClick = frmAddEditPigTypeChangeParamlItem.OKClick

         Unload frmAddEditPigTypeChangeParamlItem
         Set frmAddEditPigTypeChangeParamlItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.PigStatusChangeItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 12 Then
         Set frmAddEditPigBuyParamItem.ParentForm = Me
         Set frmAddEditPigBuyParamItem.TempCollection = m_Parameter.PigBuyItems
         frmAddEditPigBuyParamItem.ShowMode = SHOW_ADD
         frmAddEditPigBuyParamItem.HeaderText = MapText("เพิ่มการซื้อสุกร")
         Load frmAddEditPigBuyParamItem
         frmAddEditPigBuyParamItem.Show 1

         OKClick = frmAddEditPigBuyParamItem.OKClick

         Unload frmAddEditPigBuyParamItem
         Set frmAddEditPigBuyParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.PigBuyItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 13 Then
         Set frmAddEditExpenseParamItem.ParentForm = Me
         Set frmAddEditExpenseParamItem.TempCollection = m_Parameter.ExpenseSharings
         frmAddEditExpenseParamItem.ShowMode = SHOW_ADD
         frmAddEditExpenseParamItem.HeaderText = MapText("เพิ่มการปันค่าใช้จ่าย")
         Load frmAddEditExpenseParamItem
         frmAddEditExpenseParamItem.Show 1

         OKClick = frmAddEditExpenseParamItem.OKClick

         Unload frmAddEditExpenseParamItem
         Set frmAddEditExpenseParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.ExpenseSharings)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 14 Then
         Set frmAddEditPigAdjParamItem.ParentForm = Me
         Set frmAddEditPigAdjParamItem.TempCollection = m_Parameter.PigAdjustItems
         frmAddEditPigAdjParamItem.ShowMode = SHOW_ADD
         frmAddEditPigAdjParamItem.HeaderText = MapText("เพิ่มการคุมยอดสุกร")
         Load frmAddEditPigAdjParamItem
         frmAddEditPigAdjParamItem.Show 1

         OKClick = frmAddEditPigAdjParamItem.OKClick

         Unload frmAddEditPigAdjParamItem
         Set frmAddEditPigAdjParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.PigAdjustItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 15 Then
         Set frmAddEditExpenseManagenent.ParentForm = Me
         Set frmAddEditExpenseManagenent.TempCollection = m_Parameter.MenagementExpenses
         frmAddEditExpenseManagenent.ShowMode = SHOW_ADD
         frmAddEditExpenseManagenent.HeaderText = MapText("เพิ่มยอด คชจ บริหาร")
         Load frmAddEditExpenseManagenent
         frmAddEditExpenseManagenent.Show 1

         OKClick = frmAddEditExpenseManagenent.OKClick

         Unload frmAddEditExpenseManagenent
         Set frmAddEditExpenseManagenent = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.MenagementExpenses)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 16 Then
         Set frmAddEditGLAgeAmount.ParentForm = Me
         Set frmAddEditGLAgeAmount.TempCollection = m_Parameter.Glages
         frmAddEditGLAgeAmount.ShowMode = SHOW_ADD
         frmAddEditGLAgeAmount.HeaderText = MapText("เพิ่มรายละเอียด")
         Load frmAddEditGLAgeAmount
         frmAddEditGLAgeAmount.Show 1

         OKClick = frmAddEditGLAgeAmount.OKClick

         Unload frmAddEditGLAgeAmount
         Set frmAddEditGLAgeAmount = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.Glages)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 17 Then
         Set frmAddEditGLBackAmount.ParentForm = Me
         Set frmAddEditGLBackAmount.TempCollection = m_Parameter.GLbacks
         frmAddEditGLBackAmount.ShowMode = SHOW_ADD
         frmAddEditGLBackAmount.HeaderText = MapText("เพิ่มรายละเอียด")
         Load frmAddEditGLBackAmount
         frmAddEditGLBackAmount.Show 1

         OKClick = frmAddEditGLBackAmount.OKClick

         Unload frmAddEditGLBackAmount
         Set frmAddEditGLBackAmount = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.GLbacks)
            GridEX1.Rebind
         End If
      
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      Set frmAddEditApArMasAccount.TempCollection = m_Parameter.CstAccounts
'      frmAddEditApArMasAccount.ShowMode = SHOW_ADD
'      frmAddEditApArMasAccount.HeaderText = MapText("เพิ่มบัญชีลูกค้า")
'      Load frmAddEditApArMasAccount
'      frmAddEditApArMasAccount.Show 1
'
'      OKClick = frmAddEditApArMasAccount.OKClick
'
'      Unload frmAddEditApArMasAccount
'      Set frmAddEditApArMasAccount = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Parameter.CstAccounts)
'         GridEX1.Rebind
'      End If
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

End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ParamArea = 1 Then
         If ID1 <= 0 Then
            m_Parameter.BrtPrmItems.Remove (ID2)
         Else
            m_Parameter.BrtPrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.BrtPrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 2 Then
         If ID1 <= 0 Then
            m_Parameter.UsedPrmItems.Remove (ID2)
         Else
            m_Parameter.UsedPrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.UsedPrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 3 Then
         If ID1 <= 0 Then
            m_Parameter.TrnPrmItems.Remove (ID2)
         Else
            m_Parameter.TrnPrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.TrnPrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 4 Then
         If ID1 <= 0 Then
            m_Parameter.SalePrmItems.Remove (ID2)
         Else
            m_Parameter.SalePrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.SalePrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 5 Then
         If ID1 <= 0 Then
            m_Parameter.WeightPrmItems.Remove (ID2)
         Else
            m_Parameter.WeightPrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.WeightPrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 6 Then
         If ID1 <= 0 Then
            m_Parameter.CostPrmItems.Remove (ID2)
         Else
            m_Parameter.CostPrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.CostPrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 7 Then
         If ID1 <= 0 Then
            m_Parameter.AdjPrmItems.Remove (ID2)
         Else
            m_Parameter.AdjPrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.AdjPrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 9 Then
         If ID1 <= 0 Then
            m_Parameter.RvnPrmItems.Remove (ID2)
         Else
            m_Parameter.RvnPrmItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.RvnPrmItems)
         GridEX1.Rebind
      ElseIf ParamArea = 10 Then
         If ID1 <= 0 Then
            m_Parameter.CustRatioItems.Remove (ID2)
         Else
            m_Parameter.CustRatioItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.CustRatioItems)
         GridEX1.Rebind
      ElseIf ParamArea = 11 Then
         If ID1 <= 0 Then
            m_Parameter.PigStatusChangeItems.Remove (ID2)
         Else
            m_Parameter.PigStatusChangeItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.PigStatusChangeItems)
         GridEX1.Rebind
      ElseIf ParamArea = 12 Then
         If ID1 <= 0 Then
            m_Parameter.PigBuyItems.Remove (ID2)
         Else
            m_Parameter.PigBuyItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.PigBuyItems)
         GridEX1.Rebind
      ElseIf ParamArea = 13 Then
         If ID1 <= 0 Then
            m_Parameter.ExpenseSharings.Remove (ID2)
         Else
            m_Parameter.ExpenseSharings.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.ExpenseSharings)
         GridEX1.Rebind
      ElseIf ParamArea = 14 Then
         If ID1 <= 0 Then
            m_Parameter.PigAdjustItems.Remove (ID2)
         Else
            m_Parameter.PigAdjustItems.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.PigAdjustItems)
         GridEX1.Rebind
      ElseIf ParamArea = 15 Then
         If ID1 <= 0 Then
            m_Parameter.MenagementExpenses.Remove (ID2)
         Else
            m_Parameter.MenagementExpenses.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.MenagementExpenses)
         GridEX1.Rebind
      ElseIf ParamArea = 16 Then
         If ID1 <= 0 Then
            m_Parameter.Glages.Remove (ID2)
         Else
            m_Parameter.Glages.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.Glages)
         GridEX1.Rebind
      ElseIf ParamArea = 17 Then
         If ID1 <= 0 Then
            m_Parameter.GLbacks.Remove (ID2)
         Else
            m_Parameter.GLbacks.Item(ID2).Flag = "D"
         End If
   
         GridEX1.ItemCount = CountItem(m_Parameter.GLbacks)
         GridEX1.Rebind
      End If
      
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      If m_Parameter.CstAccounts.Item(ID2).MASTER_FLAG = "Y" Then
'         glbErrorLog.LocalErrorMsg = "ไม่สมารถลบบัญชีพื้นฐานได้"
'         glbErrorLog.ShowUserError
'         Exit Sub
'      End If
'
'      If ID1 <= 0 Then
'         m_Parameter.CstAccounts.Remove (ID2)
'      Else
'         m_Parameter.CstAccounts.Item(ID2).Flag = "D"
'      End If
'
'      GridEX1.itemcount = CountItem(m_Parameter.CstAccounts)
'      GridEX1.Rebind
'      m_HasModify = True
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
      If ParamArea = 1 Then
         Set frmAddEditBrthParamItem.ParentForm = Me
         frmAddEditBrthParamItem.ID = ID
         Set frmAddEditBrthParamItem.TempCollection = m_Parameter.BrtPrmItems
         frmAddEditBrthParamItem.ShowMode = SHOW_EDIT
         frmAddEditBrthParamItem.HeaderText = MapText("แก้ไขการผสม")
         Load frmAddEditBrthParamItem
         frmAddEditBrthParamItem.Show 1

         OKClick = frmAddEditBrthParamItem.OKClick

         Unload frmAddEditBrthParamItem
         Set frmAddEditBrthParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.BrtPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 2 Then
         Set frmAddEditFoodParamItem.ParentForm = Me
         frmAddEditFoodParamItem.ID = ID
         Set frmAddEditFoodParamItem.TempCollection = m_Parameter.UsedPrmItems
         frmAddEditFoodParamItem.HeaderText = MapText("แก้ไขเบอร์อาหาร/ยา")
         frmAddEditFoodParamItem.ShowMode = SHOW_EDIT
         Load frmAddEditFoodParamItem
         frmAddEditFoodParamItem.Show 1
   
         OKClick = frmAddEditFoodParamItem.OKClick
   
         Unload frmAddEditFoodParamItem
         Set frmAddEditFoodParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.UsedPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 3 Then
         Set frmAddEditLossParamItem.ParentForm = Me
         frmAddEditLossParamItem.ID = ID
         Set frmAddEditLossParamItem.TempCollection = m_Parameter.TrnPrmItems
         frmAddEditLossParamItem.HeaderText = MapText("แก้ไขการโอนตามสถานะ")
         frmAddEditLossParamItem.ShowMode = SHOW_EDIT
         Load frmAddEditLossParamItem
         frmAddEditLossParamItem.Show 1
   
         OKClick = frmAddEditLossParamItem.OKClick
   
         Unload frmAddEditLossParamItem
         Set frmAddEditLossParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.TrnPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 4 Then
         Set frmAddEditSaleParamItem.ParentForm = Me
         frmAddEditSaleParamItem.ID = ID
         Set frmAddEditSaleParamItem.TempCollection = m_Parameter.SalePrmItems
         frmAddEditSaleParamItem.HeaderText = MapText("แก้ไขราคาขายตามสถานะ")
         frmAddEditSaleParamItem.ShowMode = SHOW_EDIT
         Load frmAddEditSaleParamItem
         frmAddEditSaleParamItem.Show 1
   
         OKClick = frmAddEditSaleParamItem.OKClick
   
         Unload frmAddEditSaleParamItem
         Set frmAddEditSaleParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.SalePrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 5 Then
         Set frmAddEditWeightParamItem.ParentForm = Me
         frmAddEditWeightParamItem.ID = ID
         Set frmAddEditWeightParamItem.TempCollection = m_Parameter.WeightPrmItems
         frmAddEditWeightParamItem.HeaderText = MapText("แก้ไขน้ำหนักตามช่วงอายุ")
         frmAddEditWeightParamItem.ShowMode = SHOW_EDIT
         Load frmAddEditWeightParamItem
         frmAddEditWeightParamItem.Show 1
   
         OKClick = frmAddEditWeightParamItem.OKClick
   
         Unload frmAddEditWeightParamItem
         Set frmAddEditWeightParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.WeightPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 6 Then
         Set frmAddEditCostParamItem.ParentForm = Me
         frmAddEditCostParamItem.ID = ID
         Set frmAddEditCostParamItem.TempCollection = m_Parameter.CostPrmItems
         frmAddEditCostParamItem.HeaderText = MapText("แก้ไขราคาอาหาร/ยา")
         frmAddEditCostParamItem.ShowMode = SHOW_EDIT
         Load frmAddEditCostParamItem
         frmAddEditCostParamItem.Show 1
   
         OKClick = frmAddEditCostParamItem.OKClick
   
         Unload frmAddEditCostParamItem
         Set frmAddEditCostParamItem = Nothing
   
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.CostPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 7 Then
         Set frmAddEditAdjParamItem.ParentForm = Me
         frmAddEditAdjParamItem.ID = ID
         Set frmAddEditAdjParamItem.TempCollection = m_Parameter.AdjPrmItems
         frmAddEditAdjParamItem.HeaderText = MapText("แก้ไขยอดสุกรยกมา")
         frmAddEditAdjParamItem.ShowMode = SHOW_EDIT
         Load frmAddEditAdjParamItem
         frmAddEditAdjParamItem.Show 1

         OKClick = frmAddEditAdjParamItem.OKClick

         Unload frmAddEditAdjParamItem
         Set frmAddEditAdjParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.AdjPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 9 Then
         Set frmAddEditRevenueParamltem.ParentForm = Me
         frmAddEditRevenueParamltem.ID = ID
         Set frmAddEditRevenueParamltem.TempCollection = m_Parameter.RvnPrmItems
         frmAddEditRevenueParamltem.ShowMode = SHOW_EDIT
         frmAddEditRevenueParamltem.HeaderText = MapText("แก้ไขรายการขายอื่น ๆ")
         Load frmAddEditRevenueParamltem
         frmAddEditRevenueParamltem.Show 1

         OKClick = frmAddEditRevenueParamltem.OKClick

         Unload frmAddEditRevenueParamltem
         Set frmAddEditRevenueParamltem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.RvnPrmItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 10 Then
         Set frmAddEditCustRatioParamlItem.ParentForm = Me
         frmAddEditCustRatioParamlItem.ID = ID
         Set frmAddEditCustRatioParamlItem.TempCollection = m_Parameter.CustRatioItems
         frmAddEditCustRatioParamlItem.ShowMode = SHOW_EDIT
         frmAddEditCustRatioParamlItem.HeaderText = MapText("แก้ไข % การขาย")
         Load frmAddEditCustRatioParamlItem
         frmAddEditCustRatioParamlItem.Show 1

         OKClick = frmAddEditCustRatioParamlItem.OKClick

         Unload frmAddEditCustRatioParamlItem
         Set frmAddEditCustRatioParamlItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.CustRatioItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 11 Then
         Set frmAddEditPigTypeChangeParamlItem.ParentForm = Me
         frmAddEditPigTypeChangeParamlItem.ID = ID
         Set frmAddEditPigTypeChangeParamlItem.TempCollection = m_Parameter.PigStatusChangeItems
         frmAddEditPigTypeChangeParamlItem.ShowMode = SHOW_EDIT
         frmAddEditPigTypeChangeParamlItem.HeaderText = MapText("แก้ไขการเปลี่ยนสถานะ")
         Load frmAddEditPigTypeChangeParamlItem
         frmAddEditPigTypeChangeParamlItem.Show 1

         OKClick = frmAddEditPigTypeChangeParamlItem.OKClick

         Unload frmAddEditPigTypeChangeParamlItem
         Set frmAddEditPigTypeChangeParamlItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.PigStatusChangeItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 12 Then
         Set frmAddEditPigBuyParamItem.ParentForm = Me
         frmAddEditPigBuyParamItem.ID = ID
         Set frmAddEditPigBuyParamItem.TempCollection = m_Parameter.PigBuyItems
         frmAddEditPigBuyParamItem.ShowMode = SHOW_EDIT
         frmAddEditPigBuyParamItem.HeaderText = MapText("แก้ไขการซื้อสุกร")
         Load frmAddEditPigBuyParamItem
         frmAddEditPigBuyParamItem.Show 1

         OKClick = frmAddEditPigBuyParamItem.OKClick

         Unload frmAddEditPigBuyParamItem
         Set frmAddEditPigBuyParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.PigBuyItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 13 Then
         Set frmAddEditExpenseParamItem.ParentForm = Me
         frmAddEditExpenseParamItem.ID = ID
         Set frmAddEditExpenseParamItem.TempCollection = m_Parameter.ExpenseSharings
         frmAddEditExpenseParamItem.ShowMode = SHOW_EDIT
         frmAddEditExpenseParamItem.HeaderText = MapText("แก้ไขการปันค่าใช้จ่าย")
         Load frmAddEditExpenseParamItem
         frmAddEditExpenseParamItem.Show 1

         OKClick = frmAddEditExpenseParamItem.OKClick

         Unload frmAddEditExpenseParamItem
         Set frmAddEditExpenseParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.ExpenseSharings)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 14 Then
         Set frmAddEditPigAdjParamItem.ParentForm = Me
         frmAddEditPigAdjParamItem.ID = ID
         Set frmAddEditPigAdjParamItem.TempCollection = m_Parameter.PigAdjustItems
         frmAddEditPigAdjParamItem.ShowMode = SHOW_EDIT
         frmAddEditPigAdjParamItem.HeaderText = MapText("แก้ไขการคุมยอดสุกร")
         Load frmAddEditPigAdjParamItem
         frmAddEditPigAdjParamItem.Show 1

         OKClick = frmAddEditPigAdjParamItem.OKClick

         Unload frmAddEditPigAdjParamItem
         Set frmAddEditPigAdjParamItem = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.PigAdjustItems)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 15 Then
         Set frmAddEditExpenseManagenent.ParentForm = Me
         frmAddEditExpenseManagenent.ID = ID
         Set frmAddEditExpenseManagenent.TempCollection = m_Parameter.MenagementExpenses
         frmAddEditExpenseManagenent.ShowMode = SHOW_EDIT
         frmAddEditExpenseManagenent.HeaderText = MapText("แก้ไขยอด คชจ บริหาร")
         Load frmAddEditExpenseManagenent
         frmAddEditExpenseManagenent.Show 1

         OKClick = frmAddEditExpenseManagenent.OKClick

         Unload frmAddEditExpenseManagenent
         Set frmAddEditExpenseManagenent = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.MenagementExpenses)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 16 Then
         Set frmAddEditGLAgeAmount.ParentForm = Me
         frmAddEditGLAgeAmount.ID = ID
         Set frmAddEditGLAgeAmount.TempCollection = m_Parameter.Glages
         frmAddEditGLAgeAmount.ShowMode = SHOW_EDIT
         frmAddEditGLAgeAmount.HeaderText = MapText("แก้ไขรายละเอียด")
         Load frmAddEditGLAgeAmount
         frmAddEditGLAgeAmount.Show 1

         OKClick = frmAddEditGLAgeAmount.OKClick

         Unload frmAddEditGLAgeAmount
         Set frmAddEditGLAgeAmount = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.Glages)
            GridEX1.Rebind
         End If
      ElseIf ParamArea = 17 Then
         Set frmAddEditGLBackAmount.ParentForm = Me
         frmAddEditGLBackAmount.ID = ID
         Set frmAddEditGLBackAmount.TempCollection = m_Parameter.GLbacks
         frmAddEditGLBackAmount.ShowMode = SHOW_EDIT
         frmAddEditGLBackAmount.HeaderText = MapText("แก้ไขรายละเอียด")
         Load frmAddEditGLBackAmount
         frmAddEditGLBackAmount.Show 1

         OKClick = frmAddEditGLBackAmount.OKClick

         Unload frmAddEditGLBackAmount
         Set frmAddEditGLBackAmount = Nothing

         If OKClick Then
            GridEX1.ItemCount = CountItem(m_Parameter.GLbacks)
            GridEX1.Rebind
         End If
      
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      frmAddEditApArMasAccount.ID = ID
'      Set frmAddEditApArMasAccount.TempCollection = m_Parameter.CstAccounts
'      frmAddEditApArMasAccount.HeaderText = MapText("แก้ไขบัญชีลูกค้า")
'      frmAddEditApArMasAccount.ShowMode = SHOW_EDIT
'      Load frmAddEditApArMasAccount
'      frmAddEditApArMasAccount.Show 1
'
'      OKClick = frmAddEditApArMasAccount.OKClick
'
'      Unload frmAddEditApArMasAccount
'      Set frmAddEditApArMasAccount = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Parameter.CstAccounts)
'         GridEX1.Rebind
'      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_Parameter.GetFieldValue("PARAM_ID")
      m_Parameter.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
   
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
                  
      Call LoadProductStatus(uctlPigStatusLookup.MyCombo, m_PigStatus)
      Set uctlPigStatusLookup.MyCollection = m_PigStatus
      
      Call LoadProductType(uctlPigTypeLookup.MyCombo, m_PigTypes)
      Set uctlPigTypeLookup.MyCollection = m_PigTypes
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Parameter.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlJournalDate.ShowDate = Now
         m_Parameter.QueryFlag = 0
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
   
   Set m_Parameter = Nothing
   Set m_ApArMass = Nothing
   Set m_PigTypes = Nothing
   Set m_PigStatus = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1(Ind As Long)
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

   If Ind = 1 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2040
      Col.Caption = MapText("จากวันที่ผสม")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2190
      Col.Caption = MapText("ถึงวันที่ผสม")
   
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 1815
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("% การเข้าคลอด")
      
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 1800
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนลูกเกิด/แม่")
   
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 1410
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนลูกเกิด")
   
      Set Col = GridEX1.Columns.Add '8
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("อัตราเกิด/วัน")
   ElseIf Ind = 2 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2250
      Col.Caption = MapText("รหัสอาหาร/ยา")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 4035
      Col.Caption = MapText("ชื่ออาหาร/ยา")
   
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2910
      Col.Caption = MapText("ประเภท")
      
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("INTAKE")
   ElseIf Ind = 3 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2250
      Col.Caption = MapText("รหัสสถานะ")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 3000
      Col.Caption = MapText("ชื่อสถานะ")
         
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 3000
      Col.Caption = MapText("ประเภทสูญเสีย")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("% การสูญเสีย")
   ElseIf Ind = 4 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2000
      Col.Caption = MapText("รหัสสถานะ")
      
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2000
      Col.Caption = MapText("ชื่อสถานะ")
         
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2000
      Col.Caption = MapText("จากวันที่")
      
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2000
      Col.Caption = MapText("ถึงวันที่")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคาขาย")
   ElseIf Ind = 5 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2250
      Col.Caption = MapText("จากอายุ")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2130
      Col.Caption = MapText("ถึงอายุ")
         
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("น้ำหนักเฉลี่ย")
   
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 4785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("หมายเหตุ")
   ElseIf Ind = 6 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2250
      Col.Caption = MapText("รหัสอาหาร/ยา")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 4035
      Col.Caption = MapText("ชื่ออาหาร/ยา")
   
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2910
      Col.Caption = MapText("ประเภท")
      
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
   ElseIf Ind = 7 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2550
      Col.Caption = MapText("สัปดาห์เกิด")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 6550
      Col.Caption = MapText("รายละเอียดสุกร")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2425
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวน")
   ElseIf Ind = 9 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2040
      Col.Caption = MapText("จากวันที่ขาย")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2190
      Col.Caption = MapText("ถึงวันที่ขาย")
   
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 1815
      Col.Caption = MapText("รหัสสินค้า")
      
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 4095
      Col.Caption = MapText("ชื่อสินค้า")
   
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 1410
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนที่ขาย")
   
      Set Col = GridEX1.Columns.Add '8
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา/หน่วย")
   
      Set Col = GridEX1.Columns.Add '9
      Col.Width = 2025
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนเงิน")
   ElseIf Ind = 10 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2340
      Col.Caption = MapText("รหัสลูกค้า")
   
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 4915
      Col.Caption = MapText("ชื่อลูกค้า")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2000
      Col.Caption = MapText("ประเภท %")
      
      Set Col = GridEX1.Columns.Add '5
      Col.TextAlignment = jgexAlignRight
      Col.Width = 2000
      Col.Caption = MapText("% การขาย")
   ElseIf Ind = 11 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2550
      Col.Caption = MapText("ประเภทสุกร")
      
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 6210
      Col.Caption = MapText("รายละเอียดประเภทสุกร")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2425
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("% การเปลี่ยนสถานะ")
   ElseIf Ind = 12 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2010
      Col.Caption = MapText("วันที่ซื้อ")
      
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2070
      Col.Caption = MapText("รหัสสัปดาห์เกิด")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2850
      Col.Caption = MapText("รายละเอียด")
   
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวน")
   
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
   ElseIf Ind = 13 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2010
      Col.Caption = MapText("วันที่ ค.ช.จ.")
      
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2070
      Col.Caption = MapText("รหัสค่าใช้จ่าย")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 2850
      Col.Caption = MapText("รายละเอียด")
   
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวน")
   
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
   ElseIf Ind = 14 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2010
      Col.Caption = MapText("จากวันที่")
      
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 2070
      Col.Caption = MapText("ถึงวันที่")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 5280
      Col.Caption = MapText("สัปดาห์เกิด")
   
      Set Col = GridEX1.Columns.Add '6
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนที่คุม")
   ElseIf Ind = 15 Then
      Set Col = GridEX1.Columns.Add '3
      Col.Width = 2000
      Col.Caption = MapText("เดือนปี")
      
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX1.Columns.Add '5
      Col.Width = 8000
      Col.Caption = MapText("รายละเอียด")
   ElseIf Ind = 16 Or Ind = 17 Then
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 3500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("อายุ (วัน)")
      
      Set Col = GridEX1.Columns.Add '7
      Col.Width = 3500
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนตัว")
   End If
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
   Col.Width = 2925
   Col.Caption = MapText("เลขที่บัญชี")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 6270
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3240
   Col.Caption = MapText("แพคเกจ")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblJournalCode, MapText("เลขที่"))
   Call InitNormalLabel(lblJournalDate, MapText("วันที่"))
   Call InitNormalLabel(lblJournalDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromAge, MapText("จากอายุ"))
   Call InitNormalLabel(lblToAge, MapText("ถึงอายุ"))
   Call InitNormalLabel(Label1, MapText("สัปดาห์"))
   Call InitNormalLabel(Label2, MapText("สัปดาห์"))
   Call InitNormalLabel(lblPigType, MapText("ประเภทสุกร"))
   Call InitNormalLabel(lblPigStatus, MapText("สถานะสุกร"))
   Call InitNormalLabel(lblADG, MapText("ADG"))
   
   Call txtJournalCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtJournalDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFromAge.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtToAge.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtADG.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   uctlPigStatusLookup.Enabled = False
   txtADG.Visible = False
   lblADG.Visible = False
   If (ParamArea = 1) Or (ParamArea = 5) Or (ParamArea = 6) Or (ParamArea = 7) Or (ParamArea = 9) Then
      txtFromAge.Enabled = False
      txtToAge.Enabled = False
   ElseIf (ParamArea = 2) Then
      txtADG.Visible = True
      txtADG.Enabled = True
      lblADG.Visible = True
   ElseIf (ParamArea = 10) Then
      txtFromAge.Enabled = False
      txtToAge.Enabled = False
      uctlPigStatusLookup.Enabled = True
   ElseIf (ParamArea = 12) Then
      txtFromAge.Enabled = False
      txtToAge.Enabled = False
      uctlPigStatusLookup.Enabled = False
      uctlPigTypeLookup.Enabled = False
   ElseIf (ParamArea = 13) Or (ParamArea = 14) Or (ParamArea = 15) Then
      txtFromAge.Enabled = False
      txtToAge.Enabled = False
      uctlPigStatusLookup.Enabled = False
      uctlPigTypeLookup.Enabled = False
   ElseIf (ParamArea = 16) Or (ParamArea = 17) Then
      txtFromAge.Enabled = False
      txtToAge.Enabled = False
      uctlPigStatusLookup.Enabled = False
   End If
   
   Call InitCheckBox(chkPostFlag, "ห้ามแก้ไข")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   Call InitGrid1(ParamArea)
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   If ParamArea = 1 Then
      TabStrip1.Tabs.Add().Caption = MapText("การผสม")
   ElseIf ParamArea = 2 Then
      TabStrip1.Tabs.Add().Caption = MapText("เบอร์อาหาร/ยา")
   ElseIf ParamArea = 3 Then
      TabStrip1.Tabs.Add().Caption = MapText("% การสูญเสีย")
   ElseIf ParamArea = 4 Then
      TabStrip1.Tabs.Add().Caption = MapText("ราคาขาย/ก.ก.")
   ElseIf ParamArea = 5 Then
      TabStrip1.Tabs.Add().Caption = MapText("น้ำหนักเฉลี่ย")
   ElseIf ParamArea = 6 Then
      TabStrip1.Tabs.Add().Caption = MapText("ราคาอาหาร/ยา")
   ElseIf ParamArea = 7 Then
      TabStrip1.Tabs.Add().Caption = MapText("ยอดหมูยกมา")
   ElseIf ParamArea = 9 Then
      TabStrip1.Tabs.Add().Caption = MapText("รายการขายอื่น ๆ")
   ElseIf ParamArea = 10 Then
      TabStrip1.Tabs.Add().Caption = MapText("% การขายตามลูกค้า")
   ElseIf ParamArea = 11 Then
      TabStrip1.Tabs.Add().Caption = MapText("% เปลี่ยนสถานะสุกร")
   ElseIf ParamArea = 12 Then
      TabStrip1.Tabs.Add().Caption = MapText("การซื้อสุกร")
   ElseIf ParamArea = 13 Then
      TabStrip1.Tabs.Add().Caption = MapText("ปันค่าใช้จ่าย")
   ElseIf ParamArea = 14 Then
      TabStrip1.Tabs.Add().Caption = MapText("คุมยอดสุกร")
   ElseIf ParamArea = 15 Then
      TabStrip1.Tabs.Add().Caption = MapText("ค่าใช้จ่ายขาย/บริหาร")
   ElseIf ParamArea = 16 Then
      TabStrip1.Tabs.Add().Caption = MapText("ยกมาหมู GL")
   ElseIf ParamArea = 17 Then
      TabStrip1.Tabs.Add().Caption = MapText("กลับสัตว์")
   End If
   
   cmdOK.Enabled = (ShowMode <> SHOW_VIEW_ONLY)
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
   Set m_Parameter = New CParameter
   Set m_ApArMass = New Collection
   Set m_PigTypes = New Collection
   Set m_PigStatus = New Collection
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim Pk As CPackage
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = Val(GridEX1.Value(2))
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("COPY")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      If ParamArea = 1 Then
         Dim Bi As CBrtPrmItem
         Dim TempBi As CBrtPrmItem
         Set Bi = m_Parameter.BrtPrmItems(TempID1)
         Set TempBi = New CBrtPrmItem
         TempBi.Flag = "A"
         
         Call Bi.CopyItemCollection(TempBi)
         Call m_Parameter.BrtPrmItems.Add(TempBi)
         
         Set Bi = Nothing
         Set TempBi = Nothing
         Call TabStrip1_Click
         m_HasModify = True
      ElseIf ParamArea = 9 Then
         Dim Ji As CRvnPrmItem
         Dim TempJi As CRvnPrmItem
         Set Ji = m_Parameter.RvnPrmItems(TempID1)
         Set TempJi = New CRvnPrmItem
         TempJi.Flag = "A"
         Call Ji.CopyItemCollection(TempJi)
         Call m_Parameter.RvnPrmItems.Add(TempJi)
         
         Set Ji = Nothing
         Set TempJi = Nothing
         Call TabStrip1_Click
         m_HasModify = True
      ElseIf ParamArea = 14 Then
         Dim Pi As CParamItem
         Dim TempPi As CParamItem
         Set Pi = m_Parameter.PigAdjustItems(TempID1)
         Set TempPi = New CParamItem
         TempPi.Flag = "A"
         Call Pi.CopyItemCollection(TempPi)
         Call m_Parameter.PigAdjustItems.Add(TempPi)

         Set Pi = Nothing
         Set TempPi = Nothing
         Call TabStrip1_Click
         m_HasModify = True
      End If
   End If
   
   Call EnableForm(Me, True)
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
      If m_Parameter.UsedPrmItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If ParamArea = 1 Then
         Dim Bi As CBrtPrmItem
         If m_Parameter.BrtPrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set Bi = GetItem(m_Parameter.BrtPrmItems, RowIndex, RealIndex)
         If Bi Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Bi.GetFieldValue("BRTPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = DateToStringExtEx2(Bi.GetFieldValue("FROM_BREED"))
         Values(4) = DateToStringExtEx2(Bi.GetFieldValue("TO_BREED"))
         Values(5) = Bi.GetFieldValue("BREED_RATE")
         Values(6) = Bi.GetFieldValue("CHILD_RATE")
         Values(7) = Bi.GetFieldValue("BIRTH_AMOUNT")
         Values(8) = Bi.GetFieldValue("BIRTH_RATE")
      ElseIf ParamArea = 2 Then
         Dim CR As CUsedPrmItem
         If m_Parameter.UsedPrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set CR = GetItem(m_Parameter.UsedPrmItems, RowIndex, RealIndex)
         If CR Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = CR.GetFieldValue("USEDPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = CR.GetFieldValue("FOOD_NO")
         Values(4) = CR.GetFieldValue("FOOD_NAME")
         Values(5) = CR.GetFieldValue("PART_TYPE_NAME") & " (" & InTakeTypeToString(CR.GetFieldValue("INTAKE_TYPE")) & ")"
         Values(6) = CR.GetFieldValue("USED_RATE")
      ElseIf ParamArea = 3 Then
         Dim Ti As CTrnPrmItem
         If m_Parameter.TrnPrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set Ti = GetItem(m_Parameter.TrnPrmItems, RowIndex, RealIndex)
         If Ti Is Nothing Then
            Exit Sub
         End If

         Values(1) = Ti.GetFieldValue("TRNPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Ti.GetFieldValue("STATUS_NO")
         Values(4) = Ti.GetFieldValue("STATUS_NAME")
         Values(5) = GetLossType(Ti.GetFieldValue("LOSS_TYPE"))
         Values(6) = Ti.GetFieldValue("LOSS_RATE")
      ElseIf ParamArea = 4 Then
         Dim Si As CSalePrmItem
         If m_Parameter.SalePrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set Si = GetItem(m_Parameter.SalePrmItems, RowIndex, RealIndex)
         If Si Is Nothing Then
            Exit Sub
         End If

         Values(1) = Si.GetFieldValue("SALEPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Si.GetFieldValue("STATUS_NO")
         Values(4) = Si.GetFieldValue("STATUS_NAME")
         Values(5) = DateToStringExtEx2(Si.GetFieldValue("FROM_SALE"))
         Values(6) = DateToStringExtEx2(Si.GetFieldValue("TO_SALE"))
         Values(7) = (Si.GetFieldValue("SALE_RATE"))
      ElseIf ParamArea = 5 Then
         Dim Wi As CWeightPrmItem
         If m_Parameter.WeightPrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set Wi = GetItem(m_Parameter.WeightPrmItems, RowIndex, RealIndex)
         If Wi Is Nothing Then
            Exit Sub
         End If

         Values(1) = Wi.GetFieldValue("WEIGHTPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Wi.GetFieldValue("FROM_AGE")
         Values(4) = Wi.GetFieldValue("TO_AGE")
         Values(5) = (Wi.GetFieldValue("UNIT_WEIGHT"))
      ElseIf ParamArea = 6 Then
         Dim Ci As CCostPrmItem
         If m_Parameter.CostPrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set Ci = GetItem(m_Parameter.CostPrmItems, RowIndex, RealIndex)
         If Ci Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Ci.GetFieldValue("COSTPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Ci.GetFieldValue("FOOD_NO")
         Values(4) = Ci.GetFieldValue("FOOD_NAME")
         Values(5) = Ci.GetFieldValue("PART_TYPE_NAME")
         Values(6) = (Ci.GetFieldValue("COST_RATE"))
      ElseIf ParamArea = 7 Then
         Dim Ap As CAdjPrmItem
         If m_Parameter.AdjPrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set Ap = GetItem(m_Parameter.AdjPrmItems, RowIndex, RealIndex)
         If Ap Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Ap.GetFieldValue("ADJPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Ap.GetFieldValue("PIG_NO")
         Values(4) = Ap.GetFieldValue("PIG_NAME")
         Values(5) = (Ap.GetFieldValue("PIG_AMOUNT"))
      ElseIf ParamArea = 9 Then
         Dim Ri As CRvnPrmItem
         If m_Parameter.RvnPrmItems.Count <= 0 Then
            Exit Sub
         End If
         Set Ri = GetItem(m_Parameter.RvnPrmItems, RowIndex, RealIndex)
         If Ri Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Ri.GetFieldValue("RVNPRM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = DateToStringExtEx2(Ri.GetFieldValue("FROM_SALE"))
         Values(4) = DateToStringExtEx2(Ri.GetFieldValue("TO_SALE"))
         Values(5) = Ri.GetFieldValue("REVENUE_NO")
         Values(6) = Ri.GetFieldValue("REVENUE_NAME")
         Values(7) = (Ri.GetFieldValue("SALE_AMOUNT"))
         Values(8) = (Ri.GetFieldValue("UNIT_PRICE"))
         Values(9) = (Ri.GetFieldValue("TOTAL_PRICE"))
      ElseIf ParamArea = 10 Then
         Dim Pi1 As CParamItem
         If m_Parameter.CustRatioItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pi1 = GetItem(m_Parameter.CustRatioItems, RowIndex, RealIndex)
         If Pi1 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pi1.GetFieldValue("PARAM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Pi1.GetFieldValue("CUSTOMER_CODE")
         Values(4) = Pi1.GetFieldValue("CUSTOMER_NAME")
         Values(5) = GetSellShareType(Pi1.GetFieldValue("SHARE_SELL_TYPE"))
         Values(6) = (Pi1.GetFieldValue("SALE_RATIO"))
      ElseIf ParamArea = 11 Then
         Dim Pi2 As CParamItem
         If m_Parameter.PigStatusChangeItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pi2 = GetItem(m_Parameter.PigStatusChangeItems, RowIndex, RealIndex)
         If Pi2 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pi2.GetFieldValue("PARAM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Pi2.GetFieldValue("PIG_TYPE_NO")
         Values(4) = Pi2.GetFieldValue("PIG_TYPE_NAME")
         Values(5) = (Pi2.GetFieldValue("TRANSFER_RATE"))
      ElseIf ParamArea = 12 Then
         Dim Pi3 As CParamItem
         If m_Parameter.PigBuyItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pi3 = GetItem(m_Parameter.PigBuyItems, RowIndex, RealIndex)
         If Pi3 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pi3.GetFieldValue("PARAM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = DateToStringExtEx2(Pi3.GetFieldValue("BUY_DATE"))
         Values(4) = Pi3.GetFieldValue("PIG_NO")
         Values(5) = Pi3.GetFieldValue("PIG_DESC")
         Values(6) = (Pi3.GetFieldValue("BUY_AMOUNT"))
         Values(7) = (Pi3.GetFieldValue("BUY_TOTAL_PRICE"))
      ElseIf ParamArea = 13 Then
         Dim Pi4 As CParamItem
         If m_Parameter.ExpenseSharings.Count <= 0 Then
            Exit Sub
         End If
         Set Pi4 = GetItem(m_Parameter.ExpenseSharings, RowIndex, RealIndex)
         If Pi4 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pi4.GetFieldValue("PARAM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = DateToStringExtEx2(Pi4.GetFieldValue("EXPENSE_DATE"))
         Values(4) = Pi4.GetFieldValue("EXPENSE_TYPE_NAME")
         Values(5) = Pi4.GetFieldValue("EXPENSE_NAME")
         Values(6) = (Pi4.GetFieldValue("EXP_AMOUNT"))
         Values(7) = (Pi4.GetFieldValue("EXP_TOTAL_PRICE"))
      ElseIf ParamArea = 14 Then
         Dim Pi5 As CParamItem
         If m_Parameter.PigAdjustItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pi5 = GetItem(m_Parameter.PigAdjustItems, RowIndex, RealIndex)
         If Pi5 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pi5.GetFieldValue("PARAM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = DateToStringExtEx2(Pi5.GetFieldValue("CTRL_FROM_DATE"))
         Values(4) = DateToStringExtEx2(Pi5.GetFieldValue("CTRL_TO_DATE"))
         Values(5) = Pi5.GetFieldValue("PIG_DESC")
         Values(6) = (Pi5.GetFieldValue("CTRL_AMOUNT"))
      ElseIf ParamArea = 15 Then
         Dim Pi6 As CParamItem
         If m_Parameter.MenagementExpenses.Count <= 0 Then
            Exit Sub
         End If
         Set Pi6 = GetItem(m_Parameter.MenagementExpenses, RowIndex, RealIndex)
         If Pi6 Is Nothing Then
            Exit Sub
         End If
   
         Values(1) = Pi6.GetFieldValue("PARAM_ITEM_ID")
         Values(2) = RealIndex
         Values(3) = Left(Pi6.GetFieldValue("YYYYMM"), 4) + 543 & "-" & Right(Pi6.GetFieldValue("YYYYMM"), 2)
         Values(4) = (Pi6.GetFieldValue("EXP_AMOUNT"))
         Values(5) = Pi6.GetFieldValue("EXPENSE_NAME")
      ElseIf ParamArea = 16 Then
         Dim GLa As CGLAgeAmount
         If m_Parameter.Glages.Count <= 0 Then
            Exit Sub
         End If
         Set GLa = GetItem(m_Parameter.Glages, RowIndex, RealIndex)
         If GLa Is Nothing Then
            Exit Sub
         End If
         
         Values(1) = GLa.GL_AGE_AMOUNT_ID
         Values(2) = RealIndex
         Values(3) = GLa.GL_AGE
         Values(4) = GLa.GL_AMOUNT
      ElseIf ParamArea = 17 Then
         Dim GLb As CGLBackAmount
         If m_Parameter.GLbacks.Count <= 0 Then
            Exit Sub
         End If
         Set GLb = GetItem(m_Parameter.GLbacks, RowIndex, RealIndex)
         If GLb Is Nothing Then
            Exit Sub
         End If
         
         Values(1) = GLb.GL_BACK_AMOUNT_ID
         Values(2) = RealIndex
         Values(3) = GLb.GL_AGE
         Values(4) = GLb.GL_AMOUNT
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      If m_Parameter.CstAccounts Is Nothing Then
'         Exit Sub
'      End If
'
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'
'      Dim Ca As CAccount
'      If m_Parameter.CstAccounts.count <= 0 Then
'         Exit Sub
'      End If
'      Set Ca = GetItem(m_Parameter.CstAccounts, RowIndex, RealIndex)
'      If Ca Is Nothing Then
'         Exit Sub
'      End If
'
'      Values(1) = Ca.ACCOUNT_ID
'      Values(2) = RealIndex
'      Values(3) = Ca.ACCOUNT_NO
'      Values(4) = Ca.NOTE
'      Values(5) = ""
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub RefreshGrid(Flag As Boolean)
   If ParamArea = 1 Then
      GridEX1.ItemCount = CountItem(m_Parameter.BrtPrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 2 Then
      GridEX1.ItemCount = CountItem(m_Parameter.UsedPrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 3 Then
      GridEX1.ItemCount = CountItem(m_Parameter.TrnPrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 4 Then
      GridEX1.ItemCount = CountItem(m_Parameter.SalePrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 5 Then
      GridEX1.ItemCount = CountItem(m_Parameter.WeightPrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 6 Then
      GridEX1.ItemCount = CountItem(m_Parameter.CostPrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 7 Then
      GridEX1.ItemCount = CountItem(m_Parameter.AdjPrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 9 Then
      GridEX1.ItemCount = CountItem(m_Parameter.RvnPrmItems)
      GridEX1.Rebind
   ElseIf ParamArea = 10 Then
      GridEX1.ItemCount = CountItem(m_Parameter.CustRatioItems)
      GridEX1.Rebind
   ElseIf ParamArea = 11 Then
      GridEX1.ItemCount = CountItem(m_Parameter.PigStatusChangeItems)
      GridEX1.Rebind
   ElseIf ParamArea = 12 Then
      GridEX1.ItemCount = CountItem(m_Parameter.PigBuyItems)
      GridEX1.Rebind
   ElseIf ParamArea = 13 Then
      GridEX1.ItemCount = CountItem(m_Parameter.ExpenseSharings)
      GridEX1.Rebind
   ElseIf ParamArea = 14 Then
      GridEX1.ItemCount = CountItem(m_Parameter.PigAdjustItems)
      GridEX1.Rebind
   ElseIf ParamArea = 15 Then
      GridEX1.ItemCount = CountItem(m_Parameter.MenagementExpenses)
      GridEX1.Rebind
   ElseIf ParamArea = 16 Then
      GridEX1.ItemCount = CountItem(m_Parameter.Glages)
      GridEX1.Rebind
   ElseIf ParamArea = 17 Then
      GridEX1.ItemCount = CountItem(m_Parameter.GLbacks)
      GridEX1.Rebind
   End If
   
   m_HasModify = Flag
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1(ParamArea)
      Call RefreshGrid(False)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      Call InitGrid2
'      GridEX1.itemcount = CountItem(m_Parameter.CstAccounts)
'      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtADG_Change()
   m_HasModify = True
End Sub

Private Sub txtFromAge_Change()
   m_HasModify = True
End Sub

Private Sub txtJournalDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtToAge_Change()
   m_HasModify = True
End Sub
Private Sub txtJournalCode_Change()
   m_HasModify = True
End Sub
Private Sub uctlJournalDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlPigStatusLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlPigTypeLookup_Change()
   m_HasModify = True
End Sub
