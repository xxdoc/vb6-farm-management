VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRunSimulate 
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   Icon            =   "frmRunSimulate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   11850
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   720
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1270
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   405
         Left            =   5100
         TabIndex        =   2
         Top             =   90
         Width           =   3225
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlDate uctlDate1 
         Height          =   405
         Left            =   780
         TabIndex        =   0
         Top             =   90
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   435
         Left            =   3360
         TabIndex        =   1
         Top             =   90
         Width           =   585
         _ExtentX        =   1244
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   405
         Left            =   10080
         TabIndex        =   7
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRunSimulate.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   405
         Left            =   8400
         TabIndex        =   6
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRunSimulate.frx":2ABC
         ButtonStyle     =   3
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   5
         Top             =   90
         Width           =   615
      End
      Begin VB.Label lblCurrentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   90
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRunSimulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Batch As CBatch
Private m_ApArMass As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ParamArea As Long

Private ApArText As String
Private FileName As String
Private m_Populations As Collection
Private m_BirthParams As Collection
Private m_Houses As Collection
Private m_SaleHouses As Collection
Private m_Locations As Collection
Private m_TempFeed As Collection
Private m_TempTransf As Collection
Private m_FeedUsed As Collection
Private m_CostParams As Collection
Private m_SaleParams As Collection
Private m_RevenueParams As Collection
Private m_RevenueAccum As Collection
Private m_Feeds As Collection
Private m_ExpenseTypes As Collection
Private m_PigBuyParams As Collection
Private m_ProductTypes As Collection
Private m_Pigs As Collection
Private m_PigStatusSellItems As Collection
Private m_PigTypeStatusCustomers As Collection
Private m_ExpenseSharing As Collection
Private m_PigAdjustItems As Collection
Private m_Adgs As Collection
Private m_InTakeFoods As Collection
Private m_ExportPerDay As Collection

Private m_PartItemsLocationMonthlies As Collection
Private m_PartItemsLocations As Collection
Private CcostColls1 As Collection
Public JournalType As Long
Private m_GLAgecoll As Collection
Private m_GLBackcoll As Collection

Private BFodd(12) As Double
Private BExp(12) As Double
Private Birth(12) As Double
Private Food(12) As Double
Private Expense(12) As Double
Private PigIDBirthInMonthColl As Collection
Private DoItemBirthInMonthColl As Collection

Private SumFood As Double
Private FromInSertDate As Date
Private ToInSertDate As Date

Private Row As Long
Private Function IsExist(TempCol As Collection, Key As String) As Boolean
Dim p As CPopulation

   IsExist = False
   For Each p In TempCol
      If p.PIG_ID = Key Then
         IsExist = True
         Exit For
      End If
   Next p
End Function

Private Sub GenerateInitialPopulation()
Dim Bl As CAdjPrmItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CPopulation
Dim TempDate As Date
Dim I As Long
Dim Pi As CPartItem
Dim Pt As CProductType

   Set TempRs = New ADODB.Recordset
   
   'glbErrorLog.LocalErrorMsg = "GenerateInitialPopulation"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   
   For Each Bi In m_Batch.Balances
      Set Bl = New CAdjPrmItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Call Bl.PopulateFromRS(1, TempRs)
         If Not IsExist(m_Populations, Bl.GetFieldValue("PIG_ID")) Then
            Set Pp = New CPopulation
            Pp.PIG_ID = Bl.GetFieldValue("PIG_ID")
            
            Pp.PIG_NO = Bl.GetFieldValue("PIG_NO")
            Pp.PIG_NAME = Bl.GetFieldValue("PIG_NAME")
            Pp.PIG_TYPE = Bl.GetFieldValue("PIG_TYPE")
            Pp.CURRENT_AMOUNT = Bl.GetFieldValue("PIG_AMOUNT")
            Pp.CURRENT_AGE = GetAge(Pp.PIG_NO, FromInSertDate)
            Pp.AVG_WEIGHT = Bl.GetFieldValue("AVG_WEIGHT")
            Pp.TOTAL_WEIGHT = Pp.CURRENT_AMOUNT * Pp.AVG_WEIGHT
            Pp.FEED_COST = Bl.GetFieldValue("FEED_COST")
            Pp.EXPENSE_COST = Bl.GetFieldValue("EXPENSE_COST")
            Pp.MEDICINE_COST = Bl.GetFieldValue("MEDICINE_COST")
            Pp.BIRTH_COST = Bl.GetFieldValue("BIRTH_COST")
            
            'glbErrorLog.LocalErrorMsg = Pp.PIG_ID & "-" & Pp.CURRENT_AMOUNT
            'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
            
            Pp.Flag = "A"
            'If Pp.PIG_ID = 12641 Then
               Call m_Populations.Add(Pp, Trim(Str(Pp.PIG_ID)))
            'End If
            Set Pp = Nothing
         End If
         
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
   
   'glbErrorLog.LocalErrorMsg = "GenerateInitialPopulation"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
            
   TempDate = FromInSertDate
   While TempDate <= ToInSertDate
      TempDate = DateAdd("D", 1, TempDate)
      
'      For Each Pt In m_ProductTypes
         Set Pi = glbDaily.DateToPartItem(TempDate, PigCodeToID("N"))
         If Not (Pi Is Nothing) Then
            If Not IsExist(m_Populations, Pi.PART_ITEM_ID) Then            'เพิ่มเข้าใน Collection เฉพาะ PartItemID ที่เป็นหมู N และไม่มีใน Collection   เท่านั้น
               Set Pp = New CPopulation
               Pp.PIG_ID = Pi.PART_ITEM_ID
               Pp.PIG_NO = Pi.PART_NO
               Pp.PIG_NAME = Pi.PART_DESC
               Pp.PIG_TYPE = Pi.PIG_TYPE
               Pp.CURRENT_AMOUNT = 0
               Pp.CURRENT_AGE = 0
               Pp.TOTAL_WEIGHT = 0
               Pp.Flag = "A"
               
               'glbErrorLog.LocalErrorMsg = Pp.PIG_ID & "-" & Pp.CURRENT_AMOUNT
               'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
               
               
               'If Pp.PIG_ID = 12641 Then
                  Call m_Populations.Add(Pp, Trim(Str(Pp.PIG_ID)))
               'End If
                        
               Set Pp = Nothing
            End If
         End If
'      Next Pt
   Wend
   
   'glbErrorLog.LocalErrorMsg = "GenerateInitialPopulation"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GenerateFeedCostParam()
Dim Bl As CCostPrmItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CCostPrmItem
Dim TempDate As Date
Dim I As Long
Dim Pi As CPartItem

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.Feeds
      Set Bl = New CCostPrmItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Pp = New CCostPrmItem
         
         Call Pp.PopulateFromRS(1, TempRs)
         Pp.Flag = "A"
         Call m_CostParams.Add(Pp)
         
         Set Pp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GenerateRevenueParam()
Dim Bl As CRvnPrmItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CRvnPrmItem
Dim TempDate As Date
Dim I As Long

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.Revenues
      Set Bl = New CRvnPrmItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Pp = New CRvnPrmItem
         
         Call Pp.PopulateFromRS(1, TempRs)
         Pp.Flag = "A"
         Call m_RevenueParams.Add(Pp)
         
         Set Pp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GeneratePigBuyParam()
Dim Bl As CParamItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CRvnPrmItem
Dim TempDate As Date
Dim I As Long

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.BuyItems
      Set Bl = New CParamItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Bl = New CParamItem
         
         Call Bl.PopulateFromRS(1, TempRs)
         
         'If Bl.GetFieldValue("PIG_ID") <= 0 Then
            '''debug.print
         'End If
         Bl.Flag = "A"
         Call m_PigBuyParams.Add(Bl)
         
         Set Bl = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GeneratePigAdjParam()
Dim Bl As CParamItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CRvnPrmItem
Dim TempDate As Date
Dim I As Long

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.PigAdjItems
      Set Bl = New CParamItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Bl = New CParamItem
         
         Call Bl.PopulateFromRS(1, TempRs)
         Bl.Flag = "A"
         Call m_PigAdjustItems.Add(Bl)
         
         Set Bl = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GenerateADGParam()
   Call LoadADGParam(Nothing, m_Adgs, ID, 2)
End Sub

Private Function GetAdgRate(Pp As CPopulation) As Double
Dim Bi As CBatchItem
Dim TempADG As Double

   TempADG = 0
   For Each Bi In m_Adgs
      If (Bi.GetFieldValue("FROM_AGE") <= Pp.CURRENT_AGE) And _
         (Bi.GetFieldValue("TO_AGE") >= Pp.CURRENT_AGE) And _
         (PigTypeToCode(Bi.GetFieldValue("PIG_TYPE")) = Pp.PIG_TYPE) Then
            TempADG = Bi.GetFieldValue("ADG")
            Exit For
      End If
   Next Bi
   GetAdgRate = TempADG
End Function

Private Sub GenerateExpenseSharing()
Dim Bl As CParamItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CRvnPrmItem
Dim TempDate As Date
Dim I As Long

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.ExpenseSharingItems
      Set Bl = New CParamItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Bl = New CParamItem
         
         Call Bl.PopulateFromRS(1, TempRs)
         Bl.Flag = "A"
         Call m_ExpenseSharing.Add(Bl)
         
         Set Bl = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GeneratePigStatusCustomerParam()
Dim Bl As CParamItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CRvnPrmItem
Dim TempDate As Date
Dim I As Long

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.CustRatios
      Set Bl = New CParamItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Bl = New CParamItem
         
         Call Bl.PopulateFromRS(1, TempRs)
         Bl.Flag = "A"
         Call m_PigTypeStatusCustomers.Add(Bl)
         
         Set Bl = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GenerateBirthParam()
Dim Bl As CBrtPrmItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CBrtPrmItem
Dim TempDate As Date
Dim I As Long
Dim Pi As CPartItem

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.BirthItems
      Set Bl = New CBrtPrmItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Pp = New CBrtPrmItem
         
         Call Pp.PopulateFromRS(1, TempRs)
         Pp.Flag = "A"
         Call Pp.SetFieldValue("BIRTH_COST", 0)
         Call m_BirthParams.Add(Pp)
         
         Set Pp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GenerateSaleParam()
Dim Bl As CSalePrmItem
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Pp As CSalePrmItem
Dim TempDate As Date
Dim I As Long

   Set TempRs = New ADODB.Recordset
   
   For Each Bi In m_Batch.SaleItems
      Set Bl = New CSalePrmItem
      Call Bl.SetFieldValue("PARAM_ID", Bi.GetFieldValue("PARAM_ID"))
      Call Bl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Pp = New CSalePrmItem
         
         Call Pp.PopulateFromRS(1, TempRs)
         Pp.Flag = "A"
         Call m_SaleParams.Add(Pp)
         
         Set Pp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Bl = Nothing
   Next Bi
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_Batch.SetFieldValue("BATCH_ID", ID)
      m_Batch.QueryFlag = 1
      If Not glbDaily.QueryBatch(m_Batch, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         'glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Batch.PopulateFromRS(1, m_Rs)
      
      FromInSertDate = m_Batch.GetFieldValue("EXECUTE_FROM")
      ToInSertDate = m_Batch.GetFieldValue("EXECUTE_TO")
      Me.Caption = HeaderText & " BATCH : " & m_Batch.GetFieldValue("BATCH_NO") & "   จากวันที่ " & FromInSertDate & "-" & ToInSertDate
      
      'Call RefreshGrid(False)
   Else
      ShowMode = SHOW_ADD
   End If
   
   If Not IsOK Then
      'glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub MergeItem(B As CBatch)
Dim Bi As CBatchItem

   Set B.BatchItems = Nothing
   Set B.BatchItems = New Collection
   
   For Each Bi In B.BirthItems
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.FoodItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.SaleItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.TransferItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.WeightItems
      Call B.BatchItems.Add(Bi)
   Next Bi
End Sub

Private Sub cmdOK_Click()
      OKClick = True
      Unload Me
End Sub
Private Function GetBirthRate(BirthDate As Date, AvgWeight As Double, Bpi As CBrtPrmItem) As Double
Dim Bi As CBrtPrmItem

   GetBirthRate = 0
   AvgWeight = 0
   For Each Bi In m_BirthParams
      If (BirthDate >= Bi.GetFieldValue("FROM_BIRTH")) And (BirthDate <= Bi.GetFieldValue("TO_BIRTH")) Then
         GetBirthRate = Bi.GetFieldValue("BIRTH_RATE")
         AvgWeight = Bi.GetFieldValue("AVG_WEIGHT")
         Set Bpi = Bi
         Exit For
      End If
   Next Bi
End Function
Private Function GetSalePrice(Pp As CPopulation, PigStatus As Long, CREDIT As Double, TempDate As Date) As Double
Dim Bi As CSalePrmItem

   GetSalePrice = 0
   For Each Bi In m_SaleParams
      If (PigStatus = Bi.GetFieldValue("PIG_STATUS")) And (Pp.CURRENT_AGE >= Bi.GetFieldValue("FROM_AGE")) And (Pp.CURRENT_AGE <= Bi.GetFieldValue("TO_AGE")) And (TempDate >= Bi.GetFieldValue("FROM_SALE")) And (TempDate <= Bi.GetFieldValue("TO_SALE")) Then
         GetSalePrice = Bi.GetFieldValue("SALE_RATE")
         CREDIT = Bi.GetFieldValue("CREDIT")
         Exit For
      End If
   Next Bi
End Function

Private Function GetFeedCost(FoodID As Long) As Double
Dim Bi As CCostPrmItem

   GetFeedCost = 0
   For Each Bi In m_CostParams
      '''debug.print Bi.GetFieldValue("FOOD_NAME")
      If Bi.GetFieldValue("FOOD_ID") = FoodID Then
         GetFeedCost = Bi.GetFieldValue("COST_RATE")
         Exit For
      End If
   Next Bi
End Function

Private Sub GeneratePigBirthDocument(BirthDate As Date, ByVal BirthAmount As Double, Flag As Boolean, Pi As CPartItem, Pp As CPopulation, Bpi As CBrtPrmItem)
Dim Ivd As CInventoryDoc
Dim II As CImportItem
Dim IsOK As Boolean
'Dim Cm As CCapitalMovement
'Dim Ci As CMovementItem

Static RunNo As Long
Dim O As Object
   If Not Flag Then
      Exit Sub
   End If
   
   If BirthAmount <= 0 Then
      Exit Sub
   End If
   
   RunNo = RunNo + 1
   
   '=====
   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.INVENTORY_DOC_ID = -1
    Ivd.DOCUMENT_DATE = BirthDate
   Ivd.DOCUMENT_NO = "PB-" & Format(RunNo, "000000") & "-" & Format(ID, "0000") & "-N"
   Ivd.EMP_ID = -1
   Ivd.DOCUMENT_TYPE = 5
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.SIMULATE_FLAG = "Y"
   Ivd.BATCH_ID = ID

   Set II = New CImportItem
   II.Flag = "A"
   Call Ivd.ImportExports.Add(II)

   II.TX_TYPE = "I"
   II.BIRTH_DATE = BirthDate
   II.INCLUDE_UNIT_PRICE = 0
   II.ACTUAL_UNIT_PRICE = 0
   II.PART_ITEM_ID = Pi.PART_ITEM_ID
   II.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   II.FATHER_NO = ""
   II.MOTHER_NO = ""
   II.IMPORT_AMOUNT = BirthAmount
   II.TOTAL_WEIGHT = Pp.AVG_WEIGHT * II.IMPORT_AMOUNT
   II.CALCULATE_FLAG = "N"
       
   Call UpDateCostColls(Pi.PART_ITEM_ID, , , , , , , , , , , , Bpi.GetFieldValue("FROM_BIRTH"), Bpi.GetFieldValue("TO_BIRTH"))
   'มีหมูเกิดขึ้นมาในระบบต้องทำการเพิ่มหมูเข้าไปใน CcostColls1
   Set O = II
   O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
   O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
   Call GeneratePartItemLocationMonthly(O)
         
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   Set II = Nothing
   '=====
   Set Ivd = Nothing
End Sub

Private Function GetMatchFoodParam(TempCol As Collection, Pp As CPopulation) As CBatchItem
Dim Bi As CBatchItem
Dim PigTypeID As Long
   
   PigTypeID = PigCodeToID(Pp.PIG_TYPE)
   Set GetMatchFoodParam = Nothing
   For Each Bi In TempCol
       If (Bi.GetFieldValue("FROM_AGE") <= Pp.CURRENT_AGE) And (Pp.CURRENT_AGE <= Bi.GetFieldValue("TO_AGE")) And _
         (Bi.GetFieldValue("PIG_TYPE") = PigTypeID) Then
         Set GetMatchFoodParam = Bi
         Exit For
       End If
   Next Bi
End Function

Private Function GetMatchPigTypeChangeParam(TempCol As Collection, Pp As CPopulation) As CBatchItem
Dim Bi As CBatchItem
Dim PigTypeID As Long

   PigTypeID = PigCodeToID(Pp.PIG_TYPE)
   Set GetMatchPigTypeChangeParam = Nothing
   For Each Bi In TempCol
       If (Bi.GetFieldValue("FROM_AGE") <= Pp.CURRENT_AGE) And (Pp.CURRENT_AGE <= Bi.GetFieldValue("TO_AGE")) And _
         (Bi.GetFieldValue("PIG_TYPE") = PigTypeID) Then
         Set GetMatchPigTypeChangeParam = Bi
         Exit For
       End If
   Next Bi
End Function

Private Function MyGetParameter(TempCol As Collection, Key As String) As CParameter
Dim Pi As CParameter

   Set MyGetParameter = Nothing
   For Each Pi In TempCol
      If Pi.GetFieldValue("PARAM_ID") = Key Then
         Set MyGetParameter = Pi
         Exit For
      End If
   Next Pi
End Function

Private Function MyGetPopulation(TempCol As Collection, Key As String) As CPopulation
Dim Pi As CPopulation

   Set MyGetPopulation = Nothing
   For Each Pi In TempCol
      If Pi.PIG_ID = Key Then
         Set MyGetPopulation = Pi
         Exit For
      End If
   Next Pi
End Function
Private Sub GeneratePigFeedDocument(TempDate As Date, Bi As CBatchItem, Flag As Boolean, Pp As CPopulation, Mode As Long)
Dim Ui As CUsedPrmItem

Dim Ci As CCostPrmItem

Dim Pm As CParameter
Dim ParamID As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Static Ivd As CInventoryDoc
Dim EI As CExportItem
Dim Lc As CLocation
Dim TempPP As CPopulation
Static RunNo As Long
Static PrevDate As Date
Dim HaveCost  As Boolean
   If Not Flag Then
      Exit Sub
   End If
   
   If (Pp.CURRENT_AMOUNT <= 0) And (Mode = 2) Then
      Exit Sub
   End If
   
   If Mode = 1 Then
      RunNo = RunNo + 1
      
      Set Ivd = New CInventoryDoc
      Ivd.AddEditMode = SHOW_ADD
      Ivd.INVENTORY_DOC_ID = -1
       Ivd.DOCUMENT_DATE = TempDate
      Ivd.DOCUMENT_NO = "EXP-" & Format(RunNo, "00000") & "-" & Format(ID, "0000")
      Ivd.DELIVERY_FEE = 0
      Ivd.EMP_ID = -1
      Ivd.DOCUMENT_TYPE = 2
      Ivd.COMMIT_FLAG = "N"
      Ivd.SALE_FLAG = "N"
      Ivd.EXCEPTION_FLAG = "N"
      Ivd.SIMULATE_FLAG = "Y"
      Ivd.BATCH_ID = ID
   ElseIf Mode = 2 Then
      Set TempRs = New ADODB.Recordset
      
      ParamID = Bi.GetFieldValue("PARAM_ID")
      Set Pm = Nothing
      Set Pm = MyGetParameter(m_TempFeed, Trim(Str(ParamID)))
      If Pm Is Nothing Then
         Set Pm = New CParameter
      End If
      
      If Pm.GetFieldValue("PARAM_ID") <= 0 Then
         'ไม่พบให้คิวรี่มาแล้วใส่เข้าไปใน m_TempFeed
         Call Pm.SetFieldValue("PARAM_ID", ParamID)
         Pm.QueryFlag = 1
         Call glbDaily.QueryParameter(Pm, TempRs, iCount, IsOK, glbErrorLog)
         If Not TempRs.EOF Then
            Call Pm.PopulateFromRS(1, TempRs)
            Call m_TempFeed.Add(Pm, Trim(Str(ParamID)))
         End If
      End If
      
      Set Lc = m_Locations("00")
      For Each Ui In Pm.UsedPrmItems
         Set EI = New CExportItem
         EI.EXPORT_ITEM_ID = -1
         EI.TX_TYPE = "E"
         EI.PART_ITEM_ID = Ui.GetFieldValue("FOOD_ID")
         EI.PART_NO = Ui.GetFieldValue("FOOD_NO")
         EI.LOCATION_ID = Lc.LOCATION_ID
         EI.EXPORT_AMOUNT = Ui.GetFieldValue("USED_RATE") * Pp.CURRENT_AMOUNT
         EI.HOUSE_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         
'         ''debug.print (Ui.GetFieldValue("USED_RATE"))
'         If Ui.GetFieldValue("USED_RATE") = "0.6" Then
'            ''debug.print
'         End If

         HaveCost = False
         For Each Ci In m_CostParams
'            ''debug.print (Ci.GetFieldValue("FOOD_ID") & "-" & Ci.GetFieldValue("FOOD_NO") & "-" & Ci.GetFieldValue("FOOD_NAME"))
'            ''debug.print (PigTypeToCode(Ci.GetFieldValue("PIG_TYPE")))
            If (Ci.GetFieldValue("FOOD_ID") = EI.PART_ITEM_ID) And (Pp.PIG_TYPE = PigTypeToCode(Ci.GetFieldValue("PIG_TYPE"))) Then
               EI.EXPORT_AVG_PRICE = Ci.GetFieldValue("COST_RATE")
               EI.EXPORT_TOTAL_PRICE = EI.EXPORT_AMOUNT * Ci.GetFieldValue("COST_RATE")
               
               If EI.EXPORT_AVG_PRICE > 0 Then
                  HaveCost = True
               End If
               
               If Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L" Or Pp.PIG_TYPE = "B" Then
                  '''debug.print
                  SumFood = SumFood + EI.EXPORT_TOTAL_PRICE
               End If
               If Ui.GetFieldValue("INTAKE_TYPE") = 1 Then
                  Call UpDateCostColls(Pp.PIG_ID, , , , EI.EXPORT_TOTAL_PRICE) 'หมูการกินอาหารในระบบต้องทำการเพิ่มต้นทุนหมูเข้าไปใน CcostColls1
               ElseIf Ui.GetFieldValue("INTAKE_TYPE") = 2 Then
                  Call UpDateCostColls(Pp.PIG_ID, , , , , , , , , , EI.EXPORT_TOTAL_PRICE) 'หมูการกินอาหารในระบบต้องทำการเพิ่มต้นทุนหมูเข้าไปใน CcostColls1
               ElseIf Ui.GetFieldValue("INTAKE_TYPE") = 3 Then
                  Call UpDateCostColls(Pp.PIG_ID, , , , , , , , , , , EI.EXPORT_TOTAL_PRICE) 'หมูการกินอาหารในระบบต้องทำการเพิ่มต้นทุนหมูเข้าไปใน CcostColls1
               End If
            End If
         Next Ci
         
         If Not (HaveCost) Then
            glbErrorLog.LocalErrorMsg = "ยังไม่มีการใส่ราคา   " & Ui.GetFieldValue("FOOD_NO") & "      ประเภทหมู        " & Pp.PIG_TYPE
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            '''debug.print "ยังไม่มีการใส่ราคา   " & Ui.GetFieldValue("FOOD_NO") & "      ประเภทหมู        " & Pp.PIG_TYPE
         End If
         
         EI.PIG_ID = Pp.PIG_ID
         EI.CALCULATE_FLAG = "Y"
         EI.Flag = "A"
         
         Dim Itk As CIntake
         Set Itk = GetObject("CIntake", m_InTakeFoods, Trim(Ivd.DOCUMENT_DATE & "-" & Pp.PIG_ID & "-" & EI.PART_ITEM_ID), False)
         If Itk Is Nothing Then
            Set Itk = New CIntake
            Itk.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
            Itk.PIG_ID = Pp.PIG_ID
            Itk.PART_ITEM_ID = EI.PART_ITEM_ID
            Itk.INTAKE_COST = EI.EXPORT_TOTAL_PRICE
            Itk.INTAKE_AMOUNT = EI.EXPORT_AMOUNT
            Itk.CURRENT_PIG_AMOUNT = Pp.CURRENT_AMOUNT
            Call m_InTakeFoods.Add(Itk, Trim(Ivd.DOCUMENT_DATE & "-" & Pp.PIG_ID & "-" & EI.PART_ITEM_ID))
         Else
            Itk.INTAKE_COST = Itk.INTAKE_COST + EI.EXPORT_TOTAL_PRICE
            Itk.INTAKE_AMOUNT = Itk.INTAKE_AMOUNT + EI.EXPORT_AMOUNT
            Itk.CURRENT_PIG_AMOUNT = Itk.CURRENT_PIG_AMOUNT + Pp.CURRENT_AMOUNT
         End If
         Set Itk = Nothing
         
         Call Ivd.ImportExports.Add(EI)
         
         'คำนวณ น้ำหนักต่อตัวที่เพิ่มขึ้นมาในการกินอาหาร
         'ไม่ต้องคำนวณตรงนี้
'         Pp.AVG_WEIGHT = Pp.AVG_WEIGHT + Ui.GetFieldValue("ADG_RATE") * Ui.GetFieldValue("USED_RATE")
'         Pp.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Pp.CURRENT_AMOUNT
         
         Set TempPP = Nothing
         Set TempPP = GetPopulation(m_FeedUsed, Trim(Str(EI.PART_ITEM_ID)))
         If TempPP.PIG_ID <= 0 Then
            'ยังไม่มีการใช้อาหารเบอร์นี้มาก่อน
            
            'ยืมมาใช้สำหรับเก็บยอดใช้อาหาร   เพื่อเมื่อถึงสิ้นเดือนจะได้ทใบนำเข้าวัตถุดิบทีเดี่ยว
            Set TempPP = New CPopulation
            TempPP.PIG_ID = EI.PART_ITEM_ID
            TempPP.USED_AMOUNT = EI.EXPORT_AMOUNT
            TempPP.PIG_STATUS = Ui.GetFieldValue("INTAKE_TYPE")        'เก็บค่าประเภท INTAKE_TYPE
            Call m_FeedUsed.Add(TempPP, Trim(Str(TempPP.PIG_ID)))
            Set TempPP = Nothing
         Else
            'มีเบอร์อาหารนี้อยู่แล้ว
            TempPP.USED_AMOUNT = TempPP.USED_AMOUNT + EI.EXPORT_AMOUNT
         End If
               
         Set EI = Nothing
      Next Ui
      Set TempRs = Nothing
   ElseIf Mode = 3 Then
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
      Set Ivd = Nothing
   End If
End Sub

Private Sub CreateImportExportItems(Ivd As CInventoryDoc)
Dim Ti As CTransferItem
Dim EI As CExportItem
Dim II As CImportItem

   Set Ivd.ImportExports = Nothing
   Set Ivd.ImportExports = New Collection
   
   For Each Ti In Ivd.TransferItems
      Set EI = Ti.ExportItem
      Set II = Ti.ImportItem
      
      EI.Flag = Ti.Flag
      II.Flag = Ti.Flag
      
      Call Ivd.ImportExports.Add(EI)
      Call Ivd.ImportExports.Add(II)
   Next Ti
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
   Ivd.DOCUMENT_TYPE = 10
   Ivd.DOCUMENT_SUBTYPE = Bd.DOCUMENT_SUBTYPE
   Ivd.SIMULATE_FLAG = Bd.SIMULATE_FLAG
   Ivd.BATCH_ID = Bd.BATCH_ID
   
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

'Private Sub GeneratePigTransferDocumentOld(TempDate As Date, Bi As CBatchItem, Flag As Boolean, Pp As CPopulation, TransferAmount As Double)
'Dim Ui As CTrnPrmItem
'Dim Pm As CParameter
'Dim ParamID As Long
'Dim IsOK As Boolean
'Dim TempRs As ADODB.Recordset
'Dim iCount As Long
'Dim Ivd As CInventoryDoc
'Dim Bd As CBillingDoc
'Dim Tr As CTransferItem
'Dim Ei As CExportItem
'Dim Ii As CImportItem
'Dim Lc As CLocation
'Dim Di As CDoItem
'Dim TrnAmount As Double
'Dim DueDate As Date
'Dim CREDIT As Double
'Static RunNo As Long
'
'   If Not Flag Then
'      Exit Sub
'   End If
'
'   If Pp.CURRENT_AMOUNT <= 0 Then
'      Exit Sub
'   End If
'
'   TrnAmount = 0
'   RunNo = RunNo + 1
'   Set TempRs = New ADODB.Recordset
'
'   ParamID = Bi.GetFieldValue("PARAM_ID")
'   Set Pm = Nothing
'   Set Pm = MyGetParameter(m_TempTransf, Trim(Str(ParamID)))
'   If Pm Is Nothing Then
'      Set Pm = New CParameter
'   End If
'
'   If Pm.GetFieldValue("PARAM_ID") <= 0 Then
'      'ไม่พบให้คิวรี่มาแล้วใส่เข้าไปใน m_TempTransf
'      Call Pm.SetFieldValue("PARAM_ID", ParamID)
'      Pm.QueryFlag = 1
'      Call glbDaily.QueryParameter(Pm, TempRs, iCount, IsOK, glbErrorLog)
'      If Not TempRs.EOF Then
'         Call Pm.PopulateFromRS(1, TempRs)
'         Call m_TempTransf.Add(Pm, Trim(Str(ParamID)))
'      End If
'   End If
'
'   'มาถึงจุดนี้ Pm จะมีค่าแน่นอน
'   Set Ivd = New CInventoryDoc
'   Ivd.AddEditMode = SHOW_ADD
'   Ivd.INVENTORY_DOC_ID = -1
'    Ivd.DOCUMENT_DATE = TempDate
'   Ivd.DOCUMENT_NO = "TRN-" & Format(RunNo, "00000")
'   Ivd.DELIVERY_FEE = 0
'   Ivd.EMP_ID = -1
'   Ivd.DOCUMENT_TYPE = 7
'   Ivd.COMMIT_FLAG = "N"
'   Ivd.SALE_FLAG = "N"
'   Ivd.EXCEPTION_FLAG = "N"
'   Ivd.SIMULATE_FLAG = "Y"
'   Ivd.BATCH_ID = ID
'
'   DueDate = DateAdd("D", 30, TempDate)
'
'   Set Bd = New CBillingDoc
'   Bd.AddEditMode = SHOW_ADD
'   Bd.BILLING_DOC_ID = ID
'    Bd.DOCUMENT_DATE = TempDate
'    Bd.DUE_DATE = DueDate
'   Bd.DOCUMENT_NO = "SL-" & Format(RunNo, "00000")
''   Bd.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
'   Bd.BILLING_ADDRESS_ID = -1
'   Bd.ENTERPRISE_ADDRESS_ID = -1
'   Bd.DOCUMENT_TYPE = 1
'   Bd.DOCUMENT_SUBTYPE = 1
'   Bd.EXCEPTION_FLAG = "N"
'   Bd.ACCEPT_BY = -1
'   Bd.COMMIT_FLAG = "N"
'   Bd.PAYMENT_TYPE = -1
'   Bd.BATCH_ID = ID
'   Bd.SIMULATE_FLAG = "Y"
'   Call PopulateGuiID(Bd)
'
'   Set Lc = m_SaleHouses(1) 'เรือนขาย
'   For Each Ui In Pm.TrnPrmItems
'      Set Tr = New CTransferItem
'      Set Ei = New CExportItem
'      Set Ii = New CImportItem
'
'      Tr.Flag = "A"
'      Ei.Flag = "A"
'      Ei.CALCULATE_FLAG = "N"
'      Ii.Flag = "A"
'      Ii.CALCULATE_FLAG = "N"
'
'      Set Tr.ExportItem = Ei
'      Set Tr.ImportItem = Ii
'
'      Tr.ExportItem.PART_ITEM_ID = Pp.PIG_ID
'      Tr.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
'      Tr.ExportItem.EXPORT_AMOUNT = (Ui.GetFieldValue("LOSS_RATE") * Pp.CURRENT_AMOUNT / 100) 'Round(Ui.GetFieldValue("LOSS_RATE") * Pp.CURRENT_AMOUNT / 100)
'      TrnAmount = TrnAmount + Tr.ExportItem.EXPORT_AMOUNT
'      Tr.ExportItem.PIG_STATUS = Ui.GetFieldValue("PIG_STATUS")
'      Tr.ExportItem.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Tr.ExportItem.EXPORT_AMOUNT
'
'      Tr.ImportItem.PART_ITEM_ID = Pp.PIG_ID
'      Tr.ImportItem.LOCATION_ID = Lc.LOCATION_ID
'      Tr.ImportItem.IMPORT_AMOUNT = Tr.ExportItem.EXPORT_AMOUNT
'      Tr.ImportItem.TOTAL_WEIGHT = Tr.ExportItem.TOTAL_WEIGHT
'      Tr.ImportItem.PIG_STATUS = Tr.ExportItem.PIG_STATUS
'      If Tr.ExportItem.EXPORT_AMOUNT > 0 Then
'         Call Ivd.TransferItems.Add(Tr)
'      End If
'
'      Set Di = New CDoItem
'      Di.Flag = "A"
'      Di.PART_ITEM_ID = Pp.PIG_ID
'      Di.ITEM_AMOUNT = Tr.ExportItem.EXPORT_AMOUNT
'      Di.LOCATION_ID = Lc.LOCATION_ID
'      Di.PIG_STATUS = Tr.ExportItem.PIG_STATUS
'      Di.TOTAL_WEIGHT = Tr.ExportItem.TOTAL_WEIGHT
'      Di.AVG_PRICE = GetSalePrice(Pp, Di.PIG_STATUS, CREDIT, TempDate)
'      Di.TOTAL_PRICE = Di.AVG_PRICE * Di.TOTAL_WEIGHT
'      Di.AVG_WEIGHT = Pp.AVG_WEIGHT
'      If Tr.ExportItem.EXPORT_AMOUNT > 0 Then
'         Call Bd.DoItems.Add(Di)
'      End If
'
'      Set Tr = Nothing
'      Set Di = Nothing
'   Next Ui
'
'   Call CreateImportExportItems(Ivd)
'   If Ivd.TransferItems.Count > 0 Then
'      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
'   End If
'
'   'สร้างบิลขายเพื่อตัดออกไปเลย
'   If Bd.DoItems.Count > 0 Then
'      Call DO2InventoryDoc(Bd, Ivd)
'      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
'      Bd.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
'      Call glbDaily.AddEditBillingDoc(Bd, IsOK, False, glbErrorLog)
'   End If
'
'   Set Bd = Nothing
'   Set Ivd = Nothing
'
'   If TempRs.State = adStateOpen Then
'      Call TempRs.Close
'   End If
'   Set TempRs = Nothing
'
'   TransferAmount = TrnAmount
'End Sub
Private Sub GeneratePigTransferDocument(TempDate As Date, Bi As CBatchItem, Flag As Boolean, Pp As CPopulation, TransferAmount As Double, Mode As Long)
Dim Ui As CTrnPrmItem
Dim Pm As CParameter
Dim ParamID As Long
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim Tr As CTransferItem
Dim EI As CExportItem
Dim II As CImportItem
Dim Lc As CLocation
Dim TrnAmount As Double
Dim MyAmount As Double
Static Ivd As CInventoryDoc
Static RunNo As Long
Dim TempPP As CPopulation
Dim Key As String
Dim NeedAmount As Double
Dim O As Object
Dim ETemp As CExportTemp
   
   If Not Flag Then
      Exit Sub
   End If
   
   If (Pp.CURRENT_AMOUNT <= 0) And (Mode = 2) Then
      Exit Sub
   End If
   
   If Mode = 1 Then
      Set Ivd = New CInventoryDoc
      Ivd.AddEditMode = SHOW_ADD
      Ivd.INVENTORY_DOC_ID = -1
       Ivd.DOCUMENT_DATE = TempDate
      Ivd.DOCUMENT_NO = "TRN-" & Format(RunNo, "00000") & "-" & Format(ID, "0000")
      Ivd.DELIVERY_FEE = 0
      Ivd.EMP_ID = -1
      Ivd.DOCUMENT_TYPE = 7
      Ivd.COMMIT_FLAG = "N"
      Ivd.SALE_FLAG = "N"
      Ivd.EXCEPTION_FLAG = "N"
      Ivd.SIMULATE_FLAG = "Y"
      Ivd.BATCH_ID = ID
      RunNo = RunNo + 1
   ElseIf Mode = 2 Then
      TrnAmount = 0
      Set TempRs = New ADODB.Recordset
      
      ParamID = Bi.GetFieldValue("PARAM_ID")
      Set Pm = Nothing
      Set Pm = MyGetParameter(m_TempTransf, Trim(Str(ParamID)))
      If Pm Is Nothing Then
         Set Pm = New CParameter
      End If
      
      If Pm.GetFieldValue("PARAM_ID") <= 0 Then
         'ไม่พบให้คิวรี่มาแล้วใส่เข้าไปใน m_TempTransf
         Call Pm.SetFieldValue("PARAM_ID", ParamID)
         Pm.QueryFlag = 1
         Call glbDaily.QueryParameter(Pm, TempRs, iCount, IsOK, glbErrorLog)
         If Not TempRs.EOF Then
            Call Pm.PopulateFromRS(1, TempRs)
            Call m_TempTransf.Add(Pm, Trim(Str(ParamID)))
         End If
      End If
      
      'glbErrorLog.LocalErrorMsg = "GeneratePigTransferDocument"
      'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   
      Set Lc = m_SaleHouses(1) 'เรือนขาย
      For Each Ui In Pm.TrnPrmItems
         Set Tr = New CTransferItem
         Set EI = New CExportItem
         Set II = New CImportItem
         
         Tr.Flag = "A"
         EI.Flag = "A"
         EI.CALCULATE_FLAG = "N"
         II.Flag = "A"
         II.CALCULATE_FLAG = "N"
         
         Set Tr.ExportItem = EI
         Set Tr.ImportItem = II
'If Ui.GetFieldValue("PIG_STATUS") <= 0 Then
''''debug.print
'End If
         Tr.ExportItem.PART_ITEM_ID = Pp.PIG_ID
         If Pp.PIG_ID <= 0 Then
            '''debug.print
         End If
         Tr.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         MyAmount = Pp.CURRENT_AMOUNT - TrnAmount
         If Ui.GetFieldValue("LOSS_TYPE") = 1 Then
            'คิดแบบ %
            'Set ETemp = GetObject("CExportTemp", m_ExportPerDay, Trim(Pp.PIG_ID & "-" & Pp.CURRENT_AGE & "-" & Ui.GetFieldValue("PIG_STATUS")), False)
            'If Not (ETemp Is Nothing) Then
             '  Tr.ExportItem.EXPORT_AMOUNT = ETemp.EXPORT_PER_DAY
            'Else
               Tr.ExportItem.EXPORT_AMOUNT = (Ui.GetFieldValue("LOSS_RATE") * Pp.CURRENT_AMOUNT / 700) 'Round(Ui.GetFieldValue("LOSS_RATE") * Pp.CURRENT_AMOUNT / 100)
'               Set ETemp = New CExportTemp
'               ETemp.PIG_ID = Pp.PIG_ID
'               ETemp.PIG_AGE = Pp.CURRENT_AGE
'               ETemp.PIG_STATUS = Ui.GetFieldValue("PIG_STATUS")
'               ETemp.EXPORT_PER_DAY = Tr.ExportItem.EXPORT_AMOUNT
'               Call m_ExportPerDay.Add(ETemp, Trim(Pp.PIG_ID & "-" & Pp.CURRENT_AGE))
'            End If
            
            If MyAmount < Tr.ExportItem.EXPORT_AMOUNT Then 'เอาจำนวนที่เหลือจริง ๆ
               Tr.ExportItem.EXPORT_AMOUNT = MyAmount
            End If
         ElseIf Ui.GetFieldValue("LOSS_TYPE") = 2 Then
            'คิดแบบหักเป็นจำนวนตัว/วัน
            If MyAmount >= Ui.GetFieldValue("LOSS_RATE") Then
               Tr.ExportItem.EXPORT_AMOUNT = Ui.GetFieldValue("LOSS_RATE")
            Else
               Tr.ExportItem.EXPORT_AMOUNT = MyAmount
            End If
         ElseIf Ui.GetFieldValue("LOSS_TYPE") = 3 Then
            'คิดแบบหักแล้วให้เหลือเป็นจำนวนที่ต้องการ
'            If Pp.CURRENT_AGE = 9 Then
'               ''debug.print
'            End If
            Set ETemp = GetObject("CExportTemp", m_ExportPerDay, Trim(Pp.PIG_ID & "-" & Pp.CURRENT_AGE & "-" & Ui.GetFieldValue("PIG_STATUS")), False)
            If Not (ETemp Is Nothing) Then
               Tr.ExportItem.EXPORT_AMOUNT = ETemp.EXPORT_PER_DAY
               If ETemp.LIMIT_AMOUNT > (MyAmount - Tr.ExportItem.EXPORT_AMOUNT) Then 'ไม่ต่ำกว่าที่ตั้งไว้
                  Tr.ExportItem.EXPORT_AMOUNT = MyAmount - ETemp.LIMIT_AMOUNT
               End If
            Else
               NeedAmount = (MyAmount - Ui.GetFieldValue("LOSS_RATE"))
               If NeedAmount >= 0 Then
                  Tr.ExportItem.EXPORT_AMOUNT = MyDiff(NeedAmount, 7)
                  Set ETemp = New CExportTemp
                  ETemp.PIG_ID = Pp.PIG_ID
                  ETemp.PIG_AGE = Pp.CURRENT_AGE
                  ETemp.PIG_STATUS = Ui.GetFieldValue("PIG_STATUS")
                  ETemp.EXPORT_PER_DAY = Tr.ExportItem.EXPORT_AMOUNT
                  ETemp.LIMIT_AMOUNT = Ui.GetFieldValue("LOSS_RATE")
                  Call m_ExportPerDay.Add(ETemp, Trim(Pp.PIG_ID & "-" & Pp.CURRENT_AGE & "-" & Ui.GetFieldValue("PIG_STATUS")))
               Else
                  Tr.ExportItem.EXPORT_AMOUNT = 0 'MyAmount
               End If
            End If
            
            If MyAmount < Tr.ExportItem.EXPORT_AMOUNT Then 'เอาจำนวนที่เหลือจริง ๆ
               Tr.ExportItem.EXPORT_AMOUNT = MyAmount
            End If
         End If
         TrnAmount = TrnAmount + Tr.ExportItem.EXPORT_AMOUNT
         Tr.ExportItem.PIG_STATUS = Ui.GetFieldValue("PIG_STATUS")
         Tr.ExportItem.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Tr.ExportItem.EXPORT_AMOUNT
         
         Tr.ImportItem.PART_ITEM_ID = Pp.PIG_ID
         Tr.ImportItem.LOCATION_ID = Lc.LOCATION_ID      ' Import เข้าเรือนขาย
         Tr.ImportItem.IMPORT_AMOUNT = Tr.ExportItem.EXPORT_AMOUNT
         Tr.ImportItem.TOTAL_WEIGHT = Tr.ExportItem.TOTAL_WEIGHT
         Tr.ImportItem.PIG_STATUS = Tr.ExportItem.PIG_STATUS
         If Tr.ExportItem.EXPORT_AMOUNT > 0 Then
            Call Ivd.TransferItems.Add(Tr)
            
            Key = Pp.PIG_ID & "-" & Tr.ExportItem.PIG_STATUS
            
            Set TempPP = GetPopulationEx(m_PigStatusSellItems, Key)
            If TempPP Is Nothing Then
               Set TempPP = New CPopulation
               
               TempPP.PIG_ID = Tr.ImportItem.PART_ITEM_ID
               TempPP.PIG_NO = Pp.PIG_NO
               TempPP.PIG_STATUS = Tr.ImportItem.PIG_STATUS
               TempPP.CURRENT_AMOUNT = Tr.ImportItem.IMPORT_AMOUNT
               TempPP.CURRENT_AGE = Pp.CURRENT_AGE
               TempPP.AVG_WEIGHT = Pp.AVG_WEIGHT
               
               Call m_PigStatusSellItems.Add(TempPP, Key)
            Else
               TempPP.CURRENT_AMOUNT = TempPP.CURRENT_AMOUNT + Tr.ImportItem.IMPORT_AMOUNT
            End If
            
            Set O = Tr.ExportItem
            O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
            O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
            Call GeneratePartItemLocationMonthly(O, False)        'ที่ไม่ต้องทำอะไรเพราะว่าเป็นการโอนเข้าและออกของหมู เบอร์เดี่ยวกัน ซึ่งจะทำให้ต้นทุนต่อตัวยังเท่าเดิม
            
            Set O = Tr.ImportItem
            O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
            O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
            Call GeneratePartItemLocationMonthly(O, False)        'ที่ไม่ต้องทำอะไรเพราะว่าเป็นการโอนเข้าและออกของหมู เบอร์เดี่ยวกัน ซึ่งจะทำให้ต้นทุนต่อตัวยังเท่าเดิม
            
            Set TempPP = Nothing
         End If
               
         Set Tr = Nothing
      Next Ui
   
      If TempRs.State = adStateOpen Then
         Call TempRs.Close
      End If
      Set TempRs = Nothing
      
      TransferAmount = TrnAmount
      
      'glbErrorLog.LocalErrorMsg = "GeneratePigTransferDocument"
      'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
      
   ElseIf Mode = 3 Then
      Call CreateImportExportItems(Ivd)
      'glbErrorLog.LocalErrorMsg = "----------------------------------"
      'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
      
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   End If
End Sub

Private Sub GenerateFeedImportDocument(TempDate As Date, Flag As Boolean)
Static RunNo As Long
Dim TempPP As CPopulation
Dim Ivd As CInventoryDoc
Dim II As CImportItem
Dim IsOK As Boolean
Dim FD As Date
Dim Ed As Date
Dim Lc As CLocation

   Call GetFirstLastDate(TempDate, FD, Ed)
   
   If Not Flag Then
      Exit Sub
   End If
   
   If m_FeedUsed.Count <= 0 Then
      Exit Sub
   End If
      
   RunNo = RunNo + 1
   
   Set Ivd = New CInventoryDoc
   
   Ivd.AddEditMode = SHOW_ADD
   Ivd.INVENTORY_DOC_ID = ID
   Ivd.DOCUMENT_DATE = FD 'นำเข้าตั้งแต่ต้นเดือนเพื่อให้มีใช้ภายในเดือน
   Ivd.DUE_DATE = DateAdd("D", 32, Ivd.DOCUMENT_DATE)   'ยังไงก็ให้เหลื่อมไปจ่ายเดือนหน้า
   Ivd.DOCUMENT_NO = "IMP-" & Format(RunNo, "00000") & "-" & Format(ID, "00000")
   Ivd.SUPPLIER_ID = -1
   Ivd.DELIVERY_ID = -1
   Ivd.DOCUMENT_TYPE = 1
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.BATCH_ID = ID
   Ivd.SIMULATE_FLAG = "Y"
   
   Set Lc = m_Locations("00")
   For Each TempPP In m_FeedUsed
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I"
      II.PART_ITEM_ID = TempPP.PIG_ID
      II.LOCATION_ID = Lc.LOCATION_ID
      II.IMPORT_AMOUNT = TempPP.USED_AMOUNT
      II.ACTUAL_UNIT_PRICE = GetFeedCost(II.PART_ITEM_ID)
      II.CALCULATE_FLAG = "Y"
      II.INTAKE_TYPE = TempPP.PIG_STATUS     'นำมาใช้แทน INTAKE_TYPE
      II.TOTAL_ACTUAL_PRICE = II.ACTUAL_UNIT_PRICE * II.IMPORT_AMOUNT
      II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE
      II.TOTAL_INCLUDE_PRICE = II.TOTAL_ACTUAL_PRICE
      
      Call Ivd.ImportExports.Add(II)
      Set II = Nothing
   Next TempPP
   
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   Set Ivd = Nothing
   
   Set m_FeedUsed = Nothing
   Set m_FeedUsed = New Collection
End Sub

Private Sub GenerateRevenueSaleDocument(TempDate As Date, Flag As Boolean)
Static RunNo As Long
Dim TempPP As CPopulation
Dim Bd As CBillingDoc
Dim Di As CDoItem
Dim IsOK As Boolean

   If Not Flag Then
      Exit Sub
   End If
   
   If m_RevenueAccum.Count <= 0 Then
      Exit Sub
   End If
      
   RunNo = RunNo + 1
   
   Set Bd = New CBillingDoc
   Bd.AddEditMode = SHOW_ADD
   Bd.BILLING_DOC_ID = ID
    Bd.DOCUMENT_DATE = TempDate
    Bd.DUE_DATE = TempDate
   Bd.DOCUMENT_NO = "SO-" & Format(RunNo, "00000") & "-" & Format(ID, "0000")
'   Bd.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
   Bd.BILLING_ADDRESS_ID = -1
   Bd.ENTERPRISE_ADDRESS_ID = -1
   Bd.DOCUMENT_TYPE = 1
   Bd.DOCUMENT_SUBTYPE = 1
   Bd.EXCEPTION_FLAG = "N"
   Bd.ACCEPT_BY = -1
   Bd.COMMIT_FLAG = "N"
   Bd.PAYMENT_TYPE = -1
   Bd.BATCH_ID = ID
   Bd.SIMULATE_FLAG = "Y"
'   Call PopulateGuiID(Bd)

   For Each TempPP In m_RevenueAccum
      Set Di = New CDoItem
      Di.Flag = "A"
      Di.PART_ITEM_ID = -1
      Di.REVENUE_ID = TempPP.PIG_ID
      Di.ITEM_AMOUNT = TempPP.CURRENT_AMOUNT
      Di.LOCATION_ID = -1
      Di.PIG_STATUS = -1
      Di.TOTAL_WEIGHT = 0
      Di.AVG_PRICE = MyDiffEx(TempPP.TOTAL_PRICE, TempPP.CURRENT_AMOUNT)
      Di.TOTAL_PRICE = TempPP.TOTAL_PRICE
      Di.AVG_WEIGHT = 0
      
      Call Bd.DoItems.Add(Di)
      Set Di = Nothing
   Next TempPP

   If Bd.DoItems.Count > 0 Then
      Call glbDaily.AddEditBillingDoc(Bd, IsOK, False, glbErrorLog)
   End If
   Set Bd = Nothing
   
   Set m_RevenueAccum = Nothing
   Set m_RevenueAccum = New Collection
End Sub

Private Function IsEndOfMonth(TempDate As Date) As Boolean
Dim StartDate As Date
Dim EndDate As Date

   Call GetFirstLastDate(TempDate, StartDate, EndDate)
   
   If DateToStringExtEx2(EndDate) = DateToStringExtEx2(TempDate) Then
      IsEndOfMonth = True
   Else
      IsEndOfMonth = False
   End If
End Function

Private Sub GeneratePigBalanceDocument(TempDate As Date)
Static RunNo As Long
Dim Pp As CPopulation
Dim Ivd As CInventoryDoc
Dim IsOK As Boolean
Dim II As CImportItem

   RunNo = RunNo + 1
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.INVENTORY_DOC_ID = -1
    Ivd.DOCUMENT_DATE = TempDate
   Ivd.DOCUMENT_NO = "PI-" & Format(RunNo, "00000")
   Ivd.DELIVERY_FEE = 0
   Ivd.EMP_ID = -1
   Ivd.DOCUMENT_TYPE = 11
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.SIMULATE_FLAG = "Y"
   Ivd.BATCH_ID = ID
   
   For Each Pp In m_Populations
      If Pp.CURRENT_AMOUNT > 0 Then
         Set II = New CImportItem
         II.Flag = "A"
         
         II.TX_TYPE = "I"
         II.PART_ITEM_ID = Pp.PIG_ID
         II.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         II.PIG_STATUS = -1
         II.EXPENSE_TYPE = -1
         II.IMPORT_AMOUNT = Pp.CURRENT_AMOUNT
         II.ACTUAL_UNIT_PRICE = 0
         II.CALCULATE_FLAG = "Y"
         II.TOTAL_ACTUAL_PRICE = 0
         II.TOTAL_WEIGHT = Pp.TOTAL_WEIGHT
         
         Call Ivd.ImportExports.Add(II)
         Set II = Nothing
      End If
   Next Pp
   
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   Set Ivd = Nothing
End Sub

Private Sub AccumulateRevenue(TempDate As Date)
Dim Ri As CRvnPrmItem
Dim Pp As CPopulation
Dim DateCount As Double

   For Each Ri In m_RevenueParams
      If (Ri.GetFieldValue("FROM_SALE") <= TempDate) And (TempDate <= Ri.GetFieldValue("TO_SALE")) Then
         DateCount = Abs(DateDiff("D", Ri.GetFieldValue("FROM_SALE"), Ri.GetFieldValue("TO_SALE"))) + 1
         'ยืมมาใช้เก็บ Revenue
         Set Pp = MyGetPopulation(m_RevenueAccum, Ri.GetFieldValue("REVENUE_ID"))
         If Pp Is Nothing Then
            'ยังไม่มีข้อมูล
            Set Pp = New CPopulation
            Pp.PIG_ID = Ri.GetFieldValue("REVENUE_ID")
            Pp.CURRENT_AMOUNT = MyDiffEx(Ri.GetFieldValue("SALE_AMOUNT"), DateCount)
            Pp.TOTAL_PRICE = MyDiffEx(Ri.GetFieldValue("TOTAL_PRICE"), DateCount)
            Call m_RevenueAccum.Add(Pp, Trim(Str(Pp.PIG_ID)))
            Set Pp = Nothing
         Else
            Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT + MyDiffEx(Ri.GetFieldValue("SALE_AMOUNT"), DateCount)
            Pp.TOTAL_PRICE = Pp.TOTAL_PRICE + MyDiffEx(Ri.GetFieldValue("TOTAL_PRICE"), DateCount)
         End If
      End If
   Next Ri
End Sub

Private Sub GeneratePigBalanceDoc()
Dim Ivd As CInventoryDoc
Dim II As CImportItem
Dim p As CPopulation
Dim IsOK As Boolean
Dim O As Object
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_DATE = DateAdd("D", -1, FromInSertDate)
   Ivd.DOCUMENT_TYPE = 11
   Ivd.DOCUMENT_NO = "ยกมา-" & Format(ID, "0000")
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.BATCH_ID = ID
   
   'glbErrorLog.LocalErrorMsg = "GeneratePigBalanceDoc"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   
   For Each p In m_Populations
      If p.CURRENT_AMOUNT > 0 Then
         Set II = New CImportItem
         II.Flag = "A"
         II.PART_ITEM_ID = p.PIG_ID
         II.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         II.IMPORT_AMOUNT = p.CURRENT_AMOUNT
         II.ACTUAL_UNIT_PRICE = 0
         II.CALCULATE_FLAG = "Y"
         II.TOTAL_ACTUAL_PRICE = 0
         II.TOTAL_WEIGHT = p.TOTAL_WEIGHT
         
         Set O = II
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call GeneratePartItemLocationMonthly(O)         'มีหมูจากยอดยกมาเข้ามาในระบบ * ต้องทำการเพิ่มหมูเข้าไปใน CcostColls1
         
         Call Ivd.ImportExports.Add(II)
         Set II = Nothing
      End If
      
   Next p
   
   'glbErrorLog.LocalErrorMsg = "GeneratePigBalanceDoc"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, True, glbErrorLog)
   Set Ivd = Nothing
End Sub

Private Sub GeneratePigCostBalanceDoc()
'Dim Cm As CCapitalMovement
'Dim Mi As CMovementItem
Dim p As CPopulation
Dim IsOK As Boolean

   glbDaily.StartTransaction
      
   For Each p In m_Populations
      If p.CURRENT_AMOUNT > 0 Then
         Call UpDateCostColls(p.PIG_ID, p.FEED_COST, p.EXPENSE_COST, , , , , , , p.MEDICINE_COST, , , , , p.BIRTH_COST)
         'update เฉพาะต้องทุนของหมูยกมาเข้าไปใน CcostColls1
      End If
      
      'Set Cm = Nothing
   Next p
      
   Call glbDaily.CommitTransaction
End Sub

Private Sub GeneratePigSellDocument(TempDate As Date, Optional GID As Long, Optional LID As Long, Optional BID As Long)
Dim Pi As CParamItem
Dim Pp As CPopulation
Dim Amount As Double
Dim Di As CDoItem
Dim Bd As CBillingDoc
Dim IsOK As Boolean
Dim Lc As CLocation
Dim Ivd As CInventoryDoc
Dim AvgPrice As Double
Dim CREDIT As Double
Static RunNo As Long
Dim O As Object
Dim EI As CExportItem
Dim Cs1 As CCostSearch1
Dim TempDi As CDoItem
Dim TempPi As CPartItem
         
   Set Lc = m_SaleHouses(1) 'เรือนขาย
   For Each Pi In m_PigTypeStatusCustomers
      RunNo = RunNo + 1
      
      'สร้างหัวบิลขาย
      Set Bd = New CBillingDoc
      Bd.AddEditMode = SHOW_ADD
      Bd.DOCUMENT_NO = "DO-" & Format(RunNo, "00000") & "-" & Format(ID, "0000")
      Bd.DOCUMENT_DATE = TempDate
      Bd.DOCUMENT_TYPE = 1
      Bd.DOCUMENT_SUBTYPE = 1
      Bd.ACCOUNT_ID = Pi.GetFieldValue("ACCOUNT_ID")
      Bd.EXCEPTION_FLAG = "N"
      Bd.COMMIT_FLAG = "N"
      Bd.ACCEPT_BY = -1
      Bd.BATCH_ID = ID
      Bd.SIMULATE_FLAG = "Y"
      
      For Each Pp In m_PigStatusSellItems
         If (PigTypeToCode(Pi.GetFieldValue("PIG_TYPE")) = Pp.PIG_TYPE) And _
            (Pi.GetFieldValue("PARAM_PIG_STATUS") = Pp.PIG_STATUS) Then
            
               AvgPrice = GetSalePrice(Pp, Pi.GetFieldValue("PARAM_PIG_STATUS"), CREDIT, TempDate)
              Bd.DUE_DATE = DateAdd("D", CREDIT, Bd.DOCUMENT_DATE)
               
               If Pi.GetFieldValue("SHARE_SELL_TYPE") = 1 Then
                  'ตาม %
                  Amount = Pp.CURRENT_AMOUNT * Pi.GetFieldValue("SALE_RATIO")
               ElseIf Pi.GetFieldValue("SHARE_SELL_TYPE") = 2 Then
                  'ตามกำหนดจำนวน
                  If Pp.CURRENT_AMOUNT >= Pi.GetFieldValue("SALE_RATIO") Then
                     Amount = Pi.GetFieldValue("SALE_RATIO")
                  Else
                     Amount = Pp.CURRENT_AMOUNT
                  End If
               ElseIf Pi.GetFieldValue("SHARE_SELL_TYPE") = 3 Then
                  'เอาตามจำนวนที่มี
                  Amount = Pp.CURRENT_AMOUNT
               End If
               Pp.LEFT_AMOUNT = Pp.CURRENT_AMOUNT - Amount
               'สร้าง item บิลขาย
               
               Set Di = New CDoItem
               Di.Flag = "A"
               Di.LOCATION_ID = Lc.LOCATION_ID
'               If Pp.PIG_ID = 12793 Then
'                  ''debug.print
'               End If
               Di.PART_ITEM_ID = Pp.PIG_ID
'               If Pp.PIG_NO = "254930" Then
'                  '''debug.print
'               End If
               Di.PIG_STATUS = Pp.PIG_STATUS
               Di.ITEM_AMOUNT = Amount
               Di.TOTAL_PRICE = Amount * AvgPrice * Pp.AVG_WEIGHT
               Di.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Amount
               Di.AVG_PRICE = AvgPrice
               Di.AVG_WEIGHT = Pp.AVG_WEIGHT
               
               '' เวลาขายให้ไปดึงข้อมูลจาก Collection ที่เก็บ ข้อมูลต้นที่เอาไปใส่ เพื่อ ออกรายงานได้เลย
               If Not (Pp.PIG_ID = GID Or Pp.PIG_ID = LID Or Pp.PIG_ID = BID) Then
                  'โดยต้นทุนขายที่มีพร้อมบิลขายนั้นจะมีกับหมูทุกเบอร์ยกเว้นหมู G L B
                  Set Cs1 = GetObject("CCostSearch1", CcostColls1, Trim(Str(Pp.PIG_ID)), True)
                  Di.BFOOD_AMOUNT = Cs1.BFOOD_AMOUNT * Amount
                  Di.BMEDICINE_AMOUNT = Cs1.BMEDICINE_AMOUNT * Amount
                  Di.BEXPENSE_AMOUNT = Cs1.BEXPENSE_AMOUNT * Amount
                  
                  Di.BBIRTH_AMOUNT = Cs1.BBIRTH_AMOUNT * Amount
                  Di.BIRTH_AMOUNT = Cs1.BIRTH_AMOUNT * Amount
                  
                  Di.FOOD_AMOUNT = Cs1.FOOD_AMOUNT * Amount
                  Di.MEDICINE_AMOUNT = Cs1.MEDICINE_AMOUNT * Amount
                  Di.EXPENSE_AMOUNT = Cs1.EXPENSE_AMOUNT * Amount
                  Di.OTHER_AMOUNT = Cs1.OTHER_AMOUNT * Amount
               End If
               
               Call Bd.DoItems.Add(Di)
               Set Di = Nothing
         End If
      Next Pp
      
'      If Month(TempDate) = 2 Then
'         ''debug.print
'      End If
      'Call DebugCostSellAmount(TempDate)
      
      
      If Bd.DoItems.Count > 0 Then
         Call PopulateGuiID(Bd)
         Call DO2InventoryDoc(Bd, Ivd)
         
         'glbErrorLog.LocalErrorMsg = "GeneratePigSellDocument"
         'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
         
         For Each EI In Ivd.ImportExports
            If EI.Flag = "A" Then
               Set O = EI
               O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
               O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
'               If Ei.PART_ITEM_ID = 12967 Then
'                  '''debug.print
'               End If
               'เมื่อขายหมูไปแล้วนั้นหมูจะลดจำนวนลงโดยที่ซึ่งโดยปกติแล้วยอดต้นทุนทั้งหมดของหมูแต่ละเบอร์จะลดลงแต่ต้นทุนเฉลี่ยเพิ่มขึ้น
               If O.PART_ITEM_ID = GID Or O.PART_ITEM_ID = LID Or O.PART_ITEM_ID = BID Then
                  'แต่ยกเว้นหมู G L B ซึ่งเมื่อจำนวนหมูลดลงแล้วจริงแต่ต้นต้นทุนหมูทั้งหมดต้องเท่าเดิม และ ต้นทุนเฉลี่ยเพิ่มขึ้น
                  Call UpDateCostColls(O.PART_ITEM_ID, , , , , , EI.EXPORT_AMOUNT, , , , , , , , , True)
               End If
               Call UpDateCostColls(O.PART_ITEM_ID, , , , , , EI.EXPORT_AMOUNT, , , , , , , , , , True)
               Call GeneratePartItemLocationMonthly(O, False)
            End If
         Next
         
         'glbErrorLog.LocalErrorMsg = "GeneratePigSellDocument"
         'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
      
         Ivd.BATCH_ID = ID
         Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
         
         Bd.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
         
         Call glbDaily.AddEditBillingDoc(Bd, IsOK, False, glbErrorLog)
         
         
         For Each Di In Bd.DoItems
            For Each TempPi In PigIDBirthInMonthColl
               If Di.PART_ITEM_ID = TempPi.PART_ITEM_ID Then
                  ' Ture แสดงว่าเป็นการโอนออกของ หมูที่เกิดเดือนนี้
                  Set TempDi = New CDoItem
                  TempDi.DO_ITEM_ID = Di.DO_ITEM_ID
'                  If Di.BIRTH_AMOUNT < 0 Or Di.ITEM_AMOUNT <= 0 Then
'                     '''debug.print
'                  End If
                  TempDi.BIRTH_AMOUNT = Di.BIRTH_AMOUNT
                  TempDi.ITEM_AMOUNT = Di.ITEM_AMOUNT
                  '''debug.print (TempPI.PART_NO)
                  Call DoItemBirthInMonthColl.Add(TempDi)
                  Set TempDi = Nothing
               End If
            Next TempPi
         Next Di
      End If
      Set Bd = Nothing
      Set TempPi = Nothing
      'คอมมิตเอกสารบิลขาย
   Next Pi
   
   'เคลียร์ m_PigStatusSellItems
   Set m_PigStatusSellItems = Nothing
   Set m_PigStatusSellItems = New Collection
End Sub

Private Sub GeneratePigAdjustDocument(TempDate As Date)
Dim Ivd1 As CInventoryDoc
Dim Ivd2 As CInventoryDoc
Dim II As CImportItem
Dim EI As CExportItem
Dim Tr As CTransferItem
Dim Pi As CParamItem
Dim IsOK As Boolean
Dim Et As CExpenseType
Dim Pp As CPopulation
Dim TempID As Long
Dim NeedAmount As Double
Dim Lc As CLocation
Dim Key As String
Dim TempPP As CPopulation
Static RunNo As Long
Dim O As Object
   RunNo = RunNo + 1
   
   Set Ivd1 = New CInventoryDoc
   Ivd1.AddEditMode = SHOW_ADD
   Ivd1.DOCUMENT_DATE = TempDate
   Ivd1.DUE_DATE = TempDate
   Ivd1.DOCUMENT_TYPE = 4
   Ivd1.DOCUMENT_NO = "ปรับยอด-" & Format(RunNo, "0000") & "-" & Format(ID, "0000")
   Ivd1.COMMIT_FLAG = "N"
   Ivd1.EXCEPTION_FLAG = "N"
   Ivd1.BATCH_ID = ID
      
   'glbErrorLog.LocalErrorMsg = "GeneratePigAdjustDocument"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
         
   'สร้างหัวเอกสาร
   For Each Pi In m_PigAdjustItems
      If DateToStringInt(Pi.GetFieldValue("CTRL_FROM_DATE")) <= DateToStringInt(TempDate) And _
         DateToStringInt(Pi.GetFieldValue("CTRL_TO_DATE")) >= DateToStringInt(TempDate) Then
         
         Set Pp = Nothing
'If Pi.GetFieldValue("PIG_ID") = 12749 Then
'''debug.print
'End If
         Set Pp = GetPopulation(m_Populations, Trim(Str(Pi.GetFieldValue("PIG_ID"))))
         If Pp.PIG_ID <= 0 Then
            'สร้างรายการใหม่
            Pp.Flag = "A"
            Pp.PIG_ID = Pi.GetFieldValue("PIG_ID")
            Pp.CURRENT_AMOUNT = 0
            Pp.PIG_NO = Pi.GetFieldValue("PIG_NO")
            Pp.PIG_NAME = Pi.GetFieldValue("PIG_DESC")
            Pp.PIG_TYPE = Pi.GetFieldValue("PIG_TYPE_NO2")
            Pp.CURRENT_AGE = GetAge(Pp.PIG_NO, TempDate)
            Pp.TOTAL_WEIGHT = 0
            Call m_Populations.Add(Pp, Trim(Str(Pi.GetFieldValue("PIG_ID"))))
            
         Else
'''debug.print
         End If
         
         If Pp.CURRENT_AMOUNT > Pi.GetFieldValue("CTRL_AMOUNT") Then
            'เบิกออก
            NeedAmount = Pp.CURRENT_AMOUNT - Pi.GetFieldValue("CTRL_AMOUNT")
            
            '===
            Set Tr = New CTransferItem
            Set EI = New CExportItem
            Set II = New CImportItem
            
            Tr.Flag = "A"
            EI.Flag = "A"
            EI.CALCULATE_FLAG = "N"
            II.Flag = "A"
            II.CALCULATE_FLAG = "N"
               
            Set Tr.ExportItem = EI
            Set Tr.ImportItem = II
            
            Tr.ExportItem.PART_ITEM_ID = Pp.PIG_ID
            Tr.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
            
            Tr.ExportItem.EXPORT_AMOUNT = NeedAmount
            Tr.ExportItem.PIG_STATUS = -1
            Tr.ExportItem.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Tr.ExportItem.EXPORT_AMOUNT
            
            Tr.ImportItem.PART_ITEM_ID = Pp.PIG_ID
            Tr.ImportItem.LOCATION_ID = Tr.ExportItem.LOCATION_ID
            Tr.ImportItem.IMPORT_AMOUNT = 0
            Tr.ImportItem.TOTAL_WEIGHT = 0
            Tr.ImportItem.PIG_STATUS = -1
            If Tr.ExportItem.EXPORT_AMOUNT > 0 Then
               Call Ivd1.TransferItems.Add(Tr)
               
               'Update ค่าที่ m_populations
               Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT - Tr.ExportItem.EXPORT_AMOUNT
               'Pp.TOTAL_WEIGHT = Pp.TOTAL_WEIGHT - Tr.ExportItem.TOTAL_WEIGHT
               'Pp.AVG_WEIGHT = MyDiffEx(Pp.TOTAL_WEIGHT, Pp.CURRENT_AMOUNT)
               
               Set O = EI
               O.DOCUMENT_DATE = Ivd1.DOCUMENT_DATE
               O.DOCUMENT_TYPE = Ivd1.DOCUMENT_TYPE
               Call UpDateCostColls(Pp.PIG_ID, , , , , , Tr.ExportItem.EXPORT_AMOUNT, , , , , , , , , , True)
               Call GeneratePartItemLocationMonthly(O, False) 'การปรับยอดหมูนั้นจะทำแค่จำนวนที่เพิ่มขึ้นหรือลดลงเท่านั้นแต่ว่าต้นทุนเฉลี่ยยังคงเป็นจำนวนเดิม
               
            End If
            Set Tr = Nothing
         ElseIf Pp.CURRENT_AMOUNT < Pi.GetFieldValue("CTRL_AMOUNT") Then
            'โอนเข้ามา
            NeedAmount = Pi.GetFieldValue("CTRL_AMOUNT") - Pp.CURRENT_AMOUNT
            
            '===
            Set Tr = New CTransferItem
            Set EI = New CExportItem
            Set II = New CImportItem
            
            Tr.Flag = "A"
            EI.Flag = "A"
            EI.CALCULATE_FLAG = "N"
            II.Flag = "A"
            II.CALCULATE_FLAG = "N"
            
            Set Tr.ExportItem = EI
            Set Tr.ImportItem = II
   
            Tr.ExportItem.PART_ITEM_ID = Pp.PIG_ID
            Tr.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   
            Tr.ExportItem.EXPORT_AMOUNT = 0
            Tr.ExportItem.PIG_STATUS = -1
            Tr.ExportItem.TOTAL_WEIGHT = 0
            
            Tr.ImportItem.PART_ITEM_ID = Pp.PIG_ID
            Tr.ImportItem.LOCATION_ID = Tr.ExportItem.LOCATION_ID
            Tr.ImportItem.IMPORT_AMOUNT = NeedAmount
            Tr.ImportItem.TOTAL_WEIGHT = Pp.AVG_WEIGHT * NeedAmount
            Tr.ImportItem.PIG_STATUS = -1
            If Tr.ImportItem.IMPORT_AMOUNT > 0 Then
               Call Ivd1.TransferItems.Add(Tr)
               
               'Update ค่าที่ m_populations
               Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT + Tr.ImportItem.IMPORT_AMOUNT
'               Pp.TOTAL_WEIGHT = Pp.TOTAL_WEIGHT + Tr.ImportItem.TOTAL_WEIGHT
'               Pp.AVG_WEIGHT = MyDiffEx(Pp.TOTAL_WEIGHT, Pp.CURRENT_AMOUNT)
               
               Set O = II
               O.DOCUMENT_DATE = Ivd1.DOCUMENT_DATE
               O.DOCUMENT_TYPE = Ivd1.DOCUMENT_TYPE
               Call UpDateCostColls(Pp.PIG_ID, , , , , , Tr.ImportItem.IMPORT_AMOUNT, , , , , , , , , , , True)
               Call GeneratePartItemLocationMonthly(O, False) 'การปรับยอดหมูนั้นจะทำแค่จำนวนที่เพิ่มขึ้นหรือลดลงเท่านั้นแต่ว่าต้นทุนเฉลี่ยยังคงเป็นจำนวนเดิม
            End If
            Set Tr = Nothing
         End If
      End If
   Next Pi
   
   'glbErrorLog.LocalErrorMsg = "GeneratePigAdjustDocument"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   
   Call CreateImportExportItems(Ivd1)
      
   If Ivd1.ImportExports.Count > 0 Then
      Call glbDaily.AddEditInventoryDoc(Ivd1, IsOK, False, glbErrorLog)
   End If
   
   Set Ivd1 = Nothing
   Set Ivd2 = Nothing
End Sub

Private Sub GeneratePigBuyDocument(TempDate As Date)
Dim Ivd As CInventoryDoc
Dim II As CImportItem
Dim Pi As CParamItem
Dim IsOK As Boolean
Dim Et As CExpenseType
Dim Pp As CPopulation
Dim TempID As Long
Static RunNo As Long
Dim O As Object
   For Each Et In m_ExpenseTypes
      If Et.BUY_FLAG = "Y" Then
         TempID = Et.EXPENSE_TYPE_ID
      End If
   Next Et
   
   RunNo = RunNo + 1
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_DATE = TempDate
   Ivd.DUE_DATE = TempDate
   Ivd.DOCUMENT_TYPE = 11
   Ivd.DOCUMENT_NO = "ซื้อสุกร-" & Format(RunNo, "0000") & "-" & Format(ID, "0000")
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.BATCH_ID = ID
   
   For Each Pi In m_PigBuyParams
      '''debug.print (Pi.GetFieldValue("BUY_DATE"))
      If DateToStringExtEx2(Pi.GetFieldValue("BUY_DATE")) = DateToStringExtEx2(TempDate) Then
         Set II = New CImportItem
         
         II.Flag = "A"
         II.PART_ITEM_ID = Pi.GetFieldValue("PIG_ID")
         II.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         II.IMPORT_AMOUNT = Pi.GetFieldValue("BUY_AMOUNT")
         II.ACTUAL_UNIT_PRICE = Pi.GetFieldValue("BUY_AVG_PRICE")
         II.TOTAL_ACTUAL_PRICE = Pi.GetFieldValue("BUY_TOTAL_PRICE")
         II.CALCULATE_FLAG = "Y"
         II.TOTAL_WEIGHT = Pi.GetFieldValue("BUY_AVG_WEIGHT") * II.IMPORT_AMOUNT
         II.EXPENSE_TYPE = TempID
         Call Ivd.ImportExports.Add(II)
         
         'Update ค่าที่ m_populations
Set Pp = Nothing
         Set Pp = GetPopulation(m_Populations, Trim(Str(II.PART_ITEM_ID)))
         
         If Pp.PIG_ID <= 0 Then
            'สร้างรายการใหม่
            Set Pp = New CPopulation
            Pp.Flag = "A"
            Pp.PIG_ID = Pi.GetFieldValue("PIG_ID")
'            If Pi.GetFieldValue("PIG_ID") <= 0 Then
'               '''debug.print
'            End If
            Pp.CURRENT_AMOUNT = II.IMPORT_AMOUNT
            Pp.PIG_NO = Pi.GetFieldValue("PIG_NO")
            Pp.PIG_NAME = Pi.GetFieldValue("PIG_DESC")
            Pp.PIG_TYPE = Pi.GetFieldValue("PIG_TYPE_NO2")
            Pp.CURRENT_AGE = GetAge(Pp.PIG_NO, TempDate)
            Pp.TOTAL_WEIGHT = II.TOTAL_WEIGHT
            Pp.AVG_WEIGHT = MyDiffEx(Pp.TOTAL_WEIGHT, Pp.CURRENT_AMOUNT)
            Call m_Populations.Add(Pp, Trim(Str(Pi.GetFieldValue("PIG_ID"))))
         Else
            Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT + II.IMPORT_AMOUNT
            Pp.TOTAL_WEIGHT = Pp.TOTAL_WEIGHT + II.TOTAL_WEIGHT
            Pp.AVG_WEIGHT = MyDiffEx(Pp.TOTAL_WEIGHT, Pp.CURRENT_AMOUNT)
         End If
         'เมื่อซึ้อหมูเข้ามาจะต้องมีการต้นทุนของหมูจากราคาหมูที่ซื้อมา
         Call UpDateCostColls(Pi.GetFieldValue("PIG_ID"), , , , , II.TOTAL_ACTUAL_PRICE, Pi.GetFieldValue("BUY_AMOUNT"), , True)
         Call UpDateCostColls(Pi.GetFieldValue("PIG_ID"), , , , , , Pi.GetFieldValue("BUY_AMOUNT"), , , , , , , , , , , True)
         Set O = II
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call GeneratePartItemLocationMonthly(O, False)  'พร้อมทั้งเพิ่มจำนวนของหมูที่ซื้อมาด้วย
         
         Set II = Nothing
      End If
   Next Pi
   
   If Ivd.ImportExports.Count > 0 Then
      Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   End If
   Set Ivd = Nothing
End Sub

Private Sub GenerateExpenseSharingDocument(TempDate As Date)
Dim Ivd As CBillingDoc
Dim II As CROItem
Dim Pi As CParamItem
Dim IsOK As Boolean
Dim Et As CExpenseType
Dim Pp As CPopulation
Dim TempID As Long
Dim Er As CExpenseRatio
Static RunNo As Long
Dim Ma As CMonthlyAccum
Dim ToTalAmountInHouse As Double
   RunNo = RunNo + 1
   
   Set Ivd = New CBillingDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_DATE = TempDate
   Ivd.DUE_DATE = TempDate
   Ivd.DOCUMENT_TYPE = 5
   Ivd.DOCUMENT_NO = "ปัน ค.ช.จ.-" & Format(RunNo, "0000") & "-" & Format(ID, "0000")
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   Ivd.BATCH_ID = ID
   
   For Each Pi In m_ExpenseSharing
      If DateToStringExtEx2(Pi.GetFieldValue("EXPENSE_DATE")) = DateToStringExtEx2(TempDate) Then
         Set II = New CROItem
         
         Set Er = New CExpenseRatio
         Er.Flag = "A"
         Er.SELECT_FLAG = "Y"
         Er.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         Er.RATIO = 100
         Er.RATIO_AMOUNT = Pi.GetFieldValue("EXP_TOTAL_PRICE")
         Call II.ExpenseRatios.Add(Er)
         Set Er = Nothing

         II.Flag = "A"
         II.EXPENSE_TYPE = Pi.GetFieldValue("EXPENSE_TYPE")
         II.AVG_PRICE = Pi.GetFieldValue("EXP_AVG_PRICE")
         II.TOTAL_PRICE = Pi.GetFieldValue("EXP_TOTAL_PRICE")
         II.ITEM_AMOUNT = Pi.GetFieldValue("EXP_AMOUNT")
         II.EXPENSE_DESC = Pi.GetFieldValue("EXPENSE_NAME")
         Call Ivd.RoItems.Add(II)
         
         ToTalAmountInHouse = GetPigAmount
         For Each Ma In m_PartItemsLocations
            If Ma.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)) Then
               'ปันค่าใช้จ่ายเข้าในตัวหมู
               Call UpDateCostColls(Ma.PART_ITEM_ID, , , , , MyDiffEx(Ma.BALANCE_AMOUNT2 * II.TOTAL_PRICE, ToTalAmountInHouse))
            End If
         Next Ma
         Set II = Nothing
      End If
   Next Pi
   
   If Ivd.RoItems.Count > 0 Then
      'Call GenerateExpenseMovement(Ivd, ToTalAmountInHouse)
      Call glbDaily.AddEditBillingDoc(Ivd, IsOK, False, glbErrorLog)
   End If
   Set Ivd = Nothing
End Sub
Private Function GetPigAmount() As Double
Dim Ma As CMonthlyAccum
Dim Count  As Double
   Count = 0
   For Each Ma In m_PartItemsLocations
      If Ma.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)) Then
         Count = Count + Ma.BALANCE_AMOUNT2
      End If
   Next Ma
   GetPigAmount = Count
End Function
'Private Sub IncreasePopulation(Col As Collection, Pop As CPopulation)
'Dim Pp As CPopulation
'
'   Set Pp = Nothing
'   Set Pp = GetPopulation(Col, Trim(Str(Pop.PIG_ID)))
'   If Pp.PIG_ID <= 0 Then 'หาไม่เจอ
'      Set Pp = Nothing
'      Set Pp = New CPopulation
'
'      Pp.PIG_ID = Pop.PIG_ID
'      Pp.PIG_NO = Pop.PIG_NO
'      Pp.PIG_NAME = Pop.PIG_NAME
'      Pp.PIG_TYPE = Pop.PIG_TYPE
'      Pp.CURRENT_AMOUNT = Pop.CURRENT_AMOUNT
'      Call Col.Add(Pp, Trim(Str(Pp.PIG_ID)))
'      Set Pp = Nothing
'   Else
'      Pp.AVG_WEIGHT = Pop.AVG_WEIGHT
'      Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT + Pop.CURRENT_AMOUNT
'      Pp.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Pp.CURRENT_AMOUNT
'   End If
'
'End Sub

'Private Sub DecreasePopulation(Col As Collection, Pop As CPopulation)
'Dim Pp As CPopulation
'
'   Set Pp = Nothing
'   Set Pp = GetPopulation(Col, Trim(Str(Pop.PIG_ID)))
'   If Pp.PIG_ID <= 0 Then 'หาไม่เจอ
'      Pp.PIG_ID = Pop.PIG_ID
'      Pp.PIG_NO = Pop.PIG_NO
'      Pp.PIG_NAME = Pop.PIG_NAME
'      Pp.PIG_TYPE = Pop.PIG_TYPE
'      Pp.CURRENT_AMOUNT = -1 * Pop.CURRENT_AMOUNT
'      Call Col.Add(Pp, Trim(Str(Pp.PIG_ID)))
'      Set Pp = Nothing
'   Else
'      Pp.AVG_WEIGHT = Pop.AVG_WEIGHT
'      Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT - Pop.CURRENT_AMOUNT
'      Pp.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Pp.CURRENT_AMOUNT
'   End If
'End Sub

Private Sub cmdStart_Click()
Dim TempDate As Date
Dim BirthAmount As Double
Dim Pi As CPartItem
Dim Pp As CPopulation
Dim I As Long
Dim DateCount As Long
Dim Bi As CBatchItem
Dim TrnAmount As Double
Dim AvgWeight As Double
Dim k As Long
Dim Bpi As CBrtPrmItem
Dim NewPop As CPopulation
Dim TempADG As Double
Dim lMenuChosen  As Long
Dim Menu As cPopupMenu
Dim Wr As CWeightRecord
Dim Ma As CMonthlyAccum
Dim BirthFlag As Boolean
Dim PrevBirth As String
Dim GID As Long
Dim LID As Long
Dim BID As Long
Dim PigBirthInMonth As Double
Dim TempPi As CPartItem
Dim TempWeightRecord As Collection
Dim TempWr As CWeightRecord

   Set TempWeightRecord = New Collection
   
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo) Then
      Exit Sub
   End If
   
   DateCount = Abs(DateDiff("D", FromInSertDate, ToInSertDate))
   
   'Call EnableForm(Me, False)
   
   Set Ma = New CMonthlyAccum
   Ma.FROM_DATE = FromInSertDate
   Ma.TO_DATE = ToInSertDate
   Ma.BATCH_ID = ID
   Call Ma.ClearData
   Set Ma = Nothing
   
   Call GenerateInitialPopulation               'สร้างข้อมูลหมูยกมาและข้อมูลหมูจากวันที่รันถึงวันที่รัน                                                '***************************
   'Call RefreshGrid(False)
   
   Call GenerateInitialGL    'สร้างข้อมูลยกมาช่วงอายุต่างๆของ GL และ ข้อมูล % การกลับสัตว์ตามช่วงอายุ

   
   'สร้างยอดยกมาสุกร
   Call GeneratePigBalanceDoc                                                                '***************************
   
   'สร้างต้นทุนยกมาสุกร
   Call GeneratePigCostBalanceDoc
   
   'เก็บว่าอัตราการเกิดหมูแต่ละวันเป็นเท่าใด
   Call GenerateBirthParam
   
   'Call DebugCostAmount
   
   'เก็บว่าอาหารแต่ละเบอร์มีราคาเท่าใด
   Call GenerateFeedCostParam
   
   'เก็บว่าราคาขายหมูแต่ละช่วงอายุ สถานะ มีราคาเป็นเท่าใด
   Call GenerateSaleParam
   
   'เก็บว่าจะต้องขายอื่น ๆในแต่ละช่วงเวลา ราคาเท่าไหร่ เป็นจำนวนเท่าไหร่
   Call GenerateRevenueParam
   
   'เก็บว่าจะต้องซื้อหมูในแต่ละวันเป็นจำนวนเท่าใด
   Call GeneratePigBuyParam

   'เก็บว่า % ที่จะขายหมูแต่ละสถานะ แต่ละลูกค้าเป็นเท่าใด
   Call GeneratePigStatusCustomerParam
   
   'เก็บว่าแต่ละวันจะมีปันค่าใช้จ่ายเป็นอย่างไร
   Call GenerateExpenseSharing
   
   'เก็บว่าจะต้องปรับยอดหมูเป็นอย่างไร
   Call GeneratePigAdjParam
   
   'เก็บว่าในแต่ละช่วงอายุ ประเภทหมู จะมี ADG เท่าใด
   Call GenerateADGParam
   
   Set Wr = New CWeightRecord
   Call Wr.SetFieldValue("BATCH_ID", ID)
   Call Wr.ClearData
   Set Wr = Nothing
   
   I = 0
   
   Call glbDaily.StartTransaction
   
   PigBirthInMonth = 0
   TempDate = FromInSertDate
   While TempDate <= ToInSertDate
      I = I + 1
      uctlDate1.ShowDate = TempDate
      txtPercent.Text = FormatNumber((I - 1) / DateCount * 100)
      
      'If TempDate = "7/1/2550" Then
         '''debug.print
      'End If
      
      DoEvents
      Me.Refresh
      
      BirthAmount = Round(GetBirthRate(TempDate, AvgWeight, Bpi))
      
      'หาว่าวันนี้มีการเกิดกี่ตัว
      'Call DebugCostAmount
      
      Set Pi = glbDaily.DateToPartItem(TempDate, PigCodeToID("N"))
      If (Not (Pi Is Nothing)) And (BirthAmount > 0) Then
         Set Pp = Nothing
         Set Pp = GetPopulation(m_Populations, Trim(Str(Pi.PART_ITEM_ID)))
         If Pp.PIG_ID <= 0 Then
            Call m_Populations.Add(Pp, Trim(Str(Pi.PART_ITEM_ID)))
         End If
         
         Set TempPi = GetObject("CPartItem", PigIDBirthInMonthColl, Trim(Str(Pi.PART_ITEM_ID)), False)
         If TempPi Is Nothing Then
            Call PigIDBirthInMonthColl.Add(Pi, Trim(Str(Pi.PART_ITEM_ID)))
         End If
         
         Pp.PIG_ID = Pi.PART_ITEM_ID
         Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT + BirthAmount
         Pp.PIG_NO = Pi.PART_NO
         Pp.PIG_NAME = Pi.PART_DESC
         If PrevBirth <> Pi.PART_NO Then
            Pp.AVG_WEIGHT = AvgWeight
         End If
         Pp.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Pp.CURRENT_AMOUNT
         PigBirthInMonth = PigBirthInMonth + BirthAmount
         Call GeneratePigBirthDocument(TempDate, BirthAmount, True, Pi, Pp, Bpi)                         '***************************
         
         'สร้าง record น้ำหนักหมูแรกเกิด, อายุ = 0 สัปดาห์
         Set Wr = New CWeightRecord
         Wr.ShowMode = SHOW_ADD
         Call Wr.SetFieldValue("RECORD_DATE", TempDate)
         Call Wr.SetFieldValue("PART_ITEM_ID", Pp.PIG_ID)
         Call Wr.SetFieldValue("ITEM_AMOUNT", 1)
         Call Wr.SetFieldValue("AVG_WEIGHT", Pp.AVG_WEIGHT)
         Call Wr.SetFieldValue("WEIGHT_AMOUNT", Pp.AVG_WEIGHT)
         Call Wr.SetFieldValue("PIG_AGE", 0)
         Call Wr.SetFieldValue("PIG_AGE_INT", 0)
         Call Wr.SetFieldValue("BATCH_ID", ID)
         If PrevBirth <> Pi.PART_NO Then
            Call Wr.AddEditData
         End If
         PrevBirth = Pi.PART_NO
         
         Set Wr = Nothing
      End If
      
'      '====
      Call AddDateToGL ' เพิ่มอายุให้ GL
      
      'คำนวณอายุของหมูแต่ละ week
      'สร้างใบเบิกอาหารสำหรับหมูเอาไปกิน
      k = 0
      For Each Pp In m_Populations
         k = k + 1
         If Pp.CURRENT_AMOUNT > 0 Or k = m_Populations.Count Then
'            If Pp.PIG_NO = "255001" Then
'               ''debug.print (Pp.PIG_TYPE & "-" & Pp.CURRENT_AGE)
'            End If
            Pp.CURRENT_AGE = GetAge(Pp.PIG_NO, TempDate)
   'If False Then
            If k = 1 Then
               'สร้างหัวเอกสาร
               Call GeneratePigFeedDocument(TempDate, Bi, True, Pp, 1)
            End If
            'เบิกอาหารเพื่อมาเลี้ยงหมูในแต่ละวัน
            Set Bi = GetMatchFoodParam(m_Batch.FoodItems, Pp)
            If Not (Bi Is Nothing) Then
               'สร้าง item
               Call GeneratePigFeedDocument(TempDate, Bi, True, Pp, 2)
            End If
            If k = m_Populations.Count Then
               'คอมมิต เอกสาร
               Call GeneratePigFeedDocument(TempDate, Bi, True, Pp, 3)
            End If
            
   '         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
   '            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
   '            Call VerifyToCollectionFix(Pp.PIG_ID)
   '         End If
            'Call RefreshGrid(False)
            'คำนวณน้ำหนักหมูจาก ADG ที่ตรงนี้
   '            If Pp.PIG_TYPE = "G" Then
   '               ''debug.print
   '            End If         'If Pp.PIG_NO >= "255001" Then
            If ((Pp.CURRENT_AGE > 0) Or (PrevBirth = Pp.PIG_NO)) And (Pp.CURRENT_AMOUNT > 0) Then
               
               TempADG = GetAdgRate(Pp)
               Pp.AVG_WEIGHT = Pp.AVG_WEIGHT + TempADG
               Pp.TOTAL_WEIGHT = Pp.CURRENT_AMOUNT * Pp.AVG_WEIGHT
               
               Set Wr = New CWeightRecord
               Wr.ShowMode = SHOW_ADD
               Call Wr.SetFieldValue("RECORD_DATE", TempDate)
               Call Wr.SetFieldValue("PART_ITEM_ID", Pp.PIG_ID)
               Call Wr.SetFieldValue("ITEM_AMOUNT", 1)
               Call Wr.SetFieldValue("AVG_WEIGHT", Pp.AVG_WEIGHT)
               Call Wr.SetFieldValue("WEIGHT_AMOUNT", Pp.AVG_WEIGHT)
               Call Wr.SetFieldValue("PIG_AGE", Pp.CURRENT_AGE)
               Call Wr.SetFieldValue("PIG_AGE_INT", Pp.CURRENT_AGE)
               Call Wr.SetFieldValue("PIG_AGE_INT", Fix(Pp.CURRENT_AGE))
               Call Wr.SetFieldValue("BATCH_ID", ID)
               Call Wr.SetFieldValue("ADG", TempADG)
               Call Wr.AddEditData
               
            End If
   '         End If
            
            
            'โอนสูญเสีย ขาย ไปยังเรือนขาย
            If k = 1 Then
               Call GeneratePigTransferDocument(TempDate, Bi, True, Pp, TrnAmount, 1)
            End If
            Set Bi = GetMatchFoodParam(m_Batch.TransferItems, Pp)
            If Not (Bi Is Nothing) Then
               TrnAmount = 0
               Call GeneratePigTransferDocument(TempDate, Bi, True, Pp, TrnAmount, 2)                             '***************************
               Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT - TrnAmount
               'โอนออก AvgWeight จะคงเดิม
               'คำนวณ TotalWeight ใหม่
               Pp.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Pp.CURRENT_AMOUNT
            End If
            If k = m_Populations.Count Then
               'คอมมิต เอกสาร
               Call GeneratePigTransferDocument(TempDate, Bi, True, Pp, TrnAmount, 3)
            End If
            
   '         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
   '            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
   '            Call VerifyToCollectionFix(Pp.PIG_ID)
   '         End If
            'โอนเปลี่ยนประเภทสุกร
            If k = 1 Then
               Call GeneratePigTypeChangeDocument(TempDate, Bi, True, Pp, TrnAmount, 1)
            End If
            Set Bi = GetMatchPigTypeChangeParam(m_Batch.ChangePigTypes, Pp)
            If Not (Bi Is Nothing) Then
               TrnAmount = 0
               Call GeneratePigTypeChangeDocument(TempDate, Bi, True, Pp, TrnAmount, 2)                                 '***************************
               Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT - TrnAmount
               'โอนออก AvgWeight จะคงเดิม
               'คำนวณ TotalWeight ใหม่
               Pp.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Pp.CURRENT_AMOUNT
            End If
            If k = m_Populations.Count Then
               'คอมมิต เอกสาร
               Call GeneratePigTypeChangeDocument(TempDate, Bi, True, Pp, TrnAmount, 3)
            End If
            
   '         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
   '            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
   '            Call VerifyToCollectionFix(Pp.PIG_ID)
   '         End If
            
            'โอนเปลี่ยนประเภท จาก G เป็น L
            If Pp.PIG_TYPE = "G" Then
               Call GeneratePigChangeGToL(TempDate, Pp)
               GID = Pp.PIG_ID
            ElseIf Pp.PIG_TYPE = "L" Then
               Call GeneratePigChangeLToG(TempDate, Pp)
               LID = Pp.PIG_ID
            ElseIf Pp.PIG_TYPE = "B" Then
               BID = Pp.PIG_ID
            End If
            
   '         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
   '            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
   '            Call VerifyToCollectionFix(Pp.PIG_ID)
   '         End If
   'End If
'         ElseIf Pp.CURRENT_AGE > 0 Then
'            Call m_Populations.Remove(Trim(Str(Pp.PIG_ID)))
         End If
      Next Pp
     '===
      
      'Call VerifyToCollection(TempDate)
      
'      For Each Pp In m_Populations
'         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
'            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
'         End If
'      Next Pp
      
      Call GenerateGBackToG            ' สำหรับ กรณี หมู G หรือ L กลับสัตว์จากสถานะวันที่ใดๆไปเป็นวันที่ 0 (วันผสม)
      
'      For Each Pp In m_Populations
'         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
'            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
'         End If
'      Next Pp
      
      'Call DebugCostAmount
      'Call VerifyToCollection(TempDate)
      
      'สร้างซื้อสุกรตรงนี้
      Call GeneratePigBuyDocument(TempDate)                                                              '***************************
      
      'Call VerifyToCollection(TempDate)
      
      'Call DebugCostAmount
      'สร้างเอกสารการปันค่าใช้จ่าย
      Call GenerateExpenseSharingDocument(TempDate)
'
      'Call VerifyToCollection(TempDate)
      
      'Call DebugCostAmount
      'สร้างบิลขายตรงนี้
      Call GeneratePigSellDocument(TempDate, GID, LID, BID)                                                          '***************************
      'Call VerifyToCollection(TempDate)
      
'      For Each Pp In m_Populations
'         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
'            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
'         End If
'      Next Pp
      
      'สร้างการคุมยอดหมูตรงนี้ ต้องอยู่ก่อน GeneratePigSellDocument
      Call GeneratePigAdjustDocument(TempDate)                                                        '***************************
      
'      For Each Pp In m_Populations
'         If Pp.PIG_NO = "254400" And (Pp.PIG_TYPE = "G" Or Pp.PIG_TYPE = "L") Then
'            ''debug.print (Pp.PIG_TYPE & "-" & Pp.AVG_WEIGHT)
'         End If
'      Next Pp
      
      'Call VerifyToCollection(TempDate)
      
      Call AccumulateRevenue(TempDate)
'
      'Call VerifyToCollection(TempDate)
      
      If IsEndOfMonth(TempDate) Then
         Call GenerateFeedImportDocument(TempDate, True)

         Call GenerateRevenueSaleDocument(TempDate, True)

         Call GeneratePigBirthDocumentEndMonth(GID, LID, BID, PigBirthInMonth, TempDate)
         PigBirthInMonth = 0
      End If
      
      'Call VerifyToCollection(TempDate)
      'Call RefreshGrid(False)
      
      TempDate = DateAdd("D", 1, TempDate)
   Wend
      
   Call InsertMonthlyAccum(ID)
   Call InsertInTakeFood(ID)
   
   Call glbDaily.CommitTransaction
   
   Call EnableForm(Me, True)
End Sub
Private Sub GeneratePartItemLocationMonthly(O As Object, Optional RunUpdate As Boolean = True)
Dim Key As String
Dim Key1 As String
Dim II As CMonthlyAccum
Dim Ii1 As CMonthlyAccum
Dim TempII As CMonthlyAccum
Dim TempII1 As CMonthlyAccum
Dim Amount As Double
   
   Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID & "-" & Mid(DateToStringInt(O.DOCUMENT_DATE), 1, 7)           'เป็นเดือน
   Key1 = O.PART_ITEM_ID & "-" & O.LOCATION_ID
   
   Set II = GetMonthlyAccum(m_PartItemsLocationMonthlies, Key)
   Set Ii1 = GetMonthlyAccumEx(m_PartItemsLocations, Key1)
   If II.PART_ITEM_ID <= 0 Then
      Set TempII = New CMonthlyAccum
      If Ii1.PART_ITEM_ID <= 0 Then
         Set TempII1 = New CMonthlyAccum
      End If
      If O.TX_TYPE = "I" Then
         If II.PART_ITEM_ID <= 0 Then
            TempII.BALANCE_AMOUNT1 = Ii1.BALANCE_AMOUNT2
         End If
         If O.DOCUMENT_TYPE = 5 Then 'ใบเกิด
            TempII.BIRTH_AMOUNT = O.IMPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.BIRTH_AMOUNT = TempII1.BIRTH_AMOUNT + O.IMPORT_AMOUNT
            Else
               Ii1.BIRTH_AMOUNT = Ii1.BIRTH_AMOUNT + O.IMPORT_AMOUNT
            End If
         ElseIf O.DOCUMENT_TYPE = 11 Then 'ใบนำเข้าสุกร
            TempII.BUY_AMOUNT = O.IMPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.BUY_AMOUNT = TempII1.BUY_AMOUNT + O.IMPORT_AMOUNT
            Else
               Ii1.BUY_AMOUNT = Ii1.BUY_AMOUNT + O.IMPORT_AMOUNT
            End If
         ElseIf O.DOCUMENT_TYPE = 8 Then  'ใบขึ้นทดแทน
            TempII.STATUS_IN_AMOUNT = O.IMPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.STATUS_IN_AMOUNT = TempII1.STATUS_IN_AMOUNT + O.IMPORT_AMOUNT
            Else
               Ii1.STATUS_IN_AMOUNT = Ii1.STATUS_IN_AMOUNT + O.IMPORT_AMOUNT
            End If
         ElseIf O.DOCUMENT_TYPE = 4 Then  'ใบคุมยอดเพิ่ม
            TempII.ADJUST_IN_AMOUNT = O.IMPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.ADJUST_IN_AMOUNT = TempII1.ADJUST_IN_AMOUNT + O.IMPORT_AMOUNT
            Else
               Ii1.ADJUST_IN_AMOUNT = Ii1.ADJUST_IN_AMOUNT + O.IMPORT_AMOUNT
            End If
         ElseIf O.DOCUMENT_TYPE = 888 Then   'GL_IN
            TempII.GL_IN = O.IMPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.GL_IN = TempII1.GL_IN + O.IMPORT_AMOUNT
            Else
               Ii1.GL_IN = Ii1.GL_IN + O.IMPORT_AMOUNT
            End If
         Else
            TempII.IMPORT_AMOUNT = O.IMPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.IMPORT_AMOUNT = TempII1.IMPORT_AMOUNT + O.IMPORT_AMOUNT
            Else
               Ii1.IMPORT_AMOUNT = Ii1.IMPORT_AMOUNT + O.IMPORT_AMOUNT
            End If
         End If
         Amount = O.IMPORT_AMOUNT
      ElseIf O.TX_TYPE = "E" Then
         If II.PART_ITEM_ID <= 0 Then
            TempII.BALANCE_AMOUNT1 = Ii1.BALANCE_AMOUNT2
         End If
         If O.DOCUMENT_TYPE = 8 Then  'ใบขึ้นทดแทน
            TempII.STATUS_OUT_AMOUNT = O.EXPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.STATUS_OUT_AMOUNT = TempII1.STATUS_OUT_AMOUNT + O.EXPORT_AMOUNT
            Else
               Ii1.STATUS_OUT_AMOUNT = Ii1.STATUS_OUT_AMOUNT + O.EXPORT_AMOUNT
            End If
         ElseIf O.DOCUMENT_TYPE = 7 Then  'โอนไปเรือนขาย
            TempII.SELL_AMOUNT = O.EXPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.SELL_AMOUNT = TempII1.SELL_AMOUNT + O.EXPORT_AMOUNT
            Else
               Ii1.SELL_AMOUNT = Ii1.SELL_AMOUNT + O.EXPORT_AMOUNT
            End If
         ElseIf O.DOCUMENT_TYPE = 4 Then  'ใบคุมยอดลด
            TempII.ADJUST_OUT_AMOUNT = O.EXPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.ADJUST_OUT_AMOUNT = TempII1.ADJUST_OUT_AMOUNT + O.EXPORT_AMOUNT
            Else
               Ii1.ADJUST_OUT_AMOUNT = Ii1.ADJUST_OUT_AMOUNT + O.EXPORT_AMOUNT
            End If
         ElseIf O.DOCUMENT_TYPE = 888 Then   'GL_OUT
            TempII.GL_OUT = O.EXPORT_AMOUNT
            If Ii1.PART_ITEM_ID <= 0 Then
               TempII1.GL_OUT = TempII1.GL_OUT + O.EXPORT_AMOUNT
            Else
               Ii1.GL_OUT = Ii1.GL_OUT + O.EXPORT_AMOUNT
            End If
            
         Else
            TempII.EXPORT_AMOUNT = O.EXPORT_AMOUNT
         End If
         Amount = -O.EXPORT_AMOUNT
      End If
      
      TempII.LOCATION_ID = O.LOCATION_ID
      TempII.PART_ITEM_ID = O.PART_ITEM_ID
      TempII.DOCUMENT_DATE = O.DOCUMENT_DATE
      TempII.YYYYMM = Mid(DateToStringInt(O.DOCUMENT_DATE), 1, 7)
      TempII.BALANCE_AMOUNT2 = TempII.BALANCE_AMOUNT1 + Amount
      If TempII.BALANCE_AMOUNT2 <= 0 Then
         TempII.BALANCE_AMOUNT2 = 0
      End If
'      glbErrorLog.LocalErrorMsg = Key & "-" & TempII.BALANCE_AMOUNT1 & "-" & TempII.BALANCE_AMOUNT2
'      glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
'      ''debug.print (Key & "-" & TempII.BALANCE_AMOUNT1 & "-" & TempII.BALANCE_AMOUNT2)
      
      Call m_PartItemsLocationMonthlies.Add(TempII, Key)
      
      If Ii1.PART_ITEM_ID <= 0 Then
         TempII1.LOCATION_ID = O.LOCATION_ID
         TempII1.PART_ITEM_ID = O.PART_ITEM_ID
         TempII1.BALANCE_AMOUNT1 = TempII1.BALANCE_AMOUNT2
         TempII1.BALANCE_AMOUNT2 = TempII1.BALANCE_AMOUNT2 + Amount
         
         If RunUpdate Then
            Call UpDateCostColls(O.PART_ITEM_ID, , , , , , TempII1.BALANCE_AMOUNT2, True)
         End If
         
         Call m_PartItemsLocations.Add(TempII1, Key1)
         Set TempII1 = Nothing
      Else
         Ii1.BALANCE_AMOUNT1 = Ii1.BALANCE_AMOUNT2
         Ii1.BALANCE_AMOUNT2 = Ii1.BALANCE_AMOUNT2 + Amount
         
         If RunUpdate Then
            Call UpDateCostColls(O.PART_ITEM_ID, , , , , , Ii1.BALANCE_AMOUNT2, True)
         End If
      End If
      Set TempII = Nothing
   Else
      If O.TX_TYPE = "I" Then
         If II.PART_ITEM_ID <= 0 Then
            II.BALANCE_AMOUNT1 = Ii1.BALANCE_AMOUNT2
         End If
         If O.DOCUMENT_TYPE = 5 Then 'ใบเกิด
            II.BIRTH_AMOUNT = II.BIRTH_AMOUNT + O.IMPORT_AMOUNT
            Ii1.BIRTH_AMOUNT = Ii1.BIRTH_AMOUNT + O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 11 Then 'ใบนำเข้าสุกร
            II.BUY_AMOUNT = II.BUY_AMOUNT + O.IMPORT_AMOUNT
            Ii1.BUY_AMOUNT = Ii1.BUY_AMOUNT + O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 8 Then 'ใบขึ้นทดแทน
            II.STATUS_IN_AMOUNT = II.STATUS_IN_AMOUNT + O.IMPORT_AMOUNT
            Ii1.STATUS_IN_AMOUNT = Ii1.STATUS_IN_AMOUNT + O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 4 Then 'คุมยอดเพิ่ม
            II.ADJUST_IN_AMOUNT = II.ADJUST_IN_AMOUNT + O.IMPORT_AMOUNT
            Ii1.ADJUST_IN_AMOUNT = Ii1.ADJUST_IN_AMOUNT + O.IMPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 888 Then   'GL_IN
            II.GL_IN = II.GL_IN + O.IMPORT_AMOUNT
            Ii1.GL_IN = Ii1.GL_IN + O.IMPORT_AMOUNT
         Else
            II.IMPORT_AMOUNT = II.IMPORT_AMOUNT + O.IMPORT_AMOUNT
            Ii1.IMPORT_AMOUNT = Ii1.IMPORT_AMOUNT + O.IMPORT_AMOUNT
         End If
         Amount = O.IMPORT_AMOUNT
      ElseIf O.TX_TYPE = "E" Then
         If II.PART_ITEM_ID <= 0 Then
            II.BALANCE_AMOUNT1 = Ii1.BALANCE_AMOUNT2
         End If
         If O.DOCUMENT_TYPE = 8 Then  'ใบขึ้นทดแทน
            II.STATUS_OUT_AMOUNT = II.STATUS_OUT_AMOUNT + O.EXPORT_AMOUNT
            Ii1.STATUS_OUT_AMOUNT = Ii1.STATUS_OUT_AMOUNT + O.EXPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 7 Then  'โอนไปเรือนขาย
            II.SELL_AMOUNT = II.SELL_AMOUNT + O.EXPORT_AMOUNT
            Ii1.SELL_AMOUNT = Ii1.SELL_AMOUNT + O.EXPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 4 Then  'คุมยอดเพิ่ม
            II.ADJUST_OUT_AMOUNT = II.ADJUST_OUT_AMOUNT + O.EXPORT_AMOUNT
            Ii1.ADJUST_OUT_AMOUNT = Ii1.ADJUST_OUT_AMOUNT + O.EXPORT_AMOUNT
         ElseIf O.DOCUMENT_TYPE = 888 Then   'GL_OUT
            II.GL_OUT = II.GL_OUT + O.EXPORT_AMOUNT
            Ii1.GL_OUT = Ii1.GL_OUT + O.EXPORT_AMOUNT
'''debug.print O.EXPORT_AMOUNT
         Else
            II.EXPORT_AMOUNT = II.EXPORT_AMOUNT + O.EXPORT_AMOUNT
            Ii1.EXPORT_AMOUNT = Ii1.EXPORT_AMOUNT + O.EXPORT_AMOUNT
         End If
         Amount = -O.EXPORT_AMOUNT
      End If
      II.BALANCE_AMOUNT2 = II.BALANCE_AMOUNT2 + Amount
      
      Ii1.BALANCE_AMOUNT1 = Ii1.BALANCE_AMOUNT2
      Ii1.BALANCE_AMOUNT2 = Ii1.BALANCE_AMOUNT2 + Amount
      If Ii1.BALANCE_AMOUNT2 <= 0 Then
         Ii1.BALANCE_AMOUNT2 = 0
      End If
      If RunUpdate Then
         Call UpDateCostColls(O.PART_ITEM_ID, , , , , , Ii1.BALANCE_AMOUNT2, True)
      End If
   End If
   
End Sub


Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Houses, 1)
      Set uctlLocationLookup.MyCollection = m_Houses
      
      Call LoadLocation(Nothing, m_SaleHouses, 1, "Y")
      
      Call LoadLocationByCode(Nothing, m_Locations, 2)
      
      Call LoadPartItem(Nothing, m_Feeds, , "N")
      Call LoadExpenseType(Nothing, m_ExpenseTypes, "")
      
      Call LoadProductType(Nothing, m_ProductTypes)
      
      Call LoadPartItem(Nothing, m_Pigs, , "Y", , , , 2)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Batch.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_Batch.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      'glbErrorLog.LocalErrorMsg = Me.Name
      'glbErrorLog.ShowUserError
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
   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Batch = Nothing
   Set m_ApArMass = Nothing
   Set m_Populations = Nothing
   Set m_BirthParams = Nothing
   Set m_Houses = Nothing
   Set m_TempFeed = Nothing
   Set m_Locations = Nothing
   Set m_TempTransf = Nothing
   Set m_SaleHouses = Nothing
   Set m_FeedUsed = Nothing
   Set m_SaleParams = Nothing
   Set m_RevenueParams = Nothing
   Set m_RevenueAccum = Nothing
   Set m_Feeds = Nothing
   Set m_ExpenseTypes = Nothing
   Set m_PigBuyParams = Nothing
   Set m_ProductTypes = Nothing
   Set m_Pigs = Nothing
   Set m_PigStatusSellItems = Nothing
   Set m_PigTypeStatusCustomers = Nothing
   Set m_ExpenseSharing = Nothing
   Set m_PigAdjustItems = Nothing
   Set m_Adgs = Nothing
   Set m_PartItemsLocationMonthlies = Nothing
   Set m_PartItemsLocations = Nothing
   Set CcostColls1 = Nothing
   Set m_InTakeFoods = Nothing
   Set m_GLAgecoll = Nothing
   Set m_GLBackcoll = Nothing
   Set PigIDBirthInMonthColl = Nothing
   Set DoItemBirthInMonthColl = Nothing
   Set m_ExportPerDay = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub

'Private Sub InitGrid1()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.Add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 1000
'   Col.Caption = MapText("ส.เกิด")
'
'   Set Col = GridEX1.Columns.Add '4
'   Col.Width = 1500
'   Col.Caption = MapText("ชื่อสุกร")
'
'   Set Col = GridEX1.Columns.Add '5
'   Col.Width = 700
'   Col.Caption = MapText("อายุ")
'
'   Set Col = GridEX1.Columns.Add '6
'   Col.Width = 800
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("จำนวน")
'
'   Set Col = GridEX1.Columns.Add '7
'   Col.Width = 1000
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("นนเฉลี่ย")
'
'   Set Col = GridEX1.Columns.Add '8
'   Col.Width = 1000
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("นนรวม")
'End Sub
'Private Sub InitGrid2()
'Dim Col As JSColumn
'
'   GridEX2.Columns.Clear
'   GridEX2.BackColor = GLB_GRID_COLOR
'   GridEX2.ItemCount = 0
'   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX2.ColumnHeaderFont.Bold = True
'   GridEX2.ColumnHeaderFont.Name = GLB_FONT
'   GridEX2.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX2.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX2.Columns.Add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'   Set Col = GridEX2.Columns.Add '3
'   Col.Width = 1000
'   Col.Caption = MapText("ส.เกิด")
'
'   Set Col = GridEX2.Columns.Add '4
'   Col.Width = 1500
'   Col.Caption = MapText("ชื่อสุกร")
'
'   Set Col = GridEX2.Columns.Add '5
'   Col.Width = 700
'   Col.Caption = MapText("อายุ")
'
'   Set Col = GridEX2.Columns.Add '6
'   Col.Width = 800
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("จำนวน")
'
'   Set Col = GridEX2.Columns.Add '7
'   Col.Width = 1000
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("นนเฉลี่ย")
'
'   Set Col = GridEX2.Columns.Add '8
'   Col.Width = 1000
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("นนรวม")
'End Sub
'Private Sub InitGrid3()
'Dim Col As JSColumn
'
'   GridEX3.Columns.Clear
'   GridEX3.BackColor = GLB_GRID_COLOR
'   GridEX3.ItemCount = 0
'   GridEX3.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX3.ColumnHeaderFont.Bold = True
'   GridEX3.ColumnHeaderFont.Name = GLB_FONT
'   GridEX3.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX3.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX3.Columns.Add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"
'
'   Set Col = GridEX3.Columns.Add '3
'   Col.Width = 1000
'   Col.Caption = MapText("ส.เกิด")
'
'   Set Col = GridEX3.Columns.Add '4
'   Col.Width = 1500
'   Col.Caption = MapText("ชื่อสุกร")
'
'   Set Col = GridEX3.Columns.Add '5
'   Col.Width = 700
'   Col.Caption = MapText("อายุ")
'
'   Set Col = GridEX3.Columns.Add '6
'   Col.Width = 800
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("จำนวน")
'
'   Set Col = GridEX3.Columns.Add '7
'   Col.Width = 1000
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("นนเฉลี่ย")
'
'   Set Col = GridEX3.Columns.Add '8
'   Col.Width = 1000
'   Col.TextAlignment = jgexAlignRight
'   Col.Caption = MapText("นนรวม")
'End Sub
'
Private Sub InitFormLayout()
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitNormalLabel(lblCurrentDate, MapText("วันที่"))
   Call InitNormalLabel(lblLocation, MapText("รร"))
   
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.CODE_TYPE)
   txtPercent.Enabled = False
   
   uctlDate1.Enable = False
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
'   Call InitGrid1
'   Call InitGrid2
'   Call InitGrid3
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
   Set m_Batch = New CBatch
   Set m_ApArMass = New Collection
   Set m_Populations = New Collection
   Set m_BirthParams = New Collection
   Set m_Houses = New Collection
   Set m_TempFeed = New Collection
   Set m_Locations = New Collection
   Set m_TempTransf = New Collection
   Set m_SaleHouses = New Collection
   Set m_FeedUsed = New Collection
   Set m_CostParams = New Collection
   Set m_SaleParams = New Collection
   Set m_RevenueParams = New Collection
   Set m_RevenueAccum = New Collection
   Set m_Feeds = New Collection
   Set m_ExpenseTypes = New Collection
   Set m_PigBuyParams = New Collection
   Set m_ProductTypes = New Collection
   Set m_Pigs = New Collection
   Set m_PigStatusSellItems = New Collection
   Set m_PigTypeStatusCustomers = New Collection
   Set m_ExpenseSharing = New Collection
   Set m_PigAdjustItems = New Collection
   Set m_Adgs = New Collection
   Set m_PartItemsLocationMonthlies = New Collection
   Set m_PartItemsLocations = New Collection
   Set CcostColls1 = New Collection
   Set m_InTakeFoods = New Collection
   Set m_GLAgecoll = New Collection
   Set m_GLBackcoll = New Collection
   Set PigIDBirthInMonthColl = New Collection
   Set DoItemBirthInMonthColl = New Collection
   Set m_ExportPerDay = New Collection
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   If m_Populations Is Nothing Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Dim CR As CPopulation
   Row = Row + 1
   Call GetpopFromColl(Row)
   Set CR = m_Populations(Row)
   If CR Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = CR.PIG_ID
   Values(2) = RealIndex
   Values(3) = CR.PIG_NO
   Values(4) = CR.PIG_NAME
   Values(5) = FormatNumber(CR.CURRENT_AGE)
   Values(6) = FormatNumber(CR.CURRENT_AMOUNT)
   Values(7) = FormatNumber(CR.AVG_WEIGHT)
   Values(8) = FormatNumber(CR.TOTAL_WEIGHT)
   Exit Sub
   
ErrorHandler:
   'glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   'glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GetpopFromColl(RowI)
Dim CR As CPopulation
   
   Set CR = m_Populations(RowI)
'   If CR.CURRENT_AMOUNT > 0 Then
'      Exit Sub
'   Else
'      Row = RowI
'      Row = Row + 1
'      GetpopFromColl (Row)
'   End If
End Sub
Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   If m_Populations Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   'Row = CountItemColl(m_Populations) / 3
   
   Dim CR As CPopulation
   Row = Row + 1
   Call GetpopFromColl(Row)
   Set CR = m_Populations(Row)
   If CR Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = CR.PIG_ID
   Values(2) = RealIndex
   Values(3) = CR.PIG_NO
   Values(4) = CR.PIG_NAME
   Values(5) = FormatNumber(CR.CURRENT_AGE)
   Values(6) = FormatNumber(CR.CURRENT_AMOUNT)
   Values(7) = FormatNumber(CR.AVG_WEIGHT)
   Values(8) = FormatNumber(CR.TOTAL_WEIGHT)
   Exit Sub
   
ErrorHandler:
   'glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   'glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX3_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   If m_Populations Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   'Row = (CountItemColl(m_Populations) / 3) * 2
   
   Dim CR As CPopulation
   Row = Row + 1
   Call GetpopFromColl(Row)
   Set CR = m_Populations(Row)
   If CR Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = CR.PIG_ID
   Values(2) = RealIndex
   Values(3) = CR.PIG_NO
   Values(4) = CR.PIG_NAME
   Values(5) = FormatNumber(CR.CURRENT_AGE)
   Values(6) = FormatNumber(CR.CURRENT_AMOUNT)
   Values(7) = FormatNumber(CR.AVG_WEIGHT)
   Values(8) = FormatNumber(CR.TOTAL_WEIGHT)
   Exit Sub
   
ErrorHandler:
   'glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   'glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
'Public Sub RefreshGrid(Flag As Boolean)
'
'   Row = 0
'   GridEX1.ItemCount = (CountItemColl(m_Populations) / 3)
'   'GridEX1.Rebind
'
'   Row = CountItemColl(m_Populations) / 3
'
'   GridEX2.ItemCount = (CountItemColl(m_Populations) / 3)
'   'GridEX2.Rebind
'
'   Row = (CountItemColl(m_Populations) / 3) * 2
'
'   GridEX3.ItemCount = (CountItemColl(m_Populations) - GridEX1.ItemCount - GridEX2.ItemCount)
'   'GridEX3.Rebind
'
'End Sub
'Public Function CountItemColl(Col As Collection) As Long
'Dim I As Long
'Dim Count As Long
'Dim O As CPopulation
'   Count = 0
'   For Each O In m_Populations
''      If O.CURRENT_AMOUNT > 0 Then
'         Count = Count + 1
''      End If
'   Next O
'
'   CountItemColl = Count
'End Function

Private Sub txtJournalCode_Change()
   m_HasModify = True
End Sub
Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
Private Sub InsertMonthlyAccum(BatchID As Long)
Dim Ba As CMonthlyAccum
Dim II As CMonthlyAccum
Dim iCount As Long

   For Each II In m_PartItemsLocationMonthlies
      II.AddEditMode = SHOW_ADD
      II.BATCH_ID = BatchID
      Call II.AddEditData
   Next II
End Sub
Private Sub InsertInTakeFood(BatchID As Long)
Dim Itk As CIntake
Dim iCount As Long

   For Each Itk In m_InTakeFoods
      Itk.AddEditMode = SHOW_ADD
      Itk.BATCH_ID = BatchID
      Call Itk.AddEditData
   Next Itk
End Sub
Private Sub UpDateCostColls(PigID As Long, Optional BFood As Double, Optional BExpense As Double, Optional Birth As Double, Optional Food As Double, Optional Expense As Double, Optional CurrentAmount As Double, Optional UpdateAmountFlag As Boolean = False, Optional UpdateAmountAndCostFlag As Boolean = False, Optional BMedicine As Double, Optional Medicine As Double, Optional Other As Double, Optional FromBirth As Date, Optional ToBirth As Date, Optional BBirth As Double, Optional UpdateDecreaseFlag As Boolean = False, Optional UpdateSubAmountFlag As Boolean = False, Optional UpdateAddAmountFlag As Boolean = False)
Dim Cs1 As CCostSearch1
Dim TempCs1  As CCostSearch1
Dim Key As String
   
   Key = Trim(Str(PigID))
   'Sub นี้ เป็นการคิดจำนวนและต้นทุน ***ต่อตัว*** ทุกโรงเรือน
   Set Cs1 = GetObject("CCostSearch1", CcostColls1, Key, False)
   If Cs1 Is Nothing Then
      Set TempCs1 = New CCostSearch1
      TempCs1.PIG_ID = PigID
      TempCs1.FROM_BIRTH = FromBirth
      TempCs1.TO_BIRTH = ToBirth
      If UpdateAmountFlag Then            ' Update จำนวน
         If CurrentAmount <= 0 Then
            CurrentAmount = 0
         End If
         TempCs1.CURRENT_AMOUNT = CurrentAmount
      ElseIf UpdateAmountAndCostFlag Then    'Update ราคาเฉลี่ย โดยที่จำนวนเป็นจำนวนใหม่
         TempCs1.BFOOD_AMOUNT = MyDiffEx(BFood, CurrentAmount)
         TempCs1.BMEDICINE_AMOUNT = MyDiffEx(BMedicine, CurrentAmount)
         TempCs1.BEXPENSE_AMOUNT = MyDiffEx(BExpense, CurrentAmount)
         TempCs1.BBIRTH_AMOUNT = MyDiffEx(BBirth, CurrentAmount)
         TempCs1.BIRTH_AMOUNT = MyDiffEx(Birth, CurrentAmount)
         TempCs1.FOOD_AMOUNT = MyDiffEx(Food, CurrentAmount)
         TempCs1.MEDICINE_AMOUNT = MyDiffEx(Medicine, CurrentAmount)
         TempCs1.OTHER_AMOUNT = MyDiffEx(Other, CurrentAmount)
         TempCs1.EXPENSE_AMOUNT = MyDiffEx(Expense, CurrentAmount)
         TempCs1.COST_PER_AMOUNT = TempCs1.BFOOD_AMOUNT + TempCs1.BMEDICINE_AMOUNT + TempCs1.BEXPENSE_AMOUNT + TempCs1.BBIRTH_AMOUNT + TempCs1.BIRTH_AMOUNT + TempCs1.FOOD_AMOUNT + TempCs1.MEDICINE_AMOUNT + TempCs1.EXPENSE_AMOUNT + TempCs1.OTHER_AMOUNT
      Else 'Update ราคาเฉลี่ย โดยที่จำนวนเป็นจำนวนเดิม
         TempCs1.BFOOD_AMOUNT = MyDiffEx(BFood, TempCs1.CURRENT_AMOUNT)
         TempCs1.BMEDICINE_AMOUNT = MyDiffEx(BMedicine, TempCs1.CURRENT_AMOUNT)
         TempCs1.BEXPENSE_AMOUNT = MyDiffEx(BExpense, TempCs1.CURRENT_AMOUNT)
         TempCs1.BBIRTH_AMOUNT = MyDiffEx(BBirth, TempCs1.CURRENT_AMOUNT)
         TempCs1.BIRTH_AMOUNT = MyDiffEx(Birth, TempCs1.CURRENT_AMOUNT)
         TempCs1.FOOD_AMOUNT = MyDiffEx(Food, TempCs1.CURRENT_AMOUNT)
         TempCs1.MEDICINE_AMOUNT = MyDiffEx(Medicine, TempCs1.CURRENT_AMOUNT)
         TempCs1.OTHER_AMOUNT = MyDiffEx(Other, TempCs1.CURRENT_AMOUNT)
         TempCs1.EXPENSE_AMOUNT = MyDiffEx(Expense, TempCs1.CURRENT_AMOUNT)
         TempCs1.COST_PER_AMOUNT = TempCs1.BFOOD_AMOUNT + TempCs1.BMEDICINE_AMOUNT + TempCs1.BEXPENSE_AMOUNT + TempCs1.BBIRTH_AMOUNT + TempCs1.BIRTH_AMOUNT + TempCs1.FOOD_AMOUNT + TempCs1.MEDICINE_AMOUNT + TempCs1.EXPENSE_AMOUNT + TempCs1.OTHER_AMOUNT
      End If
      TempCs1.COST_AMOUNT = TempCs1.COST_PER_AMOUNT * TempCs1.CURRENT_AMOUNT
      
'      glbErrorLog.LocalErrorMsg = TempCs1.BFOOD_AMOUNT & "-" & TempCs1.BEXPENSE_AMOUNT & "-" & TempCs1.BIRTH_AMOUNT & "-" & TempCs1.FOOD_AMOUNT & "-" & TempCs1.EXPENSE_AMOUNT & "-" & TempCs1.COST_PER_AMOUNT & "-------------->" & FormatNumber(TempCs1.COST_AMOUNT)
'      glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
'      ''debug.print TempCs1.BFOOD_AMOUNT & "-" & TempCs1.BEXPENSE_AMOUNT & "-" & TempCs1.BIRTH_AMOUNT & "-" & TempCs1.FOOD_AMOUNT & "-" & TempCs1.EXPENSE_AMOUNT & "-" & TempCs1.COST_PER_AMOUNT & "-------------->" & FormatNumber(TempCs1.COST_AMOUNT)

      Call CcostColls1.Add(TempCs1, Key)
      Set Cs1 = Nothing
   Else
      If UpdateAmountFlag Then 'Update จำนวน
         If CurrentAmount <= 0 Then
            CurrentAmount = 0
         End If
         
         Cs1.CURRENT_AMOUNT = CurrentAmount
         
      ElseIf UpdateSubAmountFlag Then 'Update จำนวนแบบหักออก
         Cs1.CURRENT_AMOUNT = Cs1.CURRENT_AMOUNT - CurrentAmount
         If Cs1.CURRENT_AMOUNT <= 0 Then
            Cs1.CURRENT_AMOUNT = 0
         End If
      ElseIf UpdateAddAmountFlag Then 'Update จำนวนแบบหักออก
         Cs1.CURRENT_AMOUNT = Cs1.CURRENT_AMOUNT + CurrentAmount
      ElseIf UpdateAmountAndCostFlag Then 'Update ต้นทุนเฉลี่ย โดยที่จำนวนเป็นจำนวนใหม่
         Cs1.BFOOD_AMOUNT = MyDiffEx(((Cs1.BFOOD_AMOUNT * Cs1.CURRENT_AMOUNT) + BFood), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.BMEDICINE_AMOUNT = MyDiffEx(((Cs1.BMEDICINE_AMOUNT * Cs1.CURRENT_AMOUNT) + BMedicine), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.BEXPENSE_AMOUNT = MyDiffEx(((Cs1.BEXPENSE_AMOUNT * Cs1.CURRENT_AMOUNT) + BExpense), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.BBIRTH_AMOUNT = MyDiffEx(((Cs1.BBIRTH_AMOUNT * Cs1.CURRENT_AMOUNT) + BBirth), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.BIRTH_AMOUNT = MyDiffEx(((Cs1.BIRTH_AMOUNT * Cs1.CURRENT_AMOUNT) + Birth), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.FOOD_AMOUNT = MyDiffEx(((Cs1.FOOD_AMOUNT * Cs1.CURRENT_AMOUNT) + Food), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.MEDICINE_AMOUNT = MyDiffEx(((Cs1.MEDICINE_AMOUNT * Cs1.CURRENT_AMOUNT) + Medicine), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.OTHER_AMOUNT = MyDiffEx(((Cs1.OTHER_AMOUNT * Cs1.CURRENT_AMOUNT) + Other), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.EXPENSE_AMOUNT = MyDiffEx(((Cs1.EXPENSE_AMOUNT * Cs1.CURRENT_AMOUNT) + Expense), Cs1.CURRENT_AMOUNT + CurrentAmount)
         Cs1.COST_PER_AMOUNT = Cs1.BFOOD_AMOUNT + Cs1.BMEDICINE_AMOUNT + Cs1.BEXPENSE_AMOUNT + Cs1.BBIRTH_AMOUNT + Cs1.BIRTH_AMOUNT + Cs1.FOOD_AMOUNT + Cs1.MEDICINE_AMOUNT + Cs1.EXPENSE_AMOUNT + Cs1.OTHER_AMOUNT
      ElseIf UpdateDecreaseFlag Then 'Update ต้นทุนเฉลี่ย โดยที่จำนวนเป็นจำนวนที่น้องกว่าเดิมซึ่งต้องการให้ต้นทุนเฉลี่ยของหมูเพิ่มขึ้น
         Cs1.BFOOD_AMOUNT = MyDiffEx(((Cs1.BFOOD_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.BMEDICINE_AMOUNT = MyDiffEx(((Cs1.BMEDICINE_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.BEXPENSE_AMOUNT = MyDiffEx(((Cs1.BEXPENSE_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.BBIRTH_AMOUNT = MyDiffEx(((Cs1.BBIRTH_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.BIRTH_AMOUNT = MyDiffEx(((Cs1.BIRTH_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.FOOD_AMOUNT = MyDiffEx(((Cs1.FOOD_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.MEDICINE_AMOUNT = MyDiffEx(((Cs1.MEDICINE_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.OTHER_AMOUNT = MyDiffEx(((Cs1.OTHER_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.EXPENSE_AMOUNT = MyDiffEx(((Cs1.EXPENSE_AMOUNT * Cs1.CURRENT_AMOUNT)), Cs1.CURRENT_AMOUNT - CurrentAmount)
         Cs1.COST_PER_AMOUNT = Cs1.BFOOD_AMOUNT + Cs1.BMEDICINE_AMOUNT + Cs1.BEXPENSE_AMOUNT + Cs1.BBIRTH_AMOUNT + Cs1.BIRTH_AMOUNT + Cs1.FOOD_AMOUNT + Cs1.MEDICINE_AMOUNT + Cs1.EXPENSE_AMOUNT + Cs1.OTHER_AMOUNT
      Else 'Update ต้นทุนเฉลี่ย โดยที่จำนวนเป็นจำนวนเดิม + ต้นทุนใหม่ที่เพิ่มขึ้น
         Cs1.BFOOD_AMOUNT = Cs1.BFOOD_AMOUNT + MyDiffEx(BFood, Cs1.CURRENT_AMOUNT)
         Cs1.BMEDICINE_AMOUNT = Cs1.BMEDICINE_AMOUNT + MyDiffEx(BMedicine, Cs1.CURRENT_AMOUNT)
         Cs1.BEXPENSE_AMOUNT = Cs1.BEXPENSE_AMOUNT + MyDiffEx(BExpense, Cs1.CURRENT_AMOUNT)
         Cs1.BIRTH_AMOUNT = Cs1.BIRTH_AMOUNT + MyDiffEx(Birth, Cs1.CURRENT_AMOUNT)
         Cs1.BBIRTH_AMOUNT = Cs1.BBIRTH_AMOUNT + MyDiffEx(BBirth, Cs1.CURRENT_AMOUNT)
         Cs1.FOOD_AMOUNT = Cs1.FOOD_AMOUNT + MyDiffEx(Food, Cs1.CURRENT_AMOUNT)
         Cs1.MEDICINE_AMOUNT = Cs1.MEDICINE_AMOUNT + MyDiffEx(Medicine, Cs1.CURRENT_AMOUNT)
         Cs1.OTHER_AMOUNT = Cs1.OTHER_AMOUNT + MyDiffEx(Other, Cs1.CURRENT_AMOUNT)
         Cs1.EXPENSE_AMOUNT = Cs1.EXPENSE_AMOUNT + MyDiffEx(Expense, Cs1.CURRENT_AMOUNT)
         Cs1.COST_PER_AMOUNT = Cs1.BFOOD_AMOUNT + Cs1.BMEDICINE_AMOUNT + Cs1.BEXPENSE_AMOUNT + Cs1.BBIRTH_AMOUNT + Cs1.BIRTH_AMOUNT + Cs1.FOOD_AMOUNT + Cs1.MEDICINE_AMOUNT + Cs1.EXPENSE_AMOUNT + Cs1.OTHER_AMOUNT
      End If
      Cs1.COST_AMOUNT = Cs1.COST_PER_AMOUNT * Cs1.CURRENT_AMOUNT
            
   End If
      
End Sub
Public Sub DebugCostAmount()
Dim Cs1 As CCostSearch1
Dim Sum1  As Double
   Sum1 = 0
   For Each Cs1 In CcostColls1
      'Sum1 = Sum1 + (Cs1.BIRTH_AMOUNT * Cs1.CURRENT_AMOUNT)
      Sum1 = Sum1 + (Cs1.COST_AMOUNT)
   Next Cs1
'   glbErrorLog.LocalErrorMsg = "///////  " & FormatNumber(Sum1) & "    //////////////////////////////"
'   glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
'   ''debug.print "///////  " & FormatNumber(Sum1) & "    //////////////////////////////"

End Sub
Public Sub DebugCostSellAmount(TempDate As Date)
   glbErrorLog.LocalErrorMsg = "ต้นทุนยอดยกมาอาหารเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(BFodd(Month(TempDate)))
   glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   ''debug.print "ต้นทุนยอดยกมาอาหารเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(BFodd(Month(TempDate)))
   
   glbErrorLog.LocalErrorMsg = "ต้นทุนยอดยกมาค่าใช้จ่ายเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(BExp(Month(TempDate)))
   glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   ''debug.print "ต้นทุนยอดยกมาค่าใช้จ่ายเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(BExp(Month(TempDate)))
   
   glbErrorLog.LocalErrorMsg = "ต้นทุนเกิดเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(Birth(Month(TempDate)))
   glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   ''debug.print "ต้นทุนเกิดเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(Birth(Month(TempDate)))
   
   glbErrorLog.LocalErrorMsg = "ต้นทุนอาหารเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(Food(Month(TempDate)))
   glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   ''debug.print "ต้นทุนอาหารเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(Food(Month(TempDate)))
   
   glbErrorLog.LocalErrorMsg = "ต้นทุนค่าใช้จ่ายเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(Expense(Month(TempDate)))
   glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   ''debug.print "ต้นทุนค่าใช้จ่ายเดือน    " & Month(TempDate) & " ------->   " & FormatNumber(Expense(Month(TempDate)))
End Sub
'Public Sub GenerateExpenseMovement(Ivd As CBillingDoc, ToTalAmountInHouse As Double)
'Dim Cm As CCapitalMovement
'Dim Ci As CMovementItem
'Dim Ro As CROItem
'Dim Ma As CMonthlyAccum
'
'   For Each Ro In Ivd.RoItems
'      For Each Ma In m_PartItemsLocations
'         If Ma.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)) Then
'            Set Cm = New CCapitalMovement
'            Cm.AddEditMode = SHOW_ADD
'            Cm.COMMIT_FLAG = "N"
'            Cm.DOCUMENT_NO = Ivd.DOCUMENT_NO
'            Cm.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
'            Cm.BL_ID = Ivd.BILLING_DOC_ID
'            Cm.DOCUMENT_CATEGORY = 2
'            Cm.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
'            Cm.TX_TYPE = "I"
'            Cm.TX_AMOUNT = 0
'            Cm.TO_PIG_COUNT = ToTalAmountInHouse
'            Cm.FROM_HOUSE_ID = 0
'            Cm.TO_HOUSE_ID = 0
'            Cm.PIG_ID = Ma.PART_ITEM_ID
'            Cm.PIG_STATUS = 0
'            Cm.TX_SEQ = 0
'            Cm.REPLACE_FLAG = "N"
'            Cm.BATCH_ID = Ivd.BATCH_ID
'            Call Cm.AddEditData
'
'            Set Mi = New CMovementItem
'            Mi.AddEditMode = SHOW_ADD
'            Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
'            Mi.PART_ITEM_ID = 0
'            Mi.EXPENSE_TYPE = Ro.EXPENSE_TYPE
'            Mi.CAPITAL_AMOUNT = MyDiffEx(Ma.BALANCE_AMOUNT2 * Ro.TOTAL_PRICE, ToTalAmountInHouse)
'            Call Mi.AddEditData
'            Set Mi = Nothing
'            Set Cm = Nothing
'         End If
'      Next Ma
'   Next Ro
'End Sub
'Public Sub GenerateFoodMovement(Ivd As CInventoryDoc)
'Dim Ei  As CExportItem
'Dim Cm As CCapitalMovement
'Dim Ci As CMovementItem
'   For Each Ei In Ivd.ImportExports
'      Set Cm = New CCapitalMovement
'      Cm.AddEditMode = SHOW_ADD
'      Cm.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
'      Cm.DOCUMENT_NO = Ivd.DOCUMENT_NO
'      Cm.DOCUMENT_CATEGORY = 1
'      Cm.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
'      Cm.COMMIT_FLAG = "N"
'      Cm.BATCH_ID = ID
'      Cm.PIG_ID = Ei.PIG_ID
'      Cm.FROM_HOUSE_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
'      Call Cm.AddEditData
'
'      Set Mi = New CMovementItem
'      Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
'      Mi.AddEditMode = SHOW_ADD
'      Mi.PART_ITEM_ID = Ei.PART_ITEM_ID
'      Mi.EXPENSE_TYPE = -1
'      Mi.CAPITAL_AMOUNT = Ei.EXPORT_TOTAL_PRICE
'      Call Mi.AddEditData
'      Set Mi = Nothing
'      Set Cm = Nothing
'
'   Next Ei
'End Sub
'Public Sub GenerateTranferMovement(Ivd As CInventoryDoc)
'Dim Ei  As CExportItem
'Dim Cm As CCapitalMovement
'Dim Ci As CMovementItem
'   For Each Ei In Ivd.ImportExports
'      Set Cm = New CCapitalMovement
'      Cm.AddEditMode = SHOW_ADD
'      Cm.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
'      Cm.DOCUMENT_NO = Ivd.DOCUMENT_NO
'      Cm.DOCUMENT_CATEGORY = 1
'      Cm.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
'      Cm.COMMIT_FLAG = "N"
'      Cm.BATCH_ID = ID
'      Cm.PIG_ID = Ei.PIG_ID
'      Cm.FROM_HOUSE_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
'      Call Cm.AddEditData
'
'      Set Mi = New CMovementItem
'      Mi.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
'      Mi.AddEditMode = SHOW_ADD
'      Mi.PART_ITEM_ID = Ei.PART_ITEM_ID
'      Mi.EXPENSE_TYPE = -1
'      If Ei.TX_TYPE = "I" Then
'         Mi.CAPITAL_AMOUNT = Ei.EXPORT_TOTAL_PRICE
'      ElseIf Ei.TX_TYPE = "E" Then
'         Mi.CAPITAL_AMOUNT = -Ei.EXPORT_TOTAL_PRICE
'      End If
'
'      Call Mi.AddEditData
'      Set Mi = Nothing
'      Set Cm = Nothing
'
'   Next Ei
'End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   'GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 550
'   GridEX2.HEIGHT = GridEX1.HEIGHT
'   GridEX3.HEIGHT = GridEX1.HEIGHT
'   GridEX1.Width = ScaleWidth / 3
'   GridEX2.Width = ScaleWidth / 3
'   GridEX3.Width = ScaleWidth / 3
'   GridEX1.Left = 0
'   GridEX2.Left = GridEX1.Left + GridEX1.Width
'   GridEX3.Left = GridEX2.Left + GridEX1.Width
End Sub
Private Sub GeneratePigTypeChangeDocument(TempDate As Date, Bi As CBatchItem, Flag As Boolean, Pp As CPopulation, TransferAmount As Double, Mode As Long)
Dim Ui As CParamItem
Dim Pm As CParameter
Dim ParamID As Long
Dim Pi As CPartItem
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim Tr As CTransferItem
Dim EI As CExportItem
Dim II As CImportItem
Dim Lc As CLocation
Dim TrnAmount As Double
Static Ivd As CInventoryDoc
Static RunNo As Long
Dim O  As Object
Dim TempPP As CPopulation
Dim Gl As CGLAgeAmount
   
   If Not Flag Then
      Exit Sub
   End If
   
   If (Pp.CURRENT_AMOUNT <= 0) And (Mode = 2) Then
      Exit Sub
   End If
   
   If Mode = 1 Then
      RunNo = RunNo + 1
      Set Ivd = New CInventoryDoc
      Ivd.AddEditMode = SHOW_ADD
      Ivd.INVENTORY_DOC_ID = -1
       Ivd.DOCUMENT_DATE = TempDate
      Ivd.DOCUMENT_NO = "PTC-" & Format(RunNo, "00000") & "-" & Format(ID, "0000")
      Ivd.DELIVERY_FEE = 0
      Ivd.EMP_ID = -1
      Ivd.DOCUMENT_TYPE = 8
      Ivd.COMMIT_FLAG = "N"
      Ivd.SALE_FLAG = "N"
      Ivd.EXCEPTION_FLAG = "N"
      Ivd.SIMULATE_FLAG = "Y"
      Ivd.BATCH_ID = ID
   ElseIf Mode = 2 Then
      TrnAmount = 0
      RunNo = RunNo + 1
      Set TempRs = New ADODB.Recordset

      ParamID = Bi.GetFieldValue("PARAM_ID")
      Set Pm = Nothing
      Set Pm = MyGetParameter(m_TempTransf, Trim(Str(ParamID)))
      If Pm Is Nothing Then
         Set Pm = New CParameter
      End If

      If Pm.GetFieldValue("PARAM_ID") <= 0 Then
         'ไม่พบให้คิวรี่มาแล้วใส่เข้าไปใน m_TempTransf
         Call Pm.SetFieldValue("PARAM_ID", ParamID)
         Pm.QueryFlag = 1
         Call glbDaily.QueryParameter(Pm, TempRs, iCount, IsOK, glbErrorLog)
         If Not TempRs.EOF Then
            Call Pm.PopulateFromRS(1, TempRs)
            Call m_TempTransf.Add(Pm, Trim(Str(ParamID)))
         End If
      End If

      For Each Ui In Pm.PigStatusChangeItems
         Set Tr = New CTransferItem
         Set EI = New CExportItem
         Set II = New CImportItem

         Tr.Flag = "A"
         EI.Flag = "A"
         EI.CALCULATE_FLAG = "N"
         II.Flag = "A"
         II.CALCULATE_FLAG = "N"

         Set Tr.ExportItem = EI
         Set Tr.ImportItem = II
         
         Tr.ExportItem.PART_ITEM_ID = Pp.PIG_ID
         Tr.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         Tr.ExportItem.EXPORT_AMOUNT = Pp.CURRENT_AMOUNT * Ui.GetFieldValue("TRANSFER_RATE") / 100  'Ui.GetFieldValue("TRANSFER_RATE") 'ตอนนี้ transfer_rate เป็นจำนวน% ' (Ui.GetFieldValue("TRANSFER_RATE") * Pp.CURRENT_AMOUNT / 100) 'Round(Ui.GetFieldValue("TRANSFER_RATE") * Pp.CURRENT_AMOUNT / 100)
         TrnAmount = TrnAmount + Tr.ExportItem.EXPORT_AMOUNT
         Tr.ExportItem.PIG_STATUS = -1
         Tr.ExportItem.TOTAL_WEIGHT = Pp.AVG_WEIGHT * Tr.ExportItem.EXPORT_AMOUNT
         
         'If Pp.PIG_TYPE = "R" Then
            '''debug.print
         'End If
         
         If PigTypeToCode(Ui.GetFieldValue("PIG_TYPE")) = "G" Then 'ถ้าเป็นหมู R จะบังคับเปลี่ยน G 254400
            Set Pi = GetPartItem(m_Pigs, "254400" & "-" & PigTypeToCode(Ui.GetFieldValue("PIG_TYPE")))
         Else
            Set Pi = GetPartItem(m_Pigs, Pp.PIG_NO & "-" & PigTypeToCode(Ui.GetFieldValue("PIG_TYPE")))
         End If
         
         Tr.ImportItem.PART_ITEM_ID = Pi.PART_ITEM_ID
         Tr.ImportItem.LOCATION_ID = Tr.ExportItem.LOCATION_ID        'Location เดียวกับ Export
         Tr.ImportItem.IMPORT_AMOUNT = Tr.ExportItem.EXPORT_AMOUNT
         Tr.ImportItem.TOTAL_WEIGHT = Tr.ExportItem.TOTAL_WEIGHT
         Tr.ImportItem.PIG_STATUS = Tr.ExportItem.PIG_STATUS
                  
         Set TempPP = GetPopulationEx(m_Populations, Trim(Str(Pi.PART_ITEM_ID)))
         If TempPP Is Nothing Then
            Set TempPP = New CPopulation
         End If
         If TempPP.PIG_ID <= 0 Then
            Call m_Populations.Add(TempPP, Trim(Str(Pi.PART_ITEM_ID)))
         End If
         TempPP.PIG_ID = Pi.PART_ITEM_ID
         TempPP.CURRENT_AMOUNT = TempPP.CURRENT_AMOUNT + Tr.ExportItem.EXPORT_AMOUNT
         TempPP.PIG_NO = Pi.PART_NO
         TempPP.PIG_NAME = Pi.PART_DESC
         TempPP.PIG_TYPE = PigTypeToCode(Ui.GetFieldValue("PIG_TYPE"))
          
          'R ที่เปลี่ยนเป็น G นั้น น้ำหนักแค่ ประมาณ 140 แต่ ว่า น้ำหนัก G จะประมาณ 180
          'ดังนั้น ถ้าเปลี่ยน จาก R เป็น G จะให้ นน เท่าเดิม
          If Not (PigTypeToCode(Ui.GetFieldValue("PIG_TYPE")) = "G") Then
            TempPP.TOTAL_WEIGHT = TempPP.TOTAL_WEIGHT + (Pp.AVG_WEIGHT * Tr.ExportItem.EXPORT_AMOUNT) 'ทำให้ นน ของหมู เปลี่ยนแปลง
            TempPP.AVG_WEIGHT = MyDiffEx(TempPP.TOTAL_WEIGHT, TempPP.CURRENT_AMOUNT)
         End If
         Set TempPP = Nothing
         
         If PigTypeToCode(Ui.GetFieldValue("PIG_TYPE")) = "G" Then
            Call AdditionGlAmount("G", 0, Tr.ExportItem.EXPORT_AMOUNT)
         End If
         
         Set O = Tr.ExportItem
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call UpDateCostColls(O.PART_ITEM_ID, , , , , , EI.EXPORT_AMOUNT, , , , , , , , , , True)
         Call GeneratePartItemLocationMonthly(O, False)      'Export จำนวน โอนออกจะต้องลดลง โดยที่จะต้องนำต้นทุนต่างๆมาใส่ที่ partitem ที่ Import เข้ามา
         
         'Get ต้นทุนของ O.PART_ITEM_ID
         '
         Dim Cs1 As CCostSearch1
         Set Cs1 = GetObject("CCostSearch1", CcostColls1, Trim(Str(O.PART_ITEM_ID)), True)
         
         Set O = Tr.ImportItem
         'update ต้องทุนของหมูที่ Import เข้ามา
         Call UpDateCostColls(O.PART_ITEM_ID, Cs1.BFOOD_AMOUNT * O.IMPORT_AMOUNT, Cs1.BEXPENSE_AMOUNT * O.IMPORT_AMOUNT, Cs1.BIRTH_AMOUNT * O.IMPORT_AMOUNT, Cs1.FOOD_AMOUNT * O.IMPORT_AMOUNT, Cs1.EXPENSE_AMOUNT * O.IMPORT_AMOUNT, O.IMPORT_AMOUNT, , True, Cs1.BMEDICINE_AMOUNT * O.IMPORT_AMOUNT, Cs1.MEDICINE_AMOUNT * O.IMPORT_AMOUNT, Cs1.OTHER_AMOUNT * O.IMPORT_AMOUNT, , , Cs1.BBIRTH_AMOUNT * O.IMPORT_AMOUNT)
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call UpDateCostColls(O.PART_ITEM_ID, , , , , , II.IMPORT_AMOUNT, , , , , , , , , , , True)
         Call GeneratePartItemLocationMonthly(O, False)     'พร้อมกันนั้นต้องทำการเพิ่มจำนวนของหมูที่ Import ด้วย
         
         If Tr.ExportItem.EXPORT_AMOUNT > 0 Then
            Call Ivd.TransferItems.Add(Tr)
         End If
         
         Set Tr = Nothing
      Next Ui
      
      If TempRs.State = adStateOpen Then
         Call TempRs.Close
      End If
      Set TempRs = Nothing
      
      TransferAmount = TrnAmount
   ElseIf Mode = 3 Then
      Call CreateImportExportItems(Ivd)
      
      If Ivd.ImportExports.Count > 0 Then
         Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
      End If
   End If
End Sub
Private Sub GenerateInitialGL()
Dim Gl As CGLAgeAmount
Dim Gb As CGLBackAmount
Dim Bi As CBatchItem
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim TempDate As Date
Dim I As Long

   Set TempRs = New ADODB.Recordset
   
   'glbErrorLog.LocalErrorMsg = "GenerateInitialGL"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   For Each Bi In m_Batch.Glages
      Set Gl = New CGLAgeAmount
      Gl.PARAM_ID = Bi.GetFieldValue("PARAM_ID")
      Call Gl.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Gl = New CGLAgeAmount
         Call Gl.PopulateFromRS(1, TempRs)
         Call m_GLAgecoll.Add(Gl)
         TempRs.MoveNext
         Set Gl = Nothing
      Wend
      
      Set Gl = Nothing
   Next Bi
   
   For Each Bi In m_Batch.GLbacks
      Set Gb = New CGLBackAmount
      Gb.PARAM_ID = Bi.GetFieldValue("PARAM_ID")
      Call Gb.QueryData(1, TempRs, iCount)
      
      While Not TempRs.EOF
         Set Gb = New CGLBackAmount
         Call Gb.PopulateFromRS(1, TempRs)
         Call m_GLBackcoll.Add(Gb)
         TempRs.MoveNext
         Set Gb = Nothing
      Wend
      
      Set Gb = Nothing
   Next Bi
   
   'glbErrorLog.LocalErrorMsg = "GenerateInitialPopulation"
   'glbErrorLog.ShowErrorLog (LOG_TO_FILE_EX)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub
Private Sub AddDateToGL()
Dim Gl As CGLAgeAmount
   For Each Gl In m_GLAgecoll
      Gl.GL_AGE = Gl.GL_AGE + 1
   Next
End Sub
Private Sub GeneratePigChangeGToL(TempDate As Date, Pp As CPopulation)
Dim Gl As CGLAgeAmount
Dim Ivd As CInventoryDoc
Static RunNo As Long
Dim EI  As CExportItem
Dim II   As CImportItem
Dim Pi As CPartItem
Dim TempPP As CPopulation
Dim O As Object
Dim IsOK As Boolean
Dim PigType As Long
Dim TempGl   As CGLAgeAmount
   For Each Gl In m_GLAgecoll
      If Gl.PIG_TYPE_NAME = "G" And Gl.GL_AGE = 115 Then
         RunNo = RunNo + 1
         Set Ivd = New CInventoryDoc
         Ivd.AddEditMode = SHOW_ADD
         Ivd.INVENTORY_DOC_ID = -1
          Ivd.DOCUMENT_DATE = TempDate
         Ivd.DOCUMENT_NO = "GTL-" & Format(RunNo, "00000") & "-" & Format(ID, "0000")
         Ivd.DELIVERY_FEE = 0
         Ivd.EMP_ID = -1
         Ivd.DOCUMENT_TYPE = 888             ' ประเภทเอกสารพิเศษสำหรับ G เป็น L
         Ivd.COMMIT_FLAG = "N"
         Ivd.SALE_FLAG = "N"
         Ivd.EXCEPTION_FLAG = "N"
         Ivd.SIMULATE_FLAG = "Y"
         Ivd.BATCH_ID = ID
                  
                  
         Set EI = New CExportItem
         Set II = New CImportItem

         EI.Flag = "A"
         EI.CALCULATE_FLAG = "N"
         II.Flag = "A"
         II.CALCULATE_FLAG = "N"
         
         EI.PART_ITEM_ID = Pp.PIG_ID
         EI.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         EI.EXPORT_AMOUNT = Gl.GL_AMOUNT
         Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT - Gl.GL_AMOUNT
         EI.PIG_STATUS = -1
         EI.TOTAL_WEIGHT = Pp.AVG_WEIGHT * EI.EXPORT_AMOUNT
                  
         Set Pi = GetPartItem(m_Pigs, "254400" & "-" & "L")
         
         II.PART_ITEM_ID = Pi.PART_ITEM_ID
         II.LOCATION_ID = EI.LOCATION_ID
         II.IMPORT_AMOUNT = EI.EXPORT_AMOUNT
         II.TOTAL_WEIGHT = EI.TOTAL_WEIGHT
         II.PIG_STATUS = EI.PIG_STATUS
                  
         Set TempPP = GetPopulationEx(m_Populations, Trim(Str(Pi.PART_ITEM_ID)))
         If TempPP Is Nothing Then
            Set TempPP = New CPopulation
         End If
         If TempPP.PIG_ID <= 0 Then
            Call m_Populations.Add(TempPP, Trim(Str(Pi.PART_ITEM_ID)))
         End If
         TempPP.PIG_ID = Pi.PART_ITEM_ID
         TempPP.CURRENT_AMOUNT = TempPP.CURRENT_AMOUNT + EI.EXPORT_AMOUNT
         TempPP.PIG_NO = Pi.PART_NO
         TempPP.PIG_NAME = Pi.PART_DESC
         TempPP.PIG_TYPE = "L"
'         TempPP.TOTAL_WEIGHT = TempPP.TOTAL_WEIGHT + (Pp.AVG_WEIGHT * Ei.EXPORT_AMOUNT)
'         TempPP.AVG_WEIGHT = MyDiffEx(TempPP.TOTAL_WEIGHT, TempPP.CURRENT_AMOUNT)
'         มีผลให้ นน ของ L เปลี่ยนแปลงเนื่องจากนำ นน ของ G มา เฉลี่ยด้วย แต่ เรา จะไม่คิด
         Set TempPP = Nothing
         
         Call AdditionGlAmount("L", 0, EI.EXPORT_AMOUNT)
         
         Set O = EI
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call UpDateCostColls(O.PART_ITEM_ID, , , , , , EI.EXPORT_AMOUNT, , , , , , , , , , True)
         Call GeneratePartItemLocationMonthly(O, False)        'ลดจำนวนของ G โดยที่ต้นทุนเฉลี่ยยังเหมือนเดิม
         
         'Get ต้นทุนของ O.PART_ITEM_ID
         '
         Dim Cs1 As CCostSearch1
         Set Cs1 = GetObject("CCostSearch1", CcostColls1, Trim(Str(O.PART_ITEM_ID)), True)
         'Get ต้นทุนของ หมูที่ G
         Set O = II
         Call UpDateCostColls(O.PART_ITEM_ID, Cs1.BFOOD_AMOUNT * O.IMPORT_AMOUNT, Cs1.BEXPENSE_AMOUNT * O.IMPORT_AMOUNT, Cs1.BIRTH_AMOUNT * O.IMPORT_AMOUNT, Cs1.FOOD_AMOUNT * O.IMPORT_AMOUNT, Cs1.EXPENSE_AMOUNT * O.IMPORT_AMOUNT, O.IMPORT_AMOUNT, , True, Cs1.BMEDICINE_AMOUNT * O.IMPORT_AMOUNT, Cs1.MEDICINE_AMOUNT * O.IMPORT_AMOUNT, Cs1.OTHER_AMOUNT * O.IMPORT_AMOUNT, , , Cs1.BBIRTH_AMOUNT * O.IMPORT_AMOUNT)
         Call UpDateCostColls(O.PART_ITEM_ID, , , , , , II.IMPORT_AMOUNT, , , , , , , , , , , True)
         'แล้ว นำมาใส่ในต้องทุนของหมู L
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call GeneratePartItemLocationMonthly(O, False) 'พร้อมทั้งเพิ่มจำนวนของหมู L ด้วย
         
         If EI.EXPORT_AMOUNT > 0 Then
            Call Ivd.ImportExports.Add(EI)
            Call Ivd.ImportExports.Add(II)
         End If
         
         If Ivd.ImportExports.Count > 0 Then
            Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
         End If
         
         Call RemoveGlAmount("G", 115)
      End If
      
   Next Gl
   
   
End Sub
Private Sub GeneratePigChangeLToG(TempDate As Date, Pp As CPopulation)
Dim Gl As CGLAgeAmount
Dim Ivd As CInventoryDoc
Static RunNo As Long
Dim EI  As CExportItem
Dim II   As CImportItem
Dim Pi As CPartItem
Dim TempPP As CPopulation
Dim O As Object
Dim IsOK As Boolean
Dim PigType As Long
Dim TempGl As CGLAgeAmount
   
   For Each Gl In m_GLAgecoll
      If Gl.PIG_TYPE_NAME = "L" And Gl.GL_AGE = 24 Then    'วันที่ 23 ยังเป็น L อยู่
         RunNo = RunNo + 1
         Set Ivd = New CInventoryDoc
         Ivd.AddEditMode = SHOW_ADD
         Ivd.INVENTORY_DOC_ID = -1
          Ivd.DOCUMENT_DATE = TempDate
         Ivd.DOCUMENT_NO = "LTG-" & Format(RunNo, "00000") & "-" & Format(ID, "0000")
         Ivd.DELIVERY_FEE = 0
         Ivd.EMP_ID = -1
         Ivd.DOCUMENT_TYPE = 888             ' ประเภทเอกสารพิเศษสำหรับ G เป็น L
         Ivd.COMMIT_FLAG = "N"
         Ivd.SALE_FLAG = "N"
         Ivd.EXCEPTION_FLAG = "N"
         Ivd.SIMULATE_FLAG = "Y"
         Ivd.BATCH_ID = ID
                  
         Set EI = New CExportItem
         Set II = New CImportItem

         EI.Flag = "A"
         EI.CALCULATE_FLAG = "N"
         II.Flag = "A"
         II.CALCULATE_FLAG = "N"
         
         EI.PART_ITEM_ID = Pp.PIG_ID
         EI.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         EI.EXPORT_AMOUNT = Gl.GL_AMOUNT
         Pp.CURRENT_AMOUNT = Pp.CURRENT_AMOUNT - Gl.GL_AMOUNT
         EI.PIG_STATUS = -1
         EI.TOTAL_WEIGHT = Pp.AVG_WEIGHT * EI.EXPORT_AMOUNT
         
         Set Pi = GetPartItem(m_Pigs, "254400" & "-" & "G")
         
         II.PART_ITEM_ID = Pi.PART_ITEM_ID
         II.LOCATION_ID = EI.LOCATION_ID
         II.IMPORT_AMOUNT = EI.EXPORT_AMOUNT
         II.TOTAL_WEIGHT = EI.TOTAL_WEIGHT
         II.PIG_STATUS = EI.PIG_STATUS
                  
         Set TempPP = GetPopulationEx(m_Populations, Trim(Str(Pi.PART_ITEM_ID)))
         If TempPP Is Nothing Then
            Set TempPP = New CPopulation
         End If
         If TempPP.PIG_ID <= 0 Then
            Call m_Populations.Add(TempPP, Trim(Str(Pi.PART_ITEM_ID)))
         End If
         TempPP.PIG_ID = Pi.PART_ITEM_ID
         TempPP.CURRENT_AMOUNT = TempPP.CURRENT_AMOUNT + EI.EXPORT_AMOUNT
         TempPP.PIG_NO = Pi.PART_NO
         TempPP.PIG_NAME = Pi.PART_DESC
         TempPP.PIG_TYPE = "G"
'         TempPP.TOTAL_WEIGHT = TempPP.TOTAL_WEIGHT + (Pp.AVG_WEIGHT * Ei.EXPORT_AMOUNT)
'         TempPP.AVG_WEIGHT = MyDiffEx(TempPP.TOTAL_WEIGHT, TempPP.CURRENT_AMOUNT)
'         มีผลให้ นน ของ L เปลี่ยนแปลงเนื่องจากนำ นน ของ G มา เฉลี่ยด้วย แต่ เรา จะไม่คิด
         
         Set TempPP = Nothing
         
         Call AdditionGlAmount("G", 0, EI.EXPORT_AMOUNT)
         
         'Call VerifyToCollectionFix(EI.PART_ITEM_ID)
         
         Set O = EI
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call UpDateCostColls(O.PART_ITEM_ID, , , , , , EI.EXPORT_AMOUNT, , , , , , , , , , True)
         Call GeneratePartItemLocationMonthly(O, False)        '
         
         Dim Cs1 As CCostSearch1
         Set Cs1 = GetObject("CCostSearch1", CcostColls1, Trim(Str(O.PART_ITEM_ID)), True)
         'Get ต้นทุนของ L
         Set O = II
         Call UpDateCostColls(O.PART_ITEM_ID, Cs1.BFOOD_AMOUNT * O.IMPORT_AMOUNT, Cs1.BEXPENSE_AMOUNT * O.IMPORT_AMOUNT, Cs1.BIRTH_AMOUNT * O.IMPORT_AMOUNT, Cs1.FOOD_AMOUNT * O.IMPORT_AMOUNT, Cs1.EXPENSE_AMOUNT * O.IMPORT_AMOUNT, O.IMPORT_AMOUNT, , True, Cs1.BMEDICINE_AMOUNT * O.IMPORT_AMOUNT, Cs1.MEDICINE_AMOUNT * O.IMPORT_AMOUNT, Cs1.OTHER_AMOUNT * O.IMPORT_AMOUNT, , , Cs1.BBIRTH_AMOUNT * O.IMPORT_AMOUNT)
         Call UpDateCostColls(O.PART_ITEM_ID, , , , , , II.IMPORT_AMOUNT, , , , , , , , , , , True)
         'นำต้นทุนของหมู L มาเพิ่มที่หมู G
         O.DOCUMENT_DATE = Ivd.DOCUMENT_DATE
         O.DOCUMENT_TYPE = Ivd.DOCUMENT_TYPE
         Call GeneratePartItemLocationMonthly(O, False)     'พร้อมทั้งเพิ่มจำนวนของหมู G ด้วย
         
         If EI.EXPORT_AMOUNT > 0 Then
            Call Ivd.ImportExports.Add(EI)
            Call Ivd.ImportExports.Add(II)
         End If
         
         If Ivd.ImportExports.Count > 0 Then
            Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
         End If
         
         Call RemoveGlAmount("L", 24)
      End If
      
   Next Gl
   
End Sub
Private Sub AdditionGlAmount(PigType As String, AGE As Long, Amount As Double)
Dim GLa  As CGLAgeAmount
   For Each GLa In m_GLAgecoll
      If GLa.PIG_TYPE_NAME = PigType And GLa.GL_AGE = AGE Then
         GLa.GL_AMOUNT = GLa.GL_AMOUNT + Amount
         Exit Sub
      End If
   Next GLa
   
   Set GLa = New CGLAgeAmount
   GLa.GL_AGE = AGE
   GLa.PIG_TYPE_NAME = PigType
   GLa.GL_AMOUNT = Amount
   Call m_GLAgecoll.Add(GLa)
   Set GLa = Nothing
   
End Sub
Private Sub RemoveGlAmount(PigType As String, AGE As Long)
Dim GLa  As CGLAgeAmount
Dim I As Long
   I = 0
   For Each GLa In m_GLAgecoll
       I = I + 1
      If GLa.PIG_TYPE_NAME = PigType And GLa.GL_AGE = AGE Then
         Call m_GLAgecoll.Remove(I)
         Exit Sub
      End If
   Next GLa
End Sub
Private Sub GenerateGBackToG()
Dim GLa As CGLAgeAmount
Dim GLb As CGLBackAmount
   For Each GLa In m_GLAgecoll
      For Each GLb In m_GLBackcoll
         If (GLa.GL_AGE = GLb.GL_AGE) And (GLa.PIG_TYPE_NAME = GLb.PIG_TYPE_NAME) Then
            Call AdditionGlAmount(GLa.PIG_TYPE_NAME, 0, MyDiffEx(GLa.GL_AMOUNT * GLb.GL_AMOUNT, 100))
            GLa.GL_AMOUNT = GLa.GL_AMOUNT - MyDiffEx(GLa.GL_AMOUNT * GLb.GL_AMOUNT, 100)
         End If
      Next GLb
   Next GLa
End Sub
Private Sub GeneratePigBirthDocumentEndMonth(GID As Long, LID As Long, BID As Long, PigBrithInMonth As Double, TempDate As Date)
Dim TempCs1  As CCostSearch1
Dim CostAmount As Double
Dim PigBirthCost As Double
Dim FromDate As Date
Dim ToDate  As Date
Dim Bi As CBrtPrmItem
   For Each TempCs1 In CcostColls1
      If TempCs1.PIG_ID = GID Or TempCs1.PIG_ID = LID Or TempCs1.PIG_ID = BID Then   ' พ่อแม่เมื่อปันให้ลูกแล้ว จะเซ็ตต้นทุนของตัวเองเป็น 0
         CostAmount = CostAmount + (TempCs1.COST_PER_AMOUNT * TempCs1.CURRENT_AMOUNT)
         TempCs1.BFOOD_AMOUNT = 0
         TempCs1.BMEDICINE_AMOUNT = 0
         TempCs1.BEXPENSE_AMOUNT = 0
         TempCs1.BBIRTH_AMOUNT = 0
         TempCs1.BIRTH_AMOUNT = 0
         TempCs1.FOOD_AMOUNT = 0
         TempCs1.MEDICINE_AMOUNT = 0
         TempCs1.OTHER_AMOUNT = 0
         TempCs1.EXPENSE_AMOUNT = 0
         TempCs1.COST_PER_AMOUNT = 0
      End If
   Next TempCs1
   PigBirthCost = MyDiffEx(CostAmount, PigBrithInMonth)        'ต้นทุนเกิด = อาหารที่พ่อแม่กินในเดือน/หมูที่เกิดในเดือน
   '''debug.print (PigBirthCost)
   Call GetFirstLastDate(TempDate, FromDate, ToDate)
   For Each TempCs1 In CcostColls1
      If TempCs1.FROM_BIRTH > 0 And TempCs1.TO_BIRTH > 0 Then
         If TempCs1.FROM_BIRTH >= FromDate And TempCs1.TO_BIRTH <= ToDate Then
            Call UpDateCostColls(TempCs1.PIG_ID, , , PigBirthCost * TempCs1.CURRENT_AMOUNT)
            
            For Each Bi In m_BirthParams
               If Bi.GetFieldValue("FROM_BIRTH") = TempCs1.FROM_BIRTH And Bi.GetFieldValue("TO_BIRTH") = TempCs1.TO_BIRTH Then
                  Call Bi.SetFieldValue("BIRTH_COST", PigBirthCost)
                  Call Bi.UpdatePigBirthCost
               End If
            Next Bi
         ElseIf TempCs1.FROM_BIRTH >= FromDate And TempCs1.FROM_BIRTH <= ToDate Then
            Call UpDateCostColls(TempCs1.PIG_ID, , , MyDiff(PigBirthCost * TempCs1.CURRENT_AMOUNT * (DateDiff("D", TempCs1.FROM_BIRTH, ToDate) + 1), 7))
            For Each Bi In m_BirthParams
               If Bi.GetFieldValue("FROM_BIRTH") = TempCs1.FROM_BIRTH And Bi.GetFieldValue("TO_BIRTH") = TempCs1.TO_BIRTH Then
                  Call Bi.SetFieldValue("BIRTH_COST", Bi.GetFieldValue("BIRTH_COST") + MyDiff(PigBirthCost * (DateDiff("D", TempCs1.FROM_BIRTH, ToDate) + 1), 7))
                  Call Bi.UpdatePigBirthCost
               End If
            Next Bi
         ElseIf TempCs1.TO_BIRTH >= FromDate And TempCs1.TO_BIRTH <= ToDate Then
            Call UpDateCostColls(TempCs1.PIG_ID, , , MyDiffEx(PigBirthCost * TempCs1.CURRENT_AMOUNT * (DateDiff("D", FromDate, TempCs1.TO_BIRTH) + 1), 7))
            For Each Bi In m_BirthParams
               If Bi.GetFieldValue("FROM_BIRTH") = TempCs1.FROM_BIRTH And Bi.GetFieldValue("TO_BIRTH") = TempCs1.TO_BIRTH Then
                  Call Bi.SetFieldValue("BIRTH_COST", Bi.GetFieldValue("BIRTH_COST") + MyDiffEx(PigBirthCost * (DateDiff("D", FromDate, TempCs1.TO_BIRTH) + 1), 7))
                  Call Bi.UpdatePigBirthCost
               End If
            Next Bi
         End If
      End If
   Next TempCs1
   Call UpdateBirthAmountToDoItem(FromDate, ToDate, PigBirthCost)    ' Update ต้นทุนเกิดไปยัง DoItem
   
   ' เซ็ตค่าใหม่หลังจากกระจ่ายต้นทุนลูกเกิดแล้ว
   Set PigIDBirthInMonthColl = Nothing
   Set PigIDBirthInMonthColl = New Collection
   Set DoItemBirthInMonthColl = Nothing
   Set DoItemBirthInMonthColl = New Collection
End Sub
Private Sub UpdateBirthAmountToDoItem(FromDate As Date, ToDate As Date, PigBirthCost As Double)
Dim Di As CDoItem
   
   For Each Di In DoItemBirthInMonthColl
      Di.BIRTH_AMOUNT = Di.BIRTH_AMOUNT + (Di.ITEM_AMOUNT * PigBirthCost)
      Call Di.UpdateBirthCost
   Next Di
End Sub
'Private Sub VerifyToCollection(TempDate As Date)
'Dim Pp As CPopulation
'Dim Ii1 As CMonthlyAccum
'Dim Cs1 As CCostSearch1
'Dim Amount1 As Double
'Dim Amount2 As Double
'Dim Key1 As String
'Dim Key2 As String
'   For Each Pp In m_Populations
'      Key1 = Trim(Pp.PIG_ID & "-" & "492")
'      Set Ii1 = GetMonthlyAccumEx(m_PartItemsLocations, Key1)
'      Amount1 = Ii1.BALANCE_AMOUNT2
'      Key1 = Trim(Pp.PIG_ID & "-" & "500")
'      Set Ii1 = GetMonthlyAccumEx(m_PartItemsLocations, Key1)
'      Amount1 = Amount1 + Ii1.BALANCE_AMOUNT2
'
'      Key2 = Trim(Str(Pp.PIG_ID))
'      Set Cs1 = GetObject("CCostSearch1", CcostColls1, Key2, True)
'      Amount2 = Cs1.CURRENT_AMOUNT
'
'      If FormatNumber(Amount1) <> FormatNumber(Amount2) Then
'         glbErrorLog.LocalErrorMsg = Pp.PIG_NO & "-" & Pp.PIG_TYPE & " " & Amount1 & " <> " & Amount2 & " วันที่ " & TempDate
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         '''debug.print Pp.PIG_NO & "-" & Pp.PIG_TYPE & " " & Amount1 & " <> " & Amount2 & " วันที่ " & TempDate
'      End If
'   Next Pp
'
'End Sub
'Private Sub VerifyToCollectionFix(PigID As Long)
'Dim Ii1 As CMonthlyAccum
'Dim Cs1 As CCostSearch1
'Dim Amount1 As Double
'Dim Amount2 As Double
'Dim Key1 As String
'Dim Key2 As String
'
'   Key1 = Trim(PigID & "-" & "492")
'   Set Ii1 = GetMonthlyAccumEx(m_PartItemsLocations, Key1)
'   Amount1 = Ii1.BALANCE_AMOUNT2
'   Key1 = Trim(PigID & "-" & "500")
'   Set Ii1 = GetMonthlyAccumEx(m_PartItemsLocations, Key1)
'   Amount1 = Amount1 + Ii1.BALANCE_AMOUNT2
'
'   Key2 = Trim(Str(PigID))
'   Set Cs1 = GetObject("CCostSearch1", CcostColls1, Key2, True)
'   Amount2 = Cs1.CURRENT_AMOUNT
'
'   If FormatNumber(Amount1) <> FormatNumber(Amount2) Then
'      glbErrorLog.LocalErrorMsg = Amount1 & " <> " & Amount2
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      ''''debug.print Amount1 & " <> " & Amount2
'   End If
'
'End Sub
