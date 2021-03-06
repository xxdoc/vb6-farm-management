VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory048"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory048"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ExpenseTypes As Collection
Private m_PartGroup As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_ExpenseTypes = New Collection
   Set m_PartGroup = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_ExpenseTypes = Nothing
   Set m_PartGroup = Nothing
      
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Details2 = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.Add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim I As Long
   
   For I = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next I

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim Sgi As CSGroupItem
Dim TempSgi As CSGroupItem
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField
   
   Set Cf = New CReportField
                  
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ʶҹ�", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѻ�����Դ", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�ӹǹ", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "���", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�鹷ع���", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "GP", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ʶҹ�", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѻ�����Դ", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�.�./���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "���˹ѡ", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "�Ҥ�/���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Ҥ�/�.�.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�鹷ع/���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�鹷ع/�.�.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť�� �.�.", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "GP/���", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "GP/�.�.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť�� GP", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "% GP", ">")
   Call m_Headers1.AddField(Cf)

   Set Cf = Nothing
      
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
   
   Call LoadPartGroup(Nothing, m_PartGroup)
   Call LoadExpenseType(Nothing, m_ExpenseTypes)
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprFanfoldUS 'pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "�ҡ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " �֧ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim O As Object
Dim RName As String
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Lc As CProductStatus
Dim Cm As CCapitalMovement
Dim TempID As Long
Dim LeftAmount As Long
Dim CapitalMovements As Collection
Dim RevenueItemCosts As Collection
Dim Sum1 As Double
Dim TempKey As String
Dim ImportPigs As Collection
Dim ExportPigs As Collection
Dim PigStatuss As Collection
Dim Mi As CMovementItem
Dim Profit As Double
Dim Partitems As Collection
Dim Pi As CPartItem
Dim PrevID As Long
Dim RevenueCostAmount As Double
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set PigStatuss = New Collection
   Call LoadProductStatus(Nothing, PigStatuss)
   
   Set Partitems = New Collection
   Call LoadPartItem(Nothing, Partitems, , "Y")
   
   Set CapitalMovements = New Collection
   Call LoadTotalCapitalSellBalancePigStatus(Nothing, CapitalMovements, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , mcolParam("HOUSE_ID"), mcolParam("STATUS_ID"))
   Set Rs = New ADODB.Recordset
   
   Set RevenueItemCosts = New Collection
   Call LoadSumRevenueItemCost(Nothing, RevenueItemCosts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField

      Set Cm = New CCapitalMovement
      Cm.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
      Cm.CAPITAL_MOVEMENT_ID = -1
      Cm.FROM_DATE = mcolParam("FROM_DATE")
      Cm.TO_DATE = mcolParam("TO_DATE")
      Cm.PIG_STATUS = mcolParam("STATUS_ID")
      Cm.FROM_HOUSE_ID = mcolParam("HOUSE_ID")
      Cm.DocTypeSet = "(10, 13)"
      Cm.DOCUMENT_CATEGORY = 1
      Cm.EXCLUDE_STATUS_ID = mcolParam("GROUP_STATUS_ID")
      Cm.PIG_FLAG = "Y"
      Cm.OrderBy = 1 'mcolParam("ORDER_BY")
      If mcolParam("ORDER_TYPE") <= 0 Then
         Cm.OrderType = 1
      Else
         Cm.OrderType = mcolParam("ORDER_TYPE")
      End If
      Call Cm.QueryData(9, Rs, iCount)
      If Not Rs.EOF Then
         PrevID = NVLI(Rs("PIG_STATUS"), 0)
      End If
      
      While Not Rs.EOF
         Call Cm.PopulateFromRS(9, Rs)
         If PrevID <> Cm.PIG_STATUS Then
            '����ѹ����
            Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            PrevID = Cm.PIG_STATUS
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
         End If
         
         Set Mi = GetMovementItem(CapitalMovements, Cm.PIG_ID & "-" & Cm.PIG_STATUS)
         Profit = Cm.TOTAL_PRICE - Cm.DISCOUNT_AMOUNT - Mi.CAPITAL_AMOUNT

         I = I + 1
'         For j = 1 To UBound(Total1)
'            Total2(j) = 0
'         Next j
'If I = 56 Then
'''debug.print
'End If

            j = 0
            Sum1 = 0
            Call m_Details1.ClearField
            For Each HeadCf In m_Headers1.Fields
               j = j + 1
               If j = 1 Then
                  TempStr = I
               ElseIf j = 2 Then
                  If Cm.PIG_STATUS > 0 Then
                     Set Lc = PigStatuss(Trim(Str(Cm.PIG_STATUS)))
                     TempStr = Lc.PRODUCT_STATUS_NAME
                  Else
                     TempStr = ""
                  End If
               ElseIf j = 3 Then
                  If Cm.PIG_ID > 0 Then
                     Set Pi = GetPartItem(Partitems, Trim(Str(Cm.PIG_ID)))
                     TempStr = Pi.PART_NO & " (" & Pi.PIG_TYPE & ")"
                  Else
                     TempStr = ""
                  End If
               ElseIf j = 4 Then
                  TempStr = FormatNumber(Cm.TX_AMOUNT)
                  Total1(j) = Total1(j) + Cm.TX_AMOUNT
                  Total2(j) = Total2(j) + Cm.TX_AMOUNT
               ElseIf j = 5 Then
                  TempStr = FormatNumber(MyDiff(Cm.TOTAL_WEIGHT, Cm.TX_AMOUNT))
               ElseIf j = 6 Then
                  TempStr = FormatNumber(Cm.TOTAL_WEIGHT)
                  Total1(j) = Total1(j) + Cm.TOTAL_WEIGHT
                  Total2(j) = Total2(j) + Cm.TOTAL_WEIGHT
               ElseIf j = 7 Then
                  TempStr = FormatNumber(MyDiff(Cm.TOTAL_PRICE - Cm.DISCOUNT_AMOUNT, Cm.TX_AMOUNT))
               ElseIf j = 8 Then
                  TempStr = FormatNumber(MyDiff(Cm.TOTAL_PRICE - Cm.DISCOUNT_AMOUNT, Cm.TOTAL_WEIGHT))
               ElseIf j = 9 Then
                  TempStr = FormatNumber(Cm.TOTAL_PRICE - Cm.DISCOUNT_AMOUNT)
                  Total1(j) = Total1(j) + Cm.TOTAL_PRICE - Cm.DISCOUNT_AMOUNT
                  Total2(j) = Total2(j) + Cm.TOTAL_PRICE - Cm.DISCOUNT_AMOUNT
               ElseIf j = 10 Then
                  TempStr = FormatNumber(MyDiff(Mi.CAPITAL_AMOUNT, Cm.TX_AMOUNT))
                  If Mi.CAPITAL_AMOUNT <= 0 Then
                     RevenueCostAmount = 0
                     RevenueCostAmount = GetRevenueCostItemStatus(RevenueItemCosts, Cm.PIG_STATUS)
                     TempStr = FormatNumber(MyDiff(RevenueCostAmount, Cm.TX_AMOUNT))
                  End If
               ElseIf j = 11 Then
                  TempStr = FormatNumber(MyDiff(Mi.CAPITAL_AMOUNT, Cm.TOTAL_WEIGHT))
                  If Mi.CAPITAL_AMOUNT <= 0 Then
                     TempStr = FormatNumber(MyDiff(RevenueCostAmount, Cm.TOTAL_WEIGHT))
                  End If
               ElseIf j = 12 Then
                  TempStr = FormatNumber(Mi.CAPITAL_AMOUNT)
                  Total1(j) = Total1(j) + Mi.CAPITAL_AMOUNT
                  Total2(j) = Total2(j) + Mi.CAPITAL_AMOUNT
                  If Mi.CAPITAL_AMOUNT <= 0 Then
                     TempStr = FormatNumber(RevenueCostAmount)
                     Total1(j) = Total1(j) + RevenueCostAmount
                     Total2(j) = Total2(j) + RevenueCostAmount
                  End If
               ElseIf j = 13 Then
                  TempStr = FormatNumber(MyDiff(Profit, Cm.TX_AMOUNT))
               ElseIf j = 14 Then
                  TempStr = FormatNumber(MyDiff(Profit, Cm.TOTAL_WEIGHT))
               ElseIf j = 15 Then
                  TempStr = FormatNumber(Profit)
                  Total1(j) = Total1(j) + Profit
                  Total2(j) = Total2(j) + Profit
               ElseIf j = 16 Then
                  TempStr = FormatNumber(MyDiff(Profit, Cm.TOTAL_PRICE) * 100)
               Else
                  TempStr = ""
               End If

               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            Next HeadCf

            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)

            Rs.MoveNext
         Wend
         
         
         
         If iCount > 0 Then
            Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
         
         Call GenerateFooter("�ط��", "", "", Total2, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   Set CapitalMovements = Nothing
   Set PigStatuss = Nothing
   Set Partitems = Nothing
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set RevenueItemCosts = Nothing
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = ""
      ElseIf j = 5 Then
         TempStr = FormatNumber(MyDiffEx(Tot(6), Tot(4)))
      ElseIf j = 7 Then
         TempStr = FormatNumber(MyDiffEx(Tot(9), Tot(4)))
      ElseIf j = 8 Then
         TempStr = FormatNumber(MyDiffEx(Tot(9), Tot(6)))
      ElseIf j = 10 Then
         TempStr = FormatNumber(MyDiffEx(Tot(12), Tot(4)))
      ElseIf j = 11 Then
         TempStr = FormatNumber(MyDiffEx(Tot(12), Tot(6)))
      ElseIf j = 13 Then
         TempStr = FormatNumber(MyDiffEx(Tot(15), Tot(4)))
      ElseIf j = 14 Then
         TempStr = FormatNumber(MyDiffEx(Tot(15), Tot(6)))
      ElseIf j = 16 Then
         TempStr = FormatNumber(MyDiffEx(Tot(15), Tot(9)) * 100)
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|version|: " & glbParameterObj.Version)
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
Public Function GetRevenueCostItemStatus(Cl As Collection, PigStatus As Long) As Double
Dim Sum1  As Double
Dim Rci As CRevenueCostItem
   Sum1 = 0
   For Each Rci In Cl
      If Rci.GetFieldValue("PIG_STATUS") = PigStatus Then
         Sum1 = Sum1 + Rci.GetFieldValue("REVENUE_COST_ITEM_AMOUNT")
      End If
   Next Rci
   GetRevenueCostItemStatus = Sum1
End Function



