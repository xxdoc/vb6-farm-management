VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBudget002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBudget002"
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
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double

Private m_TotalExpenseCashs As Collection
Private m_FeedTotalCash As Collection
Private m_BuyTotalCash As Collection
Private m_CashPigTotalPrice As Collection
Private m_CashReveuneTotalPrice As Collection
Private m_PartTypes As Collection
Private m_TotalSellCapitals As Collection
Private m_TotalLossCapitals As Collection
Private m_MenagementExpenses  As Collection
Private m_Year As Long
Private m_FromDate As Date
Private m_ToDate As Date

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_TotalExpenseCashs = New Collection
   Set m_FeedTotalCash = New Collection
   Set m_CashPigTotalPrice = New Collection
   Set m_PartTypes = New Collection
   Set m_TotalSellCapitals = New Collection
   Set m_TotalLossCapitals = New Collection
   Set m_CashReveuneTotalPrice = New Collection
   Set m_MenagementExpenses = New Collection
   Set m_BuyTotalCash = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_TotalExpenseCashs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_FeedTotalCash = Nothing
   Set m_CashPigTotalPrice = Nothing
   Set m_PartTypes = Nothing
   Set m_TotalSellCapitals = Nothing
   Set m_TotalLossCapitals = Nothing
   Set m_CashReveuneTotalPrice = Nothing
   Set m_MenagementExpenses = Nothing
   Set m_BuyTotalCash = Nothing
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
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long
Dim YearFmt As String

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   '===
   m_Year = Val(mcolParam("YEAR_NO"))
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "รายการ", "<")
   Call m_Headers1.AddField(Cf)
   For j = 1 To 12
      YearFmt = Val(Format(m_Year, "0000")) - 543 & "-" & Format(j, "00")
      Call Cf.SetFieldValue(2, "^", IntToThaiMonthEx(j) & " " & m_Year, ">", YearFmt)
      Call m_Headers1.AddField(Cf)
   Next j
   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing

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
Dim ToPaidDate As String
Dim NewDate As Date

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   m_FromDate = DateSerial(mcolParam("YEAR_NO") - 543, 1, 1)
   m_ToDate = DateSerial(mcolParam("YEAR_NO") - 543, 12, LastDayOfMonth(m_FromDate))

   Call LoadBuyFeedIntakeTypeYYYYMM2(Nothing, m_FeedTotalCash, m_FromDate, m_ToDate, , mcolParam("BATCH_ID"), 1)      'รายจ่ายค่าอาหารแยกตาม INTAKE_TYPE ด้วย (อาหาร ยา+วัคซีน อื่นๆ)
   Call LoadBuyFeedYYYYMM2(Nothing, m_BuyTotalCash, m_FromDate, m_ToDate, , mcolParam("BATCH_ID"), 11)      'รายจ่ายค่าซื้อสุกร
   
   Call LoadTotalCashSellPig(Nothing, m_CashPigTotalPrice, m_FromDate, m_ToDate, , , mcolParam("BATCH_ID"))          ' รายได้จากการขายหมู
   Call LoadTotalRevenueCash(Nothing, m_CashReveuneTotalPrice, m_FromDate, m_ToDate, , , mcolParam("BATCH_ID"))   ' รายได้อื่นๆ
   Call LoadTotalExpenseAmountYYYYMM2(Nothing, m_TotalExpenseCashs, m_FromDate, m_ToDate, mcolParam("BATCH_ID"), "N") 'รายจ่ายอื่นๆ
'   Call LoadTotalSellCapitalYYYYMM(Nothing, m_TotalSellCapitals, m_FromDate, m_ToDate)
'   Call LoadTotalLossCapitalYYYYMM(Nothing, m_TotalLossCapitals, m_FromDate, m_ToDate)
   Call LoadSumManagementExpense(Nothing, m_MenagementExpenses, m_FromDate, m_ToDate, mcolParam("BATCH_ID"), "N")
   
   FromDate = "ปี " & mcolParam("YEAR_NO")

   MODULE_DESC = mcolParam("BATCH_NAME") & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate

   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
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
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim TempApt As CDoItem
Dim Hc As CReportField
Dim Di1 As CDoItem
Dim Di2 As CDoItem
Dim Ro As CImportItem
Dim Mi As CMovementItem
Dim Li As CLossItem
Dim PrevCustName As String
Dim PrevCustCode As String
Dim PrevCustID As Long
Dim PrevCustType As Long
Dim PrevCustTypeName As String
Dim SubHeaderFlag As Boolean
Dim Et As CPartType
Dim II As CImportItem
Dim Ri As CROItem
Dim SumLine1 As Double

Dim ParamItems As CParamItem
Dim m_Batch As CBatch
Dim CashBalance As Double
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------

   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset

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
   
   Set m_Batch = New CBatch
   Call m_Batch.SetFieldValue("BATCH_ID", mcolParam("BATCH_ID"))
   If Not glbDaily.QueryBatch(m_Batch, Rs, iCount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
      
   If iCount > 0 Then
      Call m_Batch.PopulateFromRS(1, Rs)
   End If
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Set BodyCf = New CReportField
      
      'เงินจากการขายหมู
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = 1
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินจากการขายสุกร"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set Di1 = GetDoItem(m_CashPigTotalPrice, Hc.TempKey)
         TempStr = FormatNumber(Di1.TOTAL_PRICE, , , True)
         Total1(I + 2) = Total1(I + 2) + Di1.TOTAL_PRICE
         SumLine1 = SumLine1 + Di1.TOTAL_PRICE
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total1(15) = Total1(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินจากการขายหมู
      
      'เงินจากการขายอื่น ๆ
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = 2
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินจากการขายอื่น ๆ"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set Di1 = GetDoItem(m_CashReveuneTotalPrice, Hc.TempKey)
         TempStr = FormatNumber(Di1.TOTAL_PRICE, , , True)
         Total1(I + 2) = Total1(I + 2) + Di1.TOTAL_PRICE
         SumLine1 = SumLine1 + Di1.TOTAL_PRICE
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total1(15) = Total1(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินจากการขายอื่น ๆ
            
      'รวมเงินสดรับ
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = "รวม"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = ""
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
         TempStr = FormatNumber(Total1(I + 2), , , True)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      TempStr = FormatNumber(Total1(15), , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      'รวมเงินสดรับ
            
      'เงินสดจ่ายค่าอาหาร
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = 3.1
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินจ่ายค่าอาหาร"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set II = GetImportItem(m_FeedTotalCash, Hc.TempKey & "-" & 1)
         If I = 1 Then
            TempStr = FormatNumber(II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("CASH_FIRST_MONTH"), , , True)
            Total2(I + 2) = Total2(I + 2) + II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("CASH_FIRST_MONTH")
            SumLine1 = SumLine1 + II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("CASH_FIRST_MONTH")
         Else
            TempStr = FormatNumber(II.TOTAL_INCLUDE_PRICE, , , True)
            Total2(I + 2) = Total2(I + 2) + II.TOTAL_INCLUDE_PRICE
            SumLine1 = SumLine1 + II.TOTAL_INCLUDE_PRICE
         End If
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total2(15) = Total2(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินจ่ายค่าอาหาร
                        
      'เงินจ่ายค่ายา+วัคซีน
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = 3.2
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินจ่ายค่ายา+วัคซีน"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set II = GetImportItem(m_FeedTotalCash, Hc.TempKey & "-" & 2)
         If I = 1 Then
            TempStr = FormatNumber(II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("MEDICINE_FIRST_MONTH"), , , True)
            Total2(I + 2) = Total2(I + 2) + II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("MEDICINE_FIRST_MONTH")
            SumLine1 = SumLine1 + II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("MEDICINE_FIRST_MONTH")
         Else
            TempStr = FormatNumber(II.TOTAL_INCLUDE_PRICE, , , True)
            Total2(I + 2) = Total2(I + 2) + II.TOTAL_INCLUDE_PRICE
            SumLine1 = SumLine1 + II.TOTAL_INCLUDE_PRICE
         End If
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total2(15) = Total2(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินจ่ายค่ายา+วัคซีน
      
      'เงินจ่ายอื่นๆ
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = 3.3
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินจ่ายอื่นๆ"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set II = GetImportItem(m_FeedTotalCash, Hc.TempKey & "-" & 3)
         If I = 1 Then
            TempStr = FormatNumber(II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("OTHER_FIRST_MONTH"), , , True)
            Total2(I + 2) = Total2(I + 2) + II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("OTHER_FIRST_MONTH")
            SumLine1 = SumLine1 + II.TOTAL_INCLUDE_PRICE + m_Batch.GetFieldValue("OTHER_FIRST_MONTH")
         Else
            TempStr = FormatNumber(II.TOTAL_INCLUDE_PRICE, , , True)
            Total2(I + 2) = Total2(I + 2) + II.TOTAL_INCLUDE_PRICE
            SumLine1 = SumLine1 + II.TOTAL_INCLUDE_PRICE
         End If
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total2(15) = Total2(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินจ่ายอื่นๆ
      
      'เงินจ่ายค่าซื้อสุกร
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = 4
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินจ่ายค่าซื้อสุกร"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set II = GetImportItem(m_BuyTotalCash, Hc.TempKey)
         TempStr = FormatNumber(II.TOTAL_ACTUAL_PRICE, , , True)
         Total2(I + 2) = Total2(I + 2) + II.TOTAL_ACTUAL_PRICE
         SumLine1 = SumLine1 + II.TOTAL_ACTUAL_PRICE
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total2(15) = Total2(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินจ่ายค่าซื้อสุกร
                        
                        
      'เงินสดจ่ายอื่น ๆ
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = 5
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = "ค่าใช้จ่ายผลิต"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set Ri = GetRoItem(m_TotalExpenseCashs, Hc.TempKey)
         TempStr = FormatNumber(Ri.TOTAL_PRICE, , , True)
         Total2(I + 2) = Total2(I + 2) + Ri.TOTAL_PRICE
         SumLine1 = SumLine1 + Ri.TOTAL_PRICE
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total2(15) = Total2(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'ค่าใช้จ่ายผลิต
                        
      'คชจ ขาย บริหาร
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = "6"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = "หัก คชจ ขาย บริหาร"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          Set ParamItems = GetParamItem(m_MenagementExpenses, Hc.TempKey)
         TempStr = FormatNumber(ParamItems.GetFieldValue("EXP_AMOUNT"), , , True)
         Total2(I + 2) = Total2(I + 2) + ParamItems.GetFieldValue("EXP_AMOUNT")
         SumLine1 = SumLine1 + ParamItems.GetFieldValue("EXP_AMOUNT")
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Total2(15) = Total2(15) + SumLine1
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'คชจ ขาย บริหาร
      
      'รวมรายจ่าย
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = "รวม"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = ""
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
         TempStr = FormatNumber(Total2(I + 2), , , True)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      TempStr = FormatNumber(Total2(15), , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      'รวมรายจ่าย
            
      'เงินสดคงเหลือจากการดำเนินการ
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = ""
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินสดคงเหลือจากการดำเนินการ"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
         TempStr = FormatNumber(Total1(I + 2) - Total2(I + 2), , , True)
         SumLine1 = SumLine1 + (Total1(I + 2) - Total2(I + 2))
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      TempStr = FormatNumber(SumLine1, , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินสดคงเหลือจากการดำเนินการ
            
      'เงินสดยกมา
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = ""
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินสดยกมา"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
          Set Hc = m_Headers1.Fields(I + 2)
          If I = 1 Then
            CashBalance = m_Batch.GetFieldValue("BALANCE_CASH")
         Else
            CashBalance = Total1(I + 1) - Total2(I + 1)
          End If
         TempStr = FormatNumber(CashBalance, , , True)
         Total1(I + 2) = Total1(I + 2) + CashBalance
         
         CashBalance = CashBalance + Total1(I + 2) - Total2(I + 2)
         
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      TempStr = FormatNumber(m_Batch.GetFieldValue("BALANCE_CASH"), , , True)
      Total1(15) = Total1(15) + m_Batch.GetFieldValue("BALANCE_CASH")
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินสดยกมา
      
      'เงินสดสิ้นงวด
      SumLine1 = 0
      Call m_Details1.ClearField
       Set Hc = m_Headers1.Fields(1)
      TempStr = ""
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

       Set Hc = m_Headers1.Fields(2)
      TempStr = "เงินสดสิ้นงวด"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      For I = 1 To 12
         Set Hc = m_Headers1.Fields(I + 2)
         TempStr = FormatNumber(Total1(I + 2) - Total2(I + 2), , , True)
         SumLine1 = SumLine1 + (Total1(I + 2) - Total2(I + 2))
         
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      TempStr = FormatNumber(Total1(15) - Total2(15), , , True)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
            
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      'เงินสดสิ้นงวด
      
      Set Di1 = Nothing
      Set BodyCf = Nothing
    Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateSubHeader(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   TempStr = Txt1
   TempWidth = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

   TempStr = Txt2
   TempWidth = 0
   For j = 3 To m_Headers1.Fields.Count
      TempWidth = TempWidth + m_Headers1.Fields(j).Width
   Next j
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "^")
   Call Details.AddField(BodyCf)

   Set BodyCf = Nothing
End Sub

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
      ElseIf j = m_Headers1.Fields.Count Then
         TempStr = FormatNumber(MyDiffEx(Tot(5), Tot(4)), , , True)
      Else
         TempStr = FormatNumber(Tot(j), , , True)
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
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(Now), 8))
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

