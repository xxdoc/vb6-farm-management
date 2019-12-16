VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSell010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell010"
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

Private m_PaidAmounts As Collection
Private m_BillTotalPrices As Collection
Private m_DnItemsByBill As Collection
Private m_CnItemsByBill As Collection

Private m_DistinceSellPartGroup As Collection

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
   
   Set m_PaidAmounts = New Collection
   Set m_BillTotalPrices = New Collection
   Set m_DnItemsByBill = New Collection
   Set m_CnItemsByBill = New Collection
   
   Set m_DistinceSellPartGroup = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set m_PaidAmounts = Nothing
   Set m_BillTotalPrices = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   
   Set m_DistinceSellPartGroup = Nothing
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
   Dim i As Long
   
   For i = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next i

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
'   Call Cf.SetFieldValue(1, "^", "NO.", "^")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "", "^")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "", "^")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(3, "^", "", "^")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(6, "^", "จะครบกำหนด", ">")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(10, "^", "เกินกำหนด", ">")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "", "^")
'   Call m_Headers0.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "เกินกำหนด", "^")
'   Call m_Headers0.AddField(Cf)
   '===
   Call Cf.SetFieldValue(2, "^", "วันที่ส่งของ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เลขที่ใบส่งของ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "ประเภท", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "วันที่ชำระ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เลขที่ชำระ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนวัน", "^")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing

'  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
            
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

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
    
   Call LoadTotalPriceByBill(Nothing, m_BillTotalPrices, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , , mcolParam("CUSTOMER_CODE"))
   
   Call LoadDistinctSellGroupType(Nothing, m_DistinceSellPartGroup, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , , mcolParam("CUSTOMER_CODE"))
   
   Call LoadAllPaidAmountByBill(Nothing, m_PaidAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , , , , mcolParam("CUSTOMER_CODE"))
      
'   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), , , , , mcolParam("CUSTOMER_CODE"))
'   Call LoadDnCnAmountByBill(Nothing, m_DnItemsByBill, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), 3, 2, , mcolParam("CUSTOMER_CODE"))
   'Call LoadDnCnAmountByBill(Nothing, m_CnItemsByBill, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), 4, 2, , mcolParam("CUSTOMER_CODE"))
    
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
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

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

'Private Sub GenerateDetail(alngX() As Long, I As Long, Apt As CBillingDoc, Tot1() As Double, Tot2() As Double)
'Dim HeadCf As CReportField
'Dim TempStr As String
'Dim TempStr1 As String
'Dim TempStr2 As String
'Dim BodyCf As CReportField
'Dim strFormat As String
'Dim Ari As CReceiptItem
'Dim DueCount As Long
'Dim MM As CMaxMin
'Dim j As Long
'Dim Sum As Double
'Dim PrintDate As Date
'Dim Di As CDoItem
'Dim Ri1_1 As CReceiptItem
'Dim Ri1_2 As CReceiptItem
'
'   Set BodyCf = New CReportField
'
'   If mcolParam("PRINT_DATE") <= 0 Then
'      PrintDate = Now
'   Else
'      PrintDate = mcolParam("PRINT_DATE")
'   End If
'
'   Set Di = GetDoItem(m_BillTotalPrices, Apt.BILLING_DOC_ID)
'
'   If Val(mcolParam("CREDIT_AMOUNT")) > 0 Then
'      DueCount = DateDiff("D", DateAdd("D", Val(mcolParam("CREDIT_AMOUNT")), Apt.DOCUMENT_DATE), PrintDate)
'   Else
'      DueCount = DateDiff("D", Apt.DUE_DATE, PrintDate)
'   End If
'
'
'   Set Ari = GetReceiptItem(m_PaidAmounts, Apt.BILLING_DOC_ID)
'   Set Ri1_1 = GetReceiptItem(m_DnItemsByBill, Apt.BILLING_DOC_ID)
'   Set Ri1_2 = GetReceiptItem(m_CnItemsByBill, Apt.BILLING_DOC_ID)
'
'   If (Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) > Ari.PAID_AMOUNT Then   'จ่ายยังไม่ครบ
'      Set HeadCf = m_Headers1.Fields(1)
'      TempStr = I
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
'
'      Set HeadCf = m_Headers1.Fields(2)
'      If Val(mcolParam("CREDIT_AMOUNT")) > 0 Then
'         TempStr = DateToStringExtEx2(DateAdd("D", Val(mcolParam("CREDIT_AMOUNT")), Apt.DOCUMENT_DATE))
'      Else
'         TempStr = DateToStringExtEx2(Apt.DUE_DATE)
'      End If
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
'
'      Set HeadCf = m_Headers1.Fields(3)
'      TempStr = Apt.DOCUMENT_NO
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
'
'      Set HeadCf = m_Headers1.Fields(4)
'      TempStr = DateToStringExtEx2(Apt.DOCUMENT_DATE)
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
'
'       j = 5
'      Sum = 0
'      For Each MM In m_MaxMins
'         Set HeadCf = m_Headers1.Fields(j)
'         If (MM.MIN < DueCount) And (DueCount <= MM.MAX) Then
'            TempStr = FormatNumber((Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - Ari.PAID_AMOUNT, "-")
'
'            Tot1(j) = Tot1(j) + ((Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - Ari.PAID_AMOUNT)
'            Tot2(j) = Tot2(j) + ((Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - Ari.PAID_AMOUNT)
'            Sum = Sum + ((Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - Ari.PAID_AMOUNT)
'         Else
'            TempStr = FormatNumber(0, "-")
'         End If
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         j = j + 1
'      Next MM
'
'      Set HeadCf = m_Headers1.Fields(j)
'      TempStr = FormatNumber(Sum, "-")
'      Tot1(j) = Tot1(j) + ((Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - Ari.PAID_AMOUNT)
'      Tot2(j) = Tot2(j) + ((Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT) - Ari.PAID_AMOUNT)
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
'
'      j = j + 1
'      Set HeadCf = m_Headers1.Fields(j)
'      TempStr = FormatNumberInt(DueCount)
'      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)
'
'      Call m_Details1.GetString(1, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'   Else
'      I = I - 1
'   End If
'   Set BodyCf = Nothing
'End Sub
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
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
Dim iCountBill As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim Hc As CReportField

Dim Bd As CBillingDoc
Dim TempDi As CDoItem
Dim TempDiPartGroup As CDoItem
Dim TempBD As CBillingDoc
Dim TempRi As CReceiptItem
Dim CountCollumn As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
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
      
   Set BodyCf = New CReportField
   Set Bd = New CBillingDoc
   Bd.BILLING_DOC_ID = -1
   Bd.FROM_DATE = mcolParam("FROM_DATE")
   Bd.TO_DATE = mcolParam("TO_DATE")
   Bd.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
   'Bd.DOCUMENT_TYPE = 1
   Bd.DocSubTypeSet = "(1, 2) " 'หมู, วัตถุดิบ
   Bd.OrderBy = 1
   Bd.OrderType = 1
   Call Bd.QueryData(3, Rs, iCount)
      
   i = 0
      
   While Not Rs.EOF
      i = i + 1
      
      Call Bd.PopulateFromRS(3, Rs)
      
      If PrevKey1 <> Trim(Bd.CUSTOMER_CODE) And i <> 1 Then
         Call GenerateFooter("รวม", PrevKey1, Total1, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
            
         Vsp = "     " & Bd.CUSTOMER_NAME & " ( " & Bd.CUSTOMER_CODE & " )"
      ElseIf i = 1 Then
         Vsp = "     " & Bd.CUSTOMER_NAME & " ( " & Bd.CUSTOMER_CODE & " )"
      End If
      
      PrevKey1 = Trim(Bd.CUSTOMER_CODE)
      
      Set TempDi = GetDoItem(m_BillTotalPrices, Trim(Str(Bd.BILLING_DOC_ID)))
      
      Call m_Details1.ClearField
      
      Set Hc = m_Headers1.Fields(1)
      TempStr = DateToStringExtEx2(Bd.DOCUMENT_DATE)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
         
      Set Hc = m_Headers1.Fields(2)
      TempStr = Bd.DOCUMENT_NO
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set Hc = m_Headers1.Fields(3)
      TempStr = FormatNumber(TempDi.TOTAL_PRICE - TempDi.DISCOUNT_AMOUNT)
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(3) = Total1(3) + (TempDi.TOTAL_PRICE - TempDi.DISCOUNT_AMOUNT)
      Total2(3) = Total2(3) + (TempDi.TOTAL_PRICE - TempDi.DISCOUNT_AMOUNT)
            
      Set TempDiPartGroup = GetDoItem(m_DistinceSellPartGroup, Trim(Str(Bd.BILLING_DOC_ID)))
      
      Set Hc = m_Headers1.Fields(4)
      If TempDiPartGroup.PART_GROUP_ID > 0 Then
         TempStr = TempDiPartGroup.PART_GROUP_NAME
      Else
         TempStr = "หมู"
      End If
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      CountCollumn = 1
      Set TempBD = GetBillingDoc(m_PaidAmounts, Trim(Str(Bd.BILLING_DOC_ID)))
      If TempBD.ReceiptItems.Count = 0 Then
         Set Hc = m_Headers1.Fields(5)
         TempStr = ""
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
            
         Set Hc = m_Headers1.Fields(6)
         TempStr = ""
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
   
         Set Hc = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set Hc = m_Headers1.Fields(8)
         TempStr = ""
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
               
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Else
         For Each TempRi In TempBD.ReceiptItems
            If CountCollumn > 1 Then
               Call m_Details1.ClearField
               TempStr = ""
               Set Hc = m_Headers1.Fields(1)
               TempStr = ""
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                  
               Set Hc = m_Headers1.Fields(2)
               TempStr = ""
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
         
               Set Hc = m_Headers1.Fields(3)
               TempStr = ""
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set Hc = m_Headers1.Fields(4)
               TempStr = ""
               Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            End If
            CountCollumn = CountCollumn + 1
            Set Hc = m_Headers1.Fields(5)
            TempStr = DateToStringExtEx2(TempRi.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set Hc = m_Headers1.Fields(6)
            TempStr = TempRi.DOCUMENT_NO
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set Hc = m_Headers1.Fields(7)
            TempStr = FormatNumber(TempRi.PAID_AMOUNT)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(7) = Total1(7) + TempRi.PAID_AMOUNT
            Total2(7) = Total2(7) + TempRi.PAID_AMOUNT
         
            Set Hc = m_Headers1.Fields(8)
            TempStr = DateDiff("D", Bd.DOCUMENT_DATE, TempRi.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         Next TempRi
      End If
      
      Rs.MoveNext
   Wend

   Call GenerateFooter("รวม", PrevKey1, Total1, m_Details1, 0)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
         
   Call GenerateFooter("สุทธิ", "", Total2, m_Details1, 0)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

      Vsp = ""

      
      Set BodyCf = Nothing
    Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
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

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = TXT1
      ElseIf j = 2 Then
         TempStr = TXT2
      ElseIf j = 3 Or j = 7 Then
         TempStr = FormatNumber(Tot(j))
      Else
         TempStr = ""
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








