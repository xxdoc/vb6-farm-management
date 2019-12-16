VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR002"
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
Private m_YearSeqs As Collection

Public m_PaidAmounts As Collection
Public m_DnAmounts As Collection
Public m_CnAmounts As Collection

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
   Set m_YearSeqs = New Collection
   
   Set m_DnAmounts = New Collection
   Set m_CnAmounts = New Collection
   Set m_PaidAmounts = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set m_DnAmounts = Nothing
   Set m_CnAmounts = Nothing
   Set m_PaidAmounts = Nothing
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
   
   Set Cf = New CReportField

   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(9, "^", "", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(6, "^", "ขายสุกร", ">")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(4, "^", "ขายวัตถุดิบ", ">")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(4, "^", "ขายอื่น ๆ", ">")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "รับชำระ", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เพิ่ม/ลดหนี้", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "อ้างถึง", "^")
   Call m_Headers0.AddField(Cf)
   '===
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ขายเชื่อ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายสด", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "ขายเชื่อ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายสด", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ขายเชื่อ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายสด", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รับชำระ", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "เพิ่ม/ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "อ้างถึง", "^")
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
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
      
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, mcolParam("FROM_DATE"), mcolParam("TO_PAY_DATE"))
   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, mcolParam("FROM_DATE"), mcolParam("TO_PAY_DATE"), 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, mcolParam("FROM_DATE"), mcolParam("TO_PAY_DATE"), 4, 2)

   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A") & " ถึง วันที่ชำระเงิน " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PAY_DATE")), "N/A")
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

Private Function FilterDocumentType(Bd As CBillingDoc) As Boolean
   If Bd.DOCUMENT_TYPE = 1 Then
      FilterDocumentType = True
   ElseIf Bd.DOCUMENT_TYPE = 2 Then
      FilterDocumentType = True
   Else
      FilterDocumentType = False
   End If
End Function

Private Sub GenerateBalanceItem(Txt1 As String, SellItems As Collection, PaidItems As Collection, RevenueItems As Collection, DebitCreditItems As Collection, Total() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Di As CDoItem
Dim Di2_1 As CDoItem
Dim Di2_2 As CDoItem
Dim Ri As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem
Dim KeyStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = ""
'         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = ""
'         TempStr = Txt3
      ElseIf (j = 5) Then
         KeyStr = "1-1-0" 'ขายเชื่อหมู
         Set Di = GetDoItem(SellItems, KeyStr)
         TempStr = FormatNumber(Di.TOTAL_PRICE, "-")
         Total(5) = Total(5) + Di.TOTAL_PRICE
      ElseIf (j = 6) Then
         KeyStr = "2-1-1" 'ขายสดหมู
         Set Di = GetDoItem(SellItems, KeyStr)
         TempStr = FormatNumber(Di.TOTAL_PRICE, "-")
         Total(6) = Total(6) + Di.TOTAL_PRICE
      ElseIf (j = 8) Then 'ขายเชื่อ วัตถุดิบ
         KeyStr = "1-2-0"
         Set Di = GetDoItem(SellItems, KeyStr)
         TempStr = FormatNumber(Di.TOTAL_PRICE, "-")
         Total(8) = Total(8) + Di.TOTAL_PRICE
      ElseIf (j = 9) Then
         KeyStr = "2-2-1" 'ขายสด วัตถุดิบ
         Set Di = GetDoItem(SellItems, KeyStr)
         TempStr = FormatNumber(Di.TOTAL_PRICE, "-")
         Total(9) = Total(9) + Di.TOTAL_PRICE
      
      ElseIf (j = 10) Then
         KeyStr = "1-1-0" 'ขายเชื่อหมู
         Set Di2_1 = GetDoItem(RevenueItems, KeyStr)
         KeyStr = "1-2-0" 'ขายเชื่อวัตถุดิบ
         Set Di2_2 = GetDoItem(RevenueItems, KeyStr)
         TempStr = FormatNumber(Di2_1.TOTAL_PRICE + Di2_2.TOTAL_PRICE, "-")
         Total(10) = Total(10) + Di2_1.TOTAL_PRICE + Di2_2.TOTAL_PRICE
         
      ElseIf (j = 11) Then
         KeyStr = "2-1-1" 'ขายสดหมู
         Set Di2_1 = GetDoItem(RevenueItems, KeyStr)
         KeyStr = "2-2-1" 'ขายสดวัตถุดิบ
         Set Di2_2 = GetDoItem(RevenueItems, KeyStr)
         TempStr = FormatNumber(Di2_1.TOTAL_PRICE + Di2_2.TOTAL_PRICE, "-")
         Total(11) = Total(11) + Di2_1.TOTAL_PRICE + Di2_2.TOTAL_PRICE
      ElseIf (j = 12) Then
         KeyStr = "2-0-3"
         Set Ri = GetReceiptItem(PaidItems, KeyStr)
         TempStr = FormatNumber(Ri.PAID_AMOUNT, "-")
         Total(12) = Total(12) + Ri.PAID_AMOUNT
      ElseIf (j = 13) Then
         KeyStr = "3" 'Debit note
         Set Ri1_1 = GetReceiptItem(DebitCreditItems, KeyStr)
         KeyStr = "4" 'Credit note
         Set Ri1_2 = GetReceiptItem(DebitCreditItems, KeyStr)
         TempStr = FormatNumber(Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT, "-")
         Total(13) = Total(13) + (Ri1_1.DEBIT_CREDIT_AMOUNT - Ri1_2.DEBIT_CREDIT_AMOUNT)
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Bd As CBillingDoc
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
Dim PaidItems As Collection
Dim SellItems As Collection
Dim NewDate As Date
Dim NewFromDate As Date
Dim NewToDate As Date
Dim RevenueItems As Collection
Dim DebitCreditItems As Collection
Dim Ri1_0 As CReceiptItem
Dim Ri1_1 As CReceiptItem
Dim Ri1_2 As CReceiptItem

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set PaidItems = New Collection
   Set SellItems = New Collection
   Set RevenueItems = New Collection
   Set DebitCreditItems = New Collection
   
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   
   If NewDate > 0 Then
      'ใช้ยอดยกมาตั้งแต่ต้นเดือน
      Call GetFirstLastDate(mcolParam("FROM_DATE"), NewFromDate, NewToDate)

      Call LoadDnCnAmountByDocType(Nothing, DebitCreditItems, NewFromDate, NewDate) 'ใบเพิ่มหนี้
      Call LoadSellRevenuePriceByDocTypeSubTypeReceiptType(Nothing, RevenueItems, NewFromDate, NewDate)
      Call LoadSellPriceByDocTypeSubTypeReceiptType(Nothing, SellItems, NewFromDate, NewDate)
      Call LoadReceiptByDocTypeSubTypeReceiptType(Nothing, PaidItems, NewFromDate, NewDate)
   End If
   
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
      I = 0
      Set BodyCf = New CReportField

      Set Bd = New CBillingDoc
      Bd.BILLING_DOC_ID = -1
      Bd.FROM_DATE = mcolParam("FROM_DATE")
      Bd.TO_DATE = mcolParam("TO_DATE")
      Bd.OrderBy = mcolParam("ORDER_BY")
      Bd.OrderType = mcolParam("ORDER_TYPE")
      Bd.ItemSumFlag = True
      Call glbDaily.QueryBillingDoc(Bd, Rs, iCount, IsOK, glbErrorLog)
      I = 0

      While Not Rs.EOF
         Call m_Details1.ClearField
         Call Bd.PopulateFromRS(1, Rs)

         If FilterDocumentType(Bd) Then
            Set Ri1_1 = GetReceiptItem(m_DnAmounts, Bd.BILLING_DOC_ID) 'เพิ่มหนี้
            Set Ri1_2 = GetReceiptItem(m_CnAmounts, Bd.BILLING_DOC_ID) 'ลดหนี้
            
            Bd.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
            Bd.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
         
            I = I + 1
         
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = I
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = Bd.DOCUMENT_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = DateToStringExtEx2(Bd.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(4)
            TempStr = Bd.CUSTOMER_NAME & " (" & Bd.CUSTOMER_CODE & ")"
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            '===
            Set HeadCf = m_Headers1.Fields(5)
            If (Bd.DOCUMENT_TYPE = 1) And (Bd.DOCUMENT_SUBTYPE = 1) Then
               TempStr = FormatNumber(Bd.DO_TOTAL_PRICE, "-")
               Total1(5) = Total1(5) + Bd.DO_TOTAL_PRICE
            Else
               TempStr = FormatNumber(0, "-")
               Total1(5) = Total1(5) + 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(6)
            If (Bd.DOCUMENT_TYPE = 2) And (Bd.DOCUMENT_SUBTYPE = 1) And (Bd.RECEIPT_TYPE = 1) Then  'สด
               TempStr = FormatNumber(Bd.DO_TOTAL_PRICE, "-")
               Total1(6) = Total1(6) + Bd.DO_TOTAL_PRICE
            Else
               TempStr = FormatNumber(0, "-")
               Total1(6) = Total1(6) + 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            '===
            
            '===
            Set HeadCf = m_Headers1.Fields(8)
            If (Bd.DOCUMENT_TYPE = 1) And (Bd.DOCUMENT_SUBTYPE = 2) Then
               TempStr = FormatNumber(Bd.DO_TOTAL_PRICE, "-")
               Total1(8) = Total1(8) + Bd.DO_TOTAL_PRICE
            Else
               TempStr = FormatNumber(0, "-")
               Total1(8) = Total1(8) + 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(9)
            If (Bd.DOCUMENT_TYPE = 2) And (Bd.DOCUMENT_SUBTYPE = 2) And (Bd.RECEIPT_TYPE = 1) Then  'สด
               TempStr = FormatNumber(Bd.DO_TOTAL_PRICE, "-")
               Total1(9) = Total1(9) + Bd.DO_TOTAL_PRICE
            Else
               TempStr = FormatNumber(0, "-")
               Total1(9) = Total1(9) + 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(10)
            If ((Bd.DOCUMENT_TYPE = 1) And (Bd.DOCUMENT_SUBTYPE = 1)) Or _
                  ((Bd.DOCUMENT_TYPE = 1) And (Bd.DOCUMENT_SUBTYPE = 2)) Then
               TempStr = FormatNumber(Bd.REVENUE_TOTAL_PRICE, "-")
               Total1(10) = Total1(10) + Bd.REVENUE_TOTAL_PRICE
            Else
               TempStr = FormatNumber(0, "-")
               Total1(10) = Total1(10) + 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(11)
            If ((Bd.DOCUMENT_TYPE = 2) And (Bd.DOCUMENT_SUBTYPE = 1)) Or _
                  ((Bd.DOCUMENT_TYPE = 2) And (Bd.DOCUMENT_SUBTYPE = 2) And (Bd.RECEIPT_TYPE = 1)) Then
               TempStr = FormatNumber(Bd.REVENUE_TOTAL_PRICE, "-")
               Total1(11) = Total1(11) + Bd.REVENUE_TOTAL_PRICE
            Else
               TempStr = FormatNumber(0, "-")
               Total1(11) = Total1(11) + 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            '==
            Set HeadCf = m_Headers1.Fields(12)
            If (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then 'รับชำระหนี้
               TempStr = FormatNumber(0, "-") 'Bd.RECEIPT_PAID_AMOUNT
               Total1(12) = Total1(12) + 0 'Bd.RECEIPT_PAID_AMOUNT
            Else
               TempStr = FormatNumber(0, "-")
               Total1(12) = Total1(12) + 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            '===
            
            '===
            Set HeadCf = m_Headers1.Fields(13)
            TempStr = FormatNumber((Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT), "-")
            Total1(13) = Total1(13) + (Bd.DEBIT_AMOUNT - Bd.CREDIT_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(14)
            TempStr = FormatNumber(0, "-")
            Total1(14) = Total1(14) + 0
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            '===
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            
            If (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then
               Call GenerateReceiptItem(Bd, m_Details1, Total1)
            End If
         End If
         Rs.MoveNext
      Wend

      Call GenerateFooter("สุทธิ", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateBalanceItem("ยกมา", SellItems, PaidItems, RevenueItems, DebitCreditItems, Total1, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("ยกไป", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Set DebitCreditItems = Nothing
   Set RevenueItems = Nothing
   Set PaidItems = Nothing
   Set SellItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateReceiptItem(TempBd As CBillingDoc, Details As CFieldList, Tot() As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Bd As CBillingDoc
Dim alngX() As Long
Dim j As Long
Dim Ri As CReceiptItem
Dim TempStr1 As String
Dim TempStr2 As String
Dim IsOK As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim strFormat As String
Dim Sum As Double
Dim Ri1_0 As CReceiptItem

   Set TempRs = New ADODB.Recordset
   
   Set Bd = New CBillingDoc
   Bd.QueryFlag = 1
   Bd.BILLING_DOC_ID = TempBd.BILLING_DOC_ID
   Call glbDaily.QueryBillingDoc(Bd, TempRs, iCount, IsOK, glbErrorLog)
      
   Set BodyCf = New CReportField
   
   Sum = 0
   For Each Ri In Bd.ReceiptItems
      Set Ri1_0 = GetReceiptItem(m_PaidAmounts, Ri.DO_ID)  'รับชำระ
      Ri.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
      
      j = 0
      Call Details.ClearField
      For Each HeadCf In m_Headers1.Fields
         j = j + 1
         If j = 12 Then
            TempStr1 = FormatNumber(Ri.PAID_AMOUNT, "-")
            Tot(j) = Tot(j) + Ri.PAID_AMOUNT
            Sum = Sum + Ri.PAID_AMOUNT
         ElseIf j = 14 Then
            TempStr1 = Ri.DOCUMENT_NO
            Tot(j) = Tot(j) + 0
         Else
            TempStr1 = ""
         End If
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr1, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      Next HeadCf
      
      Call Details.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
   Next Ri
   
   '====
   j = 0
   Call Details.ClearField
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr1 = "รวม"
         Tot(j) = Tot(j) + 0
      ElseIf j = 12 Then
         TempStr1 = FormatNumber(Sum, "-")
      Else
         TempStr1 = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr1, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
   
   Call Details.GetString(2, TempStr1, TempStr2)
   Vsp.FontBold = True
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   '====
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set BodyCf = Nothing
   Set Bd = Nothing
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
         TempStr = Txt3
      ElseIf (j = 4) Then
      Else
         TempStr = FormatNumber(Tot(j), "-")
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
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
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






