VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR004_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR004"
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

   Call Cf.SetFieldValue(0, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "หนี้เพิ่ม", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "หนี้ลด", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(0, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายเชื่อ (สุกร)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขายเชื่อ (อื่น ๆ)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เพิ่มหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับชำระ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
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
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   FromDate = "จาก วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
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

Private Sub GenerateBalanceItem(Txt1 As String, Total() As Double, Details As CFieldList, Bd As CBillingDoc, SellItems As Collection, PaidItems As Collection, DnBalances As Collection, CnBalances As Collection, BalAmount As Double)
Dim Ri As CReceiptItem
Dim Di1 As CDoItem
Dim Di2 As CDoItem
Dim j As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim LeftAmount As Double
Dim TempRi1 As CReceiptItem
Dim TempRi2 As CReceiptItem

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   LeftAmount = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = ""
      ElseIf j = 2 Then
         TempStr = Txt1
      ElseIf j = 3 Then
         TempStr = ""
         KeyStr = Bd.ACCOUNT_ID & "-1-1-0"
         Set Di1 = GetDoItem(SellItems, KeyStr)
      
         KeyStr = Bd.ACCOUNT_ID & "-1-2-0"
         Set Di2 = GetDoItem(SellItems, KeyStr)
      
         Set TempRi1 = GetReceiptItem(DnBalances, Trim(Str(Bd.ACCOUNT_ID)))
         
         KeyStr = Bd.ACCOUNT_ID & "-2-0-3"
         Set Ri = GetReceiptItem(PaidItems, KeyStr)
      
         Set TempRi2 = GetReceiptItem(CnBalances, Trim(Str(Bd.ACCOUNT_ID)))
         LeftAmount = Di1.TOTAL_PRICE + Di2.TOTAL_PRICE + TempRi1.DEBIT_CREDIT_AMOUNT - Ri.PAID_AMOUNT - TempRi2.DEBIT_CREDIT_AMOUNT
      
         TempStr = FormatNumber(LeftAmount, "-")
         Total(3) = Total(3) + LeftAmount
      ElseIf (j = 4) Then 'ซื้อเชื่อหมู
         TempStr = FormatNumber(0, "-")
         Total(4) = Total(4) + 0
      ElseIf (j = 5) Then 'ซื้อเชื่ออื่น ๆ
         TempStr = FormatNumber(0, "-")
         Total(5) = Total(5) + 0
      ElseIf (j = 6) Then 'เพิ่มหนี้
         TempStr = FormatNumber(0, "-")
         Total(6) = Total(6) + 0
      ElseIf (j = 7) Then 'รับชำระหนี้
         TempStr = FormatNumber(0, "-")
         Total(7) = Total(7) + 0
      ElseIf (j = 8) Then 'ลดหนี้
         TempStr = FormatNumber(0, "-")
         Total(8) = Total(8) + 0
      ElseIf (j = 9) Then 'คงเหลือ
         TempStr = FormatNumber(LeftAmount, "-")
         Total(9) = Total(9) + LeftAmount
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
   BalAmount = LeftAmount
End Sub

Private Sub GenerateDetail(Txt1 As String, Total() As Double, Details As CFieldList, Bd As CBillingDoc, iCount As Long, BalAmount As Double)
Dim j As Long
Dim KeyStr As String
Dim BodyCf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim LeftAmount As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   LeftAmount = BalAmount
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = iCount
      ElseIf j = 2 Then
      ElseIf (j = 3) Then 'ยกมา
         TempStr = FormatNumber(0, "-")
         Total(3) = Total(3) + 0
         LeftAmount = LeftAmount + 0
      ElseIf (j = 4) Then 'ซื้อเชื่อหมู
         If (Bd.DOCUMENT_TYPE = 1) And (Bd.DOCUMENT_SUBTYPE = 1) Then
            TempStr = FormatNumber(Bd.DO_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + Bd.REVENUE_TOTAL_PRICE, "-")
            Total(4) = Total(4) + Bd.DO_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + Bd.REVENUE_TOTAL_PRICE
            LeftAmount = LeftAmount + Bd.DO_TOTAL_PRICE - Bd.DISCOUNT_AMOUNT + Bd.REVENUE_TOTAL_PRICE
         Else
            TempStr = FormatNumber(0, "-")
            Total(4) = Total(4) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (j = 5) Then 'ซื้อเชื่ออื่น ๆ
         If (Bd.DOCUMENT_TYPE = 1) And (Bd.DOCUMENT_SUBTYPE = 2) Then
            TempStr = FormatNumber(Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE, "-")
            Total(5) = Total(5) + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE
            LeftAmount = LeftAmount + Bd.DO_TOTAL_PRICE + Bd.REVENUE_TOTAL_PRICE
         Else
            TempStr = FormatNumber(0, "-")
            Total(5) = Total(5) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (j = 6) Then 'เพิ่มหนี้
         If Bd.DOCUMENT_TYPE = 3 Then
            TempStr = FormatNumber(Bd.CNDN_TOTAL_PRICE, "-")  'FormatNumber(Bd.DEBIT_AMOUNT, "-")
            Total(6) = Total(6) + Bd.CNDN_TOTAL_PRICE 'Bd.DEBIT_AMOUNT
            LeftAmount = LeftAmount + Bd.CNDN_TOTAL_PRICE 'Bd.DEBIT_AMOUNT
         Else
            TempStr = FormatNumber(0, "-")
            Total(6) = Total(6) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (j = 7) Then 'รับชำระหนี้
         If (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then
            TempStr = FormatNumber(Bd.RECEIPT_PAID_AMOUNT, "-")
            Total(7) = Total(7) + Bd.RECEIPT_PAID_AMOUNT
            LeftAmount = LeftAmount - Bd.RECEIPT_PAID_AMOUNT
         Else
            TempStr = FormatNumber(0, "-")
            Total(7) = Total(7) + 0
            LeftAmount = LeftAmount + 0
         End If
      ElseIf (j = 8) Then 'ลดหนี้
         If Bd.DOCUMENT_TYPE = 4 Then
            TempStr = FormatNumber(Bd.CNDN_TOTAL_PRICE, "-") 'Bd.CREDIT_AMOUNT
            Total(8) = Total(8) + Bd.CNDN_TOTAL_PRICE 'Bd.CREDIT_AMOUNT
            LeftAmount = LeftAmount - Bd.CNDN_TOTAL_PRICE 'Bd.CREDIT_AMOUNT
         Else
            TempStr = FormatNumber(0, "-")
            Total(8) = Total(8) + 0
            LeftAmount = LeftAmount - 0
         End If
      ElseIf (j = 9) Then 'คงเหลือ
         TempStr = FormatNumber(LeftAmount, "-")
         Total(9) = Total(9) + LeftAmount
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
   BalAmount = LeftAmount
End Sub

Private Function FilterDocumentType(Bd As CBillingDoc) As Boolean
   If Bd.DOCUMENT_TYPE = 1 Then
      FilterDocumentType = True
   ElseIf (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then
      FilterDocumentType = True
   ElseIf (Bd.DOCUMENT_TYPE = 3) Or (Bd.DOCUMENT_TYPE = 4) Then
      FilterDocumentType = True
   Else
      FilterDocumentType = False
   End If
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
Dim TempID As Long
Dim NewDate As Date
Dim Accounts As Collection
Dim DistinctAccounts As Collection
Dim Bd As CBillingDoc
Dim Ac As CAccount
Dim SumWidth As Double
Dim SellItems As Collection
Dim PaidItems As Collection
Dim TempBD As CBillingDoc
Dim BalanceAmount As Double
Dim DnBalances As Collection
Dim CnBalances As Collection
Dim HaveValue As Boolean

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Accounts = New Collection
   Call LoadAccount(Nothing, Accounts)
   
   Set DistinctAccounts = New Collection
   Call LoadBillingDocDistinctAccount(Nothing, DistinctAccounts, mcolParam("CUSTOMER_CODE"))
   
   Set SellItems = New Collection
   Set PaidItems = New Collection
   Set CnBalances = New Collection
   Set DnBalances = New Collection
   
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   If NewDate > 0 Then
      Call LoadDnCnAmountByAccount(Nothing, DnBalances, -1, NewDate, 3, 2, NewDate)
      Call LoadDnCnAmountByAccount(Nothing, CnBalances, -1, NewDate, 4, 2, NewDate)
      Call LoadSellPriceByDocTypeSubTypeReceiptTypeAcc(Nothing, SellItems, -1, NewDate, NewDate)
      Call LoadReceiptByDocTypeSubTypeReceiptTypeAcc(Nothing, PaidItems, -1, NewDate, NewDate)
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
   
   Set TempBD = New CBillingDoc
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField
      For Each Bd In DistinctAccounts
         'Show customer
         Set Ac = Accounts(Trim(Str(Bd.ACCOUNT_ID)))
         
         Call m_Details1.ClearField

'         Set HeadCf = m_Headers1.Fields(1)
'         SumWidth = HeadCf.Width
'         Set HeadCf = m_Headers1.Fields(2)
'         SumWidth = SumWidth + HeadCf.Width
'         TempStr = Ac.CUSTOMER_CODE
'         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
'         Call m_Details1.AddField(BodyCf)
'
'         TempStr = ""
'         Set HeadCf = m_Headers1.Fields(3)
'         SumWidth = HeadCf.Width
'         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(4)
'         SumWidth = HeadCf.Width
'         Set HeadCf = m_Headers1.Fields(5)
'         SumWidth = SumWidth + HeadCf.Width
'         Set HeadCf = m_Headers1.Fields(6)
'         SumWidth = SumWidth + HeadCf.Width
'         Set HeadCf = m_Headers1.Fields(7)
'         SumWidth = SumWidth + HeadCf.Width
'         Set HeadCf = m_Headers1.Fields(8)
'         SumWidth = SumWidth + HeadCf.Width
'         TempStr = Ac.CUSTOMER_NAME
'         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(9)
'         SumWidth = HeadCf.Width
'         Set HeadCf = m_Headers1.Fields(10)
'         SumWidth = SumWidth + HeadCf.Width
'         TempStr = Ac.ACCOUNT_NO
'         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, "<")
'         Call m_Details1.AddField(BodyCf)
'
'         Call m_Details1.GetString(2, TempStr1, TempStr2)
'         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'         Call Vsp.AddTable(strFormat, "", TempStr2)
         'Show customer
         
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
         
         Call GenerateBalanceItem("***ยกมา***", Total1, m_Details1, Bd, SellItems, PaidItems, DnBalances, CnBalances, BalanceAmount)
'         Call m_Details1.GetString(2, TempStr1, TempStr2)
'         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'         Call Vsp.AddTable(strFormat, "", TempStr2)
            
         Set TempBD = Nothing
         Set TempBD = New CBillingDoc
         TempBD.BILLING_DOC_ID = -1
         TempBD.ACCOUNT_ID = Bd.ACCOUNT_ID
         TempBD.FROM_DATE = mcolParam("FROM_DATE")
         TempBD.TO_DATE = mcolParam("TO_DATE")
         TempBD.OrderBy = 6
         TempBD.OrderType = mcolParam("ORDER_TYPE")
         TempBD.ItemSumFlag = True
         Call TempBD.SetFlag(False, True, False, False, True, True)
         Call glbDaily.QueryBillingDoc(TempBD, Rs, iCount, IsOK, glbErrorLog)
         While Not Rs.EOF
            Call TempBD.PopulateFromRS(1, Rs)
'If TempBd.DOCUMENT_TYPE = 4 Then
'''debug.print
'End If
            If FilterDocumentType(TempBD) Then
               I = I + 1
               
               Call GenerateDetail("", Total1, m_Details1, TempBD, I, BalanceAmount)
'               Call m_Details1.GetString(2, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            Rs.MoveNext
         Wend
         
         HaveValue = False
         
         For j = 1 To UBound(Total1)
            If j <> 9 Then
               Total2(j) = Total2(j) + Total1(j)
            Else
               Total2(j) = Total2(j) + BalanceAmount
            End If
            HaveValue = CheckHaveValue(HaveValue, Total1(j))
         Next j
         
         If HaveValue Then
            Call GenerateFooter("", Ac.CUSTOMER_NAME & " (" & Ac.CUSTOMER_CODE & ")", "", Total1, m_Details1, BalanceAmount)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
      Next Bd
      Set TempBD = Nothing
      
      Call GenerateFooter("รวม", "", "", Total2, m_Details1, Total2(9))
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
   
   Set CnBalances = Nothing
   Set DnBalances = Nothing
   Set Accounts = Nothing
   Set DistinctAccounts = Nothing
   Set SellItems = Nothing
   Set PaidItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, BalAmount As Double)
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
      ElseIf j = 9 Then
         TempStr = FormatNumber(BalAmount, "-")
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






