VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalRcp002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalRcp002"
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
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

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

Private m_SupplierTypes As Collection
Private m_PaidAmounts As Collection
Private m_MaxMins As Collection
Private m_ApSupAmounts As Collection
Private m_ApSupPaids As Collection

Private m_ImageWidth As Double
Private m_ImageHeight As Double
Private m_LedgerReport As Collection
Private m_TempCollection As Collection
Private m_TempBD As CBillingDoc
Private m_Total1(100)  As Double
Private m_Total2(100)  As Double
Private m_EndItem As Boolean
Private m_PaidAmount As Double
Private m_SumDebit As Double
Private m_SumCredit As Double
Private m_CnAmount As Double

Public m_Cr As Collection
Public m_Address As CAddress
Private m_Merges As Collection
'Public OKClick As Boolean
Private ModePrint As Long
Private X1 As Long
Private X2  As Long
Private Y1  As Long
Private Y2  As Long

Private Sub Class_Initialize()
   m_EndItem = False
   
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
   Set m_TempCollection = New Collection
   Set m_Address = New CAddress
   Set m_Merges = New Collection
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
   Set m_TempCollection = Nothing
   Set m_Address = Nothing
   Set m_Merges = Nothing
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

Private Function GenerateReceiptItem(Bd As CBillingDoc, PaidAmount As Double, CnAmount As Double) As String
Dim Ri As CReceiptItem
Dim RCnDn As CReceiptCnDn
Dim TempStr As String
Dim I As Long
Dim j As Long
Dim TempPaid As Double
Dim TempCn As Double

   TempPaid = 0
   I = Bd.ReceiptItems.Count
   j = 0
   
   TempStr = ""
   For Each Ri In Bd.ReceiptItems
      j = j + 1
      TempPaid = TempPaid + Ri.PAID_AMOUNT
      If j = I Then
         TempStr = TempStr & Ri.DOCUMENT_NO
      Else
         TempStr = TempStr & Ri.DOCUMENT_NO & ", "
      End If
   Next Ri
   
   For Each RCnDn In Bd.ReceiptCnDns
      j = j + 1
      TempCn = TempCn + RCnDn.CN_AMOUNT
   Next RCnDn
   
   PaidAmount = TempPaid
   CnAmount = TempCn
   GenerateReceiptItem = TempStr
End Function
Private Sub GenerateRcpHeader(Border As TableBorderSettings, alngX() As Long, HeaderOffset)
Dim OldFont As Double
Dim TempStr As String
Dim PaidAmount As Double
Dim Pmt As PAYMENT_TYPE
Dim CnAmount As Double
Dim iCount As Long
Dim TempX As Long
Dim TempY As Long
Dim TaxBranchString As String

   Vsp.FontBold = True
   Vsp.FontSize = 16
      
   TempX = Vsp.CurrentX
   TempY = Vsp.CurrentY
      
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.5 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
                                                         
   Dim m_EnpAddress As CAddress
   Set m_EnpAddress = New CAddress
   If m_TempBD.ENTERPRISE_ADDRESS_ID > 0 Then
      m_EnpAddress.ADDRESS_ID = m_TempBD.ENTERPRISE_ADDRESS_ID
      Call m_EnpAddress.QueryData(Rs, iCount)
      If Not Rs.EOF Then
         Call m_EnpAddress.PopulateFromRS(Rs)
      End If
   End If
      
   Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
   Vsp.TableCell(tcText, 1, 1) = glbEnterPrise.ENTERPRISE_NAME & vbCrLf & "DEVELOPMENT THAI SWINE BREEDER CO.,LTD "
   
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.FontSize = 12
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 2 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
    Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
    Vsp.TableCell(tcText, 1, 1) = m_EnpAddress.PackAddress & vbCrLf & "18/1 MOO 10 SONGKHON , KAENGKHOI DISTRICT, SARABURI 18110 " & vbCrLf & "โทร / TEL " & EmptyToString(m_EnpAddress.PHONE1, "N/A") & "     แฟกซ์ / FAX " & EmptyToString(m_EnpAddress.FAX1, "N/A") & vbCrLf & "เลขประจำตัวผู้เสียภาษี / TAX ID " & glbEnterPrise.TAX_ID & vbCrLf
   
    Vsp.TableBorder = Border
    Vsp.EndTable
    
    Vsp = ""
    
   X1 = Vsp.CurrentX
   Y1 = Vsp.CurrentY
   
   Vsp.CurrentX = TempX
   Vsp.CurrentY = TempY
   
   Vsp.FontSize = 20
   
   Vsp.FontBold = True
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 3.5 * 567 + HeaderOffset
   Vsp.TableCell(tcColWidth, 1, 1) = "10cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "8.7cm"
   
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 2) = "ใบเสร็จรับเงิน" & vbCrLf & "RECEIPT"
    
    Vsp.TableBorder = Border
    Vsp.EndTable
   
   Vsp.CurrentX = X1
   Vsp.CurrentY = Y1
   
   Vsp.FontSize = 16
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 2
   Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "3.9cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "7.8cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "3.7cm"
   Vsp.TableCell(tcColWidth, 1, 4) = "3cm"
   Vsp.TableCell(tcFontSize) = 14

   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
   Vsp.TableCell(tcText, 1, 1) = "รหัสลูกค้า/Customer Code"
   
   If Len(m_TempBD.TAX_ID) > 0 Then
      TaxBranchString = "TAX ID : " & m_TempBD.TAX_ID
   End If
   If Len(m_TempBD.CUSTOMER_BRANCH) > 0 Then
      TaxBranchString = TaxBranchString & " " & m_TempBD.CUSTOMER_BRANCH
   End If
   If Len(Trim(TaxBranchString)) > 0 Then
      TaxBranchString = " (" & TaxBranchString & ")"
   End If
   
   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
   Vsp.TableCell(tcText, 1, 2) = m_TempBD.CUSTOMER_CODE & TaxBranchString
'   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
'   Vsp.TableCell(tcText, 1, 2) = m_TempBD.CUSTOMER_CODE
   
    Vsp.TableCell(tcAlign, 1, 3) = taRightTop
    Vsp.TableCell(tcText, 1, 3) = "เลขที่เอกสาร/NO : "
    
    Vsp.TableCell(tcAlign, 1, 4) = taLeftTop
    Vsp.TableCell(tcText, 1, 4) = m_TempBD.DOCUMENT_NO
   
   Vsp.TableCell(tcAlign, 2, 1) = taLeftTop
   Vsp.TableCell(tcText, 2, 1) = "ชื่อลูกค้า / Customer Name"
   
   Vsp.TableCell(tcAlign, 2, 2) = taLeftTop
    Vsp.TableCell(tcText, 2, 2) = m_TempBD.CUSTOMER_NAME
    
   Vsp.TableCell(tcAlign, 2, 3) = taRightTop
   Vsp.TableCell(tcText, 2, 3) = "วันที่ / DATE : "
   
    Vsp.TableCell(tcAlign, 2, 4) = taLeftTop
    Vsp.TableCell(tcText, 2, 4) = DateToStringExtEx2(m_TempBD.DOCUMENT_DATE)


    Vsp.TableBorder = Border
    Vsp.EndTable
 
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 2
   Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567 + HeaderOffset
   Vsp.TableCell(tcColWidth, 1, 1) = "3.9cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "14.5cm"
   Vsp.TableCell(tcFontSize) = 14

   
    Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
   Vsp.TableCell(tcText, 1, 1) = "ที่อยู่ / Address"
   
   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
   Vsp.TableCell(tcText, 1, 2) = m_Address.PackAddress

    Vsp.TableBorder = Border
    Vsp.EndTable
    
    
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 1
'   Vsp.TableCell(tcRows) = 1
'    Vsp.TableCell(tcRowHeight, 1) = 0.3 * 567 + HeaderOffset
'    Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
'
'    Vsp.TableBorder = Border
'    Vsp.EndTable

    Vsp.FontBold = False
End Sub

Private Sub GenerateRcpBody(Bd As CBillingDoc, Border As TableBorderSettings, alngX() As Long)
Dim TempStr As String
Dim IntPart As String
Dim DecimalPart As String
Dim I As Long
Dim Left As Long
Dim MaxRow As Long
Dim Ri As CReceiptItem
Dim Totals(1 To 100) As Double
Dim PigFlag As String
Dim RowHeight As Double
Dim CnAmount  As Double
Dim DnAmount As Double

   Vsp.FontSize = 14
   
   Vsp.PenStyle = psDash
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.3 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "14.9cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "3cm"
  
   Vsp.TableCell(tcAlign, 1, 1) = taCenterTop
   Vsp.TableCell(tcText, 1, 1) = "ลำดับ" & vbCrLf & "Item"
   
   Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
   Vsp.TableCell(tcText, 1, 2) = "รายการ" & vbCrLf & "Description"
   
   Vsp.TableCell(tcAlign, 1, 3) = taCenterTop
   Vsp.TableCell(tcText, 1, 3) = "ยอดชำระ" & vbCrLf & "Amount"
   
   Vsp.TableBorder = Border
   Vsp.EndTable
  
   X1 = Vsp.CurrentX
   Y1 = Vsp.CurrentY
   
   MaxRow = 12
 ' RowHeight = 4.3 * 567 / MaxRow
   Vsp.TableCell(tcFontSize) = Vsp.FontSize


  ' MaxRow = 9
   
   For I = 1 To UBound(Totals)
      Totals(I) = 0
   Next I

   I = 0
   
   
   For Each Ri In Bd.ReceiptItems
      I = I + 1
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = False
     '  Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
      Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "14.9cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "3cm"
      
       Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = I
      
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 2) = "INVOICE " & Ri.DOCUMENT_NO
      
      'Call GetRelateCnDn(Bd, Ri.DO_ID, CnAmount, DnAmount)
       
'      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
'      Vsp.TableCell(tcText, 1, 3) = FormatNumber(Ri.PAID_AMOUNT + CnAmount - DnAmount)
'      Totals(3) = Totals(3) + Ri.PAID_AMOUNT + CnAmount - DnAmount
'
'      Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
'      Vsp.TableCell(tcText, 1, 4) = FormatNumber(CnAmount)
'      Totals(4) = Totals(4) + CnAmount
'
'      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
'      Vsp.TableCell(tcText, 1, 5) = FormatNumber(DnAmount)
'      Totals(5) = Totals(5) + DnAmount
      
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = FormatNumber(Ri.PAID_AMOUNT)
      Totals(3) = Totals(3) + Ri.PAID_AMOUNT
      
       
       Vsp.TableBorder = tbColumns
       Vsp.EndTable
   Next Ri
    
   
   Left = MaxRow - I
   For I = 1 To Left
     Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = False
       'Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
      Vsp.TableCell(tcColWidth, 1, 2) = "14.9cm"
      Vsp.TableCell(tcColWidth, 1, 3) = "3cm"
      
       Vsp.TableBorder = tbColumns
       Vsp.EndTable
    Next I
      
   X2 = Vsp.CurrentX
   Y2 = Vsp.CurrentY
   
   Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
   Call Vsp.DrawLine(Vsp.MarginLeft + (19.4 * 567), Y1, Vsp.MarginLeft + (19.4 * 567), Y2)
'   Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
   Vsp.TableBorder = tbNone
   Y1 = Vsp.CurrentY
      
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcFontBold, 1) = False
   'Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
   Vsp.TableCell(tcColWidth, 1, 1) = "16.4cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3cm"
   
   Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 1) = "(" & ThaiBaht(Totals(3)) & " / " & ThaiBahtEng(Totals(3)) & ")"
   
   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
   Vsp.TableCell(tcText, 1, 2) = FormatNumber(Totals(3))
   
   Vsp.TableBorder = tbAll
   Vsp.EndTable
    'Vsp.TableCell(tcText, 1, 4) = "(" & ThaiBaht(ReceiptAmount - GeReceiptDiscount) & ")"
    
End Sub

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
   Vsp.TableBorder = tbNone
'    Vsp.TableBorder = tbAll

   Call GenerateRcpHeader(tbNone, alngX, 0)
  'Call GenerateRcpHeader(tbAll, alngX, 0)
   
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

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprLetter
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))

'   m_ImageWidth = 18.4 * 567
'   m_ImageHeight = 26.9 * 567

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub GetRelateCnDn(Bd As CBillingDoc, DoID As Long, CnAmount As Double, DnAmount As Double)
Dim RCnDn As CReceiptCnDn
Dim TempCn As Double
Dim TempDn As Double

   TempCn = 0
   TempDn = 0
   For Each RCnDn In Bd.ReceiptCnDns
      If RCnDn.DO_ID = DoID Then
         TempCn = TempCn + RCnDn.CN_AMOUNT
         TempDn = TempDn + RCnDn.DN_AMOUNT
      End If
   Next RCnDn
   
   CnAmount = TempCn
   DnAmount = TempDn
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
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim DocNum As String
Dim ReportType As Long
Static AlreadyStart As Long
Static AlreadyNewPage As Long
Dim Bd As CBillingDoc

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
   
   For j = 1 To UBound(m_Total1)
      m_Total1(j) = 0
      m_Total2(j) = 0
   Next j
   
   Set m_TempBD = New CBillingDoc
   m_TempBD.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_TempBD.QueryFlag = 1
   Call glbDaily.QueryBillingDoc(m_TempBD, Rs, iCount, IsOK, glbErrorLog)
   If Not Rs.EOF Then
      Call m_TempBD.PopulateFromRS(1, Rs)
      
      If m_TempBD.BILLING_ADDRESS_ID > 0 Then
         m_Address.ADDRESS_ID = m_TempBD.BILLING_ADDRESS_ID
         Call m_Address.QueryData(Rs, iCount)
         If Not Rs.EOF Then
            Call m_Address.PopulateFromRS(Rs)
         End If
      End If
   End If
   
   ModePrint = 1
   
   Vsp.StartDoc

   Call GenerateRcpBody(m_TempBD, tbAll, alngX)

   Call PrintBillFooterRcp(alngX)
   
   Vsp.EndDoc
   
   Set m_TempBD = Nothing
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
'Private Sub printHeader()
'   Dim strFormat As String
'   Dim strPrint As String
'   Dim tbdT As TableBorderSettings
'   Dim blnBold As Boolean
'   Dim iSize As Integer
'   Dim alngX() As Long
'
'   tbdT = Vsp.TableBorder
'   blnBold = Vsp.FontBold
'   iSize = Vsp.FontSize
'
''   Call PrintBillHeader(alngX)
'
'   Vsp.FontSize = HEADER_SIZE
'   Vsp.FontBold = True
'   Vsp.TableBorder = tbAll
'
'
'
'   strFormat = VSP_CalTable(HEADER_RECEIPT1, mdblWid, alngX)
'   strPrint = "No|รายการ|จำนวน|ราคาต่อหน่วย|จำนวนเงิน;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
'
'   Vsp.TableBorder = tbdT
'   Vsp.FontBold = blnBold
'   Vsp.FontSize = iSize
'End Sub
Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
   
   Vsp = ""
   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   Vsp.FontSize = 12
'   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
'   strPrint = ";"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "ได้รับสินค้าตามรายการข้างบนครบถ้วนเรียบร้อยแล้ว;"
  Call Vsp.AddTable(strFormat, "", strPrint)
'
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "RECEIVED THE ABOVE MENTIONED GOODS IN GOOD ORDER AND CONDITION;"
  Call Vsp.AddTable(strFormat, "", strPrint)
  
  strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "เลขที่บัญชีธนาคารสำหรับการชำระเงิน / ACCOUNT FOR PAYMENT;"
  Call Vsp.AddTable(strFormat, "", strPrint)
'
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "   -  บัญชีกระแสรายวัน ธนาคารกรุงไทย ชื่อบัญชี บจก. พันธุ์สุกรไทยพัฒนา เลขที่บัญชี 076-6-01826-1 ;"
  Call Vsp.AddTable(strFormat, "", strPrint)
  
  strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "   -  บัญชีกระแสรายวัน ธนาคารกสิกรไทย ชื่อบัญชี บจก. พันธุ์สุกรไทยพัฒนา เลขที่บัญชี 046-1-09245-4 ;"
  Call Vsp.AddTable(strFormat, "", strPrint)
  
  Vsp = ""
  
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = "..................................||.....................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strPrint = "(.........../............../............)||(.........../............../............);"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strPrint = "ผู้รับของ RECEIVED||ผู้ส่งของ DELIVERY BY;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub
Private Sub PrintBillFooterRcp(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
   
   Vsp = ""
   
   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   Vsp.FontSize = 12
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
  strPrint = "หมายเหตุ  " & m_TempBD.NOTE & ";"
  Call Vsp.AddTable(strFormat, "", strPrint)
  
  Vsp = ""
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "ชำระโดย / PAID BY;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "เงินสด / CASH.......................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strFormat = VSP_CalTable("<15|<10|<10|<12;", mdblWid, alngX)
   strPrint = "เช็คธนาคาร / BANK CHEQUE ....................|เช็คเลขที่ / NO ..........................|ลงวันที่ / DATE......../......../........|จำนวนเงิน / AMOUNT ....................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   Vsp = ""
   
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = "............................................||............................................;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strPrint = strPrint & "(.........../............../............)||(.........../............../............);"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strPrint = "ผู้รับเงิน RECEIVER||ผู้มีอำนาจลงนาม AUTHORIZED BY;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
Dim alngX() As Long

   If Not mblnEndPage Then Exit Sub
  
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
   
'   If mcolParam("PICTURE_FLAG") = "Y" Then
'      Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
'   End If
   
   If mblnHeader Then
      Call printHeader
   End If
   
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

