VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalDO003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalDO003"
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
Private FIX_COST As Double
Private DISCOUNT_AMOUNT As Double
Private SUM_DISCOUNT_REASON As String



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

   Vsp.FontBold = True
   Vsp.FontSize = 16
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 3 * 567 + HeaderOffset
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
      
   Vsp.TableCell(tcAlign, 1, 1) = taCenterTop
   Vsp.TableCell(tcText, 1, 1) = glbEnterPrise.ENTERPRISE_NAME & vbCrLf & m_EnpAddress.PackAddress & vbCrLf & "โทร. " & EmptyToString(m_EnpAddress.PHONE1, "N/A") & "     แฟกซ์. " & EmptyToString(m_EnpAddress.FAX1, "N/A") & vbCrLf & "เลขประจำตัวผู้เสียภาษี    " & glbEnterPrise.TAX_ID & vbCrLf
   
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.FontBold = True
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 2
    Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "6.7cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "6cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "2.7cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "3cm"
     Vsp.TableCell(tcFontSize) = 16
  

     Vsp.TableCell(tcAlign, 2, 2) = taCenterMiddle
    Vsp.TableCell(tcText, 2, 2) = "บิลส่งของ"
    Vsp.TableBorder = Border
    
    
    Vsp.TableBorder = Border
    Vsp.EndTable
   

   Vsp.StartTable
    Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 3
    Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "2.7cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "10cm"
     Vsp.TableCell(tcColWidth, 1, 3) = "2.7cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "3cm"
   Vsp.TableCell(tcFontSize) = 14

   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
   Vsp.TableCell(tcText, 1, 1) = "ฟาร์ม/ฝ่าย"
   
   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
   Vsp.TableCell(tcText, 1, 2) = glbEnterPrise.BRANCH_NAME
   
    Vsp.TableCell(tcAlign, 1, 3) = taRightTop
    Vsp.TableCell(tcText, 1, 3) = "เลขที่เอกสาร : "
    
    Vsp.TableCell(tcAlign, 1, 4) = taLeftTop
    Vsp.TableCell(tcText, 1, 4) = m_TempBD.DOCUMENT_NO
   
   Vsp.TableCell(tcAlign, 2, 1) = taLeftTop
   Vsp.TableCell(tcText, 2, 1) = "ชื่อลูกค้า"
   
   Vsp.TableCell(tcAlign, 2, 2) = taLeftTop
    Vsp.TableCell(tcText, 2, 2) = m_TempBD.CUSTOMER_NAME
    
    Vsp.TableCell(tcAlign, 3, 1) = taLeftTop
   Vsp.TableCell(tcText, 3, 1) = "ที่อยู่"
   
   Vsp.TableCell(tcAlign, 3, 2) = taLeftTop
   Vsp.TableCell(tcText, 3, 2) = m_Address.PackAddress
   
   Vsp.TableCell(tcAlign, 2, 3) = taRightTop
   Vsp.TableCell(tcText, 2, 3) = "วันที่ :"
    
    Vsp.TableCell(tcAlign, 2, 4) = taLeftTop
    Vsp.TableCell(tcText, 2, 4) = DateToStringExtEx2(m_TempBD.DOCUMENT_DATE)


    Vsp.TableBorder = Border
    Vsp.EndTable
 
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.3 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
                
    Vsp.TableBorder = Border
    Vsp.EndTable

    Vsp.FontBold = False
End Sub

Private Sub GenerateRcpBody(Bd As CBillingDoc, Border As TableBorderSettings, alngX() As Long)
Dim TempStr As String
Dim IntPart As String
Dim DecimalPart As String
Dim I As Long
Dim Left As Long
Dim MaxRow As Long
Dim Di As CDoItem
Dim Totals(1 To 100) As Double
Dim PigFlag As String
Dim RowHeight As Double


   Vsp.PenStyle = psDash
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 6
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "7.4cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "2.5cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "2.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "2.5cm"

   Vsp.TableCell(tcFontSize) = (Vsp.FontSize + 2)
  
    Vsp.TableCell(tcAlign, 1, 1) = taCenterTop
    Vsp.TableCell(tcText, 1, 1) = "ลำดับ"
    
    Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
     Vsp.TableCell(tcText, 1, 2) = "รายการ"
     
     Vsp.TableCell(tcAlign, 1, 3) = taCenterTop
     Vsp.TableCell(tcText, 1, 3) = "หน่วย"
     
     Vsp.TableCell(tcAlign, 1, 4) = taCenterTop
     Vsp.TableCell(tcText, 1, 4) = "จำนวน"
     
     Vsp.TableCell(tcAlign, 1, 5) = taCenterTop
     Vsp.TableCell(tcText, 1, 5) = "ราคา"
     
     Vsp.TableCell(tcAlign, 1, 6) = taCenterTop
     Vsp.TableCell(tcText, 1, 6) = "จำนวนเงิน"

    Vsp.TableBorder = Border
    Vsp.EndTable
  

   MaxRow = 19
 ' RowHeight = 4.3 * 567 / MaxRow
   Vsp.TableCell(tcFontSize) = Vsp.FontSize


  ' MaxRow = 9
   
   For I = 1 To UBound(Totals)
      Totals(I) = 0
   Next I

   I = 0
   For Each Di In m_Merges
      I = I + 1

        Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = False
     '  Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "7.4cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "2.5cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.5cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "2cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "2.5cm"

      PigFlag = Di.PIG_FLAG
      
         Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 1) = I
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = Di.PART_DESC
         
          Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 3) = Di.UNIT_NAME
         
         Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
         Vsp.TableCell(tcText, 1, 4) = FormatNumber(Di.ITEM_AMOUNT)
         Totals(5) = Totals(5) + Di.ITEM_AMOUNT
         

         Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
         Vsp.TableCell(tcText, 1, 5) = FormatNumber(MyDiffEx(Di.TOTAL_PRICE, Di.ITEM_AMOUNT))
       
       Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
       Vsp.TableCell(tcText, 1, 6) = FormatNumber(Di.TOTAL_PRICE)
       Totals(8) = Totals(8) + Di.TOTAL_PRICE
      
       Vsp.TableBorder = Border
       Vsp.EndTable
    
    Next Di
    
    
     For Each Di In m_TempBD.Revenues
         I = I + 1

        Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = False
       Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "7.4cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "2.5cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.5cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "2cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "2.5cm"
       
        PigFlag = Di.PIG_FLAG
      
         Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 1) = I
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = Di.REVENUE_NAME
       
       

'      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
'      Vsp.TableCell(tcText, 1, 1) = Di.REVENUE_NO
'      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
'
'      Vsp.TableCell(tcText, 1, 2) = Di.REVENUE_NAME
'      Vsp.TableCell(tcText, 1, 3) = ""
'      Vsp.TableCell(tcText, 1, 4) = ""

'         Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
'         Vsp.TableCell(tcText, 1, 4) = FormatNumber(Di.ITEM_AMOUNT)
'         Totals(5) = Totals(5) + Di.ITEM_AMOUNT
      
'      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
'      Vsp.TableCell(tcText, 1, 5) = FormatNumber(Di.ITEM_AMOUNT)
'      Totals(5) = Totals(5) + Di.ITEM_AMOUNT
'      Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
'      Vsp.TableCell(tcText, 1, 6) = ""

'          Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
'         Vsp.TableCell(tcText, 1, 5) = FormatNumber(MyDiffEx(Di.TOTAL_PRICE, Di.ITEM_AMOUNT))
       
       Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
       Vsp.TableCell(tcText, 1, 6) = FormatNumber(Di.TOTAL_PRICE)
       Totals(8) = Totals(8) + Di.TOTAL_PRICE
'      If Di.SHOW_AVG = "Y" Then
'         Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
'         Vsp.TableCell(tcText, 1, 7) = FormatNumberInt(MyDiffEx(Di.TOTAL_PRICE, Di.ITEM_AMOUNT))
'      End If
'      Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
'      Vsp.TableCell(tcText, 1, 8) = FormatNumber(Di.TOTAL_PRICE)
'      Totals(8) = Totals(8) + Di.TOTAL_PRICE
      
       Vsp.TableBorder = Border
       Vsp.EndTable
    Next Di
   '=== Revenue


    Left = MaxRow - I
   For I = 1 To Left
     Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = False
       'Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "7.4cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "2.5cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.5cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "2cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "2.5cm"
      
       Vsp.TableBorder = Border
       Vsp.EndTable
    Next I
   
      
     Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = False
       'Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "14.4cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "2.5cm"
       
       
      
      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
     Vsp.TableCell(tcText, 1, 1) = " รวม"
     
       Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
     Vsp.TableCell(tcText, 1, 2) = "(" & ThaiBaht(Totals(8)) & ")"
    
    Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
     Vsp.TableCell(tcText, 1, 3) = FormatNumber(Totals(8))
     
       Vsp.TableBorder = Border
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
   Vsp.PaperSize = pprA4
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

   m_ImageWidth = 18.4 * 567
   m_ImageHeight = 26.9 * 567

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

Private Function GetWeight(Idx As Long) As String
Dim Di As CDoItem
   
   If m_TempBD.DoItems.Count <= 0 Then
      Exit Function
   End If
   Set Di = m_TempBD.DoItems(1)
   If Di.PIG_FLAG <> "Y" Then
      GetWeight = ""
      Exit Function
   End If
   
   If Idx > m_TempBD.DoItems.Count Then
      GetWeight = ""
   Else
      Set Di = m_TempBD.DoItems(Idx)
      GetWeight = FormatNumber(Di.TOTAL_WEIGHT)
   End If
End Function

Private Sub GenerateRcpItem(Bd As CBillingDoc, Border As TableBorderSettings, alngX() As Long)
Dim I As Long
Dim j As Long
Dim RowHeight As Double
Dim MaxRow As Long
Dim LeftRow As Long
Dim TempStr As String
Dim IntPart As String
Dim DecimalPart As String
Dim Ri As CReceiptItem
Dim CnAmount As Double
Dim DnAmount As Double
Dim LinearIdx As Long

   MaxRow = 17
   RowHeight = 11.9 * 567 / MaxRow
   Vsp.FontSize = 16
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 16
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = "0.8cm"
    Vsp.TableCell(tcColWidth, 1, 1) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 7) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 8) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 9) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 10) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 11) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 12) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 13) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 14) = "1.3cm"
    Vsp.TableCell(tcColWidth, 1, 15) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 16) = "1.3cm"
        
    Vsp.TableBorder = Border
    Vsp.EndTable
   
                  
   I = 1
   m_SumCredit = 0
   For I = 1 To 17
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 16
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = True
       Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "1.0cm"
       
       j = 1
       LinearIdx = 17 * (j - 1) + I
       Vsp.TableCell(tcColWidth, 1, 2) = "1.3cm"
       Vsp.TableCell(tcText, 1, 2) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
       Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
       
       j = 2
       LinearIdx = 17 * (j - 1) + I
       Vsp.TableCell(tcColWidth, 1, 4) = "1.3cm"
       Vsp.TableCell(tcText, 1, 4) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
       
       Vsp.TableCell(tcColWidth, 1, 5) = "1.0cm"
       
      j = 3
      LinearIdx = 17 * (j - 1) + I
      Vsp.TableCell(tcColWidth, 1, 6) = "1.3cm"
       Vsp.TableCell(tcText, 1, 6) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
       
       Vsp.TableCell(tcColWidth, 1, 7) = "1.0cm"
       
      j = 4
      LinearIdx = 17 * (j - 1) + I
       Vsp.TableCell(tcColWidth, 1, 8) = "1.3cm"
       Vsp.TableCell(tcText, 1, 8) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
       
       Vsp.TableCell(tcColWidth, 1, 9) = "1.0cm"
       
       j = 5
       LinearIdx = 17 * (j - 1) + I
       Vsp.TableCell(tcColWidth, 1, 10) = "1.3cm"
       Vsp.TableCell(tcText, 1, 10) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 10) = taRightMiddle
       
       Vsp.TableCell(tcColWidth, 1, 11) = "1.0cm"
       
       j = 6
       LinearIdx = 17 * (j - 1) + I
       Vsp.TableCell(tcColWidth, 1, 12) = "1.3cm"
       Vsp.TableCell(tcText, 1, 12) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 12) = taRightMiddle
       
       Vsp.TableCell(tcColWidth, 1, 13) = "1.0cm"
       
       j = 7
       LinearIdx = 17 * (j - 1) + I
       Vsp.TableCell(tcColWidth, 1, 14) = "1.3cm"
       Vsp.TableCell(tcText, 1, 14) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 14) = taRightMiddle
       
       Vsp.TableCell(tcColWidth, 1, 15) = "1.0cm"
       
       j = 8
       LinearIdx = 17 * (j - 1) + I
       Vsp.TableCell(tcColWidth, 1, 16) = "1.3cm"
       Vsp.TableCell(tcText, 1, 16) = GetWeight(LinearIdx)
       Vsp.TableCell(tcAlign, 1, 16) = taRightMiddle
       
       Vsp.TableBorder = Border
       Vsp.EndTable
   Next I
   m_EndItem = True

'   LeftRow = MaxRow - I
'   For j = 1 To LeftRow
'      Vsp.StartTable
'      Vsp.TableCell(tcCols) = 16
'      Vsp.TableCell(tcRows) = 1
'       Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
'       Vsp.TableCell(tcColWidth, 1, 1) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 2) = "1.3cm"
'       Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 4) = "1.3cm"
'       Vsp.TableCell(tcColWidth, 1, 5) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 6) = "1.3cm"
'       Vsp.TableCell(tcColWidth, 1, 7) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 8) = "1.3cm"
'       Vsp.TableCell(tcColWidth, 1, 9) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 10) = "1.3cm"
'       Vsp.TableCell(tcColWidth, 1, 11) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 12) = "1.3cm"
'       Vsp.TableCell(tcColWidth, 1, 13) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 14) = "1.3cm"
'       Vsp.TableCell(tcColWidth, 1, 15) = "1.0cm"
'       Vsp.TableCell(tcColWidth, 1, 16) = "1.3cm"
'
'       Vsp.TableBorder = Border
'       Vsp.EndTable
'   Next j
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
      
      Call MergeRecord(1)
      If m_Merges.Count > 5 Then
         Set m_Merges = Nothing
         Set m_Merges = New Collection
         Call MergeRecord(2)
      End If
      
      If m_TempBD.BILLING_ADDRESS_ID > 0 Then
         m_Address.ADDRESS_ID = m_TempBD.BILLING_ADDRESS_ID
         Call m_Address.QueryData(Rs, iCount)
         If Not Rs.EOF Then
            Call m_Address.PopulateFromRS(Rs)
         End If
      End If
   End If

   Vsp.StartDoc

'   Call GenerateRcpItem(m_TempBD, tbAll, alngX)
'
'   If m_EndItem Then
      Call GenerateRcpBody(m_TempBD, tbAll, alngX)
'   End If
   
      Call PrintBillFooter(alngX)
   
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

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   Vsp.FontSize = DETAIL_SIZE
'   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
'   strPrint = ";"
'   Call Vsp.AddTable(strFormat, "", strPrint)
'
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = "ได้รับสินค้าตามรายการข้างบนครบถ้วนเรียบร้อยแล้ว;"
  Call Vsp.AddTable(strFormat, "", strPrint)
'
   
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strPrint = strPrint & "..................................||.....................................;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strPrint = strPrint & "(.........../............../............)||(.........../............../............);"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strPrint = "ผู้รับของ||ผู้ส่งของ;"
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

Private Function IsExist(Di As CDoItem, Optional TypeMerge As Long = 1) As Boolean
Dim TmpDi As CDoItem

   IsExist = False
   For Each TmpDi In m_Merges
      If TypeMerge = 1 Then
         If (Di.PIG_STATUS = TmpDi.PIG_STATUS) And (Di.PKG_TYPE = TmpDi.PKG_TYPE) _
         And (Di.PIG_TYPE = TmpDi.PIG_TYPE) And (Format(Di.AVG_PRICE, "0") = Format(TmpDi.AVG_PRICE, "0")) Then
            
            TmpDi.ITEM_AMOUNT = TmpDi.ITEM_AMOUNT + Di.ITEM_AMOUNT
            TmpDi.TOTAL_PRICE = TmpDi.TOTAL_PRICE + Di.TOTAL_PRICE - Di.PEDIGREE_COST
            TmpDi.TOTAL_WEIGHT = TmpDi.TOTAL_WEIGHT + Di.TOTAL_WEIGHT
            
            IsExist = True
            Exit Function
         End If
      ElseIf TypeMerge = 2 Then
         If (Di.PIG_STATUS = TmpDi.PIG_STATUS) And (Di.PKG_TYPE = TmpDi.PKG_TYPE) _
         And (Di.PIG_TYPE = TmpDi.PIG_TYPE) Then
            
            TmpDi.ITEM_AMOUNT = TmpDi.ITEM_AMOUNT + Di.ITEM_AMOUNT
            TmpDi.TOTAL_PRICE = TmpDi.TOTAL_PRICE + Di.TOTAL_PRICE - Di.PEDIGREE_COST
            TmpDi.TOTAL_WEIGHT = TmpDi.TOTAL_WEIGHT + Di.TOTAL_WEIGHT
            
            IsExist = True
            Exit Function
         End If
      
      End If
   Next TmpDi
End Function
Private Sub MergeRecord(Optional TypeMerge As Long = 1)
Dim Di As CDoItem
Dim TempDi As CDoItem
   
   FIX_COST = 0
   SUM_DISCOUNT_REASON = ""
   DISCOUNT_AMOUNT = 0
   
   For Each Di In m_TempBD.DoItems
      If Di.Flag <> "D" Then
         FIX_COST = FIX_COST + Di.PEDIGREE_COST
         If Di.DISCOUNT_REASON <> "" Then
            SUM_DISCOUNT_REASON = SUM_DISCOUNT_REASON & " " & Di.DISCOUNT_REASON
         End If
         If Di.DISCOUNT_AMOUNT > 0 Then
            DISCOUNT_AMOUNT = DISCOUNT_AMOUNT + Di.DISCOUNT_AMOUNT
         End If
         If Not IsExist(Di, TypeMerge) Then
            Set TempDi = New CDoItem
            TempDi.PART_ITEM_ID = Di.PART_ITEM_ID
            TempDi.PIG_STATUS = Di.PIG_STATUS
            TempDi.ITEM_AMOUNT = Di.ITEM_AMOUNT
            TempDi.TOTAL_PRICE = Di.TOTAL_PRICE - Di.PEDIGREE_COST
            TempDi.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
            TempDi.PART_NO = Di.PART_NO
            TempDi.PART_DESC = Di.PART_DESC
            TempDi.PIG_STATUS_NAME = Di.PIG_STATUS_NAME
            TempDi.PIG_STATUS_NO = Di.PIG_STATUS_NO
            TempDi.PIG_FLAG = Di.PIG_FLAG
            TempDi.PKG_TYPE = Di.PKG_TYPE
            TempDi.PIG_TYPE = Di.PIG_TYPE
             TempDi.UNIT_NAME = Di.UNIT_NAME
            If Di.TOTAL_WEIGHT > 0 Then
               'TempDi.AVG_PRICE = (Di.TOTAL_PRICE - Di.PEDIGREE_COST) / Di.TOTAL_WEIGHT
               TempDi.AVG_PRICE = Di.AVG_PRICE
            End If
            TempDi.SHOW_AVG = Di.SHOW_AVG
            
            
            Call m_Merges.Add(TempDi)
            Set TempDi = Nothing
         End If
      End If
   Next Di
      
End Sub
