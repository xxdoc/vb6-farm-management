VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalDO001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalDO001"
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
    Vsp.TableCell(tcRowHeight, 1) = 2.8 * 567 + HeaderOffset
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
   Vsp.TableCell(tcText, 1, 1) = glbEnterPrise.ENTERPRISE_NAME & vbCrLf & m_EnpAddress.PackAddress & vbCrLf & "โทร. " & EmptyToString(m_EnpAddress.PHONE1, "N/A") & "     แฟกซ์. " & EmptyToString(m_EnpAddress.FAX1, "N/A") & "    เลขประจำตัวผู้เสียภาษี    " & glbEnterPrise.TAX_ID & vbCrLf
   
    Vsp.TableBorder = Border
    Vsp.EndTable
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.5 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "16.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "2.8cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taRightTop
    Vsp.TableCell(tcText, 1, 2) = m_TempBD.DOCUMENT_NO
    Vsp.TableBorder = Border
    Vsp.EndTable
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
    Vsp.TableBorder = Border
    Vsp.EndTable
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 6
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.7 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "1.2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "8.6cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "1.6cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "2.4cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "0.7cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "3.9cm"
    
    Vsp.TableCell(tcText, 1, 2) = m_TempBD.CUSTOMER_NAME
    Vsp.TableCell(tcText, 1, 4) = glbEnterPrise.BRANCH_NAME
    Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
    Vsp.TableCell(tcText, 1, 6) = DateToStringExtEx2(m_TempBD.DOCUMENT_DATE)
    Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
    
    Vsp.TableBorder = Border
    Vsp.EndTable
        
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 7
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "1.2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "8.6cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "1.6cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "0.8cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "1.9cm"
    Vsp.TableCell(tcColWidth, 1, 7) = "2.8cm"
        
    Vsp.TableCell(tcText, 1, 2) = m_Address.PackAddress
    Vsp.TableCell(tcText, 1, 4) = m_TempBD.Credit
    Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
    
    Vsp.TableBorder = Border
    Vsp.EndTable

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 7
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "1.2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "8.6cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "1.6cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "0.8cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "1.9cm"
    Vsp.TableCell(tcColWidth, 1, 7) = "2.8cm"

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

Private Sub GenerateRcpFooter(Bd As CBillingDoc, Border As TableBorderSettings, alngX() As Long)
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

   MaxRow = 6
   RowHeight = 4.3 * 567 / MaxRow
   
   Vsp.FontSize = 16
   Vsp.StartTable
   Vsp.TableCell(tcFontBold, 1) = True
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcFontSize) = 16
   
    Vsp.TableCell(tcRowHeight) = "1.4cm"
    Vsp.TableCell(tcColWidth, 1, 1) = "18.4cm"
    
    Vsp.EndTable
       
   MaxRow = 5
   
   For I = 1 To UBound(Totals)
      Totals(I) = 0
   Next I
   
   I = 0
   For Each Di In m_Merges
      I = I + 1
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 8
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = True
       Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "2.3cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "6.6cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 7) = "1.2cm"
       Vsp.TableCell(tcColWidth, 1, 8) = "2.3cm"

      PigFlag = Di.PIG_FLAG
      If Di.PIG_FLAG = "Y" Then
         Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 1) = Di.PIG_STATUS_NO
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = Di.PIG_STATUS_NAME
         Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
         If mcolParam("OPTION_MODE") > 0 Then
            Vsp.TableCell(tcText, 1, 1) = ""
            Vsp.TableCell(tcText, 1, 2) = "ขุน"
         End If
         
         Vsp.TableCell(tcText, 1, 3) = FormatNumberInt(Di.ITEM_AMOUNT)
         Totals(3) = Totals(3) + Di.ITEM_AMOUNT
         Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
         Vsp.TableCell(tcText, 1, 4) = FormatNumber(Di.TOTAL_WEIGHT)
         Totals(4) = Totals(4) + Di.TOTAL_WEIGHT
         
         Vsp.TableCell(tcText, 1, 5) = ""
         Vsp.TableCell(tcText, 1, 6) = ""
         If Di.SHOW_AVG = "Y" Then
            Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
            Vsp.TableCell(tcText, 1, 7) = FormatNumber(Di.AVG_PRICE)
         End If
      Else
         Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 1) = Di.PART_NO
         Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
         Vsp.TableCell(tcText, 1, 2) = Di.PART_DESC
         Vsp.TableCell(tcText, 1, 3) = ""
         Vsp.TableCell(tcText, 1, 4) = ""
         
         Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
         Vsp.TableCell(tcText, 1, 5) = FormatNumber(Di.ITEM_AMOUNT)
         Totals(5) = Totals(5) + Di.ITEM_AMOUNT
         Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
         Vsp.TableCell(tcText, 1, 6) = ""
         If Di.SHOW_AVG = "Y" Then
            Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
            Vsp.TableCell(tcText, 1, 7) = FormatNumber(MyDiffEx(Di.TOTAL_PRICE, Di.ITEM_AMOUNT))
         End If
      End If
      Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
      Vsp.TableCell(tcText, 1, 8) = FormatNumber(Di.TOTAL_PRICE)
      Totals(8) = Totals(8) + Di.TOTAL_PRICE
      
       Vsp.TableBorder = Border
       Vsp.EndTable
    
    Next Di
    
    If FIX_COST > 0 Then
      I = I + 1
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 8
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = True
       Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "2.3cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "6.6cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 7) = "1.2cm"
       Vsp.TableCell(tcColWidth, 1, 8) = "2.3cm"

      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = ""
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 2) = "ค่าคงที่"
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = ""
      Vsp.TableCell(tcText, 1, 4) = ""
      
      Vsp.TableCell(tcText, 1, 5) = ""
      
      Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 6) = ""
      Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
      Vsp.TableCell(tcText, 1, 7) = ""
      
      Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
      Vsp.TableCell(tcText, 1, 8) = FormatNumber(FIX_COST)
      Totals(8) = Totals(8) + FIX_COST
   
      Vsp.TableBorder = Border
      Vsp.EndTable
    
    End If
    
   '=== Revenue
   For Each Di In m_TempBD.Revenues
      I = I + 1
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 8
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = True
       Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "2.3cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "6.6cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 7) = "1.2cm"
       Vsp.TableCell(tcColWidth, 1, 8) = "2.3cm"

      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = Di.REVENUE_NO
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle

      Vsp.TableCell(tcText, 1, 2) = Di.REVENUE_NAME
      Vsp.TableCell(tcText, 1, 3) = ""
      Vsp.TableCell(tcText, 1, 4) = ""
      
      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
      Vsp.TableCell(tcText, 1, 5) = FormatNumber(Di.ITEM_AMOUNT)
      Totals(5) = Totals(5) + Di.ITEM_AMOUNT
      Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 6) = ""
      If Di.SHOW_AVG = "Y" Then
         Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
         Vsp.TableCell(tcText, 1, 7) = FormatNumberInt(MyDiffEx(Di.TOTAL_PRICE, Di.ITEM_AMOUNT))
      End If
      Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
      Vsp.TableCell(tcText, 1, 8) = FormatNumber(Di.TOTAL_PRICE)
      Totals(8) = Totals(8) + Di.TOTAL_PRICE
      
       Vsp.TableBorder = Border
       Vsp.EndTable
    Next Di
   '=== Revenue
   
   If DISCOUNT_AMOUNT > 0 Then
      I = I + 1
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 8
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = True
       Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "2.3cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "6.6cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 7) = "1.2cm"
       Vsp.TableCell(tcColWidth, 1, 8) = "2.3cm"

      Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 1) = "ส่วนลด"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 2) = SUM_DISCOUNT_REASON
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = ""
      Vsp.TableCell(tcText, 1, 4) = ""
      
      Vsp.TableCell(tcText, 1, 5) = ""
      
      Vsp.TableCell(tcAlign, 1, 6) = taCenterMiddle
      Vsp.TableCell(tcText, 1, 6) = ""
      Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
      Vsp.TableCell(tcText, 1, 7) = ""
      
      Vsp.TableCell(tcAlign, 1, 8) = taRightMiddle
      Vsp.TableCell(tcText, 1, 8) = FormatNumber(DISCOUNT_AMOUNT)
   
      Vsp.TableBorder = Border
      Vsp.EndTable
   
   End If
   
   
    Left = MaxRow - I
   For I = 1 To Left
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 8
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontBold, 1) = True
       Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
       Vsp.TableCell(tcColWidth, 1, 1) = "2.3cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "6.6cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "1.0cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "2.0cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 7) = "1.2cm"
       Vsp.TableCell(tcColWidth, 1, 8) = "2.3cm"
      
       Vsp.TableBorder = Border
       Vsp.EndTable
    Next I

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 7
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcFontBold, 1) = True
    Vsp.TableCell(tcRowHeight, 1) = RowHeight '"0.8cm"
    Vsp.TableCell(tcColWidth, 1, 1) = "8.9cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1.0cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "2.0cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "1.2cm"
    Vsp.TableCell(tcColWidth, 1, 7) = "2.3cm"
    
   Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 1) = "" '"(" & ThaiBaht(Totals(8)) & ")"
   If PigFlag = "Y" Then
      Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
      Vsp.TableCell(tcText, 1, 2) = FormatNumberInt(Totals(3))
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = FormatNumber(Totals(4))
      Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
      Vsp.TableCell(tcText, 1, 4) = ""
      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
      Vsp.TableCell(tcText, 1, 5) = ""
      Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
      Vsp.TableCell(tcText, 1, 6) = FormatNumber(MyDiffEx(Totals(8) - DISCOUNT_AMOUNT, Totals(4)))
      Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
      Vsp.TableCell(tcText, 1, 7) = FormatNumber(Totals(8) - DISCOUNT_AMOUNT)
   Else
      Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
      Vsp.TableCell(tcText, 1, 2) = ""
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = ""
      Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
      Vsp.TableCell(tcText, 1, 4) = FormatNumber(Totals(5))
      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
      Vsp.TableCell(tcText, 1, 5) = ""
      Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
      Vsp.TableCell(tcText, 1, 6) = FormatNumber(MyDiffEx(Totals(8) - DISCOUNT_AMOUNT, Totals(5)))
      Vsp.TableCell(tcAlign, 1, 7) = taRightMiddle
      Vsp.TableCell(tcText, 1, 7) = FormatNumber(Totals(8) - DISCOUNT_AMOUNT)
   End If
   
    Vsp.TableBorder = Border
    Vsp.EndTable
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

   Call GenerateRcpHeader(tbNone, alngX, 0)
   
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
      
      If mcolParam("OPTION_MODE") > 0 Then
         Call OrderRecord
      End If
      
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

   Call GenerateRcpItem(m_TempBD, tbNone, alngX)

   If m_EndItem Then
      Call GenerateRcpFooter(m_TempBD, tbNone, alngX)
   End If
   
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
      If mcolParam("OPTION_MODE") > 0 Then
         TmpDi.ITEM_AMOUNT = TmpDi.ITEM_AMOUNT + Di.ITEM_AMOUNT
         TmpDi.TOTAL_PRICE = TmpDi.TOTAL_PRICE + Di.TOTAL_PRICE - Di.PEDIGREE_COST
         TmpDi.TOTAL_WEIGHT = TmpDi.TOTAL_WEIGHT + Di.TOTAL_WEIGHT
         
         IsExist = True
         Exit Function
      ElseIf TypeMerge = 1 Then
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
Private Sub OrderRecord()
Dim Di As CDoItem
Dim TempDi As CAdjustPigWeight
Dim sumWeight As Double
Dim AvgWeight As Double
Dim CollFromAsc As Collection
Dim CollFromDesc As Collection
Dim I As Long
Dim TempDiFind As CAdjustPigWeight
Dim TempDiAdjust As CAdjustPigWeight
Dim PrevWeight As Double
Dim TempKey As String
Dim Flag As String

Dim SumPool As Double
Dim MaxFitScale As Double

Dim RunningSwing As Long

Dim SumWeightHigh As Double
Dim SumWeightLow As Double
   
Dim Round As Long
Dim SumAdjust As Double

   Set CollFromDesc = New Collection
   Set CollFromAsc = New Collection
   sumWeight = 0
   I = 0
   For Each Di In m_TempBD.DoItems
      I = I + 1
      sumWeight = sumWeight + Di.TOTAL_WEIGHT
      If Di.TOTAL_WEIGHT > 120 Then
         SumWeightHigh = SumWeightHigh + (Di.TOTAL_WEIGHT - 120)
      ElseIf Di.TOTAL_WEIGHT < 90 Then
         SumWeightLow = SumWeightLow + (90 - Di.TOTAL_WEIGHT)
      End If
      
      Flag = ""
      PrevWeight = 0
      Set TempDi = New CAdjustPigWeight
      TempDi.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
      TempDi.REAL_INDEX = I
      
      'จากมากไปน้อย
      For Each TempDiFind In CollFromDesc
         If Di.TOTAL_WEIGHT >= TempDiFind.TOTAL_WEIGHT Then
            PrevWeight = TempDiFind.TOTAL_WEIGHT
            If Di.TOTAL_WEIGHT = TempDiFind.TOTAL_WEIGHT Then
               TempKey = Trim(Di.TOTAL_WEIGHT & "-" & I)
               Flag = "AT"
            Else
               TempKey = Trim(Str(Di.TOTAL_WEIGHT))
            End If
            Exit For
         End If
      Next TempDiFind
     
      If PrevWeight > 0 Then
         If Flag = "AT" Then
            Call CollFromDesc.Add(TempDi, TempKey, , Trim(Str(PrevWeight)))
         Else
            Call CollFromDesc.Add(TempDi, TempKey, Trim(Str(PrevWeight)))
         End If
      Else
         Call CollFromDesc.Add(TempDi, Trim(Str(Di.TOTAL_WEIGHT)))
      End If
      'จากมากไปน้อย
      
      
      'น้อยไปมาก
      Flag = ""
      PrevWeight = 0
      Set TempDi = New CAdjustPigWeight
      TempDi.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
      TempDi.REAL_INDEX = I
      
      For Each TempDiFind In CollFromAsc
         If Di.TOTAL_WEIGHT <= TempDiFind.TOTAL_WEIGHT Then
            PrevWeight = TempDiFind.TOTAL_WEIGHT
            If Di.TOTAL_WEIGHT = TempDiFind.TOTAL_WEIGHT Then
               TempKey = Trim(Di.TOTAL_WEIGHT & "-" & I)
               Flag = "AT"
            Else
               TempKey = Trim(Str(Di.TOTAL_WEIGHT))
            End If
            Exit For
         End If
      Next TempDiFind
     
      If PrevWeight > 0 Then
         If Flag = "AT" Then
            Call CollFromAsc.Add(TempDi, TempKey, , Trim(Str(PrevWeight)))
         Else
            Call CollFromAsc.Add(TempDi, TempKey, Trim(Str(PrevWeight)))
         End If
      Else
         Call CollFromAsc.Add(TempDi, Trim(Str(Di.TOTAL_WEIGHT)))
      End If
      'น้อยไปมาก
   Next Di
   
   AvgWeight = MyDiff(sumWeight, m_TempBD.DoItems.Count)
   'Debug.Print SumWeightHigh & "-" & SumWeight115_120 & "-" & SumWeight110_114 & "-" & SumWeight105_109 & "-" & SumWeight100_104 & "-" & SumWeight95_99 & "-" & SumWeight90_94 & "-" & SumWeightLow
   
   'จากมากไปน้อย
'   For Each TempDiFind In CollFromDesc
'      Debug.Print TempDiFind.TOTAL_WEIGHT & "-" & TempDiFind.REAL_INDEX
'   Next TempDiFind
   'จากมากไปน้อย
            
   'น้อยไปมาก
'   For Each TempDiFind In CollFromAsc
'      Debug.Print TempDiFind.TOTAL_WEIGHT
'   Next TempDiFind
   'น้อยไปมาก
   
   'ส่วนใหญ่มากบางตัวน้อย
   
   For Round = 1 To 2
      If SumWeightLow <= 0 Then
         Exit For
      End If
      If Round = 1 Then
         MaxFitScale = 120
      ElseIf Round = 2 Then
         MaxFitScale = 90
      End If
      For Each TempDiFind In CollFromDesc
         If TempDiFind.TOTAL_WEIGHT > (MaxFitScale) Then
            SumAdjust = 0
            SumPool = SumPool + TempDiFind.TOTAL_WEIGHT - (MaxFitScale)
            
            For Each TempDiAdjust In CollFromAsc
               If SumPool <= 0 Then
                  Exit For
               End If
               If TempDiAdjust.TOTAL_WEIGHT < 90 Then
                  TempDiAdjust.FLAG_EDIT = "Y"
                  TempDiAdjust.TOTAL_ADJUST = 90 - TempDiAdjust.TOTAL_WEIGHT
                  If TempDiAdjust.TOTAL_ADJUST > SumPool Then
                     TempDiAdjust.TOTAL_ADJUST = SumPool
                     TempDiAdjust.TOTAL_WEIGHT = TempDiAdjust.TOTAL_WEIGHT + SumPool
                     SumPool = 0
                  Else
                     TempDiAdjust.TOTAL_WEIGHT = TempDiAdjust.TOTAL_WEIGHT + TempDiAdjust.TOTAL_ADJUST
                     SumPool = SumPool - TempDiAdjust.TOTAL_ADJUST
                  End If
                  SumWeightLow = SumWeightLow - TempDiAdjust.TOTAL_ADJUST
                  SumAdjust = SumAdjust + TempDiAdjust.TOTAL_ADJUST
               End If
            Next TempDiAdjust
            
            If SumAdjust > 0 Then
               TempDiFind.FLAG_EDIT = "Y"
               TempDiFind.TOTAL_ADJUST = SumAdjust
               TempDiFind.TOTAL_WEIGHT = TempDiFind.TOTAL_WEIGHT - SumAdjust
               SumWeightHigh = SumWeightHigh - SumAdjust
            End If
         End If
      Next TempDiFind
   Next Round
   
   Dim CanAdjustFlag  As Boolean
   If SumWeightHigh > 0 Then  'ตัวมากเหลือ
      CanAdjustFlag = True
      While CanAdjustFlag
         CanAdjustFlag = False
         For Each TempDiAdjust In CollFromAsc
            If SumWeightHigh <= 0 Then
               Exit For
            End If
            If TempDiAdjust.TOTAL_WEIGHT < 120 Then
               TempDiAdjust.FLAG_EDIT = "Y"
               TempDiAdjust.TOTAL_WEIGHT = TempDiAdjust.TOTAL_WEIGHT + 1
               SumWeightHigh = SumWeightHigh - 1
               CanAdjustFlag = True
               For Each TempDiFind In CollFromDesc
                  If TempDiFind.TOTAL_WEIGHT > 120 Then
                     TempDiFind.FLAG_EDIT = "Y"
                     TempDiFind.TOTAL_WEIGHT = TempDiFind.TOTAL_WEIGHT - 1
                     Exit For
                  End If
               Next TempDiFind
            End If
         Next TempDiAdjust
      Wend
   End If
   
   If SumWeightHigh <= 0 And SumWeightLow <= 0 Then
      'ปันน้ำหนัก
      For Each TempDiAdjust In CollFromAsc
         '90 บวก 1 2 3
         If TempDiAdjust.TOTAL_WEIGHT = 90 Then
            CanAdjustFlag = False
            RunningSwing = TempDiAdjust.REAL_INDEX Mod 3
            For Each TempDiFind In CollFromDesc
               'เกิน 100 ลดลง 1 2 3
               If TempDiFind.TOTAL_WEIGHT >= 100 And TempDiFind.TOTAL_WEIGHT <= 120 Then
                  TempDiFind.TOTAL_WEIGHT = TempDiFind.TOTAL_WEIGHT - RunningSwing
                  TempDiFind.FLAG_EDIT = "Y"
                  CanAdjustFlag = True
                  Exit For
               End If
            Next TempDiFind
            If CanAdjustFlag Then
               TempDiAdjust.TOTAL_WEIGHT = TempDiAdjust.TOTAL_WEIGHT + RunningSwing
               Debug.Print TempDiAdjust.FLAG_EDIT
               TempDiAdjust.FLAG_EDIT = "Y"
            End If
         End If
      Next TempDiAdjust
      'ปันน้ำหนัก
   End If
   
   If SumWeightHigh > 0 Or SumWeightLow > 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่สามารถปัดน้ำหนักได้ครบทุกน้ำหนัก"
      glbErrorLog.ShowUserError
   End If
   Call AdjustMainWeight(CollFromAsc, CollFromDesc)
End Sub
Public Sub AdjustMainWeight(CollFromAsc As Collection, CollFromDesc As Collection)
Dim TempDiAdjust As CAdjustPigWeight
Dim Di  As CDoItem
Dim SumWeightAfterAdjust As Double

   For Each TempDiAdjust In CollFromAsc
      If TempDiAdjust.FLAG_EDIT = "Y" Then
         Set Di = m_TempBD.DoItems(TempDiAdjust.REAL_INDEX)
         Di.TOTAL_WEIGHT = TempDiAdjust.TOTAL_WEIGHT
      End If
   Next TempDiAdjust
   
   For Each TempDiAdjust In CollFromDesc
      If TempDiAdjust.FLAG_EDIT = "Y" Then
         Set Di = m_TempBD.DoItems(TempDiAdjust.REAL_INDEX)
         Di.TOTAL_WEIGHT = TempDiAdjust.TOTAL_WEIGHT
      End If
   Next TempDiAdjust
   
   For Each Di In m_TempBD.DoItems
      Debug.Print Di.PIG_STATUS_NO & "-" & Di.PIG_STATUS_NAME & "-"; Di.PIG_STATUS
      
      'SumWeightAfterAdjust = SumWeightAfterAdjust + Di.TOTAL_WEIGHT
   Next Di

'   Debug.Print SumWeightAfterAdjust
End Sub
