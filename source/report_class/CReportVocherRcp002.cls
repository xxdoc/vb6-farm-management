VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportVocherRcp002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportVocherRcp002"
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
Private m_TempBD As CCashDoc
Private m_Total1(100)  As Double
Private m_Total2(100)  As Double
Private m_EndItem As Boolean
Private m_PaidAmount As Double
Private m_SumDebit As Double
Private m_SumCredit As Double
Private m_CnAmount As Double

Public m_Cr As Collection
'Public OKClick As Boolean

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
Private Sub GenerateRcpHeader(Border As TableBorderSettings, alngX() As Long, HeaderOffset)
Dim OldFont As Double
Dim TempStr As String
Dim Pmt As CCashTran
Dim Cheque As CCheque
   
   Vsp.FontBold = True
   Vsp.FontSize = 14
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.5 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "12.4cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "3.9cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "2.9cm"
    
    Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
    Vsp.TableCell(tcText, 1, 3) = vbCrLf & m_TempBD.GetFieldValue("DOCUMENT_NO")
    
    Vsp.TableBorder = Border
    Vsp.EndTable
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "11cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "8.2cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(m_TempBD.GetFieldValue("DOCUMENT_DATE"))
    
    Vsp.TableBorder = Border
    Vsp.EndTable
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.9 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "3cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "16.2cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    If mcolParam("MODE") = 1 Then
      Vsp.TableCell(tcText, 1, 2) = m_TempBD.GetFieldValue("CUSTOMER_NAME")
    ElseIf mcolParam("MODE") = 2 Then
      Vsp.TableCell(tcText, 1, 2) = "HO"
   ElseIf mcolParam("MODE") = 3 Then
      Vsp.TableCell(tcText, 1, 2) = glbEnterPrise.SHORT_NAME
    End If
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "3cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "16.2cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = "�����ء�,��,�������,�����"
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.85 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "19.2cm"
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.85 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "4cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "15.2cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taLeftBottom
    Vsp.TableCell(tcText, 1, 2) = m_TempBD.GetFieldValue("ACCOUNT_NAME")
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.85 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "19.2cm"
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.85 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "4cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "10.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "1.2cm"
    
    Vsp.TableCell(tcAlign, 1, 2) = taRightBottom
    Vsp.TableCell(tcText, 1, 2) = FormatNumber(GenerateReceiptItem(m_TempBD) - GenerateReceiptItemFee(m_TempBD))
    Vsp.TableCell(tcAlign, 1, 4) = taLeftBottom
    Vsp.TableCell(tcText, 1, 4) = "-" & ThaiBaht(GenerateReceiptItem(m_TempBD) - GenerateReceiptItemFee(m_TempBD)) & "-"
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "15.7cm"
    
    'If Pmt = CASH_PMT Then
      Vsp.TableCell(tcAlign, 1, 1) = taRightBottom
      'Vsp.TableCell(tcText, 1, 1) = "X"
    'End If
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
    If mcolParam("MODE") = 1 Or mcolParam("MODE") = 3 Then
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 3
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567 + HeaderOffset
       Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "1.5cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "15.7cm"
       
       'If Pmt = BANKTRF_PMT Then
         Vsp.TableCell(tcAlign, 1, 1) = taRightBottom
         'Vsp.TableCell(tcText, 1, 1) = "X"
       'End If
       
       Vsp.TableBorder = Border
       Vsp.EndTable
    End If
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "6.7cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "7.5cm"
    
'    If Pmt = CHECK_PMT Then
'      Vsp.TableCell(tcAlign, 1, 1) = taRightBottom
'      Vsp.TableCell(tcText, 1, 1) = "X"
'
'      Set Cheque = m_TempBD
'
'      Vsp.TableCell(tcAlign, 1, 3) = taLeftBottom
'      Vsp.TableCell(tcText, 1, 3) = m_TempBD.GetFieldValue("ACCOUNT_NO")
'    End If
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.85 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "6.7cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "1.5cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "7.5cm"
    
'    If Pmt = CHECK_PMT Then
'      Vsp.TableCell(tcAlign, 1, 3) = taLeftBottom
'      Vsp.TableCell(tcText, 1, 3) = m_TempBD.GetFieldValue("BANK_NAME")
'      Vsp.TableCell(tcAlign, 1, 5) = taLeftBottom
'      Vsp.TableCell(tcText, 1, 5) = m_TempBD.GetFieldValue("BRANCH_NAME")
'    End If
    
    Vsp.TableBorder = Border
    Vsp.EndTable
    
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 2.1 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "19.2cm"
    Vsp.TableBorder = Border
    Vsp.EndTable
    
    Vsp.FontBold = False
End Sub

Private Sub GenerateRcpFooter(Bd As CCashDoc, Border As TableBorderSettings, alngX() As Long)
Dim TempStr As String
Dim IntPart As String
Dim DecimalPart As String

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcFontSize) = 14
   Vsp.TableCell(tcFontBold) = True
   
    Vsp.TableCell(tcRowHeight) = "1.45cm"
    Vsp.TableCell(tcColWidth, 1, 1) = "11.8cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "2.9cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "0.8cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "2.9cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "0.8cm"
    
   TempStr = FormatNumber(m_SumDebit)
   IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
   DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 1) = "                             -" & ThaiBaht(m_SumDebit) & "-"
   Vsp.TableCell(tcAlign, 1, 2) = taRightMiddle
   Vsp.TableCell(tcText, 1, 2) = IntPart
   Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 3) = "  " & DecimalPart
   
   TempStr = FormatNumber(m_SumCredit)
   IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
   DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
   Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
   Vsp.TableCell(tcText, 1, 4) = IntPart
   Vsp.TableCell(tcAlign, 1, 5) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 5) = "  " & DecimalPart
    
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
   Vsp.TableBorder = tbAll

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

   m_ImageWidth = 19.22 * 567
   m_ImageHeight = 25.57 * 567

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Sub GetRelateCnDn(Bd As CCashDoc, DoID As Long, CnAmount As Double, DnAmount As Double)
Dim RCnDn As CReceiptCnDn
Dim TempCn As Double
Dim TempDn As Double

   TempCn = 0
   TempDn = 0
   For Each RCnDn In Bd.CashTranItems
'      If RCnDn.DO_ID = DoID Then
'         TempCn = TempCn + RCnDn.CN_AMOUNT
'         TempDn = TempDn + RCnDn.DN_AMOUNT
'      End If
   Next RCnDn
   
   CnAmount = TempCn
   DnAmount = TempDn
End Sub

Private Sub GenerateRcpItem(Bd As CCashDoc, Border As TableBorderSettings, alngX() As Long)
Dim I As Long
Dim j As Long
Dim RowHeight As Double
Dim MaxRow As Long
Dim LeftRow As Long
Dim TempStr As String
Dim IntPart As String
Dim DecimalPart As String
Dim Ri As CCashTran
Dim CnAmount As Double
Dim DnAmount As Double
Dim Cheque As CCheque

   MaxRow = 12
   RowHeight = 0.725 * 12 * 567 / MaxRow

   '===============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 6
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcFontSize) = 14
   Vsp.TableCell(tcFontBold) = True

    Vsp.TableCell(tcRowHeight) = RowHeight
    Vsp.TableCell(tcColWidth, 1, 1) = "9.8cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "2.9cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "0.8cm"
    Vsp.TableCell(tcColWidth, 1, 5) = "2.9cm"
    Vsp.TableCell(tcColWidth, 1, 6) = "0.8cm"

   TempStr = FormatNumber(GenerateReceiptItem(m_TempBD) - GenerateReceiptItemFee(m_TempBD))
   m_SumDebit = GenerateReceiptItem(m_TempBD) - GenerateReceiptItemFee(m_TempBD)
   IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
   DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
   
   Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
   If mcolParam("MODE") = 1 Then
      Vsp.TableCell(tcText, 1, 1) = m_TempBD.GetFieldValue("BANK_NAME")
   ElseIf mcolParam("MODE") = 2 Then
      Vsp.TableCell(tcText, 1, 1) = "�Թ�оѴ HO"
   ElseIf mcolParam("MODE") = 3 Then
      Vsp.TableCell(tcText, 1, 1) = m_TempBD.GetFieldValue("BANK_NAME")
   End If
   
   Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
   Vsp.TableCell(tcText, 1, 3) = IntPart
   Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
   Vsp.TableCell(tcText, 1, 4) = "  " & DecimalPart
   Vsp.TableBorder = Border
    Vsp.EndTable
   '==========
         
   If mcolParam("MODE") = 1 Or mcolParam("MODE") = 2 Then
      '=============== ��Ҹ�������
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontSize) = 14
      Vsp.TableCell(tcFontBold) = True
   
       Vsp.TableCell(tcRowHeight) = RowHeight
       Vsp.TableCell(tcColWidth, 1, 1) = "9.8cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "2cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "2.9cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "0.8cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "2.9cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "0.8cm"
   
      TempStr = FormatNumber(GenerateReceiptItemFee(m_TempBD))
      m_SumDebit = m_SumDebit + GenerateReceiptItemFee(m_TempBD)
      IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
      DecimalPart = Mid(TempStr, InStr(1, TempStr, ".") + 1)
   
      Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 1) = "��Ҹ�������"
      Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
      Vsp.TableCell(tcText, 1, 3) = IntPart
      Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 4) = "  " & DecimalPart
      Vsp.TableBorder = Border
       Vsp.EndTable
      '==========
   End If
   
   I = 1
   m_SumCredit = 0
   If mcolParam("MODE") = 1 Or mcolParam("MODE") = 2 Then
      For Each Ri In Bd.CashTranItems
         If Ri.GetFieldValue("TX_TYPE") = "I" Then
            I = I + 1
            m_SumCredit = m_SumCredit + Ri.GetFieldValue("AMOUNT")
            '===============
            Vsp.StartTable
            Vsp.TableCell(tcCols) = 6
            Vsp.TableCell(tcRows) = 1
            Vsp.TableCell(tcFontSize) = 14
            Vsp.TableCell(tcFontBold) = True
            
             Vsp.TableCell(tcRowHeight) = RowHeight
             Vsp.TableCell(tcColWidth, 1, 1) = "9.8cm"
             Vsp.TableCell(tcColWidth, 1, 2) = "2cm"
             Vsp.TableCell(tcColWidth, 1, 3) = "2.9cm"
             Vsp.TableCell(tcColWidth, 1, 4) = "0.8cm"
             Vsp.TableCell(tcColWidth, 1, 5) = "2.9cm"
             Vsp.TableCell(tcColWidth, 1, 6) = "0.8cm"
            
             TempStr = FormatNumber(Ri.GetFieldValue("AMOUNT"))
             IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
             DecimalPart = "  " & Mid(TempStr, InStr(1, TempStr, ".") + 1)
            
            Set Cheque = Ri.Cheque
            
            Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
            Vsp.TableCell(tcText, 1, 1) = "                          " & Ri.GetFieldValue("PAYMENT_TYPE_NAME") & "   " & Cheque.GetFieldValue("CHEQUE_NO")
            Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
            Vsp.TableCell(tcText, 1, 2) = ""
            Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
            Vsp.TableCell(tcText, 1, 5) = IntPart
            Vsp.TableCell(tcAlign, 1, 6) = taLeftMiddle
            Vsp.TableCell(tcText, 1, 6) = DecimalPart
      
             Vsp.TableBorder = Border
             Vsp.EndTable
             '===============
          End If
      Next Ri
   ElseIf mcolParam("MODE") = 3 Then
      m_SumCredit = m_SumDebit
      '===============
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontSize) = 14
      Vsp.TableCell(tcFontBold) = True
      
       Vsp.TableCell(tcRowHeight) = RowHeight
       Vsp.TableCell(tcColWidth, 1, 1) = "9.8cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "2cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "2.9cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "0.8cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "2.9cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "0.8cm"
      
       TempStr = FormatNumber(m_SumDebit)
       IntPart = Mid(TempStr, 1, InStr(1, TempStr, ".") - 1)
       DecimalPart = "  " & Mid(TempStr, InStr(1, TempStr, ".") + 1)
      
      Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 1) = "                          �Թ�оѴ   " & glbEnterPrise.SHORT_NAME
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 2) = ""
      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
      Vsp.TableCell(tcText, 1, 5) = IntPart
      Vsp.TableCell(tcAlign, 1, 6) = taLeftMiddle
      Vsp.TableCell(tcText, 1, 6) = DecimalPart

       Vsp.TableBorder = Border
       Vsp.EndTable
   End If
   m_EndItem = True

   LeftRow = MaxRow - I

   For j = 1 To LeftRow - 1
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcFontSize) = 14

       Vsp.TableCell(tcRowHeight) = RowHeight
       Vsp.TableCell(tcColWidth, 1, 1) = "9.8cm"
       Vsp.TableCell(tcColWidth, 1, 2) = "2cm"
       Vsp.TableCell(tcColWidth, 1, 3) = "2.9cm"
       Vsp.TableCell(tcColWidth, 1, 4) = "0.8cm"
       Vsp.TableCell(tcColWidth, 1, 5) = "2.9cm"
       Vsp.TableCell(tcColWidth, 1, 6) = "0.8cm"

       Vsp.TableBorder = Border
       Vsp.EndTable
   Next j
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
Dim Bd As CCashDoc

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

   Set m_TempBD = New CCashDoc
   Call m_TempBD.SetFieldValue("CASH_DOC_ID", mcolParam("CASH_DOC_ID"))
   m_TempBD.QueryFlag = 1
   Call glbDaily.QueryCashDoc(m_TempBD, Rs, iCount, IsOK, glbErrorLog)
   If Not Rs.EOF Then
      Call m_TempBD.PopulateFromRS(1, Rs)
   End If
   
   Vsp.StartDoc

   Call GenerateRcpItem(m_TempBD, tbNone, alngX)
   
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
   
   If m_EndItem Then
      Call GenerateRcpFooter(m_TempBD, tbNone, alngX)
   End If
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
Private Function GenerateReceiptItem(Bd As CCashDoc) As Double
Dim Ri As CCashTran
Dim TempAmount As Double

   TempAmount = 0
   
   For Each Ri In Bd.CashTranItems
      If Ri.GetFieldValue("TX_TYPE") = "I" Then
         TempAmount = TempAmount + Ri.GetFieldValue("AMOUNT")
      End If
   Next Ri
   
   GenerateReceiptItem = TempAmount
End Function
Private Function GenerateReceiptItemFee(Bd As CCashDoc) As Double
Dim Ri As CCashTran
Dim TempAmount As Double

   TempAmount = 0
   
   For Each Ri In Bd.CashTranItems
      If Ri.GetFieldValue("TX_TYPE") = "I" Then
         TempAmount = TempAmount + Ri.GetFieldValue("FEE_AMOUNT")
      End If
   Next Ri
   
   GenerateReceiptItemFee = TempAmount
End Function



