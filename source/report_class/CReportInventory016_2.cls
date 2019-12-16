VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory016_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory016_2"
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
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_PartItems As Collection

Private m_BalanceDistincts As Collection

Dim g_CharSpacing%

' note: this API is declared incorrectly in the VB API Viewer.
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_PartItems = New Collection
   
   Set m_BalanceDistincts = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_PartItems = Nothing
   
   Set m_BalanceDistincts = Nothing
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
   Call Cf.SetFieldValue(6, "^", "วัตถุดิบ", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(18, "^", "จำนวน", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(18, "^", "มูลค่า", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "วัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับเข้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอนเข้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปรับเพิ่ม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เบิกใช้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอนออก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปรับลด", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับเข้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอนเข้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปรับเพิ่ม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เบิกใช้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอนออก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปรับลด", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
                                                                                                            
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
   
   ' note: SetTextCharacterExtra takes a spacing in pixels, which
    '       we get by converting from twips.
    'g_CharSpacing = -25 / Vsp.TwipsPerPixelX
    
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
   
   'Call LoadImportPartItemEx(Nothing, m_PartItems, mcolParam("PART_TYPE"), mcolParam("LOCATION_ID"))
   
   Call LoadInventoryBalanceDistinctPartItem(Nothing, m_BalanceDistincts, , mcolParam("TO_DATE"), , mcolParam("LOCATION_ID"), , , "N", mcolParam("PART_TYPE"))
   
   FromDate = "จาก " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & " - " & ToDate & " (" & EmptyToString(mcolParam("LOCATION_NAME"), "รวม") & ")"

   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetPartItem(TempCol As Collection, TempID As Long) As CPartItem
On Error Resume Next
Dim EI As CPartItem

   Set EI = TempCol(Trim(Str(TempID)))
   If EI Is Nothing Then
      Set GetPartItem = Nothing
   Else
      Set GetPartItem = EI
   End If
End Function

Private Function GetImportItem(TempCol As Collection, TempKey As String) As CImportItem
On Error Resume Next
Dim EI As CImportItem
Static TempEi As CImportItem

   Set EI = TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CImportItem
      End If
      Set GetImportItem = TempEi
   Else
      Set GetImportItem = EI
   End If
End Function

Private Function CompareKey(Key1 As String, Key2 As String, CompareType As Long) As Boolean
      If CompareType = 1 Then
         CompareKey = Key1 < Key2
      Else
         CompareKey = Key2 < Key1
      End If
End Function

Private Function GetNextTransaction(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, II As CImportItem, EI As CExportItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
Static I As Long

   I = I + 1
''debug.print "==== " & I
If I = 26 Then
''debug.print
End If

   'Export
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call EI.PopulateFromRS(13, Rs1)
   End If
   
   'Import
   EofFlag2 = Rs2.EOF
   If Not Rs2.EOF Then
      Call II.PopulateFromRS(7, Rs2)
   End If
   
   If (EofFlag1 And EofFlag2) Then
      GetNextTransaction = ""
   ElseIf (EofFlag1 And (Not EofFlag2)) Then
      GetNextTransaction = "I"
      Rs2.MoveNext
   ElseIf ((Not EofFlag1) And EofFlag2) Then
      GetNextTransaction = "E"
      Rs1.MoveNext
   Else
      '===
      'การเรียงลำดับมีผลอย่างมาก
      If EI.PART_TYPE_NO = II.PART_TYPE_NO Then
         If EI.PART_NO = II.PART_NO Then
            If DateToStringInt(EI.DOCUMENT_DATE) = DateToStringInt(II.DOCUMENT_DATE) Then
               If EI.PRIORITY1 = II.PRIORITY1 Then
                  If EI.DOCUMENT_NO < II.DOCUMENT_NO Then
                     GetNextTransaction = "E"
                  Else
                     GetNextTransaction = "I"
                  End If
               ElseIf EI.PRIORITY1 < II.PRIORITY1 Then
                  GetNextTransaction = "E"
               Else
                  GetNextTransaction = "I"
               End If
            ElseIf DateToStringInt(EI.DOCUMENT_DATE) < DateToStringInt(II.DOCUMENT_DATE) Then
               GetNextTransaction = "E"
            Else
               GetNextTransaction = "I"
            End If 'Document date
         ElseIf EI.PART_NO < II.PART_NO Then
            GetNextTransaction = "E"
         Else
            GetNextTransaction = "I"
         End If 'Part no
      ElseIf CompareKey(EI.PART_TYPE_NO, II.PART_TYPE_NO, mcolParam("ORDER_TYPE")) Then
         GetNextTransaction = "E"
      Else
         GetNextTransaction = "I"
      End If 'Part type no
      '===
      If GetNextTransaction = "I" Then
'''debug.print "I" & ";" & Ii.PART_TYPE_NO & ";" & Ii.PART_NO & ";" & DateToStringInt(Ii.DOCUMENT_DATE) & ";" & Ii.TRANSACTION_SEQ & ";" & Ii.DOCUMENT_NO
         Rs2.MoveNext
      ElseIf GetNextTransaction = "E" Then
'''debug.print "E" & ";" & Ei.PART_TYPE_NO & ";" & Ei.PART_NO & ";" & DateToStringInt(Ei.DOCUMENT_DATE) & ";" & Ei.TRANSACTION_SEQ & ";" & Ei.DOCUMENT_NO
         Rs1.MoveNext
      End If
   End If 'Eof flag
End Function

Private Function ItemToTable(EI As CExportItem, II As CImportItem, TxCode As String, Totals1() As Double, Totals2() As Double) As String
Dim TempStr As String
Static I As Long
'NO.|วันที่|เลขที่เอกสาร|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า| จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|มูลค่า;
'NO.|วันที่|เลขที่เอกสาร|ยกมา|รับเข้า (โอน+ซื้อ+ปรับเพิ่ม)|จ่ายออก (โอน+ปรับลด)|เบิกใช้|ขาย|คงเหลือ;"

   TempStr = ""
   I = I + 1

   If TxCode = "I" Then
      TempStr = I & "|"
      TempStr = TempStr & InternalDateToStringEx4(DateToStringInt(II.DOCUMENT_DATE)) & "|"
      TempStr = TempStr & II.DOCUMENT_NO & "|"
      If (II.DOCUMENT_TYPE = 1) Or (II.DOCUMENT_TYPE = 3) Or (II.DOCUMENT_TYPE = 4) Then
         TempStr = TempStr & "" & FormatNumber(II.IMPORT_AMOUNT, , 3) & "|" & FormatNumber(II.TOTAL_INCLUDE_PRICE / II.IMPORT_AMOUNT, , 3) & "|" & FormatNumber(II.TOTAL_INCLUDE_PRICE) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(II.CURRENT_AMOUNT) & "|" & FormatNumber(II.CURRENT_AMOUNT * II.NEW_PRICE, , 3) & ";"
         
         Totals1(4) = Totals1(4) + II.IMPORT_AMOUNT
         Totals2(4) = Totals2(4) + II.IMPORT_AMOUNT
         
         Totals1(5) = Totals1(5) + 0
         Totals2(5) = Totals2(5) + 0
      
         Totals1(6) = Totals1(6) + (II.TOTAL_INCLUDE_PRICE)
         Totals2(6) = Totals2(6) + (II.TOTAL_INCLUDE_PRICE)
      End If
      
      Totals1(16) = II.CURRENT_AMOUNT
      Totals1(17) = II.CURRENT_AMOUNT * II.NEW_PRICE
   ElseIf TxCode = "E" Then
      TempStr = I & "|"
      TempStr = TempStr & InternalDateToStringEx4(DateToStringInt(EI.DOCUMENT_DATE)) & "|"
      TempStr = TempStr & EI.DOCUMENT_NO & "|"
      If EI.DOCUMENT_TYPE = 2 Then
         If EI.SALE_FLAG = "Y" Then
            TempStr = TempStr & "" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(EI.EXPORT_AMOUNT, , 3) & "|" & FormatNumber(EI.EXPORT_AVG_PRICE, , 3) & "|" & FormatNumber(EI.EXPORT_TOTAL_PRICE, , 3) & "|" & FormatNumber(EI.CURRENT_AMOUNT, , 3) & "|" & FormatNumber(EI.CURRENT_AMOUNT * EI.EXPORT_AVG_PRICE, , 3) & ";"
            Totals1(13) = Totals1(13) + EI.EXPORT_AMOUNT
            Totals2(13) = Totals2(13) + EI.EXPORT_AMOUNT
            
            Totals1(14) = Totals1(14) + 0
            Totals2(14) = Totals2(14) + 0
         
            Totals1(15) = Totals1(15) + (EI.EXPORT_TOTAL_PRICE)
            Totals2(15) = Totals2(15) + (EI.EXPORT_TOTAL_PRICE)
         Else
            TempStr = TempStr & "" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(EI.EXPORT_AMOUNT, , 3) & "|" & FormatNumber(EI.EXPORT_AVG_PRICE, , 3) & "|" & FormatNumber(EI.EXPORT_TOTAL_PRICE, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(EI.CURRENT_AMOUNT, , 3) & "|" & FormatNumber(EI.CURRENT_AMOUNT * EI.EXPORT_AVG_PRICE, , 3) & ";"
            Totals1(10) = Totals1(10) + EI.EXPORT_AMOUNT
            Totals2(10) = Totals2(10) + EI.EXPORT_AMOUNT
            
            Totals1(11) = Totals1(11) + 0
            Totals2(11) = Totals2(11) + 0
         
            Totals1(12) = Totals1(12) + (EI.EXPORT_TOTAL_PRICE)
            Totals2(12) = Totals2(12) + (EI.EXPORT_TOTAL_PRICE)
         End If
      ElseIf (EI.DOCUMENT_TYPE = 3) Or (EI.DOCUMENT_TYPE = 4) Then
         TempStr = TempStr & "" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(EI.EXPORT_AMOUNT, , 3) & "|" & FormatNumber(EI.EXPORT_AVG_PRICE, , 3) & "|" & FormatNumber(EI.EXPORT_TOTAL_PRICE, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(0, , 3) & "|" & FormatNumber(EI.CURRENT_AMOUNT, , 3) & "|" & FormatNumber(EI.CURRENT_AMOUNT * EI.EXPORT_AVG_PRICE, , 3) & ";"
      
         Totals1(7) = Totals1(7) + EI.EXPORT_AMOUNT
         Totals2(7) = Totals2(7) + EI.EXPORT_AMOUNT
         
         Totals1(8) = Totals1(8) + 0
         Totals2(8) = Totals2(8) + 0
      
         Totals1(9) = Totals1(9) + (EI.EXPORT_TOTAL_PRICE)
         Totals2(9) = Totals2(9) + (EI.EXPORT_TOTAL_PRICE)
      End If
      
      Totals1(16) = EI.CURRENT_AMOUNT
      Totals1(17) = EI.CURRENT_AMOUNT * EI.EXPORT_AVG_PRICE
   End If
   
   ItemToTable = TempStr
End Function

Private Sub GenerateDetail(I As Long, O As Object, Details As CFieldList, Tot() As Double, Tot2() As Double, Tot3() As Double, Tot4() As Double, LeftAmount As Double, CurrentPrice As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Amt1 As Double
Dim Amt2 As Double
Dim Amt3 As Double
Dim Buf(30) As Double
Dim j As Long

   Call Details.ClearField
   
   For j = 1 To UBound(Buf)
      Buf(j) = 0
   Next j
   
   Set BodyCf = New CReportField
   
   Set HeadCf = m_Headers2.Fields(1)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Trim(Str(I)), HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(2)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, DateToStringExtEx2(O.DOCUMENT_DATE), HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(3)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, O.DOCUMENT_NO, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   If O.TX_TYPE = "I" Then
      Amt1 = O.IMPORT_AMOUNT
      CurrentPrice = MyDiffEx(LeftAmount * CurrentPrice + O.TOTAL_INCLUDE_PRICE, LeftAmount + O.IMPORT_AMOUNT)
      LeftAmount = LeftAmount + O.IMPORT_AMOUNT
      If O.IMPORT_AMOUNT = 0 Then
         Amt2 = 0
      Else
         Amt2 = O.TOTAL_INCLUDE_PRICE / O.IMPORT_AMOUNT
      End If
      Amt3 = O.TOTAL_INCLUDE_PRICE
      
      Buf(1) = Amt1
      Buf(2) = Amt2
      Buf(3) = Amt3
   
      If O.COMMIT_FLAG = "Y" Then
         Buf(10) = LeftAmount
         Buf(11) = CurrentPrice
         Buf(12) = LeftAmount * CurrentPrice
      Else
         Buf(10) = LeftAmount
         Buf(11) = CurrentPrice
         Buf(12) = LeftAmount * CurrentPrice
      End If
   Else
      Amt1 = O.EXPORT_AMOUNT
      Amt2 = CurrentPrice 'O.EXPORT_AVG_PRICE
      Amt3 = O.EXPORT_AMOUNT * CurrentPrice 'O.EXPORT_TOTAL_PRICE
      LeftAmount = LeftAmount - O.EXPORT_AMOUNT

      If (O.DOCUMENT_TYPE = 3) Or (O.DOCUMENT_TYPE = 4) Then
         Buf(4) = Amt1
         Buf(5) = Amt2
         Buf(6) = Amt3
      ElseIf (O.DOCUMENT_TYPE = 2) And (O.SALE_FLAG = "N") Then
         Buf(7) = Amt1
         Buf(8) = Amt2
         Buf(9) = Amt3
      ElseIf (O.DOCUMENT_TYPE = 2) And (O.SALE_FLAG = "Y") Then
         Buf(4) = Amt1
         Buf(5) = Amt2
         Buf(6) = Amt3
      ElseIf (O.DOCUMENT_TYPE = 10) And (O.DOCUMENT_SUBTYPE = 2) Then
         Buf(4) = Amt1
         Buf(5) = Amt2
         Buf(6) = Amt3
      ElseIf (O.DOCUMENT_TYPE = 13) And (O.DOCUMENT_SUBTYPE = 2) Then
         Buf(4) = Amt1
         Buf(5) = Amt2
         Buf(6) = Amt3
      End If
      
      If O.COMMIT_FLAG = "Y" Then
         Buf(10) = LeftAmount
         Buf(11) = CurrentPrice
         Buf(12) = LeftAmount * CurrentPrice
      Else
         Buf(10) = LeftAmount
         Buf(11) = CurrentPrice
         Buf(12) = LeftAmount * CurrentPrice
      End If
   End If
   
   Set HeadCf = m_Headers2.Fields(4)
   
   For j = 1 To 12
      Tot(j) = Tot(j) + Buf(j)
      If j < 13 Then
         Tot2(j) = Tot2(j) + Buf(j)
         Tot3(j) = Tot3(j) + Buf(j)
         Tot4(j) = Tot4(j) + Buf(j)
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Buf(j), , 3), HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next j
   
   Tot(10) = Buf(10)
   Tot(11) = Buf(11)
   Tot(12) = Buf(12)
   
   Set BodyCf = Nothing
End Sub

Private Sub GenerateBalanceDetail(I As Long, O As Object, Details As CFieldList, Tot() As Double, Tot2() As Double, Tot3() As Double, Tot4() As Double, CurrentPrice As Double)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Amt1 As Double
Dim Amt2 As Double
Dim Amt3 As Double
Dim Buf(30) As Double
Dim j As Long

   Call Details.ClearField
   
   For j = 1 To UBound(Buf)
      Buf(j) = 0
   Next j
   
   Set BodyCf = New CReportField
   
   Set HeadCf = m_Headers2.Fields(1)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(2)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "***ยกมา***", HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(3)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   If O.TX_TYPE = "I" Then
      CurrentPrice = O.NEW_PRICE
      
      Amt1 = O.CURRENT_AMOUNT
      Amt2 = O.NEW_PRICE
      Amt3 = O.NEW_PRICE * O.CURRENT_AMOUNT
         
      Buf(10) = Amt1
      Buf(11) = Amt2
      Buf(12) = Amt3
   Else
      CurrentPrice = O.EXPORT_AVG_PRICE
      
      Amt1 = O.CURRENT_AMOUNT
      Amt2 = O.EXPORT_AVG_PRICE
      Amt3 = O.EXPORT_AVG_PRICE * O.CURRENT_AMOUNT
      
      Buf(10) = Amt1
      Buf(11) = Amt2
      Buf(12) = Amt3
   End If
   
   Set HeadCf = m_Headers2.Fields(4)
   
   For j = 1 To 12
      Tot(j) = Tot(j) + Buf(j)
      If j < 13 Then
         Tot2(j) = Tot2(j) + Buf(j)
         Tot3(j) = Tot3(j) + Buf(j)
         Tot4(j) = Tot4(j) + Buf(j)
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Buf(j), , 3), HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next j
   
   Tot(10) = Buf(10)
   Tot(11) = Buf(11)
   Tot(12) = Buf(12)
   
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim Amt1 As Double
Dim Amt2 As Double
Dim Amt3 As Double
Dim Buf(30) As Double
Dim j As Long

   Call Details.ClearField
   
   For j = 1 To UBound(Buf)
      Buf(j) = 0
   Next j
   
   Set BodyCf = New CReportField
   
   Set HeadCf = m_Headers1.Fields(1)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt1, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers1.Fields(2)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt2, "<")
   Call Details.AddField(BodyCf)
               
   For j = 1 To 18
      Set HeadCf = m_Headers1.Fields(j + 2)
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Tot(j + 2), "-", 3), HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next j
      
   Set BodyCf = Nothing
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim O As Object
Dim RName As String
Dim cData As CPartLocation
Dim I As Long
Dim j As Long
Dim k As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim Pi As CPartItem
Dim OldPi As CPartItem
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim InventoryBals As Collection
Dim BalanceAccums As Collection
Dim Ba As CBalanceAccum
Dim PatialImports As Collection
Dim PatialExports As Collection
Dim EI As CExportItem
Dim EI2 As CExportItem
Dim EI3 As CExportItem
Dim II As CImportItem
Dim SumAmt As Double
Dim SumPrice As Double
Dim TempSum As Double
Dim p As Long
Dim Str1 As String
Dim OldBaDistinct As CBalanceAccum
Dim BaDistinct As CBalanceAccum

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
      
   Set PatialImports = New Collection
   Set PatialExports = New Collection
   Set BalanceAccums = New Collection
   
   If Not initDoc Then Exit Function
   
   If mcolParam("FROM_DATE") > 0 Then
      Call LoadInventoryBalanceExByPart(Nothing, BalanceAccums, InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))), mcolParam("TO_DATE"), "", mcolParam("LOCATION_ID"), , , "N", mcolParam("PART_TYPE"))
   End If
   
   Set PatialImports = New Collection
   Call LoadImportPrice3(Nothing, PatialImports, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
   
   Set PatialExports = New Collection
   Call LoadExportPrice8(Nothing, PatialExports, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
   
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   Set BodyCf = New CReportField
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
   p = 0
'   Set OldPi = New CPartItem
'   If m_PartItems.Count > 0 Then
'      OldPi.PART_TYPE_NO = m_PartItems(1).PART_TYPE_NO
'      OldPi.PART_TYPE_NAME = m_PartItems(1).PART_TYPE_NAME
'   End If
   
   Set OldBaDistinct = New CBalanceAccum
   If m_BalanceDistincts.Count > 0 Then
      OldBaDistinct.PART_TYPE_NO = m_BalanceDistincts(1).PART_TYPE_NO
      OldBaDistinct.PART_TYPE_NAME = m_BalanceDistincts(1).PART_TYPE_NAME
   End If
   
   I = 0
   For Each BaDistinct In m_BalanceDistincts
      If OldBaDistinct.PART_TYPE_NO <> BaDistinct.PART_TYPE_NO Then
         p = p + 1
         If mcolParam("MODE") = 1 Then
            Str1 = "รวม"
         Else
            Str1 = p
         End If
         Call GenerateFooter(Str1, OldBaDistinct.PART_TYPE_NAME & " (" & OldBaDistinct.PART_TYPE_NO & ")", "", Total1, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
         OldBaDistinct.PART_TYPE_NO = BaDistinct.PART_TYPE_NO
         OldBaDistinct.PART_TYPE_NAME = BaDistinct.PART_TYPE_NAME
      End If
      
      I = I + 1
      m_Details1.ClearField
      k = 0
      
      k = 1
      Set HeadCf = m_Headers1.Fields(k)
      TempStr = I
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
   
      k = 2
      Set HeadCf = m_Headers1.Fields(k)
      TempStr = BaDistinct.PART_DESC & " (" & BaDistinct.PART_NO & ")"
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      If BaDistinct.PART_NO = "132-D" Then
         Debug.Print
      End If
      SumAmt = 0
      SumPrice = 0
      TempSum = 0
      For k = 1 To 18
         Set HeadCf = m_Headers1.Fields(k + 2)
         If (k = 1) Then
            Set Ba = GetBalanceAccum(BalanceAccums, Trim(Str(BaDistinct.PART_ITEM_ID)))
            Amt = Ba.BALANCE_AMOUNT
            '''debug.print (Amt)
            SumAmt = SumAmt + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf (k = 2) Then
            Set II = GetImportItem(PatialImports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(1)) & "-N")
            Amt = II.SUM_ALL_IMPORT
            SumAmt = SumAmt + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 3 Then
            Set II = GetImportItem(PatialImports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(3)) & "-N")
            Amt = II.SUM_ALL_IMPORT
            SumAmt = SumAmt + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 4 Then
            Set II = GetImportItem(PatialImports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(4)) & "-N")
            Amt = II.SUM_ALL_IMPORT
            SumAmt = SumAmt + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 5 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(2)) & "-N")
            Amt = EI.SUM_ALL_EXPORT
            SumAmt = SumAmt - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 6 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(3)) & "-N")
            Amt = EI.SUM_ALL_EXPORT
            SumAmt = SumAmt - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 7 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(4)) & "-N")
            Amt = EI.SUM_ALL_EXPORT
            SumAmt = SumAmt - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 8 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(2)) & "-Y")
            Set EI2 = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(10)) & "-Y")
            Set EI3 = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(13)) & "-Y")
            Amt = EI.SUM_ALL_EXPORT + EI2.SUM_ALL_EXPORT + EI3.SUM_ALL_EXPORT
            SumAmt = SumAmt - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf (k = 9) Then
            Amt = SumAmt
         ElseIf (k = 10) Then
            Set Ba = GetBalanceAccum(BalanceAccums, Trim(Str(BaDistinct.PART_ITEM_ID)))
            Amt = Ba.TOTAL_INCLUDE_PRICE              '''''''''''''''
            SumPrice = SumPrice + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf (k = 11) Then
            Set II = GetImportItem(PatialImports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(1)) & "-N")
            Amt = II.SUM_ALL_IMPORT_AMT
            SumPrice = SumPrice + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 12 Then
            Set II = GetImportItem(PatialImports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(3)) & "-N")
            Amt = II.SUM_ALL_IMPORT_AMT
            SumPrice = SumPrice + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 13 Then
            Set II = GetImportItem(PatialImports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(4)) & "-N")
            Amt = II.SUM_ALL_IMPORT_AMT                           '''''''''''
            SumPrice = SumPrice + Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 14 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(2)) & "-N")
            Amt = EI.SUM_ALL_EXPORT_AMT
            SumPrice = SumPrice - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 15 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(3)) & "-N")
            Amt = EI.SUM_ALL_EXPORT_AMT
            SumPrice = SumPrice - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 16 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(4)) & "-N")
            Amt = EI.SUM_ALL_EXPORT_AMT
            SumPrice = SumPrice - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf k = 17 Then
            Set EI = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(2)) & "-Y")
            Set EI2 = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(10)) & "-Y")
            Set EI3 = GetExportItem(PatialExports, Trim(Str(BaDistinct.PART_ITEM_ID)) & "-" & Trim(Str(13)) & "-Y")
            Amt = EI.SUM_ALL_EXPORT_AMT + EI2.SUM_ALL_EXPORT_AMT + EI3.SUM_ALL_EXPORT_AMT
            SumPrice = SumPrice - Amt
            TempSum = TempSum + Abs(Amt)
         ElseIf (k = 18) Then
            Amt = SumPrice
         Else
            Amt = 0
         End If
         
         TempStr = FormatNumber(Amt, "-", 3)
         Total1(k + 2) = Total1(k + 2) + Amt
         Total2(k + 2) = Total2(k + 2) + Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next k
      
      If (Round(TempSum, 2) <> 0) Then
      'If (1 = 1) Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = False
         If mcolParam("MODE") = 1 Then
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      Else
         I = I - 1
      End If
   Next BaDistinct
   
   p = p + 1
   If mcolParam("MODE") = 1 Then
      Str1 = "รวม"
   Else
      Str1 = p
   End If
   Call GenerateFooter(Str1, OldBaDistinct.PART_TYPE_NAME & " (" & OldBaDistinct.PART_TYPE_NO & ")", "", Total1, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set BalanceAccums = Nothing
   Set PatialImports = Nothing
   Set PatialExports = Nothing
   
   Set OldBaDistinct = Nothing
   Set BaDistinct = Nothing
   
   Set BodyCf = Nothing
      
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
   
   ' note: this needs to be called at the NewPage event because
    '       each page has its own hDC.
    SetTextCharacterExtra Vsp.hdc, g_CharSpacing
    
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




