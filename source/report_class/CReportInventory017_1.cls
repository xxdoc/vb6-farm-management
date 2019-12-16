VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory017_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory017_1"
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

Private Const HEADER_FORMAT1 = "^1|^3|^3|^6|^6|^6|^6|^4;"
Private Const HEADER_FORMAT2 = "^1|^3|^3|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2;"
Private Const DETAIL_FORMAT1 = "<7|<6|<22;"
Private Const DETAIL_FORMAT2 = "^1|^3|^3|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
Private Const FOOTER_FORMAT1 = "^1|^6|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
Private Const FOOTER_FORMAT2 = "^1|^3|^3|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_ProductStatus As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_ProductStatus = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_ProductStatus = Nothing
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
   Call m_Headers1.AddField(Cf)
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call m_Headers2.AddField(Cf)

   Call Cf.SetFieldValue(7, "^", "", "^")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(6, "^", "รับเข้า (โอน+ซื้อ+ปรับเพิ่ม)", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
'   Call m_Headers2.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
'   Call m_Headers2.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
'   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
      
   Call Cf.SetFieldValue(6, "^", "รับเข้า", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "จ่ายออก (โอน+ปรับลด)", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "เบิกใช้", "^")
   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(6, "^", "ขาย", "^")
'   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(12, "^", "จ่ายออก", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "คงเหลือ", "^")
   Call m_Headers1.AddField(Cf)
                                                
   Call Cf.SetFieldValue(6, "^", "คงเหลือ", "^")
   Call m_Headers0.AddField(Cf)
                                                
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
                                                
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
      
  Call m_Headers2.GetString(1, TempStr1, TempStr1_1)
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
   
   Set HeadCf = m_Headers2.Fields(1)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt1, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(2)
   Call BodyCf.SetFieldValue(HeadCf.Width * 2, HeadCf.Align, Txt2, "<")
   Call Details.AddField(BodyCf)
   
'   Set HeadCf = m_Headers2.Fields(3)
'   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt3, HeadCf.BodyAlign)
'   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers2.Fields(4)
   For j = 1 To 12
      If (j <> 2) And (j <> 5) And (j <> 8) And (j <> 11) And (j <> 14) Then
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Tot(j), , 3), HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      Else
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
   Next j
      
   Set BodyCf = Nothing
End Sub

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim O As Object
Dim InventoryBals As Collection
Dim RName As String
Dim cData As CPartLocation
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim EI As CExportItem
Dim II As CImportItem
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim TxCode As String
Dim Partitems As Collection
Dim PartTypes As Collection
Dim Houses As Collection
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim HS As CLocation
Dim Pi As CPartItem
Dim Ib As CImportItem
Dim Total1(30) As Double
Dim Total2(30) As Double
Dim Total3(30) As Double
Dim Total4(30) As Double
Dim PiCount As Long
Dim iCount As Long
Dim OldColor As Double
Dim Pt As CPartType
Dim LeftAmount As Double
Dim CurrentPrice As Double
Dim HasBegin As Boolean
Dim ImportItems As Collection
Dim ExportItems As Collection
Dim Ba As CBalanceAccum
Dim PrevPartType As CPartType

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   HasBegin = False
   
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   
   Set Partitems = New Collection
   
   Set ImportItems = New Collection
   Set ExportItems = New Collection
   
   Set PartTypes = New Collection
   Call LoadPartType(Nothing, PartTypes, -1, mcolParam("PART_TYPE"))
   
   Set InventoryBals = New Collection
   Call LoadInventoryBalanceEx(Nothing, InventoryBals, InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))

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
   
   Set BodyCf = New CReportField
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   If PartTypes.Count > 0 Then
      Set PrevPartType = PartTypes(1)
   End If
   
      For Each Pt In PartTypes
         For j = 1 To UBound(Total2)
            Total2(j) = 0
         Next j
         
'         If PrevPartType.PART_GROUP_NO <> Pt.PART_GROUP_NO Then
'            m_Details1.ClearField
'            Call GenerateFooter("รวม", "กลุ่ม " & PrevPartType.PART_GROUP_NAME & " (" & PrevPartType.PART_GROUP_NO & ") ", "", Total4, m_Details1)
'            Call m_Details1.GetString(2, TempStr1, TempStr2)
'            Vsp.FontBold = True
'            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'            Call Vsp.AddTable(strFormat, "", TempStr2)
'            Vsp.FontBold = False
'
'            Set PrevPartType = Pt
'            For j = 1 To UBound(Total4)
'               Total4(j) = 0
'            Next j
'         End If
         
         Call LoadImportPartItem(Nothing, Partitems, Pt.PART_TYPE_ID, mcolParam("LOCATION_ID"), mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
         PiCount = 0
         For Each Pi In Partitems
'            Call glbDaily.StartTransaction
            HasBegin = True

            PiCount = PiCount + 1
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            'Head
            Call m_Details1.ClearField
      
            Call BodyCf.SetFieldValue(7, "<", Pt.PART_TYPE_NAME & " (" & Pt.PART_TYPE_NO & ")", "^")
            Call m_Details1.AddField(BodyCf)
            Call BodyCf.SetFieldValue(12, "<", Pi.PART_DESC & " (" & Pi.PART_NO & ") หน่วยนับเป็น " & Pi.UNIT_NAME, "^")
            Call m_Details1.AddField(BodyCf)
            Call BodyCf.SetFieldValue(12, "<", "", "^")
            Call m_Details1.AddField(BodyCf)

            Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Call Vsp.AddTable(strFormat, "", TempStr2)
            'Head

            '=== ยกมา
'            Set O = GetBalanceItem(InventoryBals, Pi.PART_ITEM_ID, mcolParam("LOCATION_ID"), InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))))
Set Ba = GetBalanceAccum(InventoryBals, mcolParam("LOCATION_ID") & "-" & Pi.PART_ITEM_ID)
Set II = New CImportItem
Call BalanceAccum2ImportItem(Ba, II)
            Call GenerateBalanceDetail(0, II, m_Details1, Total1, Total2, Total3, Total4, CurrentPrice)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Call Vsp.AddTable(strFormat, "", TempStr2)
LeftAmount = II.CURRENT_AMOUNT
Set II = Nothing
            '=== ยกมา

            '=== Detail

            Set EI = New CExportItem
            EI.EXPORT_ITEM_ID = -1
            EI.FROM_DATE = mcolParam("FROM_DATE")
            EI.TO_DATE = mcolParam("TO_DATE")
            EI.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
            EI.PIG_FLAG = "N"
            EI.PART_ITEM_ID = Pi.PART_ITEM_ID
            EI.LOCATION_ID = mcolParam("LOCATION_ID")
            EI.OrderBy = 1
            EI.OrderType = mcolParam("ORDER_TYPE")
            Call EI.QueryData(13, Rs1, iCount)

            Set II = New CImportItem
            II.IMPORT_ITEM_ID = -1
            II.FROM_DATE = mcolParam("FROM_DATE")
            II.TO_DATE = mcolParam("TO_DATE")
            II.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
            II.PIG_FLAG = "N"
            II.PART_ITEM_ID = Pi.PART_ITEM_ID
            II.LOCATION_ID = mcolParam("LOCATION_ID")
            II.OrderBy = 1
            II.OrderType = mcolParam("ORDER_TYPE")
            Call II.QueryData(7, Rs2, iCount)
            '== Detail
            
            TxCode = "X"
            PiCount = 0
            While TxCode <> ""
               PiCount = PiCount + 1
               TxCode = GetNextTransaction(Rs1, Rs2, II, EI)
               If TxCode <> "" Then
                  I = I + 1
                  If TxCode = "I" Then
                     Set O = II
                  ElseIf TxCode = "E" Then
                     Set O = EI
                  End If
                                       
                  m_Details1.ClearField
                  Call GenerateDetail(I, O, m_Details1, Total1, Total2, Total3, Total4, LeftAmount, CurrentPrice)
                  Call m_Details1.GetString(2, TempStr1, TempStr2)
                  Vsp.FontBold = False
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If 'Tx code
               DoEvents
            Wend
            
         Total2(13) = Total2(13) + Total1(13)
         Total2(14) = Total2(14) + Total1(14)
         Total2(15) = Total2(15) + Total1(15)
                        
            m_Details1.ClearField
            Call GenerateFooter("รวม", "*** " & Pi.PART_DESC & " (" & Pi.PART_NO & ") ***", "", Total1, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            Vsp.FontBold = True
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

'            Call glbDaily.CommitTransaction
            HasBegin = False
         Next Pi
         
         '=== Summary for each type
         If PiCount > 0 Then
            Total3(13) = Total3(13) + Total2(13)
            Total3(14) = Total3(14) + Total2(14)
            Total3(15) = Total3(15) + Total2(15)
            
            m_Details1.ClearField
            Call GenerateFooter("รวม", "*** " & Pt.PART_TYPE_NAME & " (" & Pt.PART_TYPE_NO & ") ***", "", Total2, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            Vsp.FontBold = True
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
         '=== Summary for each type
      Next Pt
   
'      '== Part group
'      m_Details1.ClearField
'      Call GenerateFooter("รวม", "กลุ่ม " & PrevPartType.PART_GROUP_NAME & " (" & PrevPartType.PART_GROUP_NO & ") ", "", Total4, m_Details1)
'      Call m_Details1.GetString(2, TempStr1, TempStr2)
'      Vsp.FontBold = True
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
'      '=== Part Group
      
      m_Details1.ClearField
      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      Vsp.FontBold = True
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set BodyCf = Nothing
   
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   If Rs2.State = adStateOpen Then
      Rs2.Close
   End If
   Set Rs2 = Nothing
   
   Set EI = Nothing
   Set II = Nothing
   Set InventoryBals = Nothing
   Set PartTypes = Nothing
   Set Partitems = Nothing
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   If HasBegin Then
      glbDaily.RollbackTransaction
   End If
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


