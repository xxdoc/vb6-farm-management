VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory017"
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
'NO.|วันที่|เลขที่เอกสาร|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|มูลค่า;"
'NO.|วันที่|เลขที่เอกสาร|ยกมา|รับเข้า (โอน+ซื้อ+ปรับเพิ่ม)|จ่ายออก (โอน+ปรับลด)|เบิกใช้|ขาย|คงเหลือ;"

Private Sub Class_Initialize()
   Set mcolParam = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
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
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   strFormat = VSP_CalTable(HEADER_FORMAT1, mdblWid, alngX)
   strPrint = "NO.|วันที่|เลขที่เอกสาร|รับเข้า (โอน+ซื้อ+ปรับเพิ่ม)|จ่ายออก (โอน+ปรับลด)|เบิกใช้|ขาย|คงเหลือ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable(HEADER_FORMAT2, mdblWid, alngX)
   strPrint = "NO.|วันที่|เลขที่เอกสาร|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|มูลค่า;"
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
               If EI.TRANSACTION_SEQ < II.TRANSACTION_SEQ Then
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
         TempStr = TempStr & "" & FormatNumber(II.IMPORT_AMOUNT, "-") & "|" & Format(II.TOTAL_INCLUDE_PRICE / II.IMPORT_AMOUNT, "0.00") & "|" & FormatNumber(II.TOTAL_INCLUDE_PRICE, "-") & "|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(II.CURRENT_AMOUNT, "-") & "|" & FormatNumber(II.CURRENT_AMOUNT * II.NEW_PRICE, "-") & ";"
         
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
            TempStr = TempStr & "0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(EI.EXPORT_AMOUNT, "-") & "|" & FormatNumber(EI.EXPORT_AVG_PRICE, "-") & "|" & FormatNumber(EI.EXPORT_TOTAL_PRICE, "-") & "|" & FormatNumber(EI.CURRENT_AMOUNT, "-") & "|" & FormatNumber(EI.CURRENT_AMOUNT * EI.EXPORT_AVG_PRICE, "-") & ";"
            Totals1(13) = Totals1(13) + EI.EXPORT_AMOUNT
            Totals2(13) = Totals2(13) + EI.EXPORT_AMOUNT
            
            Totals1(14) = Totals1(14) + 0
            Totals2(14) = Totals2(14) + 0
         
            Totals1(15) = Totals1(15) + (EI.EXPORT_TOTAL_PRICE)
            Totals2(15) = Totals2(15) + (EI.EXPORT_TOTAL_PRICE)
         Else
            TempStr = TempStr & "0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(EI.EXPORT_AMOUNT, "-") & "|" & FormatNumber(EI.EXPORT_AVG_PRICE, "-") & "|" & FormatNumber(EI.EXPORT_TOTAL_PRICE, "-") & "|0.00|0.00|0.00|" & FormatNumber(EI.CURRENT_AMOUNT, "-") & "|" & FormatNumber(EI.CURRENT_AMOUNT * EI.EXPORT_AVG_PRICE, "-") & ";"
            Totals1(10) = Totals1(10) + EI.EXPORT_AMOUNT
            Totals2(10) = Totals2(10) + EI.EXPORT_AMOUNT
            
            Totals1(11) = Totals1(11) + 0
            Totals2(11) = Totals2(11) + 0
         
            Totals1(12) = Totals1(12) + (EI.EXPORT_TOTAL_PRICE)
            Totals2(12) = Totals2(12) + (EI.EXPORT_TOTAL_PRICE)
         End If
      ElseIf (EI.DOCUMENT_TYPE = 3) Or (EI.DOCUMENT_TYPE = 4) Then
         TempStr = TempStr & "0.00|0.00|0.00|" & FormatNumber(EI.EXPORT_AMOUNT, "-") & "|" & FormatNumber(EI.EXPORT_AVG_PRICE, "-") & "|" & FormatNumber(EI.EXPORT_TOTAL_PRICE, "-") & "|0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(EI.CURRENT_AMOUNT, "-") & "|" & FormatNumber(EI.CURRENT_AMOUNT * EI.EXPORT_AVG_PRICE, "-") & ";"
      
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

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim InventoryBals As Collection
Dim O As Object
Dim PrevO As Object
Dim RName As String
Dim cData As CPartLocation
Dim iCount1 As Long
Dim iCount2 As Long
Dim I As Long
Dim j As Long
Dim strFormat1 As String
Dim strFormat2 As String
Dim strFormat3 As String
Dim strFormat4 As String
Dim strPrint As String
Dim strPrint2 As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim Sum2 As Double
Dim PrevPI As CPartItem
Dim Value1(30) As Double
Dim Total1(30) As Double
Dim Total2(30) As Double
Dim PrevID As Long
Dim Amt As Double
Dim Temp1 As Double
Dim Temp2  As Double
Dim EI As CExportItem
Dim II As CImportItem
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim TxCode As String
Dim PrevCode As String
Dim Ib As CImportItem

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   
   Set InventoryBals = New Collection
   Call LoadInventoryBalance(Nothing, InventoryBals, InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
   
   Set EI = New CExportItem
   EI.EXPORT_ITEM_ID = -1
   EI.LOCATION_ID = mcolParam("LOCATION_ID")
   EI.FROM_DATE = mcolParam("FROM_DATE")
   EI.TO_DATE = mcolParam("TO_DATE")
   EI.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
   EI.PIG_FLAG = "N"
   EI.OrderBy = mcolParam("ORDER_BY")
   EI.OrderType = mcolParam("ORDER_TYPE")
   Call EI.QueryData(13, Rs1, iCount1)
   
   Set II = New CImportItem
   II.IMPORT_ITEM_ID = -1
   II.LOCATION_ID = mcolParam("LOCATION_ID")
   II.FROM_DATE = mcolParam("FROM_DATE")
   II.TO_DATE = mcolParam("TO_DATE")
   II.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
   II.PIG_FLAG = "N"
   II.OrderBy = mcolParam("ORDER_BY")
   II.OrderType = mcolParam("ORDER_TYPE")
   Call II.QueryData(7, Rs2, iCount2)

   If (iCount1 < 1) And (iCount2 < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
'NO.|วันที่|เลขที่เอกสาร|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|ราคา|มูลค่า|จำนวน|มูลค่า;
'NO.|วันที่|เลขที่เอกสาร|ยกมา|รับเข้า (โอน+ซื้อ+ปรับเพิ่ม)|จ่ายออก (โอน+ปรับลด)|เบิกใช้|ขาย|คงเหลือ;"

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      strFormat1 = VSP_CalTable(DETAIL_FORMAT2, mdblWid, alngX)
      strFormat2 = VSP_CalTable(FOOTER_FORMAT1, mdblWid, alngX)
      strFormat3 = VSP_CalTable(DETAIL_FORMAT1, mdblWid, alngX)
      strFormat4 = VSP_CalTable(FOOTER_FORMAT1, mdblWid, alngX)
      Vsp.TableBorder = tbAll
      Sum1 = 0

      Sum1 = 0
      Sum2 = 0
      
      Dim PrevDesc As String
      Dim PrevPartNo As String
      
      PrevCode = "XX-XX"
      TxCode = "X"
      While (TxCode <> "")
         TxCode = GetNextTransaction(Rs1, Rs2, II, EI)
         If TxCode = "I" Then
            Set O = II
         ElseIf TxCode = "E" Then
            Set O = EI
         End If
         
         If (PrevCode) <> (O.PART_TYPE_NO & "-" & O.PART_NO) Then
            'Change type and part item
            If PrevPartNo <> "" Then
               strPrint = "รวม|" & PrevDesc & " (" & PrevPartNo & ")|" & "" & FormatNumber(Total1(4)) & "||" & FormatNumber(Total1(6)) & "|" & FormatNumber(Total1(7)) & "||" & FormatNumber(Total1(9)) & "|" & FormatNumber(Total1(10)) & "||" & FormatNumber(Total1(12)) & "|" & FormatNumber(Total1(13)) & "||" & FormatNumber(Total1(15)) & "|" & FormatNumber(Total1(16)) & "|" & FormatNumber(Total1(17)) & ";"
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat4, "", strPrint)
               Vsp.FontBold = False
            End If
            
            strPrint = O.PART_TYPE_NAME & " (" & O.PART_TYPE_NO & ")" & "|"
            strPrint = strPrint & O.PART_DESC & " (" & O.PART_NO & ")" & "|"
            strPrint = strPrint & O.UNIT_NAME & ";"
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat3, "", strPrint)
            Vsp.FontBold = False
            
            Set Ib = GetImportItem(InventoryBals, Trim(Str(O.PART_ITEM_ID)) & "-" & Trim(Str(O.LOCATION_ID)))
            strPrint = "|" & "***ยกมา***" & "||" & "0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|" & FormatNumber(Ib.CURRENT_AMOUNT, "-") & "|" & FormatNumber(Ib.CURRENT_AMOUNT * Ib.NEW_PRICE, "-") & ";"
            Vsp.FontBold = False
            Call Vsp.AddTable(strFormat1, "", strPrint)
                        
            Total1(1) = Total1(1) + Ib.CURRENT_AMOUNT
            Total2(1) = Total2(1) + Ib.CURRENT_AMOUNT
            Total1(2) = Total1(2) + 0
            Total2(2) = Total2(2) + 0
            Total1(3) = Total1(3) + (Ib.NEW_PRICE * Ib.CURRENT_AMOUNT)
            Total2(3) = Total2(3) + (Ib.NEW_PRICE * Ib.CURRENT_AMOUNT)
                        
            Total1(16) = Ib.CURRENT_AMOUNT
            Total1(17) = Ib.NEW_PRICE * Ib.CURRENT_AMOUNT
                        
            PrevCode = O.PART_TYPE_NO & "-" & O.PART_NO
            PrevDesc = O.PART_DESC
            PrevPartNo = O.PART_NO
         End If
         
         strPrint = ItemToTable(EI, II, TxCode, Total1, Total2)
         If TxCode <> "" Then
            Vsp.FontBold = False
            Call Vsp.AddTable(strFormat1, "", strPrint)
         End If
      Wend
      
      Vsp.FontBold = True
      If PrevPartNo <> "" Then
         strPrint = "รวม|" & PrevDesc & " (" & PrevPartNo & ")|" & "" & FormatNumber(Total1(4)) & "||" & FormatNumber(Total1(6)) & "|" & FormatNumber(Total1(7)) & "||" & FormatNumber(Total1(9)) & "|" & FormatNumber(Total1(10)) & "||" & FormatNumber(Total1(12)) & "|" & FormatNumber(Total1(13)) & "||" & FormatNumber(Total1(15)) & "|" & FormatNumber(Total1(16)) & "|" & FormatNumber(Total1(17)) & ";"
         Call Vsp.AddTable(strFormat4, "", strPrint)
      End If
       
      strPrint = "สุทธิ||" & "" & FormatNumber(Total2(4)) & "||" & FormatNumber(Total2(6)) & "|" & FormatNumber(Total2(7)) & "||" & FormatNumber(Total2(9)) & "|" & FormatNumber(Total2(10)) & "||" & FormatNumber(Total2(12)) & "|" & FormatNumber(Total2(13)) & "||" & FormatNumber(Total2(15)) & "||;"
      Call Vsp.AddTable(strFormat4, "", strPrint)
       
       Vsp.FontBold = False
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
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


