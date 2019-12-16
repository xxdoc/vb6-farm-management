VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory006"
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
Private Const HEADER_FORMAT = "^1|^3|^5|^3|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2;"
Private Const DETAIL_FORMAT = "^1|<3|<5|^3|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
Private Const FOOTER_FORMAT = "^1|<3|<5|^3|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
'ลำดับ|ประเภทวัตถุดิบ|วัตถุดิบ|ราคา/หน่วย|หน่วย|ยอดยกมา|รับเข้า|โอนเข้า|โอนออก|เบิกใช้|ปรับยอด เพิ่ม|ปรับยอด ลด|คงเหลือ|มูลค่ารวม;

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
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
   strPrint = "NO.|ประเภทวัตถุดิบ|วัตถุดิบ|จำนวน/มูลค่า|ราคา/หน่วย|ยอดยกมา|รับเข้า|โอนเข้า|ปรับเพิ่ม|เบิกใช้|โอนออก|ปรับลด|ขาย|คงเหลือ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim FromDate As String
Dim ToDate As String

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4
   Vsp.ORIENTATION = orLandscape
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
                                       vbCrLf & FromDate & " - " & ToDate
                                       
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim cData As CPartLocation
Dim iCount As Long
Dim i As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim Sum2 As Double
Dim LeftAmount As Double
Dim ValueAmount As Double
Dim II As CImportItem
Dim TempRs As ADODB.Recordset
Dim TempCol As Collection
Dim AvgPrice As Double
Dim TotalLeft As Double
Dim TempToDate As String
Dim Locations As Collection
Dim Lc As CLocation
Dim j As Long
Dim Total1(15) As Double
Dim Total2(15) As Double
Dim strFormat2 As String
Dim strPrint2 As String
Dim PrevID As Long
Dim Total3(15) As Double
Dim Total4(15) As Double

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   Set TempCol = New Collection
   Set Locations = New Collection
   Call LoadLocation(Nothing, Locations, 2, "")
   
   Set II = New CImportItem
   II.LOCATION_ID = -1
   II.FROM_DATE = mcolParam("FROM_DATE")
   II.TO_DATE = mcolParam("TO_DATE")
   II.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
   II.OrderBy = 4
   II.OrderType = 2
   Call II.QueryData(1, TempRs, iCount)
   Set II = Nothing
   While Not TempRs.EOF
      Set II = New CImportItem
      Call II.PopulateFromRS(1, TempRs)
      Call TempCol.Add(II)
      Set II = Nothing
      TempRs.MoveNext
   Wend
   
   Set cData = New CPartLocation
   cData.PART_LOCATION_ID = -1
   cData.FROM_DATE = mcolParam("FROM_DATE")
   cData.TO_DATE = mcolParam("TO_DATE")
   cData.LOCATION_ID = -1
   cData.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
   cData.OrderBy = mcolParam("ORDER_BY")
   cData.OrderType = mcolParam("ORDER_TYPE")
   Call cData.QueryData(3, Rs, iCount)

   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   Set cData = Nothing
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
'   "ลำดับ|ประเภทวัตถุดิบ|วัตถุดิบ|ราคา/หน่วย|หน่วย|ยอดยกมา|รับเข้า|โอนเข้า|โอนออก|เบิกใช้|ปรับยอด เพิ่ม|ปรับยอด ลด|คงเหลือ|มูลค่ารวม;"
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)
      strFormat2 = VSP_CalTable(FOOTER_FORMAT, mdblWid, alngX)
      Vsp.TableBorder = tbAll
      Sum1 = 0
      
      If mcolParam("TO_DATE") <= 0 Then
         TempToDate = "9999-99-99 99:99:99"
      Else
         TempToDate = DateToStringIntHi(mcolParam("TO_DATE"))
      End If
      
      For j = 1 To 10
         Total1(j) = 0
         Total1(2) = 0
         Total1(3) = 0
         Total1(4) = 0
      Next j
      
      For i = 1 To iCount
         If i = 1 Then
            PrevID = NVLI(Rs("PART_TYPE_ID"), 0)
         Else
            '====
            If PrevID <> NVLI(Rs("PART_TYPE_ID"), 0) Then
               strPrint2 = "รวม|||จำนวน||" & FormatNumber(Total2(1)) & "|" & FormatNumber(Total2(2)) & "|" & FormatNumber(Total2(3)) & "|" & FormatNumber(Total2(4)) & "|" & FormatNumber(Total2(5)) & "|" & FormatNumber(Total2(6)) & "|" & FormatNumber(Total2(7)) & "|" & FormatNumber(Total2(8)) & "|" & FormatNumber(Total2(9)) & ";"
               Call Vsp.AddTable(strFormat2, "", strPrint2)
               strPrint2 = "รวม|||มูลค่า||" & FormatNumber(Total4(1)) & "|" & FormatNumber(Total4(2)) & "|" & FormatNumber(Total4(3)) & "|" & FormatNumber(Total4(4)) & "|" & FormatNumber(Total4(5)) & "|" & FormatNumber(Total4(6)) & "|" & FormatNumber(Total4(7)) & "|" & FormatNumber(Total4(8)) & "|" & FormatNumber(Total4(9)) & ";"
               Call Vsp.AddTable(strFormat2, "", strPrint2)
               For j = 1 To 10
                  Total2(j) = 0
                  Total4(j) = 0
               Next j
               PrevID = NVLI(Rs("PART_TYPE_ID"), 0)
            End If
            '====
         End If
      
         AvgPrice = 0
         TotalLeft = 0
         j = 0
         For Each Lc In Locations
            While j + 1 <= TempCol.Count
               j = j + 1
               Set II = TempCol(j)
               If (II.PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)) And _
                  (II.LOCATION_ID = Lc.LOCATION_ID) And _
                  (TempToDate >= DateToStringInt(II.DOCUMENT_DATE)) Then
                     AvgPrice = AvgPrice + II.NEW_PRICE * II.CURRENT_AMOUNT
                     TotalLeft = TotalLeft + II.CURRENT_AMOUNT
                     GoTo NextLocation
               End If
            Wend
NextLocation:
         Next Lc
         If TotalLeft > 0 Then
            AvgPrice = AvgPrice / TotalLeft
         Else
            AvgPrice = 0
         End If
         
         '== จำนวน
         strPrint = i & "|"
         strPrint = strPrint & NVLS(Rs("PART_TYPE_NAME"), "") & " (" & NVLS(Rs("PART_TYPE_NO"), "") & ")|"
         strPrint = strPrint & NVLS(Rs("PART_DESC"), "") & " (" & NVLS(Rs("PART_NO"), "") & ")|"
         strPrint = strPrint & "จำนวน|"
         strPrint = strPrint & FormatNumber(AvgPrice, "-") & "|"
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_ALL_IMPORT"), 0) - NVLD(Rs("SUM_ALL_EXPORT"), 0), "-") & "|"
         Total1(1) = Total1(1) + (NVLD(Rs("SUM_ALL_IMPORT"), 0) - NVLD(Rs("SUM_ALL_EXPORT"), 0))
         Total2(1) = Total2(1) + (NVLD(Rs("SUM_ALL_IMPORT"), 0) - NVLD(Rs("SUM_ALL_EXPORT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_IMPORT1"), 0), "-") & "|"
         Total1(2) = Total1(2) + (NVLD(Rs("SUM_IMPORT1"), 0))
         Total2(2) = Total2(2) + (NVLD(Rs("SUM_IMPORT1"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_IMPORT2"), 0), "-") & "|"
         Total1(3) = Total1(3) + (NVLD(Rs("SUM_IMPORT2"), 0))
         Total2(3) = Total2(3) + (NVLD(Rs("SUM_IMPORT2"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_IMPORT3"), 0), "-") & "|"
         Total1(4) = Total1(4) + (NVLD(Rs("SUM_IMPORT3"), 0))
         Total2(4) = Total2(4) + (NVLD(Rs("SUM_IMPORT3"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT1_1"), 0), "-") & "|"
         Total1(5) = Total1(5) + (NVLD(Rs("SUM_EXPORT1_1"), 0))
         Total2(5) = Total2(5) + (NVLD(Rs("SUM_EXPORT1_1"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT2"), 0), "-") & "|"
         Total1(6) = Total1(6) + (NVLD(Rs("SUM_EXPORT2"), 0))
         Total2(6) = Total2(6) + (NVLD(Rs("SUM_EXPORT2"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT3"), 0), "-") & "|"
         Total1(7) = Total1(7) + (NVLD(Rs("SUM_EXPORT3"), 0))
         Total2(7) = Total2(7) + (NVLD(Rs("SUM_EXPORT3"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT1_2"), 0), "-") & "|"
         Total1(8) = Total1(8) + (NVLD(Rs("SUM_EXPORT1_2"), 0))
         Total2(8) = Total2(8) + (NVLD(Rs("SUM_EXPORT1_2"), 0))
         
         LeftAmount = (NVLD(Rs("SUM_ALL_IMPORT"), 0) - NVLD(Rs("SUM_ALL_EXPORT"), 0)) + _
                                    NVLD(Rs("SUM_IMPORT1"), 0) + NVLD(Rs("SUM_IMPORT2"), 0) + NVLD(Rs("SUM_IMPORT3"), 0) - _
                                     NVLD(Rs("SUM_EXPORT1_1"), 0) - NVLD(Rs("SUM_EXPORT1_2"), 0) - NVLD(Rs("SUM_EXPORT2"), 0) - NVLD(Rs("SUM_EXPORT3"), 0)
         Total1(9) = Total1(9) + LeftAmount
         Total2(9) = Total2(9) + LeftAmount
         strPrint = strPrint & FormatNumber(LeftAmount, "-") & ";"
         ValueAmount = LeftAmount * AvgPrice
         Total1(10) = Total1(10) + ValueAmount
         Total2(10) = Total2(10) + ValueAmount
         '== จำนวน
         Call Vsp.AddTable(strFormat, "", strPrint)

         '== มูลค่า
         strPrint = "" & "|"
         strPrint = strPrint & "|"
         strPrint = strPrint & "|"
         strPrint = strPrint & "มูลค่า|"
         strPrint = strPrint & FormatNumber(AvgPrice, "-") & "|"
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_ALL_IMPORT_AMT"), 0) - NVLD(Rs("SUM_ALL_EXPORT_AMT"), 0), "-") & "|"
         Total3(1) = Total3(1) + (NVLD(Rs("SUM_ALL_IMPORT_AMT"), 0) - NVLD(Rs("SUM_ALL_EXPORT_AMT"), 0))
         Total4(1) = Total4(1) + (NVLD(Rs("SUM_ALL_IMPORT_AMT"), 0) - NVLD(Rs("SUM_ALL_EXPORT_AMT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_IMPORT1_AMT"), 0), "-") & "|"
         Total3(2) = Total3(2) + (NVLD(Rs("SUM_IMPORT1_AMT"), 0))
         Total4(2) = Total4(2) + (NVLD(Rs("SUM_IMPORT1_AMT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_IMPORT2_AMT"), 0), "-") & "|"
         Total3(3) = Total3(3) + (NVLD(Rs("SUM_IMPORT2_AMT"), 0))
         Total4(3) = Total4(3) + (NVLD(Rs("SUM_IMPORT2_AMT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_IMPORT3_AMT"), 0), "-") & "|"
         Total3(4) = Total3(4) + (NVLD(Rs("SUM_IMPORT3_AMT"), 0))
         Total4(4) = Total4(4) + (NVLD(Rs("SUM_IMPORT3_AMT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT1_1_AMT"), 0), "-") & "|"
         Total3(5) = Total3(5) + (NVLD(Rs("SUM_EXPORT1_1_AMT"), 0))
         Total4(5) = Total4(5) + (NVLD(Rs("SUM_EXPORT1_1_AMT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT2_AMT"), 0), "-") & "|"
         Total3(6) = Total3(6) + (NVLD(Rs("SUM_EXPORT2_AMT"), 0))
         Total4(6) = Total4(6) + (NVLD(Rs("SUM_EXPORT2_AMT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT3_AMT"), 0), "-") & "|"
         Total3(7) = Total3(7) + (NVLD(Rs("SUM_EXPORT3_AMT"), 0))
         Total4(7) = Total4(7) + (NVLD(Rs("SUM_EXPORT3_AMT"), 0))
         strPrint = strPrint & FormatNumber(NVLD(Rs("SUM_EXPORT1_2_AMT"), 0), "-") & "|"
         Total3(8) = Total3(8) + (NVLD(Rs("SUM_EXPORT1_2_AMT"), 0))
         Total4(8) = Total4(8) + (NVLD(Rs("SUM_EXPORT1_2_AMT"), 0))
         
         LeftAmount = (NVLD(Rs("SUM_ALL_IMPORT_AMT"), 0) - NVLD(Rs("SUM_ALL_EXPORT_AMT"), 0)) + _
                                    NVLD(Rs("SUM_IMPORT1_AMT"), 0) + NVLD(Rs("SUM_IMPORT2_AMT"), 0) + NVLD(Rs("SUM_IMPORT3_AMT"), 0) - _
                                     NVLD(Rs("SUM_EXPORT1_1_AMT"), 0) - NVLD(Rs("SUM_EXPORT1_2_AMT"), 0) - NVLD(Rs("SUM_EXPORT2_AMT"), 0) - NVLD(Rs("SUM_EXPORT3_AMT"), 0)
         Total3(9) = Total3(9) + LeftAmount
         Total4(9) = Total4(9) + LeftAmount
         strPrint = strPrint & FormatNumber(LeftAmount, "-") & ";"
         ValueAmount = LeftAmount * AvgPrice
         Total3(10) = Total3(10) + ValueAmount
         Total4(10) = Total4(10) + ValueAmount
         '== มูลค่า

         Sum1 = Sum1 + ValueAmount
         Rs.MoveNext
         If Rs.EOF Then
            Call AdjustPage(Vsp, strFormat, strPrint, 2500)
         End If
         Call Vsp.AddTable(strFormat, "", strPrint)
      Next i

      '====
      If PrevID <> NVLI(Rs("PART_TYPE_ID"), 0) Then
         strPrint2 = "รวม|||จำนวน||" & FormatNumber(Total2(1)) & "|" & FormatNumber(Total2(2)) & "|" & FormatNumber(Total2(3)) & "|" & FormatNumber(Total2(4)) & "|" & FormatNumber(Total2(5)) & "|" & FormatNumber(Total2(6)) & "|" & FormatNumber(Total2(7)) & "|" & FormatNumber(Total2(8)) & "|" & FormatNumber(Total2(9)) & ";"
         Call Vsp.AddTable(strFormat2, "", strPrint2)
         strPrint2 = "รวม|||มูลค่า||" & FormatNumber(Total4(1)) & "|" & FormatNumber(Total4(2)) & "|" & FormatNumber(Total4(3)) & "|" & FormatNumber(Total4(4)) & "|" & FormatNumber(Total4(5)) & "|" & FormatNumber(Total4(6)) & "|" & FormatNumber(Total4(7)) & "|" & FormatNumber(Total4(8)) & "|" & FormatNumber(Total4(9)) & ";"
         Call Vsp.AddTable(strFormat2, "", strPrint2)
         For j = 1 To 10
            Total2(j) = 0
            Total4(j) = 0
         Next j
         PrevID = NVLI(Rs("PART_TYPE_ID"), 0)
      End If
      '====
      
      strPrint2 = "รวม|||จำนวน||" & FormatNumber(Total1(1)) & "|" & FormatNumber(Total1(2)) & "|" & FormatNumber(Total1(3)) & "|" & FormatNumber(Total1(4)) & "|" & FormatNumber(Total1(5)) & "|" & FormatNumber(Total1(6)) & "|" & FormatNumber(Total1(7)) & "|" & FormatNumber(Total1(8)) & "|" & FormatNumber(Total1(9)) & ";"
      Call Vsp.AddTable(strFormat2, "", strPrint2)
      strPrint2 = "รวม|||มูลค่า||" & FormatNumber(Total3(1)) & "|" & FormatNumber(Total3(2)) & "|" & FormatNumber(Total3(3)) & "|" & FormatNumber(Total3(4)) & "|" & FormatNumber(Total3(5)) & "|" & FormatNumber(Total3(6)) & "|" & FormatNumber(Total3(7)) & "|" & FormatNumber(Total3(8)) & "|" & FormatNumber(Total3(9)) & ";"
      Call Vsp.AddTable(strFormat2, "", strPrint2)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
      Set Rs = Nothing
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
      Set TempRs = Nothing
   End If
   
   Set Locations = Nothing
   Set TempCol = Nothing
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
   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
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


