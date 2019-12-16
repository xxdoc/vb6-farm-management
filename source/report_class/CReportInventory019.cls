VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory019"
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
Private Const HEADER_FORMAT1 = "^5|^18;"
Private Const HEADER_FORMAT2 = "^5|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2|^2;"
Private Const DETAIL_FORMAT1 = "<5|<22;"
Private Const DETAIL_FORMAT2 = "<5|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
Private Const FOOTER_FORMAT = "^5|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2|>2;"
'วัตถุดิบ|ยกมา|รับเข้า|โอนเข้า|ปรับเพิ่ม|เบิกใช้|โอนออก|ปรับลด|ขาย|คงเหลือ|ยกมา|รับเข้า|โอนเข้า|ปรับเพิ่ม|เบิกใช้|โอนออก|ปรับลด|ขาย|คงเหลือ;"

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
   
'   strFormat = VSP_CalTable(HEADER_FORMAT1, mdblWid, alngX)
'   strPrint = "วัตถุดิบ|จำนวน;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable(HEADER_FORMAT2, mdblWid, alngX)
   strPrint = "วัตถุดิบ|ราคา|มูลค่า|ยกมา|รับเข้า|โอนเข้า|ปรับเพิ่ม|เบิกใช้|โอนออก|ปรับลด|ขาย|คงเหลือ;"
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

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim cData As CPartLocation
Dim iCount As Long
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim strFormat2 As String
Dim strPrint As String
Dim strPrint2 As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim Sum2 As Double
Dim Partitems As Collection
Dim Pi As CPartItem
Dim PrevPI As CPartItem
Dim Value1(30) As Double
Dim Total1(30) As Double
Dim Total2(30) As Double
Dim PrevID As Long
Dim AllImports As Collection
Dim AllExports As Collection
Dim PatialImports As Collection
Dim PatialExports As Collection
Dim II As CImportItem
Dim EI As CExportItem
Dim Amt As Double
Dim Temp1 As Double
Dim Temp2  As Double
Dim InventoryBals As Collection
Dim Ib As CImportItem
Dim Lc As CLocation
Dim Locations As Collection
Dim BalanceAccums As Collection

   RName = "genDoc"
   
   Set Partitems = New Collection
   Call LoadPartItem(Nothing, Partitems)

   Set Locations = New Collection
   Call LoadLocation(Nothing, Locations, 2, "", mcolParam("LOCATION_ID"))
   
'   Set AllImports = New Collection
'   Call LoadImportPrice2(Nothing, AllImports, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
'
'   Set AllExports = New Collection
'   Call LoadExportPrice7(Nothing, AllExports, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
   
   Set PatialImports = New Collection
   Call LoadImportPrice3(Nothing, PatialImports, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
   
   Set PatialExports = New Collection
   Call LoadExportPrice8(Nothing, PatialExports, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
   
'   Set InventoryBals = New Collection
'   'Uset TO_DATE to be start date
'   If mcolParam("TO_DATE") <= 0 Then
'      Call LoadInventoryBalance(Nothing, InventoryBals, InternalDateToDate(DateToStringIntHi(-2)), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
'   Else
'      Call LoadInventoryBalance(Nothing, InventoryBals, InternalDateToDate(DateToStringIntHi(mcolParam("TO_DATE"))), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("LOCATION_ID"))
'   End If
   
   If mcolParam("FROM_DATE") > 0 Then
      Set BalanceAccums = New Collection
      Call LoadInventoryBalanceExByPart(Nothing, BalanceAccums, InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))), mcolParam("TO_DATE"), "", mcolParam("LOCATION_ID"), , , "N")
   End If
   Dim BaAc As CBalanceAccum
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set cData = New CPartLocation
   cData.PART_LOCATION_ID = -1
   cData.LOCATION_ID = mcolParam("LOCATION_ID")
   cData.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
   cData.OrderBy = mcolParam("ORDER_BY")
   cData.OrderType = mcolParam("ORDER_TYPE")
   Call cData.QueryData(5, Rs, iCount)

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
'    strPrint = "วัตถุดิบ|1ยกมา|2รับเข้า|3โอนเข้า|4ปรับเพิ่ม|5เบิกใช้|6โอนออก|7ปรับลด|8ขาย|9คงเหลือ|ยกมา|รับเข้า|โอนเข้า|ปรับเพิ่ม|เบิกใช้|โอนออก|ปรับลด|ขาย|คงเหลือ;"

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      strFormat = VSP_CalTable(DETAIL_FORMAT2, mdblWid, alngX)
      strFormat2 = VSP_CalTable(FOOTER_FORMAT, mdblWid, alngX)
      Vsp.TableBorder = tbAll
      Sum1 = 0
   
      If Not Rs.EOF Then
         Set Pi = GetPartItem(Partitems, NVLI(Rs("PART_ITEM_ID"), 0))
         PrevID = Pi.PART_TYPE
      End If
      
      Sum1 = 0
      Sum2 = 0
      While Not Rs.EOF
         Set PrevPI = Pi
         Set Pi = GetPartItem(Partitems, NVLI(Rs("PART_ITEM_ID"), 0))
         '===
         If (PrevPI.PART_TYPE <> Pi.PART_TYPE) Then
            strPrint2 = "รวม (" & PrevPI.PART_TYPE_NAME & ")|"
            For j = 1 To 11
               If j = 11 Then
                  strPrint2 = strPrint2 & FormatNumber(Total1(j)) & ";"
               Else
                  If j <> 1 Then
                     strPrint2 = strPrint2 & FormatNumber(Total1(j)) & "|"
                  Else
                     strPrint2 = strPrint2 & "" & "|"
                  End If
               End If
            Next j
            
            For j = 1 To 18
               Total1(j) = 0
            Next j
            Set PrevPI = Pi

            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat2, "", strPrint2)
            Vsp.FontBold = False
         End If
         '===
         Set BaAc = GetBalanceAccum(BalanceAccums, Trim(Str(Pi.PART_ITEM_ID)))
         
         strPrint = Pi.PART_DESC & " (" & Pi.PART_NO & ")|"
         For j = 1 To 18
            If j = 1 Then
               Amt = MyDiff(BaAc.TOTAL_INCLUDE_PRICE, BaAc.BALANCE_AMOUNT)
            ElseIf j = 2 Then
               Amt = BaAc.TOTAL_INCLUDE_PRICE
            ElseIf ((j = 3) Or (j = 12)) Then
               Amt = BaAc.BALANCE_AMOUNT
            ElseIf ((j = 4) Or (j = 13)) Then
               Set II = GetImportItem(PatialImports, Trim(Str(Pi.PART_ITEM_ID)) & "-" & Trim(Str(1)) & "-N")
               If j = 4 Then
                  Amt = II.SUM_ALL_IMPORT
               Else
                  Amt = II.SUM_ALL_IMPORT_AMT
               End If
            ElseIf ((j = 5) Or (j = 14)) Then
               Set II = GetImportItem(PatialImports, Trim(Str(Pi.PART_ITEM_ID)) & "-" & Trim(Str(3)) & "-N")
               If j = 5 Then
                  Amt = II.SUM_ALL_IMPORT
               Else
                  Amt = II.SUM_ALL_IMPORT_AMT
               End If
            ElseIf ((j = 6) Or (j = 15)) Then
               Set II = GetImportItem(PatialImports, Trim(Str(Pi.PART_ITEM_ID)) & "-" & Trim(Str(4)) & "-N")
               If j = 6 Then
                  Amt = II.SUM_ALL_IMPORT
               Else
                  Amt = II.SUM_ALL_IMPORT_AMT
               End If
            ElseIf ((j = 7) Or (j = 16)) Then
               Set EI = GetExportItem(PatialExports, Trim(Str(Pi.PART_ITEM_ID)) & "-" & Trim(Str(2)) & "-N")
               If j = 7 Then
                  Amt = EI.SUM_ALL_EXPORT
               Else
                  Amt = EI.SUM_ALL_EXPORT_AMT
               End If
            ElseIf ((j = 8) Or (j = 17)) Then
               Set EI = GetExportItem(PatialExports, Trim(Str(Pi.PART_ITEM_ID)) & "-" & Trim(Str(3)) & "-N")
               If j = 8 Then
                  Amt = EI.SUM_ALL_EXPORT
               Else
                  Amt = EI.SUM_ALL_EXPORT_AMT
               End If
            ElseIf ((j = 9) Or (j = 18)) Then
               Set EI = GetExportItem(PatialExports, Trim(Str(Pi.PART_ITEM_ID)) & "-" & Trim(Str(4)) & "-N")
               If j = 9 Then
                  Amt = EI.SUM_ALL_EXPORT
               Else
                  Amt = EI.SUM_ALL_EXPORT_AMT
               End If
            ElseIf ((j = 10) Or (j = 19)) Then
               Set EI = GetExportItem(PatialExports, Trim(Str(Pi.PART_ITEM_ID)) & "-" & Trim(Str(2)) & "-Y")
               If j = 10 Then
                  Amt = EI.SUM_ALL_EXPORT
               Else
                  Amt = EI.SUM_ALL_EXPORT_AMT
               End If
            Else
               Amt = 0
            End If

            If j = 11 Then
               Amt = Value1(3) + Value1(4) + Value1(5) + Value1(6) - Value1(7) - Value1(8) - Value1(9) - Value1(10)
               strPrint = strPrint & FormatNumber(Amt, "-") & "|"
            ElseIf j = 20 Then
               Amt = Value1(12) + Value1(13) + Value1(14) + Value1(15) - Value1(16) - Value1(17) - Value1(18) - Value1(19)
               strPrint = strPrint & FormatNumber(Amt, "-") & ";"
            Else
               strPrint = strPrint & FormatNumber(Amt, "-") & "|"
            End If
            Value1(j) = Amt
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
         Next j

         Rs.MoveNext
         If Rs.EOF Then
            Call AdjustPage(Vsp, strFormat, strPrint, 2500)
         End If
         Call Vsp.AddTable(strFormat, "", strPrint)
      Wend

      '===
      strPrint2 = "รวม (" & Pi.PART_TYPE_NAME & ")|"
      For j = 1 To 11
         If j = 11 Then
            strPrint2 = strPrint2 & FormatNumber(Total1(j)) & ";"
         Else
            If j <> 1 Then
               strPrint2 = strPrint2 & FormatNumber(Total1(j)) & "|"
            Else
               strPrint2 = strPrint2 & "" & "|"
            End If
         End If
      Next j

      For j = 1 To UBound(Total1)
         Total1(j) = 0
      Next j
      PrevID = Pi.PART_TYPE

      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat2, "", strPrint2)
      Vsp.FontBold = False
      '===

      '===
      strPrint2 = "รวมสุทธิ           |"
      For j = 1 To 11
         If j = 11 Then
            strPrint2 = strPrint2 & FormatNumber(Total2(j)) & ";"
         Else
            If j <> 1 Then
               strPrint2 = strPrint2 & FormatNumber(Total2(j)) & "|"
            Else
               strPrint2 = strPrint2 & "" & "|"
            End If
         End If
      Next j

      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat2, "", strPrint2)
      Vsp.FontBold = False
      '===
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   
   Set Locations = Nothing
   Set Partitems = Nothing
   Set AllImports = Nothing
   Set AllExports = Nothing
   Set PatialImports = Nothing
   Set PatialExports = Nothing
   Set BalanceAccums = Nothing
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


