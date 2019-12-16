VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory021"
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

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_ProductStatus As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_ProductStatus = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
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
   
   Call m_Headers1.ClearField
   Call m_Headers0.ClearField
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(7, "^", "", "^", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เกิด", "^", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอนเข้า", "^", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "รับเข้า", "^", "^")
   Call m_Headers0.AddField(Cf)
   
   m_Left = 0
   ExportWidth = 0
   For Each Sgi In m_ProductStatus
      If Sgi.SELECT_FLAG = "Y" Then
         Call Cf.SetFieldValue(2, "^", Sgi.STATUS_NAME, "^", Sgi.ST_STATUS_ID)
         Call m_Headers1.AddField(Cf)
         
         ExportWidth = ExportWidth + 2
         m_Left = m_Left + 2
      End If
   Next Sgi
   Call Cf.SetFieldValue(2, "^", "โอนออก", "^")
   Call m_Headers1.AddField(Cf)
   ExportWidth = ExportWidth + 2
   
   Call Cf.SetFieldValue(CDbl(ExportWidth), "^", "จ่ายออก", "^", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", "^")
   Call m_Headers1.AddField(Cf)
   Call m_Headers0.AddField(Cf)
   
   m_Left = m_Left + 4
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

   Call LoadProductStatusEx(Nothing, m_ProductStatus, mcolParam("STATUS_GROUP_ID"))
   
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
                                       vbCrLf & FromDate & " - " & ToDate & " (" & EmptyToString(mcolParam("HOUSE_NAME"), "รวม") & ")"

   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetPartItem(m_TempCol As Collection, TempID As Long) As CPartItem
On Error Resume Next
Dim EI As CPartItem

   Set EI = m_TempCol(Trim(Str(TempID)))
   If EI Is Nothing Then
      Set GetPartItem = Nothing
   Else
      Set GetPartItem = EI
   End If
End Function

Private Function GetImportItem(m_TempCol As Collection, TempKey As String) As CImportItem
On Error Resume Next
Dim EI As CImportItem
Static TempEi As CImportItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CImportItem
      End If
      Set GetImportItem = TempEi
   Else
      Set GetImportItem = EI
   End If
End Function

Private Function CompareKeyEx(Key1 As String, Key2 As String, CompareType As Long) As Boolean
      If CompareType = 1 Then
         CompareKeyEx = Key1 < Key2
      Else
         CompareKeyEx = Key2 < Key1
      End If
End Function

Private Function GetNextTransaction(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, II As CImportItem, EI As CExportItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
   
   'Export
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call EI.PopulateFromRS(13, Rs1)
'''debug.print Ei.DOCUMENT_NO
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
      ElseIf CompareKeyEx(EI.PART_TYPE_NO, II.PART_TYPE_NO, mcolParam("ORDER_TYPE")) Then
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

Private Function IsStatusIn(StatusID As Long) As Boolean
Dim D As CSGroupItem

   IsStatusIn = False
   For Each D In m_ProductStatus
      If D.ST_STATUS_ID = StatusID Then
         IsStatusIn = True
         Exit Function
      End If
   Next D
End Function

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
Dim StatusMatch As Boolean
Dim PiCount As Long
Dim iCount As Long
Dim OldColor As Double
Dim LeftAmount As Double
Dim ImportItems As Collection
Dim ExportItems As Collection
Dim NewDate As Date

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Houses = New Collection
'   If mcolParam("SALE_FLAG") = "Y" Then
      Call LoadLocation(Nothing, Houses, 1, "", mcolParam("HOUSE_ID"), 1)
'   ElseIf mcolParam("SALE_FLAG") = "N" Then
'      Call LoadLocation(Nothing, Houses, 1, "N", mcolParam("HOUSE_ID"))
'   End If
   
   Set ImportItems = New Collection
   Set ExportItems = New Collection
   
   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   If mcolParam("FROM_DATE") > 0 Then
      Call LoadPigHouseImportAmount(Nothing, ImportItems, -1, InternalDateToDate(DateToStringIntHi(NewDate)), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"), , , , mcolParam("BATCH_ID"))
      Call LoadPigHouseExportAmount(Nothing, ExportItems, -1, InternalDateToDate(DateToStringIntHi(NewDate)), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"), , , mcolParam("BATCH_ID"))
   End If

   Set Partitems = New Collection
   
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
         
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      Vsp.TableBorder = tbAll
      I = 0
      
      For Each HS In Houses
         For j = 1 To UBound(Total2)
            Total2(j) = 0
         Next j

         Call LoadImportPig(Nothing, Partitems, HS.LOCATION_ID, PigTypeToCode(mcolParam("PIG_TYPE_ID")), mcolParam("BATCH_ID"))
Call glbDaily.PigSelectionsortEx(Partitems, 1, Partitems.Count)
         PiCount = 0
         For Each Pi In Partitems
            PiCount = PiCount + 1
            '===หัว
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            
            Call m_Details1.ClearField
            Set BodyCf = New CReportField
            
            Call BodyCf.SetFieldValue(7, "<", HS.LOCATION_NAME & " (" & HS.LOCATION_NO & ")")
            Call m_Details1.AddField(BodyCf)
            Call BodyCf.SetFieldValue(4, "<", Pi.PART_NO & " (" & Pi.PIG_TYPE & ")")
            Call m_Details1.AddField(BodyCf)
            Call BodyCf.SetFieldValue(m_Left, "<", "")
            Call m_Details1.AddField(BodyCf)
            
            Set BodyCf = Nothing
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            Vsp.FontBold = False
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            '===หัว
            
            '== ยกมา
            m_Details1.ClearField
            j = 0
            For Each HeadCf In m_Headers1.Fields
               j = j + 1
               
               Set BodyCf = New CReportField
               If j = 1 Then
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
               ElseIf j = 2 Then
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "*** ยกมา ***", HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
               ElseIf j = 3 Then
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "", HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
               ElseIf j = m_Headers1.Fields.Count Then 'อันสุดท้าย
   Set II = GetImportItem(ImportItems, Pi.PART_ITEM_ID & "-" & HS.LOCATION_ID)
   Set EI = GetExportItem(ExportItems, Pi.PART_ITEM_ID & "-" & HS.LOCATION_ID)
   Amt = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
   LeftAmount = Amt
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Trim(Str(FormatNumber(Amt))), HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
               
                  Total1(j) = Amt
                  Total2(j) = Total2(j) + 0 'Amt
                  Total3(j) = Total3(j) + 0 'Amt
               Else
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, 0, HeadCf.BodyAlign)
                  Call m_Details1.AddField(BodyCf)
               End If

               Set BodyCf = Nothing
            Next HeadCf
            
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            Vsp.FontBold = False
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            '==
            
               '=== Detail
               Set EI = New CExportItem
               EI.EXPORT_ITEM_ID = -1
               EI.FROM_DATE = mcolParam("FROM_DATE")
               EI.TO_DATE = mcolParam("TO_DATE")
               EI.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
               EI.PIG_FLAG = "Y"
               EI.PART_ITEM_ID = Pi.PART_ITEM_ID
               EI.LOCATION_ID = HS.LOCATION_ID
               EI.OrderBy = mcolParam("ORDER_BY")
               EI.OrderType = mcolParam("ORDER_TYPE")
               EI.BATCH_ID = mcolParam("BATCH_ID")
               Call EI.QueryData(17, Rs1, iCount)
'               Set Ei = Nothing
               
               Set II = New CImportItem
               II.IMPORT_ITEM_ID = -1
               II.FROM_DATE = mcolParam("FROM_DATE")
               II.TO_DATE = mcolParam("TO_DATE")
               II.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
               II.PIG_FLAG = "Y"
               II.PART_ITEM_ID = Pi.PART_ITEM_ID
               II.LOCATION_ID = HS.LOCATION_ID
               II.OrderBy = mcolParam("ORDER_BY")
               II.OrderType = mcolParam("ORDER_TYPE")
               II.BATCH_ID = mcolParam("BATCH_ID")
               Call II.QueryData(10, Rs2, iCount)
'               Set II = Nothing
               
               TxCode = "X"
               While TxCode <> ""
                  TxCode = GetNextTransaction(Rs1, Rs2, II, EI)
                  If TxCode <> "" Then
                     I = I + 1
                     If TxCode = "I" Then
                        Set O = II
                     ElseIf TxCode = "E" Then
                        Set O = EI
                     End If
                                          
                     Set BodyCf = New CReportField
                     m_Details1.ClearField
                     
                     j = 0
                     StatusMatch = False
                     
                     If O.DOCUMENT_NO = "T453/22610" Then
                        Debug.Print
                     End If
                     
                     For Each HeadCf In m_Headers1.Fields
                        j = j + 1
                        If j = 1 Then
                           If I = 11 Then
                              ''debug.print
                           End If
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Trim(Str(I)), HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                        ElseIf j = 2 Then
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, DateToStringExtEx2(O.DOCUMENT_DATE), HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                        ElseIf j = 3 Then
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, O.DOCUMENT_NO, HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                        ElseIf j = 4 Then 'จำนวนเกิด
                           If O.DOCUMENT_TYPE = 5 Then
                              Amt = O.IMPORT_AMOUNT
                           Else
                              Amt = 0
                           End If
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Amt, "-", 0), HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                           
                           Total1(j) = Total1(j) + Amt
                           Total2(j) = Total2(j) + Amt
                           Total3(j) = Total3(j) + Amt
                        ElseIf j = 5 Then 'โอนเข้า
                           If (O.DOCUMENT_TYPE = 6) Or (O.DOCUMENT_TYPE = 11) Or (O.DOCUMENT_TYPE = 7) Or (O.DOCUMENT_TYPE = 8) Or (O.DOCUMENT_TYPE = 9) Or (O.DOCUMENT_TYPE = 12) Or (O.DOCUMENT_TYPE = 888) Then
                              If (TxCode = "I") Then
                                 Amt = O.IMPORT_AMOUNT
                              Else
                                 Amt = 0
                              End If
                           Else
                              Amt = 0
                           End If
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Amt, "-", 0), HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                           
                           Total1(j) = Total1(j) + Amt
                           Total2(j) = Total2(j) + Amt
                           Total3(j) = Total3(j) + Amt
                        ElseIf j = m_Headers1.Fields.Count - 1 Then 'โอนออก
                           If TxCode = "E" Then
                              If StatusMatch Then
                                 Amt = 0
                              Else
                                 Amt = O.EXPORT_AMOUNT
                              End If
                           End If
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Amt, "-", 0), HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                           
                           Total1(j) = Total1(j) + Amt
                           Total2(j) = Total2(j) + Amt
                           Total3(j) = Total3(j) + Amt
                        ElseIf j = m_Headers1.Fields.Count Then
If TxCode = "I" Then
   LeftAmount = LeftAmount + O.IMPORT_AMOUNT
ElseIf TxCode = "E" Then
   LeftAmount = LeftAmount - O.EXPORT_AMOUNT
End If
                           Amt = LeftAmount 'O.CURRENT_AMOUNT
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Amt, "-", 0), HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                        
                           Total1(j) = Amt
                           Total2(j) = Total2(j) + 0
                           Total3(j) = Total3(j) + 0
                        Else
                           If TxCode = "E" Then
                              If HeadCf.TempKey = O.PIG_STATUS Then
If O.HOUSE_SALE_FLAG = "Y" Then 'ไม่ต้องเอาสถานะมาแสดงถ้าเป็นเบิกจากเรือนขาย
                                 Amt = 0
Else
                                 Amt = O.EXPORT_AMOUNT
                                 StatusMatch = True
End If
                              Else
                                 Amt = 0
                              End If
                           Else
                              Amt = 0
                           End If
                           Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Amt, "-", 0), HeadCf.BodyAlign)
                           Call m_Details1.AddField(BodyCf)
                        
                           Total1(j) = Total1(j) + Amt
                           Total2(j) = Total2(j) + Amt
                           Total3(j) = Total3(j) + Amt
                        End If
                     Next HeadCf
                     Set BodyCf = Nothing
                  
                     Call m_Details1.GetString(2, TempStr1, TempStr2)
                     Vsp.FontBold = False
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                  End If
                  
                  DoEvents
               Wend
               '=== Detail
               
               '== Summary for each week
               Total2(j) = Total2(j) + Total1(j)
               Call CreateTotal(Vsp, Total1, TempStr1, TempStr2, "รวม", Pi.PART_NO & " (" & Pi.PIG_TYPE & ")")
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               '== Summary for each week
         Next Pi
      
         If PiCount > 0 Then
            '== Summary for each house
            Total3(j) = Total3(j) + Total2(j)
            
            Call CreateTotal(Vsp, Total2, TempStr1, TempStr2, "รวม", HS.LOCATION_NAME)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            '== Summary for each week
         End If
      Next HS
         
      '== Summary for each house
      Call CreateTotal(Vsp, Total3, TempStr1, TempStr2, "สุทธิ", "")
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      '== Summary for each week
         
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
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub CreateTotal(Vsp As VSPrinter, Tot() As Double, FormatText As String, BodyText As String, Txt1 As String, Txt2 As String)
Dim D As CReportField
Dim j As Long
Dim TempStr1 As String
Dim TempStr2 As String
Dim BodyCf As CReportField
Dim alngX() As Long
Dim strFormat As String
Dim Amt As String

   Call m_Details1.ClearField
   Set BodyCf = New CReportField
                     
   j = 0
   For Each D In m_Headers1.Fields
      j = j + 1
      
      If j = 1 Then
         Amt = Txt1
      ElseIf j = 2 Then
         Amt = Txt2
      ElseIf j > 3 Then
         Amt = FormatNumber(Tot(j), "-", 0)
      Else
         Amt = ""
      End If
      
      Call BodyCf.SetFieldValue(D.Width, D.Align, Amt, D.BodyAlign)
      Call m_Details1.AddField(BodyCf)
   Next D
   
   Call m_Details1.GetString(2, FormatText, BodyText)
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


