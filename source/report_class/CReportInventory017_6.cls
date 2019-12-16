VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory017_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory017_4"
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
Private Const HEADER_SIZE = 10
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
Private m_PartItems As Collection
Private m_PartItemTxs As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_ProductStatus = New Collection
   Set m_PartItems = New Collection
   Set m_PartItemTxs = New Collection
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
   Set m_PartItems = Nothing
   Set m_PartItemTxs = Nothing
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
   Call Cf.SetFieldValue(0, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "ราคา", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers2.AddField(Cf)
      
   Call Cf.SetFieldValue(6, "^", "รับเข้า", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "จ่ายออก (โอน+ปรับลด+ขาย+ใช้)", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "เบิกใช้", "^")
   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(6, "^", "ขาย", "^")
'   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "จ่ายออก", "^")
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
   
   Call LoadImportPartItemEx(Nothing, m_PartItems, mcolParam("PART_TYPE"), mcolParam("LOCATION_ID"))
   Call LoadImportPartItemEx(Nothing, m_PartItemTxs, mcolParam("PART_TYPE"), mcolParam("LOCATION_ID"), mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   
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

Private Function CompareKeyAsc(Key1 As String, Key2 As String) As Boolean
      CompareKeyAsc = Key1 < Key2
End Function

Private Function GetNextTransaction(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, II As CImportItem, EI As CExportItem, NextFlag As Boolean) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
Static I As Long

   I = I + 1

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
      If NextFlag Then
         Rs2.MoveNext
      End If
   ElseIf ((Not EofFlag1) And EofFlag2) Then
      GetNextTransaction = "E"
      If NextFlag Then
         Rs1.MoveNext
      End If
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
      ElseIf CompareKeyAsc(EI.PART_TYPE_NO, II.PART_TYPE_NO) Then
         GetNextTransaction = "E"
      Else
         GetNextTransaction = "I"
      End If 'Part type no
      '===
      If GetNextTransaction = "I" Then
         If NextFlag Then
            Rs2.MoveNext
         End If
      ElseIf GetNextTransaction = "E" Then
         If NextFlag Then
            Rs1.MoveNext
         End If
      End If
   End If 'Eof flag
End Function

Private Sub GenerateBalance(Txt1 As String, Txt2 As String, Txt3 As String, Bal As CImportItem, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers2.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = 4 Then
         TempStr = FormatNumber(0, "-")
      ElseIf j = 5 Then
         TempStr = FormatNumber(0, "-")
      ElseIf j = 6 Then
         TempStr = FormatNumber(0, "-", 3)
      ElseIf j = 7 Then
         TempStr = FormatNumber(0, "-")
      ElseIf j = 8 Then
         TempStr = FormatNumber(0, "-")
      ElseIf j = 9 Then
         TempStr = FormatNumber(0, "-", 3)
      ElseIf j = 10 Then
         TempStr = FormatNumber(0, "-")
      ElseIf j = 11 Then
         TempStr = FormatNumber(0, "-")
      ElseIf j = 12 Then
         TempStr = FormatNumber(0, "-", 3)
      ElseIf j = 13 Then
         TempStr = FormatNumber(Bal.CURRENT_AMOUNT, "-")
      ElseIf j = 14 Then
         TempStr = FormatNumber(Bal.NEW_PRICE, "-")
      ElseIf j = 15 Then
         TempStr = FormatNumber(Bal.CURRENT_AMOUNT * Bal.NEW_PRICE, "-", 3)
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateSubHeader(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   TempStr = Txt1
   TempWidth = m_Headers2.Fields(1).Width + m_Headers2.Fields(2).Width + m_Headers2.Fields(3).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

   TempStr = Txt2
   TempWidth = 0
   For j = 4 To m_Headers2.Fields.Count
      TempWidth = TempWidth + m_Headers2.Fields(j).Width
   Next j
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "^")
   Call Details.AddField(BodyCf)

   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, Ba As CExportItem)
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
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt2, "<")
   Call Details.AddField(BodyCf)
      
   Set HeadCf = m_Headers2.Fields(3)
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, Txt2, "<")
   Call Details.AddField(BodyCf)
         
   For j = 4 To 15
      Set HeadCf = m_Headers2.Fields(j)
      If (j = 5) Or (j = 8) Or (j = 11) Then
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, "-", HeadCf.BodyAlign)
      ElseIf (j = 4) Or (j = 6) Or (j = 7) Or (j = 9) Or (j = 10) Or (j = 12) Then
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Tot(j), "-", 3), HeadCf.BodyAlign)
      ElseIf (j = 13) Then
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Ba.CURRENT_AMOUNT, "-"), HeadCf.BodyAlign)
      ElseIf (j = 14) Then
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Ba.EXPORT_AVG_PRICE, "-"), HeadCf.BodyAlign)
      ElseIf (j = 15) Then
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, FormatNumber(Ba.EXPORT_TOTAL_PRICE, "-", 3), HeadCf.BodyAlign)
      End If
      Call Details.AddField(BodyCf)
   Next j
      
   Set BodyCf = Nothing
End Sub

Private Function IsIn(TempID As Long, Col As Collection) As Boolean
Dim Li As CPartItem

   IsIn = False
   For Each Li In Col
'''debug.print Li.PART_TYPE_NO & "|" & Li.PART_NO
      If Li.PART_ITEM_ID = TempID Then
         IsIn = True
         Exit Function
      End If
   Next Li
'IsIn = True
End Function

Public Sub InitCurrentBalance(CurrentBalance As CExportItem, II As CImportItem)
   CurrentBalance.CURRENT_AMOUNT = II.CURRENT_AMOUNT
   CurrentBalance.EXPORT_AVG_PRICE = II.NEW_PRICE
   CurrentBalance.EXPORT_TOTAL_PRICE = II.CURRENT_AMOUNT * II.NEW_PRICE
End Sub

Public Sub CalculateCurrentBalance(CurrentBalance As CExportItem, O As Object)
   If O.TX_TYPE = "I" Then
      CurrentBalance.CURRENT_AMOUNT = CurrentBalance.CURRENT_AMOUNT + O.IMPORT_AMOUNT
      CurrentBalance.EXPORT_TOTAL_PRICE = CurrentBalance.EXPORT_TOTAL_PRICE + O.TOTAL_INCLUDE_PRICE
      CurrentBalance.EXPORT_AVG_PRICE = MyDiffEx(CurrentBalance.EXPORT_TOTAL_PRICE, CurrentBalance.CURRENT_AMOUNT)
   ElseIf O.TX_TYPE = "E" Then
      CurrentBalance.CURRENT_AMOUNT = CurrentBalance.CURRENT_AMOUNT - O.EXPORT_AMOUNT
      'CurrentBalance.EXPORT_AVG_PRICE ราคาเฉลี่ยไม่เปลี่ยน
      CurrentBalance.EXPORT_TOTAL_PRICE = CurrentBalance.EXPORT_TOTAL_PRICE - (O.EXPORT_AMOUNT * CurrentBalance.EXPORT_AVG_PRICE)
      
      'ถ้าเป็น export item จะ populate ค่า O ให้
      O.EXPORT_AVG_PRICE = CurrentBalance.EXPORT_AVG_PRICE
      O.EXPORT_TOTAL_PRICE = CurrentBalance.EXPORT_AVG_PRICE * O.EXPORT_AMOUNT
   End If
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
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim OldID As Long

Dim TempLi As CPartItem
Dim k As Long
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim II As CImportItem
Dim EI As CExportItem
Dim O As Object
Dim Ba As CBalanceAccum
Dim TxCode As String
Dim CurrentBalance As CExportItem
Dim SumCurrentBalance As CExportItem

Dim BalanceAccums As Collection
Dim InventoryBals2 As Collection
Dim showFlag As Boolean

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   Set InventoryBals2 = New Collection
   Set BalanceAccums = New Collection
   If mcolParam("FROM_DATE") > 0 Then
      Call LoadInventoryBalanceEx(Nothing, BalanceAccums, InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))), mcolParam("TO_DATE"), "", mcolParam("LOCATION_ID"), , , "N")
   End If

   
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
   
      Set CurrentBalance = New CExportItem
      Set SumCurrentBalance = New CExportItem
      
      Set EI = New CExportItem
      EI.EXPORT_ITEM_ID = -1
      EI.FROM_DATE = mcolParam("FROM_DATE")
      EI.TO_DATE = mcolParam("TO_DATE")
      EI.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
      EI.PIG_FLAG = "N"
      EI.PART_ITEM_ID = -1
      EI.PART_TYPE = mcolParam("PART_TYPE")
      EI.LOCATION_ID = mcolParam("LOCATION_ID")
      EI.OrderBy = 1
      EI.OrderType = 1
'Ei.PART_ITEM_ID = 10463
      Call EI.QueryData(13, Rs1, iCount)

      Set II = New CImportItem
      II.IMPORT_ITEM_ID = -1
      II.FROM_DATE = mcolParam("FROM_DATE")
      II.TO_DATE = mcolParam("TO_DATE")
      II.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
      II.PIG_FLAG = "N"
      II.PART_ITEM_ID = -1
      II.LOCATION_ID = mcolParam("LOCATION_ID")
      II.PART_TYPE = mcolParam("PART_TYPE")
      II.OrderBy = 1
      II.OrderType = 1
'Ii.PART_ITEM_ID = 10463
      Call II.QueryData(7, Rs2, iCount)
   
      Set BodyCf = New CReportField
      For Each TempLi In m_PartItems
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
         
         Set Ba = GetBalanceAccum(BalanceAccums, mcolParam("LOCATION_ID") & "-" & TempLi.PART_ITEM_ID)
         Set II = New CImportItem
         Call BalanceAccum2ImportItem(Ba, II)
         
         showFlag = False
'If TempLi.PART_ITEM_ID = 10451 Then
'''debug.print
'End If
         If (IsIn(TempLi.PART_ITEM_ID, m_PartItemTxs)) Or (II.CURRENT_AMOUNT <> 0) Then
            showFlag = True
         End If

         If showFlag Then
            Call GenerateSubHeader(TempLi.PART_TYPE_NAME & " (" & TempLi.PART_TYPE_NO & ")", TempLi.PART_DESC & " (" & TempLi.PART_NO & ")", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            Call GenerateBalance("ยกมา", "", "", II, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            Call InitCurrentBalance(CurrentBalance, II)
'            Total1(10) = BalLi.NEW_AMOUNT
'            Total1(12) = BalLi.TOTAL_INCLUDE_PRICE
         End If

         If IsIn(TempLi.PART_ITEM_ID, m_PartItemTxs) Then
            TxCode = GetNextTransaction(Rs1, Rs2, II, EI, False) 'เอาค่ามาอย่างเดียว
            If TxCode = "I" Then
               Set O = II
            ElseIf TxCode = "E" Then
               Set O = EI
            End If
            
'            I = 0
            While (TxCode <> "") And (TempLi.PART_NO = O.PART_NO) And (TempLi.PART_TYPE_NO = O.PART_TYPE_NO)
               I = I + 1
               If TxCode = "I" Then
                  Set O = II
               ElseIf TxCode = "E" Then
                  Set O = EI
               End If
               Call CalculateCurrentBalance(CurrentBalance, O)
               
               'สร้าง detail ตรงนี้
               Call m_Details1.ClearField
               For k = 1 To 15
                  Set HeadCf = m_Headers2.Fields(k)
                  If k = 1 Then
                     TempStr = I
                  ElseIf k = 2 Then
                     TempStr = DateToStringExtEx2(O.DOCUMENT_DATE)
                  ElseIf k = 3 Then
                     TempStr = O.DOCUMENT_NO
                  ElseIf k = 4 Then
                     If O.TX_TYPE = "I" Then
                        Amt = O.IMPORT_AMOUNT
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 5 Then
                     If O.TX_TYPE = "I" Then
                        Amt = MyDiffEx(O.TOTAL_INCLUDE_PRICE, O.IMPORT_AMOUNT)
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 6 Then
                     If O.TX_TYPE = "I" Then
                        Amt = O.TOTAL_INCLUDE_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt, "-", 3)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 7 Then
                     If (O.TX_TYPE = "E") Then
                        Amt = O.EXPORT_AMOUNT
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 8 Then
                     If (O.TX_TYPE = "E") Then
                        Amt = O.EXPORT_AVG_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 9 Then
                     If (O.TX_TYPE = "E") Then
                        Amt = O.EXPORT_TOTAL_PRICE
                     Else
                        Amt = 0
                     End If
                     TempStr = FormatNumber(Amt, "-", 3)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 10 Then
                     Amt = 0
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 11 Then
                     Amt = 0
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 12 Then
                     Amt = 0
                     TempStr = FormatNumber(Amt, "-", 3)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 13 Then
                     Amt = CurrentBalance.CURRENT_AMOUNT
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 14 Then
                     Amt = CurrentBalance.EXPORT_AVG_PRICE
                     TempStr = FormatNumber(Amt, "-")
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  ElseIf k = 15 Then
                     Amt = CurrentBalance.EXPORT_TOTAL_PRICE
                     TempStr = FormatNumber(Amt, "-", 3)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  Else
                     TempStr = ""
                     Amt = 0
                  End If
                  
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               Next k
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               TxCode = GetNextTransaction(Rs1, Rs2, II, EI, True) 'เลื่อน Rs อย่างเดียว
               TxCode = GetNextTransaction(Rs1, Rs2, II, EI, False) 'Populate ค่ามาอย่างเดียว
               If TxCode = "I" Then
                  Set O = II
               ElseIf TxCode = "E" Then
                  Set O = EI
               End If
            Wend
         End If
         
         If showFlag Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1, CurrentBalance)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            SumCurrentBalance.CURRENT_AMOUNT = SumCurrentBalance.CURRENT_AMOUNT + CurrentBalance.CURRENT_AMOUNT
            SumCurrentBalance.EXPORT_TOTAL_PRICE = SumCurrentBalance.EXPORT_TOTAL_PRICE + CurrentBalance.EXPORT_TOTAL_PRICE
            SumCurrentBalance.EXPORT_AVG_PRICE = 0
         End If
      Next TempLi
   
      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, SumCurrentBalance)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set BodyCf = Nothing
      Vsp.EndDoc
      Call CloseExportFile(Vsp)
      
      Set EI = Nothing
      Set II = Nothing
      
      If Rs1.State = adStateOpen Then
         Rs1.Close
      End If
      Set Rs1 = Nothing
      
      If Rs2.State = adStateOpen Then
         Rs2.Close
      End If
      Set Rs2 = Nothing
      Set InventoryBals2 = Nothing
      Set CurrentBalance = Nothing
      Set SumCurrentBalance = Nothing
      
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



