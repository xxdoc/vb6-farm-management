VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory069"
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

'Private tempDeFor As String
'For Each tempExposeType In m_ExposeType
'   tempDeFor = tempDeFor & ">1"
'Next tempExposeType
'Private Const DETAIL_FORMAT = "^1|>1|>1|" & tempDeFor & "^1|>1|^1|^1|^1;"


Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_PartItems As Collection
Private m_ExposeType As Collection
Private m_HGroupItem As Collection
Private tempExposeType As CExportItem

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
   Set m_ExposeType = New Collection
   Set m_HGroupItem = New Collection
   
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
   Set m_ExposeType = Nothing
   Set m_HGroupItem = Nothing
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
Dim txtColumn As String


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
   Call Cf.SetFieldValue(1, "^", "ว/ด/ป", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "ยอดยกมา", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "ยอดรับเข้า", ">")
   Call m_Headers0.AddField(Cf)
   'forloop x 1 = x
   Call Cf.SetFieldValue(m_ExposeType.Count + 1, "^", "จำนวนที่เบิกใช้", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ผู้เบิก", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "คงเหลือ", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "ว/ด/ป", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "ว/ด/ป", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ผู้รับผิดชอบ", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "", ">")
   Call m_Headers1.AddField(Cf)
   For Each tempExposeType In m_ExposeType
      If Left(tempExposeType.HOUSE_GROUP_NAME, 4) = "สุกร" Then
         txtColumn = Mid(tempExposeType.HOUSE_GROUP_NAME, 5, Len(tempExposeType.HOUSE_GROUP_NAME))
      Else
         txtColumn = tempExposeType.HOUSE_GROUP_NAME
      End If
      Call Cf.SetFieldValue(1, "^", txtColumn, ">")
      Call m_Headers1.AddField(Cf)
   Next tempExposeType
'         Call Cf.SetFieldValue(1, "^", "โอนออก", ">")
'         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(1, "^", "ขาย", ">")
         Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^<", "", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "ที่ผลิต", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "ที่หมดอายุ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", "^")
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
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = 10
   


   Call LoadHGroupItem(Nothing, m_HGroupItem)
   'HGroupItems
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & "หน่วยงาน ____________   " & "ประจำเดือน " & EmptyToString(mcolParam("MONTH_NAME"), "____") & "  " & EmptyToString(mcolParam("YEAR_NO"), "_____")
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Private Function LoadColumnType(C As CExportItem, Cl As Collection)
Dim iCount_c As Long
Dim Ia As Long
Dim tempHg As CHGroupItem
Dim tempEX As CExportItem
Dim Rs_c As ADODB.Recordset
Dim Y_all As Boolean
         
      Set Rs_c = New ADODB.Recordset

      Call C.QueryData(62, Rs_c, iCount_c)   ' 1
      While Not Rs_c.EOF
      Ia = Ia + 1
      Set tempEX = New CExportItem
      Call tempEX.PopulateFromRS(62, Rs_c)
   
'      If Not (C Is Nothing) Then
'         C.AddItem (TempData.EXPOSE_TYPE_NAME)
'         C.ItemData(I) = TempData.EXPOSE_TYPE_ID
'      End If
      
      If Not (Cl Is Nothing) Then
        ' TempData.EXIST_FLAG = "N"
         Call Cl.Add(tempEX, Trim(Str(tempEX.HOUSE_GROUP_ID)))
      End If
      
      Set tempEX = Nothing
      Rs_c.MoveNext
   Wend
   
   Set Rs_c = Nothing
   'Set D = Nothing
End Function
Private Function LoadColumnTypeHG(C As CExportItem, Cl As Collection)
Dim iCount_c As Long
Dim II As Long
Dim tempHg As CHGroupItem
Dim tempEX As CHGroupItem
Dim Rs_c As ADODB.Recordset
Dim Y_all As Boolean
         
         Set Rs_c = New ADODB.Recordset

      Call C.QueryData(62, Rs_c, iCount_c)   ' 1
      For II = 1 To iCount_c
         Set tempHg = GetObject("CHGroupItem", m_ExposeType, Trim(Str(NVLI(Rs_c("HOUSE_GROUP_ID"), -1))))
         If Not (tempHg Is Nothing) Then
            tempHg.EXIST_FLAG = "Y"
          End If
          
           For Each tempEX In m_ExposeType
               If tempEX = "N" Then
                  Y_all = False
                  Exit For
               End If
            Next tempEX
            
            If Y_all = True Then
              Exit Function
            End If
            
            Y_all = True
            Rs_c.MoveNext
      Next II
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
Dim cData As CExportItem
Dim haveData As CExportItem
Dim I As Long
Dim j As Long
Dim k As Long
Dim B As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr As String
Dim TempStr1 As String
Dim TempStr2 As String
'Dim Pi As CPartItem
'Dim OldPi As CPartItem
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
'Dim OldBaDistinct As CBalanceAccum
'Dim BaDistinct As CBalanceAccum
Dim TempFromDate  As Date
Dim TempToDate As Date
Dim TempImDate  As Date
Dim TempExDate As Date
Dim iCount As Long
Dim exposetypeCount As Long
Dim strPrint As String
Dim tempDeFor As String
Dim DETAIL_FORMAT As String
Dim tempPartName As String
Dim tempBALANCE_AMOUNT As Double
Dim SUM_ALL_IMPORT As Double
Dim tempCl As CImportItem
Dim subTempCl As CImportItem
Dim tempB As Long
Dim tempC As Long
Dim tempHgroup As CHGroupItem

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
      
      
   Set Rs = New ADODB.Recordset
   Set PatialImports = New Collection
   Set BalanceAccums = New Collection
      
   If Not initDoc Then Exit Function

   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), TempFromDate, TempToDate)
   
   If TempFromDate > 0 Then
      Call LoadInventoryBalanceExByPart(Nothing, BalanceAccums, TempFromDate, TempFromDate, "", , , , "N", -1)
   End If
   
   Set PatialImports = New Collection
   Call LoadImportPrice7(Nothing, PatialImports, TempFromDate, TempToDate, mcolParam("PART_NO"), mcolParam("PART_GROUP"))

   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   Set cData = New CExportItem
   cData.EXPORT_ITEM_ID = -1
   cData.FROM_DATE = TempFromDate
   cData.TO_DATE = TempToDate
   cData.PIG_FLAG = "N"
   cData.PART_NO = mcolParam("PART_NO")
   cData.PART_GROUP_ID = mcolParam("PART_GROUP")
   cData.OrderBy = 13
   cData.OrderType = mcolParam("ORDER_TYPE")
   Call cData.QueryData(1, Rs, iCount)   ' 1
   
  Set haveData = cData
  Call LoadColumnType(haveData, m_ExposeType)
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Call StartExportFile(Vsp)

      Vsp.StartDoc
      Vsp.TableBorder = tbAll
 
      For I = 1 To iCount
         If I = 1 Then
               For Each tempExposeType In m_ExposeType
                  tempDeFor = tempDeFor & ">1|"
               Next tempExposeType
               DETAIL_FORMAT = "^1|>1|>1|" & tempDeFor & ">1|" & "<2|>1|^1|^1|^2;"     'ตัดโอน
         End If
      
         strFormat = VSP_CalTable(DETAIL_FORMAT, mdblWid, alngX)

         If TempExDate <> InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), "")) Then   ' มีการเปลี่ยนวัน
                                  
                  Set tempCl = GetObject("CImport", PatialImports, Trim(NVLS(Rs("PART_NO"), "")), False)
                  If Not (tempCl Is Nothing) Then                                   ' เจอใน Coll
                        
                          ''debug.print NVLS(Rs("PART_NO"), "")
                          
                          If NVLS(Rs("PART_DESC"), "") <> tempPartName Then               ' เช็คบรรทัดแรกของวัตถุดิบตัวเดียวกัน
                              If B <> 0 Then
                              For tempB = B To 27 - 1            'ปริ้นให้ครบบรรทัดที่เหลือ
                                  strPrint = "|||"
                                  For Each tempExposeType In m_ExposeType
                                     strPrint = strPrint & "|"
                                  Next tempExposeType
                                  strPrint = strPrint & "|"                 ' ตัดโอน
                                   strPrint = strPrint & "|||||"
                                   Call Vsp.AddTable(strFormat, "", strPrint)
                                   strPrint = ""
                                   B = B + 1
                               Next tempB
                               strPrint = "รวม|" & FormatNumber(Total1(1), "0", 3) & "|" & FormatNumber(Total1(2), "0", 3) & "|"
                               tempC = 4
                               For Each tempExposeType In m_ExposeType
                                  strPrint = strPrint & FormatNumber(Total1(tempC), "0", 3) & "|"
                                   tempC = tempC + 1
                               Next tempExposeType
                                   tempC = tempC + 1
                              strPrint = strPrint & FormatNumber(Total1(tempC), "0", 3) & "|"
                                   tempC = tempC + 1
                               strPrint = strPrint & "|" & FormatNumber(Total1(tempC), "0", 3) & "||||"
                               Call Vsp.AddTable(strFormat, "", strPrint)
                               Vsp.NewPage
                               
                               strPrint = ""
                              For j = 1 To UBound(Total1)
                                  Total1(j) = 0
                               Next j
                               
                               B = 1
                               End If
                               
                               If B = 0 Then
                                  B = 1
                               End If
                               
                              Set Ba = GetBalanceAccum(BalanceAccums, Trim(Str(NVLS(Rs("PART_ITEM_ID"), ""))))
                              tempBALANCE_AMOUNT = Ba.BALANCE_AMOUNT
                              Total1(1) = Ba.BALANCE_AMOUNT                         ' ยอดยกมา ต้องโชว์ก่อนเปลี่ยนค่า
                              tempPartName = NVLS(Rs("PART_DESC"), "")
                              TempExDate = DateAdd("D", -1, TempFromDate)
                           End If
                           
                           For Each subTempCl In tempCl.m_Import
                              If TempExDate < subTempCl.DOCUMENT_DATE And subTempCl.DOCUMENT_DATE <= InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), "")) Then
                                    strPrint = DateToStringExtEx2(subTempCl.DOCUMENT_DATE) & "|" & FormatNumber(tempBALANCE_AMOUNT, "0", 3) & "|" & FormatNumber(subTempCl.IMPORT_AMOUNT, "0", 3) & "|"
                                   For Each tempExposeType In m_ExposeType
                                      strPrint = strPrint & "|"
                                   Next tempExposeType
                                   strPrint = strPrint & "|"                          ' ตัดโอน
                                    strPrint = strPrint & "|"
                                    strPrint = strPrint & FormatNumber(subTempCl.IMPORT_AMOUNT + tempBALANCE_AMOUNT, "0", 3) & "|"          ' คงเหลือ
                                         tempBALANCE_AMOUNT = subTempCl.IMPORT_AMOUNT + tempBALANCE_AMOUNT
                                    strPrint = strPrint & "|" & "|" & "|"
                                    Total1(2) = Total1(2) + subTempCl.IMPORT_AMOUNT

                                        Call Vsp.AddTable(strFormat, "", strPrint)
                                        B = B + 1                                                                                    'นับบรรทัด

                                       strPrint = ""
                                    If B = 28 Then
                                       B = 1
                                       Vsp.NewPage
                                    End If

                                    TempImDate = subTempCl.DOCUMENT_DATE
                               End If  ' วันที่อยู่ระหว่าง
                           Next subTempCl
                  End If   ' เจอใน Coll
         End If      'มีการเปลี่ยนวัน
         
         If NVLS(Rs("PART_DESC"), "") <> tempPartName Then                                                                 ' เช็คบรรทัดแรกของวัตถุดิบตัวเดียวกัน
            If B <> 0 Then
               For tempB = B To 27 - 1             'ปริ้นให้ครบบรรทัดที่เหลือ
                  strPrint = "|||"
                  For Each tempExposeType In m_ExposeType
                     strPrint = strPrint & "|"
                  Next tempExposeType
                  strPrint = strPrint & "|"                                      ' ตัดโอน
                   strPrint = strPrint & "|||||"
                   Call Vsp.AddTable(strFormat, "", strPrint)
                   strPrint = ""
                   B = B + 1
               Next tempB
                               strPrint = "รวม|" & FormatNumber(Total1(1), "0", 3) & "|" & FormatNumber(Total1(2), "0", 3) & "|"
                               tempC = 4
                               For Each tempExposeType In m_ExposeType
                                  strPrint = strPrint & FormatNumber(Total1(tempC), "0", 3) & "|"
                                   tempC = tempC + 1
                               Next tempExposeType
'                                   strPrint = strPrint & FormatNumber(Total1(tempC), "0", 3) & "|"       ' ตัดโอน
                                   tempC = tempC + 1
                                   strPrint = strPrint & FormatNumber(Total1(tempC), "0", 3) & "|"
                                   tempC = tempC + 1
                               strPrint = strPrint & "|" & FormatNumber(Total1(tempC), "0", 3) & "||||"
               Call Vsp.AddTable(strFormat, "", strPrint)
               B = B + 1                                                                                    'นับบรรทัด
               Vsp.NewPage
               strPrint = ""
               
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j

              End If
              
            If B = 0 Then
               B = 1
            End If

               strPrint = ""
                strPrint = strPrint & DateToStringExtEx2(InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))) & "|"                               ' วันที่เอกสาร

               B = 1
               Set Ba = GetBalanceAccum(BalanceAccums, Trim(NVLS(Rs("PART_ITEM_ID"), "")))
               strPrint = strPrint & FormatNumber(Ba.BALANCE_AMOUNT, "0", 3) & "|"                                                                   ' ยอดยกมา
               tempBALANCE_AMOUNT = Ba.BALANCE_AMOUNT
               Total1(1) = Ba.BALANCE_AMOUNT                          ' ยอดยกมา ต้องโชว์ก่อนเปลี่ยน
               tempPartName = NVLS(Rs("PART_DESC"), "")
            Else
               strPrint = ""
               strPrint = strPrint & DateToStringExtEx2(InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))) & "|"                                ' วันที่เอกสา
               strPrint = strPrint & FormatNumber(tempBALANCE_AMOUNT, "0", 3) & "|"                                                                     ' ยอดยกมา
         
         End If

         strPrint = strPrint & "" & "|"                                                                                                              ' ยอดรับเข้า
         
         tempC = 4
         Set tempHgroup = GetObject("HGroupItem", m_HGroupItem, Trim(Str(NVLI(Rs("HOUSE_ID"), -1))))
         If Not (tempHgroup Is Nothing) Then
            For Each tempExposeType In m_ExposeType
               If tempHgroup.HOUSE_GROUP_ID = tempExposeType.HOUSE_GROUP_ID Then
                   
                   strPrint = strPrint & FormatNumber(NVLD(Rs("EXPORT_AMOUNT"), 0), "0", 3) & "|"
                  Total1(tempC) = Total1(tempC) + NVLD(Rs("EXPORT_AMOUNT"), 0)
                 
               Else
                    strPrint = strPrint & "|"
               End If
   
               tempC = tempC + 1
            Next tempExposeType
            strPrint = strPrint & "|"    '  ตัดโอน
         Else
              For Each tempExposeType In m_ExposeType
                    strPrint = strPrint & "|"
                    tempC = tempC + 1
            Next tempExposeType
             If Trim(NVLS(Rs("DOCUMENT_TYPE"), "")) = "10" Or Trim(NVLS(Rs("DOCUMENT_TYPE"), "")) = "13" Then                'ขายสด S

                strPrint = strPrint & FormatNumber(NVLD(Rs("EXPORT_AMOUNT"), 0), "0", 3) & "|"     ' ตัดโอน & "|"
                tempC = tempC + 1
                Total1(tempC) = Total1(tempC) + NVLD(Rs("EXPORT_AMOUNT"), 0)
                 tempC = tempC + 1
            ElseIf Trim(NVLS(Rs("DOCUMENT_TYPE"), "")) = "3" Then           'โอน B
                strPrint = strPrint & FormatNumber(NVLD(Rs("EXPORT_AMOUNT"), 0), "0", 3) & "|" & "|"
                 Total1(tempC) = Total1(tempC) + NVLD(Rs("EXPORT_AMOUNT"), 0)
                   tempC = tempC + 1
            Else
              strPrint = strPrint & "|"
            End If
         End If

       
         strPrint = strPrint & NVLS(Rs("SURETY_NAME"), "") & " " & NVLS(Rs("SURETY_LNAME"), "") & "|"                   ' ผูู้เบิก
       
         strPrint = strPrint & FormatNumber(tempBALANCE_AMOUNT - NVLD(Rs("EXPORT_AMOUNT"), 0), "0", 3) & "|"         ' คงเหลือ
                      tempBALANCE_AMOUNT = tempBALANCE_AMOUNT - NVLD(Rs("EXPORT_AMOUNT"), 0)
         strPrint = strPrint & "|"                                                              ' & DateToStringExtEx2(InternalDateToDate(NVLS(Rs("MFG_DATE"), "")))  วันที่ MFG
         strPrint = strPrint & "|"                                                                 ' & DateToStringExtEx2(InternalDateToDate(NVLS(Rs("EXP_DATE"), "")))วันที่ EXP
         
         
         strPrint = strPrint & NVLS(Rs("EMP_NAME"), "") & " " & NVLS(Rs("EMP_LNAME"), "") & "|"              'ผูู้รับผิดชอบ
'        ''debug.print (strPrint)
         Total1(tempC + 2) = tempBALANCE_AMOUNT
            
         tempPartName = NVLS(Rs("PART_DESC"), "")
         TempExDate = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
 
 If Trim(NVLS(Rs("DOCUMENT_TYPE"), "")) <> "3" Then          'ไม่ใช่ โอน B ไม่ต้องโชว์
       Call Vsp.AddTable(strFormat, "", strPrint)
       B = B + 1
End If
      strPrint = ""
        If B = 28 Then
            B = 1
            Vsp.NewPage
         End If
         If Not Rs.EOF Then
'            Set BodyCf = Nothing
'            genDoc = True
'            Exit Function
'         Else
            Rs.MoveNext
         End If
     Next I

      If B <> 0 Then
                              For tempB = B To 27 - 1           'ปริ้นให้ครบบรรทัดที่เหลือ
                                  strPrint = "|||"
                                  For Each tempExposeType In m_ExposeType
                                     strPrint = strPrint & "|"
                                  Next tempExposeType
                                  strPrint = strPrint & "|"             'ตัด โอน
                                   strPrint = strPrint & "|||||"
                                   Call Vsp.AddTable(strFormat, "", strPrint)
                                   B = B + 1
                                   strPrint = ""
                               Next tempB
                               strPrint = "รวม|" & FormatNumber(Total1(1), "0", 3) & "|" & FormatNumber(Total1(2), "0", 3) & "|"
                               tempC = 4
                               For Each tempExposeType In m_ExposeType
                                  strPrint = strPrint & FormatNumber(Total1(tempC), "0", 3) & "|"
                                   tempC = tempC + 1
                               Next tempExposeType
                                  tempC = tempC + 1
                                    strPrint = strPrint & FormatNumber(Total1(tempC), "0", 3) & "|"
                                   tempC = tempC + 1
                               strPrint = strPrint & "|" & FormatNumber(Total1(tempC), "0", 3) & "||||"
                              Call Vsp.AddTable(strFormat, "", strPrint)
                               strPrint = ""
                               B = B + 1
                               Vsp.EndDoc
                            End If
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
   
   Vsp = vbCrLf & "                                                ผู้จัดทำ _________________________________                                                                                  ผู้จัดการฟาร์ม  _________________________                                                                                  สัตวแพทย์ควบคุมฟาร์ม  ______________________" & _
            vbCrLf & "                                                       (                                                                       )                                                                                      (                                                                          )                                                                                             (                                                                          )" & _
            vbCrLf & "                                                                   ว/ด/ป   ___ / ___ / ___                                                                                                                               ว/ด/ป   ___ / ___ / ___                                                                                                                               ว/ด/ป   ___ / ___ / ___ "
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
   Vsp = "ชื่อยา/วัคซีน/สารเคมี  " & EmptyToString(NVLS(Rs("PART_DESC"), "") & " [" & NVLS(Rs("PART_NO") & "]", ""), "_________") & "                                                          เลขทะเบียน  " & EmptyToString(NVLS(Rs("REGIST_NUM"), ""), "______________") & "   ขนาดบรรจุ  " & EmptyToString(NVLS(Rs("DRUG_SIZE"), ""), "______________")
   Vsp = "ตัวยา  " & EmptyToString(NVLS(Rs("DRUG_NAME"), "_____________________________________________________________________________"), "") & "                                        บริษัทจัดจำหน่าย  " & NVLS(Rs("SUP_NAME"), "______________________") & " (" & NVLS(Rs("SUP_NUM"), "__________") & ")"
'   Vsp.Paragraph = ""
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




