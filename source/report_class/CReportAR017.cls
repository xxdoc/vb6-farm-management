VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportAR017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportAR012"
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
Private m_ReceiptItems As Collection
Private m_MaxMinDates As Collection
Private m_CashTrans As Collection
Private m_ToDate As Date

Private Sub Class_Initialize()
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
   Set m_YearSeqs = New Collection
   Set m_ReceiptItems = New Collection
   Set m_CashTrans = New Collection
   Set m_MaxMinDates = New Collection
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
   Set m_Details2 = Nothing
   Set m_ReceiptItems = Nothing
   Set m_CashTrans = Nothing
   Set m_MaxMinDates = Nothing
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
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Inv As CInventoryDoc
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long
Dim TmpDate As Date
Dim I As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
      
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ชื่อลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ใบส่งของ", "^")
   Call m_Headers1.AddField(Cf)
   
   TmpDate = mcolParam("FROM_DATE")
   For I = 1 To 7
      Call Cf.SetFieldValue(2, "^", DateToStringExtEx2(TmpDate), ">", DateToStringExtEx2(TmpDate))
      Call m_Headers1.AddField(Cf)
      TmpDate = DateAdd("D", 1, TmpDate)
   Next I
   
   Call Cf.SetFieldValue(3, "^", "เช็ครับล่วงหน้า", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เช็ค", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "พ.ขาย", "^")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
   Set TempRs = Nothing
   
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
Dim Ri As CReceiptItem

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   m_ToDate = DateAdd("D", 6, mcolParam("FROM_DATE"))
   Call LoadDoIDFromReceiptItem(Nothing, m_ReceiptItems, mcolParam("FROM_DATE"), m_ToDate)
   Call LoadMaxMinReceiptDate(Nothing, m_MaxMinDates, mcolParam("FROM_DATE"), m_ToDate)
   Set Ri = m_MaxMinDates(1)
   Call LoadChequeFromReceipt(Nothing, m_CashTrans, Ri.MIN_DOC_DATE, Ri.MAX_DOC_DATE)
   
   FromDate = "จาก วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(m_ToDate), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function MyGetItem(Key As String, Col As Collection) As CReceiptItem
On Error Resume Next
Dim Ri As CReceiptItem
   
   Set Ri = Col(Key)
   Set MyGetItem = Ri
End Function

Private Function MyGetCashTran(Key As String, Col As Collection, CheckNo As String, CheckDate As String) As Boolean
Dim Ct As CCashTran
Dim TempStr1 As String
Dim TempStr2 As String
Dim Found As Boolean

   TempStr1 = ""
   TempStr2 = ""
   Found = False
   For Each Ct In Col
      If (Ct.GetFieldValue("BILLING_DOC_ID") = Key) And (Ct.GetFieldValue("POST_FLAG") = "N") Then
         TempStr1 = TempStr1 & " " & Ct.GetFieldValue("CHEQUE_NO")
         TempStr2 = TempStr2 & " " & DateToStringExtEx2(Ct.GetFieldValue("EFFECTIVE_DATE"))
         Found = True
      End If
   Next Ct
   
   CheckNo = TempStr1
   CheckDate = TempStr2
   MyGetCashTran = Found
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim j As Long
Dim k As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Cm As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim TempDiscount As Double
Dim Ri As CReceiptItem
Dim Ct As CCashTran
Dim Diff As Double
Dim CheckNo As String
Dim CheckDate As String
Dim X As Double
Dim ItemCount As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
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
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField

      Set Cm = New CBillingDoc
      Cm.BILLING_DOC_ID = -1
      Cm.FROM_DUE_DATE = mcolParam("FROM_DATE")
      Cm.TO_DUE_DATE = m_ToDate
      Cm.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Cm.OrderBy = 10
      Cm.OrderType = 1
      Cm.DOCUMENT_TYPE = 1
      Cm.ItemSumFlag = True
      Call Cm.QueryData(1, Rs, iCount)

      I = 0

      If Not Rs.EOF Then
         Call Cm.PopulateFromRS(1, Rs)
         PrevKey1 = Cm.CUSTOMER_CODE
      End If
      
      While Not Rs.EOF
         Call Cm.PopulateFromRS(1, Rs)
         X = (Cm.DO_TOTAL_PRICE + Cm.REVENUE_TOTAL_PRICE - Cm.DISCOUNT_AMOUNT)
         X = X + (Cm.CNDN_TOTAL_PRICE) - Cm.RECEIPT_PAID_AMOUNT '- Cm.CASH_DISCOUNT
         'X เป็นค่าเดียวกันกับที่ดึงมาแสดงตอนเพิ่มรายการใบลดหนี้เพิ่มหน้า frmAddDebitCreditNoteItem
         
        
         If PrevKey1 <> Cm.CUSTOMER_CODE Then
            If ItemCount > 0 Then
               Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            
            PrevKey1 = Cm.CUSTOMER_CODE
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            ItemCount = 0
         End If
         
         Call m_Details1.ClearField
         Set Ri = MyGetItem(Cm.BILLING_DOC_ID & "-" & 1, m_ReceiptItems)
         If Ri Is Nothing Then
            I = I + 1
            'ยังไม่ถูกอ้างในใบเสร็จ
                        
            Cm.TEMP_PAID_AMOUNT = X
            Cm.TEMP_CHECK_NO = ""
            Cm.TEMP_CHECK_DATE = ""
            
            Call GenerateDetail(Cm, Total1, Total2, m_Details1, I)
            ItemCount = ItemCount + 1
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         Else
            k = 1
'If Ri.BILLING_DOC_ID = 16668 Then
'''debug.print
'End If
            'มีเช็คที่ยังไม่เคลียร์
            Set Ri = MyGetItem(Cm.BILLING_DOC_ID & "-" & k, m_ReceiptItems)
            While Not (Ri Is Nothing)
               If MyGetCashTran(Ri.BILLING_DOC_ID, m_CashTrans, CheckNo, CheckDate) Then
                  I = I + 1
                  Cm.TEMP_PAID_AMOUNT = Ri.PAID_AMOUNT
                  Cm.TEMP_CHECK_NO = CheckNo
                  Cm.TEMP_CHECK_DATE = CheckDate
                  
                  Call GenerateDetail(Cm, Total1, Total2, m_Details1, I)
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  ItemCount = ItemCount + 1
               End If
               
               k = k + 1
               Set Ri = MyGetItem(Cm.BILLING_DOC_ID & "-" & k, m_ReceiptItems)
            Wend
            
            'มาถึงจุดนี้อาจจะมีใบส่งของที่มีหนี้เหลือจากการรับชำระไม่เต็มบิล
            If X <> 0 Then
               Cm.TEMP_PAID_AMOUNT = X
               Cm.TEMP_CHECK_NO = ""
               Cm.TEMP_CHECK_DATE = ""
               I = I + 1
               
               Call GenerateDetail(Cm, Total1, Total2, m_Details1, I)
               ItemCount = ItemCount + 1
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         End If
         
         Rs.MoveNext
      Wend

      If ItemCount > 0 Then
         Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      
      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf (j >= 5) And (j <= 11) Then
         TempStr = FormatNumber(Tot(j))
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateDetail(Bd As CBillingDoc, Tot1() As Double, Tot2() As Double, Details As CFieldList, I As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = I
      ElseIf j = 2 Then
         TempStr = Bd.CUSTOMER_CODE
      ElseIf j = 3 Then
         TempStr = Bd.CUSTOMER_NAME
      ElseIf j = 4 Then
         TempStr = Bd.DOCUMENT_NO
      ElseIf (j >= 5) And (j <= 11) Then
         If HeadCf.TempKey = DateToStringExtEx2(Bd.DUE_DATE) Then
            TempStr = FormatNumber(Bd.TEMP_PAID_AMOUNT)
            Tot1(j) = Tot1(j) + Bd.TEMP_PAID_AMOUNT
            Tot2(j) = Tot2(j) + Bd.TEMP_PAID_AMOUNT
         Else
            TempStr = ""
         End If
      ElseIf j = 12 Then
         TempStr = Bd.TEMP_CHECK_DATE
      ElseIf j = 13 Then
         TempStr = Bd.TEMP_CHECK_NO
      ElseIf j = 14 Then
         TempStr = Bd.EMP_CODE
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
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


