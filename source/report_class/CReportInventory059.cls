VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory059"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory059"
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

Private m_SumAmountByBills As Collection
Private m_SumAmountByBillDocTypes As Collection

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
   Set m_SumAmountByBills = New Collection
   Set m_SumAmountByBillDocTypes = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_SumAmountByBills = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_SumAmountByBillDocTypes = Nothing
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

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   If mcolParam("SUMMARY_FLAG") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Else
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
   End If
   Call m_Headers0.AddField(Cf)
   
   If mcolParam("SUMMARY_FLAG") = "N" Then
      Call Cf.SetFieldValue(6, "^", "รายละเอียด", "^")
   Else
      Call Cf.SetFieldValue(8, "^", "รายละเอียด", "^")
   End If
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "ขายสุกร", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ขายวัตถุดิบ", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ขายอื่น ๆ", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "รวม", "^")
   Call m_Headers0.AddField(Cf)
   
   If mcolParam("SUMMARY_FLAG") = "N" Then
      Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Else
      Call Cf.SetFieldValue(0, "^", "NO.", "^")
   End If
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SUMMARY_FLAG") = "N" Then
      Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "รหัสลูกค้า", "^")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(6, "^", "ชื่อลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
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
   
   FromDate = "จาก วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
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

Private Function GetKey(Di As CBillingDoc) As String
Dim OrderBy As Long
   
   If OrderBy = 1 Then
      GetKey = Di.CUSTOMER_ID
   ElseIf OrderBy = 2 Then
      GetKey = Di.CUSTOMER_ID
   Else
      GetKey = Di.CUSTOMER_ID
   End If
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Bd As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String

Dim Di1 As CDoItem
Dim Di2 As CDoItem
Dim Di3 As CDoItem
Dim SumLine1 As Double
Dim SumLine2 As Double
Dim SumLine3 As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   Call LoadSumAmountByBill(Nothing, m_SumAmountByBills, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("PIG_STATUS"))
   Call LoadSumAmountByBillDocTypeSubType(Nothing, m_SumAmountByBillDocTypes, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("PIG_STATUS"))
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j

   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField

      Set Bd = New CBillingDoc
      Bd.BILLING_DOC_ID = -1
      Bd.FROM_DATE = mcolParam("FROM_DATE")
      Bd.TO_DATE = mcolParam("TO_DATE")
      Bd.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
      Bd.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
      Bd.CUSTOMER_CODE = mcolParam("CUSTOMER_CODE")
      Bd.FROM_EMP_CODE = mcolParam("FROM_SALE_CODE")
      Bd.TO_EMP_CODE = mcolParam("TO_SALE_CODE")
      Bd.DocSubTypeSet = "(1, 2) " 'หมู, วัตถุดิบ
      Bd.OrderBy = mcolParam("ORDER_BY")
      Bd.OrderType = mcolParam("ORDER_TYPE")
      If Bd.OrderBy <= 0 Then
         Bd.OrderBy = 8
      End If
      If Bd.OrderType <= 0 Then
         Bd.OrderType = 1
      End If
      
      Call Bd.QueryData(1, Rs, iCount)
      I = 0

      PrevKey1 = ""
      If Not Rs.EOF Then
         Call Bd.PopulateFromRS(1, Rs)
         PrevKey1 = GetKey(Bd)
         
         If mcolParam("SUMMARY_FLAG") = "N" Then
            Call GenerateHeader(Bd.CUSTOMER_CODE, Bd.CUSTOMER_NAME, "", m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
      End If
'
      Set Bd = Nothing
      While Not Rs.EOF
         I = I + 1

         Set Bd = New CBillingDoc

         Call Bd.PopulateFromRS(1, Rs)
         If PrevKey1 <> GetKey(Bd) Then
            Call GenerateFooter("รวม", PrevKey2, PrevKey3, Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            
            If mcolParam("SUMMARY_FLAG") = "N" Then
               Vsp.FontBold = True
            End If
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            If mcolParam("SUMMARY_FLAG") = "N" Then
               Call GenerateHeader(Bd.CUSTOMER_CODE, Bd.CUSTOMER_NAME, "", m_Details1)
               Call m_Details1.GetString(2, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
         End If
         PrevKey1 = GetKey(Bd)
         
         PrevKey2 = Bd.CUSTOMER_CODE
         If mcolParam("SUMMARY_FLAG") = "Y" Then
            PrevKey3 = Bd.CUSTOMER_NAME
         End If
         
         Call m_Details1.ClearField

         Set Di1 = GetDoItem(m_SumAmountByBillDocTypes, Bd.BILLING_DOC_ID & "-" & Bd.DOCUMENT_TYPE & "-" & 1)
         Set Di2 = GetDoItem(m_SumAmountByBillDocTypes, Bd.BILLING_DOC_ID & "-" & Bd.DOCUMENT_TYPE & "-" & 2)
         Set Di3 = GetDoItem(m_SumAmountByBills, Trim(Str(Bd.BILLING_DOC_ID)))

         SumLine1 = 0
         SumLine2 = 0
         SumLine3 = 0
         
         j = 0
         For Each HeadCf In m_Headers1.Fields
            j = j + 1
            If j = 1 Then
               TempStr = I
               Amt = 0
            ElseIf j = 2 Then
               TempStr = DateToStringExtEx2(Bd.DOCUMENT_DATE)
               Amt = 0
            ElseIf j = 3 Then
               TempStr = Bd.DOCUMENT_NO
               Amt = 0
            ElseIf j = 4 Then
               Amt = Di1.ITEM_AMOUNT
               TempStr = FormatNumber(Amt)
               SumLine1 = SumLine1 + Di1.ITEM_AMOUNT
            ElseIf j = 5 Then
               Amt = Di1.TOTAL_WEIGHT
               TempStr = FormatNumber(Amt)
            ElseIf j = 6 Then
               Amt = MyDiffEx(Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT, Di1.ITEM_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 7 Then
               SumLine2 = SumLine2 + Round(Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT, 2)
               Amt = Round(Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT, 2)
               TempStr = FormatNumber(Amt)
            ElseIf j = 8 Then
               SumLine1 = SumLine1 + Di2.ITEM_AMOUNT
               Amt = Di2.ITEM_AMOUNT
               TempStr = FormatNumber(Amt)
            ElseIf j = 9 Then
               Amt = MyDiffEx(Di2.TOTAL_PRICE, Di2.ITEM_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 10 Then
               Amt = Round(Di2.TOTAL_PRICE, 2)
               SumLine2 = SumLine2 + Round(Di2.TOTAL_PRICE, 2)
               TempStr = FormatNumber(Amt)
            ElseIf j = 11 Then
               Amt = Di3.ITEM_AMOUNT
               SumLine1 = SumLine1 + Di3.ITEM_AMOUNT
               TempStr = FormatNumber(Amt)
            ElseIf j = 12 Then
               Amt = MyDiffEx(Di3.TOTAL_PRICE, Di3.ITEM_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 13 Then
               Amt = Round(Di3.TOTAL_PRICE, 2)
               SumLine2 = SumLine2 + Round(Di3.TOTAL_PRICE, 2)
               TempStr = FormatNumber(Amt)
            ElseIf j = 14 Then
               Amt = SumLine1
               TempStr = FormatNumber(Amt)
            ElseIf j = 15 Then
               Amt = MyDiffEx(SumLine2, SumLine1)
               TempStr = FormatNumber(Amt)
            ElseIf j = 16 Then
               Amt = SumLine2
               TempStr = FormatNumber(Amt)
            End If

            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt

            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         Next HeadCf

         If mcolParam("SUMMARY_FLAG") = "N" Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         Rs.MoveNext
         Set Bd = Nothing
      Wend

      Call GenerateFooter("รวม", PrevKey2, PrevKey3, Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      If mcolParam("SUMMARY_FLAG") = "N" Then
         Vsp.FontBold = True
      End If
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 1)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
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
      ElseIf j = 6 Then
         TempStr = FormatNumber(MyDiffEx(Tot(7), Tot(4)))
      ElseIf j = 9 Then
         TempStr = FormatNumber(MyDiffEx(Tot(10), Tot(8)))
      ElseIf j = 12 Then
         TempStr = FormatNumber(MyDiffEx(Tot(13), Tot(11)))
      ElseIf j = 15 Then
         TempStr = FormatNumber(MyDiffEx(Tot(16), Tot(14)))
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateHeader(Txt1 As String, Txt2 As String, Txt3 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double
Dim TempAlign As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   For j = 1 To 2
      If j = 1 Then
         TempStr = Txt1
         TempWidth = 4
         TempAlign = "^"
      ElseIf j = 2 Then
         TempStr = Txt2
         TempWidth = 29
         TempAlign = "<"
      End If
      
      Call BodyCf.SetFieldValue(TempWidth, TempAlign, TempStr, TempAlign)
      Call Details.AddField(BodyCf)
   Next j
         
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




