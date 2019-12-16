VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportRC001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportRC001"
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

Private Const TITLE_SIZE = 18
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<32|<3|<5;"
'NO.|วันที่|เลขที่เอกสาร|เกิด|โอนเข้า|%P|โอนออก|คงเหลือ;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Left As Double
Private m_YearWeeks As Collection

Private m_BillingDoc As CBillingDoc
Private m_BillingAddress As CAddress
Private m_EnpAddress As CAddress
Private m_Customer As CCustomer
Private m_Account As CAccount
Private m_Rs As ADODB.Recordset

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_YearWeeks = New Collection
   
   Set m_BillingDoc = New CBillingDoc
   Set m_BillingAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_Customer = New CCustomer
   Set m_Account = New CAccount
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_BillingAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_Customer = Nothing
   Set m_Account = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing

   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
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
Dim Cf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = glbEnterPrise.ENTERPRISE_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("ที่อยู่ ") & m_EnpAddress.PackAddress & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("โทรศัพท์ ") & m_EnpAddress.PHONE1 & MapText(" แฟกซ์ ") & EmptyToString(m_EnpAddress.FAX1, "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strFormat = VSP_CalTable("<10|>10;", mdblWid, alngX)
   strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & glbEnterPrise.TAX_ID & "| " & MapText("เลขที่เอกสาร ") & m_BillingDoc.DOCUMENT_NO & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<15|>10;", mdblWid, alngX)
   strPrint = MapText("วันที่เอกสาร ") & EmptyToString(DateToStringExtEx2((m_BillingDoc.DOCUMENT_DATE)), "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10|>10;", mdblWid, alngX)
   strPrint = MapText("ชื่อลูกค้า ") & m_Customer.CUSTOMER_NAME & "|" & MapText("เลขที่บัญชีลูกค้า ") & m_Account.ACCOUNT_NO & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("ที่อยู่ลูกค้า ") & EmptyToString(m_BillingAddress.PackAddress, "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("โทรศัพท์ ") & EmptyToString(m_BillingAddress.PHONE1, "N/A") & MapText(" แฟกซ์ ") & EmptyToString(m_BillingAddress.FAX1, "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbAll
   
   Call m_Headers1.ClearField
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "ลำดับ", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(7, "^", "รายการ", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวน (ตัว)", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก (ก.ก.)", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ราคา/ก.ก.", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
   Call m_Headers1.AddField(Cf)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = False
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim iCount As Long
Dim IsOK As Boolean

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
   
   MODULE_DESC = mcolParam("REPORT_HEADER")
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   m_BillingDoc.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_BillingDoc.QueryFlag = 1
   Call glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, iCount, IsOK, glbErrorLog)
   If Not m_Rs.EOF Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)
   End If
   
   m_BillingAddress.ADDRESS_ID = m_BillingDoc.BILLING_ADDRESS_ID
   Call m_BillingAddress.QueryData(m_Rs, iCount)
   If Not m_Rs.EOF Then
      Call m_BillingAddress.PopulateFromRS(m_Rs)
   End If
   
   m_EnpAddress.ADDRESS_ID = m_BillingDoc.ENTERPRISE_ADDRESS_ID
   Call m_EnpAddress.QueryData(m_Rs, iCount)
   If Not m_Rs.EOF Then
      Call m_EnpAddress.PopulateFromRS(m_Rs)
   End If
   
   m_Account.ACCOUNT_ID = m_BillingDoc.ACCOUNT_ID
   Call m_Account.QueryData(1, m_Rs, iCount)
   If Not m_Rs.EOF Then
      Call m_Account.PopulateFromRS(1, m_Rs)
   End If
   
   m_Customer.CUSTOMER_ID = m_Account.CUSTOMER_ID
   Call m_Customer.QueryData2(m_Rs, iCount)
   If Not m_Rs.EOF Then
      Call m_Customer.PopulateFromRS(1, m_Rs)
   End If
   
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

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim Cf As CReportField
Dim I As Long
Dim Di As CDoItem
Dim Ri As CReceiptItem
Dim strPrint As String
Dim strFormat As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim alngX() As Long
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Bd As CBillingDoc
Dim iCount As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
      
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
   I = 0
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   
'   Vsp.FontBold = False
'   Vsp.TableBorder = tbAll
   
   Set Bd = New CBillingDoc
   If m_BillingDoc.RECEIPT_TYPE = 1 Then
      For Each Di In m_BillingDoc.DoItems
          I = I + 1
          Call m_Details1.ClearField
          
          Set Cf = m_Headers1.Fields(1)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, Trim(Str(I)), Cf.BodyAlign)
          Call m_Details1.AddField(Cf)
       
          Set Cf = m_Headers1.Fields(2)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, Di.PART_DESC & " (" & Di.PIG_STATUS_NAME & ")", Cf.BodyAlign)
          Call m_Details1.AddField(Cf)
       
          Set Cf = m_Headers1.Fields(3)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Di.ITEM_AMOUNT), Cf.BodyAlign)
          Sum1 = Sum1 + Di.ITEM_AMOUNT
          Call m_Details1.AddField(Cf)
       
          Set Cf = m_Headers1.Fields(4)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Di.TOTAL_WEIGHT), Cf.BodyAlign)
          Sum2 = Sum2 + Di.TOTAL_WEIGHT
          Call m_Details1.AddField(Cf)
          
          Set Cf = m_Headers1.Fields(5)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(MyDiff(Di.TOTAL_PRICE, Di.TOTAL_WEIGHT)), Cf.BodyAlign)
          Call m_Details1.AddField(Cf)
       
          Set Cf = m_Headers1.Fields(6)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Di.TOTAL_PRICE), Cf.BodyAlign)
          Sum4 = Sum4 + Di.TOTAL_PRICE
          Call m_Details1.AddField(Cf)
          
          Call m_Details1.GetString(1, TempStr1, TempStr2)
          
           strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
           strPrint = TempStr2
           
          If I >= m_BillingDoc.DoItems.Count Then
             Call AdjustPage(Vsp, strFormat, strPrint, 5000)
          End If
          Call Vsp.AddTable(strFormat, "", strPrint)
       Next Di
   ElseIf m_BillingDoc.RECEIPT_TYPE = 3 Then
      For Each Ri In m_BillingDoc.ReceiptItems
         Bd.BILLING_DOC_ID = Ri.DO_ID
         Call Bd.QueryData(1, m_Rs, iCount)
         If Not m_Rs.EOF Then
            Call Bd.PopulateFromRS(1, m_Rs)
         End If
         
          I = I + 1
          Call m_Details1.ClearField
          
          Set Cf = m_Headers1.Fields(1)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, Trim(Str(I)), Cf.BodyAlign)
          Call m_Details1.AddField(Cf)
       
          Set Cf = m_Headers1.Fields(2)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, "อ้างถึงใบส่งสินค้า " & Ri.DOCUMENT_NO, Cf.BodyAlign)
          Call m_Details1.AddField(Cf)
       
          Set Cf = m_Headers1.Fields(3)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Bd.DO_TOTAL_AMOUNT), Cf.BodyAlign)
          Sum1 = Sum1 + Bd.DO_TOTAL_AMOUNT
          Call m_Details1.AddField(Cf)
        
          Set Cf = m_Headers1.Fields(4)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Bd.TOTAL_SUM_WEIGHT), Cf.BodyAlign)
          Sum2 = Sum2 + Bd.TOTAL_SUM_WEIGHT
          Call m_Details1.AddField(Cf)
          
          Set Cf = m_Headers1.Fields(5)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(MyDiff(Bd.DO_TOTAL_PRICE, Bd.TOTAL_SUM_WEIGHT)), Cf.BodyAlign)
          Call m_Details1.AddField(Cf)
       
          Set Cf = m_Headers1.Fields(6)
          Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Bd.DO_TOTAL_PRICE), Cf.BodyAlign)
          Sum4 = Sum4 + Bd.DO_TOTAL_PRICE
          Call m_Details1.AddField(Cf)
          
          Call m_Details1.GetString(1, TempStr1, TempStr2)
          
           strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
           strPrint = TempStr2
           
          If I >= m_BillingDoc.DoItems.Count Then
             Call AdjustPage(Vsp, strFormat, strPrint, 5000)
          End If
          Call Vsp.AddTable(strFormat, "", strPrint)
      Next Ri
   End If
   Set Bd = Nothing
   
   Call PatchTable(Vsp, strFormat, "|||||", 5000)
   '===
   Call m_Details1.ClearField
   
   Set Cf = m_Headers1.Fields(1)
   Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, "รวม", Cf.BodyAlign)
   Call m_Details1.AddField(Cf)

   Set Cf = m_Headers1.Fields(2)
   Call Cf.SetFieldValue(Cf.Width, ">", "(" & ThaiBaht(Sum4) & ")", ">")
   Call m_Details1.AddField(Cf)

   Set Cf = m_Headers1.Fields(3)
   Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Sum1), Cf.BodyAlign)
   Call m_Details1.AddField(Cf)

   Set Cf = m_Headers1.Fields(4)
   Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Sum2), Cf.BodyAlign)
   Call m_Details1.AddField(Cf)
   
   Set Cf = m_Headers1.Fields(5)
   Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, "", Cf.BodyAlign)
   Call m_Details1.AddField(Cf)

   Set Cf = m_Headers1.Fields(6)
   Call Cf.SetFieldValue(Cf.Width, Cf.BodyAlign, FormatNumber(Sum4), Cf.BodyAlign)
   Call m_Details1.AddField(Cf)
   
   Call m_Details1.GetString(1, TempStr1, TempStr2)
    strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
    strPrint = TempStr2
    
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = False
   '===
   
   Call PrintBillFooter(alngX)
   
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = "ผู้รับสินค้า ............................................|ผู้ส่งสินค้า ............................................|ผู้รับเงิน ............................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = "     วันที่ ............................................|     วันที่ ............................................|     วันที่ ............................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
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


