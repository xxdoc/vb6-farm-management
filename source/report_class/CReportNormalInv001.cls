VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalInv001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalInv001"
Private MODULE_DESC  As String

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
Private Const HEADER_SIZE = 16
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<52|<3|<5;"

Private Const HEADER_RECEIPT1 = "^1|^9|^2|^2|^2;"

Private Const DETAIL_RECEIPT1 = "^1|<9|>2|>2|>2;"

Private m_cData As CInventoryDoc
Private m_Enterprise As CEnterprise
Private m_SupplierColl As Collection

Private DocumentRunStart As String
Private DocumentRunNo  As Double
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_cData = New CInventoryDoc
   Set m_Enterprise = New CEnterprise
   Set m_SupplierColl = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_cData = Nothing
   Set m_Enterprise = Nothing
   Set m_SupplierColl = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
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

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("2|<10;", mdblWid, alngX)
   strPrint = "ข้อความสำคัญ  :|1.โปรดตรวจสอบสินค้าให้ถูกต้องก่อนออกใบรายงานรับของทุกครั้ง;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("2|<10;", mdblWid, alngX)
   strPrint = "|2.เมื่อรับสินค้าแล้ว ให้ส่งใบรายงานรับของให้หน่วยงานที่เกี่ยวข้องทันที;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strPrint = strPrint & "..................................|..................................|.....................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strPrint = "ผู้รับของ|ผู้ตรวจสอบ|ผู้จัดการ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub

Private Sub PrintBillHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Dim iCount As Long
Dim OwnerFlag As String
Dim TempStr As String
Dim TempAddr As String
Dim Addr As CAddress
Dim EmpAddr As CEnterpriseAddress
Dim TempX As Long
Dim TempY As Long
Dim TempSupplier As CSupplier

   Set EmpAddr = m_Enterprise.EnpAddresses(1)
   Set Addr = EmpAddr.Addresses(1)
   
   Vsp.TableBorder = tbNone
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   
   strFormat = VSP_CalTable("^10;", mdblWid, alngX)
   strPrint = glbEnterPrise.ENTERPRISE_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = DETAIL_SIZE
   strFormat = VSP_CalTable("^10;", mdblWid, alngX)
   strPrint = MapText("") & Addr.PackAddress & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("^10;", mdblWid, alngX)
   strPrint = "โทร. " & EmptyToString(Addr.PHONE1, "N/A") & " แฟกซ์ " & EmptyToString(Addr.FAX1, "N/A") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strFormat = VSP_CalTable("^10;", mdblWid, alngX)
   strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & glbEnterPrise.TAX_ID & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   
   TempX = Vsp.CurrentX
   TempY = Vsp.CurrentY
   
   Vsp.FontSize = TITLE_SIZE
   strFormat = VSP_CalTable("^10;", mdblWid, alngX)
   strPrint = "ใบรายงานรับของ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.CurrentX = TempX
   Vsp.CurrentY = TempY
   strFormat = VSP_CalTable(">10;", mdblWid, alngX)
   If mcolParam("REPORT_TYPE") = 1 Then
      If Vsp.CurrentPage > 1 Then
        strPrint = "เลขที่เอกสาร : " & m_cData.DOCUMENT_NO & "-1" & ";"
      Else
        strPrint = "เลขที่เอกสาร : " & m_cData.DOCUMENT_NO & ";"
      End If
      Call Vsp.AddTable(strFormat, "", strPrint)
   Else
      strPrint = "หมายเลขเอกสาร : " & "_______________" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   End If
   
   Vsp.FontSize = DETAIL_SIZE
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   If mcolParam("REPORT_TYPE") = 1 Then
      strPrint = "ฟาร์ม/ฝ่าย| " & glbEnterPrise.BRANCH_NAME & "|" & "วันที่ที่รับของ" & "|" & DateToStringExtEx2(m_cData.DOCUMENT_DATE) & ";"
   Else
      strPrint = "ฟาร์ม/ฝ่าย| " & "_______________" & "|" & "วันที่ที่รับของ" & "|" & "_______________" & ";"
   End If
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   If mcolParam("REPORT_TYPE") = 1 Then
      If Vsp.CurrentPage > 1 Then
        Set TempSupplier = GetObject("CSupplier", m_SupplierColl, Trim(Str(m_cData.DELIVERY_ID)))
        strPrint = "ผู้ขนส่ง| " & Trim(TempSupplier.SUPPLIER_NAME) & "|" & "ใบสั่งซื้อเลขที่" & "|" & Trim(m_cData.PO_NO) & ";"
      Else
        strPrint = "ผู้ขาย| " & Trim(m_cData.SUPPLIER_NAME) & "|" & "ใบสั่งซื้อเลขที่" & "|" & Trim(m_cData.PO_NO) & ";"
      End If
   Else
      strPrint = "ผู้ขาย| " & "_______________" & "|" & "ใบสั่งซื้อเลขที่" & "|_______________" & ";"
   End If
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   If mcolParam("REPORT_TYPE") = 1 Then
      strPrint = "เลขที่ใบส่งของ| " & m_cData.DO_NO & "|" & "หน่วยงานที่ใช้" & "|" & ";"
   Else
      strPrint = "เลขที่ใบส่งของ| " & "_______________" & "|" & "หน่วยงานที่ใช้" & "|_______________" & ";"
   End If
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   If mcolParam("REPORT_TYPE") = 1 Then
      strPrint = "ทะเบียนรถ| " & m_cData.TRUCK_NO & "|" & "เครดิต" & "|" & ";"
   Else
      strPrint = "ทะเบียนรถ| " & "_______________" & "|" & "เครดิต" & "|_______________" & ";"
   End If
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontBold = False
End Sub

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
   
   Call PrintBillHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
      
   
   strFormat = VSP_CalTable(HEADER_RECEIPT1, mdblWid, alngX)
   strPrint = "No|รายการ|จำนวน|ราคาต่อหน่วย|จำนวนเงิน;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim ItemCount As Long
Dim IsOK As Boolean
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprLetter
   
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   If mcolParam("REPORT_TYPE") = 3 Then
      DocumentRunStart = mcolParam("REPORT_DOC_NO")
      DocumentRunNo = mcolParam("REPORT_RUN_NO")
   End If
   
   Dim m_Rs As ADODB.Recordset
   Set m_Rs = New ADODB.Recordset
   
   m_Enterprise.ENTERPRISE_ID = glbEnterPrise.ENTERPRISE_ID
   m_Enterprise.QueryFlag = 1
   If Not glbDaily.QueryEnterprise(m_Enterprise, m_Rs, ItemCount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   Set m_Rs = Nothing
   
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
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim Sum2 As Double
Dim TempCol As Collection
Dim Bi As CImportItem
Dim TempSupplier As CSupplier

   RName = "genDoc"
   Set Rs = New ADODB.Recordset
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
   
   Call StartExportFile(Vsp)
   
   Set Rs = New ADODB.Recordset
   m_cData.INVENTORY_DOC_ID = mcolParam("INVENTORY_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryInventoryDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
      
   Call m_cData.PopulateFromRS(1, Rs)
               
   Vsp.TableBorder = tbAll
   
   Vsp.PenStyle = psDash
   
   
   Vsp.StartDoc
   mblnHeader = True
           
   I = 0
   
   strFormat = VSP_CalTable(DETAIL_RECEIPT1, mdblWid, alngX)
   Set TempCol = m_cData.ImportExports
   For Each Bi In TempCol
      I = I + 1
      strPrint = ""
      If mcolParam("REPORT_TYPE") = 1 Then
         strPrint = strPrint & I & "|"
         If Len(Bi.IMPORT_DESC) > 0 Then
            strPrint = strPrint & Bi.PART_DESC & "(" & Bi.IMPORT_DESC & ")" & "|"
         Else
            strPrint = strPrint & Bi.PART_DESC & "|"
         End If
         strPrint = strPrint & FormatNumber(Bi.IMPORT_AMOUNT) & " " & Bi.UNIT_NAME & "|"
         strPrint = strPrint & FormatNumber(Bi.ACTUAL_UNIT_PRICE) & "|"
         strPrint = strPrint & FormatNumber(Bi.TOTAL_ACTUAL_PRICE) & ";"
         
         Sum1 = Sum1 + Bi.TOTAL_ACTUAL_PRICE
      Else
         strPrint = "||||;"
      End If
      If I >= 13 Then
         Call AdjustPage(Vsp, strFormat, strPrint, 5000)
      End If
      Call Vsp.AddTable(strFormat, "", strPrint)
   Next Bi
   
   If mcolParam("REPORT_TYPE") = 1 And m_cData.DELIVERY_ID <= 0 Then
      If m_cData.DELIVERY_FEE > 0 Then
         I = I + 1
         strPrint = ""
         strPrint = strPrint & I & "|"
         strPrint = strPrint & "ค่าขนส่ง" & "|"
         strPrint = strPrint & "" & "|"
         strPrint = strPrint & "" & "|"
         strPrint = strPrint & FormatNumber(m_cData.DELIVERY_FEE) & ";"
         
         Sum1 = Sum1 + m_cData.DELIVERY_FEE
         Call Vsp.AddTable(strFormat, "", strPrint)
      End If
   End If
      
   Call PatchTable(Vsp, strFormat, "||||;", 4000) '10
      
   If mcolParam("REPORT_TYPE") = 1 Then
      strFormat = VSP_CalTable("<10|>4|>2;", mdblWid, alngX)
      Vsp.FontBold = True
      strPrint = " (" & ThaiBaht(Sum1) & ")|ยอดรวมสุทธิ|" & FormatNumber(Sum1) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   Else
      strFormat = VSP_CalTable("<10|>4|>2;", mdblWid, alngX)
      Vsp.FontBold = True
      strPrint = "|ยอดรวมสุทธิ|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   End If
   
   Vsp.FontBold = False
   Call PrintBillFooter(alngX)
         
       
    If mcolParam("REPORT_TYPE") = 1 And m_cData.DELIVERY_ID > 0 And m_cData.DELIVERY_FEE > 0 Then
        Call LoadSupplier(Nothing, m_SupplierColl)
        Vsp.NewPage
             
        Vsp.TableBorder = tbAll
        Vsp.PenStyle = psDash
        
        I = 0
        strFormat = VSP_CalTable(DETAIL_RECEIPT1, mdblWid, alngX)
        I = I + 1
        strPrint = ""
        strPrint = strPrint & I & "|"
        strPrint = strPrint & "ค่าขนส่ง" & "|"
        strPrint = strPrint & "" & "|"
        strPrint = strPrint & "" & "|"
        strPrint = strPrint & FormatNumber(m_cData.DELIVERY_FEE) & ";"
        
        Sum1 = m_cData.DELIVERY_FEE
        Call Vsp.AddTable(strFormat, "", strPrint)
    
        Call PatchTable(Vsp, strFormat, "||||;", 4000) '10
       
        strFormat = VSP_CalTable("<10|>4|>2;", mdblWid, alngX)
        Vsp.FontBold = True
        strPrint = " (" & ThaiBaht(Sum1) & ")|ยอดรวมสุทธิ|" & FormatNumber(Sum1) & ";"
        Call Vsp.AddTable(strFormat, "", strPrint)
    
        Vsp.FontBold = False
        Call PrintBillFooter(alngX)
    End If
       
       
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   genDoc = True
   
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
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
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   'Vsp.Paragraph = MODULE_DESC
   'Vsp.Paragraph = ""
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
