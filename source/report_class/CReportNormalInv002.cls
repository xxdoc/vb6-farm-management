VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalInv002"
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

Private DocumentRunStart As String
Private DocumentRunNo  As Double
Private TempDocumentNo As String
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_cData = New CInventoryDoc
   Set m_Enterprise = New CEnterprise
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_cData = Nothing
   Set m_Enterprise = Nothing
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
   strPrint = "เลขที่เอกสาร : " & TempDocumentNo & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontSize = DETAIL_SIZE
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   strPrint = "ฟาร์ม/ฝ่าย| " & "_______________" & "|" & "วันที่ที่รับของ" & "|" & "_______________" & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   strPrint = "ผู้ขาย| " & "_______________" & "|" & "ใบสั่งซื้อเลขที่" & "|_______________" & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   strPrint = "เลขที่ใบส่งของ| " & "_______________" & "|" & "หน่วยงานที่ใช้" & "|_______________" & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<3|<10|<3|<4;", mdblWid, alngX)
   strPrint = "ทะเบียนรถ| " & "_______________" & "|" & "เครดิต" & "|_______________" & ";"
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
   
   DocumentRunStart = mcolParam("REPORT_DOC_NO")
   DocumentRunNo = mcolParam("REPORT_RUN_NO")

   
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
Dim i As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim Sum2 As Double
Dim TempCol As Collection
Dim Bi As CImportItem
Dim TempSupplier As CSupplier
Dim CountRun As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
   
   Call StartExportFile(Vsp)
                  
   Vsp.TableBorder = tbAll
   
   Vsp.PenStyle = psDash
   
   TempDocumentNo = DocumentRunStart
   Vsp.StartDoc
   mblnHeader = True
           
   i = 0
   
   
   For CountRun = 1 To DocumentRunNo
      If CountRun > 1 Then
         TempDocumentNo = Left(DocumentRunStart, Len(DocumentRunStart) - 3) & Format(Val(Right(DocumentRunStart, 3) + CountRun - 1), "000")
         Vsp.NewPage
      End If
             
      Vsp.TableBorder = tbAll
      Vsp.PenStyle = psDash
        
        i = 0
        strFormat = VSP_CalTable(DETAIL_RECEIPT1, mdblWid, alngX)
        i = i + 1
        strPrint = ""
        strPrint = strPrint & "|"
        strPrint = strPrint & "" & "|"
        strPrint = strPrint & "" & "|"
        strPrint = strPrint & "" & "|"
        strPrint = strPrint & ";"
        
        Call Vsp.AddTable(strFormat, "", strPrint)
    
        Call PatchTable(Vsp, strFormat, "||||;", 4000) '10
       
        strFormat = VSP_CalTable("<10|>4|>2;", mdblWid, alngX)
        Vsp.FontBold = True
        strPrint = "|ยอดรวมสุทธิ|;"
        Call Vsp.AddTable(strFormat, "", strPrint)
    
        Vsp.FontBold = False
        Call PrintBillFooter(alngX)
   Next CountRun
   
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
