VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory004"
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
Private Const HEADER_FORMAT1 = "^1|^3|^5|^17;"
Private Const HEADER_FORMAT2 = "^1|^3|^5|^3|^2|^3|^3|^3|^3;"
Private Const DETAIL_FORMAT1 = "^1|^3|^5|<17;"
Private Const DETAIL_FORMAT2 = "^1|<3|<5|>3|^2|>3|>3|<3|<3;"
Private Const FOOTER_FORMAT = "^1|^3|^5|>3|^2|>3|>3|<3|<3;"
'ลำดับ|ประเภทวัตถุดิบ|วัตถุดิบ|จำนวน|หน่วย|ราคา/หน่วย (รวมค่าขนส่ง)|ราคารวม|จากคลัง|ไปคลัง;
'ลำดับ|บิลรับของ|วันที่|ซัพพลายเออร์;

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
   
   strFormat = VSP_CalTable(HEADER_FORMAT1, mdblWid, alngX)
   strPrint = "ลำดับ|บิลรับของ|วันที่|ซัพพลายเออร์;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable(HEADER_FORMAT2, mdblWid, alngX)
   strPrint = "|ประเภทวัตถุดิบ|วัตถุดิบ|จำนวน|หน่วย|ราคา/หน่วย (รวม)|ราคารวม|จากคลัง|ไปคลัง;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
   Dim strDate As String
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
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME")
   
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
Dim cData As CInventoryDoc
Dim iCount As Long
Dim i As Long
Dim iCount2 As Long
Dim TempRs As ADODB.Recordset
Dim strFormat1 As String
Dim strFormat2 As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim II As CTransferItem

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   
   Set cData = New CInventoryDoc
   cData.INVENTORY_DOC_ID = -1
   cData.FROM_DATE = mcolParam("FROM_DATE")
   cData.TO_DATE = mcolParam("TO_DATE")
   cData.DOCUMENT_NO = mcolParam("DOCUMENT_NO")
   cData.DOCUMENT_TYPE = 3
   cData.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
   cData.OrderBy = mcolParam("ORDER_BY")
   cData.OrderType = mcolParam("ORDER_TYPE")
   If Not glbDaily.QueryInventoryDoc(cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
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
   'ลำดับ|ประเภทวัตถุดิบ|วัตถุดิบ|จำนวน|หน่วย|ราคา/หน่วย (รวมค่าขนส่ง)|ราคารวม|จากคลัง|ไปคลัง;
   'ลำดับ|บิลรับของ|วันที่|ซัพพลายเออร์;

   Call StartExportFile(Vsp)
   Vsp.StartDoc
      strFormat1 = VSP_CalTable(DETAIL_FORMAT1, mdblWid, alngX)
      strFormat2 = VSP_CalTable(DETAIL_FORMAT2, mdblWid, alngX)
      Vsp.TableBorder = tbAll
      Sum1 = 0
      
      Set cData = Nothing
      Set cData = New CInventoryDoc
      For i = 1 To iCount
         Call cData.PopulateFromRS(1, Rs)
         
         strPrint = i & "|"
         strPrint = strPrint & cData.DOCUMENT_NO & "|"
         strPrint = strPrint & InternalDateToStringEx4(DateToStringInt(cData.DOCUMENT_DATE)) & "|"
         strPrint = strPrint & cData.SUPPLIER_NAME & ";"
         Call Vsp.AddTable(strFormat1, "", strPrint)
         
         cData.FROM_DATE = -1
         cData.TO_DATE = -1
         cData.DOCUMENT_NO = ""
         cData.COMMIT_FLAG = ""
         cData.QueryFlag = 1
         Call glbDaily.QueryInventoryDoc(cData, TempRs, iCount2, IsOK, glbErrorLog)
         Set cData.TransferItems = Nothing
         Set cData.TransferItems = New Collection
         Call glbDaily.CreateTransferItems(cData)
         For Each II In cData.TransferItems
            strPrint = "" & "|"
            strPrint = strPrint & II.ExportItem.PART_TYPE_NAME & " (" & II.ExportItem.PART_TYPE_NO & ")|"
            strPrint = strPrint & II.ExportItem.PART_DESC & " (" & II.ExportItem.PART_NO & ")|"
            strPrint = strPrint & FormatNumber(II.ExportItem.EXPORT_AMOUNT, "-") & "|"
            strPrint = strPrint & II.ExportItem.UNIT_NAME & "|"
            strPrint = strPrint & FormatNumber(II.ImportItem.INCLUDE_UNIT_PRICE, "-") & "|"
            strPrint = strPrint & FormatNumber(II.ImportItem.TOTAL_ACTUAL_PRICE, "-") & "|"
            strPrint = strPrint & II.ExportItem.LOCATION_NAME & " (" & II.ExportItem.LOCATION_NO & ")|"
            strPrint = strPrint & II.ImportItem.LOCATION_NAME & " (" & II.ImportItem.LOCATION_NO & ");"

            Call Vsp.AddTable(strFormat2, "", strPrint)
         Next II
         
         Rs.MoveNext
      Next i

      Set cData = Nothing
'      strFormat = VSP_CalTable(FOOTER_FORMAT, mdblWid, alngX)
'      strPrint = "รวม|||||" & FormatNumber(Sum1) & "|"
'      Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   Set TempRs = Nothing
   
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


