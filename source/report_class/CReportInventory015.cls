VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory015"
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
Private Const HEADER_FORMAT = "^1|^2.5|^1.5|%P^2;"
Private Const DETAIL_FORMAT = "^1|^2.5|^1.5|%P>2;"
Private Const FOOTER_FORMAT = "^1|^2.5|<1.5|%P>2;"

Private m_HeaderStr As String
Private m_DetailStr As String
Private m_FooterStr As String
Private m_HeaderTextStr As String
Private m_Houses As Collection

'ลำดับ|ประเภทวัตถุดิบ|%Pรวม;

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Houses = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set m_Houses = Nothing
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
   strFormat = VSP_CalTable(m_HeaderStr, mdblWid, alngX)
   strPrint = m_HeaderTextStr
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim G As CLocation
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String

   Call LoadLocationEx(Nothing, m_Houses, mcolParam("HOUSE_GROUP_ID"))
   TempStr1 = ""
   TempStr2 = ""
   For Each G In m_Houses
      TempStr1 = TempStr1 & "^2|"
      TempStr2 = TempStr2 & ">2|"
      TempStr3 = TempStr3 & G.LOCATION_NAME & "|"
   Next G
   m_HeaderStr = Replace(HEADER_FORMAT, "%P", TempStr1)
   m_DetailStr = Replace(DETAIL_FORMAT, "%P", TempStr2)
   m_FooterStr = Replace(FOOTER_FORMAT, "%P", TempStr2)
   m_HeaderTextStr = Replace("NO.|สัปดาห์เกิด|ประเภท|%Pรวม;", "%P", TempStr3)
   
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

Private Function GetImportPrice(TempCol As Collection, TempKey As String, TotalAmount As Double) As Double
On Error Resume Next
Dim EI As CImportItem

   Set EI = TempCol(TempKey)
   If EI Is Nothing Then
      GetImportPrice = 0
      TotalAmount = 0
   Else
      GetImportPrice = EI.IMPORT_AMOUNT
      TotalAmount = EI.IMPORT_AMOUNT
   End If
End Function

Private Function GetExportPrice(TempCol As Collection, TempKey As String, TotalAmount As Double) As Double
On Error Resume Next
Dim EI As CExportItem

   Set EI = TempCol(TempKey)
   If EI Is Nothing Then
      GetExportPrice = 0
      TotalAmount = 0
   Else
      GetExportPrice = EI.EXPORT_AMOUNT
      TotalAmount = EI.EXPORT_AMOUNT
   End If
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim cData As CImportItem
Dim Pt As CPartItem
Dim Pt2 As CPartItem
Dim iCount As Long
Dim i As Long
Dim strFormat As String
Dim strPrint1 As String
Dim strPrint2 As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Sum1 As Double
Dim strFormat2 As String
Dim G As CLocation
Dim TempRs As ADODB.Recordset
Dim ImportPrices As Collection
Dim ExportPrices As Collection
Dim EI As CExportItem
Dim ImportTotalPrice As Double
Dim ExportTotalPrice As Double
Dim TotalAmount As Double
Dim SumPrice1 As Double
Dim SumPrice2 As Double
Dim Accums1(200) As Double
Dim Accums2(200) As Double
Dim j As Long
Dim m_PartItem As Collection

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------

'   Set m_PartItem = New Collection
'   Call LoadPartItem(Nothing, m_PartItem, , "")

   Set ImportPrices = New Collection
   Call LoadImportPrice1(Nothing, ImportPrices, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
   
   Set ExportPrices = New Collection
   Call LoadExportPrice6(Nothing, ExportPrices, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")))
   
   Set TempRs = New ADODB.Recordset
   
   Set Rs = New ADODB.Recordset
   Set cData = New CImportItem
   cData.PIG_FLAG = "Y"
   cData.FROM_DATE = mcolParam("FROM_DATE")
   cData.TO_DATE = mcolParam("TO_DATE")
   cData.OrderBy = mcolParam("ORDER_BY")
   cData.OrderType = mcolParam("ORDER_TYPE")
   Call cData.QueryData(4, Rs, iCount)

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

   Call StartExportFile(Vsp)
   Vsp.StartDoc
      strFormat = VSP_CalTable(m_DetailStr, mdblWid, alngX)
      strFormat2 = VSP_CalTable(m_FooterStr, mdblWid, alngX)
      Vsp.TableBorder = tbAll
      Sum1 = 0

      For j = 1 To UBound(Accums1)
         Accums1(j) = 0
      Next j
      While Not Rs.EOF
         i = i + 1

         strPrint1 = i & "|"
         strPrint1 = strPrint1 & NVLS(Rs("PART_NO"), "") & "|"
         strPrint1 = strPrint1 & NVLS(Rs("PIG_TYPE"), "") & "|"
                  
         SumPrice1 = 0

         j = 0
         For Each G In m_Houses
            j = j + 1
            ImportTotalPrice = GetImportPrice(ImportPrices, Trim(NVLS(Rs("PART_NO"), "")) & "-" & Trim(NVLS(Rs("PIG_TYPE"), "")) & "-" & Trim(Str(G.LOCATION_ID)), TotalAmount)
            ExportTotalPrice = GetExportPrice(ExportPrices, Trim(NVLS(Rs("PART_NO"), "")) & "-" & Trim(NVLS(Rs("PIG_TYPE"), "")) & "-" & Trim(Str(G.LOCATION_ID)), TotalAmount)
            Accums1(j) = Accums1(j) + (ImportTotalPrice - ExportTotalPrice)
            Accums2(j) = Accums2(j) + TotalAmount
            strPrint1 = strPrint1 & FormatNumber(ImportTotalPrice - ExportTotalPrice, "-") & "|"
            SumPrice1 = SumPrice1 + (ImportTotalPrice - ExportTotalPrice)
         Next G
         strPrint1 = strPrint1 & FormatNumber(SumPrice1, "-") & ";"

         Sum1 = Sum1 + 0
         Rs.MoveNext
         If Rs.EOF Then
            Call AdjustPage(Vsp, strFormat, strPrint1, 2500)
         End If
         Call Vsp.AddTable(strFormat, "", strPrint1)
      Wend

      Vsp.FontBold = True
      strPrint1 = "รวม|||"
      j = 0
      SumPrice1 = 0
      SumPrice2 = 0
      For Each G In m_Houses
         j = j + 1

         strPrint1 = strPrint1 & FormatNumber(Accums1(j), "-") & "|"
         SumPrice1 = SumPrice1 + Accums1(j)
      Next G
      strPrint1 = strPrint1 & FormatNumber(SumPrice1, "-") & ";"
      Call Vsp.AddTable(strFormat2, "", strPrint1)
      Vsp.FontBold = False
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set m_PartItem = Nothing
   Set ImportPrices = Nothing
   Set ExportPrices = Nothing
   
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


