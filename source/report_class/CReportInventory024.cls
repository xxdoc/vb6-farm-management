VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory024"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory024"
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
Private Const HEADER_FORMAT1 = "^1|^3|^2|^2|^2|%P^2|^2;"
Private Const HEADER_FORMAT2 = "^1|^3|^2|^2|^2|%P^2|^2;"
Private Const DETAIL_FORMAT1 = "<4|<4|%N;"
Private Const DETAIL_FORMAT2 = "^1|^3|>2|>2|>2|%P>2|>2;"
Private Const FOOTER_FORMAT1 = "^1|^3|>2|>2|%P>2|>2;"
Private Const FOOTER_FORMAT2 = "^1|^3|>2|>2|>2|%P>2|>2;"
'NO.|วันที่|เลขที่เอกสาร|เกิด|โอนเข้า|%P|โอนออก|คงเหลือ;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

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
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "สัปดาห์เกิด", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "รับเข้า", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "สัปดาห์เกิด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เกิด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอนเข้า", "^")
   Call m_Headers1.AddField(Cf)
                                                   
   m_Left = 0
   For Each Sgi In m_ProductStatus
      If Sgi.SELECT_FLAG = "Y" Then
         Call Cf.SetFieldValue(2, "^", Sgi.STATUS_NAME, "^", Sgi.ST_STATUS_ID)
         Call m_Headers1.AddField(Cf)
         
         m_Left = m_Left + 2
         ExportWidth = ExportWidth + 2
      End If
   Next Sgi
   Call Cf.SetFieldValue(2, "^", "โอนออก", "^")
   Call m_Headers1.AddField(Cf)
   ExportWidth = ExportWidth + 2

   Call Cf.SetFieldValue(CDbl(ExportWidth), "^", "จ่ายออก", "")
   Call m_Headers0.AddField(Cf)
            
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", "^")
   Call m_Headers1.AddField(Cf)
   Call m_Headers0.AddField(Cf)
                                                                                                                                                
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
Dim TempStr As String
Dim ImportPrices As Collection
Dim ExportPrices1 As Collection
Dim ExportPrices2 As Collection
Dim Sum1 As Double
Dim ImportItems As Collection
Dim ExportItems As Collection
Dim NewDate As Date

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set ImportPrices = New Collection
   Call LoadImportPrice5(Nothing, ImportPrices, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"))
   
   Set ExportPrices1 = New Collection
   Call LoadExportPrice12(Nothing, ExportPrices1, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"))
   
   Set ExportPrices2 = New Collection
   Call LoadExportPrice13(Nothing, ExportPrices2, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("STATUS_GROUP_ID"), mcolParam("HOUSE_ID"))
   
   Set Houses = New Collection
   Call LoadImportLocation(Nothing, Houses, mcolParam("HOUSE_ID"))
   
   Set Partitems = New Collection
   
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset

'   Set InventoryBals = New Collection
'   Call LoadInventoryBalance(Nothing, InventoryBals, InternalDateToDate(DateToStringIntLow(mcolParam("FROM_DATE"))), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"), "Y")

   Set Partitems = New Collection
   Call LoadImportPig(Nothing, Partitems, mcolParam("HOUSE_ID"))
   
   Set ImportItems = New Collection
   Set ExportItems = New Collection
   
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
      
      If mcolParam("FROM_DATE") > 0 Then
         NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
         Call LoadPigImportAmount(Nothing, ImportItems, -1, InternalDateToDate(DateToStringIntHi(NewDate)), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"))
         Call LoadPigExportAmount(Nothing, ExportItems, -1, InternalDateToDate(DateToStringIntHi(NewDate)), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("HOUSE_ID"))
      End If
      
      For Each Pi In Partitems
         I = I + 1
         Call m_Details1.ClearField
         Set BodyCf = New CReportField

         j = 0
         Sum1 = 0
         For Each HeadCf In m_Headers1.Fields
                j = j + 1
              If j = 1 Then
                  TempStr = Trim(Str(I))
                  Amt = 0
               ElseIf j = 2 Then
                  TempStr = Pi.PART_NO & " (" & Pi.PIG_TYPE & ")"
                  Amt = 0
               ElseIf j = 3 Then

                  Set II = GetImportItem(ImportItems, Trim(Str(Pi.PART_ITEM_ID)))
                  Set EI = GetExportItem(ExportItems, Trim(Str(Pi.PART_ITEM_ID)))
                  Amt = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
               
                  TempStr = Trim(Str(Amt))
                  Sum1 = Sum1 + Amt
               ElseIf j = 4 Then
                  Set II = GetImportItem(ImportPrices, Pi.PART_ITEM_ID & "-" & 5)
                  Amt = II.SUM_ALL_IMPORT
                  TempStr = Trim(Str(Amt))
                  Sum1 = Sum1 + Amt
               ElseIf j = 5 Then
                  Set II = GetImportItem(ImportPrices, Pi.PART_ITEM_ID & "-" & 6)
                  Amt = II.SUM_ALL_IMPORT
                  Set II = GetImportItem(ImportPrices, Pi.PART_ITEM_ID & "-" & 7)
                  Amt = Amt + II.SUM_ALL_IMPORT
                  Set II = GetImportItem(ImportPrices, Pi.PART_ITEM_ID & "-" & 8)
                  Amt = Amt + II.SUM_ALL_IMPORT
                  Set II = GetImportItem(ImportPrices, Pi.PART_ITEM_ID & "-" & 9)
                  Amt = Amt + II.SUM_ALL_IMPORT
                  Set II = GetImportItem(ImportPrices, Pi.PART_ITEM_ID & "-" & 11)
                  Amt = Amt + II.SUM_ALL_IMPORT
                  
                  TempStr = Trim(Str(Amt))
                  Sum1 = Sum1 + Amt
               ElseIf j = m_Headers1.Fields.Count Then
                  TempStr = Trim(Str(Sum1))
                  Amt = Sum1
               ElseIf j = m_Headers1.Fields.Count - 1 Then
                  Set EI = GetExportItem(ExportPrices2, Pi.PART_ITEM_ID)
                  Amt = EI.SUM_ALL_EXPORT
                  TempStr = Trim(Str(Amt))
                  Sum1 = Sum1 - Amt
               Else
                  Set EI = GetExportItem(ExportPrices1, Pi.PART_ITEM_ID & "-" & HeadCf.TempKey)
                  Amt = EI.SUM_ALL_EXPORT
                  TempStr = Trim(Str(Amt))
                  Sum1 = Sum1 - Amt
              End If
            
              Total1(j) = Total1(j) + Amt
              Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
              Call m_Details1.AddField(BodyCf)
         Next HeadCf
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         DoEvents
      Next Pi
      
      Call GenerateFooter("รวม", "", "", Total1, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
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
   Set ImportPrices = Nothing
   Set ExportPrices1 = Nothing
   Set ExportPrices2 = Nothing
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
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
         TempStr = ""
      Else
         TempStr = Trim(Str(Tot(j)))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
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


