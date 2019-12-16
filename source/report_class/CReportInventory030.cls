VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory030"
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
   
   '===
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "สถานะสุกร", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "สุกรขายรายวัน", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "สุกรขายสะสม", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "สุกรขายเฉลี่ย/วัน", "<")
   Call m_Headers0.AddField(Cf)
   '===
   
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "สถานะสุกร", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก/ตัว", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก/ตัว", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">")
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
Dim Ei1 As CExportItem
Dim EI2 As CExportItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PigStatuss As Collection
Dim Ps As CProductStatus
Dim Sum1 As Double
Dim ExportItems As Collection
Dim ExportItemExs As Collection
Dim DateCount As Long

Dim TempExport1 As CExportItem
Dim TempExport2 As CExportItem

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set PigStatuss = New Collection
   Call LoadProductStatus(Nothing, PigStatuss, mcolParam("STATUS_ID"))
   
   Set ExportItems = New Collection
   Call LoadPigSellAmount(Nothing, ExportItems, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("BILL_SUBTYPE"))
   
   Set ExportItemExs = New Collection
   Call LoadPigSellAmount(Nothing, ExportItemExs, mcolParam("TO_DATE"), mcolParam("TO_DATE"), CommitTypeToFlag(mcolParam("COMMIT_TYPE")), mcolParam("BILL_SUBTYPE"))
   
   DateCount = DateDiff("D", mcolParam("FROM_DATE"), mcolParam("TO_DATE")) + 1
   
   
   
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
      I = 0
      Set BodyCf = New CReportField
      
      For Each Ps In PigStatuss
         
         If CheckStatusNo(Ps.PRODUCT_STATUS_NO) Then
            I = I + 1
            For j = 1 To UBound(Total1)
               Total2(j) = 0
            Next j
            
            Call m_Details1.ClearField
            j = 0
            
            Set Ei1 = GetExportItem(ExportItemExs, Trim(Str(Ps.PRODUCT_STATUS_ID)))
            Set EI2 = GetExportItem(ExportItems, Trim(Str(Ps.PRODUCT_STATUS_ID)))
            
            Set TempExport1 = GetObject("CExportItem", T706Collection1, Trim(Ps.PRODUCT_STATUS_NO), False)
            Set TempExport2 = GetObject("CExportItem", T706Collection2, Trim(Ps.PRODUCT_STATUS_NO), False)
            For Each HeadCf In m_Headers1.Fields
               j = j + 1
               If j = 1 Then
                  TempStr = I
                  Amt = 0
               ElseIf j = 2 Then
                  TempStr = Ps.PRODUCT_STATUS_NAME & " (" & Ps.PRODUCT_STATUS_NO & ")"
                  Amt = 0
               ElseIf j = 3 Then
                  If TempExport1 Is Nothing Then
                     Amt = Ei1.EXPORT_AMOUNT
                  Else
                     Amt = TempExport1.EXPORT_AMOUNT
                  End If
                  TempStr = Amt
               ElseIf j = 4 Then
                  If TempExport1 Is Nothing Then
                     Amt = Ei1.TOTAL_WEIGHT
                  Else
                     Amt = TempExport1.TOTAL_WEIGHT
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 5 Then
                  If TempExport1 Is Nothing Then
                     If Ei1.TOTAL_WEIGHT <> 0 Then
                        Amt = (Ei1.SUM_ALL_EXPORT_AMT - Ei1.DISCOUNT_AMOUNT) / Ei1.TOTAL_WEIGHT
                     Else
                        Amt = 0
                     End If
                  Else
                     If TempExport1.TOTAL_WEIGHT <> 0 Then
                        Amt = TempExport1.EXPORT_AVG_PRICE
                     Else
                        Amt = 0
                     End If
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 6 Then
                  If TempExport1 Is Nothing Then
                     Amt = Ei1.SUM_ALL_EXPORT_AMT - Ei1.DISCOUNT_AMOUNT
                  Else
                     Amt = TempExport1.EXPORT_TOTAL_PRICE
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 7 Then
                  If TempExport1 Is Nothing Then
                     Amt = MyDiffEx(Ei1.TOTAL_WEIGHT, Ei1.EXPORT_AMOUNT)
                  Else
                     Amt = MyDiffEx(TempExport1.TOTAL_WEIGHT, TempExport1.EXPORT_AMOUNT)
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 8 Then
                  If TempExport2 Is Nothing Then
                     Amt = EI2.EXPORT_AMOUNT
                  Else
                     Amt = TempExport2.EXPORT_AMOUNT
                  End If
                  TempStr = FormatNumberInt(Amt)
               ElseIf j = 9 Then
                  If TempExport2 Is Nothing Then
                     Amt = EI2.TOTAL_WEIGHT
                  Else
                     Amt = TempExport2.TOTAL_WEIGHT
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 10 Then
                  If TempExport2 Is Nothing Then
                     If EI2.TOTAL_WEIGHT <> 0 Then
                        Amt = (EI2.SUM_ALL_EXPORT_AMT - EI2.DISCOUNT_AMOUNT) / EI2.TOTAL_WEIGHT
                     Else
                        Amt = 0
                     End If
                  Else
                     If TempExport2.TOTAL_WEIGHT <> 0 Then
                        Amt = (TempExport2.EXPORT_AVG_PRICE)
                     Else
                        Amt = 0
                     End If
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 11 Then
                  If TempExport2 Is Nothing Then
                     Amt = EI2.SUM_ALL_EXPORT_AMT - EI2.DISCOUNT_AMOUNT
                  Else
                     Amt = TempExport2.EXPORT_TOTAL_PRICE
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 12 Then
                  If TempExport2 Is Nothing Then
                     Amt = MyDiffEx(EI2.TOTAL_WEIGHT, EI2.EXPORT_AMOUNT)
                  Else
                     Amt = MyDiffEx(TempExport2.TOTAL_WEIGHT, TempExport2.EXPORT_AMOUNT)
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 13 Then
                  If TempExport2 Is Nothing Then
                     Amt = EI2.EXPORT_AMOUNT / DateCount
                  Else
                     Amt = TempExport2.EXPORT_AMOUNT / DateCount
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 14 Then
                  If TempExport2 Is Nothing Then
                     Amt = EI2.TOTAL_WEIGHT / DateCount
                  Else
                     Amt = TempExport2.TOTAL_WEIGHT / DateCount
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 15 Then
                  If TempExport2 Is Nothing Then
                     If EI2.TOTAL_WEIGHT <> 0 Then
                        Amt = (EI2.SUM_ALL_EXPORT_AMT - EI2.DISCOUNT_AMOUNT) / EI2.TOTAL_WEIGHT
                     Else
                        Amt = 0
                     End If
                  Else
                     If TempExport2.TOTAL_WEIGHT <> 0 Then
                        Amt = (TempExport2.EXPORT_AVG_PRICE)
                     Else
                        Amt = 0
                     End If
                  End If
                  TempStr = FormatNumber(Amt, "-")
               ElseIf j = 16 Then
                  If TempExport2 Is Nothing Then
                     Amt = (EI2.SUM_ALL_EXPORT_AMT - EI2.DISCOUNT_AMOUNT) / DateCount
                  Else
                     Amt = (TempExport2.EXPORT_TOTAL_PRICE) / DateCount
                  End If
                  TempStr = FormatNumber(Amt, "-")
               Else
                  TempStr = ""
               End If
   
               Total1(j) = Total1(j) + Amt
               Total2(j) = Total2(j) + Amt
   
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            Next HeadCf
   
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      Next Ps

      Call GenerateFooter("รวม", "", "", Total1, m_Details1, DateCount)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set PigStatuss = Nothing
   Set ExportItems = Nothing
   Set ExportItemExs = Nothing
   
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
      ElseIf (j = 3) Then
         TempStr = Tot(j)
      ElseIf j = 4 Then
         TempStr = FormatNumber(Tot(4))
      ElseIf j = 5 Then
         TempStr = FormatNumber(MyDiff(Tot(6), Tot(4)))
      ElseIf j = 6 Then
         TempStr = FormatNumber(Tot(6))
      ElseIf (j = 7) Then
         TempStr = FormatNumber(MyDiff(Tot(4), Tot(3)))
      ElseIf (j = 8) Then
         TempStr = FormatNumberInt(Tot(j))
      ElseIf j = 9 Then
         TempStr = FormatNumber(Tot(9))
      ElseIf j = 10 Then
         TempStr = FormatNumber(MyDiff(Tot(11), Tot(9))) '11, 9
      ElseIf j = 11 Then
         TempStr = FormatNumber(Tot(11))
      ElseIf j = 12 Then
         TempStr = FormatNumber(MyDiff(Tot(9), Tot(8)))
      ElseIf (j = 13) Then
         TempStr = FormatNumber(MyDiff(Tot(j), 1))
      ElseIf j = 14 Then
         TempStr = FormatNumber(MyDiff(Tot(j), 1))
      ElseIf j = 15 Then
         TempStr = FormatNumber(MyDiff(Tot(11), Tot(9))) '11, 9
      ElseIf j = 16 Then
         TempStr = FormatNumber(MyDiff(Tot(j), 1))
      Else
         TempStr = Tot(j)
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
Private Function CheckStatusNo(Status As String) As Boolean
   Dim TempStatus As Boolean
   Dim TempString As String
   Dim TempNo As String
   TempStatus = True
   TempString = mcolParam("NO_STATUS_NO")
   While Len(TempString) > 0
      If InStr(1, TempString, ",") > 0 Then
         TempNo = Left(TempString, InStr(1, TempString, ",") - 1)
      Else
         TempNo = TempString
      End If
      If TempNo = Status Then
         TempStatus = False
         CheckStatusNo = False
         Exit Function
      End If
      If TempNo = TempString Then
         TempString = ""
      Else
         TempString = Mid(TempString, InStr(1, TempString, ",") + 1)
      End If
   Wend
   CheckStatusNo = TempStatus
End Function

