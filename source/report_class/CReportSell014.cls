VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSell014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell013"
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

Private ExceptColl As Collection
Private ProductStatusTypeColl As Collection
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
   Set ExceptColl = New Collection
   Set ProductStatusTypeColl = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set ExceptColl = Nothing
   Set ProductStatusTypeColl = Nothing
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
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long

Dim TempFromDate As Date
Dim TempToDate As Date
Dim TempDate As Date
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(40, "^", "ประเภทสินค้า", "<")
   Call m_Headers0.AddField(Cf)
   
   TempDate = mcolParam("FROM_DATE")
   TempFromDate = mcolParam("FROM_DATE")
   TempToDate = mcolParam("TO_DATE")
   
   While TempDate <= mcolParam("TO_DATE")
      Call GetFirstLastDate(TempDate, TempFromDate, TempToDate)
      If mcolParam("SUMMARY_COLUMN") = "Y" Then
         Call Cf.SetFieldValue(0, "^", Day(TempFromDate) & "-" & Day(TempToDate) & "/" & Format(Month(TempFromDate), "00") & "/" & (Year(TempFromDate) + 543), "<")
      Else
         If Month(TempDate) = Month(mcolParam("TO_DATE")) Then
            Call Cf.SetFieldValue(70, "^", Day(TempFromDate) & "-" & Day(mcolParam("TO_DATE")) & "/" & Format(Month(TempFromDate), "00") & "/" & (Year(TempFromDate) + 543), "<")
         Else
            Call Cf.SetFieldValue(70, "^", Day(TempFromDate) & "-" & Day(TempToDate) & "/" & Format(Month(TempFromDate), "00") & "/" & (Year(TempFromDate) + 543), "<")
         End If
      End If
      Call m_Headers0.AddField(Cf)
      TempDate = DateAdd("M", 1, TempDate)
   Wend
   Call Cf.SetFieldValue(70, "^", "รวม", "<")
   Call m_Headers0.AddField(Cf)
      
   '===
   Call Cf.SetFieldValue(40, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   
   TempDate = mcolParam("FROM_DATE")
   TempFromDate = mcolParam("FROM_DATE")
   TempToDate = mcolParam("TO_DATE")
   
   While TempDate <= mcolParam("TO_DATE")
      Call GetFirstLastDate(TempDate, TempFromDate, TempToDate)
      If mcolParam("SUMMARY_COLUMN") = "Y" Then
         Call Cf.SetFieldValue(0, "^", "จำนวน", ">", "SHOW_DPOINT")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(0, "^", "น้ำหนัก", ">", "SHOW_DPOINT")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(0, "^", "ยอดเงิน", ">")
         Call m_Headers1.AddField(Cf)
      Else
         Call Cf.SetFieldValue(20, "^", "จำนวน", ">", "SHOW_DPOINT")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(20, "^", "น้ำหนัก", ">", "SHOW_DPOINT")
         Call m_Headers1.AddField(Cf)
         Call Cf.SetFieldValue(30, "^", "ยอดเงิน", ">")
         Call m_Headers1.AddField(Cf)
      End If
      TempDate = DateAdd("M", 1, TempDate)
   Wend
   
   Call Cf.SetFieldValue(20, "^", "จำนวน", ">", "SHOW_DPOINT")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "น้ำหนัก", ">", "SHOW_DPOINT")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(30, "^", "ยอดเงิน", ">")
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
Dim ToPaidDate As String
Dim NewDate As Date

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
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
  
  If Len(mcolParam("EXCEPT_PIG_STATUS")) > 0 Then
      Call GetExceptPigStatusNo
  End If
  
  Call GetProductStatusTypeYYYYMM(ProductStatusTypeColl, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
  
  
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")

   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & vbCrLf & mcolParam("REPORT_NAME") & _
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

Private Function GetItemFromListIndex(TempCol As Collection, Ind As Long) As Object
   Set GetItemFromListIndex = TempCol(Ind)
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
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
Dim PrevKey3  As String
Dim PrevKey4 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim Di1 As CDoItem
Dim Hc As CReportField
Dim tempPigStatus As CProductStatus
Dim showFlag As Boolean
Dim TempDi As CDoItem
Dim strDesc As String

Dim TempFromDate As Date
Dim TempToDate As Date
Dim TempDate As Date
Dim sumAmount As Double
Dim sumWeight As Double
Dim sumTotalPrice  As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   
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
   Set BodyCf = New CReportField
   Set Di1 = New CDoItem
   Di1.DO_ITEM_ID = -1
   Di1.FROM_DATE = mcolParam("FROM_DATE")
   Di1.TO_DATE = mcolParam("TO_DATE")
   Call Di1.QueryData(42, Rs, iCount)
      
   i = 0
   While Not Rs.EOF
      i = i + 1
      Call Di1.PopulateFromRS(42, Rs)
      
      showFlag = True
      
      If Len(mcolParam("EXCEPT_PIG_STATUS")) > 0 Then
         If Di1.PIG_FLAG = "Y" Then
            Set tempPigStatus = GetObject("", ExceptColl, Trim(Di1.PRODUCT_STATUS_NO), False)
            If Not (tempPigStatus Is Nothing) Then
               'อยู่ในรายการที่ยกเว้น
               showFlag = False
            End If
         End If
      End If
      
      If showFlag Then
                
         Call m_Details1.ClearField
          Set Hc = m_Headers1.Fields(1)
          If Di1.PIG_FLAG = "Y" Then
            TempStr = Di1.PRODUCT_STATUS_NAME & "(" & Di1.PRODUCT_STATUS_NO & ")"
         ElseIf Len(Di1.PART_TYPE_NO) > 0 Then
            TempStr = Di1.PART_TYPE_NAME & "(" & Di1.PART_TYPE_NO & ")"
         Else
            TempStr = Di1.REVENUE_NAME & "(" & Di1.REVENUE_NO & ")"
         End If
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                  
         TempDate = mcolParam("FROM_DATE")
         TempFromDate = mcolParam("FROM_DATE")
         TempToDate = mcolParam("TO_DATE")
                  
         sumAmount = 0
         sumWeight = 0
         sumTotalPrice = 0
         
         j = 1
         While TempDate <= mcolParam("TO_DATE")
            Call GetFirstLastDate(TempDate, TempFromDate, TempToDate)
            
            Set TempDi = GetObject("CDoItem", ProductStatusTypeColl, Trim(Di1.REVENUE_NO & "-" & Di1.PRODUCT_STATUS_NO & "-" & Di1.PART_TYPE_NO & "-" & Di1.PIG_FLAG & "-" & (Year(TempDate) & "-" & Format(Month(TempDate), "00"))))
            
            j = j + 1
            Set Hc = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(TempDi.ITEM_AMOUNT, 0)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + (TempDi.ITEM_AMOUNT)
            Total2(j) = Total2(j) + (TempDi.ITEM_AMOUNT)
            sumAmount = sumAmount + TempDi.ITEM_AMOUNT
         
            j = j + 1
            Set Hc = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(TempDi.TOTAL_WEIGHT, 0)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + (TempDi.TOTAL_WEIGHT)
            Total2(j) = Total2(j) + (TempDi.TOTAL_WEIGHT)
            sumWeight = sumWeight + TempDi.TOTAL_WEIGHT
            
            j = j + 1
            Set Hc = m_Headers1.Fields(j)
            TempStr = FormatNumber(TempDi.TOTAL_PRICE)
            Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(j) = Total1(j) + (TempDi.TOTAL_PRICE)
            Total2(j) = Total2(j) + (TempDi.TOTAL_PRICE)
            sumTotalPrice = sumTotalPrice + TempDi.TOTAL_PRICE
            
            TempDate = DateAdd("M", 1, TempDate)
         Wend
                  
         j = j + 1
         Set Hc = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(sumAmount, 0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + (sumAmount)
         Total2(j) = Total2(j) + (sumAmount)
         
         j = j + 1
         Set Hc = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(sumWeight, 0)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + (sumWeight)
         Total2(j) = Total2(j) + (sumWeight)
         
         j = j + 1
         Set Hc = m_Headers1.Fields(j)
         TempStr = FormatNumber(sumTotalPrice)
         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(j) = Total1(j) + (sumTotalPrice)
         Total2(j) = Total2(j) + (sumTotalPrice)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      
      Rs.MoveNext
   Wend
   
   Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
         
   Set Di1 = Nothing
   Set BodyCf = Nothing
    Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
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

Private Sub GenerateSubHeader(TXT1 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   TempStr = TXT1
   TempWidth = m_Headers1.Fields(1).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

'   TempStr = Txt2
'   TempWidth = 0
'   For j = 3 To m_Headers1.Fields.Count
'      TempWidth = TempWidth + m_Headers1.Fields(j).Width
'   Next j
'   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "^")
'   Call Details.AddField(BodyCf)

   Set BodyCf = Nothing
End Sub


Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList)
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
         TempStr = TXT1
      ElseIf HeadCf.TempKey = "SHOW_DPOINT" Then
         TempStr = FormatNumberToNull(Tot(j), 0)
      Else
         TempStr = FormatNumber(Tot(j))
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
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(Now), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
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
Private Function GetExceptPigStatusNo()
Dim TempStr As String
Dim tempPigStatus As CProductStatus
   GetExceptPigStatusNo = ""
   If Len(mcolParam("EXCEPT_PIG_STATUS")) > 0 Then
      TempStr = mcolParam("EXCEPT_PIG_STATUS")
      
      If Left(TempStr, 1) = "(" Then
         TempStr = Mid(TempStr, 2)
      End If
      If Right(TempStr, 1) = ")" Then
         TempStr = Left(TempStr, Len(TempStr) - 1)
      End If
      
      While Len(TempStr) > 0
         If InStr(1, TempStr, ",") > 0 Then
            Set tempPigStatus = New CProductStatus
            tempPigStatus.PRODUCT_STATUS_NO = Left(TempStr, InStr(1, TempStr, ",") - 1)
            Call ExceptColl.Add(tempPigStatus, Trim(tempPigStatus.PRODUCT_STATUS_NO))
            TempStr = Mid(TempStr, InStr(1, TempStr, ",") + 1, Len(TempStr) - InStr(1, TempStr, ","))
         Else
            Set tempPigStatus = New CProductStatus
            tempPigStatus.PRODUCT_STATUS_NO = TempStr
            Call ExceptColl.Add(tempPigStatus, Trim(tempPigStatus.PRODUCT_STATUS_NO))
            Set tempPigStatus = Nothing
            TempStr = ""
         End If
      Wend
   End If
End Function
Public Function GetKey1(Di As CDoItem) As String
   GetKey1 = Di.YYYYMM
End Function
Public Function GetKey2(Di As CDoItem) As String

If Di.PRODUCT_STATUS_NO <> "" Then
   GetKey2 = Trim(Di.PRODUCT_STATUS_NO) & "-" & Trim(Di.PRODUCT_STATUS_NAME)
Else
   GetKey2 = Trim(Di.PART_TYPE_NO) & "-" & Trim(Di.PART_TYPE_NAME)
End If
End Function

