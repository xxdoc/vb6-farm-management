VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBudget019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBudget019"
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

Private m_PigBirthAmounts As Collection
Private m_RevenueTotalPrices As Collection
Private m_SellPigTotalPrices As Collection
Private m_PigBalanceAmounts As Collection
Private m_TotalSellCapitals As Collection
Private m_TotalLossCapitals As Collection
Private m_Year As Long
Private m_FromDate As Date
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
   Set m_PigBirthAmounts = New Collection
   Set m_RevenueTotalPrices = New Collection
   Set m_SellPigTotalPrices = New Collection
   Set m_PigBalanceAmounts = New Collection
   Set m_TotalSellCapitals = New Collection
   Set m_TotalLossCapitals = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_PigBirthAmounts = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_RevenueTotalPrices = Nothing
   Set m_SellPigTotalPrices = Nothing
   Set m_PigBalanceAmounts = Nothing
   Set m_TotalSellCapitals = Nothing
   Set m_TotalLossCapitals = Nothing
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
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long
Dim YearFmt As String

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
   m_Year = Val(mcolParam("YEAR_NO"))

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "สัปดาห์เกิด", "^")
   Call m_Headers1.AddField(Cf)
   For J = 1 To 12
      YearFmt = Val(Format(m_Year, "0000")) - 543 & "-" & Format(J, "00")
      Call Cf.SetFieldValue(2, "^", IntToThaiMonthEx(J) & " " & m_Year, "^", YearFmt)
      Call m_Headers1.AddField(Cf)
   Next J
   Call Cf.SetFieldValue(0, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing

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
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   m_FromDate = DateSerial(mcolParam("YEAR_NO") - 543, 1, 1)
   m_ToDate = DateSerial(mcolParam("YEAR_NO") - 543, 12, LastDayOfMonth(m_FromDate))
      
   FromDate = "ปี " & mcolParam("YEAR_NO")

   MODULE_DESC = mcolParam("BATCH_NAME") & vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & "  " & ToPaidDate

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

Private Function IntToText(ID As Long) As String
   If ID = 1 Then
      IntToText = "ยกมา"
   ElseIf ID = 2 Then
      IntToText = "เกิด"
   ElseIf ID = 3 Then
      IntToText = "รับเข้า"
   ElseIf ID = 4 Then
      IntToText = "จ่ายออก"
   ElseIf ID = 5 Then
      IntToText = "คงเหลือ"
   End If
End Function

Private Function GetAmount(Ma As CMonthlyAccum, ID As Long) As Double
   If ID = 1 Then
      GetAmount = Ma.BALANCE_AMOUNT1
   ElseIf ID = 2 Then
      GetAmount = Ma.BIRTH_AMOUNT
   ElseIf ID = 3 Then
      GetAmount = Ma.IMPORT_AMOUNT
   ElseIf ID = 4 Then
      GetAmount = Ma.EXPORT_AMOUNT
   ElseIf ID = 5 Then
      GetAmount = Ma.BALANCE_AMOUNT2
   End If
End Function

Private Function GeneratePigAgeStr(PigNo As String, MonthStr As String) As String
Dim DateStr As String
Dim TempDate As Date
Dim FirstDate As Date
Dim LastDate As Date

   DateStr = MonthStr & "-01 00:00:00"
   TempDate = InternalDateToDate(DateStr)
   Call GetFirstLastDate(TempDate, FirstDate, LastDate)
   
   GeneratePigAgeStr = Minus2Zero(GetAge(PigNo, FirstDate)) & "--" & Minus2Zero(GetAge(PigNo, LastDate))
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim J As Long
Dim k As Long
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
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Width As Double
Dim TempRs As ADODB.Recordset
Dim TempApt As CDoItem
Dim Hc As CReportField
Dim II As CImportItem
Dim TempII As CImportItem
Dim SumLine1 As Double
Dim SumLine2 As Double
Dim PigAgeStr As String

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

   For J = 1 To UBound(Total1)
      Total1(J) = 0
      Total2(J) = 0
      Total3(J) = 0
   Next J
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Set BodyCf = New CReportField

   Set II = New CImportItem
   II.IMPORT_ITEM_ID = -1
   II.BATCH_ID = mcolParam("BATCH_ID")
'   Ii.LOCATION_ID = mcolParam("LOCATION_ID")
   II.FROM_DATE = m_FromDate
   II.TO_DATE = m_ToDate
   II.PIG_FLAG = "Y"
'   II.DOCUMENT_TYPE = 5
   II.OrderBy = 3
   Call II.QueryData(9, Rs, iCount)

   I = 0
   While Not Rs.EOF
      I = I + 1
      Call II.PopulateFromRS(9, Rs)

      Call m_Details1.ClearField
      Set Hc = m_Headers1.Fields(1)
      TempStr = I
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      Set Hc = m_Headers1.Fields(2)
      TempStr = II.PART_NO & " (" & II.PIG_TYPE & ")"
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)

      SumLine1 = 0
      SumLine2 = 0
      For J = 1 To 12
          Set Hc = m_Headers1.Fields(J + 2)
      
         PigAgeStr = GeneratePigAgeStr(II.PART_NO, Hc.TempKey)
         
         TempStr = PigAgeStr
         Total1(J + 2) = Total1(J + 2) + 0
         Total2(J + 2) = Total2(J + 2) + 0

         SumLine1 = SumLine1 + 0
         SumLine2 = SumLine2 + 0

         Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next J
      Total1(15) = Total1(15) + SumLine1
      Total2(15) = Total2(15) + SumLine2
      TempStr = FormatNumber(SumLine1, "-")
      Call BodyCf.SetFieldValue(Hc.Width, Hc.BodyAlign, TempStr)
'      Call m_Details1.AddField(BodyCf)

      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)

      Rs.MoveNext
   Wend

   Call GenerateFooter("รวม", "", "", Total2, m_Details1, 0)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
'   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

   Set II = Nothing
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

Private Sub GenerateSubHeader(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   TempStr = txt1
   TempWidth = m_Headers1.Fields(1).Width + m_Headers1.Fields(2).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

   TempStr = txt2
   TempWidth = 0
   For J = 3 To m_Headers1.Fields.Count
      TempWidth = TempWidth + m_Headers1.Fields(J).Width
   Next J
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "^")
   Call Details.AddField(BodyCf)

   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim J As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   J = 0
   For Each HeadCf In m_Headers1.Fields
      J = J + 1
      If J = 1 Then
         TempStr = txt1
      ElseIf J = 2 Then
         TempStr = txt2
      Else
         TempStr = FormatNumber(Tot(J))
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






