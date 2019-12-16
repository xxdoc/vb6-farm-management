VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory060"
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

Private m_PartItems As Collection
Private m_PigStatuss As Collection

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
   Set m_PartItems = New Collection
   Set m_PigStatuss = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_PartItems = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_PigStatuss = Nothing
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

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "อายุ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "สถานะ", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "สัปดาห์เกิด", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_PRICE") = "Y" Then
      Call Cf.SetFieldValue(2, "^", "น้ำหนัก", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "น.น./ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "ราคา/ก.ก.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "ราคา/ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "น้ำหนัก", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "น.น./ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "ราคา/ก.ก.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "ราคา/ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "มูลค่า", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_COST") = "Y" Then
      Call Cf.SetFieldValue(2, "^", "ตท/ก.ก.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "ตท/ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "ตท", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ตท/ก.ก.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "ตท/ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "ตท", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
   If mcolParam("SHOW_GP") = "Y" Then
      Call Cf.SetFieldValue(2, "^", "GP/ก.ก.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "GP/ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "GP", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "GP/ก.ก.", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "GP/ตัว", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "GP", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
   
   
   
   Set Cf = Nothing
'
'  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
            
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
                                       vbCrLf & FromDate & ToDate & " โรงเรือน " & EmptyToString(mcolParam("HOUSE_NAME"), "N/A") & " สัปดาห์เกิด " & EmptyToString(mcolParam("YEAR_NO") & mcolParam("WEEK_NO"), "N/A")
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetKey(Di As CDoItem) As String
Dim OrderBy As Long
   
   If OrderBy = 1 Then
      GetKey = Di.PIG_AGE & "-" & Di.PIG_STATUS
   ElseIf OrderBy = 2 Then
      GetKey = Di.PIG_AGE & "-" & Di.PIG_STATUS
   Else
      GetKey = Di.PIG_AGE & "-" & Di.PIG_STATUS
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
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim Di1 As CDoItem
Dim Pi As CPartItem
Dim Ps As CProductStatus
Dim CapitalMovements  As Collection
Dim Mi As CMovementItem
Dim strAvg As String
   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset

   Call LoadPartItem(Nothing, m_PartItems, , "Y")
   Call LoadProductStatus(Nothing, m_PigStatuss)
   
   Set CapitalMovements = New Collection
   Call LoadTotalCapitalSellBalancePigStatusPigID(Nothing, CapitalMovements, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("PIG_STATUS"), mcolParam("HOUSE_ID"), , DocType2Set(mcolParam("BILL_SUBTYPE")))
   
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

      Set Di1 = New CDoItem
      Di1.DO_ITEM_ID = -1
      Di1.FROM_DATE = mcolParam("FROM_DATE")
      Di1.TO_DATE = mcolParam("TO_DATE")
      Di1.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
      Di1.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
      Di1.PIG_STATUS = mcolParam("PIG_STATUS")
'      Di1.DocSubTypeSet = "(1, 2) " 'หมู, วัตถุดิบ
      Di1.PIG_FLAG = "Y"
      Di1.YEAR_SEQ_ID = mcolParam("YEAR_SEQ_ID")
      Di1.WEEK_NO = mcolParam("WEEK_NO")
      Di1.LOCATION_ID = mcolParam("HOUSE_ID")
      Di1.OrderBy = mcolParam("ORDER_BY")
      Di1.OrderType = mcolParam("ORDER_TYPE")
      If Di1.OrderBy <= 0 Then
         Di1.OrderBy = 1
      End If
      Call Di1.QueryData(15, Rs, iCount)
      I = 0

      PrevKey1 = ""
      If Not Rs.EOF Then
         Call Di1.PopulateFromRS(15, Rs)
         PrevKey1 = GetKey(Di1)
      End If

      Set Di1 = Nothing
      While Not Rs.EOF
         I = I + 1

         Set Di1 = New CDoItem

         Call Di1.PopulateFromRS(15, Rs)
         If PrevKey1 <> GetKey(Di1) Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
         End If
         PrevKey1 = GetKey(Di1)

         Call m_Details1.ClearField
         Set Pi = GetPartItem(m_PartItems, Trim(Str(Di1.PART_ITEM_ID)))
         Set Ps = GetProductStatus(m_PigStatuss, Trim(Str(Di1.PIG_STATUS)))
         Set Mi = GetMovementItem(CapitalMovements, Trim(Di1.PIG_STATUS & "-" & Di1.PIG_AGE & "-" & Di1.PART_ITEM_ID))

         j = 0
         For Each HeadCf In m_Headers1.Fields
            j = j + 1
            If j = 1 Then
               TempStr = I
               Amt = 0
            ElseIf j = 2 Then
               TempStr = Di1.PIG_AGE
               Amt = 0
            ElseIf j = 3 Then
               TempStr = Ps.PRODUCT_STATUS_NAME
               Amt = 0
            ElseIf j = 4 Then
               TempStr = Pi.PART_NO & " (" & Pi.PIG_TYPE & ")"
               Amt = 0
            ElseIf j = 5 Then
               Amt = Di1.ITEM_AMOUNT
               TempStr = FormatNumber(Amt)
            ElseIf j = 6 Then
               Amt = Di1.TOTAL_WEIGHT
               TempStr = FormatNumber(Amt)
            ElseIf j = 7 Then
               Amt = MyDiffEx(Di1.TOTAL_WEIGHT, Di1.ITEM_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 8 Then
               Amt = MyDiffEx(Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT, Di1.TOTAL_WEIGHT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 9 Then
               Amt = MyDiffEx(Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT, Di1.ITEM_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 10 Then
               Amt = Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT
               TempStr = FormatNumber(Amt)
            ElseIf j = 11 Then
               Amt = MyDiffEx(Mi.CAPITAL_AMOUNT, Di1.TOTAL_WEIGHT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 12 Then
               Amt = MyDiffEx(Mi.CAPITAL_AMOUNT, Di1.ITEM_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 13 Then
               Amt = Mi.CAPITAL_AMOUNT
               TempStr = FormatNumber(Amt)
            ElseIf j = 14 Then
               Amt = MyDiffEx(Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT - Mi.CAPITAL_AMOUNT, Di1.TOTAL_WEIGHT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 15 Then
               Amt = MyDiffEx(Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT - Mi.CAPITAL_AMOUNT, Di1.ITEM_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 16 Then
               Amt = Di1.TOTAL_PRICE - Di1.DISCOUNT_AMOUNT - Mi.CAPITAL_AMOUNT
               TempStr = FormatNumber(Amt)
            End If

            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt

            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         Next HeadCf

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

         Rs.MoveNext
         Set Di1 = Nothing
      Wend
       
      Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      

     strAvg = queryForAverageOfAge
      Call GenerateFooter("สุทธิ", strAvg, "", Total2, m_Details1, 1)
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
      ElseIf j = 4 Then
         TempStr = ""
      ElseIf j = 7 Then
         TempStr = FormatNumber(MyDiffEx(Tot(6), Tot(5)))
      ElseIf j = 8 Then
         TempStr = FormatNumber(MyDiffEx(Tot(10), Tot(6)))
      ElseIf j = 9 Then
         TempStr = FormatNumber(MyDiffEx(Tot(10), Tot(5)))
      ElseIf j = 10 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf (j = 11) Then
         TempStr = FormatNumber(MyDiffEx(Tot(13), Tot(6)))
      ElseIf j = 12 Then
         TempStr = FormatNumber(MyDiffEx(Tot(13), Tot(5)))
      ElseIf (j = 13) Then
         TempStr = FormatNumber(Tot(j))
      ElseIf (j = 14) Then
         TempStr = FormatNumber(MyDiffEx(Tot(16), Tot(6)))
      ElseIf j = 15 Then
         TempStr = FormatNumber(MyDiffEx(Tot(16), Tot(5)))
      ElseIf (j = 16) Then
         TempStr = FormatNumber(Tot(j))
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


Private Function queryForAverageOfAge() As Double
Dim Di  As CDoItem
Dim iCount As Long
Dim Avg As Double
Dim Pi As CPartItem
Dim Ps As CProductStatus
Dim CapitalMovements  As Collection
Dim Mi As CMovementItem
Dim countDay As Double
Dim TotalAmount As Double
Dim mod7days As Long
Dim diff7day As Long


    Set Di = New CDoItem
      Di.DO_ITEM_ID = -1
      Di.FROM_DATE = mcolParam("FROM_DATE")
      Di.TO_DATE = mcolParam("TO_DATE")
      Di.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
      Di.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
      Di.PIG_STATUS = mcolParam("PIG_STATUS")
      Di.PIG_FLAG = "Y"
      Di.YEAR_SEQ_ID = mcolParam("YEAR_SEQ_ID")
      Di.WEEK_NO = mcolParam("WEEK_NO")
      Di.LOCATION_ID = mcolParam("HOUSE_ID")
      Di.OrderBy = mcolParam("ORDER_BY")
      Di.OrderType = mcolParam("ORDER_TYPE")
      If Di.OrderBy <= 0 Then
         Di.OrderBy = 1
      End If
      Call Di.QueryData(36, Rs, iCount)


      Set Di = Nothing
      While Not Rs.EOF
         Set Di = New CDoItem
          Call Di.PopulateFromRS(36, Rs)


         Set Pi = GetPartItem(m_PartItems, Trim(Str(Di.PART_ITEM_ID)))
'         Set Ps = GetProductStatus(m_PigStatuss, Trim(Str(Di.PIG_STATUS)))
'         Set Mi = GetMovementItem(CapitalMovements, Trim(Di.PIG_STATUS & "-" & Di.PIG_AGE & "-" & Di.PART_ITEM_ID))
         countDay = GetAgeDay(Pi.PART_NO, Di.DOCUMENT_DATE) 'Di.PIG_AGE
         TotalAmount = TotalAmount + Di.ITEM_AMOUNT
         Avg = Avg + (countDay * Di.ITEM_AMOUNT)
         
          Rs.MoveNext
         Set Di = Nothing
      Wend
      
     Avg = FormatNumber(MyDiff(Avg, TotalAmount))
     Avg = FormatNumber(MyDiffEx(Avg, 7))
   ' diff7day = MyDiffEx(Avg, 7)
     'mod7days = Avg Mod 7


    
    queryForAverageOfAge = Avg 'diff7day & "." & mod7days
End Function
