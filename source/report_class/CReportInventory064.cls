VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory064"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory064"
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

Private m_PigFeeds As Collection
Private m_PigBalanceAmounts As Collection
Private m_YearWeeks As Collection

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
   Set m_PigFeeds = New Collection
   Set m_PigBalanceAmounts = New Collection
   Set m_YearWeeks = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_PigFeeds = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_PigBalanceAmounts = Nothing
   Set m_YearWeeks = Nothing
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

   Call Cf.SetFieldValue(4, "^", "สัปดาห์เกิด", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "รายละเอียด", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "หมูสะสม", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "หมูเฉลี่ย/วัน", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(8, "^", "", ">")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เบอร์อาหาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่ออาหาร", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ราคา/หน่วย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "INTAKE", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@จำนวนกินสะสม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@มูลค่ากินสะสม", ">")
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
                                       vbCrLf & FromDate & ToDate & " ประเภท " & EmptyToString(mcolParam("PART_TYPE_NAME"), "N/A") _
                                       & " กลุ่มโรงเรือน " & EmptyToString(mcolParam("LOCATION_GROUP_NAME"), "N/A") & " โรงเรือน " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A") _
                                       
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function GetKey(EI As CExportItem) As String
Dim OrderBy As Long
   
   If OrderBy = 1 Then
      GetKey = EI.PIG_ID
   ElseIf OrderBy = 2 Then
      GetKey = EI.PIG_ID
   Else
      GetKey = EI.PIG_ID
   End If
End Function
Private Function GetKey1(EI As CExportItem) As String
Dim OrderBy As Long
   
   If OrderBy = 1 Then
      GetKey1 = EI.PIG_ID & "-" & EI.PIG_AGE
   ElseIf OrderBy = 2 Then
      GetKey1 = EI.PIG_ID & "-" & EI.PIG_AGE
   Else
      GetKey1 = EI.PIG_ID & "-" & EI.PIG_AGE
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
Dim Bd As CBillingDoc
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
Dim Di1 As CDoItem
Dim Di2 As CDoItem
Dim Di3 As CDoItem
Dim SumLine1 As Double
Dim SumLine2 As Double
Dim SumLine3 As Double
Dim EI As CExportItem
Dim TempEi As CExportItem
Dim TempBa As CBalanceAccum
Dim PigAmt As Double
Dim DateCount As Double
Dim DateCount2 As Long
Dim AvgPig As Double
Dim PigAmt2 As Double
Dim AvgPig2 As Double

   DateCount = Abs(DateDiff("D", mcolParam("FROM_DATE"), mcolParam("TO_DATE"))) + 1
   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   Call LoadPigFeedAmountByFeedPigByPigAge(Nothing, m_PigFeeds, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", mcolParam("LOCATION_ID"), , mcolParam("LOCATION_GROUP"))
   Call LoadSumPigBalanceAmount(Nothing, m_PigBalanceAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("LOCATION_GROUP"), mcolParam("LOCATION_ID"))
   
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
      
      Set EI = New CExportItem
      EI.EXPORT_ITEM_ID = -1
      EI.FROM_DATE = mcolParam("FROM_DATE")
      EI.TO_DATE = mcolParam("TO_DATE")
      EI.PART_TYPE = mcolParam("PART_TYPE")
      EI.PART_GROUP_ID = mcolParam("PART_GROUP")
      EI.COMMIT_FLAG = CommitTypeToFlag(mcolParam("COMMIT_TYPE"))
      EI.TO_HOUSE_GROUP_ID = mcolParam("LOCATION_GROUP")
      EI.PIG_TYPE = PigTypeToCode(mcolParam("PIG_TYPE_ID"))
      EI.HOUSE_ID = mcolParam("LOCATION_ID")
      EI.OrderBy = mcolParam("ORDER_BY")
      EI.OrderType = mcolParam("ORDER_TYPE")
      EI.YEAR_SEQ_ID = mcolParam("YEAR_SEQ_ID")
      EI.WEEK_NO = mcolParam("WEEK_NO")
      EI.DOCUMENT_TYPE = 2
      If mcolParam("INTAKE_FLAG") = "Y" Then
         EI.INTAKE_FLAG = ""
      Else
         EI.INTAKE_FLAG = "Y"
      End If
      Call EI.QueryData(54, Rs, iCount)
      I = 0
      
      PrevKey1 = ""
      PrevKey2 = ""
      If Not Rs.EOF Then
         Call EI.PopulateFromRS(54, Rs)
         PrevKey1 = GetKey(EI)
         PrevKey2 = GetKey1(EI)
         PigAmt = GenerateBalanceAmt(EI.PIG_ID, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("LOCATION_GROUP"), mcolParam("LOCATION_ID"))
         AvgPig = MyDiffEx(PigAmt, DateCount)
         Call GenerateHeader(EI.PIG_NO & " (" & EI.PIG_TYPE & ")", EI.PIG_NAME, Trim(Str(PigAmt)), Trim(Str(AvgPig)), m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
                          
         PigAmt2 = GenerateBalanceAmt(EI.PIG_ID, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("LOCATION_GROUP"), mcolParam("LOCATION_ID"), , EI.PIG_AGE, EI.PIG_NO, DateCount2)
         AvgPig2 = MyDiffEx(PigAmt2, DateCount2)
         
         Call GenerateHeader(EI.PIG_AGE, "", PigAmt2, AvgPig2, m_Details1)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      End If
'
      Set EI = Nothing
      While Not Rs.EOF
         I = I + 1

         Set EI = New CExportItem

         Call EI.PopulateFromRS(54, Rs)
         If PrevKey1 <> GetKey(EI) Then
            Call GenerateFooter("รวม", "ตามอายุ", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            Call GenerateFooter("รวม", "ตามสัปดาห์เกิด", "", Total2, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            PigAmt = GenerateBalanceAmt(EI.PIG_ID, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("LOCATION_GROUP"), mcolParam("LOCATION_ID"))
            AvgPig = MyDiffEx(PigAmt, DateCount)
            
            Call GenerateHeader(EI.PIG_NO & " (" & EI.PIG_TYPE & ")", EI.PIG_NAME, Trim(Str(PigAmt)), Trim(Str(AvgPig)), m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PigAmt2 = GenerateBalanceAmt(EI.PIG_ID, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("LOCATION_GROUP"), mcolParam("LOCATION_ID"), , EI.PIG_AGE, EI.PIG_NO, DateCount2)
            AvgPig2 = MyDiffEx(PigAmt2, DateCount2)
            
            Call GenerateHeader(EI.PIG_AGE, "", PigAmt2, AvgPig2, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
            Next j
         ElseIf PrevKey2 <> GetKey1(EI) Then
            
            Call GenerateFooter("รวม", "ตามอายุ", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PigAmt2 = GenerateBalanceAmt(EI.PIG_ID, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("LOCATION_GROUP"), mcolParam("LOCATION_ID"), , EI.PIG_AGE, EI.PIG_NO, DateCount2)
            AvgPig2 = MyDiffEx(PigAmt2, DateCount2)
            
            Call GenerateHeader(EI.PIG_AGE, "", PigAmt2, AvgPig2, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
         End If
         PrevKey1 = GetKey(EI)
         PrevKey2 = GetKey1(EI)
         Call m_Details1.ClearField

         SumLine1 = 0
         SumLine2 = 0
         SumLine3 = 0
         Set TempEi = GetExportItem(m_PigFeeds, EI.PIG_ID & "-" & EI.PART_ITEM_ID & "-" & EI.PIG_AGE)
         
         j = 0
         For Each HeadCf In m_Headers1.Fields
            j = j + 1
            If j = 1 Then
               TempStr = I
               Amt = 0
            ElseIf j = 2 Then
               TempStr = EI.PART_NO
               Amt = 0
            ElseIf j = 3 Then
               TempStr = EI.PART_DESC
               Amt = 0
            ElseIf j = 4 Then
               Amt = TempEi.EXPORT_AMOUNT
               TempStr = FormatNumber(Amt)
               SumLine1 = SumLine1 + TempEi.EXPORT_AMOUNT
            ElseIf j = 5 Then
               Amt = TempEi.EXPORT_TOTAL_PRICE
               TempStr = FormatNumber(Amt)
            ElseIf j = 6 Then
               Amt = MyDiffEx(TempEi.EXPORT_TOTAL_PRICE, TempEi.EXPORT_AMOUNT)
               TempStr = FormatNumber(Amt)
            ElseIf j = 7 Then
               Amt = MyDiffEx(TempEi.EXPORT_AMOUNT, PigAmt2)
               TempStr = FormatNumber(Amt)
            ElseIf j = 8 Then
               Amt = MyDiffEx(TempEi.EXPORT_AMOUNT, AvgPig2)
               TempStr = FormatNumber(Amt)
            ElseIf j = 9 Then
               Amt = MyDiffEx(TempEi.EXPORT_TOTAL_PRICE, AvgPig2)
               TempStr = FormatNumber(Amt)
            End If

            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            Total3(j) = Total3(j) + Amt

            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         Next HeadCf

         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)

         Rs.MoveNext
         Set Bd = Nothing
      Wend
      
      Call GenerateFooter("รวม", "ตามอายุ", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("รวม", "ตามสัปดาห์เกิด", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("สุทธิ", "", "", Total3, m_Details1, 1)
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
      ElseIf j = 6 Then
         TempStr = FormatNumber(MyDiffEx(Tot(5), Tot(4)))
      ElseIf j = 12 Then
         TempStr = FormatNumber(MyDiffEx(Tot(13), Tot(11)))
      ElseIf j = 15 Then
         TempStr = FormatNumber(MyDiffEx(Tot(16), Tot(14)))
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateHeader(Txt1 As String, ByVal Txt2 As String, ByVal Txt3 As String, ByVal Txt4, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double
Dim TempAlign As String
   
   Call Details.ClearField
   Set BodyCf = New CReportField
   
   For j = 1 To 5
      If j = 1 Then
         TempStr = Txt1
         TempWidth = 4
         TempAlign = "^"
      ElseIf j = 2 Then
         TempStr = Txt2
         TempWidth = 5
         TempAlign = "<"
      ElseIf j = 3 Then
         TempStr = FormatNumber(Val(Txt3))
         TempWidth = 2
         TempAlign = ">"
      ElseIf j = 4 Then
         TempStr = FormatNumber(Val(Txt4))
         TempWidth = 2
         TempAlign = ">"
      ElseIf j = 5 Then
         TempStr = ""
         TempWidth = 8
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
