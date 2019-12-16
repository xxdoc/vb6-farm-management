VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventory043"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventory043"
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
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private m_ExpenseTypes1 As Collection
Private m_ExpenseTypes2 As Collection
Private m_PartGroup As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_ExpenseTypes1 = New Collection
   Set m_ExpenseTypes2 = New Collection
   Set m_PartGroup = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_ExpenseTypes1 = Nothing
   Set m_ExpenseTypes2 = Nothing
   Set m_PartGroup = Nothing
      
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
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
   Call m_Headers2.ClearField
   
   Set Cf = New CReportField
                  
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "โรงเรือน", "<")
   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "แม่+พ่อ", "^")
'   Call m_Headers1.AddField(Cf)
'   Call Cf.SetFieldValue(2, "^", "จำนวน", "^")
'   Call m_Headers1.AddField(Cf)
   
   Dim Et As CExpenseType
'   For Each Et In m_ExpenseTypes1
'      Call Cf.SetFieldValue(2, "^", Et.EXPENSE_TYPE_NAME, ">", "E-" & Et.EXPENSE_TYPE_ID)
'      Call m_Headers1.AddField(Cf)
'   Next Et
'
   Dim Pg As CPartGroup
   For Each Pg In m_PartGroup
      Call Cf.SetFieldValue(2, "^", Pg.PART_GROUP_NAME, ">", "I-" & Pg.PART_GROUP_ID)
      Call m_Headers1.AddField(Cf)
   Next Pg
'
'   For Each Et In m_ExpenseTypes2
'      Call Cf.SetFieldValue(2, "^", Et.EXPENSE_TYPE_NAME, ">", "E-" & Et.EXPENSE_TYPE_ID)
'      Call m_Headers1.AddField(Cf)
'   Next Et

   Call Cf.SetFieldValue(2, "^", "ค่าใช้จ่าย", ">")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(3, "^", "รวมต้นทุน", ">")
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
   
   Call LoadPartGroup(Nothing, m_PartGroup)
   Call LoadExpenseType(Nothing, m_ExpenseTypes1, "Y")
   Call LoadExpenseType(Nothing, m_ExpenseTypes2, "N")
   
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
   ToDate = " ถึง วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & " โรงเรือน " & EmptyToString(mcolParam("HOUSE_NAME"), "N/A")
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
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Lc As CLocation
Dim Pg As CPartItem
Dim InventoryBals As Collection
Dim CapitalBals As Collection
Dim HS As CLocation
Dim Cm As CCapitalMovement
Dim TempID As Long
Dim NewDate As Date
Dim LocationMaps As Collection
Dim MovementHouses As Collection
Dim LeftAmount As Long
Dim ParentUSed As Collection
Dim TempHS As CMovementItem
Dim Sum1 As Double
Dim TempKey As String
Dim II As CImportItem
Dim EI As CExportItem
Dim PigCount As Long
Dim Mi As CMovementItem
Dim SumToPigCount As Long
Dim PigBirths As Collection
Dim PigBirthCount As Long
Dim ExpenseRatios As Collection
Dim Er As CExpenseRatio
Dim HousePigBirths As Collection
Dim PigBirthInHouse As Long
Dim PigBalance As Double
Dim ImportPigs As Collection
Dim ExportPigs As Collection
Dim ImportPigParents As Collection
Dim ExportPigParents As Collection
Dim PigParentCount As Double
Dim ExpenseAmt As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------

   NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
   
   Set ParentUSed = New Collection
   Call LoadPigParentUseAmountEx2(Nothing, ParentUSed, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))

   Set MovementHouses = New Collection
   Call LoadLocation(Nothing, MovementHouses, 1, "", mcolParam("HOUSE_ID"))
'   Call LoadExpRatioLocation(Nothing, MovementHouses, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("HOUSE_ID"))
  
   Set PigBirths = New Collection
   Call LoadPigBirthAmount(Nothing, PigBirths, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   Set II = GetImportItem(PigBirths, "1")
   PigBirthCount = II.IMPORT_AMOUNT
   
   Set ExpenseRatios = New Collection
   'ไม่นับ ค่าเสื่อมพันธ์
   Call LoadHouseExpenseRatio(Nothing, ExpenseRatios, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "N")
   
   Set HousePigBirths = New Collection
   Call LoadHousePigBirthAmount(Nothing, HousePigBirths, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), "", mcolParam("HOUSE_ID"))
   
   Set ImportPigs = New Collection
   Set ExportPigs = New Collection
   If NewDate > 0 Then
      Call LoadPigImportByHouse(Nothing, ImportPigs, -1, NewDate, "", mcolParam("HOUSE_ID"))
      Call LoadPigExportByHouse(Nothing, ExportPigs, -1, NewDate, "", mcolParam("HOUSE_ID"))
   End If
    
   Set ImportPigParents = New Collection
   Set ExportPigParents = New Collection
   If NewDate > 0 Then
      Call LoadPigImportByHouse(Nothing, ImportPigParents, -1, NewDate, "", mcolParam("HOUSE_ID"), , "Y")
      Call LoadPigExportByHouse(Nothing, ExportPigParents, -1, NewDate, "", mcolParam("HOUSE_ID"), , , "Y")
   End If
         
   Set Rs = New ADODB.Recordset
   
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
      Set BodyCf = New CReportField

      For Each HS In MovementHouses
         I = I + 1
         For j = 1 To UBound(Total1)
            Total2(j) = 0
         Next j

            Set Er = GetExpenseRatio(ExpenseRatios, Trim(Str(HS.LOCATION_ID)))
            
            'เกิด
            Set II = GetImportItem(HousePigBirths, Trim(Str(HS.LOCATION_ID)))
            PigBirthInHouse = II.IMPORT_AMOUNT
            
            'ยกมา
            Set II = GetImportItem(ImportPigs, Trim(Str(HS.LOCATION_ID)))
            Set EI = GetExportItem(ExportPigs, Trim(Str(HS.LOCATION_ID)))
            PigBalance = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
            
            'หมูพ่อแม่พันธ์ปลายงวด
            Set II = GetImportItem(ImportPigParents, Trim(Str(HS.LOCATION_ID)))
            Set EI = GetExportItem(ExportPigParents, Trim(Str(HS.LOCATION_ID)))
            PigParentCount = II.IMPORT_AMOUNT - EI.EXPORT_AMOUNT
            
            j = 0
            Sum1 = 0
            Call m_Details1.ClearField
            SumToPigCount = 0
            For Each HeadCf In m_Headers1.Fields
               j = j + 1
               If j = 1 Then
                  TempStr = I
               ElseIf j = 2 Then
                  TempStr = HS.LOCATION_NAME
               ElseIf Mid(HeadCf.TempKey, 1, 1) = "I" Then
                  TempKey = HS.LOCATION_ID & "-" & Mid(HeadCf.TempKey, 3)
                  Set EI = GetExportItem(ParentUSed, TempKey)
                  TempStr = FormatNumber(EI.EXPORT_TOTAL_PRICE, "-")

                  Total1(j) = Total1(j) + EI.EXPORT_TOTAL_PRICE
                  Sum1 = Sum1 + EI.EXPORT_TOTAL_PRICE
'               ElseIf Mid(HeadCf.TempKey, 1, 1) = "E" Then
'                  TempKey = HS.LOCATION_ID & "-" & 0 & "-" & Mid(HeadCf.TempKey, 3)
'                  Set Mi = GetMovementItem(ParentUSed, TempKey)
'                  TempStr = FormatNumber(Mi.CAPITAL_AMOUNT, "-")
'
'                  Total1(j) = Total1(j) + Mi.CAPITAL_AMOUNT
'                  Sum1 = Sum1 + Mi.CAPITAL_AMOUNT
'                  SumToPigCount = SumToPigCount + Mi.TX_AMOUNT
               ElseIf j = m_Headers1.Fields.Count - 1 Then
                  ExpenseAmt = MyDiffEx(Er.RATIO_AMOUNT * PigParentCount, (PigBalance + PigBirthInHouse))
                  
                  TempStr = FormatNumber(ExpenseAmt, "-")
                  Total1(j) = Total1(j) + ExpenseAmt
                  Sum1 = Sum1 + ExpenseAmt
               ElseIf j = m_Headers1.Fields.Count Then
                  TempStr = FormatNumber(Sum1, "-")
                  Total1(j) = Total1(j) + Sum1
               End If
                                 
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            Next HeadCf
            
            If Sum1 <> 0 Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            Else
               I = I - 1
            End If
         Next HS
         
         Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      
         Call GenerateFooterEx("", "ต้นทุน/ตัว", "", Total1, m_Details1, PigBirthCount)
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Vsp = ""
         Vsp.FontBold = True
         Vsp = "จำนวนลูกเกิด : " & PigBirthCount
         Vsp.FontBold = False
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set ParentUSed = Nothing
   Set LocationMaps = Nothing
   Set MovementHouses = Nothing
   Set CapitalBals = Nothing
   Set InventoryBals = Nothing
   Set ExpenseRatios = Nothing
   Set HousePigBirths = Nothing
   Set ImportPigs = Nothing
    Set ExportPigs = Nothing
    Set ImportPigParents = Nothing
    Set ExportPigParents = Nothing
    
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
      Else
         TempStr = FormatNumber(Tot(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateFooterEx(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, PigCount As Long)
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
      Else
         TempStr = FormatNumber(MyDiffEx(Tot(j), PigCount))
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






