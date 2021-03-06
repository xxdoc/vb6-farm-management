VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSell001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSell001"
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

Private m_PigStatus As Collection
Private m_ExpenseTypes2 As Collection
Private m_PartGroup As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_PigStatus = New Collection
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
   
   Set m_PigStatus = Nothing
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
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Amt As Double
Dim iCount As Long
Dim Et As CSGroupItem

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
   
   Call Cf.SetFieldValue(1.5, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "", "<")
   Call m_Headers1.AddField(Cf)
      
   For Each Et In m_PigStatus
      If Et.SELECT_FLAG = "Y" Then
         Call Cf.SetFieldValue(6, "^", Et.STATUS_NAME, ">", "E-" & Et.STATUS_ID)
         Call m_Headers1.AddField(Cf)
      End If
   Next Et

   Call Cf.SetFieldValue(8, "^", "��� �", ">")
   Call m_Headers1.AddField(Cf)
      
   '===
   Call Cf.SetFieldValue(1.5, "^", "NO.", "^")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�Ţ����͡���", "^")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "�ѹ����͡���", "^")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "�����١���", "<")
   Call m_Headers2.AddField(Cf)
   
   For Each Et In m_PigStatus
      If Et.SELECT_FLAG = "Y" Then
         Call Cf.SetFieldValue(2, "^", "���", ">", "E-" & Et.STATUS_ID)
         Call m_Headers2.AddField(Cf)
         
         Call Cf.SetFieldValue(2, "^", "���˹ѡ", ">", "E-" & Et.STATUS_ID)
         Call m_Headers2.AddField(Cf)
      
         Call Cf.SetFieldValue(2, "^", "��Ť��", ">", "E-" & Et.STATUS_ID)
         Call m_Headers2.AddField(Cf)
      End If
   Next Et

   Call Cf.SetFieldValue(2, "^", "��Դ", "^")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "���", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "���˹ѡ", ">")
   Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��Ť��", ">")
   Call m_Headers2.AddField(Cf)
   '===
   
   Set Cf = Nothing
                  
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
  Call m_Headers2.GetString(1, TempStr1, TempStr1_1)
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
   
   Call LoadPartGroup(Nothing, m_PartGroup)
   Call LoadPigStatusInGroup(Nothing, m_PigStatus, mcolParam("GROUP_STATUS_ID"))
   
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
   
   FromDate = "�ҡ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " �֧ �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
                                       vbCrLf & mcolParam("REPORT_NAME") & _
                                       vbCrLf & FromDate & ToDate & " �ç���͹ " & EmptyToString(mcolParam("HOUSE_NAME"), "N/A")
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
Dim Di As CDoItem
Dim Sgi As CSGroupItem
Dim FoundFlag As Boolean
Dim Ps As CProductStatus
Dim ProductStatus As Collection
Dim k As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------

   Set ProductStatus = New Collection
   Call LoadProductStatus(Nothing, ProductStatus)
   
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
   
   Set Di = New CDoItem
   Di.DO_ITEM_ID = -1
   Di.FROM_DATE = mcolParam("FROM_DATE")
   Di.TO_DATE = mcolParam("TO_DATE")
   Di.LOCATION_ID = mcolParam("HOUSE_ID")
   Di.DOCUMENT_TYPE = -1 'mcolParam("BILL_SUBTYPE")
   Di.DocTypeSet = BillingDocType2Set(mcolParam("BILL_SUBTYPE"))
   Di.OrderBy = mcolParam("ORDER_BY")
   Di.OrderType = mcolParam("ORDER_TYPE")
   Call Di.QueryData(4, Rs, iCount)
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc

      I = 0
      Set BodyCf = New CReportField

      While Not Rs.EOF
         Call Di.PopulateFromRS(4, Rs)
         
         I = I + 1
         For j = 1 To UBound(Total1)
            Total2(j) = 0
         Next j

         j = 0
         Call m_Details1.ClearField

         For j = 1 To 4 'm_Headers2.Fields.Count
            Set HeadCf = m_Headers2.Fields(j)
            If j = 1 Then
               TempStr = I
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            ElseIf j = 2 Then
               TempStr = Di.DOCUMENT_NO
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            ElseIf j = 3 Then
               TempStr = DateToStringExtEx2(Di.DOCUMENT_DATE)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            ElseIf j = 4 Then
               TempStr = Di.CUSTOMER_NAME
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            End If
         Next j
         
         FoundFlag = False
         k = 5
         For Each Sgi In m_PigStatus
            If Sgi.SELECT_FLAG = "Y" Then
               If Sgi.STATUS_ID = Di.PIG_STATUS Then
                  TempStr = FormatNumber(Di.ITEM_AMOUNT, "-")
                  Call BodyCf.SetFieldValue(2, ">", TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + Di.ITEM_AMOUNT
                  k = k + 1
                  
                  TempStr = FormatNumber(Di.TOTAL_WEIGHT, "-")
                  Call BodyCf.SetFieldValue(2, ">", TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + Di.TOTAL_WEIGHT
                  k = k + 1
               
                  TempStr = FormatNumber(Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT, "-")
                  Call BodyCf.SetFieldValue(2, ">", TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT
                  k = k + 1
                  FoundFlag = True
               Else
                  TempStr = FormatNumber(0, "-")
                  Call BodyCf.SetFieldValue(2, ">", TempStr)
                  Call m_Details1.AddField(BodyCf)
                  k = k + 1
                  
                  TempStr = FormatNumber(0, "-")
                  Call BodyCf.SetFieldValue(2, ">", TempStr)
                  Call m_Details1.AddField(BodyCf)
                  k = k + 1
                  
                  TempStr = FormatNumber(0, "-")
                  Call BodyCf.SetFieldValue(2, ">", TempStr)
                  Call m_Details1.AddField(BodyCf)
                  k = k + 1
               End If
            End If
         Next Sgi
         
         If Not FoundFlag Then
'If Di.PIG_STATUS <= 0 Then
'''debug.print
'End If
            Set Ps = ProductStatus(Trim(Str(Di.PIG_STATUS)))
            TempStr = Ps.PRODUCT_STATUS_NAME
            Call BodyCf.SetFieldValue(2, "^", TempStr)
            Call m_Details1.AddField(BodyCf)
            k = k + 1
            
            TempStr = FormatNumber(Di.ITEM_AMOUNT, "-")
            Call BodyCf.SetFieldValue(2, ">", TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Di.ITEM_AMOUNT
            k = k + 1
            
            TempStr = FormatNumber(Di.TOTAL_WEIGHT, "-")
            Call BodyCf.SetFieldValue(2, ">", TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Di.TOTAL_WEIGHT
            k = k + 1
         
            TempStr = FormatNumber(Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT, "-")
            Call BodyCf.SetFieldValue(2, ">", TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Di.TOTAL_PRICE - Di.DISCOUNT_AMOUNT
            k = k + 1
            
            FoundFlag = True
         Else
            TempStr = ""
            Call BodyCf.SetFieldValue(2, "^", TempStr)
            Call m_Details1.AddField(BodyCf)
            
            TempStr = FormatNumber(0, "-")
            Call BodyCf.SetFieldValue(2, ">", TempStr)
            Call m_Details1.AddField(BodyCf)
         
            TempStr = FormatNumber(0, "-")
            Call BodyCf.SetFieldValue(2, ">", TempStr)
            Call m_Details1.AddField(BodyCf)
         
            TempStr = FormatNumber(0, "-")
            Call BodyCf.SetFieldValue(2, ">", TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Rs.MoveNext
      Wend
         
      Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
         
      Set BodyCf = Nothing
   Vsp.EndDoc
    Call CloseExportFile(Vsp)
    
    Set Di = Nothing

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing

   Set ProductStatus = Nothing
   
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
   For Each HeadCf In m_Headers2.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf j = 4 Then
         TempStr = ""
      ElseIf j = m_Headers2.Fields.Count - 3 Then
         TempStr = ""
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
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




