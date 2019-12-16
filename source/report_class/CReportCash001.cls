VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCash001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCash001"
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
Private m_RcpCndnItems As Collection
Private m_CashBals As Collection
Private m_InvoiceBals As Collection
Private m_Accounts As Collection
Private m_AccountTxs As Collection

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
   Set m_RcpCndnItems = New Collection
   Set m_CashBals = New Collection
   Set m_InvoiceBals = New Collection
   Set m_Accounts = New Collection
   Set m_AccountTxs = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_RcpCndnItems = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_CashBals = Nothing
   Set m_InvoiceBals = Nothing
   Set m_Accounts = Nothing
   Set m_AccountTxs = Nothing
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
Dim j As Long
Dim TempStr As String
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
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "รหัสบัญชี", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(5, "^", "ชื่อบัญชี", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ธนาคาร", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "สาขาธนาคาร", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(1, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "เลขที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วันที่เอกสาร", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับขายสด/ฝาก", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับชำระ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จ่าย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ค่าธรรมเนียม", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "หมายเหตุ", ">")
   Call m_Headers1.AddField(Cf)

   Set Cf = Nothing
            
  Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
            
  Call m_Headers1_1.GetString(1, TempStr1, TempStr1_1)
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
Dim Ct As CCashTran
Dim NewDate As Date

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   If mcolParam("FROM_DATE") > 0 Then
      NewDate = DateAdd("D", -1, mcolParam("FROM_DATE"))

      Set Ct = New CCashTran
      Call Ct.SetFieldValue("FROM_DATE", -1)
      Call Ct.SetFieldValue("TO_DATE", NewDate)
      Call Ct.SetFieldValue("BANK_ACCOUNT", mcolParam("BANK_ACCOUNT"))
      Call LoadSumCashTrnAmount(Ct, Nothing, m_CashBals)
      Set Ct = Nothing
   End If

   Set Ct = New CCashTran
   Call Ct.SetFieldValue("FROM_DATE", -1)
   Call Ct.SetFieldValue("TO_DATE", -1)
   Call Ct.SetFieldValue("BANK_ACCOUNT", mcolParam("BANK_ACCOUNT"))
   Call LoadBankAccountInCashTrn(Ct, Nothing, m_Accounts)
   Set Ct = Nothing

   Set Ct = New CCashTran
   Call Ct.SetFieldValue("FROM_DATE", mcolParam("FROM_DATE"))
   Call Ct.SetFieldValue("TO_DATE", mcolParam("TO_DATE"))
   Call Ct.SetFieldValue("BANK_ACCOUNT", mcolParam("BANK_ACCOUNT"))
   Call LoadBankAccountInCashTrn(Ct, Nothing, m_AccountTxs)
   Set Ct = Nothing
   
   FromDate = "จาก " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = " ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
                                          
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & " , " & glbEnterPrise.BRANCH_NAME & _
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

Private Sub GenerateSubHeader(Ct As CCashTran, TOT() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   TempStr = ""
   TempWidth = m_Headers0.Fields(1).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)
   
   TempStr = Ct.GetFieldValue("ACCOUNT_NO")
   TempWidth = m_Headers0.Fields(2).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

   TempStr = Ct.GetFieldValue("ACCOUNT_NAME")
   TempWidth = m_Headers0.Fields(3).Width
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "<")
   Call Details.AddField(BodyCf)

   TempStr = Ct.GetFieldValue("BANK_NAME")
   TempWidth = m_Headers0.Fields(4).Width
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "<")
   Call Details.AddField(BodyCf)

   TempStr = Ct.GetFieldValue("BRANCH_NAME")
   TempWidth = m_Headers0.Fields(5).Width
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "<")
   Call Details.AddField(BodyCf)

   TempStr = ""
   TempWidth = m_Headers0.Fields(6).Width
   Call BodyCf.SetFieldValue(TempWidth, "^", TempStr, "^")
   Call Details.AddField(BodyCf)

   Set BodyCf = Nothing
End Sub

Private Function GetBalanceAmount(Ct As CCashTran) As Double
Dim Ct1 As CCashTran
Dim Ct2 As CCashTran

   Set Ct1 = GetCashTran(m_CashBals, Ct.GetFieldValue("BANK_ACCOUNT") & "-" & "I")
   Set Ct2 = GetCashTran(m_CashBals, Ct.GetFieldValue("BANK_ACCOUNT") & "-" & "E")
   
   GetBalanceAmount = Ct1.GetFieldValue("NET_AMOUNT") - Ct2.GetFieldValue("NET_AMOUNT")
End Function

Private Function IsIn(TempID As Long, Col As Collection) As Boolean
Dim Bd As CCashTran

'IsIn = False
'Exit Function

   IsIn = False
   For Each Bd In Col
      If Bd.GetFieldValue("BANK_ACCOUNT") = TempID Then
         IsIn = True
         Exit Function
      End If
   Next Bd
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
Dim m_Cheque As CCheque
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim OldID As Long
Dim Ct As CCashTran
Dim TempCt As CCashTran
Dim k As Long

Dim ShowFlag As Boolean
Dim BalAmount As Double
Dim CurAmount As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
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
      k = 0
      Set BodyCf = New CReportField

      Set Ct = New CCashTran

      Call Ct.SetFieldValue("CASH_TRAN_ID", -1)
      Call Ct.SetFieldValue("FROM_DATE", mcolParam("FROM_DATE"))
      Call Ct.SetFieldValue("TO_DATE", mcolParam("TO_DATE"))
      Call Ct.SetFieldValue("ORDER_BY", 3)
      Call Ct.SetFieldValue("ORDER_TYPE", 1)

       Call Ct.QueryData(1, Rs, iCount)

      For Each TempCt In m_Accounts
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j

         BalAmount = GetBalanceAmount(TempCt)
         ShowFlag = False
         'แสดงถ้ามียอดยกมา หรือ มีการเคลื่อไหว
         If (BalAmount <> 0) Or (IsIn(TempCt.GetFieldValue("BANK_ACCOUNT"), m_AccountTxs)) Then
            ShowFlag = True
         End If
         
         If ShowFlag Then
            Call GenerateSubHeader(TempCt, Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
               
            Call GenerateBalance("ยกมา", "", "", BalAmount, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
         Total1(4) = Total1(4) + BalAmount
         Total2(4) = Total2(4) + BalAmount
         CurAmount = BalAmount

         If IsIn(TempCt.GetFieldValue("BANK_ACCOUNT"), m_AccountTxs) Then
            If Not Rs.EOF Then
               Call Ct.PopulateFromRS(1, Rs)
            End If
            
            While (Not Rs.EOF) And (TempCt.GetFieldValue("BANK_ACCOUNT") = Ct.GetFieldValue("BANK_ACCOUNT"))
               I = I + 1
               Call m_Details1.ClearField
               
               k = 0
               For k = 1 To 10
                  Set HeadCf = m_Headers1.Fields(k)
                  If k = 1 Then
                     TempStr = I
                  ElseIf k = 2 Then
                     TempStr = Ct.GetFieldValue("TX_NO")
                  ElseIf k = 3 Then
                     TempStr = DateToStringExtEx2(Ct.GetFieldValue("TX_DATE"))
                  ElseIf k = 4 Then
                     TempStr = FormatNumber(0)
                  ElseIf k = 5 Then
                     'ขายสด หรือ นำฝาก
                     If (Ct.GetFieldValue("TX_TYPE") = "I") And (Ct.GetFieldValue("RECEIPT_TYPE") = 1 Or Ct.GetFieldValue("CASH_DOC_ID") > 0) Then
                        TempStr = FormatNumber(Ct.GetFieldValue("AMOUNT"))
                        Total1(k) = Total1(k) + Ct.GetFieldValue("AMOUNT")
                        Total2(k) = Total2(k) + Ct.GetFieldValue("AMOUNT")
                        CurAmount = CurAmount + Ct.GetFieldValue("AMOUNT")
                     Else
                        TempStr = FormatNumber(0)
                     End If
                  ElseIf k = 6 Then
                     'รับชำระ
                     If (Ct.GetFieldValue("TX_TYPE") = "I") And (Ct.GetFieldValue("RECEIPT_TYPE") = 3) Then
                        TempStr = FormatNumber(Ct.GetFieldValue("AMOUNT"))
                        Total1(k) = Total1(k) + Ct.GetFieldValue("AMOUNT")
                        Total2(k) = Total2(k) + Ct.GetFieldValue("AMOUNT")
                        CurAmount = CurAmount + Ct.GetFieldValue("AMOUNT")
                     Else
                        TempStr = FormatNumber(0)
                     End If
                  ElseIf k = 7 Then
                     If Ct.GetFieldValue("TX_TYPE") = "E" Then
                        TempStr = FormatNumber(Ct.GetFieldValue("AMOUNT"))
                        Total1(k) = Total1(k) + Ct.GetFieldValue("AMOUNT")
                        Total2(k) = Total2(k) + Ct.GetFieldValue("AMOUNT")
                        CurAmount = CurAmount - Ct.GetFieldValue("AMOUNT")
                     Else
                        TempStr = FormatNumber(0)
                     End If
                  ElseIf k = 8 Then
                     TempStr = FormatNumber(Ct.GetFieldValue("FEE_AMOUNT"))
                     Total1(k) = Total1(k) + Ct.GetFieldValue("FEE_AMOUNT")
                     Total2(k) = Total2(k) + Ct.GetFieldValue("FEE_AMOUNT")
                     CurAmount = CurAmount - Ct.GetFieldValue("FEE_AMOUNT")
                  ElseIf k = 9 Then
                     TempStr = FormatNumber(CurAmount)
                     Total1(k) = CurAmount
                     Total2(k) = Total2(k) + CurAmount
                  ElseIf k = 10 Then
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               Next k
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               '===

               Call Rs.MoveNext
               Call Ct.PopulateFromRS(1, Rs)
            Wend
         End If

         If ShowFlag Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         End If
      Next TempCt

      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Set Ct = Nothing
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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, TOT() As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = ""
      ElseIf j = 9 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(TOT(j))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub GenerateBalance(Txt1 As String, Txt2 As String, Txt3 As String, BalAmount As Double, Details As CFieldList, DateCount As Long)
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
         TempStr = ""
      ElseIf j = 4 Then
         TempStr = FormatNumber(BalAmount)
      ElseIf j = 8 Then
         TempStr = FormatNumber(BalAmount)
      ElseIf j = 9 Then
         TempStr = ""
      Else
         TempStr = FormatNumber(0)
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
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
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx2(mdteDate), 8))
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






