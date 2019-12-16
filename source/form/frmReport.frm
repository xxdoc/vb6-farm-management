VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   8520
   ClientLeft      =   1740
   ClientTop       =   555
   ClientWidth     =   11910
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   7890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   2280
      TabIndex        =   5
      Top             =   8010
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VSPrinter7LibCtl.VSPrinter VSPrinter1 
      Height          =   7185
      Left            =   0
      TabIndex        =   2
      Top             =   690
      Width           =   11895
      _cx             =   20981
      _cy             =   12674
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   40.2462121212121
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   7
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   11895
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11625
      End
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   615
      Left            =   7740
      TabIndex        =   3
      Top             =   7890
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   131073
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   9840
      TabIndex        =   4
      Top             =   7890
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   131073
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmReport"

Private HasActivate As Boolean
Public HeaderText As String
Public ReportID As String
Public ReportObject As CReportInterface
Public OKClick As Boolean
Private m_ErrorFlag As Boolean
Public ClassName As String

Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

'VSPrinter1.PrintDoc
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("พิมพ์ไปเครื่องพิมพ์", "-", "บันทึกไปที่ไฟล์")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      m_ErrorFlag = True
      VSPrinter1.PrintDoc (True)
      If m_ErrorFlag Then 'Error
         glbErrorLog.LocalErrorMsg = "โปรแกรมได้ทำการพิมพ์รายงานเสร็จสิ้นแล้ว"
         glbErrorLog.ShowUserError
         Exit Sub
      Else
         glbErrorLog.LocalErrorMsg = "โปรแกรมได้ทำการพิมพ์รายงานเสร็จสิ้นแล้ว"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   ElseIf lMenuChosen = 3 Then
      CommonDialog1.Filter = "Save Files (*.html, *.htm)|*.html;*.htm;"
      CommonDialog1.DialogTitle = "Select access file to import"
      CommonDialog1.ShowSave
      If CommonDialog1.FileName = "" Then
         Exit Sub
      End If
      
      Call FileCopy(glbParameterObj.ReportFile, CommonDialog1.FileName)
   End If
   
   OKClick = True
   Unload Me
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   VSPrinter1.SaveDoc ("C:\xxx.rtf")
End Sub

Private Sub Form_Activate()
   If Not HasActivate Then
      HasActivate = True
      Me.Refresh
'      DoEvents
      
      Call EnableForm(Me, False)
      Me.Refresh
      Set ReportObject.VsPrint = VSPrinter1
      If Not ReportObject.Preview Then
         glbErrorLog.LocalErrorMsg = ReportObject.ErrorMsg
         glbErrorLog.ShowUserError
      End If
      Call EnableForm(Me, True)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = ClassName
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

   Me.BackColor = GLB_FORM_COLOR
   VSPrinter1.NavBarColor = GLB_FORM_COLOR
   VSPrinter1.PaperSize = pprA4
   Frame2.BackColor = GLB_HEAD_COLOR
   lblCaption.BackColor = GLB_HEAD_COLOR

   HasActivate = False
   m_ErrorFlag = False

   Me.Caption = HeaderText
   Call InitFormHeader(lblCaption, HeaderText)
   Call InitFormHeader(lblCaption, HeaderText)

   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"

   Call InitMainButton(cmdPrint, "พิมพ์ (F10)")
   Call InitMainButton(cmdExit, "ออก (ESC)")

   Call EnableForm(Me, True)
   Exit Sub

ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   With VSPrinter1
      .Move .Left, .Top, ScaleWidth - .Left * 2, ScaleHeight - .Top - .Left - 650
      cmdPrint.Top = .Top + ScaleHeight - .Top - .Left - 650
      cmdPrint.Left = .Left + ScaleWidth - .Left * 2 - cmdPrint.Width - cmdExit.Width
      cmdExit.Top = .Top + ScaleHeight - .Top - .Left - 650
      cmdExit.Left = .Left + ScaleWidth - .Left * 2 - cmdExit.Width
      Frame2.Width = ScaleWidth
      lblCaption.Width = ScaleWidth
      .ZoomMode = zmPageWidth
   End With
       
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ReportObject = Nothing
   Unload Me
End Sub

Private Sub VSPrinter1_Error()
   m_ErrorFlag = False
End Sub
