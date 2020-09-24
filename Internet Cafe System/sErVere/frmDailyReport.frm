VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDailyReport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4380
   ClientLeft      =   3765
   ClientTop       =   1890
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "frmDailyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   615
      Left            =   1680
      Picture         =   "frmDailyReport.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   840
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".xls"
      DialogTitle     =   "Export"
      Filter          =   "Microsoft Excel 2K (*.xls)|*.xls|"
      Flags           =   4
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   3120
      Picture         =   "frmDailyReport.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin MSACAL.Calendar Cal 
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4095
      _Version        =   524288
      _ExtentX        =   7223
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2004
      Month           =   9
      Day             =   3
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00000000&
      Caption         =   " DAILY REPORT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmDailyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdExport_Click()
On Error GoTo Monger
  cmdExport.Enabled = False
  cmdPrint.Enabled = False
  Dialog.FileName = App.Path & "\" & Format(Cal.Value, "mmmm") & "_" & Trim(Str(Cal.Day)) & "_" & Trim(Str(Cal.Year)) & " - " & Format(Cal.Value, "dddd")
  Export = True
  Dialog.ShowSave
  If Dialog.FileName = "" Then GoTo Monger
  SavePath = Dialog.FileName
    
  Screen.MousePointer = vbHourglass
  If Excel_Daily = True Then
    oWB.Close 0
    Set oXL = Nothing
    Set oSheet = Nothing
    Unload Me
    MsgBox "The report was successfully exported in excel format!"
  End If
Monger:
  cmdExport.Enabled = True
  cmdPrint.Enabled = True
  Screen.MousePointer = vbNormal
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Monger
  
  cmdExport.Enabled = False
  cmdPrint.Enabled = False
  Export = False
  Screen.MousePointer = vbHourglass
  If Excel_Daily = True Then
    oWB.Close 0
    oXL.Quit
    Unload Me
  End If

Monger:
  cmdExport.Enabled = True
  cmdPrint.Enabled = True
  Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
  Set oXL = CreateObject("Excel.Application")
  oXL.Visible = False
  Cal.Value = Format(Now, "ddddd")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set oSheet = Nothing
  Set oWB = Nothing
  Set oXL = Nothing
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me
End Sub
