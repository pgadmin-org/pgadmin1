VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.2#0"; "HighlightBox.ocx"
Begin VB.Form frmReportAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Report"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmReportAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin HighlightBox.HBX txtSQL 
      Height          =   1320
      Left            =   90
      TabIndex        =   6
      ToolTipText     =   "Enter the SQL required to provide the data for the report. "
      Top             =   2385
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   2328
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SQL"
      AutoColour      =   -1  'True
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   945
      TabIndex        =   4
      ToolTipText     =   "Enter the name of the Author of the Report."
      Top             =   1035
      Width           =   3390
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseFile 
      Caption         =   "..."
      Height          =   285
      Left            =   4005
      TabIndex        =   1
      ToolTipText     =   "Browse for a file to import from"
      Top             =   90
      Width           =   345
   End
   Begin VB.CheckBox chkGroupTree 
      Alignment       =   1  'Right Justify
      Caption         =   "S&how 'Group Tree' in the Report Viewer?"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   4410
      Width           =   4245
   End
   Begin VB.CheckBox chkSequenceCache 
      Alignment       =   1  'Right Justify
      Caption         =   "Refresh the '&Sequence Cache' before execution?"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   4095
      Width           =   4245
   End
   Begin VB.CheckBox chkTableCache 
      Alignment       =   1  'Right Justify
      Caption         =   "Refresh the '&Table Cache' before execution?"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   3780
      Width           =   4245
   End
   Begin VB.TextBox txtCategory 
      Height          =   285
      Left            =   945
      TabIndex        =   3
      ToolTipText     =   "Enter a Category for the Report."
      Top             =   720
      Width           =   3390
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   945
      TabIndex        =   2
      ToolTipText     =   "Enter a name for the Report."
      Top             =   405
      Width           =   3390
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   945
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Enter the absolute filename for the report."
      Top             =   90
      Width           =   3030
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Report"
      Height          =   330
      Left            =   3015
      TabIndex        =   0
      Top             =   4680
      Width           =   1320
   End
   Begin HighlightBox.HBX txtDescription 
      Height          =   960
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Enter a description for the report."
      Top             =   1395
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   1693
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Description"
      AutoColour      =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Author"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   14
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   13
      Top             =   765
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   12
      Top             =   450
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Report File"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   135
      Width           =   765
   End
End
Attribute VB_Name = "frmReportAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin - PostgreSQL db Administration/Management for Win32
' Copyright (C) 1998 - 2001, Dave Page

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Option Explicit

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
Dim x As Integer
Dim fNum As Integer
Dim szData As String
  If txtFile.Text = "" Then
    MsgBox "You must select a report file!", vbExclamation, "Error"
    txtFile.SetFocus
    Exit Sub
  End If
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the report!", vbExclamation, "Error"
    txtName.SetFocus
    Exit Sub
  End If
  If txtCategory.Text = "" Then
    MsgBox "You must enter a report category!", vbExclamation, "Error"
    txtCategory.SetFocus
    Exit Sub
  End If
  If txtSQL.Text = "" Then
    MsgBox "You must select an SQL query!", vbExclamation, "Error"
    txtSQL.SetFocus
    Exit Sub
  End If
  ReDim Preserve rptList(UBound(rptList) + 1)
  rptList(UBound(rptList)).szName = txtName.Text
  rptList(UBound(rptList)).szCategory = txtCategory.Text
  rptList(UBound(rptList)).szFile = txtFile.Text
  rptList(UBound(rptList)).szSQL = txtSQL.Text
  rptList(UBound(rptList)).szAuthor = txtAuthor.Text
  rptList(UBound(rptList)).szDescription = txtDescription.Text
  If chkGroupTree.Value = 1 Then rptList(UBound(rptList)).bShowTree = True
  If chkTableCache.Value = 1 Then rptList(UBound(rptList)).bRefreshTables = True
  If chkSequenceCache.Value = 1 Then rptList(UBound(rptList)).bRefreshSequences = True
  StartMsg "Writing " & app.Path & "\Reports\Reports.dat"
  Dim fso As New FileSystemObject
  Dim fReports As File
  If fso.FileExists(app.Path & "\Reports\Reports.bak") Then
    Set fReports = fso.GetFile(app.Path & "\Reports\Reports.bak")
    fReports.Delete True
  End If
  Set fReports = fso.GetFile(app.Path & "\Reports\Reports.dat")
  fReports.Move (app.Path & "\Reports\Reports.bak")
  szData = ""
  fNum = FreeFile
  Open app.Path & "\Reports\Reports.dat" For Binary Access Write As #fNum
  For x = 1 To UBound(rptList)
    szData = szData & rptList(x).szName & Chr(253) & rptList(x).szCategory & Chr(253) & rptList(x).szFile & Chr(253) & rptList(x).szSQL & Chr(253) & rptList(x).szAuthor & Chr(253) & rptList(x).szDescription & Chr(253)
    If rptList(x).bShowTree = True Then
      szData = szData & "1" & Chr(253)
    Else
      szData = szData & "0" & Chr(253)
    End If
    If rptList(x).bRefreshTables = True Then
      szData = szData & "1" & Chr(253)
    Else
      szData = szData & "0" & Chr(253)
    End If
    If rptList(x).bRefreshSequences = True Then
      szData = szData & "1" & Chr(254)
    Else
      szData = szData & "0" & Chr(254)
    End If
  Next
  Put #fNum, , szData
  Close #fNum
  EndMsg
  frmReportManager.Refresh_List
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmReportAdd, cmdAdd_Click"
End Sub

Private Sub cmdBrowseFile_Click()
On Error GoTo Err_Handler
  With CommonDialog1
    .DialogTitle = "Select Crystal Report"
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .Filter = "Crystal Report Files (*.rpt)|*.rpt"
    .FileName = ""
    .CancelError = True
    .ShowOpen
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtFile.Text = CommonDialog1.FileName
  StartMsg "Examining " & CommonDialog1.FileName
  Dim rpt As craxdrt.Report
  Dim app As New craxdrt.Application
  Set rpt = app.OpenReport(txtFile.Text, crOpenReportByTempCopy)
  txtName.Text = rpt.ReportTitle
  txtCategory.Text = rpt.ReportSubject
  txtAuthor.Text = rpt.ReportAuthor
  txtDescription.Text = rpt.ReportComments
  txtSQL.Text = rpt.ReportTemplate
  Set rpt = Nothing
  Set app = Nothing
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 And Err.Number <> 32755 Then LogError Err, "frmReportAdd, cmdBrowseFile_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4530
  Me.Height = 5445
  txtSQL.Wordlist = TextColours
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportAdd, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtDescription.Minimise
  txtSQL.Minimise
  If Me.WindowState <> 1 Then
    If Me.Width < 4530 Then Me.Width = 4530
    If Me.Height < 5445 Then Me.Height = 5445
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportAdd, Form_Resize"
End Sub
