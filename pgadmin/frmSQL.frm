VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#4.0#0"; "HighlightBox.ocx"
Begin VB.Form frmSQL 
   Caption         =   "SQL"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   Icon            =   "frmSQL.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6555
   Begin HighlightBox.HBX txtSQL 
      Height          =   2805
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   4948
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordList        =   $"frmSQL.frx":030A
      Text            =   ""
      ScrollBars      =   2
      MultiLine       =   -1  'True
   End
   Begin VB.CommandButton cmdSQLWizard 
      Caption         =   "&Wizard"
      Height          =   330
      Left            =   1710
      TabIndex        =   2
      ToolTipText     =   "Run the SQL Wizard."
      Top             =   2835
      Width           =   810
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   330
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Load a query."
      Top             =   2835
      Width           =   810
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   330
      Left            =   855
      TabIndex        =   1
      ToolTipText     =   "Save the current query."
      Top             =   2835
      Width           =   795
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute to:"
      Height          =   330
      Left            =   2565
      TabIndex        =   3
      ToolTipText     =   "Execute the SQL query to the selected output option."
      Top             =   2835
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select SQL File"
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin vsAdoSelector.VS_AdoSelector vssExporters 
      Height          =   315
      Left            =   3645
      TabIndex        =   4
      ToolTipText     =   "Select where to execute the query to."
      Top             =   2835
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectorType    =   1
   End
End
Attribute VB_Name = "frmSQL"
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
Dim bDirty As Boolean
Public szTitle As String
Dim szFilename As String

Private Sub cmdExecute_Click()
On Error GoTo Err_Handler
Dim rsQuery As New Recordset
Dim szQuery As String
  If Len(txtSQL.Text) < 5 Then Exit Sub
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Recordset Viewer", ValString, vssExporters.Caption
  szQuery = Replace(txtSQL.Text, vbCrLf, " ")
  While InStr(1, szQuery, "  ") > 0
    szQuery = Replace(szQuery, "  ", " ")
  Wend
  LogMsg "Executing: " & szQuery
  If UCase(Mid(szQuery, 1, 6)) = "SELECT" Then
    StartMsg "Executing SQL Query..."
    Select Case vssExporters.Text
      Case "SC"
        Dim DataFormRO As New frmSQLOutput
        rsQuery.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
        Load DataFormRO
        DataFormRO.Display rsQuery
        DataFormRO.Show
        DataFormRO.ZOrder 0
      Case Else
        rsQuery.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
        LogMsg "Running Exporter: " & Exporters(CInt(vssExporters.Text)).Description & " v" & Exporters(CInt(vssExporters.Text)).Version
        Exporters(CInt(vssExporters.Text)).Export rsQuery
    End Select
    EndMsg
    Exit Sub
  End If
  StartMsg "Executing SQL Query..."
  gConnection.Execute szQuery
  LogQuery szQuery
  EndMsg
  MsgBox "Query Executed OK!", vbInformation
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmSQL, cmdExecute_Click"
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Err_Handler
Dim DataLine As String
Dim fNum As Integer
  If bDirty = True Then
    If MsgBox("This query has been edited - do you wish to save it?", vbQuestion + vbYesNo, "Save Query") = vbYes Then cmdSave_Click
  End If
  With CommonDialog1
    .DialogTitle = "Load SQL Query"
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .Filter = "SQL Scripts (*.sql)|*.sql|All Files (*.*)|*.*"
    .FileName = ""
    .CancelError = True
    .ShowOpen
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtSQL.Text = ""
  fNum = FreeFile
  LogMsg "Loading " & CommonDialog1.FileName
  Open CommonDialog1.FileName For Input As #fNum
  While Not EOF(fNum)
    Line Input #fNum, DataLine
    txtSQL.Text = txtSQL.Text & DataLine & vbCrLf
  Wend
  Close #fNum
  Get_Filename
  Me.Caption = szTitle & " (" & szFilename & ")"
  bDirty = False
  Exit Sub
Err_Handler: If Err.Number <> 0 And Err.Number <> 32755 Then LogError Err, "frmSQL, cmdLoad_Click"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
Dim DataLine As String
Dim fNum As Integer
  With CommonDialog1
    .DialogTitle = "Save SQL Query"
    .Filter = "SQL Scripts (*.sql)|*.sql"
    .CancelError = True
    .ShowSave
  End With
  If CommonDialog1.FileName = "" Then
    MsgBox "No filename specified - SQL query not saved.", vbExclamation, "Warning"
    Exit Sub
  End If
  If Dir(CommonDialog1.FileName) <> "" Then
    If MsgBox("File exists - overwrite?", vbYesNo + vbQuestion, "Overwrite File") = vbNo Then cmdSave_Click
  End If
  fNum = FreeFile
  LogMsg "Writing " & CommonDialog1.FileName
  Open CommonDialog1.FileName For Output As #fNum
  Print #fNum, txtSQL.Text
  Close #fNum
  Get_Filename
  Me.Caption = szTitle & " (" & szFilename & ")"
  bDirty = False
  Exit Sub
Err_Handler: If Err.Number <> 0 And Err.Number <> 32755 Then LogError Err, "frmSQL, cmdSave_Click"
End Sub

Private Sub cmdSQLWizard_Click()
On Error GoTo Err_Handler
Dim SQLWizard As New frmSQLWizard
  Load SQLWizard
  SQLWizard.Tag = Me.hwnd
  SQLWizard.Caption = SQLWizard.Caption & " (" & Me.Caption & ")"
  SQLWizard.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQL, cmdSQLWizard_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim X As Integer
  LogMsg "Loading Form: " & Me.Name
  vssExporters.DisplayList = "Screen;"
  vssExporters.IndexList = "SC;"
  On Error Resume Next
  For X = 0 To UBound(Exporters)
    vssExporters.DisplayList = vssExporters.DisplayList & Exporters(X).Description & ";"
    vssExporters.IndexList = vssExporters.IndexList & X & ";"
  Next
  On Error GoTo Err_Handler
  vssExporters.LoadList
  vssExporters.SelectItemText RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Recordset Viewer", "Read Only Screen (Fast)")
  txtSQL.Wordlist = TextColours
  bDirty = False
  Me.Height = 3600
  Me.Width = 6705
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQL, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 6705 Then Me.Width = 6705
      If Me.Height < 3600 Then Me.Height = 3600
    End If
    txtSQL.Width = Me.ScaleWidth
    txtSQL.Height = Me.ScaleHeight - cmdExecute.Height - 50
    cmdExecute.Top = Me.ScaleHeight - cmdExecute.Height
    cmdLoad.Top = cmdExecute.Top
    cmdSave.Top = cmdExecute.Top
    cmdSQLWizard.Top = cmdExecute.Top
    vssExporters.Top = cmdExecute.Top - ((cmdExecute.Height - vssExporters.Height) / 2)
    vssExporters.Left = Me.ScaleWidth - vssExporters.Width
    cmdExecute.Left = vssExporters.Left - cmdExecute.Width - 50
    vssExporters.Left = Me.ScaleWidth - vssExporters.Width
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQL, Form_Resize"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
  If bDirty = True Then
    If MsgBox("This query has been edited - do you wish to save it?", vbQuestion + vbYesNo, "Save Query") = vbYes Then cmdSave_Click
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQL, Form_Unload"
End Sub

Private Sub txtSQL_Change()
On Error GoTo Err_Handler
  If CommonDialog1.FileName = "" Then
    Me.Caption = szTitle & " (Edited)"
  Else
    Me.Caption = szTitle & " (" & szFilename & ") (Edited)"
  End If
  bDirty = True
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQL, txtSQL_Change"
End Sub

Private Sub Get_Filename()
On Error GoTo Err_Handler
Dim iSlashes As Integer
Dim iLastPos As Integer
Dim iCount As Integer
  iSlashes = CountChar(CommonDialog1.FileName, Asc("\"))
  If iSlashes = 0 Then
    szFilename = CommonDialog1.FileName
  Else
    For iCount = 1 To Len(CommonDialog1.FileName)
      If Mid(CommonDialog1.FileName, iCount, 1) = "\" Then iLastPos = iCount
    Next
    szFilename = Mid(CommonDialog1.FileName, iLastPos + 1)
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQL, Get_Filename"
End Sub

