VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Data"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   5835
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmImport.frx":030A
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   3
      Top             =   -45
      Width           =   465
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import Data"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4545
      TabIndex        =   2
      ToolTipText     =   "Run the defined import"
      Top             =   3915
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   330
      Left            =   2475
      TabIndex        =   0
      ToolTipText     =   "Move back a stage"
      Top             =   3915
      Width           =   960
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   3510
      TabIndex        =   1
      ToolTipText     =   "Move forward a stage"
      Top             =   3915
      Width           =   960
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   540
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Import File"
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin TabDlg.SSTab tabWizard 
      CausesValidation=   0   'False
      Height          =   3840
      Left            =   495
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   176
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmImport.frx":12F3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboTableoid"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstColumns"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSelect"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDeSelect"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtFile"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdBrowse"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSample"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboTables"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmImport.frx":130F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdDown"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdUp"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lstOColumns"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.ListBox lstOColumns 
         Height          =   3180
         Left            =   -72480
         TabIndex        =   37
         ToolTipText     =   "Lists the Columns that data will be imported to in order"
         Top             =   495
         Width           =   2220
      End
      Begin VB.CommandButton cmdUp 
         Height          =   540
         Left            =   -70230
         Picture         =   "frmImport.frx":132B
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Move the selected Column up the list"
         Top             =   495
         Width           =   435
      End
      Begin VB.CommandButton cmdDown 
         Height          =   540
         Left            =   -70230
         Picture         =   "frmImport.frx":176D
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Move the selected Column down the list"
         Top             =   3150
         Width           =   435
      End
      Begin VB.Frame Frame1 
         Caption         =   "Delimiter"
         Height          =   1125
         Left            =   -74895
         TabIndex        =   29
         Top             =   180
         Width           =   2325
         Begin VB.CheckBox chkTrailing 
            Alignment       =   1  'Right Justify
            Caption         =   "Expect trailing delimiter?"
            Height          =   225
            Left            =   105
            TabIndex        =   34
            ToolTipText     =   "Should the import wizard expect to find a trailing delimiter after the last Column?"
            Top             =   840
            Value           =   1  'Checked
            Width           =   2115
         End
         Begin VB.OptionButton optDelimiter 
            Alignment       =   1  'Right Justify
            Caption         =   "Character"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   33
            ToolTipText     =   "Specify a character to use as a delimiter"
            Top             =   240
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optDelimiter 
            Alignment       =   1  'Right Justify
            Caption         =   "Ascii Value"
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   32
            ToolTipText     =   "Specify a decimal Ascii value to use as a delimiter"
            Top             =   555
            Width           =   1275
         End
         Begin VB.TextBox txtDelimChar 
            Height          =   285
            Left            =   1575
            MaxLength       =   1
            TabIndex        =   31
            Text            =   ","
            ToolTipText     =   "Enter the delimiting character"
            Top             =   210
            Width           =   645
         End
         Begin VB.TextBox txtDelimAscii 
            Height          =   285
            Left            =   1575
            MaxLength       =   3
            TabIndex        =   30
            ToolTipText     =   "Enter the decimal Ascii value of the delimiting character"
            Top             =   525
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Quoting"
         Height          =   1110
         Left            =   -74910
         TabIndex        =   23
         Top             =   1395
         Width           =   2325
         Begin VB.OptionButton optQuote 
            Alignment       =   1  'Right Justify
            Caption         =   "None"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   28
            ToolTipText     =   "Specify that there is no character 'quoting' each column"
            Top             =   210
            Width           =   1275
         End
         Begin VB.OptionButton optQuote 
            Alignment       =   1  'Right Justify
            Caption         =   "Character"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   27
            ToolTipText     =   "Specify to use a character as a quote mark"
            Top             =   495
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optQuote 
            Alignment       =   1  'Right Justify
            Caption         =   "Ascii Value"
            Height          =   225
            Index           =   2
            Left            =   105
            TabIndex        =   26
            ToolTipText     =   "Specify a decimal Ascii value to use as a quote mark"
            Top             =   810
            Width           =   1275
         End
         Begin VB.TextBox txtQuoteChar 
            Height          =   285
            Left            =   1575
            MaxLength       =   1
            TabIndex        =   25
            Text            =   "'"
            ToolTipText     =   "Enter the quote mark character"
            Top             =   450
            Width           =   645
         End
         Begin VB.TextBox txtQuoteAscii 
            Height          =   285
            Left            =   1575
            MaxLength       =   3
            TabIndex        =   24
            ToolTipText     =   "Enter the decimal Ascii value to use as a quote mark"
            Top             =   765
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pick Value Mark Conversion"
         Height          =   1140
         Left            =   -74910
         TabIndex        =   17
         Top             =   2610
         Width           =   2325
         Begin VB.OptionButton optPick 
            Alignment       =   1  'Right Justify
            Caption         =   "None"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   22
            ToolTipText     =   "Specify not to search for Pick VMs  (Ascii 253)"
            Top             =   270
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optPick 
            Alignment       =   1  'Right Justify
            Caption         =   "Character"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   21
            ToolTipText     =   "Specify to convert Pick VMs to a character"
            Top             =   540
            Width           =   1275
         End
         Begin VB.OptionButton optPick 
            Alignment       =   1  'Right Justify
            Caption         =   "Ascii Value"
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   20
            ToolTipText     =   "Specify to convert Pick VMs to a decimal Ascii value"
            Top             =   855
            Width           =   1275
         End
         Begin VB.TextBox txtPickChar 
            Height          =   285
            Left            =   1575
            MaxLength       =   1
            TabIndex        =   19
            ToolTipText     =   "Enter a character to convert Pick VMs to"
            Top             =   495
            Width           =   645
         End
         Begin VB.TextBox txtPickAscii 
            Height          =   285
            Left            =   1575
            MaxLength       =   3
            TabIndex        =   18
            Text            =   "13"
            ToolTipText     =   "Enter a decimal Ascii value to convert Pick VMs to"
            Top             =   810
            Width           =   645
         End
      End
      Begin VB.ComboBox cboTables 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Select the table to import to"
         Top             =   1860
         Width           =   3615
      End
      Begin VB.TextBox txtSample 
         Height          =   1170
         Left            =   1575
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   11
         ToolTipText     =   "Displays a sample of the import file"
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   4815
         TabIndex        =   10
         ToolTipText     =   "Browse for a file to import from"
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1590
         TabIndex        =   9
         ToolTipText     =   "Enter a filename to import from"
         Top             =   180
         Width           =   3225
      End
      Begin VB.CommandButton cmdDeSelect 
         Caption         =   "&Deselect All"
         Height          =   330
         Left            =   225
         TabIndex        =   8
         ToolTipText     =   "Deselect all listed columns"
         Top             =   3195
         Width           =   1170
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select All"
         Height          =   330
         Left            =   225
         TabIndex        =   7
         ToolTipText     =   "Select all listed columns"
         Top             =   2700
         Width           =   1170
      End
      Begin VB.ListBox lstColumns 
         Height          =   1410
         Left            =   1590
         Style           =   1  'Checkbox
         TabIndex        =   6
         ToolTipText     =   "Select the Columns to import data to"
         Top             =   2280
         Width           =   3615
      End
      Begin VB.ComboBox cboTableoid 
         Height          =   315
         Left            =   2925
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1890
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Column Order"
         Height          =   225
         Index           =   1
         Left            =   -72525
         TabIndex        =   38
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label2 
         Caption         =   "Sample:"
         Height          =   225
         Left            =   180
         TabIndex        =   16
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Import File:"
         Height          =   225
         Left            =   225
         TabIndex        =   15
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label3 
         Caption         =   "Import to Table:"
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   1860
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "Import to Columns:"
         Height          =   225
         Left            =   225
         TabIndex        =   13
         Top             =   2280
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmImport"
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
Dim bButtonPress As Boolean
Dim bProgramPress As Boolean

Private Sub cboTables_Click()
On Error GoTo Err_Handler
Dim rsColumns As New Recordset
  If cboTables.ListIndex <> -1 Then
    cboTableoid.ListIndex = cboTables.ListIndex
    lstColumns.Clear
    StartMsg "Retrieving Column Names..."
    LogMsg "Executing: SELECT column_name FROM pgadmin_tables WHERE column_position > 0 AND table_oid = " & cboTableoid.Text & " ORDER BY column_position"
    rsColumns.Open "SELECT column_name FROM pgadmin_tables WHERE column_position > 0 AND table_oid = " & cboTableoid.Text & " ORDER BY column_position", gConnection, adOpenForwardOnly
    While Not rsColumns.EOF
      lstColumns.AddItem rsColumns!column_name
      rsColumns.MoveNext
    Wend
    Set rsColumns = Nothing
    EndMsg
  End If
  Exit Sub
Err_Handler:
  Set rsColumns = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmImport, cboTables_Click"
End Sub

Private Sub tabWizard_Click(PreviousTab As Integer)
On Error GoTo Err_Handler
  If bButtonPress = False And bProgramPress = False Then
    bProgramPress = True
    tabWizard.Tab = PreviousTab
  Else
    bProgramPress = False
  End If
  bButtonPress = False
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, tabWizard_Click"
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo Err_Handler
Dim DataLine As String
Dim X As Integer
Dim fNum As Integer
  With CommonDialog1
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .ShowOpen
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtFile.Text = CommonDialog1.FileName
  txtSample.Text = ""
  fNum = FreeFile
  Open txtFile.Text For Input As #fNum
  For X = 0 To 4
    If Not EOF(1) Then
      Line Input #fNum, DataLine
      txtSample.Text = txtSample.Text & DataLine & vbCrLf
    End If
  Next
  Close #fNum
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, cmdBrowse_Click"
End Sub

Private Sub cmdDeSelect_Click()
On Error GoTo Err_Handler
Dim X As Integer
  For X = 0 To lstColumns.ListCount - 1
    lstColumns.Selected(X) = False
  Next
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, cmdDeSelect_Click"
End Sub

Private Sub cmdDown_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstOColumns.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstOColumns.ListIndex = lstOColumns.ListCount - 1 Then
    MsgBox "This column is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstOColumns.List(lstOColumns.ListIndex + 1)
  lstOColumns.List(lstOColumns.ListIndex + 1) = lstOColumns.List(lstOColumns.ListIndex)
  lstOColumns.List(lstOColumns.ListIndex) = Temp
  lstOColumns.ListIndex = lstOColumns.ListIndex + 1
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, cmdDown_Click"
End Sub

Private Sub cmdImport_Click()
On Error GoTo Err_Handler
Dim X As Long
Dim rsColumns As New Recordset
Dim Columns As String
Dim ColCount As Integer
Dim QuoteChar As String
Dim DelimChar As String
Dim PickChar As String
Dim RawLine As String
Dim TempLine As String
Dim DataLine As String
Dim InsertStr As String
Dim TupleCount As Long
Dim Tuple() As Variant
Dim Fields() As String
Dim Flag As Boolean
Dim InField As Boolean
Dim Y As Long
Dim Temp As String
Dim Msg As String
Dim ErrCount As Long
Dim ADOErr As Boolean
Dim fNum1 As Integer
Dim fNum2 As Integer

  'Check the file details, and open it.
  
  If txtFile.Text = "" Then
    MsgBox "You must select a file to export to!", vbExclamation, "Error"
    Exit Sub
  End If
  If FileDateTime(txtFile.Text) = "" Then
    MsgBox "The specified file could not be read.", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Build the import column list
  
  ColCount = 0
  For X = 0 To lstOColumns.ListCount - 1
    ReDim Preserve Fields(ColCount + 1)
    Fields(ColCount) = lstOColumns.List(X)
    ColCount = ColCount + 1
  Next
  If ColCount = 0 Then
    MsgBox "You must select at least one column to import to!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Set up the quote, delim & Pick chars
  
  If optQuote(0).Value = True Then
    QuoteChar = ""
  ElseIf optQuote(1).Value = True Then
    QuoteChar = txtQuoteChar.Text
    If QuoteChar = "" Then
      MsgBox "You must enter a quoting character!", vbExclamation, "Error"
      Exit Sub
    End If
  ElseIf optQuote(2).Value = True Then
    QuoteChar = Chr(txtQuoteAscii.Text)
    If QuoteChar = "" Then
      MsgBox "You must enter a quoting character value!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  If optPick(0).Value = True Then
    PickChar = ""
  ElseIf optPick(1).Value = True Then
    PickChar = txtPickChar.Text
    If PickChar = "" Then
      MsgBox "You must enter a replacement Pick Value Mark character!", vbExclamation, "Error"
      Exit Sub
    End If
  ElseIf optPick(2).Value = True Then
    PickChar = Chr(txtPickAscii.Text)
    If PickChar = "" Then
      MsgBox "You must enter a replacement Pick Value Mark character value!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  If optDelimiter(0).Value = True Then
    DelimChar = txtDelimChar.Text
    If DelimChar = "" Then
      MsgBox "You must enter a delimiting character!", vbExclamation, "Error"
      Exit Sub
    End If
  ElseIf optDelimiter(1).Value = True Then
    DelimChar = Chr(txtDelimAscii.Text)
    If DelimChar = "" Then
      MsgBox "You must enter a delimiting character value!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  
  'Dimension the import array.
  
  ReDim Tuple(ColCount - 1)

  'Now enter a processing loop, looping though each import line until we hit EOF
  
  StartMsg "Importing Records..."
  TupleCount = 0
  ErrCount = 0
  ADOErr = False
  fNum1 = FreeFile
  Open txtFile.Text For Input As #fNum1
  
  Do While Not EOF(1)
    If Not EOF(1) Then
      Line Input #fNum1, RawLine
      Flag = False
      
      'Attempt to figure out if we have a complete tuple, if not, read the
      'next line and concatenate it to the first one, then check again.
      'Three ways to check...
      
      'Quoted
      
      If optQuote(0).Value = False Then
        Do While Flag <> True
          If CountChar(RawLine, Asc(QuoteChar)) > 2 * (UBound(Tuple) + 1) Then
            Msg = "Tuple:    " & TupleCount + 1 & vbCrLf & _
                  "Error:    Too many quote characters found." & vbCrLf & _
                  "Expected: " & 2 * (UBound(Tuple) + 1) & vbCrLf & _
                  "Found:    " & CountChar(RawLine, Asc(QuoteChar)) & vbCrLf & vbCrLf
            GoTo Log_Error
          End If
          If CountChar(RawLine, Asc(QuoteChar)) = 2 * (UBound(Tuple) + 1) Then
            Flag = True
          Else
            Line Input #fNum1, TempLine
            RawLine = RawLine & vbCrLf & TempLine
          End If
        Loop
      End If
      
      'Delimited
      
      If optQuote(0).Value = True And chkTrailing.Value = 0 Then
        Do While Flag <> True
          If CountChar(RawLine, Asc(DelimChar)) > UBound(Tuple) Then
            Msg = "Tuple:    " & TupleCount + 1 & vbCrLf & _
                  "Error:    Too many delimiter characters found." & vbCrLf & _
                  "Expected: " & UBound(Tuple) & vbCrLf & _
                  "Found:    " & CountChar(RawLine, Asc(DelimChar)) & vbCrLf & vbCrLf
            GoTo Log_Error
          End If
          If CountChar(RawLine, Asc(DelimChar)) = UBound(Tuple) Then
            Flag = True
          Else
            Line Input #fNum1, TempLine
            RawLine = RawLine & vbCrLf & TempLine
          End If
        Loop
      End If
      
      'Delimited + Trailing
      
      If optQuote(0).Value = True And chkTrailing.Value = 1 Then
        Do While Flag <> True
          If CountChar(RawLine, Asc(DelimChar)) > UBound(Tuple) + 1 Then
            Msg = "Tuple:    " & TupleCount + 1 & vbCrLf & _
                  "Error:    Too many delimiter characters found." & vbCrLf & _
                  "Expected: " & UBound(Tuple) + 1 & vbCrLf & _
                  "Found:    " & CountChar(RawLine, Asc(DelimChar)) & vbCrLf & vbCrLf
            GoTo Log_Error
          End If
          If CountChar(RawLine, Asc(DelimChar)) = UBound(Tuple) + 1 Then
            Flag = True
          Else
            Line Input #fNum1, TempLine
            RawLine = RawLine & vbCrLf & TempLine
          End If
        Loop
      End If
      
      'Process Quoted & Delimited
      
      Flag = False
      If optQuote(0).Value = False Then
        X = -1
        InField = False
        For Y = 1 To Len(RawLine)
          If Mid(RawLine, Y, 1) = QuoteChar Then
            If InField = True Then
              InField = False
            Else
              InField = True
              X = X + 1
            End If
          ElseIf InField = True And Mid(RawLine, Y, 1) <> QuoteChar Then
            Tuple(X) = Tuple(X) & Mid(RawLine, Y, 1)
          End If
        Next
      End If
      
      'Process Delimited
      
      If optQuote(0).Value = True Then
        X = 0
        For Y = 1 To Len(RawLine)
          If Mid(RawLine, Y, 1) = DelimChar Then
            X = X + 1
          Else
            If X <= UBound(Tuple) Then
             Tuple(X) = Tuple(X) & Mid(RawLine, Y, 1)
            End If
          End If
        Next
      End If
      
      'Replace Pick MVs
      
      If optPick(0).Value = False Then
        For X = 0 To UBound(Tuple)
          Temp = ""
          For Y = 1 To Len(Tuple(X))
            If Mid(Tuple(X), Y, 1) = Chr(253) Then
              Temp = Temp & PickChar
            Else
              Temp = Temp & Mid(Tuple(X), Y, 1)
            End If
          Next
          Tuple(X) = Temp
        Next
      End If
      
      'Build SQL and Execute
      
      DataLine = ""
      Columns = ""
      For X = 0 To UBound(Tuple)
        If Tuple(X) <> "" Then
          DataLine = DataLine & "'" & Tuple(X) & "', "
          Columns = Columns & QUOTE & Fields(X) & QUOTE & ", "
        End If
      Next
      DataLine = Mid(DataLine, 1, Len(DataLine) - 2)
      Columns = Mid(Columns, 1, Len(Columns) - 2)
      InsertStr = "INSERT INTO " & QUOTE & cboTables.Text & QUOTE & " (" & Columns & ") VALUES (" & DataLine & ")"
      LogMsg "Executing: " & InsertStr
      gConnection.Execute InsertStr
      TupleCount = TupleCount + 1
Proc_Next:
      For X = 0 To UBound(Tuple)
        Tuple(X) = ""
      Next
      fMainForm.StatusBar1.Panels(1).Text = "Inserting Data - " & TupleCount & " Records"
      fMainForm.StatusBar1.Refresh
    End If
  Loop
  Close #fNum1
  EndMsg
  If ErrCount > 0 Then
      MsgBox "Data import complete. " & TupleCount & " records were imported from " & txtFile.Text & _
             vbCrLf & ErrCount & " Errors occured and were logged to " & LogFile, _
             vbExclamation, "Data Import"
      Close #fNum2
      Set rsColumns = Nothing
  Else
      MsgBox "Data import complete. " & TupleCount & " records were imported from " & txtFile.Text, _
             vbInformation, "Data Import"
      Set rsColumns = Nothing
      Unload Me
  End If

  Exit Sub
  
Err_Handler:
  Msg = "Tuple:    " & TupleCount + 1 & vbCrLf & _
        "Error:    " & Err.Number & " - " & Err.Description & vbCrLf
  ADOErr = True
Log_Error:
  fNum2 = FreeFile
  Open LogFile For Append As #fNum2
  Print #fNum2, Msg & "Data:     " & RawLine & vbCrLf
  Close #fNum2
  ErrCount = ErrCount + 1
  If ADOErr = True Then
    ADOErr = False
    Resume Proc_Next
  Else
    GoTo Proc_Next
  End If
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler
Dim X As Integer
  bButtonPress = True
  If tabWizard.Tab = 0 Then
  
    'Transfer selected columns
    
    lstOColumns.Clear
    For X = 0 To lstColumns.ListCount - 1
      If lstColumns.Selected(X) = True Then
        lstOColumns.AddItem lstColumns.List(X)
      End If
    Next
    tabWizard.Tab = 1
    cmdImport.Enabled = True
    cmdNext.Enabled = False
    cmdPrevious.Enabled = True
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, cmdNext_Click"
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Err_Handler
  bButtonPress = True
  If tabWizard.Tab = 1 Then
    cmdImport.Enabled = False
    cmdNext.Enabled = True
    cmdPrevious.Enabled = False
    tabWizard.Tab = 0
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, cmdPrevious_Click"
End Sub

Private Sub cmdSelect_Click()
On Error GoTo Err_Handler
Dim X As Integer
  For X = 0 To lstColumns.ListCount - 1
    lstColumns.Selected(X) = True
  Next
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, cmdSelect_Click"
End Sub

Private Sub cmdUp_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstOColumns.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstOColumns.ListIndex = 0 Then
    MsgBox "This column is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstOColumns.List(lstOColumns.ListIndex - 1)
  lstOColumns.List(lstOColumns.ListIndex - 1) = lstOColumns.List(lstOColumns.ListIndex)
  lstOColumns.List(lstOColumns.ListIndex) = Temp
  lstOColumns.ListIndex = lstOColumns.ListIndex - 1
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmImport, cmdUp_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 5925 Then Me.Width = 5925
    If Me.Height < 4665 Then Me.Height = 4665
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmExport, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsTables As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 4665
  Me.Width = 5925
  cboTables.Clear
  cboTableoid.Clear
  StartMsg "Retrieving Table Names..."
  LogMsg "Executing: SELECT DISTINCT ON(table_name) table_oid, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
  rsTables.Open "SELECT DISTINCT ON(table_name) table_oid, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name", gConnection, adOpenForwardOnly
  While Not rsTables.EOF
    cboTables.AddItem rsTables!table_name
    cboTableoid.AddItem rsTables!table_oid
    rsTables.MoveNext
  Wend
  If rsTables.State <> adStateClosed Then rsTables.Close
  EndMsg
  tabWizard.Tab = 0
  cmdPrevious.Enabled = False
  Set rsTables = Nothing
  Exit Sub
Err_Handler:
  Set rsTables = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmImport, Form_Load"
End Sub
