VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Begin VB.Form frmAddTable 
   Caption         =   "Table Designer"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   Icon            =   "frmAddTable.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7020
   Begin VB.ListBox lstConstraints 
      Height          =   645
      ItemData        =   "frmAddTable.frx":030A
      Left            =   3015
      List            =   "frmAddTable.frx":030C
      TabIndex        =   19
      ToolTipText     =   "Displays the list of Constraints"
      Top             =   3105
      Width           =   3960
   End
   Begin VB.Frame fraConstraints 
      Caption         =   "Constraints"
      Height          =   780
      Left            =   45
      TabIndex        =   34
      Top             =   2970
      Width           =   2910
      Begin VB.CommandButton cmdUpConstraint 
         Caption         =   "Up"
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         ToolTipText     =   "Move the selected constraint up the list"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton cmdDownConstraint 
         Caption         =   "Down"
         Height          =   315
         Left            =   1485
         TabIndex        =   17
         ToolTipText     =   "Move the selected constraint down the list"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton cmdDeleteConstraint 
         Caption         =   "Del"
         Height          =   315
         Left            =   765
         TabIndex        =   16
         ToolTipText     =   "Remove the selected constraint"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton cmdAddConstraint 
         Caption         =   "Add"
         Height          =   315
         Left            =   105
         TabIndex        =   15
         ToolTipText     =   "Add the entered constraint"
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.Frame fraInherits 
      Caption         =   "Inherits"
      Height          =   960
      Left            =   45
      TabIndex        =   31
      Top             =   3825
      Width           =   2910
      Begin vsAdoSelector.VS_AdoSelector vssTables 
         Height          =   315
         Left            =   90
         TabIndex        =   20
         ToolTipText     =   "Select a Table to be inherited by the new table"
         Top             =   225
         Width           =   2715
         _ExtentX        =   4789
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
      End
      Begin VB.CommandButton cmdAddInherit 
         Caption         =   "Add"
         Height          =   315
         Left            =   90
         TabIndex        =   21
         ToolTipText     =   "Add the selected Table to the Inherit list"
         Top             =   585
         Width           =   630
      End
      Begin VB.CommandButton cmdDeleteInherit 
         Caption         =   "Del"
         Height          =   315
         Left            =   765
         TabIndex        =   22
         ToolTipText     =   "Remove the selected Table from the Inherit list"
         Top             =   585
         Width           =   630
      End
      Begin VB.CommandButton cmdDownInherit 
         Caption         =   "Down"
         Height          =   315
         Left            =   1485
         TabIndex        =   23
         ToolTipText     =   "Move the selected table down the Inherit list"
         Top             =   585
         Width           =   630
      End
      Begin VB.CommandButton cmdUpInherit 
         Caption         =   "Up"
         Height          =   315
         Left            =   2160
         TabIndex        =   24
         ToolTipText     =   "Move the selected Table up the Inherit list"
         Top             =   585
         Width           =   630
      End
   End
   Begin VB.ListBox lstInherits 
      Height          =   840
      ItemData        =   "frmAddTable.frx":030E
      Left            =   3015
      List            =   "frmAddTable.frx":0310
      TabIndex        =   25
      ToolTipText     =   "Displays the list of Inherited Tables"
      Top             =   3960
      Width           =   3960
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Table"
      Height          =   360
      Left            =   5715
      TabIndex        =   26
      ToolTipText     =   "Creates the new Table"
      Top             =   4860
      Width           =   1290
   End
   Begin VB.ListBox lstColumns 
      Height          =   2790
      ItemData        =   "frmAddTable.frx":0312
      Left            =   3015
      List            =   "frmAddTable.frx":0314
      TabIndex        =   14
      ToolTipText     =   "Displays the current Column definitions for the new Table"
      Top             =   90
      Width           =   3960
   End
   Begin VB.TextBox txtTable 
      Height          =   285
      Left            =   1035
      TabIndex        =   0
      ToolTipText     =   "Enter a name for the new Table"
      Top             =   45
      Width           =   1890
   End
   Begin VB.Frame fraColumn 
      Caption         =   "Column Definition"
      Height          =   2550
      Left            =   45
      TabIndex        =   28
      Top             =   360
      Width           =   2895
      Begin VB.TextBox txtLength2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   3
         Text            =   "1"
         ToolTipText     =   "Specify the length of the new column"
         Top             =   900
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.ComboBox cboColumnType 
         Height          =   315
         Left            =   1035
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select the data type for the new column"
         Top             =   540
         Width           =   1770
      End
      Begin VB.CheckBox chkPrimaryKey 
         Alignment       =   1  'Right Justify
         Caption         =   "Primary Key?"
         Height          =   255
         Left            =   90
         TabIndex        =   8
         ToolTipText     =   "Select to specify that this column is the primary key."
         Top             =   1530
         Width           =   2700
      End
      Begin MSComCtl2.UpDown udLength 
         Height          =   315
         Left            =   2566
         TabIndex        =   6
         Top             =   900
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLength"
         BuddyDispid     =   196628
         OrigLeft        =   2520
         OrigTop         =   1080
         OrigRight       =   2715
         OrigBottom      =   1395
         Max             =   4096
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtLength 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1965
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Specify the length of the new column"
         Top             =   900
         Width           =   600
      End
      Begin VB.TextBox txtDefault 
         Height          =   285
         Left            =   810
         TabIndex        =   9
         ToolTipText     =   "Enter a default value for the new Column"
         Top             =   1800
         Width           =   2010
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up"
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         ToolTipText     =   "Move the selected Column down the list"
         Top             =   2160
         Width           =   630
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "D&own"
         Height          =   315
         Left            =   1485
         TabIndex        =   12
         ToolTipText     =   "Move the selected Column up the list"
         Top             =   2160
         Width           =   630
      End
      Begin VB.CommandButton cmdDeleteColumn 
         Caption         =   "D&el"
         Height          =   315
         Left            =   765
         TabIndex        =   11
         ToolTipText     =   "Delete the selected Column from the Table definition"
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdAddColumn 
         Caption         =   "&Add"
         Height          =   315
         Left            =   90
         TabIndex        =   10
         ToolTipText     =   "Add the new Column to the Table definition"
         Top             =   2160
         Width           =   630
      End
      Begin VB.CheckBox chkColumnNull 
         Alignment       =   1  'Right Justify
         Caption         =   "Restrict null values?"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         ToolTipText     =   "Select to specify that data is required"
         Top             =   1260
         Width           =   2700
      End
      Begin VB.TextBox txtColumnName 
         Height          =   285
         Left            =   765
         TabIndex        =   1
         ToolTipText     =   "Enter a name for the new Column"
         Top             =   225
         Width           =   2055
      End
      Begin MSComCtl2.UpDown udLength2 
         Height          =   315
         Left            =   1710
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtLength2"
         BuddyDispid     =   196625
         OrigLeft        =   1665
         OrigTop         =   900
         OrigRight       =   1860
         OrigBottom      =   1215
         Max             =   4096
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.Label lblLength 
         Caption         =   "Length"
         Height          =   225
         Left            =   105
         TabIndex        =   33
         Top             =   930
         Width           =   750
      End
      Begin VB.Label lblDefault 
         Caption         =   "Default"
         Height          =   255
         Left            =   105
         TabIndex        =   32
         Top             =   1845
         Width           =   855
      End
      Begin VB.Label lblType 
         Caption         =   "Data Type"
         Height          =   255
         Left            =   105
         TabIndex        =   30
         Top             =   585
         Width           =   855
      End
      Begin VB.Label lblColName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   270
         Width           =   495
      End
   End
   Begin VB.Label lblName 
      Caption         =   "Table Name"
      Height          =   255
      Left            =   45
      TabIndex        =   27
      Top             =   90
      Width           =   975
   End
End
Attribute VB_Name = "frmAddTable"
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
Dim SerialCol As String

Private Sub cmdAddColumn_Click()
On Error GoTo Err_Handler
Dim ColDef As String
  If txtColumnName.Text = "" Then
    MsgBox "You must enter a column name!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboColumnType.Text = "" Then
    MsgBox "You must select a data type!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboColumnType.Text = "serial" Then
    If txtTable.Text = "" Then
      MsgBox "You must specify a table name before you add a serial column!", vbExclamation, "Error"
      Exit Sub
    End If
    If SerialCol <> "" Then
      MsgBox "You can only specify one serial column per table!", vbExclamation, "Error"
      Exit Sub
    End If
    MsgBox "A column of type 'int4' will be created which will have a default value derived from a sequence which will also be created for you. A unique index will be created on this column.", vbInformation, "Serial Column"
    SerialCol = txtColumnName.Text
    ColDef = QUOTE & txtColumnName.Text & QUOTE & " int4 DEFAULT nextval('" & txtTable.Text & "_pkey_id') PRIMARY KEY"
    lstColumns.AddItem ColDef
    txtColumnName.Text = ""
    cboColumnType.RemoveItem cboColumnType.ListIndex
    chkPrimaryKey.Enabled = False
    txtTable.Enabled = False
    Gen_SQL
    Exit Sub
  End If
  ColDef = QUOTE & txtColumnName.Text & QUOTE & " "
  If txtLength2.Visible = True And txtLength2.Text <> "" Then
    ColDef = ColDef & cboColumnType.Text & "(" & txtLength2.Text & "," & txtLength.Text & ")"
  ElseIf txtLength.Enabled = True And txtLength.Text <> "" Then
    ColDef = ColDef & cboColumnType.Text & "(" & txtLength.Text & ")"
  Else
    ColDef = ColDef & QUOTE & cboColumnType.Text & QUOTE
  End If
  If txtDefault.Text <> "" Then
    ColDef = ColDef & " DEFAULT " & txtDefault.Text
  End If
  If chkColumnNull.Value = 1 Then
    ColDef = ColDef & " NOT NULL"
  End If
  If chkPrimaryKey.Value = 1 Then
    ColDef = ColDef & " PRIMARY KEY"
    chkPrimaryKey.Value = 0
    chkPrimaryKey.Enabled = False
  End If
  lstColumns.AddItem ColDef
  txtColumnName.Text = ""
  txtDefault.Text = ""
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdAddColumn_Click"
End Sub

Private Sub cmdAddInherit_Click()
On Error GoTo Err_Handler
  If vssTables.Text = "" Then
    MsgBox "You must select a table to inherit!", vbExclamation, "Error"
    Exit Sub
  End If
  lstInherits.AddItem QUOTE & vssTables.Caption & QUOTE
  vssTables.Text = ""
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdAddInherit_Click"
End Sub

Private Sub cmdAddConstraint_Click()
On Error GoTo Err_Handler
  Load frmTableConst
  frmTableConst.Show
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdAddConstraint_Click"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim CreateStr As String
Dim X As Integer
  If txtTable.Text = "" Then
    MsgBox "You must enter a table name!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = "CREATE TABLE " & QUOTE & txtTable.Text & QUOTE & " ("
  X = 0
  Do While X <> lstColumns.ListCount - 1
    CreateStr = CreateStr & lstColumns.List(X) & ", "
    X = X + 1
  Loop
  CreateStr = CreateStr & lstColumns.List(X)
  X = 0
  If lstConstraints.ListCount > 0 Then
    CreateStr = CreateStr & ", "
    Do While X <> lstConstraints.ListCount - 1
      CreateStr = CreateStr & lstConstraints.List(X) & ", "
      X = X + 1
    Loop
    CreateStr = CreateStr & lstConstraints.List(X)
  End If
  If lstInherits.ListCount > 0 Then
    CreateStr = CreateStr & ") INHERITS ("
    X = 0
    Do While X <> lstInherits.ListCount - 1
      CreateStr = CreateStr & lstInherits.List(X) & ", "
      X = X + 1
    Loop
    CreateStr = CreateStr & lstInherits.List(X)
  End If
  StartMsg "Creating Table..."
  CreateStr = CreateStr & ")"
  LogMsg "Executing: " & CreateStr
  gConnection.Execute CreateStr
  LogQuery CreateStr
  
  If SerialCol <> "" Then
    CreateStr = "CREATE SEQUENCE " & QUOTE & txtTable.Text & "_pkey_id" & QUOTE
    LogMsg "Executing: " & CreateStr
    gConnection.Execute CreateStr
    LogQuery CreateStr
  End If
  frmTables.cmdRefresh_Click
  SerialCol = ""
  EndMsg
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdCreate_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
Dim X As Integer
  fMainForm.txtSQLPane.Text = "CREATE TABLE " & QUOTE & txtTable.Text & QUOTE & " ("
  X = 0
  If lstColumns.ListCount > 0 Then
    Do While X <> lstColumns.ListCount - 1
      fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "    " & lstColumns.List(X) & ", "
      X = X + 1
    Loop
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "    " & lstColumns.List(X)
  X = 0
  If lstConstraints.ListCount > 0 Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & ", "
    Do While X <> lstConstraints.ListCount - 1
      fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "    " & lstConstraints.List(X) & ", "
      X = X + 1
    Loop
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "    " & lstConstraints.List(X)
  If lstInherits.ListCount > 0 Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & ")" & vbCrLf & "  INHERITS ("
    X = 0
    Do While X <> lstInherits.ListCount - 1
      fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "    " & lstInherits.List(X) & ", "
      X = X + 1
    Loop
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "    " & lstInherits.List(X)
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & ")"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdCreate_Click"
End Sub

Private Sub cmdDeleteColumn_Click()
On Error GoTo Err_Handler
  If lstColumns.ListIndex = -1 Then
    MsgBox "You must select a column to delete!", vbExclamation, "Error"
    Exit Sub
  End If
  If InStr(1, lstColumns.Text, "PRIMARY KEY") <> 0 Then chkPrimaryKey.Enabled = True
  If lstColumns.Text = QUOTE & SerialCol & QUOTE & " int4 DEFAULT nextval('" & txtTable.Text & "_pkey_id') PRIMARY KEY" Then
    SerialCol = ""
    cboColumnType.AddItem "serial"
    txtTable.Enabled = True
  End If
  lstColumns.RemoveItem lstColumns.ListIndex
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdDeleteColumn_Click"
End Sub

Private Sub cmdDeleteInherit_Click()
On Error GoTo Err_Handler
  If lstInherits.ListIndex = -1 Then
    MsgBox "You must select a table to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  lstInherits.RemoveItem lstInherits.ListIndex
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdDeleteInherit_Click"
End Sub

Private Sub cmdDeleteConstraint_Click()
On Error GoTo Err_Handler
  If lstConstraints.ListIndex = -1 Then
    MsgBox "You must select a constraint to delete!", vbExclamation, "Error"
    Exit Sub
  End If
  lstConstraints.RemoveItem lstConstraints.ListIndex
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdDeleteConstraint_Click"
End Sub

Private Sub cmdDown_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstColumns.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstColumns.ListIndex = lstColumns.ListCount - 1 Then
    MsgBox "This column is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstColumns.List(lstColumns.ListIndex + 1)
  lstColumns.List(lstColumns.ListIndex + 1) = lstColumns.List(lstColumns.ListIndex)
  lstColumns.List(lstColumns.ListIndex) = Temp
  lstColumns.ListIndex = lstColumns.ListIndex + 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdDown_Click"
End Sub

Private Sub cmdDownInherit_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstInherits.ListIndex = -1 Then
    MsgBox "You must select a table to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstInherits.ListIndex = lstInherits.ListCount - 1 Then
    MsgBox "This table is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstInherits.List(lstInherits.ListIndex + 1)
  lstInherits.List(lstInherits.ListIndex + 1) = lstInherits.List(lstInherits.ListIndex)
  lstInherits.List(lstInherits.ListIndex) = Temp
  lstInherits.ListIndex = lstInherits.ListIndex + 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdDownInherit_Click"
End Sub

Private Sub cmdDownConstraint_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstConstraints.ListIndex = -1 Then
    MsgBox "You must select a constraint to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstConstraints.ListIndex = lstConstraints.ListCount - 1 Then
    MsgBox "This constraint is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstConstraints.List(lstConstraints.ListIndex + 1)
  lstConstraints.List(lstConstraints.ListIndex + 1) = lstConstraints.List(lstConstraints.ListIndex)
  lstConstraints.List(lstConstraints.ListIndex) = Temp
  lstConstraints.ListIndex = lstConstraints.ListIndex + 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdDownConstraint_Click"
End Sub

Private Sub cmdUp_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstColumns.ListIndex = -1 Then
    MsgBox "You must select a column to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstColumns.ListIndex = 0 Then
    MsgBox "This column is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstColumns.List(lstColumns.ListIndex - 1)
  lstColumns.List(lstColumns.ListIndex - 1) = lstColumns.List(lstColumns.ListIndex)
  lstColumns.List(lstColumns.ListIndex) = Temp
  lstColumns.ListIndex = lstColumns.ListIndex - 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdUp_Click"
End Sub

Private Sub cmdUpInherit_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstInherits.ListIndex = -1 Then
    MsgBox "You must select a table to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstInherits.ListIndex = 0 Then
    MsgBox "This table is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstInherits.List(lstInherits.ListIndex - 1)
  lstInherits.List(lstInherits.ListIndex - 1) = lstInherits.List(lstInherits.ListIndex)
  lstInherits.List(lstInherits.ListIndex) = Temp
  lstInherits.ListIndex = lstInherits.ListIndex - 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdUpInherit_Click"
End Sub

Private Sub cmdUpConstraint_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstConstraints.ListIndex = -1 Then
    MsgBox "You must select a constraint to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstConstraints.ListIndex = 0 Then
    MsgBox "This constraint is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstConstraints.List(lstConstraints.ListIndex - 1)
  lstConstraints.List(lstConstraints.ListIndex - 1) = lstConstraints.List(lstConstraints.ListIndex)
  lstConstraints.List(lstConstraints.ListIndex) = Temp
  lstConstraints.ListIndex = lstConstraints.ListIndex - 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cmdUpConstraint_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsTypes As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 7140
  Me.Height = 5625
  StartMsg "Retrieving data types..."
  If rsTypes.State <> adStateClosed Then rsTypes.Close
  LogMsg "Executing: SELECT typname FROM pg_type WHERE typrelid = 0 ORDER BY typname"
  rsTypes.Open "SELECT typname FROM pg_type WHERE typrelid = 0 ORDER BY typname", gConnection, adOpenForwardOnly
  While Not rsTypes.EOF
    If Mid(rsTypes!typname, 1, 1) <> "_" Then cboColumnType.AddItem rsTypes!typname
    rsTypes.MoveNext
  Wend
  If rsTypes.State <> adStateClosed Then rsTypes.Close
  cboColumnType.AddItem "serial"
  fMainForm.StatusBar1.Panels("Status").Text = "Retrieving table names..."
  fMainForm.StatusBar1.Refresh
  vssTables.Connect = Connect
  vssTables.SQL = "SELECT DISTINCT ON(table_name) table_oid, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
  LogMsg "Executing: " & vssTables.SQL
  vssTables.LoadList
  SerialCol = ""
  EndMsg
  Gen_SQL
  Set rsTypes = Nothing
  Exit Sub
Err_Handler:
  Set rsTypes = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddTable, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      Me.Height = 5625
      If Me.Width < 7140 Then Me.Width = 7140
    End If
    lstColumns.Width = Me.ScaleWidth - lstColumns.Left
    lstConstraints.Width = lstColumns.Width
    lstInherits.Width = lstColumns.Width
    cmdCreate.Left = lstColumns.Left + lstColumns.Width - cmdCreate.Width
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, Form_Resize"
End Sub

Private Sub txtTable_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, txtTable_Change"
End Sub

Private Sub cboColumnType_Click()
On Error GoTo Err_Handler
  If cboColumnType.Text = "char" Or cboColumnType.Text = "bpchar" Or cboColumnType.Text = "varchar" Then
    txtLength.Text = 1
    txtLength.Enabled = True
    udLength.Enabled = True
    udLength2.Visible = False
    txtLength2.Visible = False
    txtLength2.Enabled = False
    udLength2.Enabled = False
    txtDefault.Enabled = True
    chkColumnNull.Enabled = True
    chkPrimaryKey.Enabled = True
  ElseIf cboColumnType.Text = "numeric" Then
    txtLength.Text = 1
    txtLength.Enabled = True
    udLength.Enabled = True
    udLength2.Visible = True
    txtLength2.Visible = True
    txtLength2.Text = 1
    txtLength2.Enabled = True
    udLength2.Enabled = True
    txtDefault.Enabled = True
    chkColumnNull.Enabled = True
    chkPrimaryKey.Enabled = True
  ElseIf cboColumnType.Text = "serial" Then
    txtLength.Text = ""
    txtLength.Enabled = False
    udLength.Enabled = False
    udLength2.Visible = False
    txtLength2.Visible = False
    txtLength2.Enabled = False
    udLength2.Enabled = False
    txtDefault.Enabled = False
    chkColumnNull.Enabled = False
    chkPrimaryKey.Enabled = False
  Else
    txtLength.Text = ""
    txtLength.Enabled = False
    udLength.Enabled = False
    udLength2.Visible = False
    txtLength2.Visible = False
    txtLength2.Enabled = False
    udLength2.Enabled = False
    txtDefault.Enabled = True
    chkColumnNull.Enabled = True
    chkPrimaryKey.Enabled = True
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTable, cboColumnType_Change"
End Sub
