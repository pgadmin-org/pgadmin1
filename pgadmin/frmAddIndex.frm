VERSION 5.00
Begin VB.Form frmAddIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Index"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmAddIndex.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Index"
      Height          =   330
      Left            =   3465
      TabIndex        =   5
      ToolTipText     =   "Build the new Index"
      Top             =   2520
      Width           =   1275
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmAddIndex.frx":030A
      Left            =   1470
      List            =   "frmAddIndex.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select the type of Index to create"
      Top             =   2520
      Width           =   960
   End
   Begin VB.CheckBox chkUnique 
      Alignment       =   1  'Right Justify
      Caption         =   "Unique Index?"
      Height          =   225
      Left            =   75
      TabIndex        =   3
      ToolTipText     =   "Will only unique values be present?"
      Top             =   2205
      Width           =   1590
   End
   Begin VB.ComboBox cboTableoid 
      Height          =   315
      Left            =   2415
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   105
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ComboBox cboTables 
      Height          =   315
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select a table to Index"
      Top             =   105
      Width           =   3270
   End
   Begin VB.ListBox lstColumns 
      Height          =   1185
      Left            =   1470
      Style           =   1  'Checkbox
      TabIndex        =   2
      ToolTipText     =   "Select the columns to include in the Index"
      Top             =   945
      Width           =   3270
   End
   Begin VB.TextBox txtIndexName 
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      ToolTipText     =   "Enter a name for the Index"
      Top             =   525
      Width           =   3270
   End
   Begin VB.Label lblType 
      Caption         =   "Index Type:"
      Height          =   225
      Left            =   105
      TabIndex        =   10
      Top             =   2545
      Width           =   1485
   End
   Begin VB.Label lblColumns 
      Caption         =   "Columns to Index:"
      Height          =   225
      Left            =   105
      TabIndex        =   9
      Top             =   970
      Width           =   1485
   End
   Begin VB.Label lblName 
      Caption         =   "Index name:"
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   550
      Width           =   1170
   End
   Begin VB.Label lblTable 
      Caption         =   "Table to Index:"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   130
      Width           =   1170
   End
End
Attribute VB_Name = "frmAddIndex"
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

Private Sub cboTables_Click()
On Error GoTo Err_Handler
Dim rsColumns As New Recordset
  cboTableoid.ListIndex = cboTables.ListIndex
  lstColumns.Clear
  txtIndexName = cboTables.Text & "_idx"
  StartMsg "Retrieving column names..."
  LogMsg "Executing: SELECT column_name FROM pgadmin_tables WHERE column_position > 0 AND table_oid = " & cboTableoid.Text & " ORDER BY column_name"
  rsColumns.Open "SELECT column_name FROM pgadmin_tables WHERE column_position > 0 AND table_oid = " & cboTableoid.Text & " ORDER BY column_name", gConnection, adOpenDynamic
  While Not rsColumns.EOF
    lstColumns.AddItem rsColumns!column_name
    rsColumns.MoveNext
  Wend
  If rsColumns.State <> adStateClosed Then rsColumns.Close
  EndMsg
  Gen_SQL
  Set rsColumns = Nothing
  Exit Sub
Err_Handler:
  Set rsColumns = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddIndex, cboTables_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
Dim x As Integer
Dim ColumnList As String
Dim ColumnCount As Integer
  ColumnCount = 0
  For x = 0 To lstColumns.ListCount - 1
    If lstColumns.Selected(x) = True Then
      ColumnList = ColumnList & ", " & QUOTE & lstColumns.List(x) & QUOTE
      ColumnCount = ColumnCount + 1
    End If
  Next x
  fMainForm.txtSQLPane.Text = "CREATE "
  If chkUnique.Value = 1 Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "UNIQUE "
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "INDEX " & QUOTE & txtIndexName.Text & QUOTE & vbCrLf & "  ON " & _
              QUOTE & cboTables.Text & QUOTE & vbCrLf & "  USING " & cboType.Text & " (" & _
              Mid(ColumnList, 3) & ")"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddIndex, cmdCreate_Click"
End Sub

Private Sub cboType_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddIndex, cboType_Click"
End Sub

Private Sub chkUnique_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddIndex, chkUnique_Click"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim CreateStr As String
Dim x As Integer
Dim ColumnList As String
Dim ColumnCount As Integer

  If txtIndexName.Text = "" Then
    MsgBox "You must enter a name for the index!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboTables.Text = "" Then
    MsgBox "You must select a table to index!", vbExclamation, "Error"
    Exit Sub
  End If
  ColumnCount = 0
  For x = 0 To lstColumns.ListCount - 1
    If lstColumns.Selected(x) = True Then
      ColumnList = ColumnList & ", " & QUOTE & lstColumns.List(x) & QUOTE
      ColumnCount = ColumnCount + 1
    End If
  Next x
  If ColumnCount = 0 Then
    MsgBox "You must select at least one column to index!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboType.Text = "hash" And ColumnCount > 1 Then
    MsgBox "You can include only one column in a hash index!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboType.Text = "rtree" And ColumnCount > 1 Then
    MsgBox "You can include only one column in a rtree index!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboType.Text <> "btree" And chkUnique.Value = 1 Then
    MsgBox "Only btree indexes may be unique!", vbExclamation, "Error"
    Exit Sub
  End If
  StartMsg "Creating Index..."
  CreateStr = "CREATE "
  If chkUnique.Value = 1 Then
    CreateStr = CreateStr & "UNIQUE "
  End If
  CreateStr = CreateStr & "INDEX " & QUOTE & txtIndexName.Text & QUOTE & " ON " & _
              QUOTE & cboTables.Text & QUOTE & " USING " & cboType.Text & " (" & _
              Mid(ColumnList, 3) & ")"
  LogMsg "Executing: " & CreateStr
  gConnection.Execute CreateStr
  LogQuery CreateStr
  frmIndexes.cmdRefresh_Click
  EndMsg
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddIndex, cmdCreate_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4890 Then Me.Width = 4890
    If Me.Height < 3285 Then Me.Height = 3285
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsTables As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4890
  Me.Height = 3285
  cboTables.Clear
  cboTableoid.Clear
  txtIndexName = cboTables.Text & "_idx"
  StartMsg "Retrieving table names..."
  LogMsg "Executing: SELECT DISTINCT ON(table_name) table_oid, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
  rsTables.Open "SELECT DISTINCT ON(table_name) table_oid, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name", gConnection, adOpenForwardOnly
  While Not rsTables.EOF
    If Mid(rsTables.Fields(1).Value, 1, 3) <> "pg_" Then
      cboTables.AddItem rsTables!table_name
      cboTableoid.AddItem rsTables!table_oid
    End If
    rsTables.MoveNext
  Wend
  Set rsTables = Nothing
  cboType.ListIndex = 0
  EndMsg
  Gen_SQL
  Exit Sub
Err_Handler:
  Set rsTables = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddIndex, Form_Load"
End Sub

Private Sub lstColumns_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddIndex, lstColumns_Click"
End Sub

Private Sub txtIndexName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddIndex, txtName_Change"
End Sub
