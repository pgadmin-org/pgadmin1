VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddColumn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Column"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmAddColumn.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1950
   ScaleWidth      =   4065
   Begin VB.TextBox txtLength2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   2250
      TabIndex        =   4
      Text            =   "1"
      ToolTipText     =   "Specify the length of the new column"
      Top             =   810
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txtLength 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   1395
      TabIndex        =   2
      Text            =   "1"
      ToolTipText     =   "Specify the length of the new column"
      Top             =   810
      Width           =   570
   End
   Begin VB.ComboBox cboColumnType 
      Height          =   315
      Left            =   1395
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select the type for the new column"
      Top             =   450
      Width           =   2580
   End
   Begin VB.CommandButton cmdAddColumn 
      Caption         =   "&Add Column"
      Height          =   345
      Left            =   2835
      TabIndex        =   7
      ToolTipText     =   "Add the new column"
      Top             =   1560
      Width           =   1155
   End
   Begin VB.TextBox txtDefault 
      Height          =   285
      Left            =   1395
      TabIndex        =   6
      ToolTipText     =   "Enter a default value for the column"
      Top             =   1215
      Width           =   2595
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1395
      TabIndex        =   0
      ToolTipText     =   "Enter a name for the new column"
      Top             =   90
      Width           =   2595
   End
   Begin MSComCtl2.UpDown udLength 
      Height          =   315
      Left            =   1950
      TabIndex        =   3
      Top             =   810
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtLength"
      BuddyDispid     =   196610
      OrigLeft        =   2055
      OrigTop         =   900
      OrigRight       =   2295
      OrigBottom      =   1215
      Max             =   4096
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   0   'False
   End
   Begin MSComCtl2.UpDown udLength2 
      Height          =   315
      Left            =   2790
      TabIndex        =   5
      Top             =   810
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtLength2"
      BuddyDispid     =   196609
      OrigLeft        =   2055
      OrigTop         =   900
      OrigRight       =   2295
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
      Left            =   60
      TabIndex        =   11
      Top             =   885
      Width           =   1425
   End
   Begin VB.Label lblDefault 
      Caption         =   "Default Value"
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   1260
      Width           =   1425
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   465
      Width           =   1425
   End
   Begin VB.Label lblName 
      Caption         =   "Column Name"
      Height          =   225
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frmAddColumn"
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

Private Sub Gen_SQL()
On Error GoTo Err_Handler
  fMainForm.txtSQLPane.Text = "ALTER TABLE " & QUOTE & frmTables.trvBrowser.SelectedItem.Text & QUOTE & vbCrLf
  
  If txtLength2.Visible = True And txtLength2.Text <> "" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " ADD COLUMN " & QUOTE & txtName.Text & QUOTE & " " & cboColumnType.Text
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "(" & txtLength.Text & "," & txtLength2.Text & ")"
  ElseIf txtLength.Enabled = True And txtLength.Text <> "" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "  ADD COLUMN " & QUOTE & txtName.Text & QUOTE & " " & cboColumnType.Text
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "(" & txtLength.Text & ")"
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "  ADD COLUMN " & QUOTE & txtName.Text & QUOTE & " " & QUOTE & cboColumnType.Text & QUOTE
  End If
  If txtDefault.Text <> "" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "ALTER TABLE " & QUOTE & frmTables.trvBrowser.SelectedItem.Text & QUOTE & " ALTER COLUMN " & QUOTE & txtName.Text & QUOTE & " SET DEFAULT " & txtDefault.Text
  End If
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, Gen_SQL"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4155 Then Me.Width = 4155
    If Me.Height < 2325 Then Me.Height = 2325
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, Form_Resize"
End Sub

Private Sub txtLength2_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, txtLength2_Change"
End Sub

Private Sub cboColumnType_Click()
On Error GoTo Err_Handler
  If cboColumnType.Text = "char" Or cboColumnType.Text = "bpchar" Or cboColumnType.Text = "varchar" Then
    txtLength.Text = 1
    txtLength.Enabled = True
    udLength.Enabled = True
    txtLength2.Visible = False
    udLength2.Visible = False
  ElseIf cboColumnType.Text = "numeric" Then
    txtLength.Text = 1
    txtLength.Enabled = True
    udLength.Enabled = True
    txtLength2.Visible = True
    udLength2.Visible = True
    txtLength2.Text = 1
    txtLength2.Enabled = True
    udLength2.Enabled = True
  Else
    txtLength.Text = ""
    txtLength.Enabled = False
    udLength.Enabled = False
    txtLength2.Visible = False
    udLength2.Visible = False
  End If
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, cboColumnType_Change"
End Sub

Private Sub cmdAddColumn_Click()
On Error GoTo Err_Handler
Dim AlterStr As String
  If txtName.Text = "" Then
    MsgBox "You must enter a column name!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboColumnType.Text = "" Then
    MsgBox "You must select a data type!", vbExclamation, "Error"
    Exit Sub
  End If
  AlterStr = "ALTER TABLE " & QUOTE & frmTables.trvBrowser.SelectedItem.Text & QUOTE

  If txtLength2.Visible = True And txtLength2.Text <> "" Then
    AlterStr = AlterStr & " ADD COLUMN " & QUOTE & txtName.Text & QUOTE & " " & cboColumnType.Text
    AlterStr = AlterStr & "(" & txtLength.Text & "," & txtLength2.Text & ")"
  ElseIf txtLength.Enabled = True And txtLength.Text <> "" Then
    AlterStr = AlterStr & " ADD COLUMN " & QUOTE & txtName.Text & QUOTE & " " & cboColumnType.Text
    AlterStr = AlterStr & "(" & txtLength.Text & ")"
  Else
    AlterStr = AlterStr & " ADD COLUMN " & QUOTE & txtName.Text & QUOTE & " " & QUOTE & cboColumnType.Text & QUOTE
  End If
  LogMsg "Executing: " & AlterStr
  gConnection.Execute AlterStr
  LogQuery AlterStr
  If txtDefault.Text <> "" Then
    AlterStr = "ALTER TABLE " & QUOTE & frmTables.trvBrowser.SelectedItem.Text & QUOTE & " ALTER COLUMN " & QUOTE & txtName.Text & QUOTE & " SET DEFAULT " & txtDefault.Text
    LogMsg "Executing: " & AlterStr
    gConnection.Execute AlterStr
    LogQuery AlterStr
  End If
  frmTables.cmdRefresh_Click
  Unload Me
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, cmdAddColumn_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsTypes As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4155
  Me.Height = 2325
  StartMsg "Retrieving Data Types..."
  If rsTypes.State <> adStateClosed Then rsTypes.Close
  LogMsg "Executing: SELECT typname, typname FROM pg_type WHERE typrelid = 0"
  rsTypes.Open "SELECT typname, typname FROM pg_type WHERE typrelid = 0", gConnection, adOpenForwardOnly
  While Not rsTypes.EOF
    If Mid(rsTypes!typname, 1, 1) <> "_" Then cboColumnType.AddItem rsTypes!typname
    rsTypes.MoveNext
  Wend
  If rsTypes.State <> adStateClosed Then rsTypes.Close
  EndMsg
  Gen_SQL
  Set rsTypes = Nothing
  Exit Sub
Err_Handler:
  Set rsTypes = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddColumn, Form_Load"
End Sub

Private Sub txtLength_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, txtLength_Change"
End Sub

Private Sub txtDefault_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, txtDefault_Change"
End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, txtName_Change"
End Sub
