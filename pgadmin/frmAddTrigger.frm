VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "VSAdoSelector.ocx"
Begin VB.Form frmAddTrigger 
   Caption         =   "Create Trigger"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   Icon            =   "frmAddTrigger.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   4245
   Begin vsAdoSelector.VS_AdoSelector vssTable 
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      ToolTipText     =   "Select the table that the trigger will be created on."
      Top             =   1620
      Width           =   2760
      _ExtentX        =   4868
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
   Begin VB.Frame Frame3 
      Caption         =   "For Each"
      Height          =   555
      Left            =   90
      TabIndex        =   13
      Top             =   1980
      Width           =   4065
      Begin VB.OptionButton optForEach 
         Caption         =   "Row"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   15
         ToolTipText     =   "Specify that the Trigger will execute for each affected row."
         Top             =   270
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton optForEach 
         Caption         =   "Statement"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   14
         ToolTipText     =   "Specify that the Trigger will execute once for a single statement."
         Top             =   270
         Width           =   1320
      End
   End
   Begin VB.ComboBox cboFunction 
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "Select the function that the trigger will execute."
      Top             =   2610
      Width           =   2715
   End
   Begin VB.Frame Frame2 
      Caption         =   "Event"
      Height          =   600
      Left            =   90
      TabIndex        =   7
      Top             =   945
      Width           =   4065
      Begin VB.CheckBox chkEvent 
         Caption         =   "Update"
         Height          =   195
         Index           =   1
         Left            =   3105
         TabIndex        =   10
         ToolTipText     =   "Specify that the trigger will occur before or after an Update."
         Top             =   270
         Width           =   870
      End
      Begin VB.CheckBox chkEvent 
         Caption         =   "Delete"
         Height          =   195
         Index           =   2
         Left            =   1620
         TabIndex        =   9
         ToolTipText     =   "Specify that the trigger will occur before or after a Deletion."
         Top             =   270
         Width           =   1140
      End
      Begin VB.CheckBox chkEvent 
         Caption         =   "Insert"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         ToolTipText     =   "Specify that the trigger will occur before or after an Insert."
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Execution Time"
      Height          =   555
      Left            =   90
      TabIndex        =   4
      Top             =   360
      Width           =   4065
      Begin VB.OptionButton optExecution 
         Caption         =   "After"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   6
         ToolTipText     =   "Specify that the Trigger will execute after the event."
         Top             =   270
         Width           =   1320
      End
      Begin VB.OptionButton optExecution 
         Caption         =   "Before"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   5
         ToolTipText     =   "Specify that the Trigger will execute before the event."
         Top             =   270
         Value           =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Trigger"
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      ToolTipText     =   "Create the new trigger."
      Top             =   3015
      Width           =   1500
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Enter a name for the new Trigger."
      Top             =   45
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Function"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   11
      Top             =   2655
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Table"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1665
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Trigger Name"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1500
   End
End
Attribute VB_Name = "frmAddTrigger"
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

Private Sub cboFunction_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, cboFunction_Click"
End Sub


Private Sub chkEvent_Click(Index As Integer)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, chkEvent_Click"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim CreateStr As String
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the trigger!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = "CREATE TRIGGER " & QUOTE & txtName.Text & QUOTE
  If optExecution(0).Value = True Then
    CreateStr = CreateStr & " BEFORE"
  Else
    CreateStr = CreateStr & " AFTER"
  End If
  If chkEvent(0).Value = 1 Then CreateStr = CreateStr & " INSERT OR"
  If chkEvent(1).Value = 1 Then CreateStr = CreateStr & " UPDATE OR"
  If chkEvent(2).Value = 1 Then CreateStr = CreateStr & " DELETE OR"
  If Mid(CreateStr, Len(CreateStr) - 1, 2) <> "OR" Then
    MsgBox "You must select at least one trigger event!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = Mid(CreateStr, 1, Len(CreateStr) - 3) & " ON "
  If vssTable.Text = "" Then
    MsgBox "You must select a table to create the trigger on!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = CreateStr & QUOTE & vssTable.Text & QUOTE & " FOR EACH"
  If optForEach(0).Value = True Then
    CreateStr = CreateStr & " ROW"
  Else
    CreateStr = CreateStr & " STATEMENT"
  End If
  CreateStr = CreateStr & " EXECUTE PROCEDURE"
  If cboFunction.Text = "" Then
    MsgBox "You must select a function for the trigger to execute!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = CreateStr & " " & QUOTE & Mid(cboFunction.Text, 1, InStr(1, cboFunction.Text, "(") - 1)
  CreateStr = CreateStr & QUOTE & Mid(cboFunction.Text, InStr(1, cboFunction.Text, "("))
  StartMsg "Creating Trigger..."
  LogMsg "Executing: " & CreateStr
  gConnection.Execute CreateStr
  LogQuery CreateStr
  frmTriggers.cmdRefresh_Click
  EndMsg
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddTrigger, cmdCreate_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
  fMainForm.txtSQLPane.Text = "CREATE TRIGGER " & QUOTE & txtName.Text & QUOTE
  If optExecution(0).Value = True Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  BEFORE"
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  AFTER"
  End If
  If chkEvent(0).Value = 1 Then fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " INSERT OR"
  If chkEvent(1).Value = 1 Then fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " UPDATE OR"
  If chkEvent(2).Value = 1 Then fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " DELETE OR"
  If Mid(fMainForm.txtSQLPane.Text, Len(fMainForm.txtSQLPane.Text) - 2, Len(fMainForm.txtSQLPane.Text)) = " OR" Then
    fMainForm.txtSQLPane.Text = Mid(fMainForm.txtSQLPane.Text, 1, Len(fMainForm.txtSQLPane.Text) - 3) & vbCrLf & "  ON "
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  ON "
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & QUOTE & vssTable.Text & QUOTE & vbCrLf & "  FOR EACH"
  If optForEach(0).Value = True Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " ROW"
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " STATEMENT"
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  EXECUTE PROCEDURE"
  If cboFunction.Text <> "" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " " & QUOTE & Mid(cboFunction.Text, 1, InStr(1, cboFunction.Text, "(") - 1)
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & QUOTE & Mid(cboFunction.Text, InStr(1, cboFunction.Text, "("))
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, cmdCreate_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Height < 3855 Then Me.Height = 3855
    If Me.Width < 4365 Then Me.Width = 4365
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsFuncs As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4365
  Me.Height = 3855
  StartMsg "Retrieving Table & Function names..."
  vssTable.Connect = Connect
  vssTable.SQL = "SELECT DISTINCT ON(table_name) table_name, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
  LogMsg "Executing: " & vssTable.SQL
  vssTable.LoadList
  LogMsg "Executing: SELECT function_name, function_arguments FROM pgadmin_functions ORDER BY function_name"
  rsFuncs.Open "SELECT function_name, function_arguments FROM pgadmin_functions ORDER BY function_name", gConnection, adOpenForwardOnly
  While Not rsFuncs.EOF
    cboFunction.AddItem rsFuncs!function_name & "(" & rsFuncs!function_arguments & ")"
    rsFuncs.MoveNext
  Wend
  EndMsg
  Gen_SQL
  Set rsFuncs = Nothing
  Exit Sub
Err_Handler:
  Set rsFuncs = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddTrigger, Form_Load"
End Sub

Private Sub optExecution_Click(Index As Integer)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, optExecution_Click"
End Sub

Private Sub optForEach_Click(Index As Integer)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, optForEach_Click"
End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, txtName_Change"
End Sub

Private Sub vssTable_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, vssTable_ItemSelected"
End Sub
