VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Begin VB.Form frmAddTrigger 
   Caption         =   "Trigger"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
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
         Enabled         =   0   'False
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
      Caption         =   "&Save Trigger"
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      ToolTipText     =   "Saves trigger."
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
Dim lng_OpenTrig_OID As Long

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
'On Error GoTo Err_Handler
Dim szTriggerName As String
Dim szTriggerTable As String
Dim szTriggerFunction As String
Dim szTriggerArguments As String
Dim szTriggerForEach As String
Dim szTriggerExecutes As String
Dim szTriggerEvent As String
  
Dim szTriggerName_backup As String
Dim szTriggerTable_backup As String
Dim szTriggerFunction_backup As String
Dim szTriggerArguments_backup As String
Dim szTriggerForEach_backup As String
Dim szTriggerExecutes_backup As String
Dim szTriggerEvent_backup As String
Dim iLoop As Integer

bContinueCompilation = True
iLoop = 0

  'Trigger Name
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the trigger!", vbExclamation, "Error"
    Exit Sub
  End If
  szTriggerName = txtName.Text
  
  'Execution time
  If optExecution(0).Value = True Then
    szTriggerExecutes = "BEFORE"
  Else
    szTriggerExecutes = "AFTER"
  End If
  
  'Event
  szTriggerEvent = ""
  If chkEvent(0).Value = 1 Then szTriggerEvent = szTriggerEvent & " INSERT OR"
  If chkEvent(1).Value = 1 Then szTriggerEvent = szTriggerEvent & " UPDATE OR"
  If chkEvent(2).Value = 1 Then szTriggerEvent = szTriggerEvent & " DELETE OR"
  If szTriggerEvent = "" Then
    MsgBox "You must select at least one trigger event!", vbExclamation, "Error"
    Exit Sub
  End If
  szTriggerEvent = Mid(szTriggerEvent, 1, Len(szTriggerEvent) - 3)
  
  'Table
  If vssTable.Text = "" Then
    MsgBox "You must select a table to create the trigger on!", vbExclamation, "Error"
    Exit Sub
  End If
  szTriggerTable = vssTable.Text
  
  'For each
  If optForEach(0).Value = True Then
    szTriggerForEach = "ROW"
  Else
    szTriggerForEach = "STATEMENT"
  End If
  
  'Function name and arguments
  If cboFunction.Text = "" Then
    MsgBox "You must select a function for the trigger to execute!", vbExclamation, "Error"
    Exit Sub
  End If

  szTriggerFunction = Mid(cboFunction.Text, 1, InStr(1, cboFunction.Text, "(") - 1)

  szTriggerArguments = Mid(cboFunction.Text, InStr(1, cboFunction.Text, "("))
  szTriggerArguments = Replace(szTriggerArguments, "(", "")
  szTriggerArguments = Replace(szTriggerArguments, ")", "")
    
  ' In case of a creation, test existence of a trigger with same name & table
  If lng_OpenTrig_OID = 0 Then
    If cmp_Trigger_Exists(0, szTriggerName, szTriggerTable) = True Then
    MsgBox "Trigger " & szTriggerName & " already exists on table " & szTriggerTable, vbExclamation, "Error"
    Exit Sub
    End If
  End If
 
  StartMsg "Creating Trigger..."

   ' Backup trigger if exists
  If lng_OpenTrig_OID <> 0 Then cmp_Trigger_GetValues lng_OpenTrig_OID, "", szTriggerName_backup, szTriggerTable_backup, szTriggerFunction_backup, szTriggerArguments_backup, szTriggerForEach_backup, szTriggerExecutes_backup, szTriggerEvent_backup
  
  ' Drop trigger if exists
  If lng_OpenTrig_OID <> 0 Then cmp_Trigger_DropIfExists lng_OpenTrig_OID
  
  ' Create trigger
  cmp_Trigger_Create szTriggerName, szTriggerTable, szTriggerFunction, szTriggerArguments, szTriggerForEach, szTriggerExecutes, szTriggerEvent

  ' In case of a problem, if the trigger was deleted, rollback
  If (bContinueCompilation = False) And (cmp_Trigger_Exists(0, szTriggerName_backup, szTriggerTable_backup) = False) And (iLoop = 0) Then
      iLoop = iLoop + 1
      cmp_Trigger_Create szTriggerName_backup, szTriggerTable_backup, szTriggerFunction_backup, szTriggerArguments_backup, szTriggerForEach_backup, szTriggerExecutes_backup, szTriggerEvent_backup
  End If
  
  EndMsg
  frmTriggers.cmdRefresh_Click
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
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddTrigger, Gen_SQL"
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
Dim szQuery As String

  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4365
  Me.Height = 3855
  StartMsg "Retrieving Table & Function names..."
  vssTable.Connect = Connect
  vssTable.SQL = "SELECT DISTINCT ON(table_name) table_name, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
  LogMsg "Executing: " & vssTable.SQL
  vssTable.LoadList
  
  szQuery = "SELECT function_name, function_arguments FROM pgadmin_functions " & _
  "WHERE function_returns = NULL AND function_name NOT LIKE 'pg_%' AND function_name NOT LIKE 'pgadmin_%' AND function_name NOT LIKE 'RI_%'" & _
  "ORDER BY function_name"
  LogMsg "Executing: " & szQuery
  rsFuncs.Open szQuery, gConnection, adOpenForwardOnly
  While Not rsFuncs.EOF
    cboFunction.AddItem rsFuncs!Function_name & "(" & rsFuncs!Function_arguments & ")"
    rsFuncs.MoveNext
  Wend
  EndMsg
  Gen_SQL
  Set rsFuncs = Nothing
  
        ' Retrieve function if exists
  lng_OpenTrig_OID = gPostgresOBJ_OID
  gPostgresOBJ_OID = 0
  If lng_OpenTrig_OID <> 0 Then
    Me.Caption = "Modify trigger"
    Trigger_Load
  Else
    Me.Caption = "Create trigger"
  End If
  
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

Private Sub Trigger_Load()
On Error GoTo Err_Handler
    Dim temp_arg_list As Variant
    Dim temp_arg_item As Variant
    
    Dim szTriggerName As String
    Dim szTriggerTable As String
    Dim szTriggerFunction As String
    Dim szTriggerArguments As String
    Dim szTriggerForEach As String
    Dim szTriggerExecutes As String
    Dim szTriggerEvent As String
    
    StartMsg "Retrieving trigger information..."
    cmp_Trigger_GetValues lng_OpenTrig_OID, "", szTriggerName, szTriggerTable, szTriggerFunction, szTriggerArguments, szTriggerForEach, szTriggerExecutes, szTriggerEvent
    
    ' Loading trigger name
    txtName = szTriggerName
    
    ' For each Row
    If szTriggerForEach = "Row" Then
      ' Row
      optForEach(0).Value = True
      optForEach(1).Value = False
    Else
       ' Statement
      optForEach(0).Value = False
      optForEach(1).Value = True
    End If
    
    If szTriggerExecutes = "Before" Then
     ' Before
      optExecution(0).Value = True
      optExecution(1).Value = False
    Else
     ' After
     optExecution(0).Value = False
     optExecution(1).Value = True
    End If
    
    If InStr(szTriggerEvent, "Insert") > 0 Then chkEvent(0).Value = 1 ' Insert
    If InStr(szTriggerEvent, "Update") > 0 Then chkEvent(1).Value = 1 ' Update
    If InStr(szTriggerEvent, "Delete") > 0 Then chkEvent(2).Value = 1 ' Delete
      
    ' Check if trigger is not broken because function was dropped
    If cmp_Function_Exists(0, szTriggerFunction & "", szTriggerArguments & "") = True Then
        cboFunction = szTriggerFunction & "(" & szTriggerArguments & ")"
    Else
        MsgBox "The function called by the trigger does not exist. " & vbCrLf & _
        "You should select a new funtion or drop the trigger."
    End If
    
    ' Loading table
    vssTable.Text = szTriggerTable
    
    EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdRefresh_Click"
End Sub
