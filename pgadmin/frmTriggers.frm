VERSION 5.00
Begin VB.Form frmTriggers 
   Caption         =   "Triggers"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdModifyTrig 
      Caption         =   "&Modify Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   22
      ToolTipText     =   "Modify the selected trigger."
      Top             =   405
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   21
      Top             =   2115
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Triggers"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Select to view system triggers."
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected Trigger."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Trigger Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   13
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtForEach 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   2760
      End
      Begin VB.TextBox txtEvent 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1485
         Width           =   2760
      End
      Begin VB.TextBox txtExecutes 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1170
         Width           =   2760
      End
      Begin VB.TextBox txtFunction 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   855
         Width           =   2760
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   2760
      End
      Begin VB.TextBox txtTable 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   2760
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   1590
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2340
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   19
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   18
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Executes"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   17
         Top             =   1215
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Event"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   1530
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "For Each"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   15
         Top             =   1845
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   14
         Top             =   2115
         Width           =   735
      End
   End
   Begin VB.ListBox lstTrig 
      Height          =   3960
      Left            =   1485
      TabIndex        =   5
      Top             =   45
      Width           =   2985
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Refresh the list of Triggers."
      Top             =   1485
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropTrig 
      Caption         =   "&Drop Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected Trigger."
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateTrig 
      Caption         =   "&Create Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new Trigger."
      Top             =   45
      Width           =   1410
   End
End
Attribute VB_Name = "frmTriggers"
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
Dim rsTrig As New Recordset

Public Sub cmdModifyTrig_Click()
' On Error GoTo Err_Handler

If txtOID <> "" Then
    ' This means we can open the function
    gPostgresOBJ_OID = Val(txtOID)
    
    ' Load form
    Load frmAddTrigger
    frmAddTrigger.Show
End If

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdModifyFunc_Click"
End Sub

Private Sub lstTrig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXTriggers
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, lstTrig_MouseUp"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, chkSystem_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsTrig = Nothing
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If txtOID.Text = "" Then
    MsgBox "You must select a Trigger to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  CallingForm = "frmTriggers"
  OID = txtOID.Text
  Load frmComments
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdComment_Click"
End Sub

Public Sub cmdCreateTrig_Click()
On Error GoTo Err_Handler
  Load frmAddTrigger
  frmAddTrigger.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdCreateTrig_Click"
End Sub

Public Sub cmdDropTrig_Click()
On Error GoTo Err_Handler
Dim szDropStr As String
  If lstTrig.Text = "" Then
    MsgBox "You must select a Trigger to delete!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to delete this Trigger?", vbYesNo + vbQuestion, _
            "Confirm Trigger Delete") = vbYes Then
    StartMsg "Dropping Trigger..."
    szDropStr = "DROP TRIGGER " & QUOTE & lstTrig.Text & QUOTE & " ON " & QUOTE & txtTable.Text & QUOTE
    fMainForm.txtSQLPane.Text = szDropStr
    LogMsg "Executing: DROP TRIGGER " & QUOTE & lstTrig.Text & QUOTE & " ON " & QUOTE & txtTable.Text & QUOTE
    gConnection.Execute szDropStr
    LogQuery szDropStr
    cmdRefresh_Click
    EndMsg
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdDropTrig_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  StartMsg "Retrieving Trigger Names..."
  lstTrig.Clear
  txtOID.Text = ""
  txtTable.Text = ""
  txtComments.Text = ""
  txtFunction.Text = ""
  txtForEach.Text = ""
  txtEvent.Text = ""
  txtExecutes.Text = ""
  If rsTrig.State <> adStateClosed Then rsTrig.Close
  If chkSystem.Value = 1 Then
    LogMsg "Executing: SELECT * FROM pgadmin_triggers ORDER BY trigger_name"
    rsTrig.Open "SELECT * FROM pgadmin_triggers ORDER BY trigger_name", gConnection, adOpenDynamic
  Else
    LogMsg "Executing: SELECT * FROM pgadmin_triggers WHERE trigger_oid > " & LAST_SYSTEM_OID & " AND trigger_name NOT LIKE 'pgadmin_%' AND trigger_name NOT LIKE 'pg_%' AND trigger_name NOT LIKE 'RI_ConstraintTrigger_%' ORDER BY trigger_name"
    rsTrig.Open "SELECT * FROM pgadmin_triggers WHERE trigger_oid > " & LAST_SYSTEM_OID & " AND trigger_name NOT LIKE 'pgadmin_%' AND trigger_name NOT LIKE 'pg_%' AND trigger_name NOT LIKE 'RI_ConstraintTrigger_%' ORDER BY trigger_name", gConnection, adOpenDynamic
  End If
  While Not rsTrig.EOF
    lstTrig.AddItem rsTrig!trigger_name & ""
    rsTrig.MoveNext
  Wend
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdRefresh_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = 8325
  Me.Height = 4455
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4455 Then Me.Height = 4455
    End If
    lstTrig.Height = Me.ScaleHeight
    lstTrig.Width = Me.ScaleWidth - lstTrig.Left - fraDetails.Width - 25
    fraDetails.Left = lstTrig.Left + lstTrig.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, Form_Resize"
End Sub

Public Sub lstTrig_Click()
On Error GoTo Err_Handler
Dim iTrigger_type As Integer
Dim iTemp As Integer
  If lstTrig.Text <> "" Then
    StartMsg "Retrieving trigger info..."
    If rsTrig.BOF <> True Then rsTrig.MoveFirst
    MoveRS rsTrig, lstTrig.ListIndex
    txtEvent.Text = ""
    txtOID.Text = rsTrig!trigger_oid & ""
    txtTable.Text = rsTrig!trigger_table & ""
    txtFunction.Text = rsTrig!trigger_function & ""
    iTrigger_type = CInt(rsTrig!trigger_type)
    If (iTrigger_type And 1) = 1 Then
      txtForEach.Text = "Row"
    Else
      txtForEach.Text = "Statement"
    End If
    If (iTrigger_type And 2) = 2 Then
      txtExecutes.Text = "Before"
    Else
      txtExecutes.Text = "After"
    End If
    If (iTrigger_type And 4) = 4 Then txtEvent.Text = txtEvent.Text & "Insert "
    If (iTrigger_type And 8) = 8 Then txtEvent.Text = txtEvent.Text & "Delete "
    If (iTrigger_type And 16) = 16 Then txtEvent.Text = txtEvent.Text & "Update "
    txtComments.Text = rsTrig!trigger_comments & ""
    If rsTrig.BOF <> True Then rsTrig.MoveFirst
    EndMsg
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTriggers, lstTrig_Click"
End Sub
