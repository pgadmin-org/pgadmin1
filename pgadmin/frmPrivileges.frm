VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrivileges 
   Caption         =   "Privileges"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   Icon            =   "frmPrivileges.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   6225
   Begin MSComctlLib.ListView lvClass 
      Height          =   1725
      Left            =   45
      TabIndex        =   1
      Top             =   270
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   3043
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Class Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ACL"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "&All"
      Height          =   225
      Left            =   1110
      TabIndex        =   0
      ToolTipText     =   "Select to apply privileges to all classes"
      Top             =   60
      Width           =   1380
   End
   Begin VB.PictureBox picControls 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   6225
      TabIndex        =   3
      Top             =   2025
      Width           =   6225
      Begin VB.ListBox lstUsers 
         Height          =   960
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   21
         ToolTipText     =   "Select individual users to grant privileges to"
         Top             =   90
         Width           =   4815
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply Privileges"
         Height          =   330
         Left            =   4770
         TabIndex        =   20
         ToolTipText     =   "Apply the selected privileges to the selected classes"
         Top             =   3240
         Width           =   1365
      End
      Begin VB.ListBox lstGroups 
         Height          =   960
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   19
         ToolTipText     =   "Select individual users to grant privileges to"
         Top             =   1125
         Width           =   4815
      End
      Begin VB.Frame fraAction 
         Caption         =   "Action"
         Height          =   1050
         Left            =   3105
         TabIndex        =   16
         Top             =   2520
         Width           =   1590
         Begin VB.OptionButton optAction 
            Alignment       =   1  'Right Justify
            Caption         =   "Revoke"
            Height          =   225
            Index           =   1
            Left            =   135
            TabIndex        =   18
            ToolTipText     =   "Select to Revoke the specified privileges"
            Top             =   630
            Width           =   1275
         End
         Begin VB.OptionButton optAction 
            Alignment       =   1  'Right Justify
            Caption         =   "Grant"
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   17
            ToolTipText     =   "Select to Grant the specified privileges"
            Top             =   405
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame fraPrivilege 
         Caption         =   "Privilege"
         Height          =   1050
         Left            =   45
         TabIndex        =   9
         Top             =   2520
         Width           =   2985
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "All"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   15
            ToolTipText     =   "Grant/Revoke ALL Privileges"
            Top             =   270
            Width           =   1230
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "Select"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   14
            ToolTipText     =   "Grant/Revoke SELECT Privilege"
            Top             =   495
            Width           =   1230
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "Insert"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   13
            ToolTipText     =   "Grant/Revoke INSERT Privilege"
            Top             =   720
            Width           =   1230
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "Update"
            Height          =   195
            Index           =   3
            Left            =   1575
            TabIndex        =   12
            ToolTipText     =   "Grant/Revoke UPDATE Privilege"
            Top             =   270
            Width           =   1230
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "Delete"
            Height          =   195
            Index           =   4
            Left            =   1575
            TabIndex        =   11
            ToolTipText     =   "Grant/Revoke DELETE Privilege"
            Top             =   495
            Width           =   1230
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "Rule"
            Height          =   195
            Index           =   5
            Left            =   1575
            TabIndex        =   10
            ToolTipText     =   "Grant/Revoke RULE Privilege"
            Top             =   720
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdAllUsers 
         Caption         =   "&Select All"
         Height          =   330
         Left            =   45
         TabIndex        =   8
         ToolTipText     =   "Select All Users"
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdNoUsers 
         Caption         =   "&Deselect All"
         Height          =   330
         Left            =   45
         TabIndex        =   7
         ToolTipText     =   "Deselect All Users"
         Top             =   720
         Width           =   1185
      End
      Begin VB.CommandButton cmdAllGroups 
         Caption         =   "&Select All"
         Height          =   330
         Left            =   45
         TabIndex        =   6
         ToolTipText     =   "Select All Groups"
         Top             =   1395
         Width           =   1185
      End
      Begin VB.CommandButton cmdNoGroups 
         Caption         =   "&Deselect All"
         Height          =   330
         Left            =   45
         TabIndex        =   5
         ToolTipText     =   "Deselect All Groups"
         Top             =   1755
         Width           =   1185
      End
      Begin VB.CheckBox chkPublic 
         Alignment       =   1  'Right Justify
         Caption         =   "&Public"
         Height          =   195
         Left            =   45
         TabIndex        =   4
         ToolTipText     =   "Set privileges to 'PUBLIC'"
         Top             =   2205
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "User Names:"
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   23
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Group Names:"
         Height          =   225
         Index           =   4
         Left            =   45
         TabIndex        =   22
         Top             =   1125
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Class Names:"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1065
   End
End
Attribute VB_Name = "frmPrivileges"
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

Private Sub chkAll_Click()
On Error GoTo Err_Handler
Dim x As Integer
  If chkAll.Value = 1 Then
    For x = 1 To lvClass.ListItems.Count
      lvClass.ListItems(x).Checked = True
    Next
    lvClass.Enabled = False
  Else
    lvClass.Enabled = True
  End If
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, chkAll_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
Dim SecString As String
Dim ClassList As String
Dim x As Integer
Dim y As Integer
  fMainForm.txtSQLPane.Text = ""
  If optAction(0).Value = True Then
    SecString = "GRANT *#-privilege-#* ON "
  Else
    SecString = "REVOKE *#-privilege-#* ON "
  End If
  For x = 1 To lvClass.ListItems.Count
    If lvClass.ListItems(x).Checked = True Then ClassList = ClassList & QUOTE & lvClass.ListItems(x).Text & QUOTE & ", "
  Next
  If ClassList <> "" Then
    ClassList = Mid(ClassList, 1, Len(ClassList) - 2)   'Remove last ", "
  End If
  If optAction(0).Value = True Then
    SecString = SecString & ClassList & " TO "
  Else
    SecString = SecString & ClassList & " FROM "
  End If
  fMainForm.txtSQLPane.Text = ""
  If chkPublic.Value = 1 Then
    If chkPrivilege(0).Value = 1 Then
      fMainForm.txtSQLPane.Text = Replace(SecString, "*#-privilege-#*", "ALL") & "public"
    Else
      fMainForm.txtSQLPane.Text = ""
      For x = 1 To 5
        If chkPrivilege(x).Value = 1 Then fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & Replace(SecString, "*#-privilege-#*", UCase(chkPrivilege(x).Caption)) & "public" & vbCrLf
      Next
    End If
  Else
    For x = 0 To lstUsers.ListCount - 1
      If lstUsers.Selected(x) = True Then
        If chkPrivilege(0).Value = 1 Then
          fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & Replace(SecString, "*#-privilege-#*", "ALL") & QUOTE & lstUsers.List(x) & QUOTE & vbCrLf
        Else
          For y = 1 To 5
            If chkPrivilege(y).Value = 1 Then fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & Replace(SecString, "*#-privilege-#*", UCase(chkPrivilege(y).Caption)) & QUOTE & lstUsers.List(x) & QUOTE & vbCrLf
          Next
        End If
      End If
    Next
    For x = 0 To lstGroups.ListCount - 1
      If lstGroups.Selected(x) = True Then
        If chkPrivilege(0).Value = 1 Then
          fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & Replace(SecString, "*#-privilege-#*", "ALL") & "GROUP " & QUOTE & lstGroups.List(x) & QUOTE & vbCrLf
        Else
          For y = 1 To 5
            If chkPrivilege(y).Value = 1 Then fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & Replace(SecString, "*#-privilege-#*", UCase(chkPrivilege(y).Caption)) & "GROUP " & QUOTE & lstGroups.List(x) & QUOTE & vbCrLf
          Next
        End If
      End If
    Next
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, cmdApply_Click"
End Sub

Private Sub chkPrivilege_Click(Index As Integer)
On Error GoTo Err_Handler
  If chkPrivilege(0).Value = 1 Then
    chkPrivilege(1).Value = 1
    chkPrivilege(2).Value = 1
    chkPrivilege(3).Value = 1
    chkPrivilege(4).Value = 1
    chkPrivilege(5).Value = 1
    chkPrivilege(1).Enabled = False
    chkPrivilege(2).Enabled = False
    chkPrivilege(3).Enabled = False
    chkPrivilege(4).Enabled = False
    chkPrivilege(5).Enabled = False
  Else
    chkPrivilege(1).Enabled = True
    chkPrivilege(2).Enabled = True
    chkPrivilege(3).Enabled = True
    chkPrivilege(4).Enabled = True
    chkPrivilege(5).Enabled = True
  End If
  Gen_SQL
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, chkPrivilege_Click"
End Sub

Private Sub chkPublic_Click()
On Error GoTo Err_Handler
  If chkPublic.Value = 1 Then
    lstUsers.Enabled = False
    lstGroups.Enabled = False
    cmdAllUsers.Enabled = False
    cmdNoUsers.Enabled = False
    cmdAllGroups.Enabled = False
    cmdNoGroups.Enabled = False
  Else
    lstUsers.Enabled = True
    lstGroups.Enabled = True
    cmdAllUsers.Enabled = True
    cmdNoUsers.Enabled = True
    cmdAllGroups.Enabled = True
    cmdNoGroups.Enabled = True
  End If
  Gen_SQL
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, chkPublic_Click"
End Sub

Private Sub cmdAllUsers_Click()
On Error GoTo Err_Handler
Dim x As Integer
  For x = 0 To lstUsers.ListCount - 1
    lstUsers.Selected(x) = True
  Next
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, cmdAllUsers_Click"
End Sub

Private Sub cmdAllGroups_Click()
On Error GoTo Err_Handler
Dim x As Integer
  For x = 0 To lstGroups.ListCount - 1
    lstGroups.Selected(x) = True
  Next
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, cmdAllGroups_Click"
End Sub

Private Sub cmdNoUsers_Click()
On Error GoTo Err_Handler
Dim x As Integer
  For x = 0 To lstUsers.ListCount - 1
    lstUsers.Selected(x) = False
  Next
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, cmdNoUsers_Click"
End Sub

Private Sub cmdNoGroups_Click()
On Error GoTo Err_Handler
Dim x As Integer
  For x = 0 To lstGroups.ListCount - 1
    lstGroups.Selected(x) = False
  Next
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, cmdNoGroups_Click"
End Sub

Private Sub cmdApply_Click()
On Error GoTo Err_Handler
Dim SecString As String
Dim ClassList As String
Dim szSQL As String
Dim bFlag As Boolean
Dim x As Integer
Dim y As Integer
  bFlag = False
  If chkAll.Value = 1 Then bFlag = True
  For x = 1 To lvClass.ListItems.Count
    If lvClass.ListItems(x).Checked = True Then bFlag = True
  Next
  If bFlag = False Then
    MsgBox "You must select at least one class!", vbExclamation, "Error"
    lvClass.SetFocus
    Exit Sub
  End If
  bFlag = False
  If chkPublic.Value = 1 Then bFlag = True
  For x = 0 To lstUsers.ListCount - 1
    If lstUsers.Selected(x) = True Then bFlag = True
  Next
  For x = 0 To lstGroups.ListCount - 1
    If lstGroups.Selected(x) = True Then bFlag = True
  Next
  If bFlag = False Then
    MsgBox "You must select at least one user or group!", vbExclamation, "Error"
    lstUsers.SetFocus
    Exit Sub
  End If
  bFlag = False
  For x = 0 To 5
    If chkPrivilege(x).Value = 1 Then bFlag = True
  Next
  If bFlag = False Then
    MsgBox "You must select at least one privilege!", vbExclamation, "Error"
    chkPrivilege(0).SetFocus
    Exit Sub
  End If
  StartMsg "Applying Privileges..."
  If optAction(0).Value = True Then
    SecString = "GRANT *#-privilege-#* ON "
  Else
    SecString = "REVOKE *#-privilege-#* ON "
  End If
  For x = 1 To lvClass.ListItems.Count
    If lvClass.ListItems(x).Checked = True Then ClassList = ClassList & QUOTE & lvClass.ListItems(x).Text & QUOTE & ", "
  Next
  If ClassList <> "" Then
    ClassList = Mid(ClassList, 1, Len(ClassList) - 2)   'Remove last ", "
  End If
  If optAction(0).Value = True Then
    SecString = SecString & ClassList & " TO "
  Else
    SecString = SecString & ClassList & " FROM "
  End If
  If chkPublic.Value = 1 Then
    If chkPrivilege(0).Value = 1 Then
      szSQL = Replace(SecString, "*#-privilege-#*", "ALL") & "public"
      LogMsg "Executing: " & szSQL
      gConnection.Execute szSQL
      LogQuery szSQL
    Else
      For x = 1 To 5
        If chkPrivilege(x).Value = 1 Then
          szSQL = Replace(SecString, "*#-privilege-#*", UCase(chkPrivilege(x).Caption)) & "public"
          LogMsg "Executing: " & szSQL
          gConnection.Execute szSQL
          LogQuery szSQL
        End If
      Next
    End If
  Else
    For x = 0 To lstUsers.ListCount - 1
      If lstUsers.Selected(x) = True Then
        If chkPrivilege(0).Value = 1 Then
          szSQL = Replace(SecString, "*#-privilege-#*", "ALL") & QUOTE & lstUsers.List(x) & QUOTE
          LogMsg "Executing: " & szSQL
          gConnection.Execute szSQL
          LogQuery szSQL
        Else
          For y = 1 To 5
            If chkPrivilege(y).Value = 1 Then
              szSQL = Replace(SecString, "*#-privilege-#*", UCase(chkPrivilege(y).Caption)) & QUOTE & lstUsers.List(x) & QUOTE
              LogMsg "Executing: " & szSQL
              gConnection.Execute szSQL
              LogQuery szSQL
            End If
          Next
        End If
      End If
    Next
    For x = 0 To lstGroups.ListCount - 1
      If lstGroups.Selected(x) = True Then
        If chkPrivilege(0).Value = 1 Then
          szSQL = Replace(SecString, "*#-privilege-#*", "ALL") & "GROUP " & QUOTE & lstGroups.List(x) & QUOTE
          LogMsg "Executing: " & szSQL
          gConnection.Execute szSQL
          LogQuery szSQL
        Else
          For y = 1 To 5
            If chkPrivilege(y).Value = 1 Then
              szSQL = Replace(SecString, "*#-privilege-#*", UCase(chkPrivilege(y).Caption)) & "GROUP " & QUOTE & lstGroups.List(x) & QUOTE
              LogMsg "Executing: " & szSQL
              gConnection.Execute szSQL
              LogQuery szSQL
            End If
          Next
        End If
      End If
    Next
  End If
  EndMsg
  Refresh_Lists
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmPrivileges, cmdApply_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 6360 Then Me.Width = 6360
    If Me.Height < 5970 Then Me.Height = 5970
  End If
  
  
    lvClass.Width = Me.ScaleWidth - lvClass.Left
    lstUsers.Width = Me.ScaleWidth - lstUsers.Left
    lstGroups.Width = Me.ScaleWidth - lstGroups.Left
    lvClass.Height = Me.ScaleHeight - lvClass.Top - picControls.Height
    picControls.Top = lvClass.Top + lvClass.Height
    cmdApply.Left = picControls.Width - cmdApply.Width

  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, Form_Load"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 6360
  Me.Height = 5970
  Refresh_Lists
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, Form_Load"
End Sub

Private Sub Refresh_Lists()
On Error GoTo Err_Handler
Dim rs As New Recordset
Dim itmX As ListItem
Dim x As Integer
  lvClass.ListItems.Clear
  lstUsers.Clear
  lstGroups.Clear
  chkAll.Value = 0
  chkPublic.Value = 0
  chkPrivilege(0).Value = 0
  chkPrivilege(1).Value = 0
  chkPrivilege(2).Value = 0
  chkPrivilege(3).Value = 0
  chkPrivilege(4).Value = 0
  chkPrivilege(5).Value = 0
  chkPrivilege(0).Enabled = True
  chkPrivilege(1).Enabled = True
  chkPrivilege(2).Enabled = True
  chkPrivilege(3).Enabled = True
  chkPrivilege(4).Enabled = True
  chkPrivilege(5).Enabled = True
  lvClass.Enabled = True
  lstUsers.Enabled = True
  lstGroups.Enabled = True
  cmdAllUsers.Enabled = True
  cmdNoUsers.Enabled = True
  cmdAllGroups.Enabled = True
  cmdNoGroups.Enabled = True
  
  StartMsg "Retrieving Class User & Group Names..."
  
  'Tables
  LogMsg "Executing: SELECT DISTINCT ON (table_name) table_name, table_acl FROM pgadmin_tables WHERE table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' AND table_oid > " & LAST_SYSTEM_OID
  rs.Open "SELECT DISTINCT ON (table_name) table_name, table_acl FROM pgadmin_tables WHERE table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' AND table_oid > " & LAST_SYSTEM_OID, gConnection, adOpenForwardOnly
  While Not rs.EOF
    Set itmX = lvClass.ListItems.Add(, , rs!table_name)
    itmX.SubItems(1) = "Table"
    itmX.SubItems(2) = rs!table_acl & ""
    rs.MoveNext
  Wend
  If rs.State <> adStateClosed Then rs.Close
  
  'Sequences
  LogMsg "Executing: SELECT relname, relacl FROM pg_class WHERE relname NOT LIKE 'pgadmin_%' AND oid > " & LAST_SYSTEM_OID & " AND relkind = 'S'"
  rs.Open "SELECT relname, relacl FROM pg_class WHERE relname NOT LIKE 'pgadmin_%' AND oid > " & LAST_SYSTEM_OID & " AND relkind = 'S'", gConnection, adOpenForwardOnly
  While Not rs.EOF
    Set itmX = lvClass.ListItems.Add(, , rs!relname)
    itmX.SubItems(1) = "Sequence"
    itmX.SubItems(2) = rs!relacl & ""
    rs.MoveNext
  Wend
  If rs.State <> adStateClosed Then rs.Close
  
  'Views
  LogMsg "Executing: SELECT view_name, view_acl FROM pgadmin_views WHERE view_name NOT LIKE 'pgadmin_%' AND view_oid > " & LAST_SYSTEM_OID
  rs.Open "SELECT view_name, view_acl FROM pgadmin_views WHERE view_name NOT LIKE 'pgadmin_%' AND view_oid > " & LAST_SYSTEM_OID, gConnection, adOpenForwardOnly
  While Not rs.EOF
    Set itmX = lvClass.ListItems.Add(, , rs!view_name)
    itmX.SubItems(1) = "View"
    itmX.SubItems(2) = rs!view_acl & ""
    rs.MoveNext
  Wend
  If rs.State <> adStateClosed Then rs.Close
  
  'Users
  LogMsg "Executing: SELECT user_name FROM pgadmin_users"
  rs.Open "SELECT user_name FROM pgadmin_users", gConnection, adOpenForwardOnly
  While Not rs.EOF
    lstUsers.AddItem rs!user_name
    rs.MoveNext
  Wend
  If rs.State <> adStateClosed Then rs.Close
  
  'Groups
  LogMsg "Executing: SELECT group_name FROM pgadmin_groups"
  rs.Open "SELECT group_name FROM pgadmin_groups", gConnection, adOpenForwardOnly
  While Not rs.EOF
    lstGroups.AddItem rs.Fields(0).Value
    rs.MoveNext
  Wend
  Set rs = Nothing
  EndMsg
  Exit Sub
Err_Handler:
  Set rs = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmPrivileges, Refresh_Lists"
End Sub
Private Sub lvClass_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, lvClass_Click"
End Sub

Private Sub lstGroups_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, lstGroups_Click"
End Sub

Private Sub lstUsers_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, lstUsers_Click"
End Sub

Private Sub optAction_Click(Index As Integer)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrivileges, optAction_Click"
End Sub

