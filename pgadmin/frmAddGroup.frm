VERSION 5.00
Begin VB.Form frmAddGroup 
   Caption         =   "Add Group"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmAddGroup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Group"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Create the new group."
      Top             =   2880
      Width           =   1410
   End
   Begin VB.ListBox lstUsers 
      Height          =   2085
      Left            =   900
      Style           =   1  'Checkbox
      TabIndex        =   2
      ToolTipText     =   "Select the users who will be members of the new group."
      Top             =   720
      Width           =   3750
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Text            =   "101"
      ToolTipText     =   "Enter the ID for the new group."
      Top             =   405
      Width           =   1050
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   900
      TabIndex        =   0
      ToolTipText     =   "Enter the name of the new group."
      Top             =   90
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Members"
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   5
      Top             =   765
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "Group ID"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   4
      Top             =   450
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   135
      Width           =   825
   End
End
Attribute VB_Name = "frmAddGroup"
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

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim szQry As String
Dim szUsers As String
Dim X As Integer
Dim rs As New Recordset
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the new group!", vbExclamation, "Error"
    txtName.SetFocus
    Exit Sub
  End If
  If InStr(1, txtName.Text, " ") <> 0 Then
    MsgBox "Group names may not contain spaces!", vbExclamation, "Error"
    txtName.SetFocus
    Exit Sub
  End If
  If Validate(txtID.Text, vdtInteger, False) <> True Then
    MsgBox "You must enter a valid (integer) ID for the new group!", vbExclamation, "Error"
    txtID.SetFocus
    Exit Sub
  End If
  szQry = " CREATE GROUP " & txtName.Text & " "
  For X = 0 To lstUsers.ListCount - 1
    If lstUsers.Selected(X) = True Then
      szUsers = szUsers & lstUsers.List(X) & ", "
    End If
  Next
  If (txtID.Text <> "") Or (Len(szUsers) > 2) Then szQry = szQry & "WITH "
  If txtID.Text <> "" Then szQry = szQry & "SYSID " & txtID.Text & " "
  If Len(szUsers) > 2 Then
    If Mid(szUsers, Len(szUsers) - 1, 2) = ", " Then szUsers = Mid(szUsers, 1, Len(szUsers) - 2)
    szQry = szQry & "USER " & szUsers
  End If
  LogMsg "Executing: " & szQry
  gConnection.Execute szQry
  frmGroups.cmdRefresh_Click
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddGroup, cmdCreate_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
Dim szQry As String
Dim szUsers As String
Dim X As Integer
  szQry = "CREATE GROUP " & txtName.Text & " "
  For X = 0 To lstUsers.ListCount - 1
    If lstUsers.Selected(X) = True Then
      szUsers = szUsers & lstUsers.List(X) & ", "
    End If
  Next
  If (txtID.Text <> "") Or (Len(szUsers) > 2) Then szQry = szQry & "WITH "
  If txtID.Text <> "" Then szQry = szQry & vbCrLf & "  SYSID " & txtID.Text & " "
  If Len(szUsers) > 2 Then
    If Mid(szUsers, Len(szUsers) - 1, 2) = ", " Then szUsers = Mid(szUsers, 1, Len(szUsers) - 2)
    szQry = szQry & vbCrLf & "  USER " & szUsers
  End If
  fMainForm.txtSQLPane.Text = szQry
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddGroup, Gen_SQL"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rs As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4800
  Me.Height = 3690
  StartMsg "Retrieving user names..."
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: SELECT user_name FROM pgadmin_users ORDER BY user_name"
  rs.Open "SELECT user_name FROM pgadmin_users ORDER BY user_name", gConnection, adOpenForwardOnly
  While Not rs.EOF
    lstUsers.AddItem rs!user_name
    rs.MoveNext
  Wend
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: SELECT max(group_id) FROM pgadmin_groups"
  rs.Open "SELECT max(group_id) FROM pgadmin_groups", gConnection, adOpenForwardOnly
  If rs!Max <> "" Then txtID.Text = rs!Max + 1
  EndMsg
  Exit Sub
Err_Handler:
  Set rs = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddGroup, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4800 Then Me.Width = 4800
    If Me.Height < 3690 Then Me.Height = 3690
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddGroup, Form_Resize"
End Sub

Private Sub lstUsers_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddGroup, lstUsers_Click"
End Sub

Private Sub txtID_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddGroup, txtID_Change"
End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddGroup, txtName_Change"
End Sub
