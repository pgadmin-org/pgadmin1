VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Object = "{65BD1FDD-C469-464B-98C7-8C7683B4AEE1}#17.1#0"; "adodatagrid.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmGroups 
   Caption         =   "Groups"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   Icon            =   "frmGroups.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   5730
   Begin HighlightBox.HBX txtUsers 
      Height          =   1995
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "List the users that are members of the selected group."
      Top             =   2025
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   3519
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Caption         =   "Group Members"
   End
   Begin vsAdoSelector.VS_AdoSelector vssUsers 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Select a user to add or remove from a group."
      Top             =   4455
      Width           =   2445
      _ExtentX        =   4313
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
      SQL             =   "SELECT user_name, user_name FROM pgadmin_users ORDER BY user_name"
   End
   Begin VB.CommandButton cmdRemoveUser 
      Caption         =   "&Remove User"
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      ToolTipText     =   "Remove the selected user from the selected group."
      Top             =   4095
      Width           =   1170
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "&Add User"
      Height          =   330
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Add the selected user to the selected group."
      Top             =   4095
      Width           =   1170
   End
   Begin adoDataGrid.DataGrid dgGroups 
      Align           =   1  'Align Top
      Height          =   2025
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "List the user groups on the system."
      Top             =   0
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   3572
      ViewTools       =   0   'False
      HideFirstColumn =   -1  'True
      HeaderText      =   "Group Name;Group ID"
      ColumnWidths    =   "2600;2600"
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   2565
      TabIndex        =   7
      ToolTipText     =   "Refresh the list of users"
      Top             =   4455
      Width           =   1170
   End
   Begin VB.CommandButton cmdDrop 
      Caption         =   "&Drop Group"
      Height          =   330
      Left            =   3780
      TabIndex        =   6
      ToolTipText     =   "Delete the selected user"
      Top             =   4095
      Width           =   1170
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Group"
      Height          =   330
      Left            =   2565
      TabIndex        =   5
      ToolTipText     =   "Create a new user"
      Top             =   4095
      Width           =   1170
   End
End
Attribute VB_Name = "frmGroups"
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
Dim rsGroups As New Recordset

Private Sub dgGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXGroups
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmGroups, dgGroups_MouseUp"
End Sub

Private Sub dgGroups_RowClick(RowNumber As Long, RowData() As String)
On Error GoTo Err_Handler
Dim szUsers() As String
Dim szUserList As String
Dim rsTemp As New Recordset
Dim x As Integer
  If Not rsGroups.EOF Then
    StartMsg "Retrieving Group Members..."
    If rsGroups!group_members <> "" Then
      szUsers = Split(Mid(rsGroups!group_members, 2, Len(rsGroups!group_members) - 2), ",")
      For x = 0 To UBound(szUsers)
        If rsTemp.State <> adStateClosed Then rsTemp.Close
        LogMsg "Executing: SELECT pg_get_userbyid('" & szUsers(x) & "') AS username"
        rsTemp.Open "SELECT pg_get_userbyid('" & szUsers(x) & "') AS username", gConnection, adOpenForwardOnly
        szUserList = szUserList & rsTemp!Username & ", "
      Next
    End If
    If Len(szUserList) > 2 Then
      If Mid(szUserList, Len(szUserList) - 1, 2) = ", " Then szUserList = Mid(szUserList, 1, Len(szUserList) - 2)
    End If
    txtUsers.Text = szUserList
    EndMsg
  Else
    txtUsers.Text = ""
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmGroups, dgGroups_RowClick"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsGroups = Nothing
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
Dim szDummy(0) As String
  StartMsg "Retrieving Group Information..."
  If rsGroups.State <> adStateClosed Then rsGroups.Close
  LogMsg "Executing: SELECT group_members, group_name, group_id FROM pgadmin_groups ORDER BY group_name"
  rsGroups.Open "SELECT group_members, group_name, group_id FROM pgadmin_groups ORDER BY group_name", gConnection, adOpenDynamic
  Set dgGroups.Recordset = rsGroups
  dgGroups_RowClick 0, szDummy
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number = -2147217887 Then
    MsgBox "Access Denied!", vbExclamation
    Unload Me
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err, "frmGroups, cmdRefresh_click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 4845
  Me.Width = 5850
  vssUsers.Connect = Connect
  vssUsers.LoadList
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmGroups, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtUsers.Minimise
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If Me.Width < 5850 Then Me.Width = 5850
      If Me.Height < 4845 Then Me.Height = 4845
    End If
    
    
        txtUsers.Width = Me.ScaleWidth
        dgGroups.Height = ((Me.ScaleHeight - ((cmdRefresh.Height + 50) * 2)) / 3) * 2
        txtUsers.Height = dgGroups.Height / 2
        txtUsers.Top = dgGroups.Top + dgGroups.Height
        cmdRefresh.Top = Me.ScaleHeight - cmdRefresh.Height
        cmdDrop.Top = Me.ScaleHeight - (cmdRefresh.Height * 2) - 50
        cmdCreate.Top = cmdDrop.Top
        cmdAddUser.Top = cmdDrop.Top
        cmdRemoveUser.Top = cmdDrop.Top
        vssUsers.Top = cmdRefresh.Top

  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmGroups, Form_Resize"
End Sub

Public Sub cmdCreate_Click()
On Error GoTo Err_Handler
  If SuperUser <> True Then
    MsgBox "You do not have sufficient authorisation to modify User Groups!", vbExclamation, "Error"
    Exit Sub
  End If
  Load frmAddGroup
  frmAddGroup.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmGroups, cmdCreate_click"
End Sub

Public Sub cmdDrop_Click()
On Error GoTo Err_Handler
  If rsGroups!group_name & "" = "" Then Exit Sub
  If SuperUser <> True Then
    MsgBox "You do not have sufficient authorisation to modify User Groups!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to delete group '" & rsGroups!group_name & "'?", vbYesNo + vbQuestion, _
            "Confirm Group Delete") = vbYes Then
    StartMsg "Dropping Group..."
    fMainForm.txtSQLPane.Text = "DROP GROUP " & rsGroups!group_name
    LogMsg "Executing: DROP GROUP " & rsGroups!group_name
    gConnection.Execute " DROP GROUP " & rsGroups!group_name
    cmdRefresh_Click
    EndMsg
   End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmGroups, cmdDrop_Click"
End Sub

Public Sub cmdAddUser_Click()
On Error GoTo Err_Handler
Dim szDummy(0) As String
  If rsGroups!group_name & "" = "" Then Exit Sub
  If SuperUser <> True Then
    MsgBox "You do not have sufficient authorisation to modify User Groups!", vbExclamation, "Error"
    Exit Sub
  End If
  If vssUsers.Text = "" Then
    MsgBox "You must select a user to add!", vbExclamation, "Error"
    vssUsers.SetFocus
    Exit Sub
  End If
  StartMsg "Adding User to Group..."
  fMainForm.txtSQLPane.Text = "ALTER GROUP " & rsGroups!group_name & " ADD USER " & vssUsers.Text
  LogMsg "Executing: ALTER GROUP " & rsGroups!group_name & " ADD USER " & vssUsers.Text
  gConnection.Execute " ALTER GROUP " & rsGroups!group_name & " ADD USER " & vssUsers.Text
  cmdRefresh_Click
  EndMsg
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmGroups, cmdAddUser_Click"
End Sub

Public Sub cmdRemoveUser_Click()
On Error GoTo Err_Handler
Dim szDummy(0) As String
  If rsGroups!group_name & "" = "" Then Exit Sub
  If SuperUser <> True Then
    MsgBox "You do not have sufficient authorisation to modify User Groups!", vbExclamation, "Error"
    Exit Sub
  End If
  If vssUsers.Text = "" Then
    MsgBox "You must select a user to remove!", vbExclamation, "Error"
    vssUsers.SetFocus
    Exit Sub
  End If
  StartMsg "Dropping User from Group..."
  fMainForm.txtSQLPane.Text = "ALTER GROUP " & rsGroups!group_name & " DROP USER " & vssUsers.Text
  LogMsg "Executing: ALTER GROUP " & rsGroups!group_name & " DROP USER " & vssUsers.Text
  gConnection.Execute " ALTER GROUP " & rsGroups!group_name & " DROP USER " & vssUsers.Text
  cmdRefresh_Click
  EndMsg
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmGroups, cmdDropUser_Click"
End Sub
