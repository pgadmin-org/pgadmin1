VERSION 5.00
Object = "{65BD1FDD-C469-464B-98C7-8C7683B4AEE1}#17.1#0"; "adoDataGrid.ocx"
Begin VB.Form frmUsers 
   Caption         =   "Users"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   7485
   Begin adoDataGrid.DataGrid dgUsers 
      Align           =   1  'Align Top
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   5106
      ViewTools       =   0   'False
      HideFirstColumn =   -1  'True
      HeaderText      =   "Username;User ID;Create dbs;Superuser;Valid until"
      ColumnWidths    =   "1500;1000;1000;1000;2000"
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify User"
      Height          =   330
      Left            =   3825
      TabIndex        =   4
      ToolTipText     =   "Modify the selected user"
      Top             =   2970
      Width           =   1170
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create User"
      Height          =   330
      Left            =   1305
      TabIndex        =   2
      ToolTipText     =   "Create a new user"
      Top             =   2970
      Width           =   1170
   End
   Begin VB.CommandButton cmdDrop 
      Caption         =   "&Drop User"
      Height          =   330
      Left            =   2565
      TabIndex        =   3
      ToolTipText     =   "Delete the selected user"
      Top             =   2970
      Width           =   1170
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Refresh the list of users"
      Top             =   2970
      Width           =   1170
   End
End
Attribute VB_Name = "frmUsers"
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
Dim rsUsers As New Recordset

Private Sub dgUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXUsers
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUsers, grdUsers_MouseUp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsUsers = Nothing
End Sub

Public Sub cmdCreate_Click()
On Error GoTo Err_Handler
  If SuperUser <> True Then
    MsgBox "You do not have sufficient authorisation to modify user accounts!", vbExclamation, "Error"
    Exit Sub
  End If
  Load frmAddUser
  frmAddUser.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUsers, cmdCreate_click"
End Sub

Public Sub cmdDrop_Click()
On Error GoTo Err_Handler
  If rsUsers!user_name & "" = "" Then Exit Sub
  If SuperUser <> True Then
    MsgBox "You do not have sufficient authorisation to modify user accounts!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to delete user '" & rsUsers!user_name & "'?", vbYesNo + vbQuestion, _
            "Confirm User Delete") = vbYes Then
    StartMsg "Dropping User..."
    fMainForm.txtSQLPane.Text = "DROP USER " & QUOTE & rsUsers!user_name & QUOTE
    LogMsg "Executing: DROP USER " & QUOTE & rsUsers!user_name & QUOTE
    gConnection.Execute " DROP USER " & QUOTE & rsUsers!user_name & QUOTE
    cmdRefresh_Click
    EndMsg
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmUsers, cmdDrop_click"
End Sub

Public Sub cmdModify_Click()
On Error GoTo Err_Handler
  If rsUsers!user_name & "" = "" Then Exit Sub
  If SuperUser <> True Then
    MsgBox "You do not have sufficient authorisation to modify user accounts!", vbExclamation, "Error"
    Exit Sub
  End If
  Load frmModifyUser
  frmModifyUser.ModifyUser rsUsers!user_name
  frmModifyUser.Show
  frmModifyUser.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUsers, cmdModify_click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
Dim UserInfo As String
  StartMsg "Retrieving User Information..."
  If rsUsers.State <> adStateClosed Then rsUsers.Close
  LogMsg "Executing: SELECT * FROM pgadmin_users ORDER BY user_name"
  rsUsers.Open "SELECT * FROM pgadmin_users ORDER BY user_name", gConnection, adOpenForwardOnly
  Set dgUsers.Recordset = rsUsers
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number = -2147217887 Then
    MsgBox "Access Denied!", vbExclamation
    Unload Me
    Exit Sub
  End If
  If Err.Number <> 0 Then LogError Err, "frmUsers, cmdRefresh_click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 3735
  Me.Width = 7000
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUsers, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 5145 Then Me.Width = 5145
      If Me.Height < 3000 Then Me.Height = 3000
    End If
    dgUsers.Height = Me.ScaleHeight - cmdRefresh.Height - 50
    cmdRefresh.Top = Me.ScaleHeight - cmdRefresh.Height
    cmdDrop.Top = Me.ScaleHeight - cmdRefresh.Height
    cmdCreate.Top = Me.ScaleHeight - cmdRefresh.Height
    cmdModify.Top = Me.ScaleHeight - cmdRefresh.Height
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUsers, Form_Resize"
End Sub

Private Sub dgUsers_DblClick()
On Error GoTo Err_Handler
  cmdModify_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUsers, dgUsers_DblClick"
End Sub
