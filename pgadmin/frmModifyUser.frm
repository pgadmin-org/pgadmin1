VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Begin VB.Form frmModifyUser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify User"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmModifyUser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   4200
   Begin MSComCtl2.MonthView calUser 
      Height          =   2370
      Left            =   1440
      TabIndex        =   2
      Top             =   765
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   59506690
      CurrentDate     =   36587
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   330
      Left            =   2835
      TabIndex        =   5
      Top             =   3510
      Width           =   1275
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Confirm the new password for the user"
      Top             =   420
      Width           =   2115
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Enter a new password for the user"
      Top             =   105
      Width           =   2115
   End
   Begin vsAdoSelector.VS_AdoSelector vssUser 
      Height          =   315
      Index           =   0
      Left            =   1425
      TabIndex        =   3
      ToolTipText     =   "Can the new user create databases?"
      Top             =   3195
      Width           =   870
      _ExtentX        =   1535
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
      SelectorType    =   1
      DisplayList     =   "No;Yes;"
      IndexList       =   "0;1;"
   End
   Begin vsAdoSelector.VS_AdoSelector vssUser 
      Height          =   315
      Index           =   1
      Left            =   1425
      TabIndex        =   4
      ToolTipText     =   "Can the new user create other users?"
      Top             =   3540
      Width           =   870
      _ExtentX        =   1535
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
      SelectorType    =   1
      DisplayList     =   "No;Yes;"
      IndexList       =   "0;1;"
   End
   Begin VB.Label Label1 
      Caption         =   "Create dbs"
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   10
      Top             =   3240
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Superuser"
      Height          =   225
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   3600
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Expiry date:"
      Height          =   225
      Index           =   2
      Left            =   105
      TabIndex        =   8
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Confirm Password:"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   445
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   130
      Width           =   1590
   End
End
Attribute VB_Name = "frmModifyUser"
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
Dim szMuser As String

Private Sub calUser_DateClick(ByVal DateClicked As Date)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmModifyUser, calUser_DateChange"
End Sub

Private Sub cmdApply_Click()
On Error GoTo Err_Handler
Dim UpdateStr As String
  If txtPassword(0).Text <> txtPassword(1).Text Then
    MsgBox "The passwords entered do not match!", vbExclamation, "Error"
    Exit Sub
  End If
  StartMsg "Updating User..."
  UpdateStr = " ALTER USER " & szMuser
  If txtPassword(0).Text <> "" Then
    UpdateStr = UpdateStr & " WITH PASSWORD '" & txtPassword(0).Text & "'"
  End If
  If vssUser(0).Caption = "No" Then
    UpdateStr = UpdateStr & " NOCREATEDB"
  Else
    UpdateStr = UpdateStr & " CREATEDB"
  End If
  If vssUser(1).Caption = "No" Then
    UpdateStr = UpdateStr & " NOCREATEUSER"
  Else
    UpdateStr = UpdateStr & " CREATEUSER"
  End If
  UpdateStr = UpdateStr & " VALID UNTIL '" & Format(calUser.Value, "MM/dd/yyyy") & " 01:00:00" & "'"
  LogMsg "Executing: " & Replace(UpdateStr, "PASSWORD '" & txtPassword(0).Text, "PASSWORD '********")
  gConnection.Execute UpdateStr
  frmUsers.cmdRefresh_Click
  EndMsg
  If txtPassword(0).Text = "" Then MsgBox "The password was not changed because it cannnot be set to a null value.", vbInformation, "Warning"
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmModifyUser, cmdApply_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
  fMainForm.txtSQLPane.Text = "ALTER USER " & szMuser
  If txtPassword(0).Text <> "" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  WITH PASSWORD '********'"
  End If
  If vssUser(0).Caption = "No" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  NOCREATEDB"
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  CREATEDB"
  End If
  If vssUser(1).Caption = "No" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  NOCREATEUSER"
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  CREATEUSER"
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  VALID UNTIL '" & Format(calUser.Value, "MM/dd/yyyy") & " 01:00:00" & "'"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmModifyUser, Gen_SQL"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4290
  Me.Height = 4275
  vssUser(0).LoadList
  vssUser(1).LoadList
  vssUser(0).Text = "0"
  vssUser(1).Text = "0"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmModifyUser, Form_Load"
End Sub

Public Sub ModifyUser(szUser As String)
On Error GoTo Err_Handler
Dim rsUser As New Recordset
  StartMsg "Retrieving User Info..."
  If rsUser.State <> adStateClosed Then rsUser.Close
  LogMsg "Executing: SELECT passwd, valuntil, usecreatedb, usesuper FROM pg_shadow WHERE usename = '" & szUser & "'"
  rsUser.Open "SELECT passwd, valuntil, usecreatedb, usesuper FROM pg_shadow WHERE usename = '" & szUser & "'", gConnection, adOpenForwardOnly
  txtPassword(0).Text = rsUser!passwd & ""
  txtPassword(1).Text = txtPassword(0).Text
  If rsUser!usecreatedb = "t" Or rsUser!usecreatedb = True Then vssUser(0).Text = "1"
  If rsUser!usesuper = "t" Or rsUser!usesuper = True Then vssUser(1).Text = "1"
  Me.Caption = "Modify User - " & szUser
  szMuser = szUser
  If rsUser!valuntil <> "" Then
    If InStr(1, rsUser!valuntil, " ") <> 0 Then
      calUser.Value = Mid(rsUser!valuntil, 1, InStr(1, rsUser!valuntil, " ") - 1)
    Else
      calUser.Value = rsUser!valuntil
    End If
  End If
  If rsUser.State <> adStateClosed Then rsUser.Close
  Set rsUser = Nothing
  EndMsg
  Gen_SQL
  Exit Sub
Err_Handler:
  If rsUser.State <> adStateClosed Then rsUser.Close
  Set rsUser = Nothing
  If Err.Number <> 0 Then LogError Err, "frmModifyUser, ModifyUser"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4290 Then Me.Width = 4290
    If Me.Height < 3945 Then Me.Height = 3945
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmModifyUser, Form_Resize"
End Sub

Private Sub txtPassword_Change(Index As Integer)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmModifyUser, txtPassword_Change"
End Sub

Private Sub vssUser_ItemSelected(Index As Integer, Item As String, ItemText As String)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmModifyUser, vssUser_ItemSelected"
End Sub
