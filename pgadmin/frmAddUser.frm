VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add a User"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   4245
   Begin MSComCtl2.MonthView calUser 
      Height          =   2370
      Left            =   1440
      TabIndex        =   12
      Top             =   1080
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   79757314
      CurrentDate     =   36587
   End
   Begin vsAdoSelector.VS_AdoSelector vssUser 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "Can the new user create databases?"
      Top             =   3825
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
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create User"
      Height          =   330
      Left            =   2835
      TabIndex        =   9
      ToolTipText     =   "Create the new user with the information supplied"
      Top             =   4185
      Width           =   1275
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Enter the new password again to confirm it"
      Top             =   750
      Width           =   1800
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter a password for the new user"
      Top             =   420
      Width           =   1800
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Index           =   0
      Left            =   1470
      TabIndex        =   0
      ToolTipText     =   "Enter a username for the new user"
      Top             =   105
      Width           =   1800
   End
   Begin vsAdoSelector.VS_AdoSelector vssUser 
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      ToolTipText     =   "Can the new user create other users?"
      Top             =   4170
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
      Index           =   2
      Left            =   1440
      TabIndex        =   13
      ToolTipText     =   "Can the new user create databases?"
      Top             =   3480
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.Label Label1 
      Caption         =   "Group"
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   3510
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Confirm Password"
      Height          =   225
      Index           =   5
      Left            =   105
      TabIndex        =   8
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   225
      Index           =   4
      Left            =   105
      TabIndex        =   7
      Top             =   450
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Account expiry"
      Height          =   225
      Index           =   3
      Left            =   105
      TabIndex        =   6
      Top             =   1110
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Superuser"
      Height          =   225
      Index           =   2
      Left            =   105
      TabIndex        =   5
      Top             =   4230
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Create dbs"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   3870
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   130
      Width           =   960
   End
End
Attribute VB_Name = "frmAddUser"
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

Private Sub calUser_DateClick(ByVal DateClicked As Date)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddUser, calUser_DateClick"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim rsUserNames As New Recordset
Dim CreateStr As String
  If txtUser(1).Text <> txtUser(2).Text Then
    MsgBox "The passwords entered do not match!", vbExclamation, "Error"
    Exit Sub
  End If
  StartMsg "Loading Existing Usernames..."
  LogMsg "Executing: SELECT " & QUOTE & "usename" & QUOTE & " FROM " & QUOTE & "pg_shadow" & QUOTE
  rsUserNames.Open "SELECT " & QUOTE & "usename" & QUOTE & " FROM " & QUOTE & "pg_shadow" & QUOTE, gConnection, adOpenDynamic
  While Not rsUserNames.EOF
    If rsUserNames.Fields(0).Value = txtUser(0).Text Then
      MsgBox "The username entered already exists!", vbExclamation, "Error"
      EndMsg
      Exit Sub
    End If
    rsUserNames.MoveNext
  Wend
  EndMsg
  CreateStr = " CREATE USER " & QUOTE & txtUser(0).Text & QUOTE
  If txtUser(1).Text <> "" Then
    CreateStr = CreateStr & " WITH PASSWORD '" & txtUser(1).Text & "'"
  End If
  If vssUser(0).Text = "0" Then
    CreateStr = CreateStr & " NOCREATEDB"
  Else
    CreateStr = CreateStr & " CREATEDB"
  End If
  If vssUser(1).Text = "0" Then
    CreateStr = CreateStr & " NOCREATEUSER"
  Else
    CreateStr = CreateStr & " CREATEUSER"
  End If
  If vssUser(2).Text <> "" Then
    CreateStr = CreateStr & " IN GROUP " & QUOTE & vssUser(2).Text & QUOTE
  End If
  CreateStr = CreateStr & " VALID UNTIL '" & Format(calUser.Value, "yyyy-mm-dd") & " 01:00:00'"
  LogMsg "Executing: " & Replace(CreateStr, "PASSWORD '" & txtUser(1).Text, "PASSWORD '********")
  gConnection.Execute CreateStr
  frmUsers.cmdRefresh_Click
  Unload Me
  Set rsUserNames = Nothing
  Exit Sub
Err_Handler:
  Set rsUserNames = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddUser, cmdCreate_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
  fMainForm.txtSQLPane.Text = "CREATE USER " & QUOTE & txtUser(0).Text & QUOTE
  If txtUser(1).Text <> "" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  WITH PASSWORD " & QUOTE & "******** " & QUOTE
  End If
  If vssUser(0).Text = "0" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  NOCREATEDB"
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  CREATEDB"
  End If
  If vssUser(1).Text = "0" Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " NOCREATEUSER"
  Else
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " CREATEUSER"
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  VALID UNTIL '" & calUser.Value & " 01:00:00'"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddUser, Gen_SQL"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Height < 4980 Then Me.Height = 4980
    If Me.Width < 4335 Then Me.Width = 4335
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddUser, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 4980
  Me.Width = 4335
  StartMsg "Retrieving Groups..."
  vssUser(0).LoadList
  calUser.MinDate = Date
  calUser.Value = DateAdd("d", 90, Date)
  vssUser(1).LoadList
  vssUser(0).Text = "0"
  vssUser(1).Text = "0"
  vssUser(2).Connect = Connect
  vssUser(2).SQL = "SELECT groname, groname FROM pg_group ORDER BY groname"
  LogMsg "Executing: " & vssUser(2).SQL
  vssUser(2).LoadList
  Gen_SQL
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddUser, Form_Load"
End Sub

Private Sub txtUser_Change(Index As Integer)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddUser, txtUser_Change"
End Sub

Private Sub vssUser_ItemSelected(Index As Integer, Item As String, ItemText As String)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddUser, vssUser_ItemSelected"
End Sub
