VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "VSAdoSelector.ocx"
Begin VB.Form frmAddLanguage 
   Caption         =   "Create Language"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   Icon            =   "frmAddLanguage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   4155
   Begin vsAdoSelector.VS_AdoSelector vssHandler 
      Height          =   315
      Left            =   1395
      TabIndex        =   8
      ToolTipText     =   "Select the Handler Function to use for this language."
      Top             =   630
      Width           =   2715
      _ExtentX        =   4789
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
   Begin VB.CheckBox chkTrusted 
      Height          =   285
      Left            =   1395
      TabIndex        =   7
      ToolTipText     =   "Specifies whether or not the language is trusted."
      Top             =   360
      Width           =   330
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1395
      TabIndex        =   2
      ToolTipText     =   "Enter a name for the language."
      Top             =   45
      Width           =   2715
   End
   Begin VB.TextBox txtCompiler 
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      ToolTipText     =   "Enter the compiler for the language. This field is currently unused by PostgreSQL."
      Top             =   990
      Width           =   2715
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Language"
      Height          =   375
      Left            =   2610
      TabIndex        =   0
      ToolTipText     =   "Create the new language."
      Top             =   1350
      Width           =   1500
   End
   Begin VB.Label lblName 
      Caption         =   "Language Name:"
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   90
      Width           =   1500
   End
   Begin VB.Label lblHandler 
      Caption         =   "Handler:"
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label lblCompiler 
      Caption         =   "Compiler (Unused):"
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   1035
      Width           =   1500
   End
   Begin VB.Label lblTrusted 
      Caption         =   "Trusted?:"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   405
      Width           =   1050
   End
End
Attribute VB_Name = "frmAddLanguage"
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

Private Sub chkTrusted_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddLanguage, chkTrusted_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4275
  Me.Height = 2160
  StartMsg "Retrieving handler names..."
  vssHandler.Connect = Connect
  vssHandler.SQL = "SELECT function_name, function_name FROM pgadmin_functions ORDER BY function_name"
  LogMsg "Executing: " & vssHandler.SQL
  vssHandler.LoadList
  EndMsg
  Gen_SQL
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddLanguage, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 4275 Then Me.Width = 4275
      If Me.Height < 2160 Then Me.Height = 2160
    End If
    cmdCreate.Left = Me.ScaleWidth - cmdCreate.Width - 50
    cmdCreate.Top = Me.ScaleHeight - cmdCreate.Height - 50
    txtName.Width = Me.ScaleWidth - txtName.Left - 50
    vssHandler.Width = txtName.Width
    txtCompiler.Width = txtName.Width
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguage, Form_Resize"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim CreateStr As String
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the language!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = "CREATE "
  If chkTrusted.Value = 1 Then
    CreateStr = CreateStr & "TRUSTED "
  End If
  CreateStr = CreateStr & "PROCEDURAL LANGUAGE '" & txtName.Text & "'"
  If vssHandler.Text = "" Then
    MsgBox "You must select a handler for the language!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = CreateStr & " HANDLER " & QUOTE & vssHandler.Text & QUOTE & " LANCOMPILER '" & txtCompiler.Text & "'"
  StartMsg "Creating new language..."
  LogMsg "Executing: " & CreateStr
  gConnection.Execute CreateStr
  LogQuery CreateStr
  frmLanguages.cmdRefresh_Click
  EndMsg
  Unload Me
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddLanguage, cmdCreate_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
  fMainForm.txtSQLPane.Text = "CREATE "
  If chkTrusted.Value = 1 Then
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "TRUSTED "
  End If
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & "PROCEDURAL LANGUAGE '" & txtName.Text & "'"
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  HANDLER " & vssHandler.Text & vbCrLf & "  LANCOMPILER '" & txtCompiler.Text & "'"
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddLanguage, Gen_SQL"
End Sub

Private Sub txtCompiler_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddLanguage, txtCompiler_Change"
End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddLanguage, txtName_Change"
End Sub

Private Sub vssHandler_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddLanguage, vssHandler_ItemSelected"
End Sub
