VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#4.1#0"; "HighlightBox.ocx"
Begin VB.Form frmAddFunction 
   Caption         =   "Create Function"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   Icon            =   "frmAddFunction.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   4155
   Begin HighlightBox.HBX txtPath 
      Height          =   735
      Left            =   1080
      TabIndex        =   15
      ToolTipText     =   "Enter the function code or Library Path."
      Top             =   2475
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   1296
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      ScrollBars      =   2
      MultiLine       =   -1  'True
   End
   Begin vsAdoSelector.VS_AdoSelector vssLanguage 
      Height          =   315
      Left            =   1080
      TabIndex        =   14
      ToolTipText     =   "Select the the language that the function is written in."
      Top             =   2115
      Width           =   3030
      _ExtentX        =   5345
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
      SQL             =   "SELECT ""lanname"", ""lanname"" FROM ""pg_language"" WHERE ""lanname"" <> 'Internal'"
   End
   Begin VB.ComboBox cboArguments 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   675
      Width           =   2085
   End
   Begin VB.ComboBox cboReturnType 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   360
      Width           =   2085
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "&Down"
      Height          =   285
      Left            =   3195
      TabIndex        =   11
      Top             =   1410
      Width           =   915
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   285
      Left            =   3195
      TabIndex        =   10
      Top             =   1755
      Width           =   915
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "&Up"
      Height          =   285
      Left            =   3195
      TabIndex        =   9
      Top             =   1080
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   285
      Left            =   3195
      TabIndex        =   8
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Function"
      Height          =   375
      Left            =   2655
      TabIndex        =   7
      ToolTipText     =   "Create the new function."
      Top             =   3285
      Width           =   1455
   End
   Begin VB.ListBox lstArguments 
      Height          =   1035
      Left            =   1080
      TabIndex        =   4
      Top             =   1035
      Width           =   2085
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   45
      Width           =   3030
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "Function or Library Path"
      Height          =   390
      Left            =   90
      TabIndex        =   6
      Top             =   2565
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLanguage 
      AutoSize        =   -1  'True
      Caption         =   "Language"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lblArguments 
      AutoSize        =   -1  'True
      Caption         =   "Arguments"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   765
      Width           =   750
   End
   Begin VB.Label lblReturnType 
      AutoSize        =   -1  'True
      Caption         =   "Return Type"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   405
      Width           =   885
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   465
   End
End
Attribute VB_Name = "frmAddFunction"
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

Private Sub cboReturnType_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, cboReturnType_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
  lstArguments.AddItem cboArguments.Text
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdAdd_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
Dim CreateStr As String
Dim X As Integer
  fMainForm.txtSQLPane.Text = "CREATE FUNCTION " & QUOTE & txtName.Text & QUOTE & vbCrLf & "  ("
  For X = 0 To lstArguments.ListCount - 2
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & lstArguments.List(X) & ", "
  Next X
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & lstArguments.List(X) & ") "
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  RETURNS " & cboReturnType.Text & " "
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  AS '" & txtPath.Text & "' "
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  LANGUAGE '" & vssLanguage.Text & "'"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, Gen_SQL"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim CreateStr As String
Dim X As Integer
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the function!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboReturnType.Text = "" Then
    MsgBox "You must select a return data type for the function!", vbExclamation, "Error"
    Exit Sub
  End If
  If vssLanguage.Text = "" Then
    MsgBox "You must select a language for the function!", vbExclamation, "Error"
    Exit Sub
  End If
  If vssLanguage.Text = "sql" Then
    If txtPath.Text = "" Then
      MsgBox "You must enter an SQL statement!", vbExclamation, "Error"
      Exit Sub
    End If
  Else
    If txtPath.Text = "" Then
      MsgBox "You must enter the path to the object library containing the function!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  CreateStr = "CREATE FUNCTION " & QUOTE & txtName.Text & QUOTE & " ("
  For X = 0 To lstArguments.ListCount - 2
    CreateStr = CreateStr & lstArguments.List(X) & ", "
  Next X
  CreateStr = CreateStr & lstArguments.List(X) & ") "
  CreateStr = CreateStr & "RETURNS " & cboReturnType.Text & " "
  CreateStr = CreateStr & "AS '" & txtPath.Text & "' "
  CreateStr = CreateStr & "LANGUAGE '" & vssLanguage.Text & "'"
  LogMsg "Executing: " & CreateStr
  gConnection.Execute CreateStr
  LogQuery CreateStr
  frmFunctions.cmdRefresh_Click
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdAdd_Click"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
  If lstArguments.ListIndex = -1 Then
    MsgBox "You must select an argument to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  lstArguments.RemoveItem lstArguments.ListIndex
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdAdd_Click"
End Sub

Private Sub cmdUp_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstArguments.ListIndex = -1 Then
    MsgBox "You must select an argument to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstArguments.ListIndex = 0 Then
    MsgBox "This argument is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstArguments.List(lstArguments.ListIndex - 1)
  lstArguments.List(lstArguments.ListIndex - 1) = lstArguments.List(lstArguments.ListIndex)
  lstArguments.List(lstArguments.ListIndex) = Temp
  lstArguments.ListIndex = lstArguments.ListIndex - 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdUp_Click"
End Sub

Private Sub cmdDown_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstArguments.ListIndex = -1 Then
    MsgBox "You must select an argument to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstArguments.ListIndex = lstArguments.ListCount - 1 Then
    MsgBox "This argument is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstArguments.List(lstArguments.ListIndex + 1)
  lstArguments.List(lstArguments.ListIndex + 1) = lstArguments.List(lstArguments.ListIndex)
  lstArguments.List(lstArguments.ListIndex) = Temp
  lstArguments.ListIndex = lstArguments.ListIndex + 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdDown_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4275 Then Me.Width = 4275
    If Me.Height < 4110 Then Me.Height = 4110
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsTypes As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 4110
  Me.Width = 4275
  txtPath.Wordlist = TextColours
  StartMsg "Retrieving data types and languages..."
  If rsTypes.State <> adStateClosed Then rsTypes.Close
  LogMsg "Executing: SELECT typname FROM pg_type WHERE typrelid = 0 ORDER BY typname"
  rsTypes.Open "SELECT typname FROM pg_type WHERE typrelid = 0 ORDER BY typname", gConnection, adOpenForwardOnly
  cboReturnType.Clear
  cboArguments.Clear
  cboReturnType.AddItem "opaque"
  While Not rsTypes.EOF
    If Mid(rsTypes!typname, 1, 1) <> "_" Then
      cboReturnType.AddItem rsTypes!typname
      cboArguments.AddItem rsTypes!typname
    End If
    rsTypes.MoveNext
  Wend
  If rsTypes.BOF <> True Then rsTypes.MoveFirst
  vssLanguage.Connect = Connect
  vssLanguage.SQL = "SELECT language_name, language_name FROM pgadmin_languages ORDER BY language_name"
  LogMsg "Executing: " & vssLanguage.SQL
  vssLanguage.LoadList
  lstArguments.Clear
  EndMsg
  Gen_SQL
  Set rsTypes = Nothing
  Exit Sub
Err_Handler:
  Set rsTypes = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddFunction, Form_Load"
End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, txtName_Change"
End Sub

Private Sub txtPath_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, txtPath_Change"
End Sub

Private Sub vssLanguage_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, vssLanguage_ItemSelected"
End Sub
