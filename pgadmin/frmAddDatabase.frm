VERSION 5.00
Begin VB.Form frmAddDatabase 
   Caption         =   "Create Database"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   Icon            =   "frmAddDatabase.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   960
   ScaleWidth      =   4365
   Begin VB.CommandButton cmdCreatedb 
      Caption         =   "&Create Database"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      ToolTipText     =   "Create the new database."
      Top             =   540
      Width           =   1410
   End
   Begin VB.TextBox txtDatabase 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      ToolTipText     =   "Enter the name for the new database."
      Top             =   135
      Width           =   2940
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Database Name"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "frmAddDatabase"
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

Private Sub Gen_SQL()
On Error GoTo Err_Handler
  fMainForm.txtSQLPane.Text = "CREATE DATABASE " & txtDatabase.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, Gen_SQL"
End Sub

Private Sub cmdCreatedb_Click()
On Error GoTo Err_Handler
  If txtDatabase.Text = "" Then
    MsgBox "You must specify a database name!", vbExclamation, "Error"
    Exit Sub
  End If
  If InStr(1, txtDatabase.Text, " ") <> 0 Then
    MsgBox "Database names may not contain spaces!", vbExclamation, "Error"
    Exit Sub
  End If
  StartMsg "Creating Database..."
  LogMsg "Executing: CREATE DATABASE " & txtDatabase.Text
  gConnection.Execute " CREATE DATABASE " & txtDatabase.Text
  frmDatabases.cmdRefresh_Click
  EndMsg
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddDatabase, cmdCreatedb_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 1365
  Me.Width = 4485
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4485 Then Me.Width = 4485
    If Me.Height < 1365 Then Me.Height = 1365
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, Form_Resize"
End Sub

Private Sub txtDatabase_Change()
On Error GoTo Err_Handler
  txtDatabase.Text = LCase(txtDatabase.Text)
  txtDatabase.SelStart = Len(txtDatabase.Text)
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, txtDatabase_Change"
End Sub
