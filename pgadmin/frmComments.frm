VERSION 5.00
Begin VB.Form frmComments 
   Caption         =   "Comments"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "frmComments.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   4650
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3510
      TabIndex        =   1
      Top             =   2430
      Width           =   1095
   End
   Begin VB.TextBox txtComments 
      Height          =   2325
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   45
      Width           =   4635
   End
End
Attribute VB_Name = "frmComments"
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
Dim rs As New Recordset
Dim lOID As String
Dim lCallingForm As String
Dim Update As Boolean

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rs = Nothing
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
Dim szComment As String
  StartMsg "Updating Comment..."
  szComment = Replace(Replace(txtComments.Text, "\", "\\"), "'", "\'")
  If lOID > LAST_SYSTEM_OID Then
    If Update = True Then
      LogMsg "Executing: UPDATE pgadmin_desc SET description = '" & szComment & "' WHERE objoid = " & lOID
      gConnection.Execute "UPDATE pgadmin_desc SET description = '" & szComment & "' WHERE objoid = " & lOID
      fMainForm.txtSQLPane.Text = "UPDATE pgadmin_desc SET description = '" & szComment & "' WHERE objoid = " & lOID
    Else
      LogMsg "Executing: INSERT INTO pgadmin_desc (objoid, description) VALUES (" & lOID & ", '" & szComment & "')"
      gConnection.Execute "INSERT INTO pgadmin_desc (objoid, description) VALUES (" & lOID & ", '" & szComment & "')"
      fMainForm.txtSQLPane.Text = "INSERT INTO pgadmin_desc (objoid, description) VALUES (" & lOID & ", '" & szComment & "')"
    End If
  Else
    If Update = True Then
      LogMsg "Executing: UPDATE pg_description SET description = '" & txtComments.Text & "' WHERE objoid = " & lOID
      gConnection.Execute "UPDATE pg_description SET description = '" & txtComments.Text & "' WHERE objoid = " & lOID
      fMainForm.txtSQLPane.Text = "UPDATE pg_description SET description = '" & txtComments.Text & "' WHERE objoid = " & lOID
    Else
      LogMsg "Executing: INSERT INTO pg_description (objoid, description) VALUES (" & lOID & ", '" & txtComments.Text & "')"
      gConnection.Execute "INSERT INTO pg_description (objoid, description) VALUES (" & lOID & ", '" & txtComments.Text & "')"
      fMainForm.txtSQLPane.Text = "INSERT INTO pg_description (objoid, description) VALUES (" & lOID & ", '" & txtComments.Text & "')"
    End If
  End If
  
  Select Case lCallingForm
    Case "frmDatabases"
      frmDatabases.cmdRefresh_Click
    Case "frmSequences"
      frmSequences.lstSeq_Click
    Case "frmIndexes"
      frmIndexes.cmdRefresh_Click
    Case "frmTables"
      frmTables.cmdRefresh_Click
    Case "frmLanguages"
      frmLanguages.cmdRefresh_Click
    Case "frmFunctions"
      frmFunctions.cmdRefresh_Click
    Case "frmTriggers"
      frmTriggers.cmdRefresh_Click
    Case "frmViews"
      frmViews.cmdRefresh_Click
  End Select
  EndMsg
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmComments, cmdSave_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 5000
  Me.Height = 3500
  lCallingForm = CallingForm
  lOID = OID
  Update = False
  If rs.State <> adStateClosed Then rs.Close
  If lOID > LAST_SYSTEM_OID Then
    LogMsg "Executing: SELECT description FROM pgadmin_desc WHERE objoid = " & lOID
    rs.Open "SELECT description FROM pgadmin_desc WHERE objoid = " & lOID, gConnection, adOpenDynamic
  Else
  LogMsg "Executing: SELECT description FROM pg_description WHERE objoid = " & lOID
  rs.Open "SELECT description FROM pg_description WHERE objoid = " & lOID, gConnection, adOpenDynamic
  End If
  If rs.EOF <> True Then
    txtComments.Text = rs!Description
    Update = True
  Else
    txtComments.Text = ""
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmComments, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.Width < 2000 Then Me.Width = 2000
    If Me.Height < 2000 Then Me.Height = 2000
    txtComments.Width = Me.ScaleWidth
    txtComments.Height = Me.ScaleHeight - cmdSave.Height - 100
    cmdSave.Top = txtComments.Top + txtComments.Height + 50
    cmdSave.Left = Me.ScaleWidth - cmdSave.Width
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmComments, Form_Resize"
End Sub

