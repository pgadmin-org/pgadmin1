VERSION 5.00
Begin VB.Form frmAddSequence 
   Caption         =   "Create Sequence"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   Icon            =   "frmAddSequence.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   4200
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Sequence"
      Height          =   375
      Left            =   2655
      TabIndex        =   14
      ToolTipText     =   "Create the new sequence."
      Top             =   2250
      Width           =   1500
   End
   Begin VB.CheckBox chkCycle 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      ToolTipText     =   "Specifies whether or not the sequence restarts."
      Top             =   1935
      Width           =   330
   End
   Begin VB.TextBox txtCache 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Text            =   "1"
      ToolTipText     =   "Select a cache value for the sequence."
      Top             =   1620
      Width           =   1500
   End
   Begin VB.TextBox txtMaximum 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Text            =   "2147483647"
      ToolTipText     =   "Select the maximum value for the sequence."
      Top             =   1305
      Width           =   1500
   End
   Begin VB.TextBox txtMinimum 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "1"
      ToolTipText     =   "Select the minimum value for the sequence."
      Top             =   990
      Width           =   1500
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "1"
      ToolTipText     =   "Select the start value for the sequence."
      Top             =   675
      Width           =   1500
   End
   Begin VB.TextBox txtIncrement 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "1"
      ToolTipText     =   "Select a value for the sequence to increment by."
      Top             =   360
      Width           =   1500
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Enter a name for the new sequence."
      Top             =   45
      Width           =   2715
   End
   Begin VB.Label lblCycle 
      Caption         =   "Cycle:"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1980
      Width           =   915
   End
   Begin VB.Label lblCache 
      Caption         =   "Cache Value:"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   1665
      Width           =   1500
   End
   Begin VB.Label lblMaximum 
      Caption         =   "Maximum Value:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1350
      Width           =   1500
   End
   Begin VB.Label lblMinimum 
      Caption         =   "Minimum Value:"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   1035
      Width           =   1500
   End
   Begin VB.Label lblStart 
      Caption         =   "Start Value:"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label lblIncrement 
      Caption         =   "Increment:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   1500
   End
   Begin VB.Label lblName 
      Caption         =   "Sequence Name:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1500
   End
End
Attribute VB_Name = "frmAddSequence"
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
  fMainForm.txtSQLPane.Text = "CREATE SEQUENCE " & QUOTE & txtName.Text & QUOTE
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  INCREMENT " & txtIncrement.Text
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  START " & txtStart.Text
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  MINVALUE " & txtMinimum.Text
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  MAXVALUE " & txtMaximum.Text
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  CACHE " & txtCache.Text
  If chkCycle.Value = 1 Then fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & " cycle"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, Gen_SQL"
End Sub

Private Sub chkCycle_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, chkCycle_Click"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim CreateStr As String
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the new sequence!", vbExclamation, "Error"
    Exit Sub
  End If
  CreateStr = "CREATE SEQUENCE " & QUOTE & txtName.Text & QUOTE
  If txtIncrement.Text <> "" Then
    If IsNumeric(txtIncrement.Text) <> True Then
      MsgBox "The Increment value must be numeric!", vbExclamation, "Error"
      Exit Sub
    Else
      CreateStr = CreateStr & " INCREMENT " & txtIncrement.Text
    End If
  End If
  If txtStart.Text <> "" Then
    If IsNumeric(txtStart.Text) <> True Then
      MsgBox "The Start value must be numeric!", vbExclamation, "Error"
      Exit Sub
    Else
      CreateStr = CreateStr & " START " & txtStart.Text
    End If
  End If
  If txtMinimum.Text <> "" Then
    If IsNumeric(txtMinimum.Text) <> True Then
      MsgBox "The Minimum value must be numeric!", vbExclamation, "Error"
      Exit Sub
    Else
      CreateStr = CreateStr & " MINVALUE " & txtMinimum.Text
    End If
  End If
  If txtMaximum.Text <> "" Then
    If IsNumeric(txtMaximum.Text) <> True Then
      MsgBox "The Maximum value must be numeric!", vbExclamation, "Error"
      Exit Sub
    Else
      CreateStr = CreateStr & " MAXVALUE " & txtMaximum.Text
    End If
  End If
  If txtCache.Text <> "" Then
    If IsNumeric(txtCache.Text) <> True Then
      MsgBox "The Cache value must be numeric!", vbExclamation, "Error"
      Exit Sub
    Else
      CreateStr = CreateStr & " CACHE " & txtCache.Text
    End If
  End If
  If chkCycle.Value = 1 Then CreateStr = CreateStr & " cycle"
  StartMsg "Creating new sequence..."
  LogMsg "Executing: " & CreateStr
  gConnection.Execute CreateStr
  LogQuery CreateStr
  frmSequences.cmdRefresh_Click
  EndMsg
  Unload Me
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddSequence, cmdCreate_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = 4320
  Me.Height = 3075
  LogMsg "Loading Form: " & Me.Name
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 4320 Then Me.Width = 4320
      If Me.Height < 3075 Then Me.Height = 3075
    End If
    cmdCreate.Left = Me.ScaleWidth - cmdCreate.Width - 50
    cmdCreate.Top = Me.ScaleHeight - cmdCreate.Height - 50
    txtName.Width = Me.ScaleWidth - txtName.Left - 50
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, Form_Resize"
End Sub

Private Sub txtCache_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, txtCache_Change"
End Sub

Private Sub txtIncrement_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, txtIncremement_Change"
End Sub

Private Sub txtMaximum_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, txtMaximum_Change"
End Sub

Private Sub txtMinimum_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, txtMinimum_Change"
End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, txtName_Change"
End Sub

Private Sub txtStart_Change()
On Error Resume Next
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddSequence, txtStart_Change"
End Sub
