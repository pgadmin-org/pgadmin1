VERSION 5.00
Begin VB.Form frmLanguages 
   Caption         =   "Languages"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmLanguages.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   14
      Top             =   1125
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Languages"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Select to view system languages."
         Top             =   225
         Width           =   1155
      End
   End
   Begin VB.ListBox lstLang 
      Height          =   3960
      Left            =   1485
      TabIndex        =   4
      Top             =   45
      Width           =   2985
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Refresh the list of languages."
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropLang 
      Caption         =   "&Drop Language"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected language."
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateLang 
      Caption         =   "&Create Language"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new language."
      Top             =   45
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Language Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   9
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtCompiler 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1170
         Width           =   2670
      End
      Begin VB.TextBox txtHandler 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   855
         Width           =   2670
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   2670
      End
      Begin VB.TextBox txtTrusted 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trusted?"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Handler"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Compiler"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   1215
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmLanguages"
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
Dim rsLang As New Recordset

Private Sub lstLang_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXLanguages
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguages, lstLang_MouseUp"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguages, chkSystem_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsLang = Nothing
End Sub

Public Sub cmdCreateLang_Click()
On Error GoTo Err_Handler
  Load frmAddLanguage
  frmAddLanguage.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguages, cmdCreateLang_Click"
End Sub

Public Sub cmdDropLang_Click()
On Error GoTo Err_Handler
  If lstLang.Text = "" Then
    MsgBox "You must select a language to delete!", vbExclamation, "Error"
    Exit Sub
  End If
  If UCase(lstLang.Text) = "C" Or UCase(lstLang.Text) = "SQL" Or UCase(lstLang.Text) = "INTERNAL" Then
    MsgBox "You cannot delete the languages: 'C', 'sql' or 'internal'!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to delete this Language?", vbYesNo + vbQuestion, _
            "Confirm Language Delete") = vbYes Then
    fMainForm.txtSQLPane.Text = "DROP PROCEDURAL LANGUAGE '" & lstLang.Text & "'"
    LogMsg "Executing: DROP PROCEDURAL LANGUAGE '" & lstLang.Text & "'"
    gConnection.Execute "DROP PROCEDURAL LANGUAGE '" & lstLang.Text & "'"
    LogQuery "DROP PROCEDURAL LANGUAGE '" & lstLang.Text & "'"
    cmdRefresh_Click
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguages, cmdDropLang_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  StartMsg "Retrieving Language Names..."
  lstLang.Clear
  txtOID.Text = ""
  txtHandler.Text = ""
  txtTrusted.Text = ""
  txtCompiler.Text = ""
  If rsLang.State <> adStateClosed Then rsLang.Close
  If chkSystem.Value = 1 Then
    LogMsg "Executing: SELECT * FROM pgadmin_languages ORDER BY language_name"
    rsLang.Open "SELECT * FROM pgadmin_languages ORDER BY language_name", gConnection, adOpenForwardOnly
  Else
    LogMsg "Executing: SELECT * FROM pgadmin_languages WHERE language_name NOT LIKE 'pgadmin_%' AND language_name NOT LIKE 'pg_%' AND language_oid > " & LAST_SYSTEM_OID & " ORDER BY language_name"
    rsLang.Open "SELECT * FROM pgadmin_languages WHERE language_name NOT LIKE 'pgadmin_%' AND language_name NOT LIKE 'pg_%' AND language_oid > " & LAST_SYSTEM_OID & " ORDER BY language_name", gConnection, adOpenForwardOnly
  End If
  While Not rsLang.EOF
    lstLang.AddItem rsLang!language_name & ""
    rsLang.MoveNext
  Wend
  If rsLang.BOF <> True Then rsLang.MoveFirst
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmLanguages, cmdRefresh_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = 8435
  Me.Height = 4455
  LogMsg "Loading Form: " & Me.Name
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguages, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4455 Then Me.Height = 4455
    End If
    lstLang.Height = Me.ScaleHeight
    lstLang.Width = Me.ScaleWidth - lstLang.Left - fraDetails.Width - 25
    fraDetails.Left = lstLang.Left + lstLang.Width + 25
    fraDetails.Height = Me.ScaleHeight
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguages, Form_Resize"
End Sub

Public Sub lstLang_Click()
On Error GoTo Err_Handler
  If lstLang.Text <> "" Then
    If rsLang.BOF <> True Then rsLang.MoveFirst
    MoveRS rsLang, lstLang.ListIndex
    txtOID.Text = rsLang!language_oid & ""
    txtCompiler.Text = rsLang!language_compiler & ""
    txtTrusted.Text = rsLang!language_is_trusted & ""
    txtHandler.Text = rsLang!language_handler & ""
    If rsLang.BOF <> True Then rsLang.MoveFirst
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmLanguages, lstLang_Click"
End Sub
