VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUpgrade 
   Caption         =   "Upgrade Script"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3300
   Icon            =   "frmUpgrade.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1950
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   135
      Top             =   1485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreateScript 
      Caption         =   "&Create Script"
      Height          =   330
      Left            =   1890
      TabIndex        =   4
      Top             =   1575
      Width           =   1365
   End
   Begin VB.TextBox txtVersion 
      Height          =   285
      Left            =   2430
      TabIndex        =   3
      ToolTipText     =   "Enter the version number to create upgrades from. Queries with this or a greater version will be included."
      Top             =   1125
      Width           =   825
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   2925
      TabIndex        =   1
      Top             =   360
      Width           =   330
   End
   Begin VB.CheckBox chkConnect 
      Alignment       =   1  'Right Justify
      Caption         =   "Include \CONNECT statements:"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   $"frmUpgrade.frx":030A
      Top             =   810
      Width           =   3165
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Select a file to write the upgrade script to."
      Top             =   360
      Width           =   2850
   End
   Begin VB.Label lblVersion 
      Caption         =   "Produce upgrade from version:"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1170
      Width           =   2310
   End
   Begin VB.Label lblFilename 
      Caption         =   "Output file:"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   1005
   End
End
Attribute VB_Name = "frmUpgrade"
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

Private Sub cmdBrowse_Click()
On Error GoTo Err_Handler
  With CommonDialog1
    .Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    .Filter = "SQL Scripts (*.sql)|*.sql"
    .ShowSave
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtFile.Text = CommonDialog1.FileName
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUpgrade, cmdBrowse_click"
End Sub

Private Sub cmdCreateScript_Click()
On Error GoTo Err_Handler
Dim rsLog As New Recordset
Dim Count As Integer
Dim fNum As Integer
  If txtFile.Text = "" Then
    MsgBox "You must select a file to write the script to!", vbExclamation, "Error"
    Exit Sub
  End If
  If txtVersion.Text = "" Then
    MsgBox "You must enter a version to produce the upgrade from!", vbExclamation, "Error"
    Exit Sub
  End If
  If IsNumeric(txtVersion.Text) = False Then
    MsgBox "The version number must be a valid decimal number!", vbExclamation, "Error"
    Exit Sub
  End If
  LogMsg "Executing: SELECT username, query FROM pgadmin_rev_log WHERE version >= " & txtVersion.Text & "::float4"
  rsLog.Open "SELECT username, query FROM pgadmin_rev_log WHERE version >= " & txtVersion.Text & "::float4", gConnection, adOpenForwardOnly
  StartMsg "Writing upgrade script..."
  fNum = FreeFile
  Open txtFile.Text For Output As #fNum
  Count = 0
  While Not rsLog.EOF
    If chkConnect.Value = 1 Then Print #fNum, "\CONNECT - " & rsLog!Username
    Print #fNum, rsLog!Query
    Count = Count + 1
    fMainForm.StatusBar1.Panels("Status").Text = "Writing upgrade script - query: " & Count
    fMainForm.StatusBar1.Refresh
    rsLog.MoveNext
  Wend
  EndMsg
  Close #fNum
  MsgBox "Upgrade script from version " & txtVersion.Text & " has been written to " & txtFile.Text & "." & vbCrLf & Count & " Queries were written.", vbInformation
  Set rsLog = Nothing
  Unload Me
Err_Handler:
  Set rsLog = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmUpgrade, cmdCreateScript_click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 3420 Then Me.Width = 3420
    If Me.Height < 2355 Then Me.Height = 2355
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUpgrade, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  txtVersion.Text = TrackVer
  Me.Width = 3420
  Me.Height = 2355
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmUpgrade, Form_Load"
End Sub

