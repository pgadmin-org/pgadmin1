VERSION 5.00
Object = "{65BD1FDD-C469-464B-98C7-8C7683B4AEE1}#17.1#0"; "adoDataGrid.ocx"
Begin VB.Form frmTracking 
   Caption         =   "Revision Tracking"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmTracking.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   6630
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "&Clear Tracking Log"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Delete all records from the tracking log."
      Top             =   4320
      Width           =   2040
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate Upgrade Script"
      Height          =   375
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Generate an Auto Upgrade SQL Script in PostgreSQL pg_dump format."
      Top             =   4320
      Width           =   2040
   End
   Begin adoDataGrid.DataGrid grdLog 
      Align           =   1  'Align Top
      Height          =   1800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   3175
      ViewTools       =   0   'False
      HeaderText      =   "Timestamp;Username;Version;Query"
      ColumnWidths    =   "2000;1440;1000;6000"
   End
End
Attribute VB_Name = "frmTracking"
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
Dim rsLog As New Recordset

Private Sub cmdClearLog_Click()
On Error GoTo Err_Handler
Dim Response As Integer
  Response = MsgBox("Clearing the Tracking Log will prevent you from" & _
                    " creating upgrade scripts. " & vbCrLf & _
                    "Do you wish to continue?", _
                    vbInformation + vbYesNo, "Clear Tracking Log")
  If Response = vbNo Then Exit Sub
  Response = MsgBox("Clearing the Tracking Log cannot be undone." & vbCrLf & _
                    "Are you sure you wish to continue?", _
                    vbExclamation + vbYesNo, "Clear Tracking Log")
  If Response = vbNo Then Exit Sub
  StartMsg "Clearing Tracking Log..."
  LogMsg "Executing: DELETE FROM pgadmin_rev_log"
  gConnection.Execute "DELETE FROM pgadmin_rev_log"
  rsLog.Requery
  Set grdLog.Recordset = rsLog
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTracking, cmdClearLog_Click"
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo Err_Handler
  Load frmUpgrade
  frmUpgrade.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTracking, cmdGenerate_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim Response As Integer
Dim rs As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 6750
  Me.Height = 5130
  StartMsg "Loading Revision Tracking Log..."
  LogMsg "Executing: SELECT * FROM pgadmin_rev_log"
  If rsLog.State <> adStateClosed Then rsLog.Close
  rsLog.Open "SELECT * FROM pgadmin_rev_log", gConnection, adOpenForwardOnly, adLockReadOnly
  Set grdLog.Recordset = rsLog
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  EndMsg
  Exit Sub
Err_Handler:
  Set rs = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTracking, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 6650 Then Me.Width = 6650
      If Me.Height < 2500 Then Me.Height = 2500
    End If
    cmdGenerate.Top = Me.ScaleHeight - cmdGenerate.Height - 50
    cmdClearLog.Top = cmdGenerate.Top
    grdLog.Height = cmdGenerate.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTracking, Form_Resize"
End Sub


