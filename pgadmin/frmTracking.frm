VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "VSAdoSelector.ocx"
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
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1050
      Left            =   45
      TabIndex        =   5
      Top             =   3690
      Width           =   6540
      Begin vsAdoSelector.VS_AdoSelector vssTracking 
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         ToolTipText     =   "Select whether or not to use Revision Tracking."
         Top             =   270
         Width           =   1050
         _ExtentX        =   1852
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
         DisplayList     =   "Yes;No;"
         IndexList       =   "Y;N;"
      End
      Begin VB.TextBox txtTracking 
         Height          =   285
         Left            =   1710
         TabIndex        =   2
         ToolTipText     =   "Set the Version number of the last software release."
         Top             =   630
         Width           =   1050
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate Upgrade Script"
         Height          =   690
         Left            =   2880
         TabIndex        =   3
         ToolTipText     =   "Generate an Auto Upgrade SQL Script in PostgreSQL pg_dump format."
         Top             =   225
         Width           =   1725
      End
      Begin VB.CommandButton cmdClearLog 
         Caption         =   "&Clear Tracking Log"
         Height          =   690
         Left            =   4680
         TabIndex        =   4
         ToolTipText     =   "Delete all records from the tracking log."
         Top             =   225
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Enable Tracking"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Last Release Version"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   675
         Width           =   1725
      End
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
  StartMsg "Loading Revision Tracking Options..."
  LogMsg "Executing: SELECT * FROM pgadmin_rev_log"
  If rsLog.State <> adStateClosed Then rsLog.Close
  rsLog.Open "SELECT * FROM pgadmin_rev_log", gConnection, adOpenForwardOnly, adLockReadOnly
  Set grdLog.Recordset = rsLog
  
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: SELECT param_value FROM pgadmin_param WHERE param_id = 2"
  rs.Open "SELECT param_value FROM pgadmin_param WHERE param_id = 2", gConnection, adOpenForwardOnly
  vssTracking.LoadList
  vssTracking.SelectItem rs!param_value
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: SELECT param_value FROM pgadmin_param WHERE param_id = 3"
  rs.Open "SELECT param_value FROM pgadmin_param WHERE param_id = 3", gConnection, adOpenForwardOnly
  txtTracking.Text = rs!param_value
  EndMsg
  Set rs = Nothing
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
    fraOptions.Top = Me.ScaleHeight - fraOptions.Height - 50
    grdLog.Height = fraOptions.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTracking, Form_Resize"
End Sub

Private Sub txtTracking_Change()
On Error GoTo Err_Handler
  If Validate(txtTracking.Text, vdtNumeric, True) = False Then Exit Sub
  LogMsg "Executing: UPDATE pgadmin_param SET param_value = '" & txtTracking.Text & "' WHERE param_id = 3"
  gConnection.Execute "UPDATE pgadmin_param SET param_value = '" & txtTracking.Text & "' WHERE param_id = 3"
  TrackVer = txtTracking.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTracking, txtTracking_Change"
End Sub

Private Sub vssTracking_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
  If Item = "Y" Then
    Tracking = True
  Else
    Tracking = False
  End If
  LogMsg "Executing: UPDATE pgadmin_param SET param_value = '" & Item & "' WHERE param_id = 2"
  gConnection.Execute "UPDATE pgadmin_param SET param_value = '" & Item & "' WHERE param_id = 2"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTracking, vssTracking_ItemSelected"
End Sub

