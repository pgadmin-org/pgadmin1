VERSION 5.00
Begin VB.Form frmSequences 
   Caption         =   "Sequences"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmSequences.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   26
      Top             =   1485
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Sequences"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Select to view system sequences."
         Top             =   225
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected sequence."
      Top             =   765
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Sequence Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   16
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtACL 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   855
         Width           =   2670
      End
      Begin VB.TextBox txtCycled 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2745
         Width           =   2670
      End
      Begin VB.TextBox txtCache 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2430
         Width           =   2670
      End
      Begin VB.TextBox txtMinimum 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2115
         Width           =   2670
      End
      Begin VB.TextBox txtMaximum 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   2670
      End
      Begin VB.TextBox txtIncrement 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1485
         Width           =   2670
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   2670
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txtLastValue 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1170
         Width           =   2670
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   600
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   3330
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACL"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   27
         Top             =   900
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Value"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   1215
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Increment"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   22
         Top             =   1530
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Maximum"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   21
         Top             =   1845
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Minimum"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   20
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cache"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   19
         Top             =   2475
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cycled"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   18
         Top             =   2790
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   17
         Top             =   3105
         Width           =   735
      End
   End
   Begin VB.ListBox lstSeq 
      Height          =   3960
      Left            =   1485
      TabIndex        =   5
      Top             =   45
      Width           =   2985
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Refresh the list of sequences."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropSeq 
      Caption         =   "&Drop Sequence"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected sequence."
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateSeq 
      Caption         =   "&Create Sequence"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new sequence."
      Top             =   45
      Width           =   1410
   End
End
Attribute VB_Name = "frmSequences"
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
Dim rsSeq As New Recordset

Private Sub lstSeq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXSequences
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSequences, lstSeq_MouseUp"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSequences, chkSystem_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsSeq = Nothing
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If txtOID.Text = "" Then
    MsgBox "You must select a sequence to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  Load frmComments
  frmComments.Setup "frmSequences", QUOTE & lstSeq.Text & QUOTE, Val(txtOID.Text)
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSequences, cmdComment_Click"
End Sub

Public Sub cmdCreateSeq_Click()
On Error GoTo Err_Handler
  Load frmAddSequence
  frmAddSequence.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSequences, cmdCreateSeq_Click"
End Sub

Public Sub cmdDropSeq_Click()
On Error GoTo Err_Handler
  If lstSeq.Text = "" Then
    MsgBox "You must select a sequence to delete!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to delete this sequence?", vbYesNo + vbQuestion, _
            "Confirm Sequence Delete") = vbYes Then
    fMainForm.txtSQLPane.Text = "DROP SEQUENCE " & QUOTE & lstSeq.Text & QUOTE
    LogMsg "DROP SEQUENCE " & QUOTE & lstSeq.Text & QUOTE
    gConnection.Execute "DROP SEQUENCE " & QUOTE & lstSeq.Text & QUOTE
    LogQuery "DROP SEQUENCE " & QUOTE & lstSeq.Text & QUOTE
    cmdRefresh_Click
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSequences, cmdDropSeq_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  StartMsg "Retrieving Sequence Names..."
  lstSeq.Clear
  txtOID.Text = ""
  txtCache.Text = ""
  txtComments.Text = ""
  txtCycled.Text = ""
  txtIncrement.Text = ""
  txtLastValue.Text = ""
  txtMaximum.Text = ""
  txtMinimum.Text = ""
  txtOwner.Text = ""
  If rsSeq.State <> adStateClosed Then rsSeq.Close
  If chkSystem.Value = 1 Then
    LogMsg "Executing: SELECT oid, relname, pg_get_userbyid(relowner) AS usename, relacl FROM pg_class AND relkind = 'S' ORDER BY relname"
    rsSeq.Open "SELECT oid, relname, pg_get_userbyid(relowner) AS usename, relacl FROM pg_class AND relkind = 'S' ORDER BY relname", gConnection, adOpenDynamic
  Else
    LogMsg "Executing: SELECT oid, relname, pg_get_userbyid(relowner) AS usename, relacl FROM pg_class WHERE relname NOT LIKE 'pgadmin_%' AND relname NOT LIKE 'pg_%' AND oid > " & LAST_SYSTEM_OID & " AND relkind = 'S' ORDER BY relname"
    rsSeq.Open "SELECT oid, relname, pg_get_userbyid(relowner) AS usename, relacl FROM pg_class WHERE relname NOT LIKE 'pgadmin_%' AND relname NOT LIKE 'pg_%' AND oid > " & LAST_SYSTEM_OID & " AND relkind = 'S' ORDER BY relname", gConnection, adOpenDynamic
  End If
  While Not rsSeq.EOF
    lstSeq.AddItem rsSeq!relname
    rsSeq.MoveNext
  Wend
  If rsSeq.BOF <> True Then rsSeq.MoveFirst
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmSequences, cmdRefresh_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = 8325
  Me.Height = 4455
  LogMsg "Loading Form: " & Me.Name
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSequences, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4455 Then Me.Height = 4455
    End If
    lstSeq.Height = Me.ScaleHeight
    lstSeq.Width = Me.ScaleWidth - lstSeq.Left - fraDetails.Width - 25
    fraDetails.Left = lstSeq.Left + lstSeq.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSequences, Form_Resize"
End Sub

Public Sub lstSeq_Click()
On Error GoTo Err_Handler
Dim rsInfo As New Recordset
  If lstSeq.Text <> "" Then
    StartMsg "Retrieving Sequence Info..."
    If rsSeq.BOF <> True Then rsSeq.MoveFirst
    MoveRS rsSeq, lstSeq.ListIndex
    txtOID.Text = rsSeq!OID & ""
    txtOwner.Text = rsSeq!usename & ""
    txtACL.Text = rsSeq!relacl & ""
    If rsInfo.State <> adStateClosed Then rsInfo.Close
    LogMsg "Executing: SELECT description FROM pg_description WHERE objoid = " & rsSeq!OID
    rsInfo.Open "SELECT description FROM pg_description WHERE objoid = " & rsSeq!OID, gConnection, adOpenForwardOnly
    If Not rsInfo.EOF Then
      txtComments.Text = rsInfo!Description & ""
    Else
      txtComments.Text = ""
    End If
    If rsSeq.BOF <> True Then rsSeq.MoveFirst
    If rsInfo.State <> adStateClosed Then rsInfo.Close
    LogMsg "Executing: SELECT last_value, increment_by, max_value, min_value, cache_value, is_cycled FROM " & QUOTE & lstSeq.Text & QUOTE
    rsInfo.Open "SELECT last_value, increment_by, max_value, min_value, cache_value, is_cycled FROM " & QUOTE & lstSeq.Text & QUOTE, gConnection, adOpenForwardOnly
    txtLastValue.Text = rsInfo!last_value & ""
    txtIncrement.Text = rsInfo!increment_by & ""
    txtMaximum.Text = rsInfo!max_value & ""
    txtMinimum.Text = rsInfo!min_value & ""
    txtCache.Text = rsInfo!cache_value & ""
    If rsInfo!is_cycled & "" = "t" Then
      txtCycled.Text = "Yes"
    Else
      txtCycled.Text = "No"
    End If
    EndMsg
  End If
  Set rsInfo = Nothing
  Exit Sub
Err_Handler:
  Set rsInfo = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmSequences, lstSeq_Click"
End Sub
