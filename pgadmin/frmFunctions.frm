VERSION 5.00
Begin VB.Form frmFunctions 
   Caption         =   "Functions"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmFunctions.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   21
      Top             =   1485
      Width           =   1380
      Begin VB.CheckBox chkFunctions 
         Caption         =   "Functions"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Select to view system functions."
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected function."
      Top             =   765
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Function Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   13
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtLanguage 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2655
         Width           =   2670
      End
      Begin VB.TextBox txtFunction 
         BackColor       =   &H8000000F&
         Height          =   870
         Left            =   900
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1755
         Width           =   2670
      End
      Begin VB.TextBox txtReturns 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
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
      Begin VB.TextBox txtArguments 
         BackColor       =   &H8000000F&
         Height          =   555
         Left            =   900
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   855
         Width           =   2670
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   735
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   3195
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arguments"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   900
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Returns"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   17
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Language"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   15
         Top             =   2700
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   14
         Top             =   2970
         Width           =   735
      End
   End
   Begin VB.ListBox lstFunc 
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
      ToolTipText     =   "Refresh the list of function."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropFunc 
      Caption         =   "&Drop Function"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected function."
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateFunc 
      Caption         =   "&Create Function"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new function."
      Top             =   45
      Width           =   1410
   End
End
Attribute VB_Name = "frmFunctions"
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
Dim rsFunc As New Recordset

Private Sub lstFunc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXFunctions
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, lstFunc_MouseUp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsFunc = Nothing
End Sub

Private Sub chkFunctions_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, chkFunctions_Click"
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If txtOID.Text = "" Then
    MsgBox "You must select a function to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  CallingForm = "frmFunctions"
  OID = txtOID.Text
  Load frmComments
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdComment_Click"
End Sub

Public Sub cmdCreateFunc_Click()
On Error GoTo Err_Handler
  Load frmAddFunction
  frmAddFunction.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdCreateFunc_Click"
End Sub

Public Sub cmdDropFunc_Click()
On Error GoTo Err_Handler
Dim DropStr As String
  If lstFunc.Text = "" Then
    MsgBox "You must select a function to delete!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to delete this Function?", vbYesNo + vbQuestion, _
            "Confirm Function Delete") = vbYes Then
    DropStr = "DROP FUNCTION " & QUOTE & lstFunc.Text & QUOTE & " (" & txtArguments.Text & ")"
    fMainForm.txtSQLPane.Text = DropStr
    StartMsg "Dropping Function..."
    LogMsg "Executing: " & DropStr
    gConnection.Execute DropStr
    LogQuery DropStr
    cmdRefresh_Click
    EndMsg
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdDropFunc_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  StartMsg "Retrieving Function Names..."
  lstFunc.Clear
  txtOID.Text = ""
  txtArguments.Text = ""
  txtComments.Text = ""
  txtReturns.Text = ""
  txtFunction.Text = ""
  txtLanguage.Text = ""
  txtOwner.Text = ""
  If rsFunc.State <> adStateClosed Then rsFunc.Close
  If chkFunctions.Value = 1 Then
    LogMsg "Executing: SELECT * FROM pgadmin_functions ORDER BY function_name"
    rsFunc.Open "SELECT * FROM pgadmin_functions ORDER BY function_name", gConnection, adOpenDynamic
  Else
    LogMsg "Executing: SELECT * FROM pgadmin_functions WHERE function_name NOT LIKE 'pgadmin_%' AND function_name NOT LIKE 'pg_%' AND function_oid > " & LAST_SYSTEM_OID & " ORDER BY function_name"
    rsFunc.Open "SELECT * FROM pgadmin_functions WHERE function_name NOT LIKE 'pgadmin_%' AND function_name NOT LIKE 'pg_%' AND function_oid > " & LAST_SYSTEM_OID & " ORDER BY function_name", gConnection, adOpenDynamic
  End If
  While Not rsFunc.EOF
    lstFunc.AddItem rsFunc!function_name & ""
    rsFunc.MoveNext
  Wend
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdRefresh_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8325
  Me.Height = 4455
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4455 Then Me.Height = 4455
    End If
    lstFunc.Height = Me.ScaleHeight
    lstFunc.Width = Me.ScaleWidth - lstFunc.Left - fraDetails.Width - 25
    fraDetails.Left = lstFunc.Left + lstFunc.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, Form_Resize"
End Sub

Public Sub lstFunc_Click()
On Error GoTo Err_Handler
  If lstFunc.Text <> "" Then
    StartMsg "Retrieving Function Info..."
    If rsFunc.BOF <> True Then rsFunc.MoveFirst
    MoveRS rsFunc, lstFunc.ListIndex
    txtOID.Text = rsFunc!function_oid & ""
    txtOwner.Text = rsFunc!function_owner & ""
    txtReturns.Text = rsFunc!function_returns & ""
    txtArguments.Text = rsFunc!function_arguments & ""
    txtFunction.Text = rsFunc!function_source & ""
    txtLanguage.Text = rsFunc!function_language & ""
    txtComments.Text = rsFunc!function_comments & ""
    If rsFunc.BOF <> True Then rsFunc.MoveFirst
    EndMsg
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, lstFunc_Click"
End Sub
