VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmQueryPlan 
   Caption         =   "Query Plan"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "frmQueryPlan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   5670
   Begin HighlightBox.HBX txtQuery 
      Height          =   1860
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Displays the SQL Query."
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3281
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Caption         =   "SQL Query"
   End
   Begin HighlightBox.HBX txtPlan 
      Height          =   1860
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Displays the Query Execution Plan."
      Top             =   1890
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   3281
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Caption         =   "Query Execution Plan"
   End
End
Attribute VB_Name = "frmQueryPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgadmin - PostgreSQL db Administration/Management for Win32
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

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 5790
  Me.Height = 4200
  txtQuery.Wordlist = TextColours
  txtPlan.Wordlist = TextColours
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmQueryPlan, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtQuery.Minimise
  txtPlan.Minimise
  If Me.WindowState = 0 Then
    If Me.Width < 5790 Then Me.Width = 5790
    If Me.Height < 4200 Then Me.Height = 4200
    txtQuery.Width = Me.ScaleWidth
    txtPlan.Width = Me.ScaleWidth
    txtQuery.Height = (Me.ScaleHeight / 5) * 2
    txtPlan.Height = ((Me.ScaleHeight / 5) * 3) - 50
    txtPlan.Top = txtQuery.Height + 50
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmQueryPlan, Form_Resize"
End Sub

Public Sub Explain(szSQL As String)
On Error GoTo Err_Handler
Dim lEnv As Long
Dim lDBC As Long
Dim lRet As Long
Dim lStmt As Long
Dim lErr As Long
Dim iSize As Integer
Dim szResult As String * 256
Dim szSqlState As String * 1024
Dim szErrorMsg As String * 1024
Dim szPlan As String

  txtQuery.Text = szSQL
  txtQuery.ColourText

  'Execute the statement. In theory, the ADO connection object can does this, and the plan can
  'be picked up as a series of 512Byte strings in the Errors collection. This is unreliable though, so
  'we'll use ODBC directly insted <shudder>
  
  StartMsg "Requesting Query Execution Plan..."
  
  'Initialisze the ODBC subsystem
  If SQLAllocEnv(lEnv) <> 0 Then
    LogMsg "Unable to initialize ODBC API drivers!"
    MsgBox "Unable to initialize ODBC API drivers!", vbCritical, "Error"
    GoTo Cleanup
  End If

  If SQLAllocConnect(lEnv, lDBC) <> 0 Then
    LogMsg "Could not allocate memory for connection Handle!"
    MsgBox "Could not allocate memory for connection Handle!", vbCritical, "Error"
    GoTo Cleanup
  End If

  lRet = SQLDriverConnect(lDBC, Me.hWnd, Connect, Len(Connect), szResult, Len(szResult), iSize, 1)
  If lRet <> SQL_SUCCESS Then
    LogMsg "Could not establish connection to ODBC driver! Error: " & lRet
    MsgBox "Could not establish connection to ODBC driver!" & vbCrLf & "Error: " & lRet, vbCritical, "Error"
    GoTo Cleanup
  End If
  
  'Check the ODBC Driver version. EXPLAIN will only work with 07.01.0006 or higher.
  SQLGetInfoString lDBC, SQL_DBMS_VER, szResult, Len(szResult), vbNull
  LogMsg "ODBC Driver Version: " & szResult
  If Val(Mid(szResult, 1, 2)) < 7 Then
     LogMsg "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)"
     MsgBox "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", vbExclamation, "Error"
     GoTo Cleanup
  Else
    If Val(Mid(szResult, 4, 2)) < 1 Then
      LogMsg "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)"
      MsgBox "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", vbExclamation, "Error"
      GoTo Cleanup
    Else
      If Val(Mid(szResult, 7, 4)) < 6 Then
        LogMsg "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)"
        MsgBox "The installed ODBC driver is not the required version or higher (psqlODBC 07.01.0006)", vbExclamation, "Error"
        GoTo Cleanup
      End If
    End If
  End If
  
  'Allocate memory for the statement handle.
  If SQLAllocStmt(lDBC, lStmt) <> 0 Then
    LogMsg "Could not allocate memory for a statement handle!"
    MsgBox "Could not allocate memory for a statement handle!", vbCritical, "Error"
    Exit Sub
  End If
  
  szSQL = "EXPLAIN " & szSQL
  LogMsg "SQLExecDirect: " & szSQL
  If SQLExecDirect(lStmt, szSQL, Len(szSQL)) = SQL_SUCCESS_WITH_INFO Then
    While SQLError(lEnv, lDBC, lStmt, szSqlState, lErr, szErrorMsg, 1024, iSize) <> SQL_NO_DATA_FOUND
      If iSize > 512 Then iSize = 512
      szPlan = szPlan & Left(szErrorMsg, iSize)
    Wend
  End If
  
  If Len(szPlan) > 22 Then szPlan = Mid(szPlan, 23)
  If szPlan <> "" Then
    txtPlan.Text = szPlan
    txtPlan.ColourText
  Else
    LogMsg "A Query Execution Plan could not be calculated for the specified SQL query."
    txtPlan.Text = "A Query Execution Plan could not be calculated for the specified SQL query."
  End If

Cleanup:
  'Log out and cleanup
  If lDBC <> 0 Then
    SQLDisconnect lDBC
  End If
  SQLFreeConnect lDBC
  If lEnv <> 0 Then
    SQLFreeEnv lEnv
  End If
 
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmQueryPlan, Explain"
End Sub
