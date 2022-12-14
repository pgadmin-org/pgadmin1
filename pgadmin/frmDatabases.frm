VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmDatabases 
   Caption         =   "Databases"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmDatabases.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdSystemDSN 
      Caption         =   "&System DSN"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Click to create a System DSN for the selected database."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdUserDSN 
      Caption         =   "&User DSN"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Click to create a User DSN for the selected database."
      Top             =   765
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   16
      Top             =   2205
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Databases"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Select to view system databases."
         Top             =   225
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Edit the comment for the selected database."
      Top             =   1485
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Database Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   12
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtUsername 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1035
         Width           =   2625
      End
      Begin VB.TextBox txtTimeOut 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1350
         Width           =   2625
      End
      Begin VB.TextBox txtCompiler 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2295
         Width           =   2625
      End
      Begin VB.TextBox txtPlatform 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1980
         Width           =   2625
      End
      Begin VB.TextBox txtdbVer 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1665
         Width           =   2625
      End
      Begin HighlightBox.HBX txtComments 
         Height          =   1230
         Left            =   90
         TabIndex        =   11
         Top             =   2700
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2170
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
         Caption         =   "Comments"
      End
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000F&
         Height          =   240
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   765
         Width           =   2625
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   240
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   495
         Width           =   2625
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   240
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   26
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Timeout"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   25
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Compiler"
         Height          =   195
         Index           =   23
         Left            =   90
         TabIndex        =   24
         Top             =   2340
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   23
         Top             =   2025
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DBMS"
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   22
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   810
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   540
         Width           =   465
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
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   5
      ToolTipText     =   "Refresh the list."
      Top             =   1845
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropdb 
      Caption         =   "&Drop db"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Drop the selected database."
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreatedb 
      Caption         =   "&Create db"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new database."
      Top             =   45
      Width           =   1410
   End
   Begin VB.ListBox lstDB 
      Height          =   3960
      Left            =   1530
      TabIndex        =   7
      ToolTipText     =   "Lists the databases on the current server."
      Top             =   45
      Width           =   2940
   End
End
Attribute VB_Name = "frmDatabases"
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
Dim rsDatabases As New Recordset
Dim rsVersion As New Recordset


Public Sub cmdSystemDSN_Click()
On Error GoTo Err_Handler
Dim iRet As Long
Dim szAttributes As String
Dim szName As String
  If lstDB.Text = "" Then
    MsgBox "You must select a database!", vbExclamation, "Error"
    lstDB.SetFocus
    Exit Sub
  End If
  szName = InputBox("Enter a name for the new System DSN:", "Create System DSN", lstDB.Text)
  If DSN_Exists(szName) = True Then
    MsgBox "A DSN with this name already exists!", vbExclamation, "Error"
    Exit Sub
  End If

  szAttributes = "DSN=" & szName & Chr(0) & _
      "Servername=" & DSNServer & Chr(0) & _
      "Port=" & DSNPort & Chr(0) & _
      "Database=" & lstDB.Text & Chr(0) & _
      "ReadOnly=0" & Chr(0) & _
      "Protocol=6.4" & Chr(0) & _
      "ShowOidColumn=1" & Chr(0) & _
      "FakeOidIndex=1" & Chr(0) & _
      "RowVersioning=0" & Chr(0)

  iRet = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, "PostgreSQL", szAttributes)
  If DSN_Exists(szName) = False Then
    LogMsg "Failed to create System DSN: " & szName
    MsgBox "System DSN creation failed!", vbExclamation, "Error"
  Else
    LogMsg "Created System DSN: " & szName
    MsgBox "System DSN successfully created!", vbExclamation, "Error"
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, cmdSystemDSN_Click"
End Sub

Public Sub cmdUserDSN_Click()
On Error GoTo Err_Handler
Dim iRet As Long
Dim szAttributes As String
Dim szName As String
  If lstDB.Text = "" Then
    MsgBox "You must select a database!", vbExclamation, "Error"
    lstDB.SetFocus
    Exit Sub
  End If
  szName = InputBox("Enter a name for the new User DSN:", "Create User DSN", lstDB.Text)
  If DSN_Exists(szName) = True Then
    MsgBox "A DSN with this name already exists!", vbExclamation, "Error"
    Exit Sub
  End If

  szAttributes = "DSN=" & szName & Chr(0) & _
      "Servername=" & DSNServer & Chr(0) & _
      "Port=" & DSNPort & Chr(0) & _
      "Database=" & lstDB.Text & Chr(0) & _
      "ReadOnly=0" & Chr(0) & _
      "Protocol=6.4" & Chr(0) & _
      "ShowOidColumn=1" & Chr(0) & _
      "FakeOidIndex=1" & Chr(0) & _
      "RowVersioning=0" & Chr(0)

  iRet = SQLConfigDataSource(0&, ODBC_ADD_DSN, "PostgreSQL", szAttributes)
  If DSN_Exists(szName) = False Then
    LogMsg "Failed to create User DSN: " & szName
    MsgBox "User DSN creation failed!", vbExclamation, "Error"
  Else
    LogMsg "Created User DSN: " & szName
    MsgBox "User DSN successfully created!", vbExclamation, "Error"
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, cmdUserDSN_Click"
End Sub

Private Sub lstDB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXDatabases
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, lstDB_MouseUp"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, chkSystem_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsDatabases = Nothing
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If txtOID.Text = "" Then
    MsgBox "You must select a database to edit the comment for!", vbExclamation, "Error"
    Exit Sub
  End If
  Load frmComments
  frmComments.Setup "frmDatabases", QUOTE & lstDB.Text & QUOTE, Val(txtOID.Text)
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, cmdComment_Click"
End Sub

Public Sub cmdCreatedb_Click()
On Error GoTo Err_Handler
  Load frmAddDatabase
  frmAddDatabase.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, cmdCreatedb_Click"
End Sub

Public Sub cmdDropdb_Click()
On Error GoTo Err_Handler

  If lstDB.Text = "" Then
    MsgBox "You must select a database to delete!", vbExclamation, "Error"
    Exit Sub
  End If
  If UCase(lstDB.Text) = "TEMPLATE1" Then
    MsgBox "You cannot delete the template1 database!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to delete this database?" & vbCrLf & "Database name: " & lstDB.Text, vbYesNo + vbQuestion, "Delete Database?") = vbYes Then
    StartMsg "Dropping Database..."
    fMainForm.txtSQLPane.Text = "DROP DATABASE " & QUOTE & lstDB.Text & QUOTE
    LogMsg "Executing: DROP DATABASE " & QUOTE & lstDB.Text & QUOTE
    gConnection.Execute " DROP DATABASE " & QUOTE & lstDB.Text & QUOTE
    cmdRefresh_Click
    EndMsg
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmDatabases, cmdDropdb_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  StartMsg "Retrieving database names..."
  If rsDatabases.State <> adStateClosed Then rsDatabases.Close
  lstDB.Clear
  txtOID.Text = ""
  txtPath.Text = ""
  txtOwner.Text = ""
  txtComments.Text = ""
  If chkSystem.Value = 1 Then
    LogMsg "Executing: SELECT * FROM pgadmin_databases ORDER BY database_name"
    rsDatabases.Open "SELECT * FROM pgadmin_databases ORDER BY database_name", gConnection, adOpenDynamic
  Else
    LogMsg "Executing: SELECT * FROM pgadmin_databases WHERE database_name NOT LIKE 'pgadmin_%' AND database_name NOT LIKE 'pg_%' AND database_oid > " & LAST_SYSTEM_OID & " AND database_name != 'template0' ORDER BY database_name"
    rsDatabases.Open "SELECT * FROM pgadmin_databases WHERE database_name NOT LIKE 'pgadmin_%' AND database_name NOT LIKE 'pg_%' AND database_oid > " & LAST_SYSTEM_OID & " AND database_name != 'template0' ORDER BY database_name", gConnection, adOpenDynamic
  End If
  While Not rsDatabases.EOF
    lstDB.AddItem rsDatabases!database_name & ""
    rsDatabases.MoveNext
  Wend
  If rsDatabases.BOF <> True Then rsDatabases.MoveFirst
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmDatabases, cmdRefresh_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtComments.Minimise
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.Width < 8325 Then Me.Width = 8325
    If Me.Height < 4455 Then Me.Height = 4455
    lstDB.Height = Me.ScaleHeight
    lstDB.Width = Me.ScaleWidth - lstDB.Left - fraDetails.Width - 25
    fraDetails.Left = lstDB.Left + lstDB.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8325
  Me.Height = 4455
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, Form_Load"
End Sub

Public Sub lstDB_Click()
On Error GoTo Err_Handler
  If lstDB.Text <> "" Then
    StartMsg "Retrieving database info..."
    If rsDatabases.BOF <> True Then rsDatabases.MoveFirst
    MoveRS rsDatabases, lstDB.ListIndex
    txtOID.Text = rsDatabases!database_oid & ""
    txtOwner.Text = rsDatabases!database_owner & ""
    txtPath.Text = rsDatabases!database_path & ""
    txtComments.Text = rsDatabases!database_comments & ""
    
    rsVersion.Open "SELECT version()", gConnection, adOpenForwardOnly
    txtdbVer.Text = Mid(rsVersion!Version, 1, InStr(1, rsVersion!Version, " on ") - 1)
    txtPlatform.Text = Mid(rsVersion!Version, InStr(1, rsVersion!Version, " on") + 4, InStr(1, rsVersion!Version, ", compiled by ") - InStr(1, rsVersion!Version, " on") - 4)
    txtCompiler.Text = Mid(rsVersion!Version, InStr(1, rsVersion!Version, ", compiled by ") + 14, Len(rsVersion!Version))
    rsVersion.Close
    
    txtUsername.Text = Username
    txtTimeOut.Text = gConnection.ConnectionTimeout
    
    EndMsg
  Else
    txtOID.Text = ""
    txtOwner.Text = ""
    txtPath.Text = ""
    txtComments.Text = ""
    txtdbVer.Text = ""
    txtPlatform.Text = ""
    txtCompiler.Text = ""
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmDatabases, lstDB_Click"
End Sub

