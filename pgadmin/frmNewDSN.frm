VERSION 5.00
Begin VB.Form frmNewDSN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New PostgreSQL DSN"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmNewDSN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create &System DSN"
      Height          =   330
      Index           =   1
      Left            =   2745
      TabIndex        =   5
      ToolTipText     =   "Create a new System DSN."
      Top             =   1575
      Width           =   1725
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create &User DSN"
      Height          =   330
      Index           =   0
      Left            =   945
      TabIndex        =   4
      ToolTipText     =   "Create a new User DSN."
      Top             =   1575
      Width           =   1725
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "5432"
      ToolTipText     =   "Enter the port number that the PostgreSQL server is listening on."
      Top             =   1170
      Width           =   915
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Enter the name of the Database on the server."
      Top             =   810
      Width           =   3390
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      ToolTipText     =   "Enter the hostname or IP address of the PostgreSQL server."
      Top             =   450
      Width           =   3390
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Enter a name for the new Datasource."
      Top             =   90
      Width           =   3390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   9
      Top             =   180
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   1260
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   900
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   540
      Width           =   465
   End
End
Attribute VB_Name = "frmNewDSN"
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

Private Sub cmdCreate_Click(Index As Integer)
On Error GoTo Err_Handler
Dim iRet As Long
Dim szAttributes As String
Dim X As Integer
Dim bCreated As Boolean
  If txtName.Text = "" Then
    MsgBox "You must specify a name for the datasource!", vbExclamation, "Error"
    txtName.SetFocus
    Exit Sub
  End If
  If txtServer.Text = "" Then
    MsgBox "You must specify a server!", vbExclamation, "Error"
    txtServer.SetFocus
    Exit Sub
  End If
  If txtDatabase.Text = "" Then
    MsgBox "You must specify a database!", vbExclamation, "Error"
    txtDatabase.SetFocus
    Exit Sub
  End If
  If Validate(txtPort.Text, vdtInteger, False) = False Then
    MsgBox "You must specify an integer value for the port!", vbExclamation, "Error"
    txtPort.SetFocus
    Exit Sub
  End If
  If DSN_Exists(txtName.Text) = True Then
    MsgBox "A DSN with this name already exists!", vbExclamation, "Error"
    txtName.SetFocus
    Exit Sub
  End If

  szAttributes = "DSN=" & txtName.Text & Chr(0) & _
      "Servername=" & txtServer.Text & Chr(0) & _
      "Port=" & txtPort.Text & Chr(0) & _
      "Database=" & txtDatabase.Text & Chr(0) & _
      "ReadOnly=0" & Chr(0) & _
      "Protocol=6.4" & Chr(0) & _
      "ShowOidColumn=1" & Chr(0) & _
      "FakeOidIndex=1" & Chr(0) & _
      "RowVersioning=0" & Chr(0)

  If Index = 0 Then
    iRet = SQLConfigDataSource(0&, ODBC_ADD_DSN, "PostgreSQL", szAttributes)
  Else
    iRet = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, "PostgreSQL", szAttributes)
  End If
  frmODBCLogon.GetDSNsAndDrivers
  bCreated = False
  For X = 0 To frmODBCLogon.cboDSNList.ListCount - 1
    If frmODBCLogon.cboDSNList.List(X) = txtName.Text Then
      frmODBCLogon.cboDSNList.ListIndex = X
      bCreated = True
      Exit For
    End If
  Next
  If bCreated = False Then
    MsgBox "DSN creation failed - please check the options entered and try again.", vbExclamation, "Error"
    txtName.SetFocus
    Exit Sub
  End If
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmNewDSN, cmdUser_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4650
  Me.Height = 2325
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmNewDSN, Form_Load"
End Sub

