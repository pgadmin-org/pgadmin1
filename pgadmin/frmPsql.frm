VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPsql 
   Caption         =   "Psql"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   Icon            =   "frmPsql.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   5175
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   1125
      TabIndex        =   7
      Top             =   945
      Width           =   2805
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1125
      TabIndex        =   6
      Top             =   315
      Width           =   2805
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   330
      Left            =   3960
      TabIndex        =   4
      Top             =   2610
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuration"
      Height          =   1140
      Left            =   0
      TabIndex        =   0
      Top             =   1395
      Width           =   5145
      Begin VB.CheckBox chkDebug 
         Caption         =   "&Debug Mode"
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   810
         Width           =   2220
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   4725
         TabIndex        =   2
         Top             =   450
         Width           =   285
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Text            =   "psql.exe"
         Top             =   450
         Width           =   4605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Path to Psql EXE file"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Database Name"
      Height          =   195
      Left            =   1125
      TabIndex        =   8
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Database Server Hostname/IP Address"
      Height          =   195
      Left            =   1125
      TabIndex        =   5
      Top             =   90
      Width           =   2805
   End
End
Attribute VB_Name = "frmPsql"
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

Private Sub cmdExecute_Click()
On Error GoTo Err_Handler
Dim Scr_hDC As Long
Dim x As Long
  Scr_hDC = GetDesktopWindow()
  LogMsg "Executing psql.exe..."
  If chkDebug.Value = 0 Then
    x = ShellExecute(Scr_hDC, "Open", txtPath.Text, " -U " & Username & " -h " & txtHost.Text & " -d " & txtDatabase.Text, "C:\", SW_SHOWNORMAL)
  Else
    x = ShellExecute(Scr_hDC, "Open", txtPath.Text, " -U " & Username & " -h " & txtHost.Text & " -d " & txtDatabase.Text & " -E", "C:\", SW_SHOWNORMAL)
  End If
  If x <= 32 Then
    MsgBox "An error occured executing psql.exe. Please check the executable path.", vbCritical, "Error!"
    LogMsg "Could not execute psql.exe (Error: " & x & ")."
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPsql, cmdExecute_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  txtHost.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\db Servers", Datasource, gConnection.Properties("Server Name").Value)
  txtDatabase.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\db Names", Datasource, Username)
  txtPath.Text = RegRead(HKEY_LOCAL_MACHINE, "Software\pgAdmin", "Psql Path", "psql.exe")
  Me.Width = 5295
  Me.Height = 3360
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPsql, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 5295 Then Me.Width = 5295
    If Me.Height < 3360 Then Me.Height = 3360
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPsql, Form_Resize"
End Sub

Private Sub txtHost_Change()
On Error GoTo Err_Handler
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin\db Servers", Datasource, ValString, txtHost.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPsql, txtHost_Change"
End Sub

Private Sub txtDatabase_Change()
On Error GoTo Err_Handler
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin\db Names", Datasource, ValString, txtDatabase.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPsql, txtDatabase_Change"
End Sub

Private Sub txtPath_Change()
On Error GoTo Err_Handler
  RegWrite HKEY_LOCAL_MACHINE, "Software\pgAdmin", "Psql Path", ValString, txtPath.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPsql, txtPath_Change"
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo Err_Handler
  With CommonDialog1
    .FileName = txtPath.Text
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .Filter = "Executables (*.exe)|*.exe"
    .ShowOpen
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtPath.Text = CommonDialog1.FileName
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPsql, cmdBrowse_Click"
End Sub
