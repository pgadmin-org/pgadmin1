VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   3330
      TabIndex        =   1
      Top             =   2835
      Width           =   1320
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4789
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Logging"
      TabPicture(0)   =   "frmOptions.frx":128A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkEnableLogging"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkMaskPassword"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtLogfile"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBrowse"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CommonDialog1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   45
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   330
         Left            =   3690
         TabIndex        =   6
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtLogfile 
         Height          =   285
         Left            =   495
         TabIndex        =   4
         Top             =   1845
         Width           =   3165
      End
      Begin VB.CheckBox chkMaskPassword 
         Caption         =   "&Mask password in logfile"
         Height          =   195
         Left            =   495
         TabIndex        =   3
         Top             =   1260
         Width           =   3435
      End
      Begin VB.CheckBox chkEnableLogging 
         Caption         =   "&Enable advanced logging"
         Height          =   195
         Left            =   495
         TabIndex        =   2
         Top             =   810
         Width           =   3435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Logfile"
         Height          =   195
         Left            =   495
         TabIndex        =   5
         Top             =   1620
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmOptions"
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
    .FileName = txtLogfile.Text
    .CancelError = True
    .Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    .Filter = "Log Files (*.log)|*.log"
    .ShowSave
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtLogfile.Text = CommonDialog1.FileName
  Exit Sub
Err_Handler: If Err.Number <> 0 And Err.Number <> 32755 Then LogError Err, "frmOptions, cmdBrowse_click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
  Logging = chkEnableLogging.Value
  MaskPassword = chkMaskPassword.Value
  LogFile = txtLogfile.Text
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Logging", ValString, CStr(Logging)
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Mask Password", ValString, CStr(MaskPassword)
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Log File", ValString, CStr(LogFile)
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmOptions, cmdOK_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim X As Printer
  Me.Width = 4770
  Me.Height = 3570
  chkEnableLogging.Value = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Logging", 0)
  chkMaskPassword.Value = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Mask Password", 1)
  txtLogfile.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Log File", "C:\pgAdmin.log")
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, Form_Load"
End Sub
