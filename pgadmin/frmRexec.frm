VERSION 5.00
Object = "{0006467F-5D0B-11D2-AD1C-0060978DBC90}#1.0#0"; "vsrexec.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmRexec 
   Caption         =   "Remote Execute"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "frmRexec.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   Begin HighlightBox.HBX txtOP 
      Height          =   1680
      Left            =   0
      TabIndex        =   9
      ToolTipText     =   "This textbox displays the output received from the host."
      Top             =   1485
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   2963
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
      Caption         =   "Command Output"
   End
   Begin vsRExec.VS_rExec vsRexec 
      Left            =   4230
      Top             =   495
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Default         =   -1  'True
      Height          =   375
      Left            =   2745
      TabIndex        =   8
      ToolTipText     =   "Execute the command on the remote host."
      Top             =   540
      Width           =   1410
   End
   Begin VB.TextBox txtCmd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1170
      TabIndex        =   7
      ToolTipText     =   "Enter the command to execute on the remote host."
      Top             =   1125
      Width           =   3075
   End
   Begin VB.TextBox txtPWD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   6
      ToolTipText     =   "Enter a username to log on to the remote host with."
      Top             =   765
      Width           =   1455
   End
   Begin VB.TextBox txtUID 
      Height          =   285
      Left            =   1170
      TabIndex        =   5
      ToolTipText     =   "Enter a username to log on to the remote host with."
      Top             =   405
      Width           =   1455
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1170
      TabIndex        =   4
      ToolTipText     =   "Enter the TCP/IP node name of the remote host."
      Top             =   45
      Width           =   3075
   End
   Begin VB.Label lblCommand 
      Caption         =   "Command:"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   1170
      Width           =   1455
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   810
      Width           =   1455
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username:"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label lblHost 
      Caption         =   "Remote Host:"
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1455
   End
End
Attribute VB_Name = "frmRexec"
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
  StartMsg "Executing command on remote host..."
  vsRexec.HostName = txtHost.Text
  vsRexec.Username = txtUID.Text
  vsRexec.Password = txtPWD.Text
  txtOP.Text = vsRexec.rExec(txtCmd.Text)
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmRexec, cmdExecute_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Handler
  txtCmd.SetFocus
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmRexec, Form_Activate"
End Sub

Private Sub txtHost_Change()
On Error GoTo Err_Handler
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmrExec, txtHost_Change"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4365
  Me.Height = 3075
  txtOP.Wordlist = TextColours
  txtUID.Text = Username
  txtPWD.Text = Password
  txtHost.Text = gConnection.Properties("Server Name").Value
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmRexec, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtOP.Minimise
  If Me.WindowState = 0 Then
    If Me.Width < 4365 Then Me.Width = 4365
    If Me.Height < 3075 Then Me.Height = 3075
  End If
  
  
    txtOP.Width = Me.ScaleWidth
    txtOP.Height = Me.ScaleHeight - txtOP.Top

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmRexec, Form_Resize"
End Sub
