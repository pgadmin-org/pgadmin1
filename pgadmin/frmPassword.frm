VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3465
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   3465
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2385
      TabIndex        =   6
      Top             =   1350
      Width           =   1050
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   945
      Width           =   1860
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   540
      Width           =   1860
   End
   Begin VB.TextBox txtCurrent 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   90
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Confirm Password"
      Height          =   240
      Index           =   2
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Current Password"
      Height          =   240
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   1545
   End
End
Attribute VB_Name = "frmPassword"
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

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
  If txtCurrent.Text <> Password Then
    MsgBox "Incorrect Password!", vbExclamation, "Error"
    Exit Sub
  End If
  If txtNew.Text <> txtConfirm.Text Then
    MsgBox "Passwords do not match!", vbExclamation, "Error"
    Exit Sub
  End If
  If Len(txtNew.Text) < 6 Then
    MsgBox "Password must be at least 6 characters long!", vbExclamation, "Error"
    Exit Sub
  End If
  If InStr(1, txtNew.Text, " ") Or InStr(1, txtNew.Text, "'") Or InStr(1, txtNew.Text, QUOTE) Then
    MsgBox "Illegal characters in password!", vbExclamation, "Error"
    Exit Sub
  End If
  gConnection.Execute " ALTER USER " & QUOTE & Username & QUOTE & " WITH PASSWORD '" & txtNew.Text & "'"
  MsgBox "Password successfully changed!", vbInformation, "Success!!"
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPassword, cmdOK_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = 3555
  Me.Height = 2130
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPassword, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 3555 Then Me.Width = 3555
    If Me.Height < 2130 Then Me.Height = 2130
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPassword, Form_Resize"
End Sub
