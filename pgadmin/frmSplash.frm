VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3810
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3810
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      Begin VB.Label lblStatus 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00C0FFFF&
         Height          =   600
         Left            =   45
         TabIndex        =   3
         Top             =   90
         Width           =   5145
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDEV 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DEVELOPMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   0
         TabIndex        =   2
         Top             =   3465
         Width           =   2505
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4140
         TabIndex        =   1
         Top             =   3510
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Err_Handler
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSplash, Form_KeyPress"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = picLogo.Width
  Me.Height = picLogo.Height
  lblStatus.Caption = "Loading pgAdmin version " & app.Major & "." & app.Minor & "." & app.Revision & " - Searching for Exporters..."
  lblVersion.Caption = "Version " & app.Major & "." & app.Minor & "." & app.Revision
  If DEVELOPMENT Then
    lblDEV.Visible = True
  Else
    lblDEV.Visible = False
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSplash, Form_Load"
End Sub

