VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5550
   ClientLeft      =   2385
   ClientTop       =   870
   ClientWidth     =   5220
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLicence 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0C0&
      Height          =   1275
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmAbout.frx":030A
      Top             =   2745
      Width           =   4920
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3870
      TabIndex        =   0
      Top             =   5130
      Width           =   1260
   End
   Begin VB.Label lblGfxURL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.lpk-computers.co.uk/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   945
      TabIndex        =   9
      Top             =   4725
      Width           =   3150
   End
   Begin VB.Label lblGfxEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "max@lpk-computers.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   945
      TabIndex        =   8
      Top             =   4455
      Width           =   3150
   End
   Begin VB.Label lblGraphics 
      BackStyle       =   0  'Transparent
      Caption         =   "Many thanks to Max at LPK for the logo and splash screen:"
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   4185
      Width           =   4605
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.greatbridge.org/project/pgadmin/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   945
      TabIndex        =   6
      Top             =   2160
      Width           =   3330
   End
   Begin VB.Label lblLiability 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pgAdmin is released under the GNU Public License (GPL):"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   2520
      Width           =   4110
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pgadmin-support@greatbridge.org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   1035
      TabIndex        =   4
      Top             =   1890
      Width           =   3150
   End
   Begin VB.Label lblBugs 
      BackStyle       =   0  'Transparent
      Caption         =   "Please reports any bugs or faults to:"
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   135
      TabIndex        =   3
      Top             =   1530
      Width           =   5940
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   4272
      TabIndex        =   2
      Top             =   948
      Width           =   780
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1998-2001, Dave Page"
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   1215
      Width           =   5985
   End
   Begin VB.Image imgpgAdmin 
      Height          =   1230
      Left            =   0
      Picture         =   "frmAbout.frx":0639
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   5220
   End
End
Attribute VB_Name = "frmAbout"
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
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, cmdOK_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  If DEVELOPMENT Then
    lblVersion.Caption = "Version " & app.Major & "." & app.Minor & "." & app.Revision & " DEV"
  Else
    lblVersion.Caption = "Version " & app.Major & "." & app.Minor & "." & app.Revision
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, Form_Load"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  lblURL.ForeColor = &HFFC0C0
  lblEmail.ForeColor = &HFFC0C0
  lblGfxURL.ForeColor = &HFFC0C0
  lblGfxEmail.ForeColor = &HFFC0C0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, Form_MouseMove"
End Sub

Private Sub lblEmail_Click()
On Error GoTo Err_Handler
  StartURL "mailto:pgadmin-support@greatbridge.org"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblEmail_Click"
End Sub

Private Sub lblURL_Click()
On Error GoTo Err_Handler
  StartURL "http://www.greatbridge.org/project/pgadmin/"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblEmail_Click"
End Sub

Private Sub lblGfxEmail_Click()
On Error GoTo Err_Handler
  StartURL "mailto:max@lpk-computers.co.uk"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblGfxEmail_Click"
End Sub

Private Sub lblGfxURL_Click()
On Error GoTo Err_Handler
  StartURL "http://www.lpk-computers.co.uk/"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblGfxEmail_Click"
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  lblEmail.ForeColor = &HC0C0FF
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblEmail_MouseMove"
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  lblURL.ForeColor = &HC0C0FF
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblURL_MouseMove"
End Sub

Private Sub lblGfxEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  lblGfxEmail.ForeColor = &HC0C0FF
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblGfxEmail_MouseMove"
End Sub

Private Sub lblGfxURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  lblGfxURL.ForeColor = &HC0C0FF
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAbout, lblGfxURL_MouseMove"
End Sub
