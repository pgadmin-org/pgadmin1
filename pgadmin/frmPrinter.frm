VERSION 5.00
Begin VB.Form frmPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Printer"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   Icon            =   "frmPrinter.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      ToolTipText     =   "Accept the selected printer."
      Top             =   2790
      Width           =   1005
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   "Accept the highlighted printer."
      Top             =   2790
      Width           =   1005
   End
   Begin VB.ListBox lstPrinters 
      Height          =   2010
      Left            =   45
      TabIndex        =   2
      Top             =   720
      Width           =   4560
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Printer"
      Height          =   555
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4560
      Begin VB.Label lblCurrent 
         Caption         =   "No Current Printer Selected."
         Height          =   240
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   4380
      End
   End
End
Attribute VB_Name = "frmPrinter"
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
Dim X As Integer
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Printer", ValString, lblCurrent.Caption
  For X = 0 To lstPrinters.ListCount - 1
    If lstPrinters.List(X) = lblCurrent.Caption Then
      Set Printer = Printers(X)
      Exit For
    End If
  Next
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, cmdOK_Click"
End Sub

Private Sub cmdSelect_Click()
On Error GoTo Err_Handler
  If lstPrinters.Text = "" Then
    MsgBox "You must select a printer!", vbExclamation, "Error"
    Exit Sub
  End If
  lblCurrent.Caption = lstPrinters.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, cmdSelect_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim X As Printer
  Me.Width = 4770
  Me.Height = 3600
  For Each X In Printers
    lstPrinters.AddItem X.DeviceName
  Next
  lblCurrent.Caption = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Printer", "No Default Printer Currently Selected")
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4770 Then Me.Width = 4770
    If Me.Height < 3600 Then Me.Height = 3600
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, Form_Load"
End Sub

Private Sub lstPrinters_DblClick()
On Error GoTo Err_Handler
  If lstPrinters.Text = "" Then
    Exit Sub
  End If
  lblCurrent.Caption = lstPrinters.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmPrinter, lstPrinters_DblClick"
End Sub
