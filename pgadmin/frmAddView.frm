VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddView 
   Caption         =   "Create View"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "frmAddView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   4650
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2070
      Top             =   2745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Query"
      Height          =   375
      Left            =   45
      TabIndex        =   5
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create View"
      Height          =   375
      Left            =   3330
      TabIndex        =   4
      Top             =   2880
      Width           =   1275
   End
   Begin VB.TextBox txtSQL 
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   630
      Width           =   4605
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   45
      Width           =   3615
   End
   Begin VB.Label lblSQL 
      AutoSize        =   -1  'True
      Caption         =   "SQL Query"
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Width           =   780
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "View Name"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   810
   End
End
Attribute VB_Name = "frmAddView"
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

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the View!", vbExclamation, "Error"
    Exit Sub
  End If
  If txtSQL.Text = "" Then
    MsgBox "You must enter an SQL query for the View!", vbExclamation, "Error"
    Exit Sub
  End If
  StartMsg "Creating View..."
  LogMsg "Executing: CREATE VIEW " & txtName.Text & " AS " & txtSQL.Text
  gConnection.Execute "CREATE VIEW " & txtName.Text & " AS " & txtSQL.Text
  frmViews.cmdRefresh_Click
  EndMsg
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddView, cmdCreate_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
  fMainForm.txtSQLPane.Text = "CREATE VIEW " & txtName.Text & vbCrLf & "  AS " & vbCrLf & txtSQL.Text
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddView, Gen_SQL"
End Sub

Private Sub cmdLoad_Click()
On Error GoTo Err_Handler
Dim DataLine As String
Dim fNum As Integer
  With CommonDialog1
    .DialogTitle = "Load SQL Query"
    .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    .Filter = "SQL Scripts (*.sql)|*.sql|All Files (*.*)|*.*"
    .FileName = ""
    .CancelError = True
    .ShowOpen
  End With
  If CommonDialog1.FileName = "" Then Exit Sub
  txtSQL.Text = ""
  fNum = FreeFile
  LogMsg "Loading " & CommonDialog1.FileName
  Open CommonDialog1.FileName For Input As #fNum
  While Not EOF(fNum)
    Line Input #fNum, DataLine
    txtSQL.Text = txtSQL.Text & DataLine & vbCrLf
  Wend
  Close #fNum
  Exit Sub
Err_Handler: If Err.Number <> 0 And Err.Number <> 32755 Then LogError Err, "frmSQL, cmdLoad_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Height < 3675 Then Me.Height = 3675
    If Me.Width < 4770 Then Me.Width = 4770
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddView, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 3675
  Me.Width = 4770
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddView, Form_Load"
End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddView, txtName_Change"
End Sub

Private Sub txtSQL_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddView, txtSQL_Change"
End Sub
