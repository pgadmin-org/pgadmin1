VERSION 5.00
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmComments 
   Caption         =   "Comments"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "frmComments.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   4650
   Begin HighlightBox.HBX txtComments 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Enter or Edit the comments for the object."
      Top             =   0
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4154
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ControlBarVisible=   0   'False
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3510
      TabIndex        =   1
      Top             =   2430
      Width           =   1095
   End
End
Attribute VB_Name = "frmComments"
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
Dim szCaller As String
Dim szIdentifier As String
Dim lOID As Long

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
Dim szComment As String
  StartMsg "Updating Comment..."
  szComment = Replace(Replace(txtComments.Text, "\", "\\"), "'", "\'")
    
  Select Case szCaller
    Case "frmDatabases"
      LogMsg "Executing: COMMENT ON DATABASE " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON DATABASE " & szIdentifier & " IS '" & szComment & "'"
      frmDatabases.cmdRefresh_Click
    Case "frmSequences"
      LogMsg "Executing: COMMENT ON SEQUENCE " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON SEQUENCE " & szIdentifier & " IS '" & szComment & "'"
      frmSequences.lstSeq_Click
    Case "frmIndexes"
      LogMsg "Executing: COMMENT ON INDEX " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON INDEX " & szIdentifier & " IS '" & szComment & "'"
      frmIndexes.cmdRefresh_Click
    Case "frmTables"
      LogMsg "Executing: COMMENT ON TABLE " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON TABLE " & szIdentifier & " IS '" & szComment & "'"
      frmTables.cmdRefresh_Click
    Case "frmTables - Column"
      LogMsg "Executing: COMMENT ON COLUMN " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON COLUMN " & szIdentifier & " IS '" & szComment & "'"
      frmTables.cmdRefresh_Click
    Case "frmFunctions"
      LogMsg "Executing: COMMENT ON FUNCTION " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON FUNCTION " & szIdentifier & " IS '" & szComment & "'"
      frmFunctions.cmdRefresh_Click
    Case "frmTriggers"
      LogMsg "Executing: COMMENT ON TRIGGER " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON TRIGGER " & szIdentifier & " IS '" & szComment & "'"
      frmTriggers.cmdRefresh_Click
    Case "frmViews"
      LogMsg "Executing: COMMENT ON VIEW " & szIdentifier & " IS '" & szComment & "'"
      gConnection.Execute "COMMENT ON VIEW " & szIdentifier & " IS '" & szComment & "'"
      frmViews.cmdRefresh_Click
  End Select
  EndMsg
  Unload Me
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmComments, cmdSave_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 5000
  Me.Height = 3500
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmComments, Form_Load"
End Sub

Public Sub Setup(szMCaller As String, szMIdentifier As String, lMOID As Long)
On Error GoTo Err_Handler
Dim rs As New Recordset
  szCaller = szMCaller
  szIdentifier = szMIdentifier
  lOID = lMOID
  LogMsg "Executing: SELECT description FROM pg_description WHERE objoid = " & lOID
  rs.Open "SELECT description FROM pg_description WHERE objoid = " & lOID, gConnection, adOpenDynamic
  If rs.EOF <> True Then
    txtComments.Text = rs!Description
  Else
    txtComments.Text = ""
  End If
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  If Err.Number <> 0 Then LogError Err, "frmComments, Setup"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.Width < 2000 Then Me.Width = 2000
    If Me.Height < 2000 Then Me.Height = 2000
    
    txtComments.Width = Me.ScaleWidth
    txtComments.Height = Me.ScaleHeight - cmdSave.Height - 100
    cmdSave.Top = txtComments.Top + txtComments.Height + 50
    cmdSave.Left = Me.ScaleWidth - cmdSave.Width

  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmComments, Form_Resize"
End Sub

