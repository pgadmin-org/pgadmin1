VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "VSAdoSelector.ocx"
Begin VB.Form frmTunedb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tune db"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmTunedb.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1320
   ScaleWidth      =   4305
   Begin vsAdoSelector.VS_AdoSelector vssLocal 
      Height          =   315
      Left            =   3150
      TabIndex        =   0
      Top             =   90
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectorType    =   1
      DisplayList     =   "No;Yes;"
      IndexList       =   "0;1;"
   End
   Begin VB.TextBox txtCache 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3150
      TabIndex        =   2
      Top             =   945
      Width           =   1065
   End
   Begin VB.TextBox txtDelay 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3150
      TabIndex        =   1
      Top             =   540
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Background retrieval batch size (records):"
      Height          =   225
      Index           =   2
      Left            =   105
      TabIndex        =   5
      Top             =   1010
      Width           =   3270
   End
   Begin VB.Label Label1 
      Caption         =   "Background retrieval delay (seconds):"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   570
      Width           =   3270
   End
   Begin VB.Label Label1 
      Caption         =   "Allow local storage of UID and PWD:"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   150
      Width           =   2745
   End
End
Attribute VB_Name = "frmTunedb"
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

'Everything to do with msysconf is done unquoted to prevent
'compatibility errors with older versions

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 4395 Then Me.Width = 4395
    If Me.Height < 1695 Then Me.Height = 1695
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTunedb, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rs As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 4395
  Me.Height = 1695
  If ObjectExists("msysconf", tTable) = 0 Then
    If MsgBox("The MSysConf table does not exist. Do you want it to be created?", vbExclamation + vbOKCancel) = vbOK Then
      CreateMSysConf
    Else
       ActionCancelled = True
       Unload Me
       Exit Sub
    End If
  End If
  vssLocal.LoadList
  StartMsg "Retrieving db Config data..."
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: SELECT * FROM msysconf WHERE Config = 101"
  rs.Open "SELECT nvalue FROM msysconf WHERE Config = 101", gConnection, adOpenForwardOnly
  If Not rs.EOF Then vssLocal.SelectItem rs!nvalue
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: SELECT * FROM msysconf WHERE Config = 102"
  rs.Open "SELECT nvalue FROM msysconf WHERE Config = 102", gConnection, adOpenForwardOnly
  If Not rs.EOF Then txtDelay.Text = rs!nvalue
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: SELECT * FROM msysconf WHERE Config = 103"
  rs.Open "SELECT nvalue FROM msysconf WHERE Config = 103", gConnection, adOpenForwardOnly
  If Not rs.EOF Then txtCache.Text = rs!nvalue
  EndMsg
  Set rs = Nothing
  Exit Sub
Err_Handler:
  Set rs = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTunedb, Form_Load"
End Sub

Private Sub txtCache_Change()
On Error GoTo Err_Handler
  If txtCache.Text = "" Then Exit Sub
  If Validate(txtCache.Text, vdtInteger, True) = False Then Exit Sub
  gConnection.Execute "UPDATE msysconf SET nvalue = '" & txtCache.Text & "' WHERE config = 103"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTunedb, txtCache_Change"
End Sub

Private Sub txtDelay_Change()
On Error GoTo Err_Handler
  If txtDelay.Text = "" Then Exit Sub
  If Validate(txtDelay.Text, vdtInteger, True) = False Then Exit Sub
  gConnection.Execute "UPDATE msysconf SET nvalue = '" & txtDelay.Text & "' WHERE config = 102"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTunedb, txtDelay_Change"
End Sub


Private Sub vssLocal_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
  gConnection.Execute "UPDATE msysconf SET nvalue = '" & Item & "' WHERE config = 101"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTunedb, vssLocal_ItemSelected"
End Sub
