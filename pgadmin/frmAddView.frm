VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#5.0#0"; "HighlightBox.ocx"
Begin VB.Form frmAddView 
   Caption         =   "Create View"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmAddView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.Frame fraDetails 
      Caption         =   "View Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   3
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtComments 
         Height          =   2220
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1755
         Width           =   3480
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   855
         Width           =   2670
      End
      Begin VB.TextBox txtACL 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1170
         Width           =   2670
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   900
         TabIndex        =   7
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   2670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   13
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACL"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   1215
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   5
         Top             =   585
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   3510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Query"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   450
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Save View"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   1410
   End
   Begin HighlightBox.HBX txtSQL 
      Height          =   3960
      Left            =   1485
      TabIndex        =   0
      Top             =   45
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   6985
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      ScrollBars      =   2
      MultiLine       =   -1  'True
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
Dim lngView_oid_old As Long
Dim szView_name_old As String

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
  bContinueRebuilding = True
  
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the View!", vbExclamation, "Error"
    Exit Sub
  End If
  If txtSQL.Text = "" Or txtSQL.Text = "Not a view" Then
    MsgBox "You must enter an SQL query for the View!", vbExclamation, "Error"
    Exit Sub
  End If
  StartMsg "Creating View..."
    

    If szView_name_old <> "" Then cmp_View_DropIfExists "pgadmin_dev_views", 0, szView_name_old
    
    ' Create view
    cmp_View_Create "pgadmin_dev_views", txtName.Text, txtSQL.Text, txtOwner.Text, txtACL.Text, txtComments.Text
    
    If bContinueRebuilding = True Then
        frmViews.cmdRefresh_Click
        Unload Me
    End If
    
    EndMsg
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
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4455 Then Me.Height = 4455
    End If
    txtSQL.Height = Me.ScaleHeight
    txtSQL.Width = Me.ScaleWidth - txtSQL.Left - fraDetails.Width - 25
    fraDetails.Left = txtSQL.Left + txtSQL.Width + 25
    fraDetails.Height = Me.ScaleHeight
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddViews, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
    Dim lngView_oid As Long
    Dim szView_name As String
    Dim szView_definition As String
    Dim szView_owner As String
    Dim szView_acl As String
    Dim szView_comments As String
    
    szView_name_old = gView_Name
    szView_name = gView_Name
    gView_Name = ""
    
    LogMsg "Loading Form: " & Me.Name
    Me.Height = 3675
    Me.Width = 4770
    txtSQL.Wordlist = TextColours
    
    ' Retrieve view if exists

    If szView_name_old <> "" Then
      Me.Caption = "Modify view"
      
      ' Load View data
      lngView_oid = 0
      cmp_View_GetValues "pgadmin_dev_views", lngView_oid, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
      
      txtName.Text = szView_name
      txtSQL.Text = szView_definition
      txtComments.Text = szView_comments
      
      If (lngView_oid = 0) Then
            txtOID.Text = "N.S."
            txtACL.Text = "N.S."
            txtOwner.Text = "N.S."
      Else
            txtOID.Text = lngView_oid
            txtACL.Text = szView_acl
            txtOwner.Text = szView_owner
      End If
    Else
      Me.Caption = "Create view"
      txtOID.Text = "N.S."
      txtOwner.Text = "N.S."
      txtACL.Text = "N.S."
    End If
    
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

