VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmIndexes 
   Caption         =   "Indexes"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmIndexes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected object."
      Top             =   765
      Width           =   1365
   End
   Begin VB.CommandButton cmdAddIndex 
      Caption         =   "&Create Index"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new Index"
      Top             =   45
      Width           =   1380
   End
   Begin VB.CommandButton cmdDropIndex 
      Caption         =   "&Drop Index"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Drop the selected Index from the database"
      Top             =   405
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   21
      Top             =   1485
      Width           =   1380
      Begin VB.CheckBox chkIndexes 
         Caption         =   "Indexes"
         Height          =   225
         Left            =   105
         TabIndex        =   4
         ToolTipText     =   "Select to show system Indexes"
         Top             =   225
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Reload the Index definitions from the database"
      Top             =   1125
      Width           =   1380
   End
   Begin MSComctlLib.TreeView trvBrowser 
      Height          =   4005
      Left            =   1485
      TabIndex        =   5
      ToolTipText     =   "Browse Indexes and Indexed Columns"
      Top             =   0
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7064
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ilBrowser"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilBrowser 
      Left            =   810
      Top             =   2295
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndexes.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndexes.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndexes.frx":093E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraIndex 
      Caption         =   "Index Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   25
      Top             =   0
      Width           =   3660
      Begin HighlightBox.HBX txtComments 
         Height          =   2130
         Left            =   90
         TabIndex        =   11
         Top             =   1800
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   3757
         BackColor       =   -2147483633
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
         Caption         =   "Comments"
      End
      Begin VB.TextBox txtLossy 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1485
         Width           =   2625
      End
      Begin VB.TextBox txtTable 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   2625
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   2625
      End
      Begin VB.TextBox txtUnique 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   855
         Width           =   2625
      End
      Begin VB.TextBox txtPrimary 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1170
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lossy?"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   38
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Primary?"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   29
         Top             =   1215
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unique?"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   28
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   27
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   26
         Top             =   315
         Width           =   285
      End
   End
   Begin VB.Frame fraColumn 
      Caption         =   "Column Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   30
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtType 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1170
         Width           =   2625
      End
      Begin VB.TextBox txtColOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   2625
      End
      Begin VB.TextBox txtNumber 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   540
         Width           =   2625
      End
      Begin VB.TextBox txtLength 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   855
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   21
         Left            =   90
         TabIndex        =   34
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number"
         Height          =   195
         Index           =   20
         Left            =   90
         TabIndex        =   33
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   240
         Index           =   19
         Left            =   90
         TabIndex        =   32
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   31
         Top             =   1215
         Width           =   360
      End
   End
   Begin VB.Frame fraDatasource 
      Caption         =   "Datasource Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   22
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtdbVer 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   855
         Width           =   2805
      End
      Begin VB.TextBox txtPlatform 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1170
         Width           =   2805
      End
      Begin VB.TextBox txtCompiler 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1485
         Width           =   2805
      End
      Begin VB.TextBox txtTimeOut 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   2490
      End
      Begin VB.TextBox txtUsername 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DBMS"
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   37
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   36
         Top             =   1215
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Compiler"
         Height          =   195
         Index           =   23
         Left            =   90
         TabIndex        =   35
         Top             =   1530
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Timeout"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   585
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   270
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmIndexes"
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
Dim rsIndexes As New Recordset
Dim rsFields As New Recordset

Private Sub trvBrowser_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXIndexes
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmIndexes, trvBrowser_MouseUp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsIndexes = Nothing
  Set rsFields = Nothing
End Sub

Private Sub chkIndexes_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmIndexes, chkIndexes_Click"
End Sub

Public Sub cmdAddIndex_Click()
On Error GoTo Err_Handler
  Load frmAddIndex
  frmAddIndex.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmIndexes, cmdAddIndex_Click"
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If txtOID.Text = "" Or fraColumn.Visible = True Then
    MsgBox "You must select an index to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  Load frmComments
  frmComments.Setup "frmIndexes", QUOTE & trvBrowser.SelectedItem.Text & QUOTE, Val(txtOID.Text)
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmIndexes, cmdComment_Click"
End Sub

Public Sub cmdDropIndex_Click()
On Error GoTo Err_Handler
  If Left(trvBrowser.SelectedItem.Key, 1) <> "I" Then
    MsgBox "That object is not an index!", vbExclamation, "Error"
    Exit Sub
  Else
    If MsgBox("Are you sure you wish to delete " & trvBrowser.SelectedItem.Text & "?", vbYesNo + vbQuestion, _
              "Confirm Index Delete") = vbYes Then
      StartMsg "Dropping Index..."
      fMainForm.txtSQLPane.Text = " DROP INDEX " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      LogMsg "Executing: DROP INDEX " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      gConnection.Execute "DROP INDEX " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      LogQuery "DROP INDEX " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      trvBrowser.Nodes.Remove trvBrowser.SelectedItem.Key
      EndMsg
    End If
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmIndexes, cmdDropIndex_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
Dim NodeX As Node
Dim rsDesc As New Recordset

  fraIndex.Visible = False
  fraColumn.Visible = False
  fraDatasource.Visible = False
  Me.Refresh
  txtUsername.Text = Username
  txtTimeOut.Text = gConnection.CommandTimeout
  
  StartMsg "Retrieving Index Definitions..."
  If rsIndexes.State <> adStateClosed Then rsIndexes.Close
  If rsFields.State <> adStateClosed Then rsFields.Close
  If chkIndexes.Value = 1 Then
    LogMsg "Executing: SELECT DISTINCT ON(index_name) index_oid, index_name, index_table, index_is_unique, index_is_primary, index_is_lossy, index_comments FROM pgadmin_indexes ORDER BY index_name"
    rsIndexes.Open "SELECT DISTINCT ON(index_name) index_oid, index_name, index_table, index_is_unique, index_is_primary, index_is_lossy, index_comments FROM pgadmin_indexes ORDER BY index_name", gConnection, adOpenDynamic
  Else
    LogMsg "Executing: SELECT DISTINCT ON(index_name) index_oid, index_name, index_table, index_is_unique, index_is_primary, index_is_lossy, index_comments FROM pgadmin_indexes WHERE index_oid > " & LAST_SYSTEM_OID & " AND index_name NOT LIKE 'pgadmin_%' AND index_name NOT LIKE 'pg_%' ORDER BY index_name"
    rsIndexes.Open "SELECT DISTINCT ON(index_name) index_oid, index_name, index_table, index_is_unique, index_is_primary, index_is_lossy, index_comments FROM pgadmin_indexes WHERE index_oid > " & LAST_SYSTEM_OID & " AND index_name NOT LIKE 'pgadmin_%' AND index_name NOT LIKE 'pg_%' ORDER BY index_name", gConnection, adOpenDynamic
  End If
  trvBrowser.Nodes.Clear
  Set NodeX = trvBrowser.Nodes.Add(, tvwChild, "D:" & Datasource, Datasource, 1)
  While Not rsIndexes.EOF
    Set NodeX = trvBrowser.Nodes.Add("D:" & Datasource, tvwChild, "I:" & rsIndexes!index_oid, rsIndexes!index_name, 2)
    rsIndexes.MoveNext
  Wend
  If rsIndexes.BOF <> True Then rsIndexes.MoveFirst
  
  trvBrowser.Nodes(1).Expanded = True
  LogMsg "Executing: SELECT version()"
  rsDesc.Open "SELECT version()", gConnection, adOpenForwardOnly
  txtdbVer.Text = Mid(rsDesc!Version, 1, InStr(1, rsDesc!Version, " on ") - 1)
  txtPlatform.Text = Mid(rsDesc!Version, InStr(1, rsDesc!Version, " on") + 4, InStr(1, rsDesc!Version, ", compiled by ") - InStr(1, rsDesc!Version, " on") - 4)
  txtCompiler.Text = Mid(rsDesc!Version, InStr(1, rsDesc!Version, ", compiled by ") + 14, Len(rsDesc!Version))
  fraDatasource.Visible = True
  EndMsg
  Set rsDesc = Nothing
  Exit Sub
Err_Handler:
  Set rsDesc = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmIndexes, cmdRefresh_Click"
  Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = 8325
  Me.Height = 4455
  LogMsg "Loading Form: " & Me.Name
  cmdRefresh_Click
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmIndexes, Form_Load"
  Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtComments.Minimise
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If frmIndexes.Width < 8325 Then frmIndexes.Width = 8325
      If frmIndexes.Height < 4455 Then frmIndexes.Height = 4455
    End If
    
    trvBrowser.Height = frmIndexes.ScaleHeight
    trvBrowser.Width = frmIndexes.ScaleWidth - trvBrowser.Left - fraDatasource.Width - 25
    fraDatasource.Left = trvBrowser.Left + trvBrowser.Width + 25
    fraDatasource.Height = Me.ScaleHeight
    txtComments.Height = fraDatasource.Height - txtComments.Top - 100
    fraIndex.Left = fraDatasource.Left
    fraIndex.Height = fraDatasource.Height
    fraColumn.Left = fraDatasource.Left
    fraColumn.Height = fraDatasource.Height

  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmIndexes, Form_Resize"
End Sub

Private Sub trvBrowser_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
Dim NodeX As Node
Dim lOID As Long
Dim x As Integer
Dim rsTemp As New Recordset

  'If a index was clicked, set the data in the grid, and create children
  'if necessary

  Select Case Mid(Node.Key, 1, 1)
  Case "I"
    StartMsg "Retrieving Index Info..."
    fraDatasource.Visible = False
    fraColumn.Visible = False
    While Not rsIndexes.EOF
      If rsIndexes!index_name = Node.Text Then
        lOID = rsIndexes!index_oid
        txtOID.Text = rsIndexes!index_oid & ""
        txtTable.Text = rsIndexes!index_table & ""
        txtUnique.Text = rsIndexes!index_is_unique & ""
        txtPrimary.Text = rsIndexes!index_is_primary & ""
        txtLossy.Text = rsIndexes!index_is_lossy & ""
        txtComments.Text = rsIndexes!index_comments & ""
        rsIndexes.MoveLast
      End If
      rsIndexes.MoveNext
    Wend
    If rsIndexes.BOF <> True Then rsIndexes.MoveFirst
    
    'Get Columns
    If rsFields.State = adStateClosed Then
      LogMsg "Executing: SELECT index_oid, index_name, column_oid, column_name, column_position, column_type, column_length, column_comments FROM pgadmin_indexes ORDER BY column_position"
      rsFields.Open "SELECT index_oid, index_name, column_oid, column_name, column_position, column_type, column_length, column_comments FROM pgadmin_indexes ORDER BY column_position", gConnection, adOpenStatic
      On Error Resume Next
      While Not rsFields.EOF
        Set NodeX = trvBrowser.Nodes.Add("I:" & rsFields!index_oid, tvwChild, "F:" & rsFields!column_oid & ":" & rsFields!column_name, rsFields!column_name, 3)
      rsFields.MoveNext
      Wend
      On Error GoTo Err_Handler
      If rsFields.BOF <> True Then rsFields.MoveFirst
    End If
    EndMsg
    fraIndex.Visible = True

  Case "F"

    StartMsg "Retrieving Index Column Definitions..."
    fraDatasource.Visible = False
    fraIndex.Visible = False
    While Not rsFields.EOF
      If rsFields!column_name = Node.Text And rsFields!index_name = Node.Parent.Text Then
        txtColOID.Text = rsFields!column_oid & ""
        txtNumber.Text = rsFields!column_position & ""
        If rsFields!column_type & "" = "numeric" Then
          x = Hex((rsFields!column_length - 4) And &HFFFF)
          txtLength.Text = CLng("&H" & Mid(x, 1, Len(x) - 4)) & "," & CLng("&H" & Mid(x, Len(x) - 3, Len(x)))
        Else
          txtLength.Text = rsFields!column_length & ""
        End If
        txtType.Text = rsFields!column_type & ""
      End If
      rsFields.MoveNext
    Wend
    If rsFields.BOF <> True Then rsFields.MoveFirst
    fraColumn.Visible = True
    
  Case "D"

    fraIndex.Visible = False
    fraColumn.Visible = False
    txtUsername.Text = Username
    txtTimeOut.Text = gConnection.ConnectionTimeout
    fraDatasource.Visible = True
    LogMsg "Executing: SELECT version()"
    rsTemp.Open "SELECT version()", gConnection, adOpenForwardOnly
    txtdbVer.Text = Mid(rsTemp!Version, 1, InStr(1, rsTemp!Version, " on "))
    txtPlatform.Text = Mid(rsTemp!Version, InStr(1, rsTemp!Version, " on") + 4, InStr(1, rsTemp!Version, ", compiled by ") - InStr(1, rsTemp!Version, " on") - 4)
    txtCompiler.Text = Mid(rsTemp!Version, InStr(1, rsTemp!Version, ", compiled by ") + 14, Len(rsTemp!Version))
    
  End Select

  'This stuff can always be done.
  
  Set rsTemp = Nothing
  Node.Expanded = True
  EndMsg
  Exit Sub
Err_Handler:
  Set rsTemp = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmIndexes, trvBrowser_NodeClick"
End Sub

