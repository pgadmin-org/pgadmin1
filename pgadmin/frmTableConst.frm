VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.1#0"; "HighlightBox.ocx"
Begin VB.Form frmTableConst 
   Caption         =   "Table Constraints"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmTableConst.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   4875
   Begin vsAdoSelector.VS_AdoSelector vssType 
      Height          =   315
      Left            =   1485
      TabIndex        =   8
      ToolTipText     =   "Select the type of Constraint to add."
      Top             =   45
      Width           =   2760
      _ExtentX        =   4868
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
      DisplayList     =   "Primary Key;Unique;Check;Foreign Key;"
      IndexList       =   "Primary Key;Unique;Check;Foreign Key;"
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1485
      TabIndex        =   2
      ToolTipText     =   "Enter a name for the Constraint."
      Top             =   405
      Width           =   2760
   End
   Begin VB.Frame fraCheck 
      Caption         =   "Check Constraint"
      Height          =   3120
      Left            =   45
      TabIndex        =   5
      Top             =   765
      Visible         =   0   'False
      Width           =   4785
      Begin HighlightBox.HBX txtCheck 
         Height          =   2400
         Left            =   90
         TabIndex        =   24
         ToolTipText     =   "Enter an expression that evaluates to TRUE or FALSE"
         Top             =   225
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   4233
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Expression"
      End
      Begin VB.CommandButton cmdAddCheck 
         Caption         =   "&Add Check"
         Height          =   330
         Left            =   3240
         TabIndex        =   7
         ToolTipText     =   "Add the Check Constraint"
         Top             =   2700
         Width           =   1455
      End
   End
   Begin VB.Frame fraPrimaryKey 
      Caption         =   "Primary Key Constraint"
      Height          =   3120
      Left            =   45
      TabIndex        =   4
      Top             =   765
      Visible         =   0   'False
      Width           =   4785
      Begin VB.CommandButton cmdAddPrimaryKey 
         Caption         =   "&Add Primary Key"
         Height          =   330
         Left            =   3240
         TabIndex        =   10
         ToolTipText     =   "add the Primary Key Constraint"
         Top             =   2700
         Width           =   1455
      End
      Begin VB.ListBox lstPrimaryKeyCols 
         Height          =   2085
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   9
         ToolTipText     =   "Select the columns to include in the Primary Key Constraint"
         Top             =   540
         Width           =   4605
      End
      Begin VB.Label Label4 
         Caption         =   "Select the Primary Key Columns"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   2895
      End
   End
   Begin VB.Frame fraUnique 
      Caption         =   "Unique Constraint"
      Height          =   3120
      Left            =   45
      TabIndex        =   3
      Top             =   765
      Visible         =   0   'False
      Width           =   4785
      Begin VB.ListBox lstUniqueCols 
         Height          =   2085
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   14
         ToolTipText     =   "Select the columns to be included in the Unique Constraint."
         Top             =   540
         Width           =   4605
      End
      Begin VB.CommandButton cmdAddUnique 
         Caption         =   "&Add Unique"
         Height          =   330
         Left            =   3240
         TabIndex        =   11
         ToolTipText     =   "Add the Unique Constraint"
         Top             =   2700
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Select the Unique Columns"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   270
         Width           =   2895
      End
   End
   Begin VB.Frame fraForeignKey 
      Caption         =   "Foreign Key Constraint"
      Height          =   3120
      Left            =   45
      TabIndex        =   6
      Top             =   765
      Visible         =   0   'False
      Width           =   4785
      Begin TabDlg.SSTab SSTab1 
         Height          =   2400
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   4233
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Columns"
         TabPicture(0)   =   "frmTableConst.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstForeignKeyCols"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Referenced Objects"
         TabPicture(1)   =   "frmTableConst.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label7"
         Tab(1).Control(1)=   "Label8"
         Tab(1).Control(2)=   "Label9"
         Tab(1).Control(3)=   "vssTables"
         Tab(1).Control(4)=   "lstReferencedColumns"
         Tab(1).ControlCount=   5
         Begin VB.ListBox lstReferencedColumns 
            Height          =   1185
            ItemData        =   "frmTableConst.frx":0342
            Left            =   -74100
            List            =   "frmTableConst.frx":0344
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   1050
            Width           =   3570
         End
         Begin vsAdoSelector.VS_AdoSelector vssTables 
            Height          =   315
            Left            =   -74100
            TabIndex        =   21
            Top             =   645
            Width           =   3570
            _ExtentX        =   6297
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
         End
         Begin VB.ListBox lstForeignKeyCols 
            Height          =   1635
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   17
            ToolTipText     =   "Select the columns to include in the Foreign Key Constraint"
            Top             =   645
            Width           =   4425
         End
         Begin VB.Label Label9 
            Caption         =   "Columns"
            Height          =   195
            Left            =   -74865
            TabIndex        =   22
            Top             =   1095
            Width           =   690
         End
         Begin VB.Label Label8 
            Caption         =   "Table"
            Height          =   195
            Left            =   -74865
            TabIndex        =   20
            Top             =   735
            Width           =   690
         End
         Begin VB.Label Label7 
            Caption         =   "Select the Referenced table and Columns:"
            Height          =   285
            Left            =   -74955
            TabIndex        =   19
            Top             =   420
            Width           =   3660
         End
         Begin VB.Label Label6 
            Caption         =   "Select the Foreign Key Columns:"
            Height          =   240
            Left            =   45
            TabIndex        =   18
            Top             =   420
            Width           =   4425
         End
      End
      Begin VB.CommandButton cmdAddForeignKey 
         Caption         =   "&Add Foreign Key"
         Height          =   330
         Left            =   3240
         TabIndex        =   12
         ToolTipText     =   "Add the Foreign Key Constraint"
         Top             =   2700
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Name (optional)"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Constraint Type"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
End
Attribute VB_Name = "frmTableConst"
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

Private Sub AddConstraint(szConstraint As String)
On Error GoTo Err_Handler
Dim x As Integer
  For x = 0 To Forms.Count - 1
    If Forms(x).Name = "frmAddTable" Then Exit For
  Next
  If x = Forms.Count Then
    MsgBox "The create table dialogue appears to have been closed!", vbCritical, "Fatal Error"
    Unload Me
    Exit Sub
  End If
  Forms(x).lstConstraints.AddItem szConstraint
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTableConst, AddConstraint"
End Sub

Private Sub cmdAddCheck_Click()
On Error GoTo Err_Handler
  If txtName.Text = "" Then
    AddConstraint "CHECK (" & txtCheck.Text & ")"
  Else
    AddConstraint "CONSTRAINT " & QUOTE & txtName.Text & QUOTE & " CHECK (" & txtCheck.Text & ")"
  End If
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTableConst, cmdAddCheck_Click"
End Sub

Private Sub cmdAddForeignKey_Click()
On Error GoTo Err_Handler
Dim szConst As String
Dim x As Integer
Dim Flag As Boolean
  Flag = False
  If txtName.Text = "" Then
    szConst = "FOREIGN KEY ("
  Else
    szConst = "CONSTRAINT " & QUOTE & txtName.Text & QUOTE & " FOREIGN KEY ("
  End If
  For x = 0 To lstForeignKeyCols.ListCount - 1
    If lstForeignKeyCols.Selected(x) = True Then
      szConst = szConst & QUOTE & lstForeignKeyCols.List(x) & QUOTE & ", "
      Flag = True
    End If
  Next
  If Flag = False Then
    MsgBox "You must select at least one column!", vbExclamation, "Error"
    Exit Sub
  End If
  szConst = Mid(szConst, 1, Len(szConst) - 2) & ") REFERENCES " & QUOTE & vssTables.Caption & QUOTE & " ("
  Flag = False
  For x = 0 To lstReferencedColumns.ListCount - 1
    If lstReferencedColumns.Selected(x) = True Then
      szConst = szConst & QUOTE & lstReferencedColumns.List(x) & QUOTE & ", "
      Flag = True
    End If
  Next
  If Flag = False Then
    MsgBox "You must select at least one referenced column!", vbExclamation, "Error"
    Exit Sub
  End If
  szConst = Mid(szConst, 1, Len(szConst) - 2) & ")"
  AddConstraint szConst
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTableConst, cmdAddPrimaryKey_Click"
End Sub

Private Sub cmdAddPrimaryKey_Click()
On Error GoTo Err_Handler
Dim szConst As String
Dim x As Integer
Dim Flag As Boolean
  Flag = False
  If txtName.Text = "" Then
    szConst = "PRIMARY KEY ("
  Else
    szConst = "CONSTRAINT " & QUOTE & txtName.Text & QUOTE & " PRIMARY KEY ("
  End If
  For x = 0 To lstPrimaryKeyCols.ListCount - 1
    If lstPrimaryKeyCols.Selected(x) = True Then
      szConst = szConst & QUOTE & lstPrimaryKeyCols.List(x) & QUOTE & ", "
      Flag = True
    End If
  Next
  If Flag = False Then
    MsgBox "You must select at least one column!", vbExclamation, "Error"
    Exit Sub
  End If
  szConst = Mid(szConst, 1, Len(szConst) - 2) & ")"
  AddConstraint szConst
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTableConst, cmdAddPrimaryKey_Click"
End Sub

Private Sub cmdAddUnique_Click()
On Error GoTo Err_Handler
Dim szConst As String
Dim x As Integer
Dim Flag As Boolean
  Flag = False
  If txtName.Text = "" Then
    szConst = "UNIQUE ("
  Else
    szConst = "CONSTRAINT " & QUOTE & txtName.Text & QUOTE & " UNIQUE ("
  End If
  For x = 0 To lstUniqueCols.ListCount - 1
    If lstUniqueCols.Selected(x) = True Then
      szConst = szConst & QUOTE & lstUniqueCols.List(x) & QUOTE & ", "
      Flag = True
    End If
  Next
  If Flag = False Then
    MsgBox "You must select at least one column!", vbExclamation, "Error"
    Exit Sub
  End If
  szConst = Mid(szConst, 1, Len(szConst) - 2) & ")"
  AddConstraint szConst
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTableConst, cmdAddUnique_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim x As Integer
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 4335
  Me.Width = 4995
  txtCheck.Wordlist = TextColours
  StartMsg "Retrieving Table Names..."
  LogMsg "Executing: SELECT DISTINCT ON(table_name) table_oid, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
  vssTables.SQL = "SELECT DISTINCT ON(table_name) table_oid, table_name FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
  vssTables.Connect = Connect
  vssTables.LoadList
  vssType.LoadList
  For x = 0 To frmAddTable.lstColumns.ListCount - 1
    lstPrimaryKeyCols.AddItem Mid(frmAddTable.lstColumns.List(x), 2, InStr(2, frmAddTable.lstColumns.List(x), QUOTE) - 2)
    lstForeignKeyCols.AddItem Mid(frmAddTable.lstColumns.List(x), 2, InStr(2, frmAddTable.lstColumns.List(x), QUOTE) - 2)
    lstUniqueCols.AddItem Mid(frmAddTable.lstColumns.List(x), 2, InStr(2, frmAddTable.lstColumns.List(x), QUOTE) - 2)
  Next
  vssType.SelectItem ("Primary Key")
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTableConst, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtCheck.Minimise
  If Me.WindowState = 0 Then
    Me.Height = 4335
    Me.Width = 4995
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTableConst, Form_Resize"
End Sub

Private Sub vssTables_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
Dim rsColumns As New Recordset
  lstReferencedColumns.Clear
  StartMsg "Retrieving column names..."
  LogMsg "Executing: SELECT column_name FROM pgadmin_tables WHERE column_position > 0 AND table_name = '" & ItemText & "' ORDER BY column_position"
  rsColumns.Open "SELECT column_name FROM pgadmin_tables WHERE column_position > 0 AND table_name = '" & ItemText & "' ORDER BY column_position", gConnection, adOpenForwardOnly
  While Not rsColumns.EOF
    lstReferencedColumns.AddItem rsColumns!column_name
    rsColumns.MoveNext
  Wend
  If rsColumns.State <> adStateClosed Then rsColumns.Close
  EndMsg
  Set rsColumns = Nothing
  Exit Sub
Err_Handler:
  Set rsColumns = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTableConst, vssTables_ItemSelected"
End Sub

Private Sub vssType_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
Dim rs As New Recordset
  Select Case Item
    Case "Primary Key"
      fraPrimaryKey.Visible = True
      fraUnique.Visible = False
      fraCheck.Visible = False
      fraForeignKey.Visible = False
    Case "Unique"
      fraPrimaryKey.Visible = False
      fraUnique.Visible = True
      fraCheck.Visible = False
      fraForeignKey.Visible = False
    Case "Check"
      fraPrimaryKey.Visible = False
      fraUnique.Visible = False
      fraCheck.Visible = True
      fraForeignKey.Visible = False
    Case "Foreign Key"
      LogMsg "Executing: SELECT version()"
      rs.Open "SELECT version()", gConnection, adOpenForwardOnly
      If Val(Mid(rs!Version, 11, 14)) < 7 Then
        MsgBox "Foreign Key constraints are only available with PostgreSQL 7.0 or higher!", vbExclamation, "Error"
        vssType.SelectItem "Primary Key"
        Exit Sub
      End If
      fraPrimaryKey.Visible = False
      fraUnique.Visible = False
      fraCheck.Visible = False
      fraForeignKey.Visible = True
  End Select
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTableConst, vssType_ItemSelected"
End Sub
