VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmViews 
   Caption         =   "Views"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmViews.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdRebuild 
      BackColor       =   &H80000018&
      Caption         =   "Rebuild &project"
      Height          =   330
      Left            =   45
      TabIndex        =   23
      ToolTipText     =   "Checks and rebuilds dependencies on functions, triggers and views."
      Top             =   3555
      Width           =   1410
   End
   Begin VB.CommandButton cmdExportView 
      Caption         =   "Export View"
      Enabled         =   0   'False
      Height          =   330
      Left            =   45
      TabIndex        =   22
      ToolTipText     =   "Modify the selected View."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdModifyView 
      Caption         =   "&Modify View"
      Height          =   330
      Left            =   45
      TabIndex        =   21
      ToolTipText     =   "Modify the selected View."
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton cmdViewData 
      Caption         =   "&View Data"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Edit the comment for the selected View."
      Top             =   2205
      Width           =   1410
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected View."
      Top             =   1485
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "View Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   12
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txtACL 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1170
         Width           =   2670
      End
      Begin VB.TextBox txtDefinition 
         BackColor       =   &H8000000F&
         Height          =   1230
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1755
         Width           =   3480
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   2670
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   855
         Width           =   2670
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   645
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3285
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Top             =   585
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACL"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   1215
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Definition"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   14
         Top             =   1530
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   13
         Top             =   3060
         Width           =   735
      End
   End
   Begin VB.ListBox lstView 
      Height          =   3960
      Left            =   1485
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   45
      Width           =   2985
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Refresh the list of Views."
      Top             =   1845
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropView 
      Caption         =   "&Drop View"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected View."
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateView 
      Caption         =   "&Create View"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new View."
      Top             =   45
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   17
      Top             =   2970
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Views"
         Height          =   225
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   "Select to view system views"
         Top             =   225
         Width           =   1065
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   945
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select SQL File"
      Filter          =   "All Files (*.*)|*.*"
   End
End
Attribute VB_Name = "frmViews"
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
Dim rsView As New Recordset
Dim szView_PostgreSqlTable As String

Private Sub cmdExportView_Click()
On Error GoTo Err_Handler

    Dim iLoop As Long
    Dim iListCount As Long
    Dim szExport As String
    Dim bExport As Boolean
    Dim szHeader As String
    
    Dim lngView_oid As Long
    Dim szView_name As String
    Dim szView_owner As String
    Dim szView_acl As String
    Dim szView_comments As String
    Dim szView_definition As String
    
    bExport = False
    szExport = ""

    iListCount = lstView.ListCount
        
    For iLoop = 0 To iListCount - 1
        If lstView.Selected(iLoop) = True Then
            bExport = True
            szView_name = lstView.List(iLoop)
            cmp_View_GetValues szView_PostgreSqlTable, 0, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
            
            ' Header
            szExport = szExport & "/*" & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & szView_name & vbCrLf
            If szView_comments <> "" Then szExport = szExport & szView_comments & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & "*/" & vbCrLf
            
            ' Function
            szExport = szExport & cmp_View_CreateSQL(szView_name, szView_definition) & vbCrLf & vbCrLf
        End If
    Next iLoop
    
    If bExport Then
        szHeader = "/*" & vbCrLf
        szHeader = szHeader & Format(Now, "d mmmm yyyy") & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "The choice of a GNU generation, " & Format(Now, "d mmmm yyyy") & vbCrLf
        szHeader = szHeader & "PostgreSQL     www.postgresql.org" & vbCrLf
        szHeader = szHeader & "PgAdmin        www.greatbridge.org/project/pgadmin" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "*/" & vbCrLf & vbCrLf
        szExport = szHeader & szExport
        MsgExportToFile CommonDialog1, szExport, "sql", "Export views"
    End If
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdExportView_Click"
End Sub

Public Sub cmdModifyView_Click()
 On Error GoTo Err_Handler

If lstView.Text <> "" Then
    gView_Name = lstView.Text
    Load frmAddView
    frmAddView.Show
End If

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdModifyView_Click"
End Sub

Private Sub cmdRebuild_Click()
    cmp_Project_Rebuild
End Sub

Private Sub lstView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXViews
  
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, lstViews_MouseUp"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, ChkSystem_Click"
End Sub

Public Sub cmdViewData_Click()
On Error GoTo Err_Handler
Dim Response As Integer
Dim Tuples As Long
Dim rsQuery As New Recordset
  If lstView.Text = "" Then
    MsgBox "You must select a view to view!", vbExclamation, "Error"
    Exit Sub
  End If
  If rsQuery.State <> adStateClosed Then rsQuery.Close
  LogMsg "Executing: SELECT count(*) As records FROM " & QUOTE & lstView.Text & QUOTE
  rsQuery.Open "SELECT count(*) As records FROM " & QUOTE & lstView.Text & QUOTE, gConnection, adOpenForwardOnly
  If Not rsQuery.EOF Then
    Tuples = rsQuery!Records
  Else
    Tuples = 0
  End If
  If rsQuery.State <> adStateClosed Then rsQuery.Close
  If Tuples > 1000 Then
    Response = MsgBox("That table contains " & Tuples & " rows which may take some time to load! Do you wish to continue?", _
    vbExclamation + vbYesNo, "Warning")
    If Response = vbNo Then Exit Sub
  End If
  Dim DataForm As New frmSQLOutput
  LogMsg "Executing: SELECT * FROM " & QUOTE & lstView.Text & QUOTE
  rsQuery.Open "SELECT * FROM " & QUOTE & lstView.Text & QUOTE, gConnection, adOpenForwardOnly, adLockReadOnly
  Load DataForm
  DataForm.Display rsQuery
  DataForm.Show
  DataForm.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdViewData_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsView = Nothing
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If txtOID.Text = "" Then
    MsgBox "You must select a View to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  CallingForm = "frmViews"
  OID = txtOID.Text
  Load frmComments
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdComment_Click"
End Sub

Public Sub cmdCreateView_Click()
On Error GoTo Err_Handler
  gView_Name = ""
  Load frmAddView
  frmAddView.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdCreateView_Click"
End Sub

Public Sub cmdDropView_Click()
On Error GoTo Err_Handler
  Dim szTrigger_name As String
  Dim iListCount As Long
  Dim iLoop As Long
  Dim szDropStr As String
  
  If MsgBox("Are you sure you wish to delete Views?", vbYesNo + vbQuestion, _
            "Confirm View(s) Delete") = vbYes Then
        StartMsg "Dropping View..."
        
        iListCount = lstView.ListCount
        
        For iLoop = 0 To iListCount - 1
           If lstView.Selected(iLoop) = True Then
            szTrigger_name = lstView.List(iLoop)
            cmp_View_DropIfExists szView_PostgreSqlTable, 0, szTrigger_name
            End If
        Next iLoop
            
        cmdRefresh_Click
        EndMsg
  End If
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmViews, cmdDropView_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  Dim iLoop As Long
  Dim iUbound As Long
  Dim szView() As Variant
  Dim szView_name As String
  Dim szQuery As String
  
  StartMsg "Retrieving View Names..."
  lstView.Clear
  txtOID.Text = ""
  txtDefinition.Text = ""
  txtComments.Text = ""
  txtOwner.Text = ""
  If rsView.State <> adStateClosed Then rsView.Close
  If chkSystem.Value = 1 Then
    szView_PostgreSqlTable = "pgadmin_views"
    szQuery = "SELECT view_name FROM " & szView_PostgreSqlTable & " WHERE view_oid < " & LAST_SYSTEM_OID & " OR view_name LIKE 'pgadmin_%'OR view_name LIKE 'pg_%' ORDER BY view_name"
    LogMsg "Executing: " & szQuery
    rsView.Open szQuery, gConnection, adOpenDynamic
  Else
    szView_PostgreSqlTable = "pgadmin_dev_views"
    szQuery = "SELECT view_name FROM " & szView_PostgreSqlTable & " WHERE view_name NOT LIKE 'pgadmin_%' AND view_name NOT LIKE 'pg_%' ORDER BY view_name"
    LogMsg "Executing: " & szQuery
    rsView.Open szQuery, gConnection, adOpenDynamic
  End If
  
  If Not (rsView.EOF) Then
    szView = rsView.GetRows
    iUbound = UBound(szView, 2)
    For iLoop = 0 To iUbound
      szView_name = szView(0, iLoop)
      lstView.AddItem szView_name
    Next iLoop
  End If
  
  Erase szView
  lstView_Click
  
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmViews, cmdRefresh_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8325
  Me.Height = 4455
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4455 Then Me.Height = 4455
    End If
    lstView.Height = Me.ScaleHeight
    lstView.Width = Me.ScaleWidth - lstView.Left - fraDetails.Width - 25
    fraDetails.Left = lstView.Left + lstView.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, Form_Resize"
End Sub

Public Sub lstView_dblClick()
    cmdModifyView_Click
End Sub

Public Sub lstView_Click()
On Error GoTo Err_Handler
    Dim lngView_oid As Long
    Dim szView_name As String
    Dim szView_owner As String
    Dim szView_acl As String
    Dim szView_comments As String
    Dim szView_definition As String
    
    If lstView.SelCount > 0 Then
        szView_name = lstView.Text
    Else
        szView_name = ""
    End If
    
    StartMsg "Retrieving View Info..."
    lngView_oid = 0
    cmp_View_GetValues szView_PostgreSqlTable, lngView_oid, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
    
    txtOID.Text = Trim(Str(lngView_oid))
    If txtOID.Text = 0 Then txtOID.Text = ""

    txtName.Text = szView_name
    txtDefinition.Text = szView_definition
    txtOwner.Text = szView_owner
    txtACL.Text = szView_acl
    txtComments.Text = szView_comments
        
    If txtName.Text <> "" Then
        If txtOID.Text = "" Then txtOID.Text = "N.S."
        If txtOwner.Text = "" Then txtOwner.Text = "N.S."
        If txtACL.Text = "" Then txtACL.Text = "N.S."
        If txtComments.Text = "" Then txtComments.Text = "N.S."
    End If
    
    CmdViewButton
    EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmViews, lstView_Click"
End Sub

Public Sub CmdViewButton()
On Error GoTo Err_Handler

    Dim bSystem As Boolean
    bSystem = (chkSystem.Value = 1)
    cmdButtonActivate bSystem, lstView.SelCount, cmdCreateView, cmdModifyView, cmdDropView, cmdExportView, cmdComment, cmdRefresh, cmdViewData

    If cmp_Project_IsRebuilt = True Then
        cmdRebuild.Enabled = False
    Else
       cmdRebuild.Enabled = True
    End If
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmViews, CmdViewButton"
End Sub
