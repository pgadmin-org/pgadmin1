VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTriggers 
   Caption         =   "Triggers"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdRebuild 
      BackColor       =   &H80000018&
      Caption         =   "Rebuild &project"
      Height          =   330
      Left            =   45
      TabIndex        =   26
      ToolTipText     =   "Checks and rebuilds dependencies on functions, triggers and views."
      Top             =   3555
      Width           =   1410
   End
   Begin VB.CommandButton cmdExportTrig 
      Caption         =   "Export Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   23
      ToolTipText     =   "Modify the selected trigger."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdModifyTrig 
      Caption         =   "&Modify Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   22
      ToolTipText     =   "Modify the selected trigger."
      Top             =   405
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   21
      Top             =   2970
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Triggers"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Select to view system triggers."
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected Trigger."
      Top             =   1485
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Trigger Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   13
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   540
         Width           =   2760
      End
      Begin VB.TextBox txtForEach 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2115
         Width           =   2760
      End
      Begin VB.TextBox txtEvent 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   2760
      End
      Begin VB.TextBox txtExecutes 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1485
         Width           =   2760
      End
      Begin VB.TextBox txtFunction 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1170
         Width           =   2760
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   2760
      End
      Begin VB.TextBox txtTable 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   855
         Width           =   2760
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   1230
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2700
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   25
         Top             =   585
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   19
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   18
         Top             =   1215
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Executes"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   17
         Top             =   1530
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Event"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   1845
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "For Each"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   15
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   14
         Top             =   2430
         Width           =   735
      End
   End
   Begin VB.ListBox lstTrig 
      Height          =   3960
      ItemData        =   "frmTriggers.frx":0000
      Left            =   1485
      List            =   "frmTriggers.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   45
      Width           =   2985
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Refresh the list of Triggers."
      Top             =   1845
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropTrig 
      Caption         =   "&Drop Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected Trigger."
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateTrig 
      Caption         =   "&Create Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new Trigger."
      Top             =   45
      Width           =   1410
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select SQL File"
      Filter          =   "All Files (*.*)|*.*"
   End
End
Attribute VB_Name = "frmTriggers"
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
Dim rsTrig As New Recordset
Dim szTrigger_PostgreSqlTable As String
Private Sub cmdExportTrig_Click()
    Dim iLoop As Long
    Dim iListCount As Long
    Dim szExport As String
    Dim bExport As Boolean
    Dim szHeader As String
    
    Dim szTrigger_oid As Long
    Dim szTrigger_name As String
    Dim szTrigger_table As String
    Dim szTrigger_function As String
    Dim szTrigger_arguments As String
    Dim szTrigger_foreach As String
    Dim szTrigger_event As String
    Dim szTrigger_executes As String
    Dim szTrigger_Comments As String
    
    bExport = False
    szExport = ""

    iListCount = lstTrig.ListCount
        
    For iLoop = 0 To iListCount - 1
        If lstTrig.Selected(iLoop) = True Then
            bExport = True
            cmp_Trigger_ParseName lstTrig.List(iLoop), szTrigger_name, szTrigger_table
            cmp_Trigger_GetValues szTrigger_PostgreSqlTable, 0, szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_foreach, szTrigger_executes, szTrigger_event, szTrigger_Comments
            
            ' Header
            szExport = szExport & "/*" & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & szTrigger_name & " ON " & szTrigger_table & vbCrLf
            If szTrigger_Comments <> "" Then szExport = szExport & szTrigger_Comments & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & "*/" & vbCrLf
            
            ' Function
            szExport = szExport & Replace(cmp_Trigger_CreateSQL(szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_foreach, szTrigger_executes, szTrigger_event), vbCrLf, " ") & vbCrLf & vbCrLf
        End If
    Next iLoop
    
    If bExport Then
        szHeader = "/*" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "The choice of a GNU generation, " & Format(Now, "d mmmm yyyy") & vbCrLf
        szHeader = szHeader & "PostgreSQL     www.postgresql.org" & vbCrLf
        szHeader = szHeader & "PgAdmin        www.greatbridge.org/project/pgadmin" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "*/" & vbCrLf & vbCrLf
        szExport = szHeader & szExport
        MsgExportToFile CommonDialog1, szExport, "sql", "Export triggers"
    End If
End Sub

Public Sub cmdModifyTrig_Click()
 On Error GoTo Err_Handler

If txtOID <> "" Then
    ' This means we can open the function
    cmp_Trigger_ParseName lstTrig.Text, gTrigger_Name, gTrigger_Table
    
    ' Load form
    Load frmAddTrigger
    frmAddTrigger.Show
End If

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdModifyFunc_Click"
End Sub

Private Sub cmdRebuild_Click()
    cmp_Project_Rebuild
End Sub

Private Sub lstTrig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXTriggers
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, lstTrig_MouseUp"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, chkSystem_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsTrig = Nothing
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If txtOID.Text = "" Then
    MsgBox "You must select a Trigger to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  CallingForm = "frmTriggers"
  OID = txtOID.Text
  Load frmComments
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdComment_Click"
End Sub

Public Sub cmdCreateTrig_Click()
On Error GoTo Err_Handler
    gTrigger_Name = ""
    gTrigger_Table = ""
    Load frmAddTrigger
    frmAddTrigger.Show
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdCreateTrig_Click"
End Sub

Public Sub cmdDropTrig_Click()
On Error GoTo Err_Handler
Dim szDropStr As String
Dim szTrigger_name As String
Dim szTrigger_table As String
Dim iLoop As Long
Dim iListCount As Long

  If MsgBox("Are you sure you wish to drop Trigger(s)?", vbYesNo + vbQuestion, _
            "Confirm Trigger(s) deletion") = vbYes Then
                   
         StartMsg "Dropping Trigger..."
         
         iListCount = lstTrig.ListCount
         For iLoop = 0 To iListCount - 1
            If lstTrig.Selected(iLoop) = True Then
                cmp_Trigger_ParseName lstTrig.List(iLoop), szTrigger_name, szTrigger_table
                cmp_Trigger_GetValues szTrigger_PostgreSqlTable, 0, szTrigger_name, szTrigger_table
                cmp_Trigger_DropIfExists szTrigger_PostgreSqlTable, 0, szTrigger_name, szTrigger_table
             End If
          Next iLoop

          EndMsg
          cmdRefresh_Click
  End If

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdDropTrig_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler
  Dim iLoop As Long
  Dim iUbound As Long
  Dim szTrigger() As Variant
  Dim szTrigger_name As String
  Dim szTrigger_table As String
  Dim szQuery As String
  
  LogMsg "Loading Form: " & Me.Name
  StartMsg "Retrieving Trigger Names..."
  lstTrig.Clear
  lstTrig = 0
  
  If rsTrig.State <> adStateClosed Then rsTrig.Close
  If chkSystem.Value = 1 Then
    szTrigger_PostgreSqlTable = "pgadmin_triggers"
    szQuery = "SELECT trigger_name, trigger_table FROM " & szTrigger_PostgreSqlTable & " WHERE trigger_oid < " & LAST_SYSTEM_OID & " OR trigger_name LIKE 'pgadmin_%' OR trigger_name  LIKE 'pg_%' OR trigger_name LIKE 'RI_%' ORDER BY trigger_name"
    LogMsg "Executing: " & szQuery
    rsTrig.Open szQuery, gConnection, adOpenDynamic
  Else
    szTrigger_PostgreSqlTable = "pgadmin_dev_triggers"
    szQuery = "SELECT trigger_name, trigger_table FROM " & szTrigger_PostgreSqlTable & " WHERE trigger_name NOT LIKE 'pgadmin_%' AND trigger_name NOT LIKE 'pg_%' AND trigger_name NOT LIKE 'RI_%' ORDER BY trigger_name"
    LogMsg "Executing: " & szQuery
    rsTrig.Open szQuery, gConnection, adOpenDynamic
  End If
  
  If Not (rsTrig.EOF) Then
    szTrigger = rsTrig.GetRows
    iUbound = UBound(szTrigger, 2)
    For iLoop = 0 To iUbound
      szTrigger_name = szTrigger(0, iLoop)
      szTrigger_table = szTrigger(1, iLoop)
      lstTrig.AddItem szTrigger_name & " ON " & szTrigger_table
    Next iLoop
  End If
  
  Erase szTrigger
  lstTrig_Click
  
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdRefresh_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  Me.Width = 8325
  Me.Height = 4455
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4455 Then Me.Height = 4455
    End If
    lstTrig.Height = Me.ScaleHeight
    lstTrig.Width = Me.ScaleWidth - lstTrig.Left - fraDetails.Width - 25
    fraDetails.Left = lstTrig.Left + lstTrig.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, Form_Resize"
End Sub

Public Sub lstTrig_dblClick()
    cmdModifyTrig_Click
End Sub

Public Sub lstTrig_Click()
On Error GoTo Err_Handler
Dim iTrigger_type As Integer
Dim iTemp As Integer

Dim szTrigger_oid As Long
Dim szTrigger_name As String
Dim szTrigger_table As String
Dim szTrigger_function As String
Dim szTrigger_arguments As String
Dim szTrigger_foreach As String
Dim szTrigger_event As String
Dim szTrigger_executes As String
Dim szTrigger_Comments As String
Dim iInstr As Integer

    '----------------------------------------------------------------------------------
    ' Parse trigger name and arguments from List
    '----------------------------------------------------------------------------------
    If lstTrig.SelCount > 0 Then
        cmp_Trigger_ParseName lstTrig.Text, szTrigger_name, szTrigger_table
    Else
        szTrigger_name = ""
        szTrigger_table = ""
    End If
    '----------------------------------------------------------------------------------
    ' Lookup database
    '----------------------------------------------------------------------------------

    StartMsg "Retrieving trigger info..."
    szTrigger_oid = 0
    cmp_Trigger_GetValues szTrigger_PostgreSqlTable, szTrigger_oid, szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_foreach, szTrigger_executes, szTrigger_event, szTrigger_Comments
    txtOID.Text = Trim(Str(szTrigger_oid))
    If szTrigger_name <> "" Then
        If txtOID.Text = "0" Then txtOID.Text = "N.S."
    Else
        txtOID.Text = ""
    End If
    txtName.Text = szTrigger_name
    txtTable.Text = szTrigger_table
    txtFunction.Text = szTrigger_function
    txtForEach.Text = szTrigger_foreach
    txtExecutes.Text = szTrigger_executes
    txtEvent.Text = szTrigger_event
    txtComments.Text = szTrigger_Comments
    
    CmdTrigButton
    
    EndMsg

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTriggers, lstTrig_Click"
End Sub

Public Sub CmdTrigButton()
    Dim bSystem As Boolean
    bSystem = (chkSystem.Value = 1)
    cmdButtonActivate bSystem, lstTrig.SelCount, cmdCreateTrig, cmdModifyTrig, cmdDropTrig, cmdExportTrig, cmdComment, cmdRefresh
End Sub
