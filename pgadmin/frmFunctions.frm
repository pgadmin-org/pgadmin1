VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#5.0#0"; "HighlightBox.ocx"
Begin VB.Form frmFunctions 
   Caption         =   "Functions"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   8880
   Begin VB.CommandButton cmdExportFunc 
      Caption         =   "Export Function"
      Height          =   330
      Left            =   45
      TabIndex        =   26
      ToolTipText     =   "Delete the selected function."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdRebuild 
      BackColor       =   &H80000018&
      Caption         =   "&Rebuild project"
      Height          =   330
      Left            =   45
      TabIndex        =   22
      ToolTipText     =   "Checks and rebuilds dependencies on functions, triggers and views."
      Top             =   2565
      Width           =   1410
   End
   Begin VB.CommandButton cmdModifyFunc 
      Caption         =   "&Modify Function"
      Height          =   330
      Left            =   45
      TabIndex        =   21
      ToolTipText     =   "Create a new function."
      Top             =   405
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   20
      Top             =   2970
      Width           =   1380
      Begin VB.CheckBox chkFunctions 
         Caption         =   "Functions"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Select to view system functions."
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected function."
      Top             =   1485
      Width           =   1410
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Refresh the list of function."
      Top             =   1845
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropFunc 
      Caption         =   "&Drop Function"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected function."
      Top             =   765
      Width           =   1410
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Function Details"
      Height          =   5595
      Left            =   4500
      TabIndex        =   12
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1170
         Width           =   3345
      End
      Begin VB.TextBox txtLanguage 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   855
         Width           =   3345
      End
      Begin VB.TextBox txtReturns 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1500
         Width           =   3345
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   3345
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   3345
      End
      Begin VB.TextBox txtArguments 
         BackColor       =   &H8000000F&
         Height          =   1635
         Left            =   900
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1845
         Width           =   3345
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   795
         Left            =   900
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3525
         Width           =   3345
      End
      Begin HighlightBox.HBX txtFunction 
         Height          =   1095
         Left            =   45
         TabIndex        =   25
         Top             =   4500
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   1931
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
         ScrollBars      =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   23
         Top             =   1215
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arguments"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Top             =   1845
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Returns"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   16
         Top             =   1530
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   15
         Top             =   4275
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Language"
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   14
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   13
         Top             =   3510
         Width           =   735
      End
   End
   Begin VB.ListBox lstFunc 
      Height          =   5520
      ItemData        =   "frmFunctions.frx":0000
      Left            =   1485
      List            =   "frmFunctions.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   45
      Width           =   2985
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   3735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select SQL File"
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdCreateFunc 
      Caption         =   "&Create Function"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new function."
      Top             =   45
      Width           =   1410
   End
End
Attribute VB_Name = "frmFunctions"
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
Dim rsFunc As New Recordset

Private Sub cmdExportFunc_Click()
    Dim iLoop As Long
    Dim iListCount As Long
    Dim szExport As String
    Dim bExport As Boolean
    Dim szHeader As String
    
    Dim szFunction_OID As Long
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szFunction_returns As String
    Dim szFunction_source As String
    Dim szFunction_language As String
    Dim szFunction_owner As String
    Dim szFunction_comments As String
    
    bExport = False
    szExport = ""

    iListCount = lstFunc.ListCount
        
    For iLoop = 0 To iListCount - 1
        If lstFunc.Selected(iLoop) = True Then
            bExport = True
            szFunction_OID = 0
            ParseFunction lstFunc.List(iLoop), szFunction_name, szFunction_arguments
            cmp_Function_GetValues szFunction_OID, "", szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
            
            ' Header
            szExport = szExport & "/*" & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & szFunction_name & "(" & szFunction_arguments & ")" & " -> " & szFunction_returns & vbCrLf
            If szFunction_comments <> "" Then szExport = szExport & szFunction_comments & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & "*/" & vbCrLf
            
            ' Function
            szExport = szExport & cmp_Function_CreateSQL(szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language) & vbCrLf & vbCrLf
        End If
    Next iLoop
    
    If bExport Then
        szHeader = "/*" & vbCrLf
        szHeader = szHeader & Format(Now, "d mmmm yyyy") & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "The choice of the best developers, " & Format(Now, "d mmmm yyyy") & vbCrLf
        szHeader = szHeader & "PostgreSQL     www.postgresql.org" & vbCrLf
        szHeader = szHeader & "PgAdmin        www.greatbridge.org/project/pgadmin" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "*/" & vbCrLf & vbCrLf
        szExport = szHeader & szExport
        MsgExportToFile CommonDialog1, szExport, "sql", "Export functions"
    End If
End Sub

Public Sub cmdModifyFunc_Click()
' On Error GoTo Err_Handler

If txtOID <> "" Then
    ' This means we can open the function
    gPostgresOBJ_OID = Val(txtOID)
    
    ' Load form
    Load frmAddFunction
    frmAddFunction.Show
End If

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdModifyFunc_Click"
End Sub

Private Sub cmdRebuild_Click()
    If MsgBox("For the moment, the rebuilding feature does not keep comments and views ACL." & vbCrLf & "Please confirm you wish to continue.", vbYesNo + vbQuestion, _
            "Rebuild project") = vbYes Then
        comp_Project_Initialize
        comp_Project_Compile
    End If
End Sub

Private Sub lstFunc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXFunctions
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, lstFunc_MouseUp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsFunc = Nothing
End Sub

Private Sub chkFunctions_Click()
' On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, chkFunctions_Click"
End Sub

Public Sub cmdComment_Click()
' On Error GoTo Err_Handler
  If txtOID.Text = "" Then
    MsgBox "You must select a function to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  CallingForm = "frmFunctions"
  OID = txtOID.Text
  Load frmComments
  frmComments.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdComment_Click"
End Sub

Public Sub cmdCreateFunc_Click()
' On Error GoTo Err_Handler
  ' This means we will create the function
  gPostgresOBJ_OID = 0
  
  ' Load form
  Load frmAddFunction
  frmAddFunction.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdCreateFunc_Click"
End Sub

Public Sub cmdDropFunc_Click()
' On Error GoTo Err_Handler
    Dim szDropStr As String
    Dim iLoop As Long
    Dim iListCount As Long
    Dim szFunction_name As String
    Dim szFunction_arguments As String
   
    If MsgBox("Are you sure you wish to drop Function(s)?", vbYesNo + vbQuestion, _
            "Confirm Function Deletion") = vbYes Then
        
        StartMsg "Dropping Function(s)..."
        
        iListCount = lstFunc.ListCount
        For iLoop = 0 To iListCount - 1
            If lstFunc.Selected(iLoop) = True Then
                ParseFunction lstFunc.List(iLoop), szFunction_name, szFunction_arguments
                cmp_Function_GetValues 0, "", szFunction_name, szFunction_arguments
                
                szDropStr = "DROP FUNCTION " & QUOTE & szFunction_name & QUOTE & " (" & szFunction_arguments & ")"
                fMainForm.txtSQLPane.Text = szDropStr
                LogMsg "Executing: " & szDropStr
                gConnection.Execute szDropStr
                LogQuery szDropStr
             End If
        Next iLoop
        
        EndMsg
        
        cmdRefresh_Click
    End If
    
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdDropFunc_Click"
End Sub

Public Sub cmdRefresh_Click()
 On Error GoTo Err_Handler
  Dim szFunc() As Variant
  Dim iLoop As Long
  Dim iUbound As Long
  Dim szFunction_name As String
  Dim szFunction_arguments As String
  
  StartMsg "Retrieving Function Names..."
  lstFunc.Clear
  
  If rsFunc.State <> adStateClosed Then rsFunc.Close
  If chkFunctions.Value = 1 Then
    LogMsg "Executing: SELECT function_name, function_arguments FROM pgadmin_functions ORDER BY function_name"
    rsFunc.Open "SELECT function_name, function_arguments FROM pgadmin_functions ORDER BY function_name", gConnection, adOpenDynamic
  Else
    LogMsg "Executing: SELECT function_name, function_arguments FROM pgadmin_functions WHERE function_name NOT LIKE 'pgadmin_%' AND function_name NOT LIKE 'pg_%' AND function_oid > " & LAST_SYSTEM_OID & " ORDER BY function_name"
    rsFunc.Open "SELECT function_name, function_arguments FROM pgadmin_functions WHERE function_name NOT LIKE 'pgadmin_%' AND function_name NOT LIKE 'pg_%' AND function_oid > " & LAST_SYSTEM_OID & " ORDER BY function_name", gConnection, adOpenDynamic
  End If
  
  If Not (rsFunc.EOF) Then
    szFunc = rsFunc.GetRows
    iUbound = UBound(szFunc, 2)
    For iLoop = 0 To iUbound
         szFunction_name = szFunc(0, iLoop)
         szFunction_arguments = szFunc(1, iLoop)
        If szFunction_arguments <> "" Then
            lstFunc.AddItem szFunction_name & " (" & szFunction_arguments & ")"
        Else
            lstFunc.AddItem szFunction_name
        End If
    Next iLoop
  End If
  Erase szFunc
  txtName.Text = lstFunc
  lstFunc_Click
  CmdFuncButton
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdRefresh_Click"
End Sub

Private Sub Form_Load()
' On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8325
  Me.Height = 4455
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, Form_Load"
End Sub

Private Sub Form_Resize()
' On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 9000 Then Me.Width = 9000
      If Me.Height < 6000 Then Me.Height = 6000
    End If
    lstFunc.Height = Me.ScaleHeight
    lstFunc.Width = Me.ScaleWidth - lstFunc.Left - fraDetails.Width - 25
    fraDetails.Left = lstFunc.Left + lstFunc.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtFunction.Height = fraDetails.Height - txtFunction.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, Form_Resize"
End Sub

Public Sub lstFunc_dblClick()
    cmdModifyFunc_Click
End Sub

Public Sub lstFunc_Click()
On Error GoTo Err_Handler
    Dim szFunction_OID As Long
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szFunction_returns As String
    Dim szFunction_source As String
    Dim szFunction_language As String
    Dim szFunction_owner As String
    Dim szFunction_comments As String

    Dim iInstr As Integer
    
    '----------------------------------------------------------------------------------
    ' Retrieve function name and arguments from List
    '----------------------------------------------------------------------------------
    If lstFunc.SelCount > 0 Then
        ParseFunction lstFunc.Text, szFunction_name, szFunction_arguments
    Else
        szFunction_name = ""
        szFunction_arguments = ""
    End If
    '----------------------------------------------------------------------------------
    ' Lookup database
    '----------------------------------------------------------------------------------
    StartMsg "Retrieving Function Info..."
    szFunction_OID = 0
    cmp_Function_GetValues szFunction_OID, "pgadmin_functions", szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner
    txtOID.Text = Trim(Str(szFunction_OID))
    If txtOID.Text = 0 Then txtOID.Text = ""
    txtOwner.Text = szFunction_owner
    txtReturns.Text = szFunction_returns
    txtArguments.Text = szFunction_arguments
    txtFunction.Text = szFunction_source
    txtLanguage.Text = szFunction_language
    txtComments.Text = szFunction_comments
    txtName.Text = szFunction_name
    
    CmdFuncButton
    EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, lstFunc_Click"
End Sub

Private Sub ParseFunction(szInput As String, szFunction_name As String, szFunction_arguments As String)
    Dim iInstr As Integer
    iInstr = InStr(szInput, "(")
    If iInstr > 0 Then
        szFunction_name = Left(szInput, iInstr - 2)
        szFunction_arguments = Mid(szInput, iInstr + 1, Len(szInput) - iInstr - 1)
    Else
        szFunction_name = szInput
        szFunction_arguments = ""
    End If
End Sub

Public Sub CmdFuncButton()
    cmdButtonActivate lstFunc.SelCount, cmdCreateFunc, cmdModifyFunc, cmdDropFunc, cmdExportFunc, cmdComment, cmdRefresh
End Sub

