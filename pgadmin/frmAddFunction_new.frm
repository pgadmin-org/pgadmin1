VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Begin VB.Form frmAddFunction_new 
   Caption         =   "Function"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmAddFunction_new.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleMode       =   0  'User
   ScaleWidth      =   8931.034
   Begin VB.Frame fraDetails 
      Caption         =   "Function Details"
      Height          =   5595
      Left            =   4500
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.ListBox lstArguments 
         Height          =   1230
         Left            =   900
         TabIndex        =   17
         ToolTipText     =   "List of input arguments."
         Top             =   2205
         Width           =   2355
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   3330
         TabIndex        =   16
         Top             =   1845
         Width           =   915
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up"
         Height          =   330
         Left            =   3330
         TabIndex        =   15
         Top             =   2205
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   330
         Left            =   3330
         TabIndex        =   14
         Top             =   3105
         Width           =   915
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "&Down"
         Height          =   330
         Left            =   3330
         TabIndex        =   13
         Top             =   2565
         Width           =   915
      End
      Begin VB.ComboBox cboArguments 
         Height          =   315
         ItemData        =   "frmAddFunction_new.frx":030A
         Left            =   900
         List            =   "frmAddFunction_new.frx":030C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Input arguments of your function."
         Top             =   1845
         Width           =   2310
      End
      Begin VB.ComboBox cboReturnType 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   $"frmAddFunction_new.frx":030E
         Top             =   1485
         Width           =   3345
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   900
         MaxLength       =   31
         TabIndex        =   8
         ToolTipText     =   $"frmAddFunction_new.frx":03BF
         Top             =   1170
         Width           =   3345
      End
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   2010
         Left            =   900
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3525
         Width           =   3345
      End
      Begin vsAdoSelector.VS_AdoSelector vssLanguage 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         ToolTipText     =   $"frmAddFunction_new.frx":04F1
         Top             =   855
         Width           =   3345
         _ExtentX        =   5900
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
         SQL             =   "SELECT ""lanname"", ""lanname"" FROM ""pg_language"" WHERE ""lanname"" <> 'Internal'"
      End
      Begin VB.Label lblReturnType 
         AutoSize        =   -1  'True
         Caption         =   "Returns"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   1530
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label lblLanguage 
         AutoSize        =   -1  'True
         Caption         =   "Language"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   5
         Top             =   3510
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arguments"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1890
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Save function"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   $"frmAddFunction_new.frx":0602
      Top             =   45
      Width           =   1402
   End
   Begin VB.TextBox txtPath 
      Height          =   5520
      Left            =   1477
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "List of available functions."
      Top             =   45
      Width           =   2968
   End
   Begin VB.Label New 
      Caption         =   "Label2"
      Height          =   285
      Left            =   225
      TabIndex        =   18
      Top             =   765
      Width           =   600
   End
End
Attribute VB_Name = "frmAddFunction_new"
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
Dim lng_OpenFunction_OID As Long

Private Sub cboReturnType_Click()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, cboReturnType_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
  lstArguments.AddItem cboArguments.Text
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdAdd_Click"
End Sub

Private Sub Gen_SQL()
On Error GoTo Err_Handler
Dim szCreateStr As String
Dim X As Integer
  fMainForm.txtSQLPane.Text = "CREATE FUNCTION " & QUOTE & txtName.Text & QUOTE & vbCrLf & "  ("
  For X = 0 To lstArguments.ListCount - 2
    fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & lstArguments.List(X) & ", "
  Next X
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & lstArguments.List(X) & ") "
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  RETURNS " & cboReturnType.Text & " "
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  AS '" & txtPath.Text & "' "
  fMainForm.txtSQLPane.Text = fMainForm.txtSQLPane.Text & vbCrLf & "  LANGUAGE '" & vssLanguage.Text & "'"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, Gen_SQL"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim szCreateStr As String
Dim ArgList As String
Dim X As Integer
  If txtName.Text = "" Then
    MsgBox "You must enter a name for the function!", vbExclamation, "Error"
    Exit Sub
  End If
  If cboReturnType.Text = "" Then
    MsgBox "You must select a return data type for the function!", vbExclamation, "Error"
    Exit Sub
  End If
  If vssLanguage.Text = "" Then
    MsgBox "You must select a language for the function!", vbExclamation, "Error"
    Exit Sub
  End If
  If vssLanguage.Text = "sql" Then
    If txtPath.Text = "" Then
      MsgBox "You must enter an SQL statement!", vbExclamation, "Error"
      Exit Sub
    End If
  Else
    If txtPath.Text = "" Then
      MsgBox "You must enter the path to the object library containing the function!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  
  StartMsg "Saving function..."
  
  ' Build function arguments
  ArgList = ""
  For X = 0 To lstArguments.ListCount - 1
    ArgList = ArgList & lstArguments.List(X) & ", "
  Next X
  If ArgList <> "" Then ArgList = Left(ArgList, Len(ArgList) - 2)
  
   ' Drop function if exists
  If lng_OpenFunction_OID <> 0 Then cmp_Function_DropIfExists lng_OpenFunction_OID
  
  ' Create function
  cmp_Function_Create txtName.Text, ArgList, cboReturnType.Text, txtPath.Text, vssLanguage.Text
  
  ' Refresh function list
  frmFunctions.cmdRefresh_Click
  
  EndMsg
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdAdd_Click"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
  If lstArguments.ListIndex = -1 Then
    MsgBox "You must select an argument to remove!", vbExclamation, "Error"
    Exit Sub
  End If
  lstArguments.RemoveItem lstArguments.ListIndex
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdAdd_Click"
End Sub

Private Sub cmdUp_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstArguments.ListIndex = -1 Then
    MsgBox "You must select an argument to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstArguments.ListIndex = 0 Then
    MsgBox "This argument is already at the top!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstArguments.List(lstArguments.ListIndex - 1)
  lstArguments.List(lstArguments.ListIndex - 1) = lstArguments.List(lstArguments.ListIndex)
  lstArguments.List(lstArguments.ListIndex) = Temp
  lstArguments.ListIndex = lstArguments.ListIndex - 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdUp_Click"
End Sub

Private Sub cmdDown_Click()
On Error GoTo Err_Handler
Dim Temp As String
  If lstArguments.ListIndex = -1 Then
    MsgBox "You must select an argument to move!", vbExclamation, "Error"
    Exit Sub
  End If
  If lstArguments.ListIndex = lstArguments.ListCount - 1 Then
    MsgBox "This argument is already at the bottom!", vbExclamation, "Error"
    Exit Sub
  End If
  Temp = lstArguments.List(lstArguments.ListIndex + 1)
  lstArguments.List(lstArguments.ListIndex + 1) = lstArguments.List(lstArguments.ListIndex)
  lstArguments.List(lstArguments.ListIndex) = Temp
  lstArguments.ListIndex = lstArguments.ListIndex + 1
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, cmdDown_Click"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState = 0 Then
    If Me.Width < 9000 Then Me.Width = 9000
    If Me.Height < 6000 Then Me.Height = 6000
  End If
    txtPath.Height = Me.ScaleHeight
    txtPath.Width = Me.ScaleWidth - txtPath.Left - fraDetails.Width - 25
    fraDetails.Left = txtPath.Left + txtPath.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtPath.Height = fraDetails.Height - txtPath.Top - 100
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsTypes As New Recordset
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 4110
  Me.Width = 4275
  
  ' Retrieve data types
  StartMsg "Retrieving data types and languages..."
  If rsTypes.State <> adStateClosed Then rsTypes.Close
  LogMsg "Executing: SELECT typname FROM pg_type WHERE typrelid = 0 ORDER BY typname"
  rsTypes.Open "SELECT typname FROM pg_type WHERE typrelid = 0 ORDER BY typname", gConnection, adOpenForwardOnly
  cboReturnType.Clear
  cboArguments.Clear
  cboReturnType.AddItem "opaque"
  While Not rsTypes.EOF
    If Mid(rsTypes!typname, 1, 1) <> "_" Then
      cboReturnType.AddItem rsTypes!typname
      cboArguments.AddItem rsTypes!typname
    End If
    rsTypes.MoveNext
  Wend
  If rsTypes.BOF <> True Then rsTypes.MoveFirst
  
    ' Retrieve languages
  vssLanguage.Connect = Connect
  vssLanguage.SQL = "SELECT language_name, language_name FROM pgadmin_languages ORDER BY language_name"
  LogMsg "Executing: " & vssLanguage.SQL
  vssLanguage.LoadList
  lstArguments.Clear
  EndMsg
   
    ' Write query
  Gen_SQL
  Set rsTypes = Nothing
  
      ' Retrieve function if exists
  lng_OpenFunction_OID = gPostgresOBJ_OID
  gPostgresOBJ_OID = 0
    If lng_OpenFunction_OID <> 0 Then
    Me.Caption = "Modify function"
    Function_Load
  Else
    Me.Caption = "Create function"
  End If
  
  Exit Sub
Err_Handler:
  Set rsTypes = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddFunction, Form_Load"
End Sub

Private Sub New_Click()

End Sub

Private Sub txtName_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, txtName_Change"
End Sub

Private Sub txtPath_Change()
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, txtPath_Change"
End Sub

Private Sub vssLanguage_ItemSelected(Item As String, ItemText As String)
On Error GoTo Err_Handler
  Gen_SQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddDatabase, vssLanguage_ItemSelected"
End Sub

Private Sub Function_Load()
On Error GoTo Err_Handler
    Dim temp_arg_list As Variant
    Dim temp_arg_item As Variant
    Dim rsFunc As New Recordset
    StartMsg "Retrieving Function information..."

    If rsFunc.State <> adStateClosed Then rsFunc.Close
          LogMsg "Executing: SELECT * FROM pgadmin_functions WHERE Function_OID = " & lng_OpenFunction_OID
          rsFunc.Open "SELECT * FROM pgadmin_functions WHERE function_OID = " & lng_OpenFunction_OID, gConnection, adOpenDynamic
    
    ' Initialize form with values from pgadmin_function
    txtName = rsFunc!Function_name & ""
    txtPath = Replace(rsFunc!Function_source & "", "'", "''")
    vssLanguage.Text = rsFunc!Function_language & ""
    
    If (rsFunc!Function_returns & "" <> "") Then
        cboReturnType.Text = rsFunc!Function_returns & ""
    Else
        cboReturnType.Text = "opaque" ' Review if we could not put it in pgadmin_functions
    End If
    
    temp_arg_list = Split(rsFunc!Function_arguments, ",")
    For Each temp_arg_item In temp_arg_list
         cboArguments.Text = Trim(temp_arg_item)
         cmdAdd_Click
    Next
    
    rsFunc.Close
    EndMsg
    
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdRefresh_Click"
End Sub

