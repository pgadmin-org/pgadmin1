VERSION 5.00
Object = "{D4E5B983-69B8-11D3-9975-009027427025}#1.4#0"; "vsadoselector.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmAddFunction 
   Caption         =   "Function"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmAddFunction.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   8880
   Begin VB.Frame fraDetails 
      Caption         =   "Function Details"
      Height          =   5595
      Left            =   4500
      TabIndex        =   13
      Top             =   0
      Width           =   4335
      Begin HighlightBox.HBX txtComments 
         Height          =   2085
         Left            =   135
         TabIndex        =   12
         ToolTipText     =   "Enter or update the comment for this object."
         Top             =   3420
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   3678
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Comments"
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   3345
      End
      Begin VB.ListBox lstArguments 
         Height          =   1425
         Left            =   900
         TabIndex        =   7
         ToolTipText     =   "List of input arguments."
         Top             =   1890
         Width           =   2355
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   3330
         TabIndex        =   8
         Top             =   1890
         Width           =   915
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up"
         Height          =   330
         Left            =   3330
         TabIndex        =   9
         Top             =   2250
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   330
         Left            =   3330
         TabIndex        =   11
         Top             =   2970
         Width           =   915
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "&Down"
         Height          =   330
         Left            =   3330
         TabIndex        =   10
         Top             =   2610
         Width           =   915
      End
      Begin VB.ComboBox cboArguments 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Input arguments of your function."
         Top             =   1530
         Width           =   2310
      End
      Begin VB.ComboBox cboReturnType 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1170
         Width           =   3345
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   900
         MaxLength       =   31
         TabIndex        =   4
         Top             =   855
         Width           =   3345
      End
      Begin vsAdoSelector.VS_AdoSelector vssLanguage 
         Height          =   315
         Left            =   900
         TabIndex        =   3
         Top             =   540
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblReturnType 
         AutoSize        =   -1  'True
         Caption         =   "Returns"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   1215
         Width           =   510
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   900
         Width           =   420
      End
      Begin VB.Label lblLanguage 
         AutoSize        =   -1  'True
         Caption         =   "Language"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   585
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arguments"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   1530
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Save function"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1402
   End
   Begin HighlightBox.HBX txtPath 
      Height          =   5520
      Left            =   1477
      TabIndex        =   1
      ToolTipText     =   "Enter the Object path or Function code."
      Top             =   45
      Width           =   2968
      _ExtentX        =   5239
      _ExtentY        =   9737
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Function Definition"
   End
End
Attribute VB_Name = "frmAddFunction"
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
Dim szFunction_name_old As String
Dim szFunction_arguments_old As String

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
    Dim intLoop As Integer
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szFunction_returns As String
    Dim szFunction_source As String
    Dim szFunction_language As String
    
    szFunction_name = txtName.Text
    szFunction_arguments = ""
    For intLoop = 0 To lstArguments.ListCount - 2
        If szFunction_arguments <> "" Then szFunction_arguments = szFunction_arguments & ", "
        szFunction_arguments = szFunction_arguments & lstArguments.List(intLoop)
    Next intLoop
    
    szFunction_returns = cboReturnType.Text
    szFunction_source = txtPath.Text
    szFunction_language = vssLanguage.Text
    
    fMainForm.txtSQLPane.Text = cmp_Function_CreateSQL(szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language)

    Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddFunction, Gen_SQL"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo Err_Handler
Dim szCreateStr As String
Dim ArgList As String
Dim x As Integer
Dim szFunction_pgTable As String

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
  For x = 0 To lstArguments.ListCount - 1
    ArgList = ArgList & lstArguments.List(x) & ", "
  Next x
  If ArgList <> "" Then ArgList = Left(ArgList, Len(ArgList) - 2)
  
    If DevMode = True Then
          szFunction_pgTable = gDevPostgresqlTables & "_functions"
      Else
          szFunction_pgTable = "pgadmin_functions"
    End If
    
    If szFunction_name_old <> "" Then cmp_Function_DropIfExists szFunction_pgTable, szFunction_name_old, szFunction_arguments_old
    cmp_Function_DropIfExists szFunction_pgTable, txtName.Text, ArgList
    cmp_Function_Create szFunction_pgTable, txtName.Text, ArgList, cboReturnType.Text, txtPath.Text, vssLanguage.Text, "", txtComments.Text
    
    
    ' Refresh function list
    frmFunctions.cmdRefresh_Click
   
Unload Me

  
  EndMsg
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
  txtPath.Minimise
  txtComments.Minimise
  If Me.WindowState <> 1 Then
    If Me.WindowState = 0 Then
      If Me.Width < 9000 Then Me.Width = 9000
      If Me.Height < 6000 Then Me.Height = 6000
    End If
    txtPath.Height = Me.ScaleHeight
    txtPath.Width = Me.ScaleWidth - txtPath.Left - fraDetails.Width - 25
    fraDetails.Left = txtPath.Left + txtPath.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtPath.Height = fraDetails.Height - txtPath.Top - 100
    txtComments.Height = fraDetails.Height - txtComments.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmAddColumn, Form_Resize"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim rsTypes As New Recordset
    
Dim temp_arg_list As Variant
Dim temp_arg_item As Variant
    
Dim szFunction_pgTable As String
Dim szFunction_name As String
Dim szFunction_arguments As String
Dim szFunction_returns As String
Dim szFunction_source As String
Dim szFunction_language As String
Dim szFunction_owner As String
Dim szFunction_comments As String
   
  ' Log
  LogMsg "Loading Form: " & Me.Name
  Me.Height = 4110
  Me.Width = 4275
  txtPath.Wordlist = TextColours
    
    ' Remember initial values of function_name and function_arguments
    szFunction_name_old = gFunction_Name
    szFunction_name = gFunction_Name
    
    szFunction_arguments_old = gFunction_Arguments
    szFunction_arguments = gFunction_Arguments
    
    gFunction_Name = ""
    gFunction_Arguments = ""
      
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
           
        If szFunction_name_old <> "" Then
              Me.Caption = "Modify function"
              
              ' get function values
              If DevMode = True Then
                szFunction_pgTable = gDevPostgresqlTables & "_functions"
              Else
                szFunction_pgTable = "pgadmin_functions"
              End If
              
              cmp_Function_GetValues szFunction_pgTable, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
              
              ' Initialize form
              txtName = szFunction_name
              txtPath.Text = szFunction_source
              vssLanguage.Text = szFunction_language
              cboReturnType.Text = szFunction_returns
              txtOwner = szFunction_owner
              txtComments.Text = szFunction_comments
              
              temp_arg_list = Split(szFunction_arguments, ",")
              For Each temp_arg_item In temp_arg_list
                   cboArguments.Text = Trim(temp_arg_item)
                   cmdAdd_Click
              Next
              
            If txtOwner = "" Then txtOwner = "N.S."
        Else
           Me.Caption = "Create function"
           txtOwner = "N.S."
        End If
    
        ' Write query
    Gen_SQL
    Set rsTypes = Nothing
    
    Exit Sub
Err_Handler:
  Set rsTypes = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmAddFunction, Form_Load"
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

