VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Object = "{44DFA8BA-326E-4D0F-8941-25E814743439}#1.0#0"; "TreeToys.ocx"
Begin VB.Form frmFunctions 
   Caption         =   "Functions"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmFunctions.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleMode       =   0  'User
   ScaleWidth      =   9000
   Begin MSComctlLib.ImageList ilBrowser 
      Left            =   585
      Top             =   2295
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":08D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":0A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctions.frx":0B8C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TreeToys.TreeToy trvBrowser 
      Height          =   5505
      Left            =   1485
      TabIndex        =   11
      Top             =   45
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9710
      NodeTips        =   1
      BorderStyle     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      Indentation     =   299.906
      LabelEdit       =   1
      LineStyle       =   1
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
   Begin VB.CommandButton cmdCopyProToDev 
      BackColor       =   &H80000018&
      Caption         =   "Load->Developt"
      Height          =   330
      Left            =   45
      TabIndex        =   23
      ToolTipText     =   "Compiles a repository function."
      Top             =   4275
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdCopyDevToPro 
      BackColor       =   &H80000018&
      Caption         =   "Compile unsafe"
      Height          =   330
      Left            =   45
      TabIndex        =   22
      ToolTipText     =   "Compiles a repository function."
      Top             =   3915
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdExportFunc 
      Caption         =   "Export Function(s)"
      Height          =   330
      Left            =   45
      TabIndex        =   21
      ToolTipText     =   "Delete the selected function."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdRebuild 
      BackColor       =   &H80000018&
      Caption         =   "Rebuild &Project"
      Height          =   330
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   "Checks and rebuilds project dependencies."
      Top             =   3555
      Width           =   1410
   End
   Begin VB.CommandButton cmdModifyFunc 
      Caption         =   "&Modify Function"
      Height          =   330
      Left            =   45
      TabIndex        =   17
      ToolTipText     =   "Create a new function."
      Top             =   405
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   16
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
      Visible         =   0   'False
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
      Caption         =   "&Drop Function(s)"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   0
      Width           =   4155
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   855
         Width           =   3210
      End
      Begin VB.TextBox txtLanguage 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   3210
      End
      Begin VB.TextBox txtReturns 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1185
         Width           =   3210
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   3210
      End
      Begin HighlightBox.HBX txtFunction 
         Height          =   1275
         Left            =   90
         TabIndex        =   8
         Top             =   2880
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2249
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
         Caption         =   "Function/Path to Object Code"
      End
      Begin HighlightBox.HBX txtComments 
         Height          =   1275
         Left            =   90
         TabIndex        =   9
         Top             =   4230
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2249
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
      Begin HighlightBox.HBX txtArguments 
         Height          =   1275
         Left            =   90
         TabIndex        =   7
         Top             =   1530
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2249
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
         Caption         =   "Arguments"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   19
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Returns"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   14
         Top             =   1215
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Language"
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   13
         Top             =   585
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   2295
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
' pgadmin - PostgreSQL db Administration/Management for Win32
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
Option Compare Text

Public dragNode As Node, dropNode As Node

Private iPro_Index As Integer
Private iDev_Index As Integer
Private iSys_Index As Integer

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Form
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 9000
  Me.Height = 6000
  
  txtFunction.Wordlist = TextColours
  Set trvBrowser.ImageList = Me.ilBrowser
  trvBrowser.Refresh
  cmdRefresh_Click

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtArguments.Minimise
  txtComments.Minimise
  txtFunction.Minimise
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If Me.Width < 9000 Then Me.Width = 9000
      If Me.Height < 6000 Then Me.Height = 6000
    End If

    trvBrowser.Height = Me.ScaleHeight
    trvBrowser.Width = Me.ScaleWidth - trvBrowser.Left - fraDetails.Width - 25
    fraDetails.Left = trvBrowser.Left + trvBrowser.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100

  End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, Form_Resize"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Buttons
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdCopyDevToPro_Click()
On Error GoTo Err_Handler
    cmp_function_tree_copy_devtopro trvBrowser
    cmdRefresh_Click
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdCopyDevToPro_Click"
End Sub

Private Sub cmdCopyProToDev_Click()
On Error GoTo Err_Handler
    cmp_function_tree_copy_protodev trvBrowser
    cmdRefresh_Click
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdCopyProToDev_Click"
End Sub

Private Sub cmdExportFunc_Click()
On Error GoTo Err_Handler
    cmp_function_tree_export trvBrowser, CommonDialog1
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdExportFunc_Click"
End Sub

Public Sub cmdModifyFunc_Click()
'On Error GoTo Err_Handler
    If txtName <> "" Then
        ModifyFunc txtName.Text & "(" & txtArguments.Text & ")"
    End If
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdModifyFunc_Click"
End Sub

Private Sub ModifyFunc(szFunction As String)
On Error GoTo Err_Handler

If szFunction <> "" Then
    ' Get name and arguments
    cmp_Function_ParseName szFunction, gFunction_Name, gFunction_Arguments
     
    ' Load form
    Load frmAddFunction
    frmAddFunction.Show
End If

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdModifyFunc_Click"
End Sub

Private Sub cmdRebuild_Click()
On Error GoTo Err_Handler

    cmp_Project_Rebuild
    cmdRefresh_Click
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdRebuild_Click"
End Sub

Private Sub chkFunctions_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, chkFunctions_Click"
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  'If txtName.Text = "" Then
  '  MsgBox "You must select a function to edit the comment for.", vbExclamation, "Error"
  '  Exit Sub
  'End If
  'Load frmComments
  'frmComments.Setup "frmFunctions", QUOTE & txtName.Text & QUOTE & "(" & txtArguments.Text & ")", Val(txtOID.Text)
  'frmComments.Show

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdComment_Click"
End Sub

Public Sub cmdCreateFunc_Click()
On Error GoTo Err_Handler
  ' This means we will create the function
  gFunction_Name = ""
  gFunction_Arguments = ""
  
  ' Load form
  Load frmAddFunction
  frmAddFunction.Show

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdCreateFunc_Click"
End Sub

Public Sub cmdDropFunc_Click()
On Error GoTo Err_Handler
Dim szText As String

szText = trvBrowser.TreeTextChecked
If MsgBox("Are you sure you wish to drop:" & vbCrLf & vbCrLf & szText, vbYesNo + vbQuestion, _
            "Connfirm?") = vbYes Then
    cmp_function_tree_drop trvBrowser
    cmdRefresh_Click
End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdDropFunc_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler

cmp_function_tree_refresh trvBrowser, CBool(chkFunctions), iPro_Index, iSys_Index, iDev_Index

cmdCopyDevToPro.Visible = DevMode
cmdCopyProToDev.Visible = DevMode
cmdRebuild.Visible = DevMode

CmdFuncButton

  Exit Sub
Err_Handler:
EndMsg
If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdRefresh_Click"
End Sub

Public Sub CmdFuncButton()
On Error GoTo Err_Handler

cmdButtonActivate trvBrowser, CBool(chkFunctions), iPro_Index, iSys_Index, iDev_Index, cmdCreateFunc, cmdModifyFunc, cmdDropFunc, cmdExportFunc, cmdComment, cmdRefresh

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, CmdFuncButton"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Treeview
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub trvBrowser_NodeCheck(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler

trvBrowser.FreezeCtl
trvBrowser.TreeSelectiveCheck Node
trvBrowser.UnFreezeCtl
    
CmdFuncButton

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, trvBrowser_NodeCheck"
End Sub

Private Sub trvBrowser_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler

    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szFunction_returns As String
    Dim szFunction_source As String
    Dim szFunction_language As String
    Dim szFunction_owner As String
    Dim szFunction_comments As String
    Dim szFunction_table As String
    
    '----------------------------------------------------------------------------------
    ' Retrieve function name and arguments from List
    '----------------------------------------------------------------------------------
        
    Dim szRoot As String
    If Node.Text <> "" Then
        cmp_Function_ParseName Node.Text, szFunction_name, szFunction_arguments
        szRoot = Left(Node.Key, 2)
        If szRoot = "P:" Or szRoot = "S:" Then
            szFunction_table = "pgadmin_functions"
        Else
            szFunction_table = gDevPostgresqlTables & "_functions"
        End If
    Else
        szFunction_name = ""
        szFunction_arguments = ""
    End If
    '----------------------------------------------------------------------------------
    ' Lookup database
    '----------------------------------------------------------------------------------
    StartMsg "Retrieving Function Info..."
    cmp_Function_GetValues szFunction_table, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
    
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
If Err.Number <> 0 Then LogError Err, "frmFunctions, trvBrowser_NodeClick"
End Sub

Private Sub trvBrowser_dblClick()
On Error GoTo Err_Handler

    If trvBrowser.SelectedItem.Parent Is Nothing Then Exit Sub
    If DevMode = True And trvBrowser.SelectedItem.Parent.Key = "Pro:" Or trvBrowser.SelectedItem.Parent.Key = "Sys:" Then Exit Sub
    
    ModifyFunc trvBrowser.SelectedItem.Text
  
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, trvBrowser_dblClick"
End Sub

Private Sub trvBrowser_OLEStartDrag(Data As MSComctlLib.DataObject, _
AllowedEffects As Long)
On Error GoTo Err_Handler

Set dragNode = trvBrowser.SelectedItem

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, trvBrowser_OLEStartDrag"
End Sub

Private Sub trvBrowser_MouseDown(Button As Integer, Shift As Integer, _
x As Single, y As Single)
On Error GoTo Err_Handler

    Set dragNode = trvBrowser.HitTest(x, y)
    Set dropNode = Nothing
    If Not (dragNode Is Nothing) Then
        dragNode.Selected = True
        trvBrowser_NodeClick dragNode
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, trvBrowser_MouseDown"
End Sub

Private Sub trvBrowser_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler

Dim sz_drag_key As String
Dim sz_drop_key As String

    Set dropNode = trvBrowser.HitTest(x, y)

    If Not (dragNode Is Nothing) And Not (dropNode Is Nothing) Then
        If dragNode.Key <> dropNode.Key Then
            If dragNode.Parent Is Nothing Then
               sz_drag_key = dragNode.Key
            Else
               sz_drag_key = dragNode.Parent.Key
            End If
            
            If dropNode.Parent Is Nothing Then
               sz_drop_key = dropNode.Key
            Else
               sz_drop_key = dropNode.Parent.Key
            End If
            
            Select Case sz_drag_key
                Case "Pro:"
                    If (sz_drop_key = "Dev:") Then
                            cmdCopyProToDev_Click
                    End If
                
                Case "Sys:"
                    If (sz_drop_key = "Dev:") Then
                            cmdCopyProToDev_Click
                    End If
                    
                 Case "Dev:"
                    If (sz_drop_key = "Pro:" Or sz_drop_key = "Sys:") Then
                       cmdCopyDevToPro_Click
                    End If
            End Select
         Else
            If Not (dropNode Is Nothing) Then
                trvBrowser_NodeCheck dropNode
            End If
        End If
    End If
    
    Set dragNode = Nothing
    Set dropNode = Nothing
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmFunctions, trvBrowser_OLEDragDrop"
End Sub
